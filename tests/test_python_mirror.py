"""Offline regression tests: stdlib only, no real SharePoint or Visio calls."""

import json
from dataclasses import replace
from pathlib import Path
import sys
import tempfile
from types import ModuleType, SimpleNamespace
import unittest
from unittest import mock
from uuid import UUID

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "python"))
import visio_sharepoint_mirror as mirror


SITE = "https://sharepoint.example.invalid/sites/test"
ROOT = "/sites/test/Documents/Mirror"
TARGET_ID = UUID("11111111-1111-4111-8111-111111111111")
OTHER_ID = UUID("22222222-2222-4222-8222-222222222222")


class Response:
    def __init__(self, body=None, status=200):
        self.status_code = status
        self.body = {} if body is None else body
        self.headers = {}
        self.text = json.dumps(self.body)
        self.content = self.text.encode("utf-8")
        self.url = ""

    def json(self):
        return self.body

    def raise_for_status(self):
        if self.status_code >= 400:
            raise RuntimeError(f"Mock HTTP {self.status_code}")


class Session:
    def __init__(self, responder=None):
        self.calls = []
        self.responder = responder or self.default_response
        self.headers = {}
        self.closed = False

    @staticmethod
    def default_response(call):
        url = call["url"]
        if url.endswith("/_api/contextinfo"):
            return Response({"d": {"GetContextWebInformation": {"FormDigestValue": "test-digest"}}})
        if "/lists/getbytitle(" in url:
            return Response({"d": {"ForceCheckout": False, "RootFolder": {"ServerRelativeUrl": "/sites/test/Documents"}}})
        if call["method"] == "GET" and "/GetFolderByServerRelativeUrl(" in url:
            return Response({"d": {"UniqueId": str(TARGET_ID), "ServerRelativeUrl": ROOT}})
        if url.endswith("/recycle()"):
            return Response({"d": "recycled"})
        raise AssertionError(f"Unexpected mocked request: {call['method']} {url}")

    def request(self, method, url, **kwargs):
        call = {"method": method.upper(), "url": url, **kwargs}
        call["headers"] = dict(kwargs.get("headers", {}))
        self.calls.append(call)
        response = self.responder(call)
        response.url = url
        return response

    def close(self):
        self.closed = True


class FakeSharePoint:
    """Only the orchestration protocol; every action stays in memory."""

    def __init__(self, remote, upload_error=False):
        self.root = ROOT
        self.remote = remote
        self.upload_error = upload_error
        self.events = []

    def initialize(self):
        self.events.append("initialize")

    def assert_target(self):
        self.events.append("assert_target")

    def ensure_folders(self, plan):
        self.events.append("ensure_folders")

    def upload(self, artifacts):
        self.events.append("upload")
        if self.upload_error:
            raise mirror.MirrorError("Mock upload failure")

    def inventory(self):
        self.events.append("inventory")
        return self.remote

    def recycle(self, remote, plan):
        self.events.append("recycle")
        files = sum(mirror.path_key(item.relative) not in plan.file_keys for item in remote.files)
        folders = sum(mirror.path_key(item.relative) not in plan.folder_keys for item in remote.folders)
        return files, folders

    def close(self):
        self.events.append("close")


class MirrorTests(unittest.TestCase):
    def setUp(self):
        self.sandbox = tempfile.TemporaryDirectory(prefix="python-mirror-tests-")
        self.addCleanup(self.sandbox.cleanup)
        self.base = Path(self.sandbox.name)
        self.source = self.base / "Source"
        self.source.mkdir()
        self.config = mirror.Config(self.source, SITE, "Documents", "Mirror", TARGET_ID, self.base / "mirror.log")

    def source_file(self, target="Source/A.pdf", size=10, mtime=100):
        return mirror.SourceFile(self.source / "A.vsdx", "A.vsdx", target, size, mtime)

    @staticmethod
    def remote_item(relative):
        return mirror.RemoteItem(relative, ROOT + "/" + relative)

    def complete_remote(self):
        return mirror.RemoteInventory([self.remote_item("Source/A.pdf")], [self.remote_item("Source")])

    @staticmethod
    def converter(sources, workdir):
        artifacts = []
        for index, source in enumerate(sources):
            pdf = workdir / f"{index}.pdf"
            pdf.write_bytes(b"%PDF-1.7\nmock\n")
            artifacts.append(mirror.Artifact(source.target, pdf))
        return artifacts

    def config_json(self):
        return {
            "SourcePath": str(self.source), "SharePointSiteUrl": SITE, "LibraryName": "Documents",
            "TargetFolderPath": "Mirror", "TargetFolderUniqueId": str(TARGET_ID),
            "LogPath": str(self.base / "mirror.log"),
        }

    def load_config(self, data):
        path = self.base / "mirror.json"
        path.write_text(json.dumps(data), encoding="utf-8")
        return mirror.Config.load(path)

    @staticmethod
    def com_modules(documents):
        pythoncom = ModuleType("pythoncom")
        pythoncom.COINIT_APARTMENTTHREADED = 2
        pythoncom.CLSCTX_LOCAL_SERVER = 4
        pythoncom.CoInitializeEx = mock.Mock()
        pythoncom.CoUninitialize = mock.Mock()
        visio = SimpleNamespace(Documents=SimpleNamespace(OpenEx=mock.Mock(side_effect=documents)), Quit=mock.Mock())
        client = ModuleType("win32com.client")
        client.DispatchEx = mock.Mock(return_value=visio)
        win32com = ModuleType("win32com")
        win32com.client = client
        return {"pythoncom": pythoncom, "win32com": win32com, "win32com.client": client}, visio

    @staticmethod
    def exporting_document():
        document = mock.Mock()
        document.ExportAsFixedFormat.side_effect = lambda fixed, output, intent, pages: Path(output).write_bytes(b"%PDF-1.7\nmock\n")
        return document

    def test_config_accepts_complete_values_and_rejects_placeholders(self):
        self.assertEqual(self.load_config(self.config_json()), self.config)
        for name in self.config_json():
            with self.subTest(field=name):
                data = self.config_json()
                data[name] = "__PLATZHALTER_TEST__"
                with self.assertRaises(mirror.MirrorError):
                    self.load_config(data)

    def test_config_rejects_unsafe_targets_source_root_and_empty_guid(self):
        invalid = [
            ("TargetFolderPath", value)
            for value in ("", ".", "/", "../Other", "Mirror/../Other", "\\Mirror", "Mirror//Child", "Bad#Name")
        ]
        invalid += [("SourcePath", self.source.anchor), ("TargetFolderUniqueId", str(UUID(int=0))), ("SharePointSiteUrl", "http://sharepoint.example.invalid/sites/test")]
        for name, value in invalid:
            with self.subTest(field=name, value=value):
                data = self.config_json()
                data[name] = value
                with self.assertRaises(mirror.MirrorError):
                    self.load_config(data)

    def test_mutex_identity_uses_only_target_guid(self):
        name = mirror.mutex_name(self.config)
        self.assertEqual(name, "Global\\PPSI_VisioSharePointMirror_" + TARGET_ID.hex.upper())
        for config in (
            replace(self.config, site_url=SITE.upper() + "/"),
            replace(self.config, site_url="https://sharepoint-alias.example.invalid/sites/test", target_folder="Renamed/Folder"),
            replace(self.config, library_name="RenamedLibrary", source=self.base / "OtherSource"),
        ):
            with self.subTest(config=config):
                self.assertEqual(mirror.mutex_name(config), name)
        self.assertNotEqual(mirror.mutex_name(replace(self.config, target_id=OTHER_ID)), name)

    def test_scan_filters_visio_files_and_preserves_relative_structure(self):
        (self.source / "Sub").mkdir()
        for name in ("Z.vsd", "Sub/A.vsdx", "Sub/B.vsdm", "Ignore.pdf", "~$Lock.vsdx"):
            (self.source / name).write_bytes(b"sample")
        files = mirror.scan_source(self.source)
        self.assertEqual({item.target for item in files}, {"Source/Z.pdf", "Source/Sub/A.pdf", "Source/Sub/B.pdf"})
        self.assertTrue(all(item.size == 6 and item.mtime_ns > 0 for item in files))
        self.assertEqual(files, mirror.scan_source(self.source))

    def test_scan_rejects_missing_source_root_and_reparse_point(self):
        cases = ((self.base / "Missing", (mirror.MirrorError, FileNotFoundError)), (Path(self.source.anchor), mirror.MirrorError))
        for path, error in cases:
            with self.subTest(path=path), self.assertRaises(error):
                mirror.scan_source(path)
        # lstat must reject the reparse attribute before traversing the root.
        actual_lstat = Path.lstat

        def reparse_lstat(path, *args, **kwargs):
            result = actual_lstat(path, *args, **kwargs)
            if path == self.source:
                values = {name: getattr(result, name) for name in dir(result) if name.startswith("st_")}
                values["st_file_attributes"] = values.get("st_file_attributes", 0) | 0x400
                return SimpleNamespace(**values)
            return result

        with mock.patch.object(Path, "lstat", reparse_lstat), self.assertRaises(mirror.MirrorError):
            mirror.scan_source(self.source)

    def test_scan_rejects_two_visio_formats_with_same_pdf_name(self):
        (self.source / "Same.vsd").write_bytes(b"old")
        (self.source / "Same.vsdx").write_bytes(b"new")
        with self.assertRaises(mirror.MirrorError):
            mirror.scan_source(self.source)

    def test_plan_handles_empty_input_keys_values_and_file_folder_collisions(self):
        empty = mirror.build_plan([])
        self.assertEqual(empty.file_keys, set())
        self.assertEqual(empty.folder_keys, set())
        self.assertEqual(empty.folders, ())
        plan = mirror.build_plan([self.source_file("Values/Sub/A.pdf"), self.source_file("Keys/Sub/B.pdf")])
        self.assertEqual(plan.file_keys, {"values/sub/a.pdf", "keys/sub/b.pdf"})
        self.assertEqual(set(plan.folders), {"Values", "Values/Sub", "Keys", "Keys/Sub"})
        for parent in ("Values", "Keys"):
            self.assertLess(plan.folders.index(parent), plan.folders.index(parent + "/Sub"))
        with self.assertRaises(mirror.MirrorError):
            mirror.build_plan([self.source_file("Source/A.pdf"), self.source_file("Source/a.pdf/Child.pdf")])

    def test_remote_requires_expected_files_and_folders_including_keys(self):
        plan = mirror.build_plan([self.source_file()])
        mirror.validate_remote(ROOT, self.complete_remote(), plan)
        for remote in (mirror.RemoteInventory([], [self.remote_item("Source")]), mirror.RemoteInventory([self.remote_item("Source/A.pdf")], [])):
            with self.subTest(remote=remote), self.assertRaises(mirror.MirrorError):
                mirror.validate_remote(ROOT, remote, plan)
        plan = mirror.build_plan([self.source_file("Keys/Sub/A.pdf")])
        remote = mirror.RemoteInventory([self.remote_item("Keys/Sub/A.pdf")], [self.remote_item("Keys")])
        with self.assertRaises(mirror.MirrorError):
            mirror.validate_remote(ROOT, remote, plan)

    def test_remote_rejects_outside_paths_and_target_root(self):
        plan = mirror.build_plan([])
        for item in (mirror.RemoteItem("Outside.pdf", "/sites/test/Documents/Outside.pdf"), mirror.RemoteItem("", ROOT)):
            with self.subTest(item=item), self.assertRaises(mirror.MirrorError):
                mirror.validate_remote(ROOT, mirror.RemoteInventory([], [item]), plan)

    def test_http_get_and_binary_write_use_correct_digest_flow(self):
        def respond(call):
            if call["url"].endswith("/_api/contextinfo"):
                return Session.default_response(call)
            return Response({"d": "ok"})

        session = Session(respond)
        sp = mirror.SharePoint(self.config, session=session)
        self.assertEqual(sp._request("GET", SITE + "/_api/web"), {"d": "ok"})
        self.assertEqual(len(session.calls), 1)
        self.assertNotIn("X-RequestDigest", session.calls[0]["headers"])
        body = bytes((0, 1, 127, 128, 255))
        sp._request("POST", SITE + "/_api/upload", write=True, data=body)
        self.assertEqual(len(session.calls), 3)
        self.assertTrue(session.calls[1]["url"].endswith("/_api/contextinfo"))
        self.assertNotIn("X-RequestDigest", session.calls[1]["headers"])
        self.assertEqual(session.calls[2]["headers"]["X-RequestDigest"], "test-digest")
        self.assertEqual(session.calls[2]["data"], body)
        for call in session.calls:
            self.assertEqual(call["timeout"], (30, 300))
            self.assertFalse(call["allow_redirects"])
            self.assertEqual(call["headers"]["Accept"], "application/json;odata=verbose")
            self.assertEqual(call["headers"]["X-FORMS_BASED_AUTH_ACCEPTED"], "f")

    def test_http_suppresses_only_allowed_get_404(self):
        session = Session(lambda call: Response(status=404))
        sp = mirror.SharePoint(self.config, session=session)
        self.assertIsNone(sp._request("GET", SITE + "/_api/missing", allow_missing=True))
        for method, allowed in (("GET", False), ("POST", True)):
            with self.subTest(method=method, allowed=allowed), self.assertRaises(mirror.MirrorError):
                sp._request(method, SITE + "/_api/missing", allow_missing=allowed)
        session.responder = lambda call: Response(status=403)
        with self.assertRaises(mirror.MirrorError):
            sp._request("GET", SITE + "/_api/forbidden", allow_missing=True)
        session.responder = mock.Mock(side_effect=OSError("Mock transport failure"))
        with self.assertRaises((mirror.MirrorError, OSError)):
            sp._request("GET", SITE + "/_api/unavailable", allow_missing=True)

    def test_initialize_rejects_wrong_target_guid_and_path(self):
        session = Session()
        sp = mirror.SharePoint(self.config, session=session)
        sp.initialize()
        self.assertEqual(sp.root, ROOT)
        for field, value in (("UniqueId", str(OTHER_ID)), ("ServerRelativeUrl", "/sites/test/Documents/Other")):
            def respond(call, field=field, value=value):
                result = Session.default_response(call)
                if "/GetFolderByServerRelativeUrl(" in call["url"]:
                    result.body["d"][field] = value
                return result

            session.responder = respond
            with self.subTest(field=field), self.assertRaises(mirror.MirrorError):
                sp.initialize()
        self.assertTrue(all(call["method"] == "GET" for call in session.calls))

    def test_paging_follows_valid_continuation(self):
        first = SITE + "/_api/items"
        second = SITE + "/_api/items?next=2"
        pages = {first: {"d": {"results": [{"Name": "A"}], "__next": second}}, second: {"d": {"results": [{"Name": "B"}]}}}
        session = Session(lambda call: Response(pages[call["url"]]))
        sp = mirror.SharePoint(self.config, session=session)
        self.assertEqual(sp.pages(first), [{"Name": "A"}, {"Name": "B"}])
        self.assertEqual([call["url"] for call in session.calls], [first, second])

    def test_paging_rejects_bad_links_loops_and_incomplete_pages(self):
        first = SITE + "/_api/items"
        for next_url in ("https://foreign.example.invalid/_api/items", SITE + "/ordinary/page", first):
            with self.subTest(next_url=next_url):
                def respond(call):
                    if len(session.calls) > 3:
                        self.fail("Paging loop was not detected.")
                    return Response({"d": {"results": [], "__next": next_url}})

                session = Session(respond)
                sp = mirror.SharePoint(self.config, session=session)
                with self.assertRaises(mirror.MirrorError):
                    sp.pages(first)
                self.assertEqual(len(session.calls), 1, "Unsafe continuation must not be requested.")
        session = Session(lambda call: Response({"d": {"results": [{}] * 5000}}))
        with self.assertRaises(mirror.MirrorError):
            mirror.SharePoint(self.config, session=session).pages(first)

    def test_recycle_removes_only_extras_and_children_before_parents(self):
        session = Session()
        sp = mirror.SharePoint(self.config, session=session)
        sp.initialize()
        plan = mirror.build_plan([self.source_file()])
        remote = mirror.RemoteInventory(
            [self.remote_item("Source/A.pdf"), self.remote_item("Old/Deep/A.pdf")],
            [self.remote_item("Source"), self.remote_item("Old"), self.remote_item("Old/Deep")],
        )
        mirror.validate_remote(ROOT, remote, plan)
        self.assertEqual(sp.recycle(remote, plan), (1, 2))
        recycled = [call["url"] for call in session.calls if call["url"].endswith("/recycle()")]
        self.assertEqual(len(recycled), 3)
        self.assertTrue(all("Source" not in url for url in recycled))
        folder_calls = [url for url in recycled if "GetFolderByServerRelativeUrl" in url]
        self.assertIn("/Old/Deep", folder_calls[0])
        self.assertNotIn("/Old/Deep", folder_calls[1])
        self.assertTrue(all(ROOT + "/" in url for url in recycled))

    def test_recycle_stops_on_failed_request(self):
        def respond(call):
            if call["url"].endswith("/recycle()"):
                return Response(status=500)
            return Session.default_response(call)

        session = Session(respond)
        sp = mirror.SharePoint(self.config, session=session)
        sp.initialize()
        remote = mirror.RemoteInventory([self.remote_item("A.pdf"), self.remote_item("B.pdf")], [])
        with self.assertRaises(mirror.MirrorError):
            sp.recycle(remote, mirror.build_plan([]))
        self.assertEqual(sum(call["url"].endswith("/recycle()") for call in session.calls), 1)

    def test_conversion_exports_all_pages_and_closes_every_com_object(self):
        for name in ("A.vsdx", "B.vsdm"):
            (self.source / name).write_bytes(b"mock Visio input")
        sources = mirror.scan_source(self.source)
        workdir = self.base / "converted"
        workdir.mkdir()
        documents = [self.exporting_document(), self.exporting_document()]
        modules, visio = self.com_modules(documents)
        with mock.patch.dict(sys.modules, modules):
            artifacts = mirror.convert_to_pdf(sources, workdir)
        self.assertEqual({item.target for item in artifacts}, {item.target for item in sources})
        self.assertTrue(all(item.pdf.read_bytes().startswith(b"%PDF-") for item in artifacts))
        self.assertEqual(visio.Documents.OpenEx.call_count, 2)
        for call in visio.Documents.OpenEx.call_args_list:
            self.assertEqual(call.args[1], 458)
        for document, artifact in zip(documents, artifacts):
            document.ExportAsFixedFormat.assert_called_once_with(1, str(artifact.pdf), 1, 0)
            document.Close.assert_called_once_with()
        modules["pythoncom"].CoInitializeEx.assert_called_once_with(2)
        modules["win32com.client"].DispatchEx.assert_called_once_with("Visio.Application", clsctx=4)
        visio.Quit.assert_called_once_with()
        modules["pythoncom"].CoUninitialize.assert_called_once_with()

    def test_export_failure_cleans_up_com_and_prevents_sharepoint_writes(self):
        for name in ("A.vsdx", "B.vsdm"):
            (self.source / name).write_bytes(b"mock Visio input")
        sources = mirror.scan_source(self.source)
        documents = [self.exporting_document(), self.exporting_document()]
        documents[1].ExportAsFixedFormat.side_effect = RuntimeError("Mock export failure")
        modules, visio = self.com_modules(documents)
        sp = FakeSharePoint(self.complete_remote())
        with mock.patch.dict(sys.modules, modules), self.assertRaises(mirror.MirrorError):
            mirror.run_mirror(self.config, sp, scanner=lambda root: sources, converter=mirror.convert_to_pdf)
        self.assertFalse({"ensure_folders", "upload", "recycle"} & set(sp.events))
        for document in documents:
            document.Close.assert_called_once_with()
        visio.Quit.assert_called_once_with()
        modules["pythoncom"].CoUninitialize.assert_called_once_with()

    def test_run_success_checks_and_recycles_after_upload(self):
        sources = [self.source_file()]
        sp = FakeSharePoint(self.complete_remote())
        result = mirror.run_mirror(self.config, sp, scanner=lambda root: sources, converter=self.converter)
        self.assertEqual(result, (0, 0))
        self.assertLess(sp.events.index("upload"), sp.events.index("recycle"))
        self.assertLess(sp.events.index("inventory"), sp.events.index("recycle"))

    def test_run_empty_source_recycles_children(self):
        remote = mirror.RemoteInventory([self.remote_item("Old/A.pdf")], [self.remote_item("Old")])
        sp = FakeSharePoint(remote)
        self.assertEqual(mirror.run_mirror(self.config, sp, scanner=lambda root: [], converter=self.converter), (1, 1))
        self.assertIn("recycle", sp.events)

    def test_run_conversion_failure_prevents_remote_writes_and_recycle(self):
        sp = FakeSharePoint(self.complete_remote())
        converter = mock.Mock(side_effect=mirror.MirrorError("Mock conversion failure"))
        with self.assertRaises(mirror.MirrorError):
            mirror.run_mirror(self.config, sp, scanner=lambda root: [self.source_file()], converter=converter)
        self.assertFalse({"ensure_folders", "upload", "recycle"} & set(sp.events))

    def test_run_incomplete_artifacts_prevent_remote_writes(self):
        for artifacts in ([], [mirror.Artifact("Source/Wrong.pdf", self.base / "wrong.pdf")]):
            with self.subTest(artifacts=artifacts):
                sp = FakeSharePoint(self.complete_remote())
                with self.assertRaises(mirror.MirrorError):
                    mirror.run_mirror(self.config, sp, scanner=lambda root: [self.source_file()], converter=lambda sources, workdir: artifacts)
                self.assertFalse({"ensure_folders", "upload", "recycle"} & set(sp.events))

    def test_run_source_change_before_upload_prevents_remote_writes(self):
        sp = FakeSharePoint(self.complete_remote())
        scanner = mock.Mock(side_effect=[[self.source_file()], [self.source_file(size=11)]])
        with self.assertRaises(mirror.MirrorError):
            mirror.run_mirror(self.config, sp, scanner=scanner, converter=self.converter)
        self.assertFalse({"ensure_folders", "upload", "recycle"} & set(sp.events))

    def test_run_source_change_after_upload_prevents_recycle(self):
        for changed_scan in (3, 4):
            with self.subTest(changed_scan=changed_scan):
                sp = FakeSharePoint(self.complete_remote())
                scans = 0

                def scanner(root):
                    nonlocal scans
                    scans += 1
                    return [self.source_file(mtime=100 if scans < changed_scan else 101)]

                with self.assertRaises(mirror.MirrorError):
                    mirror.run_mirror(self.config, sp, scanner=scanner, converter=self.converter)
                self.assertIn("upload", sp.events)
                self.assertNotIn("recycle", sp.events)

    def test_run_upload_failure_prevents_recycle(self):
        sp = FakeSharePoint(self.complete_remote(), upload_error=True)
        with self.assertRaises(mirror.MirrorError):
            mirror.run_mirror(self.config, sp, scanner=lambda root: [self.source_file()], converter=self.converter)
        self.assertIn("upload", sp.events)
        self.assertNotIn("recycle", sp.events)

    def test_run_missing_uploaded_pdf_prevents_recycle(self):
        sp = FakeSharePoint(mirror.RemoteInventory([], [self.remote_item("Source")]))
        with self.assertRaises(mirror.MirrorError):
            mirror.run_mirror(self.config, sp, scanner=lambda root: [self.source_file()], converter=self.converter)
        self.assertIn("upload", sp.events)
        self.assertNotIn("recycle", sp.events)


if __name__ == "__main__":
    unittest.main(verbosity=2)
