"""Visio-PDF-Vollspiegel fuer SharePoint Server 2019 unter Windows.

Aufruf: python visio_sharepoint_mirror.py --config mirror.json
Die Konsolen-Demo steht separat in demo_mirror.py.
"""

from __future__ import annotations

import argparse
from contextlib import contextmanager
from dataclasses import dataclass
import json
import logging
import os
from pathlib import Path, PureWindowsPath
import re
import shutil
import stat
import sys
import tempfile
import unicodedata
from urllib.parse import quote, unquote, urljoin, urlsplit
from uuid import UUID, uuid4

LOG = logging.getLogger("mirror")
EXTENSIONS = {".vsd", ".vsdx", ".vsdm"}
HEADERS = {"Accept": "application/json;odata=verbose", "X-FORMS_BASED_AUTH_ACCEPTED": "f"}


class MirrorError(RuntimeError):
    """Ein sicherer, vollstaendiger Spiegel-Lauf ist nicht moeglich."""


def path_key(path: str) -> str:
    return unicodedata.normalize("NFC", path.replace("\\", "/").strip("/")).lower()


def safe_segment(name: str) -> None:
    if (not name or name != name.strip() or name in {".", ".."}
            or name.endswith(".") or len(name) > 128
            or any(c in '~"#%&*:<>?/\\{|}[]' or unicodedata.category(c) == "Cc" for c in name)
            or re.match(r"^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)|^_vti_", name, re.I)):
        raise MirrorError(f"Nicht unterstuetzter SharePoint-Name: {name!r}")


def safe_relative(path: str) -> str:
    path = path.replace("\\", "/")
    for segment in path.split("/"):
        safe_segment(segment)
    return path


def join_url(root: str, relative: str) -> str:
    return root.rstrip("/") + "/" + relative


def odata(value: str) -> str:
    # OData maskiert Apostrophe, URL-Encoding schuetzt die restliche URL-Syntax.
    return quote(value.replace("'", "''"), safe="/")


@dataclass(frozen=True)
class Config:
    source: Path
    site_url: str
    library_name: str
    target_folder: str
    target_id: UUID
    log_path: Path

    @classmethod
    def load(cls, path: Path) -> Config:
        raw = json.loads(Path(path).read_text(encoding="utf-8-sig"))
        fields = {"SourcePath", "SharePointSiteUrl", "LibraryName", "TargetFolderPath", "TargetFolderUniqueId", "LogPath"}
        if not isinstance(raw, dict) or set(raw) != fields:
            raise MirrorError("Die Konfiguration muss genau die sechs dokumentierten Werte enthalten.")
        for name, value in raw.items():
            if not isinstance(value, str) or not value.strip() or value.startswith("__PLATZHALTER_"):
                raise MirrorError(f"Konfigurationswert {name} fehlt, ist leer oder noch ein Platzhalter.")
        source, log = (PureWindowsPath(raw[name]) for name in ("SourcePath", "LogPath"))
        for value in (source, log):
            if not value.is_absolute() or not value.name or str(value).startswith(("\\\\?\\", "\\\\.\\")):
                raise MirrorError("Quelle und Logdatei brauchen absolute Windows-Pfade unterhalb einer Wurzel.")
        site = urlsplit(raw["SharePointSiteUrl"])
        if (site.scheme != "https" or not site.hostname or site.username is not None
                or site.password is not None or site.query or site.fragment
                or any(p in {".", ".."} for p in unquote(site.path).split("/"))):
            raise MirrorError("SharePointSiteUrl muss eine HTTPS-Site ohne Zugangsdaten, Query oder Fragment sein.")
        site.port  # Ungueltige Portangaben vor dem ersten Netzwerkzugriff ablehnen.
        library = raw["LibraryName"].strip()
        safe_segment(library)
        target = safe_relative(raw["TargetFolderPath"])
        target_id = UUID(raw["TargetFolderUniqueId"])
        if not target_id.int:
            raise MirrorError("TargetFolderUniqueId darf keine leere GUID sein.")
        return cls(Path(source), raw["SharePointSiteUrl"].rstrip("/"), library, target, target_id, Path(log))


@dataclass(frozen=True)
class SourceFile:
    source: Path
    relative: str
    target: str
    size: int
    mtime_ns: int


@dataclass(frozen=True)
class Artifact:
    target: str
    pdf: Path


@dataclass(frozen=True)
class Plan:
    file_keys: set[str]
    folder_keys: set[str]
    folders: tuple[str, ...]


@dataclass(frozen=True)
class RemoteItem:
    relative: str
    url: str


@dataclass(frozen=True)
class RemoteInventory:
    files: list[RemoteItem]
    folders: list[RemoteItem]


def is_reparse(info: os.stat_result) -> bool:
    return stat.S_ISLNK(info.st_mode) or bool(getattr(info, "st_file_attributes", 0) & 0x400)


def scan_source(root: Path) -> list[SourceFile]:
    root = Path(root)
    if not root.is_absolute() or root == Path(root.anchor):
        raise MirrorError("Die Quelle muss ein absoluter Ordner unterhalb einer Wurzel sein.")
    if is_reparse(root.lstat()) or not root.is_dir():
        raise MirrorError("Die Quelle muss ein erreichbarer Ordner ohne Reparse Point sein.")
    safe_segment(root.name)
    stack, files, owners = [root], [], set()
    while stack:
        # scandir/lstat-Fehler brechen ab; ein unvollstaendiger Scan darf nie loeschen.
        with os.scandir(stack.pop()) as entries:
            for entry in entries:
                info = entry.stat(follow_symlinks=False)
                if is_reparse(info):
                    raise MirrorError(f"Reparse Point in der Quelle: {entry.path}")
                path = Path(entry.path)
                if stat.S_ISDIR(info.st_mode):
                    stack.append(path)
                elif stat.S_ISREG(info.st_mode) and path.suffix.lower() in EXTENSIONS and not path.name.startswith("~$"):
                    relative = safe_relative(path.relative_to(root).as_posix())
                    target = safe_relative(f"{root.name}/{Path(relative).with_suffix('.pdf').as_posix()}")
                    key = path_key(target)
                    if key in owners:
                        raise MirrorError(f"Mehrere Quellen erzeugen denselben PDF-Zielpfad: {target}")
                    owners.add(key)
                    files.append(SourceFile(path, relative, target, info.st_size, info.st_mtime_ns))
    return sorted(files, key=lambda item: path_key(item.relative))


def build_plan(sources: list[SourceFile]) -> Plan:
    files, folders = set(), {}
    for item in sources:
        target = safe_relative(item.target)
        key = path_key(target)
        if key in files:
            raise MirrorError(f"Mehrere Quellen erzeugen denselben PDF-Zielpfad: {target}")
        files.add(key)
        parts = target.split("/")
        for depth in range(1, len(parts)):
            folder = "/".join(parts[:depth])
            key = path_key(folder)
            if key in folders and folders[key] != folder:
                raise MirrorError(f"Mehrdeutiger Zielordner: {folder}")
            folders[key] = folder
    if files.intersection(folders):
        raise MirrorError("Ein PDF-Zielpfad kollidiert mit einem benoetigten Ordner.")
    return Plan(files, set(folders), tuple(sorted(folders.values(), key=lambda p: (p.count("/"), p))))


def validate_remote(root: str, remote: RemoteInventory, plan: Plan) -> None:
    keys = []
    for items in (remote.files, remote.folders):
        found = set()
        for item in items:
            safe_relative(item.relative)
            if (path_key(item.url) != path_key(join_url(root, item.relative))
                    or len(item.url) > 260 or path_key(item.relative) in found):
                raise MirrorError("Die Zielinventur enthaelt einen ungueltigen, fremden oder doppelten Pfad.")
            found.add(path_key(item.relative))
        keys.append(found)
    if keys[0] & keys[1] or not plan.file_keys <= keys[0] or not plan.folder_keys <= keys[1]:
        raise MirrorError("Nach dem Upload fehlen erwartete PDFs/Ordner oder es bestehen Typkonflikte.")


class SharePoint:
    def __init__(self, config: Config, session=None):
        self.config, self.root = config, ""
        if session is None:
            import requests
            from requests_negotiate_sspi import HttpNegotiateAuth

            session = requests.Session()
            session.auth = HttpNegotiateAuth()  # Aktuelles Windows-Konto, kein Kennwort in der Konfiguration.
        self.session = session

    def close(self) -> None:
        self.session.close()

    def api(self, path: str) -> str:
        return self.config.site_url + "/_api/" + path

    def folder_url(self, server_relative: str, suffix: str = "") -> str:
        return self.api(f"web/GetFolderByServerRelativeUrl('{odata(server_relative)}'){suffix}")

    def _check_api_url(self, url: str) -> None:
        site, candidate = urlsplit(self.config.site_url), urlsplit(url)
        if (candidate.scheme != "https" or candidate.hostname != site.hostname
                or (candidate.port or 443) != (site.port or 443)
                or candidate.username is not None or candidate.password is not None or candidate.fragment
                or not unquote(candidate.path).startswith(unquote(site.path).rstrip("/") + "/_api/")
                or any(p in {".", ".."} for p in unquote(candidate.path).split("/"))):
            raise MirrorError("SharePoint lieferte eine unsichere API-/Paging-URL.")

    def _request(self, method: str, url: str, *, write=False, allow_missing=False, data=None, json=None):
        self._check_api_url(url)
        headers = HEADERS.copy()
        if write:
            digest = self._request("POST", self.api("contextinfo"))["d"]["GetContextWebInformation"]["FormDigestValue"]
            if not isinstance(digest, str) or not digest:
                raise MirrorError("SharePoint lieferte keinen FormDigest.")
            headers["X-RequestDigest"] = digest
        if data is not None:
            headers["Content-Type"] = "application/octet-stream"
        elif json is not None:
            headers["Content-Type"] = "application/json;odata=verbose"
        response = self.session.request(method, url, headers=headers, timeout=(30, 300),
                                        allow_redirects=False, data=data, json=json)
        if method == "GET" and allow_missing and response.status_code == 404:
            return None
        if not 200 <= response.status_code < 300:
            raise MirrorError(f"SharePoint {method} fehlgeschlagen: HTTP {response.status_code} ({url})")
        return response.json() if response.content else None

    def initialize(self) -> None:
        library = self._request("GET", self.api(
            f"web/lists/getbytitle('{odata(self.config.library_name)}')?$select=ForceCheckout,RootFolder/ServerRelativeUrl&$expand=RootFolder"))["d"]
        if library["ForceCheckout"]:
            raise MirrorError("Die Zielbibliothek verlangt Auschecken; ForceCheckout=false ist erforderlich.")
        library_root = library["RootFolder"]["ServerRelativeUrl"]
        if not isinstance(library_root, str) or not library_root.startswith("/"):
            raise MirrorError("SharePoint lieferte keinen gueltigen Bibliothekspfad.")
        self.root = join_url(library_root, self.config.target_folder)
        self.assert_target()

    def assert_target(self) -> None:
        folder = self._request("GET", self.folder_url(self.root, "?$select=ServerRelativeUrl,UniqueId"))["d"]
        if UUID(folder["UniqueId"]) != self.config.target_id or path_key(folder["ServerRelativeUrl"]) != path_key(self.root):
            raise MirrorError("Der SharePoint-Zielordner stimmt nicht mit Pfad und TargetFolderUniqueId ueberein.")

    def ensure_folders(self, plan: Plan) -> None:
        for folder in plan.folders:
            url = join_url(self.root, folder)
            if self._request("GET", self.folder_url(url), allow_missing=True) is None:
                self.assert_target()
                self._request("POST", self.api("web/folders"), write=True,
                              json={"__metadata": {"type": "SP.Folder"}, "ServerRelativeUrl": url})
                self._request("GET", self.folder_url(url))

    def upload(self, artifacts: list[Artifact]) -> None:
        for artifact in artifacts:
            parent, name = artifact.target.rsplit("/", 1)
            url = self.folder_url(join_url(self.root, parent), f"/Files/add(url='{odata(name)}',overwrite=true)")
            self.assert_target()
            # Bytes bleiben auch ueber mehrere NTLM-Challenges vollstaendig wiederholbar.
            self._request("POST", url, write=True, data=artifact.pdf.read_bytes())

    def pages(self, url: str) -> list[dict]:
        result, visited = [], set()
        while url:
            if url in visited:
                raise MirrorError("SharePoint lieferte einen zyklischen Paging-Link.")
            visited.add(url)
            page = self._request("GET", url)["d"]
            items, next_url = page["results"], page.get("__next")
            if not isinstance(items, list) or (next_url is not None and not isinstance(next_url, str)):
                raise MirrorError("SharePoint lieferte eine ungueltige Inventurseite.")
            if not next_url and len(items) >= 5000:
                raise MirrorError("Moeglicherweise abgeschnittene Inventurseite ohne Fortsetzungslink.")
            result.extend(items)
            url = urljoin(url, next_url) if next_url else None
        return result

    def inventory(self) -> RemoteInventory:
        files, folders, stack, visited = [], [], [(self.root, "")], set()
        while stack:
            current, relative = stack.pop()
            if path_key(current) in visited:
                raise MirrorError("Zyklischer oder doppelter Ordner in der SharePoint-Inventur.")
            visited.add(path_key(current))
            for kind, collection in (("Files", files), ("Folders", folders)):
                for item in self.pages(self.folder_url(current, f"/{kind}?$select=Name,ServerRelativeUrl&$top=5000")):
                    safe_segment(item["Name"])
                    rel = join_url(relative, item["Name"]) if relative else item["Name"]
                    server_url = item["ServerRelativeUrl"]
                    if path_key(server_url) != path_key(join_url(self.root, rel)) or len(server_url) > 260:
                        raise MirrorError("SharePoint lieferte einen Ordner-/Dateipfad ausserhalb des erwarteten Ziels.")
                    collection.append(RemoteItem(rel, server_url))
                    if kind == "Folders":
                        stack.append((server_url, rel))
        return RemoteInventory(files, folders)

    def recycle(self, remote: RemoteInventory, plan: Plan) -> tuple[int, int]:
        validate_remote(self.root, remote, plan)
        counts = []
        for kind, items, expected in (("File", remote.files, plan.file_keys), ("Folder", remote.folders, plan.folder_keys)):
            extras = sorted((item for item in items if path_key(item.relative) not in expected),
                            key=lambda item: (-item.relative.count("/"), item.relative))
            for item in extras:
                self.assert_target()
                self._request("POST", self.api(f"web/Get{kind}ByServerRelativeUrl('{odata(item.url)}')/recycle()"), write=True)
            counts.append(len(extras))
        return counts[0], counts[1]


def convert_to_pdf(sources: list[SourceFile], workdir: Path) -> list[Artifact]:
    if not sources:
        return []
    import pythoncom
    from win32com.client import DispatchEx

    artifacts, visio = [], None
    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    try:
        visio = DispatchEx("Visio.Application", clsctx=pythoncom.CLSCTX_LOCAL_SERVER)
        visio.Visible, visio.AlertResponse = False, 7
        for item in sources:
            copy = workdir / (uuid4().hex + item.source.suffix)
            pdf = workdir / (uuid4().hex + ".pdf")
            shutil.copyfile(item.source, copy)
            document = None
            try:
                # Read-only, DontList, Hidden, MacrosDisabled, NoWorkspace.
                document = visio.Documents.OpenEx(str(copy), 2 | 8 | 64 | 128 | 256)
                document.ExportAsFixedFormat(1, str(pdf), 1, 0)
            except Exception as exc:
                raise MirrorError(f"Visio-Konvertierung fehlgeschlagen: {item.relative}") from exc
            finally:
                if document is not None:
                    try:
                        document.Close()
                    finally:
                        document = None
            with pdf.open("rb") as result:
                if result.read(5) != b"%PDF-":
                    raise MirrorError(f"Visio erzeugte keine gueltige PDF: {item.relative}")
            artifacts.append(Artifact(item.target, pdf))
    finally:
        try:
            if visio is not None:
                visio.Quit()
        finally:
            visio = None
            pythoncom.CoUninitialize()
    return artifacts


def mutex_name(config: Config) -> str:
    # Derselbe Name wie PowerShell, unabhaengig von URL-Alias oder Pfadschreibweise.
    return "Global\\PPSI_VisioSharePointMirror_" + config.target_id.hex.upper()


@contextmanager
def target_lock(config: Config):
    import win32event
    import win32api

    handle, acquired = win32event.CreateMutex(None, False, mutex_name(config)), False
    try:
        acquired = win32event.WaitForSingleObject(handle, 0) in (0, 128)  # OBJECT_0 / ABANDONED
        if not acquired:
            raise MirrorError("Fuer dieses SharePoint-Ziel laeuft bereits ein Spiegel-Lauf.")
        yield
    finally:
        try:
            if acquired:
                win32event.ReleaseMutex(handle)
        finally:
            win32api.CloseHandle(handle)


def run_mirror(config: Config, sp: SharePoint, *, scanner=scan_source, converter=convert_to_pdf) -> tuple[int, int]:
    LOG.info("Spiegel-Lauf gestartet.")
    source = scanner(config.source)
    plan = build_plan(source)
    LOG.info("%s Visio-Datei(en) gefunden.", len(source))
    sp.initialize()
    if any(len(path) > 260 for path in [sp.root] + [join_url(sp.root, p) for p in (*plan.folders, *(s.target for s in source))]):
        raise MirrorError("Mindestens ein SharePoint-Zielpfad ueberschreitet 260 Zeichen.")
    LOG.info("SharePoint-Ziel und Windows-Anmeldung wurden bestaetigt.")
    with tempfile.TemporaryDirectory(prefix="PPSI-VisioSharePointMirror-") as temporary:
        artifacts = converter(source, Path(temporary))
        if len(artifacts) != len(source) or {path_key(a.target) for a in artifacts} != plan.file_keys:
            raise MirrorError("Die Konvertierung lieferte nicht den vollstaendigen PDF-Bestand.")
        LOG.info("%s PDF-Datei(en) erzeugt.", len(artifacts))
        if scanner(config.source) != source:
            raise MirrorError("Die Quelle hat sich waehrend der Konvertierung geaendert; SharePoint blieb unveraendert.")
        sp.assert_target()
        sp.ensure_folders(plan)
        sp.upload(artifacts)
        LOG.info("%s PDF-Datei(en) nach SharePoint hochgeladen.", len(artifacts))
        if scanner(config.source) != source:
            raise MirrorError("Die Quelle hat sich waehrend des Uploads geaendert; Papierkorbaktionen wurden unterdrueckt.")
        sp.assert_target()
        remote = sp.inventory()
        sp.assert_target()
        validate_remote(sp.root, remote, plan)
        if scanner(config.source) != source:
            raise MirrorError("Die Quelle hat sich waehrend der Zielinventur geaendert; Papierkorbaktionen wurden unterdrueckt.")
        sp.assert_target()
        result = sp.recycle(remote, plan)
        LOG.info("Spiegelabgleich abgeschlossen: %s Datei(en) und %s Ordner recycelt.", *result)
    return result


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--config", required=True, type=Path, help="Pfad zur mirror.json")
    args = parser.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s",
                        datefmt="%Y-%m-%d %H:%M:%S", stream=sys.stdout)
    handler = None
    try:
        config = Config.load(args.config)
        if sys.platform != "win32":
            raise MirrorError("Der produktive Lauf benoetigt Windows und lokales Microsoft Visio.")
        config.log_path.parent.mkdir(parents=True, exist_ok=True)
        handler = logging.FileHandler(config.log_path, encoding="utf-8")
        handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S"))
        LOG.addHandler(handler)
        with target_lock(config):
            sp = SharePoint(config)
            try:
                run_mirror(config, sp)
            finally:
                sp.close()
        return 0
    except ImportError as exc:
        LOG.error("Python-Abhaengigkeit fehlt (%s). Installieren mit: python -m pip install -r requirements.txt", exc.name)
    except Exception as exc:
        LOG.error("%s", exc)
    finally:
        if handler is not None:
            LOG.removeHandler(handler)
            handler.close()
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
