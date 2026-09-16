#!/usr/bin/env python3
"""Internal upload API for the Dockerized hanger worker."""

from __future__ import annotations

import io
import json
import os
import re
import stat
import sys
import tempfile
import zipfile
from dataclasses import dataclass
from email import policy
from email.parser import BytesParser
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any

from deepseek_classifier import DeepSeekClassifier
from hanger_automation import AutomationError, load_config, run


XLS_MAGIC = bytes.fromhex("D0CF11E0A1B11AE1")
SOF_FIELD_RE = re.compile(r"^sof_workbook(?:_\d+)?$")


@dataclass(frozen=True)
class Upload:
    filename: str
    content: bytes


def safe_file_name(raw_name: Any, field_name: str) -> str:
    if not isinstance(raw_name, str) or not raw_name.strip():
        raise AutomationError(f"{field_name} has no filename")
    # Browsers can send either POSIX names or Windows-style fake paths.
    filename = raw_name.replace("\\", "/").rsplit("/", 1)[-1].strip()
    if filename in {"", ".", ".."} or len(filename) > 240:
        raise AutomationError(f"{field_name} has an invalid filename")
    return filename


# Orders and SOFs accept the same formats. Word/PDF is read for the text it
# states; a scan without a text layer is rejected rather than guessed at.
UPLOAD_SUFFIXES = {".xls", ".xlsx", ".docx", ".pdf"}


def safe_upload_name(raw_name: Any, field_name: str) -> str:
    filename = safe_file_name(raw_name, field_name)
    if Path(filename).suffix.casefold() not in UPLOAD_SUFFIXES:
        raise AutomationError(f"{field_name} must be {', '.join(sorted(UPLOAD_SUFFIXES))}")
    return filename


def validate_workbook_bytes(upload: Upload, field_name: str) -> None:
    suffix = Path(upload.filename).suffix.casefold()
    if not upload.content:
        raise AutomationError(f"{field_name} is empty")
    if suffix == ".xlsx" and not upload.content.startswith(b"PK"):
        raise AutomationError(f"{field_name} is not a valid .xlsx file")
    if suffix == ".xls" and not upload.content.startswith(XLS_MAGIC):
        raise AutomationError(f"{field_name} is not a valid legacy .xls file")


def validate_upload_bytes(upload: Upload, field_name: str = "sof_workbook") -> None:
    suffix = Path(upload.filename).suffix.casefold()
    if not upload.content:
        raise AutomationError(f"{field_name} is empty")
    if suffix in {".xls", ".xlsx"}:
        validate_workbook_bytes(upload, field_name)
    elif suffix == ".docx":
        if not upload.content.startswith(b"PK"):
            raise AutomationError(f"{field_name} is not a valid .docx file")
        try:
            with zipfile.ZipFile(io.BytesIO(upload.content)) as archive:
                members = archive.infolist()
                if "word/document.xml" not in archive.namelist():
                    raise AutomationError(f"{field_name} is not a Word .docx document")
                if len(members) > 3000 or sum(item.file_size for item in members) > 100 * 1024 * 1024:
                    raise AutomationError(f"{field_name} .docx expands beyond the safe limit")
                if any(item.flag_bits & 0x1 for item in members):
                    raise AutomationError("Password-protected .docx files are not supported")
        except zipfile.BadZipFile as exc:
            raise AutomationError(f"{field_name} is not a valid .docx file") from exc
    elif suffix == ".pdf" and not upload.content.startswith(b"%PDF-"):
        raise AutomationError(f"{field_name} is not a valid .pdf file")


def validate_sof_bytes(upload: Upload) -> None:
    validate_upload_bytes(upload, "sof_workbook")


def parse_multipart(
    body: bytes, content_type: str
) -> tuple[dict[str, list[Upload]], dict[str, str]]:
    if not content_type.casefold().startswith("multipart/form-data"):
        raise AutomationError("Content-Type must be multipart/form-data")
    header = (
        f"Content-Type: {content_type}\r\n"
        "MIME-Version: 1.0\r\n\r\n"
    ).encode("utf-8")
    message = BytesParser(policy=policy.default).parsebytes(header + body)
    if not message.is_multipart():
        raise AutomationError("Invalid multipart request")
    files: dict[str, list[Upload]] = {}
    fields: dict[str, str] = {}
    for part in message.iter_parts():
        if part.get_content_disposition() != "form-data":
            continue
        field_name = part.get_param("name", header="content-disposition")
        if not field_name:
            continue
        filename = part.get_filename()
        content = part.get_payload(decode=True) or b""
        if filename is None:
            fields[field_name] = content.decode(part.get_content_charset() or "utf-8", errors="replace")
            continue
        files.setdefault(field_name, []).append(
            Upload(safe_file_name(filename, field_name), content)
        )
    return files, fields


def unpack_sof_archive(
    upload: Upload, max_sof_files: int, max_uncompressed_bytes: int
) -> list[Upload]:
    """Read the n8n-generated SOF ZIP without extracting untrusted paths."""
    if Path(upload.filename).suffix.casefold() != ".zip" or not upload.content.startswith(b"PK"):
        raise AutomationError("sof_archive must be a valid .zip file")
    try:
        archive = zipfile.ZipFile(io.BytesIO(upload.content))
    except zipfile.BadZipFile as exc:
        raise AutomationError("sof_archive is not a valid .zip file") from exc

    with archive:
        members = [member for member in archive.infolist() if not member.is_dir()]
        if not members:
            raise AutomationError("sof_archive contains no SOF files")
        if len(members) > max_sof_files:
            raise AutomationError(
                f"At most {max_sof_files} SOF files may be uploaded"
            )
        declared_size = sum(member.file_size for member in members)
        if declared_size > max_uncompressed_bytes:
            raise AutomationError("Uncompressed SOF files exceed the upload limit")

        uploads: list[Upload] = []
        actual_size = 0
        for member in members:
            if member.flag_bits & 0x1:
                raise AutomationError("Password-protected SOF archives are not supported")
            mode = (member.external_attr >> 16) & 0xFFFF
            if mode and stat.S_ISLNK(mode):
                raise AutomationError("SOF archives may not contain symbolic links")
            filename = safe_upload_name(member.filename, "sof_archive")
            content = archive.read(member)
            actual_size += len(content)
            if actual_size > max_uncompressed_bytes:
                raise AutomationError("Uncompressed SOF files exceed the upload limit")
            workbook = Upload(filename, content)
            validate_sof_bytes(workbook)
            uploads.append(workbook)
        return uploads


def collect_uploaded_workbooks(
    files: dict[str, list[Upload]], max_sof_files: int, max_upload_bytes: int
) -> tuple[Upload, list[Upload]]:
    """Normalize direct n8n fields or the ZIP produced by the workflow."""
    unexpected = [
        name
        for name in files
        if name != "order_workbook"
        and name != "sof_archive"
        and not SOF_FIELD_RE.fullmatch(name)
    ]
    if unexpected:
        raise AutomationError("Unexpected uploaded file field(s): " + ", ".join(unexpected))

    order_uploads = files.get("order_workbook", [])
    if len(order_uploads) != 1:
        raise AutomationError("Exactly one order_workbook must be uploaded")
    order_upload = Upload(
        safe_upload_name(order_uploads[0].filename, "order_workbook"),
        order_uploads[0].content,
    )
    validate_upload_bytes(order_upload, "order_workbook")

    direct_sofs = [
        upload
        for field_name, uploads in files.items()
        if SOF_FIELD_RE.fullmatch(field_name)
        for upload in uploads
    ]
    archives = files.get("sof_archive", [])
    if archives and direct_sofs:
        raise AutomationError("Upload SOFs directly or as sof_archive, not both")
    if len(archives) > 1:
        raise AutomationError("Only one sof_archive is allowed")
    if archives:
        sofs = unpack_sof_archive(archives[0], max_sof_files, max_upload_bytes)
    else:
        sofs = []
        for upload in direct_sofs:
            workbook = Upload(
                safe_upload_name(upload.filename, "sof_workbook"), upload.content
            )
            validate_sof_bytes(workbook)
            sofs.append(workbook)

    if not sofs:
        raise AutomationError("At least one SOF file must be uploaded")
    if len(sofs) > max_sof_files:
        raise AutomationError(f"At most {max_sof_files} SOF files may be uploaded")
    duplicate_names = sorted({
        upload.filename
        for upload in sofs
        if sum(other.filename.casefold() == upload.filename.casefold() for other in sofs) > 1
    })
    if duplicate_names:
        raise AutomationError("Duplicate SOF filename(s): " + ", ".join(duplicate_names))
    return order_upload, sofs


class Handler(BaseHTTPRequestHandler):
    server_version = "HangerWorker/2.0"

    def log_message(self, fmt: str, *args: Any) -> None:
        sys.stderr.write("%s - %s\n" % (self.address_string(), fmt % args))

    def send_json(self, status: HTTPStatus, payload: dict[str, Any]) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def send_result_file(self, query: str) -> None:
        """Serve a generated result workbook by name from the output directory."""
        from urllib.parse import parse_qs

        requested = (parse_qs(query).get("file") or [""])[0]
        # Only a bare filename is accepted, so a crafted name cannot escape the
        # output directory.
        if not requested or requested != Path(requested).name or requested.startswith("."):
            self.send_json(HTTPStatus.BAD_REQUEST,
                           {"status": "ERROR", "error": "Invalid result file name"})
            return
        config = load_config(self.server.config_path)
        target = (Path(config["output_directory"]) / requested).resolve()
        output_root = Path(config["output_directory"]).resolve()
        if output_root not in target.parents or not target.is_file():
            self.send_json(HTTPStatus.NOT_FOUND,
                           {"status": "ERROR", "error": "Result file not found"})
            return
        payload = target.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header(
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        self.send_header("Content-Disposition", f'attachment; filename="{target.name}"')
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def do_GET(self) -> None:  # noqa: N802
        if self.path.startswith("/result?"):
            self.send_result_file(self.path.split("?", 1)[1])
        elif self.path == "/health":
            self.send_json(HTTPStatus.OK, {
                "status": "ok",
                "mode": "direct-upload",
                "max_sof_files": self.server.max_sof_files,
                "llm_enabled": os.environ.get("HANGER_LLM_ENABLED", "false").casefold() == "true",
            })
        else:
            self.send_json(HTTPStatus.NOT_FOUND, {"status": "ERROR", "error": "Not found"})

    def do_POST(self) -> None:  # noqa: N802
        if self.path != "/run-upload":
            self.send_json(HTTPStatus.NOT_FOUND, {"status": "ERROR", "error": "Not found"})
            return
        try:
            length = int(self.headers.get("Content-Length", "0"))
            if length <= 0 or length > self.server.max_upload_bytes:
                limit_mb = self.server.max_upload_bytes // (1024 * 1024)
                raise AutomationError(f"Upload is empty or larger than {limit_mb} MiB")
            files, fields = parse_multipart(
                self.rfile.read(length), self.headers.get("Content-Type", "")
            )
            order_upload, sof_uploads = collect_uploaded_workbooks(
                files, self.server.max_sof_files, self.server.max_upload_bytes
            )
            overwrite_text = fields.get("overwrite", "false").strip().casefold()
            if overwrite_text not in {"true", "false"}:
                raise AutomationError("overwrite must be true or false")
            with tempfile.TemporaryDirectory(prefix="hanger_upload_") as temp_name:
                temp_root = Path(temp_name)
                order_dir = temp_root / "order"
                sof_dir = temp_root / "sof"
                order_dir.mkdir()
                sof_dir.mkdir()
                order_path = order_dir / order_upload.filename
                order_path.write_bytes(order_upload.content)
                for sof_upload in sof_uploads:
                    (sof_dir / sof_upload.filename).write_bytes(sof_upload.content)
                config = load_config(self.server.config_path)
                llm_classifier = DeepSeekClassifier.from_env()
                audit = run(
                    input_path=order_path,
                    sof_root=sof_dir,
                    rules_path=Path(config["rules_file"]),
                    output_dir=Path(config["output_directory"]),
                    overwrite=overwrite_text == "true"
                    or bool(config.get("overwrite_existing_hanger_data", False)),
                    order_sheet=config.get("order_sheet"),
                    header_scan_rows=int(config.get("header_scan_rows", 50)),
                    llm_classifier=llm_classifier,
                )
            audit["input_file"] = order_upload.filename
            audit["uploaded_sof_files"] = [upload.filename for upload in sof_uploads]
            audit["uploaded_sof_count"] = len(sof_uploads)
            # Retain the original field for consumers created before multi-upload support.
            audit["uploaded_sof_file"] = sof_uploads[0].filename
            audit_path = Path(str(audit.get("audit_file", "")))
            if audit_path.is_file():
                audit_path.write_text(
                    json.dumps(audit, ensure_ascii=False, indent=2), encoding="utf-8"
                )
            self.send_json(HTTPStatus.OK, audit)
        except (AutomationError, json.JSONDecodeError, ValueError) as exc:
            self.send_json(HTTPStatus.BAD_REQUEST, {"status": "ERROR", "error": str(exc)})
        except Exception as exc:
            self.send_json(HTTPStatus.INTERNAL_SERVER_ERROR, {"status": "ERROR", "error": str(exc)})


class WorkerServer(ThreadingHTTPServer):
    def __init__(
        self, address, handler, config_path: Path,
        max_upload_bytes: int, max_sof_files: int,
    ):
        super().__init__(address, handler)
        self.config_path = config_path
        self.max_upload_bytes = max_upload_bytes
        self.max_sof_files = max_sof_files


def main() -> None:
    port = int(os.environ.get("HANGER_API_PORT", "8080"))
    config_path = Path(os.environ.get("HANGER_CONFIG", "/app/config.docker.json"))
    max_upload_mb = int(os.environ.get("HANGER_MAX_UPLOAD_MB", "100"))
    max_sof_files = int(os.environ.get("HANGER_MAX_SOF_FILES", "20"))
    if not config_path.is_file():
        raise SystemExit(f"Configuration file does not exist: {config_path}")
    if max_upload_mb < 1 or max_upload_mb > 500:
        raise SystemExit("HANGER_MAX_UPLOAD_MB must be between 1 and 500")
    if max_sof_files < 1 or max_sof_files > 100:
        raise SystemExit("HANGER_MAX_SOF_FILES must be between 1 and 100")
    server = WorkerServer(
        ("0.0.0.0", port), Handler, config_path,
        max_upload_mb * 1024 * 1024, max_sof_files,
    )
    print(f"Hanger upload worker listening on port {port}", flush=True)
    server.serve_forever()


if __name__ == "__main__":
    main()
