#!/usr/bin/env python3
"""Internal upload API for the Dockerized hanger worker."""

from __future__ import annotations

import json
import os
import sys
import tempfile
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


@dataclass(frozen=True)
class Upload:
    filename: str
    content: bytes


def safe_upload_name(raw_name: Any, field_name: str) -> str:
    if not isinstance(raw_name, str) or not raw_name.strip():
        raise AutomationError(f"{field_name} has no filename")
    # Browsers can send either POSIX names or Windows-style fake paths.
    filename = raw_name.replace("\\", "/").rsplit("/", 1)[-1].strip()
    if filename in {"", ".", ".."} or len(filename) > 240:
        raise AutomationError(f"{field_name} has an invalid filename")
    if Path(filename).suffix.casefold() not in {".xls", ".xlsx"}:
        raise AutomationError(f"{field_name} must be an .xls or .xlsx workbook")
    return filename


def validate_workbook_bytes(upload: Upload, field_name: str) -> None:
    suffix = Path(upload.filename).suffix.casefold()
    if not upload.content:
        raise AutomationError(f"{field_name} is empty")
    if suffix == ".xlsx" and not upload.content.startswith(b"PK"):
        raise AutomationError(f"{field_name} is not a valid .xlsx file")
    if suffix == ".xls" and not upload.content.startswith(XLS_MAGIC):
        raise AutomationError(f"{field_name} is not a valid legacy .xls file")


def parse_multipart(body: bytes, content_type: str) -> tuple[dict[str, Upload], dict[str, str]]:
    if not content_type.casefold().startswith("multipart/form-data"):
        raise AutomationError("Content-Type must be multipart/form-data")
    header = (
        f"Content-Type: {content_type}\r\n"
        "MIME-Version: 1.0\r\n\r\n"
    ).encode("utf-8")
    message = BytesParser(policy=policy.default).parsebytes(header + body)
    if not message.is_multipart():
        raise AutomationError("Invalid multipart request")
    files: dict[str, Upload] = {}
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
        if field_name in files:
            raise AutomationError(f"Only one file is allowed for {field_name}")
        files[field_name] = Upload(safe_upload_name(filename, field_name), content)
    return files, fields


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
            missing = [name for name in ("order_workbook", "sof_workbook") if name not in files]
            if missing:
                raise AutomationError("Missing uploaded file field(s): " + ", ".join(missing))
            order_upload = files["order_workbook"]
            sof_upload = files["sof_workbook"]
            validate_workbook_bytes(order_upload, "order_workbook")
            validate_workbook_bytes(sof_upload, "sof_workbook")
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
                sof_path = sof_dir / sof_upload.filename
                order_path.write_bytes(order_upload.content)
                sof_path.write_bytes(sof_upload.content)
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
            audit["uploaded_sof_file"] = sof_upload.filename
            self.send_json(HTTPStatus.OK, audit)
        except (AutomationError, json.JSONDecodeError, ValueError) as exc:
            self.send_json(HTTPStatus.BAD_REQUEST, {"status": "ERROR", "error": str(exc)})
        except Exception as exc:
            self.send_json(HTTPStatus.INTERNAL_SERVER_ERROR, {"status": "ERROR", "error": str(exc)})


class WorkerServer(ThreadingHTTPServer):
    def __init__(self, address, handler, config_path: Path, max_upload_bytes: int):
        super().__init__(address, handler)
        self.config_path = config_path
        self.max_upload_bytes = max_upload_bytes


def main() -> None:
    port = int(os.environ.get("HANGER_API_PORT", "8080"))
    config_path = Path(os.environ.get("HANGER_CONFIG", "/app/config.docker.json"))
    max_upload_mb = int(os.environ.get("HANGER_MAX_UPLOAD_MB", "100"))
    if not config_path.is_file():
        raise SystemExit(f"Configuration file does not exist: {config_path}")
    if max_upload_mb < 1 or max_upload_mb > 500:
        raise SystemExit("HANGER_MAX_UPLOAD_MB must be between 1 and 500")
    server = WorkerServer(
        ("0.0.0.0", port), Handler, config_path, max_upload_mb * 1024 * 1024
    )
    print(f"Hanger upload worker listening on port {port}", flush=True)
    server.serve_forever()


if __name__ == "__main__":
    main()
