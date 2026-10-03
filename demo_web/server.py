#!/usr/bin/env python3
"""Full-featured Live Demo Web Server for Hanger Automation.

Provides:
- Web App UI serving (HTML, CSS, JS, Samples)
- Real file upload handling (.xlsx, .xls, .docx, .pdf)
- Direct pipeline execution via hanger_automation.run()
- Output generation of colored TONG HOP HANGER .xlsx workbooks
- Real-time audit inspect & export endpoints
"""

import io
import json
import mimetypes
import os
import re
import shutil
import sys
import tempfile
import time
import urllib.parse
from datetime import datetime
from email import policy
from email.parser import BytesParser
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

# Add project root and src to path
BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR / "src"))
sys.path.insert(0, str(BASE_DIR))

# Load environment variables from .env
try:
    import dotenv
    dotenv.load_dotenv(BASE_DIR / ".env", override=True)
except ImportError:
    env_file = BASE_DIR / ".env"
    if env_file.exists():
        for line in env_file.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                os.environ[k.strip()] = v.strip()

import threading
import traceback
import uuid

try:
    from hanger_automation import AutomationError, run, build_result_workbook
    from deepseek_classifier import DeepSeekClassifier
except ImportError:
    run = None
    build_result_workbook = None
    DeepSeekClassifier = None

JOBS = {}

WEB_DIR = Path(__file__).resolve().parent
SAMPLES_DIR = WEB_DIR / "samples"
OUTPUTS_DIR = WEB_DIR / "outputs"
SOF_STORAGE = WEB_DIR / "sof_storage"

SAMPLES_DIR.mkdir(parents=True, exist_ok=True)
OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)
SOF_STORAGE.mkdir(parents=True, exist_ok=True)


def parse_multipart(body: bytes, content_type: str):
    header = (
        f"Content-Type: {content_type}\r\n"
        "MIME-Version: 1.0\r\n\r\n"
    ).encode("utf-8")
    message = BytesParser(policy=policy.default).parsebytes(header + body)
    if not message.is_multipart():
        return {}, {}
    files = {}
    fields = {}
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
        else:
            files.setdefault(field_name, []).append({
                "filename": Path(filename).name,
                "content": content
            })
    return files, fields


class DemoHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(WEB_DIR), **kwargs)

    def do_GET(self):
        url_parsed = urllib.parse.urlparse(self.path)
        path = url_parsed.path

        if path == "/api/status":
            self.send_json(HTTPStatus.OK, {
                "status": "ONLINE",
                "engine": "Deterministic + LLM DeepSeek",
                "python": sys.version.split()[0],
                "time": datetime.now().isoformat()
            })
            return

        if path == "/api/job-status":
            query = urllib.parse.parse_qs(url_parsed.query)
            job_id = query.get("job_id", [""])[0]
            if not job_id or job_id not in JOBS:
                self.send_json(HTTPStatus.NOT_FOUND, {"error": "Không tìm thấy tiến trình"})
                return
            self.send_json(HTTPStatus.OK, JOBS[job_id])
            return

        if path == "/api/samples":
            sample_files = []
            for p in SAMPLES_DIR.glob("*"):
                if p.is_file():
                    sample_files.append({
                        "name": p.name,
                        "size_bytes": p.stat().st_size,
                        "size_formatted": f"{p.stat().st_size / 1024:.1f} KB"
                    })
            self.send_json(HTTPStatus.OK, {"samples": sample_files})
            return

        if path.startswith("/api/download/"):
            filename = urllib.parse.unquote(path.replace("/api/download/", "")).strip()
            # Security: avoid directory traversal
            clean_name = Path(filename).name
            file_path = OUTPUTS_DIR / clean_name
            if not file_path.exists():
                file_path = SAMPLES_DIR / clean_name
            if not file_path.exists():
                self.send_json(HTTPStatus.NOT_FOUND, {"error": "File not found"})
                return

            mime_type, _ = mimetypes.guess_type(file_path.name)
            if not mime_type:
                mime_type = "application/octet-stream"

            with open(file_path, "rb") as f:
                content = f.read()

            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", mime_type)
            self.send_header("Content-Disposition", f'attachment; filename="{file_path.name}"')
            self.send_header("Content-Length", str(len(content)))
            self.end_headers()
            self.wfile.write(content)
            return

        # Serve static file from directory
        return super().do_GET()

    def do_POST(self):
        url_parsed = urllib.parse.urlparse(self.path)
        path = url_parsed.path

        if path == "/api/process":
            self.handle_real_process()
            return

        if path == "/api/export-clean":
            self.handle_export_clean()
            return

        self.send_json(HTTPStatus.NOT_FOUND, {"error": "Endpoint not found"})

    def handle_export_clean(self):
        content_length = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(content_length) if content_length > 0 else b"{}"
        try:
            payload = json.loads(body.decode("utf-8")) if body else {}
        except Exception:
            payload = {}

        rows = payload.get("rows", [])
        if not rows:
            data_sample_path = WEB_DIR / "data_samples.js"
            if data_sample_path.exists():
                text = data_sample_path.read_text(encoding="utf-8")
                m = re.search(r"window\.DEMO_ROWS\s*=\s*(\[.*?\]);", text, re.DOTALL)
                if m:
                    try:
                        rows = json.loads(m.group(1))
                    except Exception:
                        rows = []

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp_path = Path(tmp.name)

        try:
            if build_result_workbook is not None:
                build_result_workbook(rows, tmp_path, include_audit=False)
            data = tmp_path.read_bytes()
        except Exception as e:
            self.send_json(HTTPStatus.INTERNAL_SERVER_ERROR, {"error": f"Lỗi tạo file Excel gọn: {e}"})
            return
        finally:
            if tmp_path.exists():
                try:
                    tmp_path.unlink()
                except OSError:
                    pass

        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.send_header("Content-Disposition", 'attachment; filename="TONG_HOP_HANGER_XUONG_GON.xlsx"')
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def handle_real_process(self):
        content_type = self.headers.get("Content-Type", "")
        content_length = int(self.headers.get("Content-Length", 0))

        if not content_type.startswith("multipart/form-data"):
            self.send_json(HTTPStatus.BAD_REQUEST, {"error": "Content-Type must be multipart/form-data"})
            return

        body = self.rfile.read(content_length)
        files, fields = parse_multipart(body, content_type)

        job_id = uuid.uuid4().hex[:12]
        JOBS[job_id] = {
            "status": "RUNNING",
            "progress": "Đang khởi động tiến trình xử lý...",
            "start_time": time.time()
        }

        # Start processing in background thread
        t = threading.Thread(target=self.run_background_pipeline, args=(job_id, files, fields), daemon=True)
        t.start()

        self.send_json(HTTPStatus.OK, {
            "status": "RUNNING",
            "job_id": job_id
        })

    @staticmethod
    def run_background_pipeline(job_id: str, files: dict, fields: dict):
        try:
            use_preset = fields.get("use_preset") == "true"

            with tempfile.TemporaryDirectory(prefix="demo_run_") as temp_run_dir:
                temp_dir = Path(temp_run_dir)
                temp_sof_dir = temp_dir / "sofs"
                temp_sof_dir.mkdir(parents=True, exist_ok=True)

                order_file_path = None
                rules_file_path = BASE_DIR / "rules" / "hanger_rules.json"

                # 1. Prepare Order File
                if use_preset or "order_file" not in files:
                    order_file_path = BASE_DIR / "testn8n" / "filedonhang" / "VLK VLH LPO 09.09.26.XLS"
                else:
                    upload_order = files["order_file"][0]
                    order_file_path = temp_dir / upload_order["filename"]
                    order_file_path.write_bytes(upload_order["content"])

                # 2. Prepare SOForm Files: Luôn nạp kho SOForm chuẩn làm nền tảng, rồi ghi đè file tải lên (nếu có)
                for p in (BASE_DIR / "testn8n" / "fileSOForm").glob("*.xlsx"):
                    shutil.copy(p, temp_sof_dir / p.name)

                if not use_preset and "sof_files" in files:
                    for sof_item in files.get("sof_files", []):
                        sof_dest = temp_sof_dir / sof_item["filename"]
                        sof_dest.write_bytes(sof_item["content"])

                # 3. Prepare Rule File
                if "rule_file" in files:
                    rule_item = files["rule_file"][0]
                    rule_dest = temp_dir / rule_item["filename"]
                    rule_dest.write_bytes(rule_item["content"])
                    if rule_dest.suffix.casefold() == ".json":
                        rules_file_path = rule_dest

                t0 = time.time()
                llm_classifier = None
                if DeepSeekClassifier is not None:
                    try:
                        llm_classifier = DeepSeekClassifier.from_env()
                    except Exception as llm_err:
                        print(f"Warning: could not init DeepSeek classifier: {llm_err}")

                audit = run(
                    input_path=order_file_path,
                    sof_root=temp_sof_dir,
                    rules_path=rules_file_path,
                    output_dir=OUTPUTS_DIR,
                    overwrite=True,
                    llm_classifier=llm_classifier
                )

                duration = round(time.time() - t0, 2)
                result_file_path = Path(audit.get("result_file", ""))
                result_download_url = f"/api/download/{result_file_path.name}" if result_file_path.name else ""

                clean_file_path = Path(audit.get("result_clean_file", ""))
                clean_download_url = f"/api/download/{clean_file_path.name}" if clean_file_path.name else ""

                # Prepare client row results (limit 150 rows for snappy frontend rendering)
                row_results = audit.get("row_results", [])
                display_rows = row_results[:150]

                response_data = {
                    "success": True,
                    "duration_seconds": duration,
                    "input_file": Path(order_file_path).name,
                    "total_order_rows": audit.get("total_order_rows", len(row_results)),
                    "matched": audit.get("matched", 0),
                    "review": audit.get("review", 0),
                    "mismatch": audit.get("mismatch", 0),
                    "result_file_name": result_file_path.name,
                    "result_download_url": result_download_url,
                    "clean_file_name": clean_file_path.name,
                    "clean_download_url": clean_download_url,
                    "accounts_checked": audit.get("accounts_checked", []),
                    "display_rows": display_rows,
                    "llm_info": audit.get("llm", {}),
                    "rule_matched_count": sum(1 for r in row_results if r.get("match_method") != "LLM_DEEPSEEK_VERIFIED_SOURCE" and r.get("status") == "MATCHED"),
                    "llm_matched_count": sum(1 for r in row_results if r.get("match_method") == "LLM_DEEPSEEK_VERIFIED_SOURCE")
                }

                JOBS[job_id] = {
                    "status": "COMPLETED",
                    "result": response_data,
                    "duration": duration
                }
        except Exception as exc:
            traceback.print_exc()
            JOBS[job_id] = {
                "status": "ERROR",
                "error": f"Lỗi xử lý file: {str(exc)}"
            }

    def send_json(self, status: HTTPStatus, data: dict):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


def start_server(port=8085):
    server = ThreadingHTTPServer(("0.0.0.0", port), DemoHandler)
    print(f"Hanger Live Demo Server running at http://localhost:{port}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    start_server(8085)
