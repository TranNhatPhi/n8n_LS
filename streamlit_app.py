#!/usr/bin/env python3
"""Streamlit test UI for the Hanger Automation engine."""

from __future__ import annotations

import json
import os
import shutil
import sys
import tempfile
import time
from pathlib import Path

import streamlit as st


BASE_DIR = Path(__file__).resolve().parent
SRC_DIR = BASE_DIR / "src"
sys.path.insert(0, str(SRC_DIR))

from deepseek_classifier import DeepSeekClassifier  # noqa: E402
from hanger_automation import AutomationError, run  # noqa: E402


st.set_page_config(
    page_title="Hanger Automation — Full Batch",
    page_icon="🧥",
    layout="wide",
)

st.markdown(
    """
    <style>
    .stApp { background: #0a0d14; }
    [data-testid="stHeader"] { background: rgba(10, 13, 20, .85); }
    .hero {
        padding: 1.6rem 1.8rem;
        border: 1px solid #1e293b;
        border-radius: 18px;
        background: linear-gradient(135deg, rgba(6,182,212,.12), rgba(139,92,246,.10));
        margin-bottom: 1.25rem;
    }
    .hero h1 { margin: 0 0 .45rem; font-size: 2rem; }
    .hero p { margin: 0; color: #b8c2d3; }
    .mode-box {
        padding: .8rem 1rem;
        border: 1px solid rgba(16,185,129,.45);
        border-radius: 12px;
        background: rgba(16,185,129,.10);
        color: #d1fae5;
        margin-bottom: 1rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def apply_streamlit_secrets() -> None:
    """Copy only known runtime settings from Streamlit Secrets to the process."""
    allowed = (
        "HANGER_LLM_ENABLED",
        "HANGER_LLM_BATCH_SIZE",
        "HANGER_LLM_MAX_CALLS",
        "HANGER_LLM_MAX_EVIDENCE_CHARS",
        "HANGER_LLM_MIN_CONFIDENCE",
        "HANGER_LLM_CONCURRENCY",
        "DEEPSEEK_API_KEY",
        "DEEPSEEK_API_KEY_2",
        "DEEPSEEK_API_KEY_3",
        "DEEPSEEK_MODEL",
        "DEEPSEEK_BASE_URL",
        "DEEPSEEK_TIMEOUT_SECONDS",
        "HANGER_CACHE_URL",
        "HANGER_CACHE_TTL_DAYS",
    )
    for name in allowed:
        try:
            value = st.secrets.get(name)
        except Exception:
            value = None
        if value is not None and str(value).strip():
            os.environ[name] = str(value).strip()


def save_upload(upload, directory: Path, prefix: str = "") -> Path:
    safe_name = Path(upload.name).name
    target = directory / f"{prefix}{safe_name}"
    target.write_bytes(upload.getbuffer())
    return target


def prepare_inputs(
    temp_dir: Path,
    use_preset: bool,
    order_upload,
    sof_uploads,
    rule_upload,
) -> tuple[Path, Path, Path]:
    sof_dir = temp_dir / "sofs"
    sof_dir.mkdir(parents=True, exist_ok=True)

    if use_preset:
        order_path = BASE_DIR / "testn8n" / "filedonhang" / "VLK VLH LPO 09.09.26.XLS"
        for source in (BASE_DIR / "testn8n" / "fileSOForm").glob("*"):
            if source.is_file() and source.suffix.casefold() in {".xls", ".xlsx", ".docx", ".pdf"}:
                shutil.copy2(source, sof_dir / source.name)
    else:
        order_path = save_upload(order_upload, temp_dir)
        for index, upload in enumerate(sof_uploads):
            save_upload(upload, sof_dir, prefix=f"{index + 1:02d}_")

    rules_path = BASE_DIR / "rules" / "hanger_rules.json"
    if rule_upload is not None:
        rules_path = save_upload(rule_upload, temp_dir)

    return order_path, sof_dir, rules_path


def build_classifier():
    apply_streamlit_secrets()
    try:
        return DeepSeekClassifier.from_env(), ""
    except Exception as exc:
        return None, f"Không thể khởi tạo DeepSeek: {exc}"


def result_table(rows: list[dict]) -> list[dict]:
    columns = (
        "po_number",
        "account",
        "ref_number",
        "style",
        "product_description",
        "product_category",
        "size_configuration",
        "order_hang_flat",
        "hanger_code",
        "hanger_color",
        "color_sizer",
        "status",
        "source_sheet",
        "source_cells",
        "validation_note",
    )
    return [{name: row.get(name, "") for name in columns} for row in rows]


def execute_full_batch(use_preset, order_upload, sof_uploads, rule_upload) -> None:
    started = time.monotonic()
    with tempfile.TemporaryDirectory(prefix="hanger_streamlit_") as temp_name:
        temp_dir = Path(temp_name)
        output_dir = temp_dir / "output"
        output_dir.mkdir()
        order_path, sof_dir, rules_path = prepare_inputs(
            temp_dir, use_preset, order_upload, sof_uploads, rule_upload
        )
        classifier, classifier_warning = build_classifier()

        with st.status("Đang xử lý toàn bộ file…", expanded=True) as status:
            st.write(f"Đơn hàng: `{order_path.name}`")
            st.write(f"Số SOForm: `{len(list(sof_dir.iterdir()))}`")
            if classifier is None:
                st.write("DeepSeek đang tắt — dòng chưa có rule chắc chắn sẽ chuyển REVIEW.")
            else:
                st.write("DeepSeek đã bật — mọi kết quả vẫn phải vượt qua kiểm tra bằng chứng nguồn.")
            if classifier_warning:
                st.warning(classifier_warning)

            audit = run(
                input_path=order_path,
                sof_root=sof_dir,
                rules_path=rules_path,
                output_dir=output_dir,
                overwrite=True,
                llm_classifier=classifier,
            )
            status.update(label="Đã xử lý xong toàn bộ file", state="complete", expanded=False)

        result_path = Path(audit.get("result_file", ""))
        clean_path = Path(audit.get("result_clean_file", ""))
        audit_path = Path(audit.get("audit_file", ""))
        st.session_state["hanger_result"] = {
            "duration": round(time.monotonic() - started, 2),
            "summary": {
                "total": audit.get("total_order_rows", 0),
                "checked": audit.get("total_checked_rows", 0),
                "matched": audit.get("matched", 0),
                "review": audit.get("review", 0),
                "mismatch": audit.get("mismatch", 0),
                "skipped": audit.get("skipped_no_sof", 0),
            },
            "rows": audit.get("row_results", []),
            "errors": audit.get("errors", []),
            "llm": audit.get("llm", {}),
            "result_name": result_path.name,
            "result_bytes": result_path.read_bytes() if result_path.is_file() else b"",
            "clean_name": clean_path.name,
            "clean_bytes": clean_path.read_bytes() if clean_path.is_file() else b"",
            "audit_name": audit_path.name,
            "audit_bytes": audit_path.read_bytes() if audit_path.is_file() else json.dumps(
                audit, ensure_ascii=False, indent=2
            ).encode("utf-8"),
        }


st.markdown(
    """
    <div class="hero">
      <h1>🧥 Hanger Automation</h1>
      <p>Đối chiếu đơn hàng với SOForm, kiểm tra bằng chứng nguồn và xuất bảng tổng hợp Hanger.</p>
    </div>
    <div class="mode-box"><strong>Chế độ duy nhất:</strong> xử lý toàn bộ file (Full Batch), không cắt mẫu.</div>
    """,
    unsafe_allow_html=True,
)

use_preset = st.toggle(
    "Dùng bộ dữ liệu đầy đủ có sẵn để thử nghiệm",
    value=False,
    help="Sử dụng file VLK/VLH/LPO đầy đủ và toàn bộ SOForm trong dự án.",
)

left, right = st.columns(2)
with left:
    order_upload = st.file_uploader(
        "1. File đơn hàng",
        type=["xls", "xlsx", "docx", "pdf"],
        disabled=use_preset,
    )
with right:
    sof_uploads = st.file_uploader(
        "2. Các file SOForm tương ứng (tối đa 20)",
        type=["xls", "xlsx", "docx", "pdf"],
        accept_multiple_files=True,
        disabled=use_preset,
    )

rule_upload = st.file_uploader(
    "3. Rule JSON tùy chỉnh (không bắt buộc)",
    type=["json"],
    help="Nếu bỏ trống, hệ thống dùng rules/hanger_rules.json đã kiểm duyệt.",
)

if use_preset:
    st.info("Bộ mẫu đầy đủ đã sẵn sàng. Quá trình có thể mất vài phút tùy cấu hình DeepSeek.")

ready = use_preset or (order_upload is not None and 0 < len(sof_uploads) <= 20)
if len(sof_uploads) > 20:
    st.error("Chỉ được tải tối đa 20 file SOForm trong một lần chạy.")

if st.button(
    "⚡ Chạy toàn bộ file",
    type="primary",
    width="stretch",
    disabled=not ready,
):
    try:
        execute_full_batch(use_preset, order_upload, sof_uploads, rule_upload)
    except AutomationError as exc:
        st.error(f"Không thể xử lý: {exc}")
    except Exception as exc:
        st.exception(exc)

result = st.session_state.get("hanger_result")
if result:
    st.divider()
    st.subheader("Kết quả Full Batch")
    summary = result["summary"]
    cols = st.columns(6)
    for container, label, value in zip(
        cols,
        ("Tổng dòng", "Đã kiểm tra", "MATCHED", "REVIEW", "MISMATCH", "Bỏ qua"),
        (
            summary["total"],
            summary["checked"],
            summary["matched"],
            summary["review"],
            summary["mismatch"],
            summary["skipped"],
        ),
    ):
        container.metric(label, value)

    st.caption(f"Thời gian xử lý: {result['duration']} giây")
    if result["errors"]:
        st.warning("\n".join(result["errors"]))
    llm_failures = result.get("llm", {}).get("failures", [])
    if llm_failures:
        st.warning("DeepSeek: " + "; ".join(llm_failures))

    download_cols = st.columns(3)
    if result["result_bytes"]:
        download_cols[0].download_button(
            "⬇️ Excel đầy đủ + Audit",
            result["result_bytes"],
            result["result_name"],
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )
    if result["clean_bytes"]:
        download_cols[1].download_button(
            "⬇️ Excel gọn",
            result["clean_bytes"],
            result["clean_name"],
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )
    download_cols[2].download_button(
        "⬇️ Audit JSON",
        result["audit_bytes"],
        result["audit_name"] or "hanger.audit.json",
        "application/json",
        width="stretch",
    )

    rows = result_table(result["rows"])
    st.caption(f"Hiển thị {min(len(rows), 500)} dòng đầu tiên trên tổng số {len(rows)} dòng đã kiểm tra.")
    st.dataframe(rows[:500], width="stretch", hide_index=True)
