import json
import sys
import threading
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from hanger_automation import find_sof, run  # noqa: E402
from deepseek_classifier import (  # noqa: E402
    DecisionCache, DeepSeekClassifier, DeepSeekSettings,
)
from worker_api import (  # noqa: E402
    AutomationError,
    Upload,
    parse_multipart,
    safe_upload_name,
    validate_workbook_bytes,
)


HEADERS = [
    "Division", "Season", "Year", "Ref#", "Style", "Product Description",
    "Color", "Label", "Account", "PO#", "PO Qty", "Size Configuration",
    "Pack Ratio", "Master Box Quantity", "Hang/Flat", "Hanger code",
    "Hanger color", "Color Sizer", "Sticker hanger", "SIZE Sticker hanger",
]

CASES = [
    (751419, "26M316", "WH", "2T - 3T - 4T", "Hang", "6110"),
    (751422, "26M316", "WH", "2T - 3T - 4T", "Hang", "6110"),
    (751423, "3XR316", "", "6X", "Flat", "NO"),
    (751424, "34R316", "", "4", "Flat", "NO"),
    (751425, "35R316", "", "5", "Flat", "NO"),
    (751427, "36M316", "WH", "4 - 5 - 6 - 6X", "Hang", "6110"),
    (766452, "36M316", "WH", "4 - 5 - 6 - 6X", "Hang", "6110"),
    (751430, "36R316", "", "6", "Flat", "NO"),
]


def make_sof(root: Path, name: str = "H040M Stock Replenishment SO Form 8.31.26.xlsx") -> Path:
    path = root / name
    wb = Workbook()
    ws = wb.active
    ws.title = "BOTTOMS"
    ws["B15"] = "Nike leggings hanger requirement"
    ws["C16"] = "Hang: 6110 WHITE"
    ws["D17"] = "Flat styles: no hanger"
    ws["G18"] = "White size clip/black lettering"
    wb.save(path)
    return path


def make_hosiery_sof(root: Path) -> Path:
    path = root / "H040M Stock Replenishment SO Form 8.31.26.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "HOSIERY"
    ws["A3"] = "Bibs, Socks"
    ws["A4"] = "Hanger Required"
    ws["B4"] = "No"
    ws["A5"] = "Flat pack"
    wb.save(path)
    return path


class FakeClassifier:
    min_confidence = 0.85

    def __init__(self, decision):
        self.decision = decision
        self.calls = 0

    def classify(self, rows, allowed_categories, sof_wb, selected_sof):
        self.calls += 1
        return {row_number: dict(self.decision) for row_number, _ in rows}

    def audit_info(self):
        return {"enabled": True, "provider": "deepseek", "model": "test", "api_calls": self.calls}


def make_order(path: Path, cases=CASES):
    wb = Workbook()
    ws = wb.active
    ws.title = "Orders"
    ws.freeze_panes = "A2"
    ws.append(HEADERS)
    for po, style, label, sizes, hang_flat, _ in cases:
        values = {
            "Division": "G", "Ref#": "NKG-LEM316", "Style": style,
            "Product Description": "KNIT LEGGING", "Label": label,
            "Account": "H040M", "PO#": po, "Size Configuration": sizes,
            "Hang/Flat": hang_flat,
        }
        ws.append([values.get(header, "") for header in HEADERS])
        ws.cell(ws.max_row, HEADERS.index("PO#") + 1).number_format = "0000000"
    wb.save(path)


def test_all_eight_acceptance_rows(tmp_path):
    sof_root = tmp_path / "sof"
    sof_root.mkdir()
    make_sof(sof_root)
    order = tmp_path / "3. TONG HOP HANGER-2026.xlsx"
    make_order(order)
    original_bytes = order.read_bytes()
    output = tmp_path / "out"
    rules = Path(__file__).resolve().parents[1] / "rules" / "hanger_rules.json"

    audit = run(order, sof_root, rules, output, now=datetime(2026, 9, 14, 12, 30, 45))

    assert audit["matched"] == 8
    assert audit["mismatch"] == 0
    assert audit["review"] == 0
    assert order.read_bytes() == original_bytes
    assert len(audit["row_results"]) == 8
    assert audit["row_results"][0]["po_number"] == "0751419"
    assert audit["row_results"][0]["source_sheet"] == "BOTTOMS"
    assert audit["row_results"][0]["source_cells"] == "B15:G18"
    assert audit["row_results"][0]["status"] == "MATCHED"
    assert Path(audit["output_file"]).name == "3. TONG HOP HANGER-2026_checked_20260914_123045.xlsx"
    wb = load_workbook(audit["output_file"])
    ws = wb["Orders"]
    columns = {cell.value: cell.column for cell in ws[1]}
    for index, case in enumerate(CASES, 2):
        expected_po = f"{case[0]:07d}"
        assert ws.cell(index, columns["PO#"]).number_format == "0000000"
        assert f"{ws.cell(index, columns['PO#']).value:07d}" == expected_po
        assert ws.cell(index, columns["Hanger code"]).value == case[-1]
        assert ws.cell(index, columns["Hanger color"]).value == ("WHITE" if case[-1] == "6110" else "NO")
        assert ws.cell(index, columns["Color Sizer"]).value == (
            "White size clip/black lettering" if case[-1] == "6110" else "NO"
        )
        assert ws.cell(index, columns["Sticker hanger"]).value == "NO"
        assert ws.cell(index, columns["SIZE Sticker hanger"]).value == "NO"
    assert wb["HANGER REVIEW"].max_row == 1
    wb.close()
    with Path(audit["audit_file"]).open(encoding="utf-8") as handle:
        stored = json.load(handle)
    assert stored["total_order_rows"] == 8


def test_both_mismatch_directions_and_blank_are_reviewed(tmp_path):
    sof_root = tmp_path / "sof"
    sof_root.mkdir()
    make_sof(sof_root)
    cases = [
        (751419, "26M316", "WH", "2T - 3T - 4T", "Flat", ""),
        (751423, "3XR316", "", "6X", "Hang", ""),
        (751427, "36M316", "WH", "4 - 5 - 6 - 6X", "", ""),
    ]
    order = tmp_path / "orders.xlsx"
    make_order(order, cases)
    rules = Path(__file__).resolve().parents[1] / "rules" / "hanger_rules.json"

    audit = run(order, sof_root, rules, tmp_path / "out")

    assert audit["matched"] == 0
    assert audit["mismatch"] == 2
    assert audit["review"] == 1
    wb = load_workbook(audit["output_file"])
    review = wb["HANGER REVIEW"]
    statuses = [review.cell(row, 19).value for row in range(2, 5)]
    assert statuses == ["MISMATCH", "MISMATCH", "REVIEW"]
    wb.close()


def test_ambiguous_latest_sof_is_not_guessed(tmp_path):
    make_sof(tmp_path, "H040M SO Form 8.31.26.xlsx")
    make_sof(tmp_path, "archive H040M SO Form 8.31.26.xlsx")
    selection = find_sof(tmp_path, "H040M")
    assert selection.status == "REVIEW"
    assert selection.path is None
    assert "share the latest revision date" in selection.reason


def test_undated_candidate_blocks_latest_selection(tmp_path):
    make_sof(tmp_path, "H040M SO Form 8.31.26.xlsx")
    make_sof(tmp_path, "H040M current.xlsx")
    selection = find_sof(tmp_path, "H040M")
    assert selection.status == "REVIEW"
    assert "not verifiable" in selection.reason


def test_existing_hanger_data_is_preserved(tmp_path):
    sof_root = tmp_path / "sof"
    sof_root.mkdir()
    make_sof(sof_root)
    order = tmp_path / "orders.xlsx"
    make_order(order, [CASES[0]])
    wb = load_workbook(order)
    ws = wb["Orders"]
    ws.cell(2, HEADERS.index("Hanger code") + 1, "MANUAL-CODE")
    wb.save(order)
    wb.close()
    rules = Path(__file__).resolve().parents[1] / "rules" / "hanger_rules.json"

    audit = run(order, sof_root, rules, tmp_path / "out")

    assert audit["skipped_existing_hanger_data"] == 1
    assert audit["row_results"] == []
    wb = load_workbook(audit["output_file"])
    assert wb["Orders"].cell(2, HEADERS.index("Hanger code") + 1).value == "MANUAL-CODE"
    wb.close()


def test_uncited_rule_becomes_review(tmp_path):
    sof_root = tmp_path / "sof"
    sof_root.mkdir()
    sof = sof_root / "H040M SO Form 8.31.26.xlsx"
    Workbook().save(sof)
    order = tmp_path / "orders.xlsx"
    make_order(order, [CASES[0]])
    rules = Path(__file__).resolve().parents[1] / "rules" / "hanger_rules.json"

    audit = run(order, sof_root, rules, tmp_path / "out")

    assert audit["review"] == 1
    assert "Source sheet is missing" in audit["row_results"][0]["validation_note"]


def test_worker_sanitizes_and_validates_uploaded_workbooks(tmp_path):
    assert safe_upload_name(r"C:\\fakepath\\orders.xlsx", "order") == "orders.xlsx"
    valid_path = tmp_path / "valid.xlsx"
    Workbook().save(valid_path)
    validate_workbook_bytes(Upload("valid.xlsx", valid_path.read_bytes()), "order")
    try:
        safe_upload_name("orders.pdf", "order")
    except AutomationError as exc:
        assert "must be an .xls or .xlsx" in str(exc)
    else:
        raise AssertionError("Worker accepted a non-workbook upload")


def test_worker_parses_n8n_multipart_upload(tmp_path):
    workbook = tmp_path / "orders.xlsx"
    Workbook().save(workbook)
    content = workbook.read_bytes()
    boundary = "hanger-test-boundary"
    parts = []
    for field, filename in (
        ("order_workbook", "orders.xlsx"),
        ("sof_workbook", "H040M SO Form 8.31.26.xlsx"),
    ):
        parts.append(
            f"--{boundary}\r\n"
            f'Content-Disposition: form-data; name="{field}"; filename="{filename}"\r\n'
            "Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n"
        .encode() + content + b"\r\n")
    parts.append(
        f"--{boundary}\r\nContent-Disposition: form-data; name=\"overwrite\"\r\n\r\nfalse\r\n"
        f"--{boundary}--\r\n".encode()
    )
    files, fields = parse_multipart(b"".join(parts), f"multipart/form-data; boundary={boundary}")
    assert set(files) == {"order_workbook", "sof_workbook"}
    assert fields["overwrite"] == "false"


def test_verified_llm_fallback_can_fill_an_unruled_flat_item(tmp_path):
    sof_root = tmp_path / "sof"
    sof_root.mkdir()
    make_hosiery_sof(sof_root)
    order = tmp_path / "orders.xlsx"
    make_order(order, [(711997, "BJ0871", "79", "7-9", "Flat", "")])
    wb = load_workbook(order)
    wb["Orders"].cell(2, HEADERS.index("Product Description") + 1, "6PK CREW SOCK")
    wb.save(order)
    wb.close()
    decision = {
        "status": "MATCHED", "product_category": "OTHER", "sof_hang_flat": "Flat",
        "hanger_code": "NO", "hanger_color": "NO", "color_sizer": "NO",
        "sticker_hanger": "NO", "size_sticker_hanger": "NO",
        "source_sheet": "HOSIERY", "source_cells": "A3:B5", "confidence": 0.96,
        "reasoning": "SOF says hosiery has no hanger and is flat packed.",
    }
    classifier = FakeClassifier(decision)
    rules = Path(__file__).resolve().parents[1] / "rules" / "hanger_rules.json"

    audit = run(order, sof_root, rules, tmp_path / "out", llm_classifier=classifier)

    assert audit["matched"] == 1
    assert audit["row_results"][0]["match_method"] == "LLM_DEEPSEEK_VERIFIED_SOURCE"
    assert audit["row_results"][0]["source_cells"] == "A3:B5"
    result_wb = load_workbook(audit["output_file"])
    columns = {cell.value: cell.column for cell in result_wb["Orders"][1]}
    assert result_wb["Orders"].cell(2, columns["Hanger code"]).value == "NO"
    result_wb.close()


def test_llm_hallucinated_value_is_rejected(tmp_path):
    sof_root = tmp_path / "sof"
    sof_root.mkdir()
    make_sof(sof_root)
    order = tmp_path / "orders.xlsx"
    make_order(order, [(711999, "UNKNOWN", "WH", "2T - 3T", "Hang", "")])
    decision = {
        "status": "MATCHED", "product_category": "BOTTOMS", "sof_hang_flat": "Hang",
        "hanger_code": "9999", "hanger_color": "BLUE", "color_sizer": "Blue clip",
        "sticker_hanger": "NO", "size_sticker_hanger": "NO",
        "source_sheet": "BOTTOMS", "source_cells": "B15:G18", "confidence": 0.99,
        "reasoning": "Unsupported values.",
    }
    rules = Path(__file__).resolve().parents[1] / "rules" / "hanger_rules.json"

    audit = run(
        order, sof_root, rules, tmp_path / "out", llm_classifier=FakeClassifier(decision)
    )

    assert audit["review"] == 1
    assert "Rejected DeepSeek decision" in audit["row_results"][0]["validation_note"]
    assert "hanger_code" in audit["row_results"][0]["validation_note"]


def test_deepseek_groups_duplicate_order_patterns_before_api_call(tmp_path):
    sof = make_hosiery_sof(tmp_path)
    wb = load_workbook(sof, data_only=True)
    settings = DeepSeekSettings(api_keys=("test",), batch_size=12, max_calls=2)
    classifier = DeepSeekClassifier(settings)
    captured = []

    def fake_request(groups, allowed, source_name, evidence, api_key):
        captured.append(groups)
        return {"decisions": [
            {"group_id": group_id, "status": "REVIEW", "reasoning": "test"}
            for group_id in groups
        ]}

    classifier._request = fake_request
    row = {
        "Account": "H040M", "Division": "G", "Label": "79", "Ref#": "R1",
        "Style": "S1", "Product Description": "6PK CREW SOCK",
        "Size Configuration": "7-9", "Hang/Flat": "Flat",
    }
    decisions = classifier.classify([(2, row), (3, dict(row))], {"OTHER"}, wb, sof)
    wb.close()

    assert len(captured) == 1
    assert len(captured[0]) == 1
    assert set(decisions) == {2, 3}
    assert classifier.groups_requested == 1


def test_unique_pending_sheet_is_preferred_over_completed_history(tmp_path):
    sof_root = tmp_path / "sof"
    sof_root.mkdir()
    make_sof(sof_root)
    order = tmp_path / "multi-sheet.xlsx"
    wb = Workbook()
    completed = wb.active
    completed.title = "Completed"
    completed.append(HEADERS)
    completed.append(["" for _ in HEADERS])
    pending = wb.create_sheet("Pending")
    pending_headers = HEADERS[:15]
    pending.append(pending_headers)
    values = {
        "Division": "G", "Ref#": "NKG-LEM316", "Style": "26M316",
        "Product Description": "KNIT LEGGING", "Label": "WH", "Account": "H040M",
        "PO#": 751419, "Size Configuration": "2T - 3T - 4T", "Hang/Flat": "Hang",
    }
    pending.append([values.get(header, "") for header in pending_headers])
    wb.save(order)
    wb.close()
    rules = Path(__file__).resolve().parents[1] / "rules" / "hanger_rules.json"

    audit = run(order, sof_root, rules, tmp_path / "out")

    assert audit["matched"] == 1
    result_wb = load_workbook(audit["output_file"])
    assert result_wb["HANGER REVIEW"].max_row == 1
    pending_columns = {cell.value: cell.column for cell in result_wb["Pending"][1]}
    assert result_wb["Pending"].cell(2, pending_columns["Hanger code"]).value == "6110"
    result_wb.close()


def test_multiple_api_keys_run_batches_in_parallel(tmp_path):
    sof = make_hosiery_sof(tmp_path)
    wb = load_workbook(sof, data_only=True)
    settings = DeepSeekSettings(
        api_keys=("key-a", "key-b"), batch_size=1, max_calls=10
    )
    classifier = DeepSeekClassifier(settings)
    barrier = threading.Barrier(2, timeout=5)
    used_keys = []
    lock = threading.Lock()

    def fake_request(groups, allowed, source_name, evidence, api_key):
        with lock:
            used_keys.append(api_key)
        # Deadlocks unless two batches are genuinely in flight together.
        barrier.wait()
        return {"decisions": [
            {"group_id": group_id, "status": "REVIEW", "reasoning": "test"}
            for group_id in groups
        ]}

    classifier._request = fake_request
    rows = []
    for index in range(2):
        rows.append((index + 2, {
            "Account": "H040M", "Division": "G", "Label": "79", "Ref#": f"R{index}",
            "Style": f"S{index}", "Product Description": "6PK CREW SOCK",
            "Color": "023", "Size Configuration": "", "Hang/Flat": "",
        }))

    classifier.classify(rows, {"HOSIERY"}, wb, Path(sof))
    wb.close()

    assert sorted(used_keys) == ["key-a", "key-b"]


def test_api_keys_parsed_from_comma_list_and_numbered_slots(monkeypatch):
    monkeypatch.setenv("HANGER_LLM_ENABLED", "true")
    monkeypatch.setenv("DEEPSEEK_API_KEY", " key-one , key-two ")
    monkeypatch.setenv("DEEPSEEK_API_KEY_2", "key-three")
    monkeypatch.setenv("DEEPSEEK_API_KEY_3", "key-two")  # duplicate is dropped
    settings = DeepSeekSettings.from_env()
    assert settings.api_keys == ("key-one", "key-two", "key-three")


def test_sof_sourced_columns_get_yellow_header(tmp_path):
    sof_root = tmp_path / "sof"
    sof_root.mkdir()
    make_sof(sof_root)
    order = tmp_path / "3. TONG HOP HANGER-2026.xlsx"
    make_order(order)
    rules = Path(__file__).resolve().parents[1] / "rules" / "hanger_rules.json"

    audit = run(order, sof_root, rules, tmp_path / "out")

    result_wb = load_workbook(audit["output_file"])
    ws = result_wb[result_wb.sheetnames[0]]
    header_row = 1
    headers = {cell.value: cell.column for cell in ws[header_row]}
    for field in ("Hanger code", "Hanger color", "Color Sizer",
                  "Sticker hanger", "SIZE Sticker hanger"):
        cell = ws.cell(header_row, headers[field])
        assert cell.fill.fill_type == "solid"
        assert cell.fill.fgColor.rgb.endswith("FFFF00"), field
    result_wb.close()


class FakeRedis:
    """Minimal stand-in for the redis client surface DecisionCache uses."""

    def __init__(self, broken=False):
        self.store = {}
        self.broken = broken
        self.writes = 0

    def mget(self, keys):
        if self.broken:
            raise ConnectionError("redis down")
        return [self.store.get(k) for k in keys]

    def pipeline(self, transaction=False):
        return FakePipeline(self)


class FakePipeline:
    def __init__(self, parent):
        self.parent = parent
        self.queued = []

    def set(self, key, value, ex=None):
        self.queued.append((key, value))

    def execute(self):
        if self.parent.broken:
            raise ConnectionError("redis down")
        for key, value in self.queued:
            self.parent.store[key] = value
            self.parent.writes += 1


def _classifier_rows():
    return [(2, {
        "Account": "H040M", "Division": "G", "Label": "79", "Ref#": "R1",
        "Style": "S1", "Product Description": "6PK CREW SOCK",
        "Color": "023", "Size Configuration": "", "Hang/Flat": "",
    })]


def test_cached_decision_is_reused_without_calling_the_api(tmp_path):
    sof = make_hosiery_sof(tmp_path)
    wb = load_workbook(sof, data_only=True)
    fake = FakeRedis()
    settings = DeepSeekSettings(api_keys=("k",), batch_size=12, max_calls=5)

    calls = []

    def fake_request(groups, allowed, source_name, evidence, api_key):
        calls.append(groups)
        return {"decisions": [
            {"group_id": gid, "status": "MATCHED", "hanger_code": "6110"}
            for gid in groups
        ]}

    first = DeepSeekClassifier(settings, DecisionCache(fake, 3600))
    first._request = fake_request
    first.classify(_classifier_rows(), {"HOSIERY"}, wb, Path(sof))
    assert len(calls) == 1 and fake.writes == 1

    # A second classifier sharing the cache must not hit the API at all.
    second = DeepSeekClassifier(settings, DecisionCache(fake, 3600))
    second._request = fake_request
    out = second.classify(_classifier_rows(), {"HOSIERY"}, wb, Path(sof))
    wb.close()

    assert len(calls) == 1, "cached group was re-sent to the API"
    assert out[2]["hanger_code"] == "6110"
    assert second.cache.hits == 1


def test_changed_sof_contents_do_not_reuse_cached_answers(tmp_path):
    sof_a = make_hosiery_sof(tmp_path)
    wb = load_workbook(sof_a, data_only=True)
    fake = FakeRedis()
    settings = DeepSeekSettings(api_keys=("k",), batch_size=12, max_calls=5)
    key_a = DecisionCache.build_key("m", "digestA", ("x",), {"HOSIERY"})
    key_b = DecisionCache.build_key("m", "digestB", ("x",), {"HOSIERY"})
    assert key_a != key_b

    other = tmp_path / "other.xlsx"
    other.write_bytes(sof_a.read_bytes() + b"different")
    from deepseek_classifier import _file_digest
    assert _file_digest(sof_a) != _file_digest(other)
    wb.close()


def test_broken_cache_degrades_to_plain_api_calls(tmp_path):
    sof = make_hosiery_sof(tmp_path)
    wb = load_workbook(sof, data_only=True)
    fake = FakeRedis(broken=True)
    settings = DeepSeekSettings(api_keys=("k",), batch_size=12, max_calls=5)
    classifier = DeepSeekClassifier(settings, DecisionCache(fake, 3600))
    calls = []

    def fake_request(groups, allowed, source_name, evidence, api_key):
        calls.append(groups)
        return {"decisions": [
            {"group_id": gid, "status": "MATCHED", "hanger_code": "6110"}
            for gid in groups
        ]}

    classifier._request = fake_request
    out = classifier.classify(_classifier_rows(), {"HOSIERY"}, wb, Path(sof))
    wb.close()

    assert len(calls) == 1, "a dead cache must not stop the API path"
    assert out[2]["hanger_code"] == "6110"
    assert classifier.cache.errors >= 1


def test_result_sheet_has_tong_hop_columns_and_header_colours(tmp_path):
    sof_root = tmp_path / "sof"
    sof_root.mkdir()
    make_sof(sof_root)
    order = tmp_path / "3. TONG HOP HANGER-2026.xlsx"
    make_order(order)
    rules = Path(__file__).resolve().parents[1] / "rules" / "hanger_rules.json"

    audit = run(order, sof_root, rules, tmp_path / "out")

    result = Path(audit["result_file"])
    assert result.is_file()
    wb = load_workbook(result)
    assert wb.sheetnames == ["KET QUA"], "deliverable must be a single tab"
    ws = wb["KET QUA"]
    headers = [c.value for c in ws[1]]
    assert headers == [
        "Division", "Season", "Year", "Ref#", "Style", "Product Description",
        "Color", "Label", "Account", "PO#", "PO Qty", "Size Configuration",
        "Pack Ratio", "Master Box Quantity", "Hang/Flat",
        "Hanger code", "Hanger color", "Color Sizer", "Sticker hanger",
        "SIZE Sticker hanger", "NCC",
        "Status", "Validation Note", "Match Method", "Confidence",
        "SOF File", "SOF Sheet", "SOF Cells",
    ]
    for cell in ws[1]:
        if cell.column <= 15:
            expected = "33CCCC"
        elif cell.column <= 21:
            expected = "FFFF00"
        else:
            expected = "D9D9D9"
        assert cell.fill.fill_type == "solid"
        assert cell.fill.fgColor.rgb.endswith(expected), cell.value
    assert ws.max_row == len(audit["row_results"]) + 1
    wb.close()


def test_result_sheet_leaves_unverified_rows_blank(tmp_path):
    sof_root = tmp_path / "sof"
    sof_root.mkdir()
    make_sof(sof_root)
    order = tmp_path / "3. TONG HOP HANGER-2026.xlsx"
    make_order(order)
    rules = Path(__file__).resolve().parents[1] / "rules" / "hanger_rules.json"

    audit = run(order, sof_root, rules, tmp_path / "out")
    wb = load_workbook(Path(audit["result_file"]))
    ws = wb["KET QUA"]
    headers = {c.value: c.column for c in ws[1]}
    for offset, entry in enumerate(audit["row_results"], start=2):
        value = ws.cell(offset, headers["Hanger code"]).value or ""
        if entry["status"] == "MATCHED":
            assert value == entry["hanger_code"]
        else:
            assert value == "", "unverified row must not carry a hanger value"
        assert (ws.cell(offset, headers["NCC"]).value or "") == ""
    wb.close()


def test_result_sheet_explains_every_row_including_reviews(tmp_path):
    sof_root = tmp_path / "sof"
    sof_root.mkdir()
    make_sof(sof_root)
    order = tmp_path / "3. TONG HOP HANGER-2026.xlsx"
    make_order(order)
    rules = Path(__file__).resolve().parents[1] / "rules" / "hanger_rules.json"

    audit = run(order, sof_root, rules, tmp_path / "out")
    wb = load_workbook(Path(audit["result_file"]))
    ws = wb["KET QUA"]
    headers = {c.value: c.column for c in ws[1]}

    for offset, entry in enumerate(audit["row_results"], start=2):
        # Status and its explanation are present on every row, matched or not.
        assert ws.cell(offset, headers["Status"]).value == entry["status"]
        assert (ws.cell(offset, headers["Validation Note"]).value or "") == entry["validation_note"]
        assert (ws.cell(offset, headers["SOF Cells"]).value or "") == entry["source_cells"]
    wb.close()
