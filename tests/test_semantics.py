import json

import pytest

from xlsprint import formulas, report, semantics


def _sidecar(path, annotations, workbook_sha256=None):
    path.write_text(json.dumps({
        "schema": semantics.SCHEMA,
        "workbook_sha256": workbook_sha256,
        "annotations": annotations,
    }), encoding="utf-8")
    return semantics.load_manifest(path)


def test_manifest_binds_clear_workbook_selectors_to_redacted_targets(tmp_path):
    salt = "test-salt"
    sheet = formulas.redact_name("Operations", salt)
    name = formulas.redact_name("NetGeneration", salt)
    workbook_sha = "a" * 64
    manifest = _sidecar(tmp_path / "meaning.json", [
        {"id": "generation", "label": "Net generation", "intent": "Convert gross output to saleable MWh.",
         "category": "Operations", "source": "Belfort route map", "match": {"defined_name": "NetGeneration"}},
        {"id": "convergence", "label": "Convergence calculation", "intent": "Recalculate the model after a convergence write.",
         "match": {"span": {"kind": "vba.proc", "name": "CopyPasteMainRoutine"}}},
        {"id": "monthly-output", "label": "Monthly output block", "intent": "Calculate the forecast output block.",
         "match": {"range": {"sheet": "Operations", "address": "$F$12:$F$23"}}},
        {"id": "generation-formulas", "label": "Generation formulas", "intent": "Apply the generation calculation family.",
         "match": {"formula_group": {"sheet": "Operations", "address": "F12:F23"}}},
    ], workbook_sha)
    formulas_obj = {
        "workbook_sha256": workbook_sha,
        "names": [{"name": name, "scope": "workbook", "sheet": sheet,
                   "refers_to_range": f"{sheet}!$F$12:$F$23", "is_range": True}],
    }
    plan = {
        "redaction_map": {sheet: "Operations", name: "NetGeneration"},
        "sheets": [{"sheet": "Operations"}],
        "groups": [{"group": "G0001", "sheet": "Operations", "areas": ["$F$12:$F$23"]}],
    }
    bound = semantics.bind_manifest(manifest, workbook_sha256=workbook_sha,
                                    formulas=formulas_obj, plan=plan, names="hashed", salt=salt)
    by_id = {item["id"]: item for item in bound["annotations"]}
    assert by_id["generation"]["targets"][0]["name"] == name
    assert by_id["generation"]["targets"][0]["sheet"] == sheet
    assert by_id["convergence"]["targets"][0]["name"] == formulas.redact_name("CopyPasteMainRoutine", salt)
    assert by_id["monthly-output"]["targets"][0]["sheet"] == sheet
    assert by_id["generation"]["binding"] == "bound"
    assert "NetGeneration" not in json.dumps(bound)
    assert '"selector"' not in json.dumps(bound)


def test_manifest_rejects_wrong_workbook_hash_and_ambiguous_or_unsafe_input(tmp_path):
    manifest = _sidecar(tmp_path / "meaning.json", [
        {"id": "net-generation", "label": "Net generation", "intent": "Convert output.",
         "match": {"defined_name": "NetGeneration"}},
    ], "a" * 64)
    with pytest.raises(semantics.SemanticsError, match="does not match"):
        semantics.bind_manifest(manifest, workbook_sha256="b" * 64,
                                 formulas={"workbook_sha256": "b" * 64}, plan={}, names="clear", salt="")

    with pytest.raises(semantics.SemanticsError, match="exactly one"):
        _sidecar(tmp_path / "ambiguous.json", [
            {"id": "bad", "label": "Bad", "intent": "Bad map.",
             "match": {"span": {"kind": "calc.range"}, "range": {"sheet": "S", "address": "A1"}}},
        ])
    with pytest.raises(semantics.SemanticsError, match="workbook_sha256 is required"):
        _sidecar(tmp_path / "unbound-range.json", [
            {"id": "range", "label": "Range", "intent": "Describe a range.",
             "match": {"range": {"sheet": "S", "address": "A1"}}},
        ])


def test_report_shows_semantic_intent_beside_defined_names_and_escapes_text():
    name_token = formulas.redact_name("NetGeneration", "s")
    sheet_token = formulas.redact_name("Operations", "s")
    formulas_obj = {
        "schema": "xlsprint.formulas/1",
        "workbook_sha256": "a" * 64,
        "workbook_sha256_prefix": "a" * 10,
        "redaction": {"names": "hashed"},
        "names": [{"name": name_token, "scope": "workbook", "sheet": sheet_token,
                   "refers_to_range": f"{sheet_token}!$F$12:$F$23", "is_range": True,
                   "formula_cells": 12}],
        "semantics": {
            "schema": semantics.SCHEMA, "workbook_sha256": "a" * 64, "unbound_count": 0,
            "annotations": [{"id": "generation", "label": "Net generation",
                             "intent": "Convert gross production into saleable MWh.",
                             "category": "Operations", "source": "Analyst map", "binding": "bound",
                             "targets": [{"target": "defined_name", "name": name_token,
                                          "scope": "workbook", "sheet": sheet_token,
                                          "address": "F12:F23", "span_kind": "calc.name"}] }],
        },
    }
    fm = report._formulas_model(formulas_obj)
    model = {"schema": report.REPORT_SCHEMA, "meta": {"label": "test", "source_kind": "real"},
             "validation": {"ok": True}, "formulas": fm, "timeline": {"rows": []},
             "drilldown": {"sheets": []}, "hotspots": [], "vba_procs": []}
    report._apply_semantics(model)
    html = report.render_html(model)
    assert "Model meaning" in html
    assert "Convert gross production into saleable MWh." in html
    assert "Net generation" in html
    assert "h:" in html
    assert html.count("Net generation") >= 2  # map summary and defined-name inventory
    assert "NetGeneration" not in html
    assert "a" * 64 not in html
    assert "formula_text" not in html
    assert 'id="semantics"' in html


def test_profile_cli_accepts_semantics_sidecar(tmp_path):
    from xlsprint.cli import build_parser

    args = build_parser().parse_args([
        "profile", "model.xlsx", "--out", str(tmp_path / "out"),
        "--semantics", str(tmp_path / "meaning.json"),
    ])
    assert args.semantics == str(tmp_path / "meaning.json")


def test_region_intent_is_attached_to_measured_drilldown_and_hotspot_rows():
    annotation = {"id": "forecast", "label": "Operating forecast",
                  "intent": "Calculate monthly output after applying availability assumptions.",
                  "category": "Operations", "source": "Analyst", "binding": "bound",
                  "targets": [{"target": "range", "sheet": "h:sheet", "address": "F12:F23",
                               "span_kinds": ["calc.range", "calc.name", "calc.group"]}]}
    formulas_model = report._formulas_model({
        "schema": "xlsprint.formulas/1", "redaction": {"names": "hashed"},
        "semantics": {"schema": semantics.SCHEMA, "workbook_bound": True,
                       "annotations": [annotation]},
    })
    model = {"schema": report.REPORT_SCHEMA, "meta": {"label": "run", "source_kind": "real"},
             "validation": {"ok": True}, "formulas": formulas_model,
             "timeline": {"rows": []},
             "drilldown": {"sheets": [{"name": "h:sheet", "children": [
                 {"kind": "calc.range", "name": "auto1", "address": "$F$12:$F$23"}
             ]}]},
             "hotspots": [{"kind": "calc.range", "name": "auto1", "sheet": "h:sheet",
                            "address": "$F$12:$F$23"}],
             "vba_procs": []}
    report._apply_semantics(model)
    assert model["drilldown"]["sheets"][0]["children"][0]["annotations"][0]["id"] == "forecast"
    assert model["hotspots"][0]["annotations"][0]["label"] == "Operating forecast"
    assert model["formulas"]["semantics"]["annotations"][0]["timed_match_count"] == 2
    html = report.render_html(model)
    assert "Calculate monthly output after applying availability assumptions." in html
