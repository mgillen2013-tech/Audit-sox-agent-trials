"""The joint: agent conclusions -> a workpaper that mirrors PY.

These are the tests that would have caught the two halves drifting apart.
Both existed and neither called the other for several rounds of work, so
the interesting assertions here are about the SEAM, not either side.
"""

from __future__ import annotations

from pathlib import Path

import openpyxl
import pytest

from agent.build_cy_workpaper import build_cy_workpaper
from agent.schemas import ConclusionOutput, ModelMetadata, SampleItem, SamplePopulationManifest


def _template(path: Path, *, sample: str = "Sample", population: str = "Population") -> Path:
    """A workbook shaped like a PY workpaper.

    The distinct number formats are not decoration. check_template refuses
    a template whose styles.xml defines fewer cell formats than the
    STYLE_* constants address, and a default openpyxl workbook has about
    two -- so a "minimal" fixture silently exercises the FALLBACK path
    while looking like it tests the mirroring one. Real Excel workpapers
    carry plenty; this makes the fixture resemble one.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sample
    ws["A1"] = "Invoice Number"
    ws["A2"] = "existing sample row"  # the row the builder replaces
    pop = wb.create_sheet(population)
    pop["A5"] = "Invoice Number"
    for r in range(6, 12):
        pop.cell(r, 1, f"population row {r}")
    for i in range(20):
        c = ws.cell(50 + i, 1, i)
        c.number_format = f"0.{'0' * (i % 8)}#{'#' * (i % 3)}"
    wb.save(path)
    return path


def _real_pdf(path: Path) -> Path:
    """A PDF that actually opens. b"%PDF-1.4" does not -- it is a valid file
    header and nothing else, and every renderer rejects it."""
    from reportlab.lib.pagesizes import letter as PAGE
    from reportlab.pdfgen import canvas

    c = canvas.Canvas(str(path), pagesize=PAGE)
    c.setFont("Helvetica", 12)
    c.drawString(72, 700, "Invoice Number: 21022452")
    c.save()
    return path


def _spec() -> dict:
    return {
        "control_id": "C-14",
        "control_objective_ref": "CO-1",
        "control_objective_text": "Invoices are approved before payment.",
        "py_testing_file": "PY.xlsx",
        "cy_support_files": [],
        "test_steps": [{"test_step_id": "TS-1", "test_step_text": "Verify approval precedes payment."}],
    }


def _results(**overrides) -> dict:
    base = dict(
        test_step_id="TS-1",
        control_objective_ref="CO-1",
        conclusion="satisfied",
        narrative="Approval preceded payment.",
        evidence_citations=[
            {
                "evidence_id": "ev_0001",
                "source_file": "invoice.pdf",
                "location": "invoice.pdf p.1",
                "quote_or_summary": "Invoice 21022452",
                "relevance": "the invoice",
                "sample_id": "1",
            }
        ],
        procedures_performed=["inspection"],
        sample_results=[{"sample_id": "1", "conclusion": "satisfied"}],
        attribute_results=[
            {
                "attribute": "Invoice number agrees to the listing",
                "sample_id": "1",
                "result": "satisfied",
                "value_observed": "Invoice 21022452",
                "evidence_ids": ["ev_0001"],
            }
        ],
        exceptions=[],
        additional_support_requests=[],
        confidence="high",
        confidence_rationale="Clear.",
        sample_coverage=None,
        model_metadata=ModelMetadata(
            model="claude-opus-5", prompt_version="v1", timestamp="2026-09-03T00:00:00Z", tool_call_count=3
        ),
    )
    base.update(overrides)
    return {"TS-1": {"conclusion": ConclusionOutput(**base), "audit_log": []}}


def _manifests() -> dict:
    return {
        "TS-1": SamplePopulationManifest(
            test_step_id="TS-1",
            sample_size=1,
            samples=[
                SampleItem(
                    sample_id="1",
                    test_step_id="TS-1",
                    identifying_details="Invoice 21022452",
                    key_fields={"Invoice Number": "21022452"},
                )
            ],
        )
    }


def test_a_usable_py_template_is_mirrored(tmp_path: Path):
    _template(tmp_path / "PY.xlsx")
    _real_pdf(tmp_path / "invoice.pdf")
    outcome = build_cy_workpaper(
        _spec(), _results(), "PY.xlsx", tmp_path / "out",
        support_dir=tmp_path, sample_manifests=_manifests(),
    )
    assert outcome.mirrored_py
    wb = openpyxl.load_workbook(outcome.path)
    # The template's own tabs survive, with the Summary added in front.
    assert wb.sheetnames[0] == "Summary"
    assert {"Sample", "Population"} <= set(wb.sheetnames)
    wb.close()


def test_a_pdf_prior_year_falls_back_instead_of_failing(tmp_path: Path):
    # Plenty of PY workpapers are PDFs. A run that reached real conclusions
    # has already spent real money and real evidence review -- a
    # worse-looking file carrying them beats an error every time.
    (tmp_path / "PY.pdf").write_bytes(b"%PDF-1.4\n")
    outcome = build_cy_workpaper(
        _spec(), _results(), "PY.pdf", tmp_path / "out",
        support_dir=tmp_path, sample_manifests=_manifests(),
    )
    assert outcome.path.exists()
    assert not outcome.mirrored_py
    # And it must SAY so -- otherwise the only clue is that the file looks
    # nothing like last year's, which reads as a regression.
    assert any("could not mirror" in w for w in outcome.warnings)


def test_a_wrongly_shaped_workbook_falls_back_rather_than_writing_into_it(tmp_path: Path):
    # Named .xlsx but shaped for a different control. Writing into it would
    # put the sample rows in whatever tab happened to be second.
    _template(tmp_path / "PY.xlsx", sample="SomethingElse", population="Other")
    outcome = build_cy_workpaper(
        _spec(), _results(), "PY.xlsx", tmp_path / "out",
        support_dir=tmp_path, sample_manifests=_manifests(),
    )
    assert not outcome.mirrored_py
    assert outcome.path.exists()


def test_a_control_where_every_step_failed_still_produces_a_workpaper(tmp_path: Path):
    _template(tmp_path / "PY.xlsx")
    failed = {"TS-1": {"error": "token budget exceeded", "reason": "token_budget_exceeded",
                       "tokens_used": 1, "turns_used": 1, "audit_log": []}}
    outcome = build_cy_workpaper(
        _spec(), failed, "PY.xlsx", tmp_path / "out", support_dir=tmp_path,
    )
    assert outcome.path.exists()
    assert not outcome.mirrored_py


def test_the_population_is_never_emptied_by_the_joint(tmp_path: Path):
    # The template's population must survive when the caller has none of
    # its own -- deleting it removes the basis for the sample.
    _template(tmp_path / "PY.xlsx")
    _real_pdf(tmp_path / "invoice.pdf")
    outcome = build_cy_workpaper(
        _spec(), _results(), "PY.xlsx", tmp_path / "out",
        support_dir=tmp_path, sample_manifests=_manifests(),
    )
    wb = openpyxl.load_workbook(outcome.path)
    rows = [r for r in wb["Population"].iter_rows(min_row=6, values_only=True) if any(v is not None for v in r)]
    wb.close()
    assert len(rows) == 6


def test_a_corrupt_evidence_file_loses_one_exhibit_not_the_workpaper(tmp_path: Path):
    # Support arrives from scanners, mail clients and portals, and some of
    # it is broken. Losing one exhibit is recoverable; losing the workpaper
    # -- and with it every conclusion the run paid for -- is not.
    _template(tmp_path / "PY.xlsx")
    (tmp_path / "invoice.pdf").write_bytes(b"%PDF-1.4\n")  # a header and nothing else
    outcome = build_cy_workpaper(
        _spec(), _results(), "PY.xlsx", tmp_path / "out",
        support_dir=tmp_path, sample_manifests=_manifests(),
    )
    assert outcome.path.exists() and outcome.mirrored_py
    assert any("could not be read" in reason for *_x, reason in outcome.unplaced_callouts)
    assert "Wrote" in outcome.summary_line


# --------------------------------------------------------------------------
# One tab per sampled item
# --------------------------------------------------------------------------


def _two_sample_results():
    r = _results(
        conclusion="not_satisfied",
        sample_results=[
            {"sample_id": "1", "conclusion": "satisfied"},
            {"sample_id": "2", "conclusion": "not_satisfied", "note": "no approval"},
        ],
        attribute_results=[
            {"attribute": "Invoice agrees", "sample_id": "1", "result": "satisfied",
             "value_observed": "21022452", "evidence_ids": ["ev_0001"]},
            {"attribute": "Invoice agrees", "sample_id": "2", "result": "satisfied",
             "value_observed": "21022453", "evidence_ids": ["ev_0001"]},
        ],
    )
    return r


def _two_manifests():
    return {
        "TS-1": SamplePopulationManifest(
            test_step_id="TS-1", sample_size=2,
            samples=[
                SampleItem(sample_id="1", test_step_id="TS-1", identifying_details="a",
                           key_fields={"Invoice Number": "21022452"}),
                SampleItem(sample_id="2", test_step_id="TS-1", identifying_details="b",
                           key_fields={"Invoice Number": "21022453"}),
            ],
        )
    }


def test_each_sampled_item_gets_its_own_tab(tmp_path: Path):
    # The prior-year workpaper documents ONE selection per Sample tab, so
    # two selections mean two tabs shaped like that one -- not two rows
    # crammed onto it, which is what stacked their exhibits and produced
    # the large empty regions a reviewer flagged.
    _template(tmp_path / "PY.xlsx")
    _real_pdf(tmp_path / "invoice.pdf")
    outcome = build_cy_workpaper(
        _spec(), _two_sample_results(), "PY.xlsx", tmp_path / "out",
        support_dir=tmp_path, sample_manifests=_two_manifests(),
    )
    wb = openpyxl.load_workbook(outcome.path)
    assert "Sample 1" in wb.sheetnames and "Sample 2" in wb.sheetnames
    # Sibling tabs stay adjacent. Appending clones at the end scattered them
    # after Population and Parameters, which reads as a mistake in a
    # workpaper a reviewer pages through in order.
    names = wb.sheetnames
    assert names.index("Sample 2") == names.index("Sample 1") + 1
    wb.close()


def test_a_single_sample_keeps_the_templates_own_tab_name(tmp_path: Path):
    # The ordinary case must stay indistinguishable from last year's file.
    _template(tmp_path / "PY.xlsx")
    _real_pdf(tmp_path / "invoice.pdf")
    outcome = build_cy_workpaper(
        _spec(), _results(), "PY.xlsx", tmp_path / "out",
        support_dir=tmp_path, sample_manifests=_manifests(),
    )
    wb = openpyxl.load_workbook(outcome.path)
    assert "Sample" in wb.sheetnames
    assert not any(n.startswith("Sample ") for n in wb.sheetnames)
    wb.close()


def test_every_summary_link_points_at_a_tab_that_exists(tmp_path: Path):
    # Tab names and summary links are derived from the same sample_id for
    # exactly this reason. A first version named tabs after the check number
    # while the links used the id, so every link pointed at a tab that did
    # not exist -- and Excel does not complain, the click just does nothing.
    _template(tmp_path / "PY.xlsx")
    _real_pdf(tmp_path / "invoice.pdf")
    outcome = build_cy_workpaper(
        _spec(), _two_sample_results(), "PY.xlsx", tmp_path / "out",
        support_dir=tmp_path, sample_manifests=_two_manifests(),
    )
    wb = openpyxl.load_workbook(outcome.path)
    ws = wb["Summary"]
    targets = [
        ws.cell(r, 5).hyperlink.location.split("!")[0].strip("'")
        for r in range(2, ws.max_row + 1)
        if ws.cell(r, 5).hyperlink
    ]
    assert targets
    assert all(t in wb.sheetnames for t in targets), (targets, wb.sheetnames)
    wb.close()


# --------------------------------------------------------------------------
# A previous CY workpaper as the template
#
# A real run uploaded last cycle's OUTPUT as the prior year. Its "Sample"
# tab had become "Sample 1"/"Sample 2", Excel had compacted its formats on
# save, and the build fell back to the standalone layout -- no PY format,
# no tickmark boxes -- with the reason in small grey text.
# --------------------------------------------------------------------------

import re
import zipfile

from agent.build_cy_workpaper import check_py_template


def _real_pdf_saying(path: Path, text: str) -> Path:
    from reportlab.lib.pagesizes import letter as PAGE
    from reportlab.pdfgen import canvas

    c = canvas.Canvas(str(path), pagesize=PAGE)
    c.setFont("Helvetica", 12)
    c.drawString(72, 700, text)
    c.save()
    return path


def _previous_output(tmp_path: Path) -> Path:
    """A real two-sample output of this builder, to use as next cycle's PY."""
    _template(tmp_path / "PY.xlsx")
    _real_pdf(tmp_path / "invoice.pdf")
    first = build_cy_workpaper(
        _spec(), _two_sample_results(), "PY.xlsx", tmp_path / "cycle1",
        support_dir=tmp_path, sample_manifests=_two_manifests(),
    )
    assert first.mirrored_py
    prior = tmp_path / "PRIOR.xlsx"
    prior.write_bytes(first.path.read_bytes())
    return prior


def test_a_previous_cy_workpaper_is_accepted_as_the_prior_year(tmp_path: Path):
    _previous_output(tmp_path)
    ok, why, notes = check_py_template(tmp_path / "PRIOR.xlsx")
    assert ok, why
    # And the run is told what was done to it, not left to discover it.
    assert any("previous CY workpaper" in n for n in notes)

    outcome = build_cy_workpaper(
        _spec(), _results(), "PRIOR.xlsx", tmp_path / "cycle2",
        support_dir=tmp_path, sample_manifests=_manifests(),
    )
    assert outcome.mirrored_py, outcome.fallback_reason
    wb = openpyxl.load_workbook(outcome.path)
    names = wb.sheetnames
    wb.close()
    # One sample this time: the plain tab name, and NONE of last cycle's
    # numbered tabs -- those would carry last cycle's items and exhibits.
    assert "Sample" in names
    assert not any(n.startswith("Sample ") for n in names), names
    # Exactly one Summary: the previous run's was replaced, not duplicated.
    assert names.count("Summary") == 1 and names[0] == "Summary"


def test_a_previous_cy_workpaper_with_more_samples_this_cycle(tmp_path: Path):
    _previous_output(tmp_path)
    outcome = build_cy_workpaper(
        _spec(), _two_sample_results(), "PRIOR.xlsx", tmp_path / "cycle2",
        support_dir=tmp_path, sample_manifests=_two_manifests(),
    )
    assert outcome.mirrored_py
    wb = openpyxl.load_workbook(outcome.path)
    assert [n for n in wb.sheetnames if n.startswith("Sample")] == ["Sample 1", "Sample 2"]
    wb.close()


def test_a_template_with_few_cell_formats_is_still_mirrored(tmp_path: Path):
    # Excel compacts styles.xml on every save, so a file someone has opened
    # and saved can have far fewer formats than the one the builder was
    # written against. Formats are now copied from the template's own cells
    # or appended, so the count no longer matters.
    wb = openpyxl.Workbook()
    wb.active.title = "Sample"
    wb.active["A1"] = "Invoice Number"
    wb.create_sheet("Population")["A5"] = "Invoice Number"
    wb.save(tmp_path / "PY.xlsx")
    _real_pdf(tmp_path / "invoice.pdf")
    outcome = build_cy_workpaper(
        _spec(), _results(), "PY.xlsx", tmp_path / "out",
        support_dir=tmp_path, sample_manifests=_manifests(),
    )
    assert outcome.mirrored_py, outcome.fallback_reason


def test_written_cells_take_the_templates_own_format_for_that_column(tmp_path: Path):
    # Every PY has its own styles.xml, so "format #2" means something
    # different in each. The sampled row's cells must look like the PY's.
    path = _template(tmp_path / "PY.xlsx")
    wb = openpyxl.load_workbook(path)
    wb["Sample"]["A2"].number_format = "@"
    wb.save(path)
    _real_pdf(tmp_path / "invoice.pdf")
    outcome = build_cy_workpaper(
        _spec(), _results(), "PY.xlsx", tmp_path / "out",
        support_dir=tmp_path, sample_manifests=_manifests(),
    )
    out = openpyxl.load_workbook(outcome.path)
    assert out["Sample"]["A2"].number_format == "@"
    out.close()


def test_each_sample_tab_keeps_its_own_exhibit_images(tmp_path: Path):
    # Every tab's drawing numbered its images from image1.png into ONE
    # shared folder, so sample 2's pages overwrote sample 1's.
    _template(tmp_path / "PY.xlsx")
    _real_pdf_saying(tmp_path / "invoice.pdf", "Invoice Number: 21022452")
    _real_pdf_saying(tmp_path / "invoice2.pdf", "Invoice Number: 21022453 SECOND")
    results = _two_sample_results()
    concl = results["TS-1"]["conclusion"]
    citations = [
        {"evidence_id": "ev_0001", "source_file": "invoice.pdf", "location": "invoice.pdf p.1",
         "quote_or_summary": "21022452", "relevance": "invoice", "sample_id": "1"},
        {"evidence_id": "ev_0002", "source_file": "invoice2.pdf", "location": "invoice2.pdf p.1",
         "quote_or_summary": "21022453", "relevance": "invoice", "sample_id": "2"},
    ]
    attrs = [dict(a.model_dump(), evidence_ids=["ev_0002"] if a.sample_id == "2" else ["ev_0001"])
             for a in concl.attribute_results]
    results = _results(
        conclusion=concl.conclusion,
        sample_results=[s.model_dump() for s in concl.sample_results],
        attribute_results=attrs,
        evidence_citations=citations,
    )
    outcome = build_cy_workpaper(
        _spec(), results, "PY.xlsx", tmp_path / "out",
        support_dir=tmp_path, sample_manifests=_two_manifests(),
    )
    z = zipfile.ZipFile(outcome.path)
    per_drawing = [
        tuple(re.findall(r'media/([^"]+)', z.read(n).decode()))
        for n in z.namelist() if n.startswith("xl/drawings/_rels/")
    ]
    images = [m for refs in per_drawing for m in refs]
    assert len(images) == 2, per_drawing
    assert len(set(images)) == 2, f"two tabs share one image file: {per_drawing}"
    assert z.read(f"xl/media/{images[0]}") != z.read(f"xl/media/{images[1]}")


def test_population_print_settings_stay_on_the_population_tab(tmp_path: Path):
    # Print titles are scoped to a tab by POSITION. Inserting the Summary
    # in front without shifting moved them onto the tab to the left.
    path = _template(tmp_path / "PY.xlsx")
    wb = openpyxl.load_workbook(path)
    wb["Population"].print_title_rows = "5:5"
    wb.save(path)
    _real_pdf(tmp_path / "invoice.pdf")
    outcome = build_cy_workpaper(
        _spec(), _two_sample_results(), "PY.xlsx", tmp_path / "out",
        support_dir=tmp_path, sample_manifests=_two_manifests(),
    )
    wbx = zipfile.ZipFile(outcome.path).read("xl/workbook.xml").decode()
    order = re.findall(r'<sheet name="([^"]+)"', wbx)
    scoped = re.findall(r'<definedName name="_xlnm.Print_Titles" localSheetId="(\d+)"', wbx)
    assert scoped and order[int(scoped[0])] == "Population", (order, scoped)


def test_the_mirrored_workpaper_is_stamped_draft(tmp_path: Path):
    # The standalone layout always carried a DRAFT banner; the mirrored
    # one -- the file that looks most finished -- carried none.
    _template(tmp_path / "PY.xlsx")
    _real_pdf(tmp_path / "invoice.pdf")
    outcome = build_cy_workpaper(
        _spec(), _results(), "PY.xlsx", tmp_path / "out",
        support_dir=tmp_path, sample_manifests=_manifests(),
    )
    assert outcome.mirrored_py
    assert "DRAFT" in outcome.path.name
    wb = openpyxl.load_workbook(outcome.path)
    assert "DRAFT" in wb["Summary"]["A1"].value
    wb.close()
    z = zipfile.ZipFile(outcome.path)
    sheets = [n for n in z.namelist() if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n)]
    assert all("DRAFT" in z.read(n).decode() for n in sheets)


def test_the_workbook_opens_on_one_tab_not_a_group(tmp_path: Path):
    # Two tabs flagged selected opens them GROUPED, and typing then edits both.
    # A file saved by Excel always has one tab flagged (the one it was on);
    # openpyxl's fixture has none unless told, which made this test unable
    # to fail.
    path = _template(tmp_path / "PY.xlsx")
    wb = openpyxl.load_workbook(path)
    wb["Population"].sheet_view.tabSelected = True
    wb.save(path)
    _real_pdf(tmp_path / "invoice.pdf")
    outcome = build_cy_workpaper(
        _spec(), _two_sample_results(), "PY.xlsx", tmp_path / "out",
        support_dir=tmp_path, sample_manifests=_two_manifests(),
    )
    z = zipfile.ZipFile(outcome.path)
    selected = [n for n in z.namelist()
                if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n) and 'tabSelected="1"' in z.read(n).decode()]
    assert len(selected) == 1


def test_a_templates_own_summary_tab_is_never_deleted(tmp_path: Path):
    # Our previous Summary is recognised by its headings, not its name; a
    # PY's own "Summary" is prior-year content and must survive.
    path = _template(tmp_path / "PY.xlsx")
    wb = openpyxl.load_workbook(path)
    wb.create_sheet("Summary", 0)["A1"] = "Control summary prepared by the PY tester"
    wb.save(path)
    _real_pdf(tmp_path / "invoice.pdf")
    outcome = build_cy_workpaper(
        _spec(), _results(), "PY.xlsx", tmp_path / "out",
        support_dir=tmp_path, sample_manifests=_manifests(),
    )
    out = openpyxl.load_workbook(outcome.path)
    assert out["Summary"]["A1"].value == "Control summary prepared by the PY tester"
    assert "CY Summary" in out.sheetnames
    out.close()


def test_a_sample_selections_tab_is_not_mistaken_for_a_sample_tab(tmp_path: Path):
    # Only "Sample <something with a digit>" is a numbered sample tab.
    path = _template(tmp_path / "PY.xlsx", sample="Sample 1")
    wb = openpyxl.load_workbook(path)
    wb.create_sheet("Sample Selections")["A1"] = "keep me"
    wb.save(path)
    _real_pdf(tmp_path / "invoice.pdf")
    outcome = build_cy_workpaper(
        _spec(), _results(), "PY.xlsx", tmp_path / "out",
        support_dir=tmp_path, sample_manifests=_manifests(),
    )
    assert outcome.mirrored_py
    out = openpyxl.load_workbook(outcome.path)
    assert out["Sample Selections"]["A1"].value == "keep me"
    out.close()


def test_an_unusable_prior_year_is_reported_before_the_run_and_after_it(tmp_path: Path):
    (tmp_path / "PY.pdf").write_bytes(b"%PDF-1.4\n")
    ok, why, _notes = check_py_template(tmp_path / "PY.pdf")
    assert not ok and ".pdf" in why
    outcome = build_cy_workpaper(
        _spec(), _results(), "PY.pdf", tmp_path / "out",
        support_dir=tmp_path, sample_manifests=_manifests(),
    )
    assert not outcome.mirrored_py
    assert ".pdf" in outcome.fallback_reason
    assert "NOT last year's format" in outcome.summary_line
