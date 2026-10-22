"""Build module_summary.docx and module_summary.pdf (the ML Analysis Report) from metrics.json.

No APA template (.docx) was provided with this project, so the document is created in APA 7
student-paper format: US Letter, 1-inch margins, Times New Roman 12 pt, double spacing,
page numbers top right, a title page, bold centred Level 1 headings, APA-style tables and
figures, and a References page with hanging indents.

Every number in the text is read from metrics.json (written by modeling.ipynb); none is typed by hand.

Usage:  python build_report.py
"""
import datetime as dt
import json
import re
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt

ROOT = Path(__file__).resolve().parent
METRICS = json.loads((ROOT / "metrics.json").read_text(encoding="utf-8"))
FIG = ROOT / "figures"
DOCX_OUT, PDF_OUT = ROOT / "module_summary.docx", ROOT / "module_summary.pdf"

STUDENT_NAME = "Mohammad Afsharnejad"
TITLE = "Predicting Term-Deposit Subscription From Bank Telemarketing Data"


# ----------------------------------------------------------------------------- number helpers
def f3(x):
    """Format a metric with three decimals."""
    return f"{x:.3f}"


def pct(x, d=1):
    """Format a proportion (0-1) as a percentage string."""
    return f"{x * 100:.{d}f}%"


def per100(x):
    """Round a proportion to 'out of 100' for plain-language text."""
    return f"{round(x * 100)}"


def signed(x):
    """Format a small signed value with three decimals and a typographic minus sign."""
    return f"{x:.3f}".replace("-", "−")


def count_words(n):
    """Write a zero count as 'no' (APA prefers words over '0' in running text)."""
    return "no" if n == 0 else f"{n:,}"


# ----------------------------------------------------------------------------- document helpers
def set_base_styles(doc):
    """Times New Roman 12 pt, double spacing, no extra space before/after (APA 7)."""
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(12)
    style.element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    pf = style.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.DOUBLE
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)


def set_page_layout(section):
    """US Letter with 1-inch margins and a right-aligned page number in the header."""
    section.page_width, section.page_height = Inches(8.5), Inches(11)
    for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(section, side, Inches(1))
    section.header_distance = Inches(0.5)
    hp = section.header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = hp.add_run()
    for tag, text in (("begin", None), (None, "PAGE"), ("end", None)):
        if tag:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), tag)
        else:
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = text
        run._r.append(el)


def add_runs(par, text):
    """Add text to a paragraph, rendering *asterisk* spans in italics."""
    for i, chunk in enumerate(re.split(r"\*(.+?)\*", text)):
        if chunk:
            par.add_run(chunk).italic = bool(i % 2)
    return par


def body(doc, text, indent=True):
    """APA body paragraph: first-line indent 0.5 in, left aligned."""
    p = add_runs(doc.add_paragraph(), text)
    if indent:
        p.paragraph_format.first_line_indent = Inches(0.5)
    return p


def centered(doc, text="", bold=False):
    """Centred paragraph (title page lines, headings)."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if text:
        p.add_run(text).bold = bold
    return p


def heading1(doc, text):
    """APA Level 1 heading: centred, bold, title case."""
    p = centered(doc, text, bold=True)
    p.paragraph_format.keep_with_next = True
    return p


def label_and_title(doc, label, title):
    """APA table/figure label (bold) and title (italic, title case) on separate lines."""
    p1 = doc.add_paragraph()
    p1.add_run(label).bold = True
    p1.paragraph_format.keep_with_next = True
    p2 = doc.add_paragraph()
    p2.add_run(title).italic = True
    p2.paragraph_format.keep_with_next = True


def note(doc, text):
    """APA note below a table or figure: 'Note.' in italics, then the text."""
    p = doc.add_paragraph()
    p.add_run("Note.").italic = True
    add_runs(p, " " + text)
    return p


def set_cell_border(cell, **edges):
    """Set borders on a table cell, e.g. top='single', bottom='single'."""
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge, val in edges.items():
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), val)
        el.set(qn("w:sz"), "8")
        el.set(qn("w:color"), "000000")
        borders.append(el)


def apa_table(doc, header, rows, col_widths):
    """APA 7 table: horizontal rules above/below the header and at the bottom, no vertical lines."""
    table = doc.add_table(rows=1 + len(rows), cols=len(header))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r_idx, values in enumerate([header] + rows):
        for c_idx, val in enumerate(values):
            cell = table.cell(r_idx, c_idx)
            cell.width = Inches(col_widths[c_idx])
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx == 0 else WD_ALIGN_PARAGRAPH.CENTER
            p.add_run(str(val))
            if r_idx == 0:
                set_cell_border(cell, top="single", bottom="single")
            if r_idx == len(rows):
                set_cell_border(cell, bottom="single")
    return table


def figure(doc, number, title, image, width, note_text):
    """APA 7 figure: bold label, italic title, image, then a note."""
    label_and_title(doc, f"Figure {number}", title)
    p = doc.add_paragraph()
    p.paragraph_format.keep_with_next = True
    p.add_run().add_picture(str(image), width=Inches(width))
    note(doc, note_text)


def reference(doc, text):
    """References entry with a 0.5-inch hanging indent."""
    p = add_runs(doc.add_paragraph(), text)
    p.paragraph_format.left_indent = Inches(0.5)
    p.paragraph_format.first_line_indent = Inches(-0.5)
    return p


# ----------------------------------------------------------------------------- pull numbers
m = METRICS
ds, sp, pre, gs = m["dataset"], m["split"], m["preprocessing"], m["grid_search"]
dummy, lr = m["test"]["dummy"], m["test"]["logistic_regression"]
rf05, rf = m["test"]["random_forest_default"], m["test"]["random_forest_tuned"]
cm = m["confusion_matrix"]["random_forest_tuned"]
cv_rf = m["cv"]["Random forest"]["pr_auc"]
thr = m["threshold"]["value"]
prev = sp["test_positive_rate"]
lift = rf["pr_auc"] / prev
tc = m["temporal_check"]
t_lift = tc["pr_auc"] / tc["test_positive_rate"]
imp = {r["feature"]: r["mean"] for r in m["permutation_importance"]}
top3 = [r["feature"] for r in m["permutation_importance"][:3]]
ages = {r["age_band"]: r for r in m["subgroup_age"]}
prec_by_age = [r["precision"] for r in m["subgroup_age"]]
flagged = cm["tp"] + cm["fp"]
n_pos_test, n_neg_test = cm["tp"] + cm["fn"], cm["tn"] + cm["fp"]
grid_rows = gs["results"]
most_reg = next(r for r in grid_rows if r["max_depth"] == 12 and r["min_samples_leaf"] == 20)
of = m["overfitting"]
bp = gs["best_params"]
depth_txt = "unlimited depth" if bp["max_depth"] is None else f"a maximum depth of {bp['max_depth']}"
assert top3 == ["month", "contact", "poutcome"], "Text below describes these three features; re-check if they change."

# ----------------------------------------------------------------------------- build document
doc = Document()
set_base_styles(doc)
set_page_layout(doc.sections[0])

# Title page (APA 7 student paper)
for _ in range(3):
    centered(doc)
centered(doc, TITLE, bold=True)
centered(doc)
for line in (STUDENT_NAME,
             "Udacity and Woolf University, Master's Degree in Artificial Intelligence",
             "Capstone Project 3: Applied Machine Learning: Model Design, Training, and Performance Evaluation",
             dt.date.today().strftime("%B %d, %Y").replace(" 0", " ")):
    centered(doc, line)
doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

# First page of text repeats the title
centered(doc, TITLE, bold=True)

# 1. Overview
heading1(doc, "Overview")
body(doc, (
    "This report describes a model that predicts, before a marketing call is made, whether a client of a "
    "Portuguese bank will subscribe to a term deposit. The problem type is supervised learning – binary "
    "classification, with the target *y* (yes or no). The data are the Bank Marketing dataset from the UCI "
    f"Machine Learning Repository ({ds['url']}), and the main model is a class-weighted random forest that is "
    "compared with a majority-class dummy and a logistic regression baseline. On held-out clients, the forest "
    f"ranks likely subscribers about {lift:.1f} times better than a random call list (PR-AUC {f3(rf['pr_auc'])} "
    f"against a baseline of {f3(prev)})."
))

# 2. Dataset Description
heading1(doc, "Dataset Description")
body(doc, (
    f"The data record {ds['n_rows']:,} clients of a Portuguese bank who were contacted by phone in direct-marketing "
    "campaigns between May 2008 and November 2010 (Moro et al., 2014). The file bank-full.csv has "
    f"{ds['n_columns']} columns: {ds['n_input_features']} input features and the target *y*, which records whether "
    "the client subscribed to a term deposit. The features describe the client (age, job, marital status, "
    "education, credit default, average yearly balance, housing loan and personal loan), the last contact of the "
    "current campaign (contact type, day, month, call duration and number of contacts) and the previous campaign "
    "(days since the last contact, number of earlier contacts and the earlier outcome). Only "
    f"{ds['n_positive']:,} clients ({pct(ds['positive_rate'])}) subscribed, so the classes are strongly imbalanced. "
    f"The file has {count_words(ds['missing_values'])} missing cells and {count_words(ds['duplicates'])} duplicate "
    "rows, but the placeholder "
    f"“unknown” fills {ds['unknown_pct']['poutcome']:.1f}% of the previous-outcome values and "
    f"{ds['unknown_pct']['contact']:.1f}% of the contact-type values. The dataset is public under a "
    f"{ds['license']} license."
))

# 3. Modeling Approach
heading1(doc, "Modeling Approach")
body(doc, (
    "The call-duration column was removed before modeling. Moro et al. (2014) note that call duration strongly "
    "affects the outcome but is only known after the call has ended, so it must be excluded from a realistic "
    "predictive model. A binary flag for previously contacted clients was added, because the days-since-last-contact "
    f"feature uses −1 as a code for “never contacted” ({ds['pdays_never_contacted_pct']:.1f}% of clients) "
    f"rather than a real number of days. The data were then split into a stratified training set ({sp['n_train']:,} "
    f"clients) and test set ({sp['n_test']:,} clients), and both sets keep the {pct(prev)} subscription rate. All "
    "preprocessing "
    f"(imputation, one-hot encoding of the {pre['n_categorical']} categorical features and standardization of the "
    f"{pre['n_numeric']} numeric features, giving {pre['n_encoded_features']} model inputs) was placed inside a "
    "scikit-learn pipeline. The pipeline learns every statistic from training data only, including inside each "
    "cross-validation fold, which is the recommended defense against data leakage (scikit-learn developers, 2026). "
    "The “unknown” labels were kept as their own category, because a missing record of a previous outcome is "
    "informative in itself."
))
body(doc, (
    "A random forest was chosen as the main model because the inputs mix numeric and categorical features, and "
    "the useful signals, such as call timing and the outcome of earlier campaigns, may interact in non-linear ways. "
    "Single decision trees handle such inputs with little preparation but overfit and change a great deal with small "
    "changes in the data. A random forest reduces this variance by averaging many trees, each grown on a bootstrap "
    "sample with a random subset of features considered at each split, at the cost of a slight increase in bias "
    "(scikit-learn developers, 2026). Both the forest and the logistic regression weighted the classes in inverse "
    "proportion to their frequency, so the rare subscribers were not ignored. The forest's depth and minimum leaf "
    f"size were tuned with a {gs['n_settings']}-setting grid search over {gs['n_folds']}-fold stratified "
    f"cross-validation ({gs['n_estimators']} trees). The best setting used {depth_txt} and at least "
    f"{bp['min_samples_leaf']} clients per leaf. The pipeline structure and the permutation-importance procedure "
    "were adapted from examples in the scikit-learn user guide."
))
body(doc, (
    f"Accuracy is not a useful headline metric here, because always answering “no” is already "
    f"{pct(dummy['accuracy'])} accurate. The evaluation therefore centers on precision (the share of flagged clients "
    "who subscribe), recall (the share of subscribers who are flagged), their harmonic mean, F1, and the area under "
    "the precision–recall curve (PR-AUC, computed as average precision). Saito and Rehmsmeier (2015) show that "
    "when positives are rare, precision–recall analysis describes a classifier's practical performance more "
    "faithfully than the ROC curve, whose false-positive rate is diluted by the large number of negatives. ROC-AUC "
    "is still reported for comparison. PR-AUC was also the score optimized during tuning, because the business task "
    "is to rank clients so that the likeliest subscribers are called first. The decision threshold "
    f"({thr:.2f}) was chosen to maximize F1 on out-of-fold training predictions, so the test set played no part in it."
))
body(doc, (
    "The random forest makes no assumption of linear relationships, normal distributions or equal feature scales, "
    "and tree-based models do not need normalized inputs (scikit-learn developers, 2026). Logistic regression, by "
    "contrast, models the log-odds of subscribing as a linear combination of the inputs, which limits how well it "
    "captures interactions. Both models assume that the clients they will score in the future come from the same "
    "population as the training clients. Mehrabi et al. (2021) describe two ways this assumption can fail: "
    "representation bias, when the sample does not reflect the target population, and temporal bias, when "
    "populations and behaviors change over time."
))

# 4. Results
heading1(doc, "Results")
body(doc, (
    f"Table 1 compares the models on the {sp['n_test']:,} held-out test clients. The dummy has the highest accuracy "
    f"({f3(dummy['accuracy'])}) but finds no subscribers at all. The random forest beats logistic regression on both "
    f"ranking metrics (PR-AUC {f3(rf['pr_auc'])} vs. {f3(lr['pr_auc'])}; ROC-AUC {f3(rf['roc_auc'])} vs. "
    f"{f3(lr['roc_auc'])}) and on F1 ({f3(rf['f1'])} vs. {f3(lr['f1'])}). Logistic regression flags more clients, "
    f"which gives it higher recall ({f3(lr['recall'])}) but much lower precision ({f3(lr['precision'])}). The forest's "
    f"test PR-AUC is close to its cross-validated value ({f3(cv_rf['mean'])} ± {f3(cv_rf['std'])}), so the "
    f"estimate is stable. Moving the threshold from 0.50 to {thr:.2f} changed little: recall rose by "
    f"{rf['recall'] - rf05['recall']:.3f} and precision fell by {rf05['precision'] - rf['precision']:.3f}."
))
label_and_title(doc, "Table 1", "Test-Set Performance of the Baseline and Main Models")
apa_table(
    doc,
    ["Model", "Accuracy", "Precision", "Recall", "F1", "ROC‑AUC", "PR‑AUC"],   # non-breaking hyphens
    [[name] + [f3(r[k]) for k in ("accuracy", "precision", "recall", "f1", "roc_auc", "pr_auc")]
     for name, r in [("Dummy (majority class)", dummy), ("Logistic regression", lr),
                     ("Random forest, t = 0.50", rf05), (f"Random forest, t = {thr:.2f}", rf)]],
    [1.85, 0.8, 0.8, 0.65, 0.55, 0.85, 0.8],
)
note(doc, (
    f"n = {sp['n_test']:,} test clients, of whom {pct(prev)} subscribed. t = decision threshold; the tuned threshold "
    "was chosen on out-of-fold training predictions. PR-AUC = average precision; a random ranking scores the "
    f"subscription rate ({f3(prev)})."
))
body(doc, (
    f"Figure 1 shows that at the tuned threshold the forest flags {flagged:,} of the {sp['n_test']:,} test clients "
    f"({pct(flagged / sp['n_test'])}), and {cm['tp']:,} of them subscribe. It misses {cm['fn']:,} of the "
    f"{n_pos_test:,} subscribers ({pct(cm['fn'] / n_pos_test)}) and wrongly flags {cm['fp']:,} of the "
    f"{n_neg_test:,} non-subscribers ({pct(cm['fp'] / n_neg_test)})."
))
figure(doc, 1, "Confusion Matrices of the Random Forest on the Test Set",
       FIG / "02_confusion_matrix.png", 6.3,
       "Left: default threshold of 0.50. Right: tuned threshold. Percentages are shares of each actual class (row).")
body(doc, (
    "Figure 2 shows that the forest's ROC curve lies above that of logistic regression almost everywhere. The "
    "precision–recall panel shows the cost of finding more subscribers: precision falls steadily as recall rises, so "
    f"no threshold gives high precision and high recall together. At the chosen operating point, a recall of "
    f"{f3(rf['recall'])} comes with a precision of {f3(rf['precision'])}, which means most flagged clients still do "
    f"not subscribe even though the ROC-AUC of {f3(rf['roc_auc'])} looks strong."
))
figure(doc, 2, "ROC and Precision–Recall Curves of the Three Models on the Test Set",
       FIG / "03_roc_pr_curves.png", 6.5,
       "The dashed lines show random guessing (ROC) and the subscription rate (precision–recall). The dot marks the "
       "random forest's tuned operating point.")
body(doc, (
    f"Figure 3 shows that the forest relies mostly on the month of contact (mean drop in PR-AUC "
    f"{imp['month']:.3f} when shuffled), the contact type ({imp['contact']:.3f}) and the outcome of the previous "
    f"campaign ({imp['poutcome']:.3f}). Client characteristics such as age ({signed(imp['age'])}), education "
    f"({signed(imp['education'])}) and job ({signed(imp['job'])}) add almost nothing. The model has therefore learned mainly "
    "when and how the bank called and how the client responded before, rather than who the client is."
))
figure(doc, 3, "Permutation Feature Importance of the Random Forest on the Test Set",
       FIG / "04_feature_importance.png", 5.6,
       "Importance is the mean drop in test-set PR-AUC when one original column is randomly shuffled, over 10 "
       "repeats; error bars show ±1 standard deviation.")

# 5. Interpretation for a Non-Technical Audience
heading1(doc, "Interpretation for a Non-Technical Audience")
body(doc, (
    f"Imagine the bank has a list of 100 clients it could call. Without the model, about {per100(prev)} of them would "
    f"say yes. If the bank calls only the clients the model picks out, about {per100(rf['precision'])} out of every "
    f"100 calls end with a new deposit, roughly {rf['precision'] / prev:.1f} times as many as calling at random. "
    f"The model is far from perfect: it still overlooks about {per100(1 - rf['recall'])} out of every 100 clients "
    f"who would have said yes, and more than half of the clients it picks out (about "
    f"{per100(1 - rf['precision'])} in 100) still say no."
))
body(doc, (
    "The model is therefore best used to decide whom to call first, not to decide that some clients should never "
    "be called. Its most useful clues are the time of year of the call, whether the client was reached on a mobile "
    "phone or a landline (or the bank did not record which), and whether the client said yes to an earlier "
    "campaign. Personal "
    "details such as age or job matter much less. One caution matters for planning: the model learned from calls "
    "made between 2008 and 2010, and when it was tested only on the most recent of those calls, it was clearly "
    "less reliable. It should be rechecked on new calls before the bank depends on it."
))

# 6. Limitations and Potential Bias
heading1(doc, "Limitations and Potential Bias")
body(doc, (
    "The most important limitation is that the headline results come from a random split of data that are ordered "
    "in time. When the same tuned model was trained on the earliest 80% of rows and tested on the latest 20%, the "
    f"subscription rate rose from {pct(tc['train_positive_rate'])} to {pct(tc['test_positive_rate'])} between the two "
    f"periods, and ROC-AUC fell from {f3(rf['roc_auc'])} to {f3(tc['roc_auc'])}. PR-AUC on the later period "
    f"({f3(tc['pr_auc'])}) looks similar, but it is only {t_lift:.1f} times its new baseline of "
    f"{f3(tc['test_positive_rate'])}, compared with {lift:.1f} times under the random split. This matches what "
    "Mehrabi et al. (2021) call temporal bias, and it means Table 1 overstates how well the model would forecast "
    "a future campaign. The heavy reliance on month and day of contact in Figure 3 makes the model especially "
    "sensitive to changes in the campaign calendar."
))
body(doc, (
    f"A second limitation is variance. The forest scores a PR-AUC of {f3(of['train_pr_auc'])} on its own training "
    f"data against {f3(of['cv_pr_auc'])} in cross-validation, so its deep trees partly memorize the training clients. "
    f"The most regularized setting in the grid narrowed this gap (training PR-AUC {f3(most_reg['mean_train_pr_auc'])}) "
    f"but also lowered cross-validated PR-AUC to {f3(most_reg['mean_test_pr_auc'])}, so some extra variance was "
    "accepted in exchange for better ranking. The metrics themselves also have limits. The baseline of a "
    "precision–recall curve equals the share of positives (Saito & Rehmsmeier, 2015), so PR-AUC values cannot be "
    "compared directly across periods or datasets with different subscription rates. This is why the temporal check "
    "above is read against its own baseline."
))
r = ages
body(doc, (
    "A specific ethical concern is how the model treats clients of different ages. Table 2 shows that precision is "
    f"similar across age bands ({min(prec_by_age):.2f} to {max(prec_by_age):.2f}), but recall is not: subscribers "
    f"aged 60 or over are flagged {pct(r['60+']['recall'])} of the time, compared with {pct(r['30-44']['recall'])} "
    f"for ages 30 to 44 and {pct(r['45-59']['recall'])} for ages 45 to 59. The model flags "
    f"{pct(r['60+']['flagged_rate'])} of all clients aged 60 or over, against {pct(r['30-44']['flagged_rate'])} of "
    "those aged 30 to 44. Under the equal-opportunity criterion, which asks that truly positive people be selected "
    "at the same rate in every group (Mehrabi et al., 2021), the model disadvantages middle-aged subscribers. In "
    "practice, it would also direct a large share of sales calls toward older and retired clients, which raises the "
    "question of whether that group would face more pressure to buy financial products than others. The 60-and-over "
    f"band contains only {r['60+']['n']:,} test clients, so its estimates are less certain."
))
label_and_title(doc, "Table 2", "Random Forest Test-Set Performance by Age Band at the Tuned Threshold")
apa_table(
    doc,
    ["Age band", "n", "Subscription rate", "Share flagged", "Precision", "Recall"],
    [[b["age_band"].replace("-", "–"), f"{b['n']:,}", pct(b["actual_yes_rate"]), pct(b["flagged_rate"]),
      f3(b["precision"]), f3(b["recall"])] for b in m["subgroup_age"]],
    [1.0, 0.8, 1.3, 1.2, 0.9, 0.8],
)
note(doc, f"Threshold t = {thr:.2f}. Subscription rate is the actual share of subscribers in each band.")
body(doc, (
    "One step was taken and one is proposed to reduce this risk. The subgroup audit in Table 2 was run before any "
    "use of the model, so the gap is documented rather than hidden. The proposed step is to make the audit a "
    "release gate: at every retraining on recent data, recall and the share flagged are recomputed for each age "
    "band, and flagged clients aged 60 or over receive a human suitability check and a cap on how often they can "
    "be called."
))

# 7. References (new page, alphabetical, hanging indent)
doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
heading1(doc, "References")
reference(doc, (
    "Mehrabi, N., Morstatter, F., Saxena, N., Lerman, K., & Galstyan, A. (2021). A survey on bias and fairness in "
    "machine learning. *ACM Computing Surveys, 54*(6), 1–35. https://doi.org/10.1145/3457607"
))
reference(doc, (
    "Moro, S., Rita, P., & Cortez, P. (2014). *Bank marketing* [Dataset]. UCI Machine Learning Repository. "
    "https://doi.org/10.24432/C5K306"
))
reference(doc, (
    "Saito, T., & Rehmsmeier, M. (2015). The precision-recall plot is more informative than the ROC plot when "
    "evaluating binary classifiers on imbalanced datasets. *PLOS ONE, 10*(3), Article e0118432. "
    "https://doi.org/10.1371/journal.pone.0118432"
))
reference(doc, (
    "scikit-learn developers. (2026). *User guide* (Version 1.9.1). scikit-learn. "
    "https://scikit-learn.org/1.9/user_guide.html"
))

doc.save(DOCX_OUT)
print(f"Saved {DOCX_OUT.name}")

# ----------------------------------------------------------------------------- convert to PDF
try:
    from docx2pdf import convert          # uses Microsoft Word on Windows / macOS
    convert(str(DOCX_OUT), str(PDF_OUT))
except Exception as exc:                  # fall back to LibreOffice if Word is unavailable
    import shutil
    import subprocess
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise RuntimeError(f"PDF conversion failed ({exc}) and LibreOffice was not found.") from exc
    subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", str(ROOT), str(DOCX_OUT)], check=True)
print(f"Saved {PDF_OUT.name}")
