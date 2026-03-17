#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Jun  6 10:41:44 2025

vPowerAnalysisTool  v0.4.0

═══════════════════════════════════════════════════════
CHANGES FROM v0.3.0
═══════════════════════════════════════════════════════

STATISTICAL FIXES
  [S1] Cohen's f (correct metric) replaces averaged Cohen's d for all ANOVA analyses.
       FTestAnovaPower expects Cohen's f = σ_between / σ_within, not Cohen's d.
       Using d produced fundamentally wrong sample-size estimates.

  [S2] Cohen's d is now taken as abs().  A negative d (group2 > group1) caused
       TTestIndPower to receive an invalid effect size and return garbage or crash.

  [S3] Two-way ANOVA completely rewritten.
       • Proper cell-mean algebra decomposes pilot data into Factor A, Factor B,
         and A×B interaction Cohen's f values independently.
       • Each effect's required n is computed separately; the most conservative
         (largest) is reported as the design-driving requirement.
       • FTestAnovaPower is called with k_groups=2 for each 1-df effect.
         Returned nobs equals n per factor level (= 2 × n_per_cell for a 2×2
         design), so the result is divided by 2 to give n per cell.
       • FTestAnovaPower is an approximation for two-way designs (it ignores
         the shared error df structure); this is noted in the output and is
         consistent with how G*Power handles this case.

  [S4] Normality checking and transformation deferred to analysis time on
       selected columns only.  v0.3.0 transformed every column at file load,
       silently corrupting non-selected columns and computing effect sizes on
       transformed data without informing the user which columns were changed.

  [S5] Shapiro-Wilk guards added:
       • n < 3  → assume normal (scipy raises an error below 3).
       • n > 5000 → subsample to 5000 (scipy warns and gives unreliable p above
         5000; a subsample is sufficient for the test to be well-powered).

  [S6] .dropna() applied before every statistical calculation.  NaN values
       previously propagated silently through mean(), std(), shapiro(), etc.

  [S7] Biological-replicate assumption documented explicitly in the results panel
       and in the Excel report.  The formula (required_n / n_obs_in_pilot) only
       makes sense if each row in the pilot data represents one technical
       replicate from a single biological replicate; this is now stated clearly.

CODE / LOGIC FIXES
  [C1] Frame-based layout replaces hardcoded row numbers.  In v0.3.0, alpha/
       power/test-type widgets were pinned to rows 6-8; with ≥5 One-way ANOVA
       groups they overlapped those labels.

  [C2] One-way ANOVA helper widgets (num_groups entry + Update button) are now
       hidden when switching away from One-way ANOVA.

  [C3] All group dropdowns validated as non-empty and present in the loaded
       DataFrame before the analysis is run.

  [C4] Normality + transformation computed once per analysis run and cached.
       generate_excel_report (save_report) reads from cache instead of
       re-running the computations redundantly (and potentially inconsistently).

  [C5] transform_data() now returns (series, method_label) so the report can
       record which transformation was applied to each column.

  [C6] Non-numeric columns in the loaded file are skipped at load time with an
       informational warning, preventing silent KeyErrors during analysis.

UX IMPROVEMENTS
  [U1] Data preview table (first 5 rows, horizontally scrollable) is displayed
       immediately after a file is loaded.

  [U2] Per-column normality badge (Normal / Not normal / Failed) shown in the
       inline results panel after each analysis run.

  [U3] Results are displayed in a persistent inline panel instead of a
       messagebox that disappears as soon as the user clicks OK.

  [U4] "Perform Analysis" and "Save Excel Report" are separate buttons.
       The Save button is disabled until an analysis has been run.

  [U5] Alpha and Power entry fields are highlighted red when the value is
       outside the valid open interval (0, 1).

  [U6] Optional manual effect-size entry for prospective (a priori) power
       analysis when no pilot data is available.  A tooltip explains the
       conventional benchmarks (small / medium / large) for d and f.

  [U7] Power curve plot embedded directly in the window (requires matplotlib).
       Shows power vs. n-per-group with dashed lines at the target power and
       the required sample size.

  [U8] Scrollable main window with a minimum size so the layout never collapses
       when many groups are selected.

  [U9] Reset button clears all state and returns the tool to its initial
       condition without needing to restart the script.

  [U10] Tooltips on Alpha, Power, Effect Size, and Biological Replicates fields
        explain what each parameter means and give typical values.
        
@author: vincentpb
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from statsmodels.stats.power import TTestIndPower, FTestAnovaPower
import math
import os
import scipy.stats as stats
import numpy as np

try:
    import matplotlib
    matplotlib.use("TkAgg")
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    MATPLOTLIB_AVAILABLE = True
except ImportError:
    MATPLOTLIB_AVAILABLE = False

VERSION = "0.4.0"


# ══════════════════════════════════════════════════════════════════════════════
#  Tooltip widget
# ══════════════════════════════════════════════════════════════════════════════

class ToolTip:
    """Simple hover tooltip for any tkinter widget."""

    def __init__(self, widget, text):
        self.widget = widget
        self.text   = text
        self._win   = None
        widget.bind("<Enter>", self._show)
        widget.bind("<Leave>", self._hide)

    def _show(self, _=None):
        if self._win:
            return
        x = self.widget.winfo_rootx() + 22
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        self._win = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        tk.Label(tw, text=self.text, justify="left", background="#fffde7",
                 relief="solid", borderwidth=1, padx=5, pady=4,
                 font=("TkDefaultFont", 9), wraplength=360).pack()

    def _hide(self, _=None):
        if self._win:
            self._win.destroy()
            self._win = None


# ══════════════════════════════════════════════════════════════════════════════
#  Statistical helpers
# ══════════════════════════════════════════════════════════════════════════════

def is_normal(series, alpha=0.05):
    """
    Shapiro-Wilk normality test.
    [S5] Guards: n<3 → True (cannot test); n>5000 → subsample to 5000.
    """
    s = series.dropna()                  # [S6]
    n = len(s)
    if n < 3:
        return True
    if n > 5000:
        s = s.sample(5000, random_state=42)
    _, p = stats.shapiro(s)
    return p > alpha


def transform_data(series):
    """
    Try log then Box-Cox transformation.
    [C5] Returns (transformed_series, label) where label in {'log','Box-Cox','none'}.

    The shift (s - min + 1) is applied ONLY when the data contains non-positive
    values.  Applying it unconditionally to already-positive data (e.g. a true
    log-normal distribution) changes the shape of the distribution and prevents
    the log transform from restoring normality.
    """
    s = series.dropna()                  # [S6]

    # Shift to strictly > 0 only if necessary
    if s.min() <= 0:
        s_pos = s - s.min() + 1
    else:
        s_pos = s

    # ── Log ──────────────────────────────────────────────────────────────────
    try:
        log_t = pd.Series(np.log(s_pos.values))
        if log_t.std() > 0 and is_normal(log_t):    # guard against constant result
            return log_t, "log"
    except Exception:
        pass

    # ── Box-Cox (requires all-positive values) ────────────────────────────────
    try:
        bc_vals, _ = stats.boxcox(s_pos.values)
        bc_t = pd.Series(bc_vals)
        if bc_t.std() > 0 and is_normal(bc_t):      # guard against constant result
            return bc_t, "Box-Cox"
    except Exception:
        pass

    return s.reset_index(drop=True), "none"


def prepare_column(original_data, col):
    """
    [S4] Check normality of the raw column; transform only if needed.
    Returns (processed_series, normality_label, transform_label).
    Normality is assessed on this column alone; nothing is written back to the
    DataFrame so other columns are never silently altered.
    """
    raw = original_data[col].dropna()    # [S6]
    if is_normal(raw):
        return raw.reset_index(drop=True), "Normal", "—"
    transformed, method = transform_data(raw)
    if method != "none":
        return transformed, "Not normal", method
    return raw.reset_index(drop=True), "Not normal", "Failed — used raw data"


# ── Effect-size calculators ───────────────────────────────────────────────────

def cohen_d(g1, g2):
    """
    [S1, S2] Pooled, absolute Cohen's d for an independent-samples t-test.
    Always returns a non-negative float (required by TTestIndPower).
    """
    g1, g2 = g1.dropna(), g2.dropna()   # [S6]
    n1, n2 = len(g1), len(g2)
    if n1 < 2 or n2 < 2:
        raise ValueError("Each group needs at least 2 observations.")
    pooled_var = (
        (n1 - 1) * g1.var(ddof=1) + (n2 - 1) * g2.var(ddof=1)
    ) / (n1 + n2 - 2)
    if pooled_var <= 0:
        raise ValueError("Pooled variance is zero — the two groups are identical.")
    return abs(g1.mean() - g2.mean()) / math.sqrt(pooled_var)


def cohen_f_oneway(groups):
    """
    [S1] Cohen's f for one-way ANOVA.
    f = sigma_between / sigma_within   (correct metric for FTestAnovaPower).

    sigma_between is computed as sqrt(sum nI*(muI - muGrand)^2 / N)  — weighted by group size.
    sigma_within  is the square root of the pooled within-group variance.
    """
    groups = [g.dropna() for g in groups]   # [S6]
    if any(len(g) < 2 for g in groups):
        raise ValueError("Every group needs at least 2 observations.")

    ns          = np.array([len(g) for g in groups])
    n_total     = int(ns.sum())
    grand_mean  = np.concatenate([g.values for g in groups]).mean()
    group_means = np.array([g.mean() for g in groups])

    # Between-group variance (sample means weighted by group size)
    sigma_m = math.sqrt(
        float(np.sum(ns * (group_means - grand_mean) ** 2)) / n_total
    )

    # Pooled within-group variance
    df_within   = n_total - len(groups)
    within_var  = sum((n - 1) * float(g.var(ddof=1)) for g, n in zip(groups, ns)) / df_within
    sigma_within = math.sqrt(within_var)

    if sigma_within <= 0:
        raise ValueError("Pooled within-group SD is zero — all groups have identical values.")
    return sigma_m / sigma_within


def cohen_f_twoway(groups):
    """
    [S3] Two-way 2x2 ANOVA effect-size decomposition.
    groups = [A1B1, A1B2, A2B1, A2B2]   (each a pd.Series)

    Returns {'Factor A': f_A, 'Factor B': f_B, 'Interaction AxB': f_AB}

    Method — cell-mean algebra (Cohen 1988, section 10):
      grand mean and marginal means computed from sample cell means.
      sigma_A  = sqrt( mean of squared Factor-A marginal deviations from grand )
      sigma_B  = sqrt( mean of squared Factor-B marginal deviations from grand )
      sigma_AB = sqrt( mean of squared interaction cell deviations )
      sigma_error = pooled within-cell SD   (common denominator for all three f)
      f_e = sigma_e / sigma_error  for each effect e in {A, B, AB}

    Power call convention — each 1-df effect in a 2x2 design uses k_groups=2.
    FTestAnovaPower(k_groups=2) returns nobs = n per factor level = 2 x n_per_cell.
    Callers MUST divide by 2 to obtain n_per_cell.
    """
    groups = [g.dropna() for g in groups]   # [S6]
    if len(groups) != 4:
        raise ValueError("Two-way ANOVA requires exactly 4 columns (A1B1, A1B2, A2B1, A2B2).")
    if any(len(g) < 2 for g in groups):
        raise ValueError("Every cell group needs at least 2 observations.")

    a1b1, a1b2, a2b1, a2b2 = groups

    # Cell means
    m = {(1, 1): a1b1.mean(), (1, 2): a1b2.mean(),
         (2, 1): a2b1.mean(), (2, 2): a2b2.mean()}
    grand = sum(m.values()) / 4

    # Marginal means
    mA = {1: (m[1, 1] + m[1, 2]) / 2,   # Factor A level-1 (average over B)
          2: (m[2, 1] + m[2, 2]) / 2}   # Factor A level-2

    mB = {1: (m[1, 1] + m[2, 1]) / 2,   # Factor B level-1 (average over A)
          2: (m[1, 2] + m[2, 2]) / 2}   # Factor B level-2

    # sigma_A = sqrt( mean squared A-marginal deviations )
    sigma_A = math.sqrt(
        sum((mA[a] - grand) ** 2 for a in (1, 2)) / 2
    )

    # sigma_B = sqrt( mean squared B-marginal deviations )
    sigma_B = math.sqrt(
        sum((mB[b] - grand) ** 2 for b in (1, 2)) / 2
    )

    # Interaction deviations: alpha_beta_ij = mu_ij - mu.. - (mu_i. - mu..) - (mu._j - mu..)
    ab_devs = [
        m[a, b] - grand - (mA[a] - grand) - (mB[b] - grand)
        for a in (1, 2) for b in (1, 2)
    ]
    sigma_AB = math.sqrt(sum(d ** 2 for d in ab_devs) / 4)

    # Pooled within-cell SD  (sigma_error)
    ns      = [len(g) for g in groups]
    n_total = sum(ns)
    df_err  = n_total - 4
    if df_err <= 0:
        raise ValueError(
            "Not enough observations for a 2x2 ANOVA "
            "(total n must be > 4, i.e. at least 2 per cell).")
    within_var  = sum((n - 1) * float(g.var(ddof=1)) for g, n in zip(groups, ns)) / df_err
    sigma_error = math.sqrt(within_var)

    if sigma_error <= 0:
        raise ValueError("Pooled within-cell SD is zero — all cells have identical values.")

    return {
        "Factor A":        sigma_A  / sigma_error,
        "Factor B":        sigma_B  / sigma_error,
        "Interaction AxB": sigma_AB / sigma_error,
    }


# ══════════════════════════════════════════════════════════════════════════════
#  Application
# ══════════════════════════════════════════════════════════════════════════════

class PowerAnalysisApp:

    def __init__(self, root):
        self.root = root
        self.root.title(f"Power Analysis Tool  v{VERSION}")
        self.root.minsize(720, 540)

        # ── State ────────────────────────────────────────────────────────────
        self.original_data = None   # raw DataFrame (numeric columns only)
        self.file_path     = None
        self.last_results  = None   # cached for the Save Report button
        self.norm_cache    = {}     # {col: (normality_label, transform_label)}

        self.group_menus    = []    # list of ttk.Combobox widgets
        self.group_vars     = []    # list of tk.StringVar objects
        self.group_row_wids = []    # list of (label_widget, combobox_widget)

        # ── Scrollable outer canvas  [U8] ─────────────────────────────────────
        outer = ttk.Frame(root)
        outer.pack(fill="both", expand=True)

        self._canvas  = tk.Canvas(outer, borderwidth=0, highlightthickness=0)
        self._vscroll = ttk.Scrollbar(outer, orient="vertical",
                                      command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=self._vscroll.set)
        self._vscroll.pack(side="right", fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)

        self.mf  = ttk.Frame(self._canvas)   # main / inner frame
        self._wid = self._canvas.create_window((0, 0), window=self.mf, anchor="nw")

        self.mf.bind(
            "<Configure>",
            lambda e: self._canvas.configure(scrollregion=self._canvas.bbox("all"))
        )
        self._canvas.bind(
            "<Configure>",
            lambda e: self._canvas.itemconfig(self._wid, width=e.width)
        )
        root.bind_all(
            "<MouseWheel>",
            lambda e: self._canvas.yview_scroll(int(-e.delta / 120), "units")
        )

        self._build()

    # ─────────────────────────────────────────────────────────────────────────
    #  UI builder
    # ─────────────────────────────────────────────────────────────────────────

    def _build(self):
        p   = self.mf
        PX  = dict(padx=12, pady=6)

        # Title
        ttk.Label(p, text=f"Power Analysis Tool  v{VERSION}",
                  font=("TkDefaultFont", 13, "bold")).pack(fill="x", **PX)
        ttk.Label(
            p,
            text=("Load an Excel or CSV file with one group/condition per numeric column.  "
                  "Select the columns for each group, set alpha and power, then run the analysis."),
            foreground="#555", wraplength=660
        ).pack(fill="x", padx=12, pady=(0, 8))

        # ── 1 · Data File ────────────────────────────────────────────────────
        self.file_frame = ttk.LabelFrame(p, text="1 · Data File")
        self.file_frame.pack(fill="x", **PX)

        btn_row = ttk.Frame(self.file_frame)
        btn_row.pack(fill="x", padx=8, pady=6)
        ttk.Button(btn_row, text="Load Data File",
                   command=self.load_data).pack(side="left", padx=(0, 8))
        ttk.Button(btn_row, text="Reset",           # [U9]
                   command=self.reset).pack(side="left")

        self.file_lbl = ttk.Label(self.file_frame,
                                  text="No file loaded.", foreground="#888")
        self.file_lbl.pack(anchor="w", padx=8, pady=(0, 6))

        # Preview frame — packed after load via _show_preview()  [U1]
        self.preview_frame = ttk.LabelFrame(p, text="Data Preview  (first 5 rows)")

        # ── 2 · Configuration ────────────────────────────────────────────────
        cfg = ttk.LabelFrame(p, text="2 · Configuration")
        cfg.pack(fill="x", **PX)
        cfg.columnconfigure(1, weight=1)

        # Test-type selector
        ttk.Label(cfg, text="Statistical Test:").grid(
            row=0, column=0, sticky="w", padx=8, pady=5)
        self.test_var = tk.StringVar(value="t-test")
        self.test_var.trace_add("write", self._on_test_change)
        ttk.Combobox(
            cfg, textvariable=self.test_var,
            values=["t-test", "One-way ANOVA", "Two-way ANOVA"],
            state="readonly", width=22
        ).grid(row=0, column=1, sticky="w", padx=8, pady=5)

        # One-way ANOVA helper (number-of-groups row) — hidden until needed  [C2]
        self.ng_frame = ttk.Frame(cfg)
        ttk.Label(self.ng_frame, text="Number of Groups:").grid(
            row=0, column=0, sticky="w")
        self.ng_var = tk.StringVar(value="3")
        ttk.Entry(self.ng_frame, textvariable=self.ng_var, width=5).grid(
            row=0, column=1, padx=6)
        ttk.Button(self.ng_frame, text="Update Groups",
                   command=self._update_group_count).grid(row=0, column=2, padx=4)

        # Group-selector container  [C1]
        self.grp_frame = ttk.Frame(cfg)
        self.grp_frame.grid(row=2, column=0, columnspan=2, sticky="ew", padx=8)
        self.grp_frame.columnconfigure(1, weight=1)

        # Alpha / Power row
        ap = ttk.Frame(cfg)
        ap.grid(row=3, column=0, columnspan=2, sticky="ew", padx=8, pady=(8, 2))

        ttk.Label(ap, text="Alpha (alpha):").grid(row=0, column=0, sticky="w",
                                                   padx=(0, 4), pady=5)
        self.alpha_var = tk.StringVar(value="0.05")
        self.alpha_ent = ttk.Entry(ap, textvariable=self.alpha_var, width=9)
        self.alpha_ent.grid(row=0, column=1, sticky="w", pady=5)
        self.alpha_var.trace_add(
            "write", lambda *_: self._val_param(self.alpha_ent, self.alpha_var))
        ToolTip(self.alpha_ent,                                              # [U10]
                "Significance level alpha — probability of a false positive (Type I error).\n"
                "Typical values: 0.05  (standard)  or  0.01  (stricter).")

        ttk.Label(ap, text="Power (1-beta):").grid(row=0, column=2, sticky="w",
                                                    padx=(18, 4), pady=5)
        self.power_var = tk.StringVar(value="0.80")
        self.power_ent = ttk.Entry(ap, textvariable=self.power_var, width=9)
        self.power_ent.grid(row=0, column=3, sticky="w", pady=5)
        self.power_var.trace_add(
            "write", lambda *_: self._val_param(self.power_ent, self.power_var))
        ToolTip(self.power_ent,                                              # [U10]
                "Statistical power 1-beta — probability of detecting a true effect.\n"
                "Common targets: 0.80  (standard)  or  0.90  (stricter).")

        # Manual effect-size override  [U6]
        self.manual_var = tk.BooleanVar(value=False)
        man_chk = ttk.Checkbutton(
            cfg, text="Override effect size manually (prospective analysis without pilot data)",
            variable=self.manual_var, command=self._toggle_manual)
        man_chk.grid(row=4, column=0, columnspan=2, sticky="w", padx=8, pady=(8, 0))
        ToolTip(man_chk,
                "Enable this when you have no pilot data and want to plan a new study.\n"
                "Enter Cohen's d (t-test) or Cohen's f (ANOVA) from the literature.\n"
                "Conventional benchmarks:\n"
                "  Cohen's d — small: 0.2  medium: 0.5  large: 0.8\n"
                "  Cohen's f — small: 0.1  medium: 0.25  large: 0.4")

        self.man_frame = ttk.Frame(cfg)
        ttk.Label(self.man_frame, text="Effect size (d or f):").grid(
            row=0, column=0, sticky="w")
        self.man_ent = ttk.Entry(self.man_frame, width=10)
        self.man_ent.grid(row=0, column=1, padx=6)
        ToolTip(self.man_ent,
                "Cohen's d for t-test (standardised mean difference).\n"
                "Cohen's f = sigma_between / sigma_within for ANOVA.\n"
                "Note: d and f are related by  f = d/2  only for balanced two-group ANOVA.")

        # ── Action row  [U4] ─────────────────────────────────────────────────
        act = ttk.Frame(p)
        act.pack(fill="x", padx=12, pady=8)
        ttk.Button(act, text="Perform Power Analysis",
                   command=self.analyse).pack(side="left", padx=(0, 10))
        self.save_btn = ttk.Button(act, text="Save Excel Report",
                                   command=self.save_report, state="disabled")
        self.save_btn.pack(side="left")

        # ── 3 · Results  [U3] ────────────────────────────────────────────────
        self.res_frame = ttk.LabelFrame(p, text="3 · Results")
        # Not packed until analysis completes

        self.res_txt = tk.Text(
            self.res_frame, height=14, state="disabled",
            font=("Courier", 10), relief="flat", wrap="none",
            bg=p.winfo_toplevel().cget("bg"))
        xsb_res = ttk.Scrollbar(self.res_frame, orient="horizontal",
                                 command=self.res_txt.xview)
        self.res_txt.configure(xscrollcommand=xsb_res.set)
        self.res_txt.pack(fill="both", padx=8, pady=(6, 0))
        xsb_res.pack(fill="x", padx=8, pady=(0, 6))

        # ── 4 · Power Curve  [U7] ────────────────────────────────────────────
        if MATPLOTLIB_AVAILABLE:
            self.plt_frame = ttk.LabelFrame(p, text="4 · Power Curve")
            # Not packed until analysis completes
            self.fig = Figure(figsize=(6.8, 3.4), dpi=90, tight_layout=True)
            self.ax  = self.fig.add_subplot(111)
            self.plt_canvas = FigureCanvasTkAgg(self.fig, master=self.plt_frame)
            self.plt_canvas.get_tk_widget().pack(
                fill="both", expand=True, padx=8, pady=6)

        # Initialise with t-test group dropdowns
        self._create_groups(2, ["Group 1", "Group 2"])

    # ─────────────────────────────────────────────────────────────────────────
    #  Parameter validation helpers
    # ─────────────────────────────────────────────────────────────────────────

    @staticmethod
    def _val_param(widget, var):
        """[U5] Highlight entry red when the value is outside (0, 1)."""
        try:
            v  = float(var.get())
            ok = 0 < v < 1
        except ValueError:
            ok = False
        widget.configure(foreground="red" if not ok else "black")

    def _toggle_manual(self):
        """[U6] Show / hide the manual effect-size entry row."""
        if self.manual_var.get():
            self.man_frame.grid(row=5, column=0, columnspan=2,
                                sticky="w", padx=8, pady=4)
        else:
            self.man_frame.grid_forget()

    # ─────────────────────────────────────────────────────────────────────────
    #  File loading
    # ─────────────────────────────────────────────────────────────────────────

    def load_data(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")])
        if not path:
            return
        try:
            df = (pd.read_excel(path) if path.endswith(".xlsx")
                  else pd.read_csv(path))

            # [C6] Silently drop non-numeric columns with a warning
            num_cols = df.select_dtypes(include=[np.number]).columns.tolist()
            dropped  = [c for c in df.columns if c not in num_cols]
            if dropped:
                messagebox.showwarning(
                    "Non-numeric columns ignored",
                    f"The following columns were skipped (non-numeric data):\n"
                    f"{', '.join(dropped)}")
            if not num_cols:
                messagebox.showerror(
                    "No numeric data",
                    "The file contains no numeric columns and cannot be used.")
                return

            self.original_data = df[num_cols]
            self.file_path     = path
            self.last_results  = None
            self.norm_cache    = {}
            self.save_btn.configure(state="disabled")

            short = os.path.basename(path)
            self.file_lbl.configure(
                text=(f"Loaded: {short}   "
                      f"({len(self.original_data)} rows, "
                      f"{len(num_cols)} numeric columns)"),
                foreground="black")
            self._update_dropdowns()
            self._show_preview()

        except Exception as e:
            messagebox.showerror("Load error", f"Failed to load file:\n{e}")

    def _show_preview(self):
        """[U1] Render a scrollable table of the first 5 rows."""
        for w in self.preview_frame.winfo_children():
            w.destroy()

        df   = self.original_data.head(5)
        cols = list(df.columns)

        tree = ttk.Treeview(self.preview_frame,
                             columns=cols, show="headings", height=5)
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=max(80, len(c) * 9),
                         anchor="center", stretch=False)
        for _, row in df.iterrows():
            tree.insert("", "end", values=[
                f"{v:.4g}" if isinstance(v, (float, np.floating)) else v
                for v in row])

        xsb = ttk.Scrollbar(self.preview_frame, orient="horizontal",
                              command=tree.xview)
        tree.configure(xscrollcommand=xsb.set)
        tree.pack(fill="x", padx=8, pady=(6, 0))
        xsb.pack(fill="x", padx=8, pady=(0, 6))

        if not self.preview_frame.winfo_manager():
            self.preview_frame.pack(fill="x", padx=12, pady=6,
                                    after=self.file_frame)

    # ─────────────────────────────────────────────────────────────────────────
    #  Group dropdown management
    # ─────────────────────────────────────────────────────────────────────────

    def _update_dropdowns(self):
        """Populate all Combobox menus with the current column list."""
        if self.original_data is None:
            return
        cols = list(self.original_data.columns)
        for m in self.group_menus:
            m.configure(values=cols)

    def _clear_groups(self):
        for lbl, cb in self.group_row_wids:
            lbl.destroy()
            cb.destroy()
        self.group_menus.clear()
        self.group_vars.clear()
        self.group_row_wids.clear()

    def _create_groups(self, n, labels):
        self._clear_groups()
        for i, lbl_text in enumerate(labels):
            lbl = ttk.Label(self.grp_frame, text=f"{lbl_text}:")
            lbl.grid(row=i, column=0, sticky="w", padx=(0, 8), pady=3)
            var = tk.StringVar()
            cb  = ttk.Combobox(self.grp_frame, textvariable=var,
                                state="readonly")
            cb.grid(row=i, column=1, sticky="ew", pady=3)
            self.group_menus.append(cb)
            self.group_vars.append(var)
            self.group_row_wids.append((lbl, cb))
        self._update_dropdowns()

    def _on_test_change(self, *_):
        """[C2] Rebuild group menus and hide/show One-way ANOVA helpers."""
        self._clear_groups()
        self.ng_frame.grid_forget()          # always hide; re-show if needed
        t = self.test_var.get()

        if t == "t-test":
            self._create_groups(2, ["Group 1", "Group 2"])

        elif t == "One-way ANOVA":
            self.ng_frame.grid(row=1, column=0, columnspan=2,
                               sticky="w", padx=8, pady=4)
            n = int(self.ng_var.get()) if self.ng_var.get().isdigit() else 3
            self._create_groups(n, [f"Group {i+1}" for i in range(n)])

        elif t == "Two-way ANOVA":
            self._create_groups(4, [
                "A1B1  (Factor A: lvl 1, Factor B: lvl 1)",
                "A1B2  (Factor A: lvl 1, Factor B: lvl 2)",
                "A2B1  (Factor A: lvl 2, Factor B: lvl 1)",
                "A2B2  (Factor A: lvl 2, Factor B: lvl 2)",
            ])

    def _update_group_count(self):
        """[C2] Re-create group menus with the user-specified group count."""
        try:
            n = int(self.ng_var.get())
            if n < 2:
                raise ValueError("At least 2 groups required.")
            self._create_groups(n, [f"Group {i+1}" for i in range(n)])
        except ValueError as e:
            messagebox.showerror("Input error", str(e))

    # ─────────────────────────────────────────────────────────────────────────
    #  Analysis entry point
    # ─────────────────────────────────────────────────────────────────────────

    def analyse(self):
        """Validate inputs then dispatch to the appropriate analysis routine."""
        # Validate alpha / power  [U5]
        try:
            alpha = float(self.alpha_var.get())
            power = float(self.power_var.get())
            if not (0 < alpha < 1) or not (0 < power < 1):
                raise ValueError
        except ValueError:
            messagebox.showerror(
                "Input error",
                "Alpha and Power must both be numbers strictly between 0 and 1.")
            return

        test   = self.test_var.get()
        manual = self.manual_var.get()

        if not manual:
            # [C3] Require loaded data and valid column selections
            if self.original_data is None:
                messagebox.showerror("No data",
                    "Please load a data file before running the analysis.")
                return
            cols = [v.get() for v in self.group_vars]
            if any(c == "" or c not in self.original_data.columns for c in cols):
                messagebox.showerror(
                    "Input error",
                    "Please select a valid column for every group.")
                return
        else:
            cols = []

        try:
            if manual:
                self._analyse_manual(alpha, power, test)
            elif test == "t-test":
                self._analyse_ttest(cols, alpha, power)
            elif test == "One-way ANOVA":
                self._analyse_oneway(cols, alpha, power)
            elif test == "Two-way ANOVA":
                self._analyse_twoway(cols, alpha, power)
        except Exception as e:
            messagebox.showerror("Analysis error", str(e))

    # ── Analysis routines ─────────────────────────────────────────────────────

    def _analyse_manual(self, alpha, power, test):
        """[U6] Prospective power analysis with a user-supplied effect size."""
        try:
            es = float(self.man_ent.get())
            if es <= 0:
                raise ValueError("Effect size must be a positive number.")
        except ValueError as e:
            messagebox.showerror("Input error", str(e))
            return

        if test == "t-test":
            n        = TTestIndPower().solve_power(
                           effect_size=es, alpha=alpha, power=power)
            es_label = "Cohen's d"
            k        = 2
        else:
            k = (int(self.ng_var.get()) if test == "One-way ANOVA"
                                        and self.ng_var.get().isdigit() else 4)
            n        = FTestAnovaPower().solve_power(
                           effect_size=es, alpha=alpha, power=power,
                           k_groups=k)
            es_label = "Cohen's f"

        n_ceil = math.ceil(n)
        lines  = [
            f"Test Type             : {test}",
            f"Effect Size           : {es:.4f}  [{es_label}, manual entry]",
            f"Alpha (alpha)         : {alpha}",
            f"Power (1-beta)        : {power}",
            "-" * 58,
            f"Required n per group  : {n_ceil}",
            "",
            "NOTE: No pilot data loaded — normality check not performed.",
            "      Biological-replicate estimate requires pilot data.",
        ]

        self.last_results = dict(
            cols=[], es=es, es_label=es_label,
            alpha=alpha, power=power, test=test,
            n=n, n_per_cell=n, norm_info={}, details={}, k=k)
        self._show_results("\n".join(lines), es, alpha, power, test, n, k,
                           n_per_cell=n)

    def _analyse_ttest(self, cols, alpha, power):
        """[S1,S2,S4] Independent-samples t-test power analysis."""
        g1, n1_lbl, t1 = prepare_column(self.original_data, cols[0])   # [S4]
        g2, n2_lbl, t2 = prepare_column(self.original_data, cols[1])
        norm_info       = {cols[0]: (n1_lbl, t1), cols[1]: (n2_lbl, t2)}
        self.norm_cache = norm_info                                       # [C4]

        es     = cohen_d(g1, g2)                                          # [S1,S2]
        n      = TTestIndPower().solve_power(
                     effect_size=es, alpha=alpha, power=power)
        n_ceil = math.ceil(n)
        n_obs  = len(self.original_data[cols[0]].dropna())                # [S6]
        bio    = n / n_obs

        norm_block = "\n".join(
            f"  {c:<32s}  {s}  (transform: {t})"
            for c, (s, t) in norm_info.items())
        lines = [
            f"Test Type              : t-test (independent samples)",
            f"Columns                : {cols[0]}  vs  {cols[1]}",
            f"Effect Size (Cohen's d): {es:.4f}",
            f"Alpha (alpha)          : {alpha}",
            f"Power (1-beta)         : {power}",
            "-" * 58,
            f"Required n per group   : {n_ceil}",
            f"Biological replicates  : {bio:.2f}",
            f"  [S7] Assumption — each row in pilot = 1 technical replicate",
            f"       from a single biological replicate.",
            "",
            "Normality (Shapiro-Wilk):",
            norm_block,
        ]

        self.last_results = dict(
            cols=cols, es=es, es_label="Cohen's d",
            alpha=alpha, power=power, test="t-test",
            n=n, n_per_cell=n, norm_info=norm_info, details={}, k=2)
        self._show_results("\n".join(lines), es, alpha, power, "t-test", n, 2,
                           n_per_cell=n)

    def _analyse_oneway(self, cols, alpha, power):
        """[S1,S4] One-way ANOVA power analysis using Cohen's f."""
        groups_data = []
        norm_info   = {}
        for col in cols:
            s, ns, tr = prepare_column(self.original_data, col)           # [S4]
            groups_data.append(s)
            norm_info[col] = (ns, tr)
        self.norm_cache = norm_info                                        # [C4]

        k      = len(cols)
        es     = cohen_f_oneway(groups_data)                              # [S1]
        n      = FTestAnovaPower().solve_power(
                     effect_size=es, alpha=alpha, power=power, k_groups=k)
        n_ceil = math.ceil(n)
        n_obs  = len(self.original_data[cols[0]].dropna())                # [S6]
        bio    = n / n_obs

        norm_block = "\n".join(
            f"  {c:<32s}  {s}  (transform: {t})"
            for c, (s, t) in norm_info.items())
        lines = [
            f"Test Type              : One-way ANOVA  ({k} groups)",
            f"Columns                : {', '.join(cols)}",
            f"Effect Size (Cohen's f): {es:.4f}  [= sigma_between / sigma_within]",
            f"Alpha (alpha)          : {alpha}",
            f"Power (1-beta)         : {power}",
            "-" * 58,
            f"Required n per group   : {n_ceil}",
            f"Biological replicates  : {bio:.2f}",
            f"  [S7] Assumption — each row in pilot = 1 technical replicate",
            f"       from a single biological replicate.",
            "",
            "Normality (Shapiro-Wilk):",
            norm_block,
        ]

        self.last_results = dict(
            cols=cols, es=es, es_label="Cohen's f",
            alpha=alpha, power=power, test="One-way ANOVA",
            n=n, n_per_cell=n, norm_info=norm_info, details={}, k=k)
        self._show_results("\n".join(lines), es, alpha, power,
                           "One-way ANOVA", n, k, n_per_cell=n)

    def _analyse_twoway(self, cols, alpha, power):
        """
        [S3] Two-way 2x2 ANOVA power analysis.

        Each of the three effects (Factor A, Factor B, Interaction AxB) has its
        own Cohen's f; FTestAnovaPower(k_groups=2) is applied per effect.

        IMPORTANT — unit conversion:
        In a 2x2 balanced design FTestAnovaPower(k_groups=2) models each 1-df
        effect as a 2-group one-way ANOVA.  The returned nobs equals the number
        of observations per factor level, which is 2 x n_per_cell.  We divide
        by 2 to convert back to n per cell (the quantity of practical interest).

        The most conservative n_per_cell (across all three effects) is reported.
        Note: FTestAnovaPower is an approximation here; the shared error df
        structure of two-way ANOVA is not perfectly captured, but this is
        consistent with the convention used by G*Power.
        """
        groups_data = []
        norm_info   = {}
        for col in cols:
            s, ns, tr = prepare_column(self.original_data, col)           # [S4]
            groups_data.append(s)
            norm_info[col] = (ns, tr)
        self.norm_cache = norm_info                                        # [C4]

        f_dict   = cohen_f_twoway(groups_data)                            # [S3]
        analysis = FTestAnovaPower()

        # [S3] Solve per effect; divide nobs by 2 to get n_per_cell
        n_per_cell_dict = {}
        for eff, f_val in f_dict.items():
            if f_val > 0:
                nobs           = analysis.solve_power(
                                     effect_size=f_val, alpha=alpha,
                                     power=power, k_groups=2)
                n_per_cell_dict[eff] = nobs / 2   # convert n-per-level to n-per-cell
            else:
                n_per_cell_dict[eff] = float("nan")

        valid_ns   = [v for v in n_per_cell_dict.values() if not math.isnan(v)]
        n_max      = max(valid_ns) if valid_ns else float("nan")
        n_ceil     = math.ceil(n_max) if not math.isnan(n_max) else "N/A"
        n_obs      = len(self.original_data[cols[0]].dropna())            # [S6]
        bio_str    = (f"{n_max / n_obs:.2f}" if not math.isnan(n_max)
                      else "N/A")

        eff_block = "\n".join(
            f"  {eff:<22s}  f = {f_val:.4f}   "
            f"required n/cell = "
            f"{math.ceil(n_per_cell_dict[eff]) if not math.isnan(n_per_cell_dict[eff]) else 'N/A'}"
            for eff, f_val in f_dict.items())
        norm_block = "\n".join(
            f"  {c:<32s}  {s}  (transform: {t})"
            for c, (s, t) in norm_info.items())
        lines = [
            f"Test Type              : Two-way ANOVA  (2x2 design)",
            f"Columns                : {', '.join(cols)}",
            f"Alpha (alpha)          : {alpha}",
            f"Power (1-beta)         : {power}",
            "-" * 58,
            "Per-effect Cohen's f  and required n per cell:",
            eff_block,
            "",
            f"Required n per cell    : {n_ceil}  (most conservative effect)",
            f"Biological replicates  : {bio_str}",
            f"  [S7] Assumption — each row in pilot = 1 technical replicate",
            f"       from a single biological replicate.",
            "",
            "NOTE: Power calculation uses FTestAnovaPower per 1-df effect",
            "      (approximation consistent with G*Power conventions).",
            "",
            "Normality (Shapiro-Wilk):",
            norm_block,
        ]

        # For the power curve, use the dominant (largest-f) effect
        dom_eff = max(f_dict, key=f_dict.get)
        dom_f   = f_dict[dom_eff]

        self.last_results = dict(
            cols=cols, es=dom_f,
            es_label=f"Cohen's f (dominant: {dom_eff})",
            alpha=alpha, power=power, test="Two-way ANOVA",
            n=n_max, n_per_cell=n_max,
            norm_info=norm_info,
            details={"f_dict": f_dict, "n_per_cell": n_per_cell_dict},
            k=2)
        self._show_results("\n".join(lines), dom_f, alpha, power,
                           "Two-way ANOVA", n_max, 2, n_per_cell=n_max)

    # ─────────────────────────────────────────────────────────────────────────
    #  Display results
    # ─────────────────────────────────────────────────────────────────────────

    def _show_results(self, text, es, alpha, power, test, n, k, n_per_cell=None):
        """[U3] Write text to the persistent results panel and draw the curve."""
        # Show results panel
        self.res_frame.pack(fill="x", padx=12, pady=6)
        self.res_txt.configure(state="normal")
        self.res_txt.delete("1.0", "end")
        self.res_txt.insert("end", text)
        self.res_txt.configure(state="disabled")
        self.save_btn.configure(state="normal")

        # Draw power curve  [U7]
        if MATPLOTLIB_AVAILABLE:
            self.plt_frame.pack(fill="x", padx=12, pady=6)
            req = n_per_cell if n_per_cell is not None else n
            self._draw_curve(es, alpha, power, test, req, k)

    def _draw_curve(self, es, alpha, power_target, test, req_n, k):
        """[U7] Plot power vs. n-per-group for the computed effect size."""
        self.ax.clear()

        if math.isnan(req_n) or es <= 0:
            self.ax.text(0.5, 0.5,
                         "Cannot draw curve — effect size <= 0 or required n undefined.",
                         transform=self.ax.transAxes, ha="center", va="center")
            self.plt_canvas.draw()
            return

        n_max  = max(req_n * 2.5, 60)
        n_vals = np.linspace(3, n_max, 300)
        pwrs   = []

        for nv in n_vals:
            try:
                if test == "t-test":
                    p = TTestIndPower().solve_power(
                            effect_size=es, alpha=alpha, nobs1=nv)
                else:
                    # For two-way ANOVA, nv is n_per_cell so nobs = 2 x nv
                    nobs_arg = (nv * 2) if test == "Two-way ANOVA" else nv
                    p = FTestAnovaPower().solve_power(
                            effect_size=es, alpha=alpha,
                            nobs=nobs_arg, k_groups=k)
                pwrs.append(p)
            except Exception:
                pwrs.append(float("nan"))

        pwrs = np.array(pwrs, dtype=float)
        ok   = ~np.isnan(pwrs)

        self.ax.plot(n_vals[ok], pwrs[ok],
                     color="#2563eb", linewidth=2.2, label="Power")
        self.ax.axhline(power_target, color="#dc2626", ls="--", lw=1.4,
                        label=f"Target = {power_target}")
        self.ax.axvline(math.ceil(req_n), color="#16a34a", ls="--", lw=1.4,
                        label=f"Required n = {math.ceil(req_n)}")

        x_label = ("n per cell" if test == "Two-way ANOVA"
                   else "n per group")
        self.ax.set_xlabel(x_label, fontsize=10)
        self.ax.set_ylabel("Power  (1-beta)", fontsize=10)
        self.ax.set_title(
            f"{test}  |  effect = {es:.3f}  |  alpha = {alpha}", fontsize=10)
        self.ax.set_ylim(0, 1.05)
        self.ax.legend(fontsize=9)
        self.ax.grid(True, ls=":", alpha=0.4)
        self.fig.tight_layout()
        self.plt_canvas.draw()

    # ─────────────────────────────────────────────────────────────────────────
    #  Save Excel report  [U4, C4]
    # ─────────────────────────────────────────────────────────────────────────

    def save_report(self):
        if not self.last_results:
            messagebox.showerror("No results", "Run the analysis first.")
            return
        r = self.last_results

        n_cell  = r.get("n_per_cell", r["n"])
        n_obs   = (len(self.original_data[r["cols"][0]].dropna())
                   if r["cols"] and self.original_data is not None else None)
        bio     = (round(n_cell / n_obs, 4)
                   if (n_obs and not math.isnan(n_cell)) else "N/A")
        n_ceil  = math.ceil(n_cell) if not math.isnan(n_cell) else "N/A"

        # ── Summary sheet ────────────────────────────────────────────────────
        summary = {
            "Test Type":                            [r["test"]],
            "Columns": [", ".join(r["cols"]) if r["cols"] else "manual entry"],
            f'Effect Size ({r["es_label"]})':       [round(r["es"], 6)],
            "Alpha":                                [r["alpha"]],
            "Power":                                [r["power"]],
            "Required n per group/cell (ceiling)":  [n_ceil],
            "Biological Replicates (estimated)":    [bio],
            "Assumption (bio. replicates)":
                ["Each row in pilot data = 1 technical replicate "
                 "from 1 biological replicate"],
        }
        sum_df = pd.DataFrame(summary)

        # ── Normality sheet  [C4] ─────────────────────────────────────────────
        norm_rows = [
            {"Column": c, "Normality (Shapiro-Wilk)": s,
             "Transformation Applied": t}
            for c, (s, t) in r["norm_info"].items()
        ]
        norm_df = (pd.DataFrame(norm_rows) if norm_rows else
                   pd.DataFrame(columns=["Column", "Normality (Shapiro-Wilk)",
                                         "Transformation Applied"]))

        # ── Two-way ANOVA detail sheet ────────────────────────────────────────
        detail_df = None
        if r["test"] == "Two-way ANOVA" and r.get("details"):
            rows = []
            for eff, fv in r["details"]["f_dict"].items():
                nv = r["details"]["n_per_cell"].get(eff, float("nan"))
                rows.append({
                    "Effect":                       eff,
                    "Cohen's f":                    round(fv, 6),
                    "Required n per cell (ceiling)":
                        math.ceil(nv) if not math.isnan(nv) else "N/A",
                })
            detail_df = pd.DataFrame(rows)

        # ── Save ──────────────────────────────────────────────────────────────
        default = (os.path.splitext(os.path.basename(self.file_path))[0]
                   + "_PowerAnalysis.xlsx"
                   if self.file_path else "PowerAnalysis.xlsx")
        save_path = filedialog.asksaveasfilename(
            initialfile=default, defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")])
        if not save_path:
            return

        try:
            with pd.ExcelWriter(save_path, engine="openpyxl") as w:
                sum_df.to_excel(w,  sheet_name="Summary",   index=False)
                norm_df.to_excel(w, sheet_name="Normality", index=False)
                if detail_df is not None:
                    detail_df.to_excel(w,
                        sheet_name="Two-way ANOVA Details", index=False)
            messagebox.showinfo("Saved", f"Report saved to:\n{save_path}")
        except Exception as e:
            messagebox.showerror("Save error", f"Failed to save report:\n{e}")

    # ─────────────────────────────────────────────────────────────────────────
    #  Reset  [U9]
    # ─────────────────────────────────────────────────────────────────────────

    def reset(self):
        """Return the tool to its initial state without restarting."""
        self.original_data = None
        self.file_path     = None
        self.last_results  = None
        self.norm_cache    = {}

        self.file_lbl.configure(text="No file loaded.", foreground="#888")

        if self.preview_frame.winfo_manager():
            self.preview_frame.pack_forget()
        if self.res_frame.winfo_manager():
            self.res_frame.pack_forget()
        if MATPLOTLIB_AVAILABLE and self.plt_frame.winfo_manager():
            self.plt_frame.pack_forget()

        self.save_btn.configure(state="disabled")
        self.alpha_var.set("0.05")
        self.power_var.set("0.80")
        self.manual_var.set(False)
        self._toggle_manual()
        self.test_var.set("t-test")   # triggers _on_test_change -> rebuilds menus


# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    root = tk.Tk()
    app  = PowerAnalysisApp(root)
    root.mainloop()
