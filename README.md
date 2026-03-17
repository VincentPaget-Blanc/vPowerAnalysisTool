# vPowerAnalysisTool

The vPowerAnalysis Tool is a graphical user interface (GUI) application designed to help researchers perform power analysis for t-tests, one-way ANOVA, and two-way ANOVA. This tool calculates the required sample size and the number of biological replicates needed to achieve a specified statistical power. It also checks for data normality and applies transformations if necessary to ensure the validity of parametric tests.

> **Note on pilot data:** The tool expects pilot data in which each row represents one technical replicate from a single biological replicate. The biological replicate estimate is derived from this assumption and should be interpreted accordingly.

## Table of Contents

1. [Installation](#installation)
2. [Usage](#usage)
3. [Data Organization](#data-organization)
4. [Calculations](#calculations)
5. [Examples](#examples)
6. [Troubleshooting](#troubleshooting)
7. [Changelog](#changelog)
8. [License](#license)

## Installation

To run the Power Analysis Tool, you need Python 3.8 or later installed on your computer. Install the required libraries with:

```bash
pip install pandas statsmodels scipy numpy openpyxl matplotlib
```

> `tkinter` is included with most standard Python distributions and does not need to be installed separately. `matplotlib` is optional — the embedded power curve will not appear if it is not installed, but all other features remain fully functional.

## Usage

### Step 1: Load Data

1. Click **Load Data File** to open a file dialog.
2. Select an Excel (`.xlsx`) or CSV (`.csv`) file containing your data.
3. A scrollable preview of the first five rows is displayed immediately after loading so you can verify the correct file was selected.
4. Non-numeric columns are automatically excluded with a warning. All analysis operates on numeric columns only.

### Step 2: Select Columns

- **For t-test:** Select the columns representing the two groups from the **Group 1** and **Group 2** dropdown menus.

- **For One-way ANOVA:**
  1. Enter the number of groups in the **Number of Groups** field.
  2. Click **Update Groups** to generate the appropriate dropdown menus.
  3. Select the column for each group.

- **For Two-way ANOVA:** Select the four columns representing each cell of the 2×2 design from the labelled dropdown menus (A1B1, A1B2, A2B1, A2B2).

> **Prospective analysis (no pilot data):** Check **Override effect size manually** to enter a Cohen's d or Cohen's f value directly from the literature. Hover over the checkbox or the entry field for conventional benchmarks (small / medium / large).

### Step 3: Set Parameters

| Parameter | Description | Default |
|---|---|---|
| **Alpha (α)** | Significance level (Type I error rate) | 0.05 |
| **Power (1−β)** | Probability of detecting a true effect | 0.80 |

Both fields are highlighted red if the entered value is outside the valid range (0, 1).

### Step 4: Choose Statistical Test

Select the test type from the **Statistical Test** dropdown:

- **t-test** — independent-samples t-test between two groups
- **One-way ANOVA** — comparison across two or more groups
- **Two-way ANOVA** — 2×2 factorial design with main effects and interaction

### Step 5: Perform Power Analysis

Click **Perform Power Analysis**. Results appear immediately in the persistent **Results** panel below, including:

- Effect size (Cohen's d for t-test; Cohen's f per effect for ANOVA)
- Required sample size (ceiling)
- Estimated biological replicates
- Per-column normality status and any transformation applied

An embedded **power curve** (power vs. n-per-group or n-per-cell) is drawn below the results panel when matplotlib is available.

### Step 6: Save the Report

Click **Save Excel Report** (enabled only after a successful analysis) to save a multi-sheet Excel file containing:

- **Summary** — key parameters and results
- **Normality** — per-column Shapiro-Wilk result and transformation applied
- **Two-way ANOVA Details** *(two-way ANOVA only)* — Cohen's f and required n per cell for each effect (Factor A, Factor B, Interaction A×B)

### Step 7: Reset

Click **Reset** at any time to clear all loaded data and results and return the tool to its initial state.

## Data Organization

### For t-test

Organise data into two columns, one per group.

| Group1 | Group2 |
|--------|--------|
| 10 | 7 |
| 12 | 9 |
| 14 | 11 |

### For One-way ANOVA

Each group occupies a separate column.

| Group1 | Group2 | Group3 |
|--------|--------|--------|
| 10 | 7 | 8 |
| 12 | 9 | 10 |
| 14 | 11 | 12 |

### For Two-way ANOVA

Each cell of the 2×2 design occupies a separate column. Column names should clearly identify the factor-level combination.

| A1B1 | A1B2 | A2B1 | A2B2 |
|------|------|------|------|
| 10 | 7 | 8 | 9 |
| 12 | 9 | 10 | 11 |
| 14 | 11 | 12 | 13 |

Where the column labels denote:

| Column | Factor A level | Factor B level |
|--------|---------------|---------------|
| A1B1 | 1 | 1 |
| A1B2 | 1 | 2 |
| A2B1 | 2 | 1 |
| A2B2 | 2 | 2 |

## Calculations

### Normality Check and Transformation

Each selected column is tested for normality using the **Shapiro-Wilk test** (α = 0.05) at analysis time. Normality checks are performed on the selected columns only and do not alter other columns in the dataset.

- If n < 3, the column is assumed normal (the test cannot run).
- If n > 5,000, a random subsample of 5,000 observations is used.
- Missing values (NaN) are removed before all calculations.

If a column is not normally distributed, the tool attempts:

1. **Log transformation** — applied to the original values (shifted to be strictly positive only if values ≤ 0 are present).
2. **Box-Cox transformation** — applied if log transformation does not achieve normality.

If neither transformation achieves normality, the raw data is used and a warning is recorded in the results and report. In this case, consider using a non-parametric test.

### Cohen's d (t-test)

Cohen's d measures the standardised difference between two group means:

```
d = |M1 − M2| / s_pooled

s_pooled = sqrt( ((n1−1)·s1² + (n2−1)·s2²) / (n1 + n2 − 2) )
```

The absolute value is taken to ensure the effect size passed to the power function is always non-negative.

### Cohen's f (ANOVA)

Cohen's f is the correct effect-size metric for ANOVA power analysis and is defined as:

```
f = σ_between / σ_within
```

where σ_between is the weighted standard deviation of group means around the grand mean, and σ_within is the square root of the pooled within-group variance.

**One-way ANOVA:**

```
σ_between = sqrt( Σ nᵢ·(μᵢ − μ̄)² / N )
σ_within  = sqrt( Σ (nᵢ−1)·sᵢ² / (N − k) )
```

**Two-way ANOVA (2×2 design):** Cell-mean algebra (Cohen 1988) is used to decompose the pilot data into three independent Cohen's f values:

- **Factor A** — based on the deviation of Factor A marginal means from the grand mean
- **Factor B** — based on the deviation of Factor B marginal means from the grand mean
- **Interaction A×B** — based on the residual cell deviations after removing main effects

A separate required sample size is computed for each effect. The most conservative (largest) n per cell is reported as the design requirement. Each effect is modelled as a 2-group one-way ANOVA (consistent with G*Power conventions); note that this is an approximation because it does not account for the shared error degrees of freedom structure of a two-way design.

### Power Analysis

Required sample sizes are calculated using the `statsmodels` library:

- **t-test:** `TTestIndPower.solve_power(effect_size=d, alpha=α, power=1−β)`
- **One-way / Two-way ANOVA:** `FTestAnovaPower.solve_power(effect_size=f, alpha=α, power=1−β, k_groups=k)`

For the two-way ANOVA, `FTestAnovaPower` is called with `k_groups=2` per effect (each is a 1-df contrast). The returned `nobs` represents observations per factor level, which equals 2 × n per cell in a balanced 2×2 design; the tool divides by 2 to report n per cell.

### Biological Replicates

```
Biological Replicates = Required Sample Size / n_pilot
```

where n_pilot is the number of non-missing observations in the first selected column.

**Assumption:** Each row in the pilot data represents one technical replicate from a single biological replicate. If your pilot data contains multiple technical replicates per biological replicate, the biological replicate estimate will need to be adjusted manually.

## Examples

### Example 1: t-test

1. Load a data file with two numeric columns representing two groups.
2. Select the columns for **Group 1** and **Group 2**.
3. Set Alpha to `0.05` and Power to `0.80`.
4. Choose **t-test** from the Statistical Test dropdown.
5. Click **Perform Power Analysis**.
6. Review the results panel and power curve, then click **Save Excel Report**.

### Example 2: One-way ANOVA

1. Load a data file with one column per group.
2. Enter the number of groups and click **Update Groups**.
3. Select the column for each group from the generated dropdown menus.
4. Set Alpha to `0.05` and Power to `0.80`.
5. Choose **One-way ANOVA** from the Statistical Test dropdown.
6. Click **Perform Power Analysis**, then **Save Excel Report**.

### Example 3: Two-way ANOVA

1. Load a data file with four columns representing the 2×2 cell combinations (e.g. A1B1, A1B2, A2B1, A2B2).
2. Select the correct column for each labelled dropdown (A1B1 through A2B2).
3. Set Alpha to `0.05` and Power to `0.80`.
4. Choose **Two-way ANOVA** from the Statistical Test dropdown.
5. Click **Perform Power Analysis**. The results panel shows Cohen's f and required n per cell for each of the three effects (Factor A, Factor B, Interaction A×B) separately.
6. Click **Save Excel Report** to save the full results including the per-effect breakdown sheet.

### Example 4: Prospective analysis (no pilot data)

1. Check **Override effect size manually**.
2. Enter an effect size from the literature (Cohen's d for t-test; Cohen's f for ANOVA).
3. Set Alpha and Power, choose the test type, and click **Perform Power Analysis**.

## Troubleshooting

| Problem | Solution |
|---|---|
| **File fails to load** | Ensure the file is a valid `.xlsx` or `.csv` with at least one numeric column and no password protection. |
| **Columns not appearing in dropdowns** | Only numeric columns are loaded. Check that your data columns contain numbers and not text or mixed types. |
| **"Please select a valid column" error** | Ensure a column is selected in every group dropdown before running the analysis. |
| **Alpha / Power field highlighted red** | Both values must be numbers strictly between 0 and 1. |
| **"Not normal — Failed — used raw data"** | Neither log nor Box-Cox normalised the column. Results may be less reliable; consider a non-parametric test or consult a statistician. |
| **Power curve not shown** | Install matplotlib (`pip install matplotlib`) and restart the tool. |
| **Two-way ANOVA gives very large n** | This can occur when one effect has a very small Cohen's f. Check whether the interaction or a main effect is the limiting factor in the per-effect breakdown. |

## Changelog

### v0.4.0
- **[S1] Correct effect size for ANOVA:** Replaced averaged Cohen's d with Cohen's f (σ_between / σ_within) throughout all ANOVA analyses. Using d with `FTestAnovaPower` produced statistically invalid sample-size estimates.
- **[S2] Absolute Cohen's d:** The t-test effect size is now taken as `abs(d)` so that the sign of the mean difference does not affect the power calculation.
- **[S3] Two-way ANOVA rewritten:** Cell-mean algebra now correctly decomposes pilot data into three independent Cohen's f values (Factor A, Factor B, Interaction A×B). Each effect is powered separately and the most conservative n per cell is reported.
- **[S4] Deferred normalisation:** Normality checking and transformation now run at analysis time on the selected columns only. Previously, every column was transformed at file load, silently corrupting non-selected columns.
- **[S5] Shapiro-Wilk guards:** Added protection for n < 3 (assume normal) and n > 5,000 (subsample to 5,000).
- **[S6] NaN handling:** `dropna()` is applied before all statistical calculations to prevent silent propagation of missing values.
- **[S7] Biological-replicate assumption documented:** The assumption underlying the estimate is now stated explicitly in the results panel and Excel report.
- **[C1–C2] Layout fixes:** Frame-based layout prevents widget overlap when many groups are selected; One-way ANOVA helper widgets are correctly hidden when switching test types.
- **[C3] Input validation:** All group dropdowns are checked for a valid column selection before analysis runs.
- **[C4] Single-pass computation:** Normality and transformation results are cached and reused by the Save Report function — no redundant recalculation.
- **[U1] Data preview:** First five rows displayed after loading.
- **[U2–U3] Persistent results panel:** Results are shown inline instead of in a dismissible message box.
- **[U4] Separate Save button:** "Perform Analysis" and "Save Excel Report" are now independent buttons; Save is disabled until analysis has been run.
- **[U5] Inline validation:** Alpha and Power fields turn red when out of range.
- **[U6] Prospective analysis:** Manual effect-size entry supports power analysis without pilot data.
- **[U7] Power curve:** Embedded matplotlib plot shows power vs. n with target and required-n markers.
- **[U8–U10] Scrollable window, Reset button, and tooltips** on all key parameters.

### v0.3.0
- Initial public release.

## License

This project is licensed under the GNU General Public License v3.0. See the [LICENSE](LICENSE) file for details.
