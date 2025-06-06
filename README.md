# vPowerAnalysisTool

The vPowerAnalysis Tool is a graphical user interface (GUI) application designed to help researchers perform power analysis for t-tests, one-way ANOVA, and two-way ANOVA. This tool calculates the required sample size and the number of biological replicates needed to achieve a specified statistical power. It also checks for data normality and applies transformations if necessary to ensure the validity of parametric tests. It is expecting n data from one biological replicate.

## Table of Contents

1. [Installation](#installation)
2. [Usage](#usage)
3. [Data Organization](#data-organization)
4. [Calculations](#calculations)
5. [Examples](#examples)
6. [Troubleshooting](#troubleshooting)
7. [License](#license)

## Installation

To run the Power Analysis Tool, you need to have Python installed on your computer. Additionally, you need to install the following Python libraries:

- pandas
- tkinter
- statsmodels
- scipy
- numpy
- openpyxl

```bash
pip install pandas tkinter statsmodels scipy numpy openpyxl
```

## Usage

### Step 1: Load Data

1. Click the "Load Data File" button to open a file dialog.
2. Select an Excel or CSV file containing your data.
3. The tool will automatically check for data normality and apply transformations if necessary.
4. The column names from the file will be displayed in the dropdown menus.

### Step 2: Select Columns

- **For t-test**: Select the columns representing the two groups you want to compare from the "Group 1 Column" and "Group 2 Column" dropdown menus.

- **For One-way ANOVA**:
  1. Enter the number of groups in the "Number of Groups" field.
  2. Click "Update" to generate the appropriate number of dropdown menus.
  3. Select the columns representing each group from the generated dropdown menus.

- **For Two-way ANOVA**: Select the columns representing the factor levels from the "Factor 1 Level 1 Column," "Factor 1 Level 2 Column," "Factor 2 Level 1 Column," and "Factor 2 Level 2 Column" dropdown menus.

### Step 3: Set Parameters

- **Alpha**: The significance level (default is 0.05).
- **Power**: The statistical power (default is 0.8).

### Step 4: Choose Statistical Test

Select the type of statistical test you want to perform from the "Statistical Test" dropdown menu:

- t-test
- One-way ANOVA
- Two-way ANOVA

### Step 5: Perform Power Analysis

Click the "Perform Power Analysis and Generate Report" button to perform the power analysis. The results will be displayed in a message box, and you will be prompted to save an Excel report.

### Step 6: Save the Report

1. An Excel report will be generated with the results of the power analysis, including information on data normality and transformation success.
2. Choose the location and name for the report file.
3. Click "Save" to save the report.

## Data Organization

### For t-test

Organize your data into two columns, each representing a different group.

| Group1 | Group2 |
|--------|--------|
| 10 | 7 |
| 12 | 9 |
| 14 | 11 |

### For One-way ANOVA

Each group should be in a separate column. Use descriptive column names.

| Group1 | Group2 | Group3 |
|--------|--------|--------|
| 10 | 7 | 8 |
| 12 | 9 | 10 |
| 14 | 11 | 12 |

### For Two-way ANOVA

Each combination of factor levels should be in a separate column. For example, if you have factors A and B with two levels each, use columns named A1B1, A1B2, A2B1, A2B2.

| A1B1 | A1B2 | A2B1 | A2B2 |
|------|------|------|------|
| 10 | 7 | 8 | 9 |
| 12 | 9 | 10 | 11 |
| 14 | 11 | 12 | 13 |

## Calculations

### Cohen's d

Cohen's d is a measure of effect size used to indicate the standard difference between two means. It is calculated as follows:

d = (M1 - M2) / s_pooled
\[ d = \frac{M_1 - M_2}{s_{\text{pooled}}} \]

where \( M_1 \) and \( M_2 \) are the means of the two groups, and \( s_{\text{pooled}} \) is the pooled standard deviation, calculated as:

s_pooled = sqrt( ((n1 - 1) * s1^2 + (n2 - 1) * s2^2) / (n1 + n2 - 2) )
\[ s_{\text{pooled}} = \sqrt{\frac{(n_1 - 1)s_1^2 + (n_2 - 1)s_2^2}{n_1 + n_2 - 2}} \]

where \( s_1 \) and \( s_2 \) are the standard deviations of the two groups, and \( n_1 \) and \( n_2 \) are the sample sizes.

### Power Analysis

Power analysis is used to determine the sample size required to detect an effect of a given size with a certain level of confidence. The required sample size is calculated using the `statsmodels` library in Python, which provides functions for performing power analysis for t-tests and ANOVA.

### Biological Replicates

The number of biological replicates required is calculated by dividing the required sample size by the number of samples in the provided dataset:

Biological Replicates = Required Sample Size / n

\[ \text{Biological Replicates} = \frac{\text{Required Sample Size}}{n} \]

where \( n \) is the number of samples in the provided dataset.

## Examples

### Example 1: t-test

1. Load a data file with two columns representing two groups.
2. Select the columns for Group 1 and Group 2.
3. Set Alpha to 0.05 and Power to 0.8.
4. Choose "t-test" from the Statistical Test dropdown menu.
5. Click "Perform Power Analysis and Generate Report."
6. Save the Excel report.

### Example 2: One-way ANOVA

1. Load a data file with multiple columns representing different groups.
2. Enter the number of groups in the "Number of Groups" field and click "Update."
3. Select the columns for each group from the generated dropdown menus.
4. Set Alpha to 0.05 and Power to 0.8.
5. Choose "One-way ANOVA" from the Statistical Test dropdown menu.
6. Click "Perform Power Analysis and Generate Report."
7. Save the Excel report.

### Example 3: Two-way ANOVA

1. Load a data file with four columns representing the combinations of factor levels.
2. Select the columns for Factor 1 Level 1, Factor 1 Level 2, Factor 2 Level 1, and Factor 2 Level 2.
3. Set Alpha to 0.05 and Power to 0.8.
4. Choose "Two-way ANOVA" from the Statistical Test dropdown menu.
5. Click "Perform Power Analysis and Generate Report."
6. Save the Excel report.

## Troubleshooting

- **Data Loading Issues**: Ensure that your data file is correctly formatted and that the column names are properly defined.
- **Dropdown Menus Not Updating**: If the dropdown menus do not update correctly when switching test types, ensure that the data is loaded and the column names are correctly specified.
- **Analysis Errors**: If you encounter errors during the analysis, verify that the input parameters (Alpha, Power) are correctly set and that the selected columns contain valid data.
- **Normality and Transformation Issues**: If the tool indicates that data is not normally distributed and transformations were not successful, consider using non-parametric tests or consulting a statistician for further advice.

## License

This project is licensed under the GNU General Public License v3.0. See the [LICENSE](LICENSE) file for details.
