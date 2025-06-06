#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Jun  6 10:41:44 2025

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

# Version number
VERSION = "0.3.0"

# Global variables
data = None
original_data = None
file_path = None
group_menus = []
group_vars = []
group_labels = []

def is_normal(data, alpha=0.05):
    """Test if the data is normally distributed using the Shapiro-Wilk test."""
    statistic, p_value = stats.shapiro(data)
    return p_value > alpha

def transform_data(data):
    """Apply a transformation to the data to make it more normally distributed."""
    # Try log transformation
    try:
        transformed_data = np.log(data - data.min() + 1)  # Adding a shift to avoid log(0)
        if is_normal(transformed_data):
            return transformed_data
    except ValueError:
        pass

    # Try Box-Cox transformation
    try:
        transformed_data, _ = stats.boxcox(data - data.min() + 1)  # Adding a shift to avoid non-positive values
        if is_normal(transformed_data):
            return transformed_data
    except ValueError:
        pass

    # If transformations do not work, return the original data
    return data

def load_data():
    global data, original_data, file_path
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")])
    if file_path:
        try:
            if file_path.endswith('.xlsx'):
                original_data = pd.read_excel(file_path)
            else:
                original_data = pd.read_csv(file_path)

            # Check and transform data for normality
            data = original_data.copy()
            for column in data.columns:
                if not is_normal(data[column]):
                    data[column] = transform_data(data[column])

            update_column_dropdowns()
            messagebox.showinfo("Success", "Data loaded and checked for normality!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {e}")

def update_column_dropdowns():
    if data is not None:
        columns = list(data.columns)
        for menu in group_menus:
            menu['values'] = columns

def calculate_cohen_d(group1, group2):
    mean1, mean2 = group1.mean(), group2.mean()
    std1, std2 = group1.std(), group2.std()
    n1, n2 = len(group1), len(group2)

    pooled_std = math.sqrt(((n1 - 1) * std1**2 + (n2 - 1) * std2**2) / (n1 + n2 - 2))
    d = (mean1 - mean2) / pooled_std
    return d

def perform_power_analysis(columns, alpha, power, test_type, num_groups=None):
    try:
        if test_type == "t-test":
            group1 = data[columns[0]]
            group2 = data[columns[1]]
            effect_size = calculate_cohen_d(group1, group2)
            analysis = TTestIndPower()
            sample_size = analysis.solve_power(effect_size=effect_size, alpha=alpha, power=power)
        elif test_type == "One-way ANOVA":
            effect_sizes = []
            for i in range(len(columns)):
                for j in range(i + 1, len(columns)):
                    group1 = data[columns[i]]
                    group2 = data[columns[j]]
                    effect_sizes.append(calculate_cohen_d(group1, group2))
            effect_size = sum(effect_sizes) / len(effect_sizes)
            analysis = FTestAnovaPower()
            sample_size = analysis.solve_power(effect_size=effect_size, alpha=alpha, power=power, k_groups=num_groups)
        elif test_type == "Two-way ANOVA":
            effect_sizes = []
            for i in range(len(columns)):
                for j in range(i + 1, len(columns)):
                    group1 = data[columns[i]]
                    group2 = data[columns[j]]
                    effect_sizes.append(calculate_cohen_d(group1, group2))
            effect_size = sum(effect_sizes) / len(effect_sizes)
            analysis = FTestAnovaPower()
            sample_size = analysis.solve_power(effect_size=effect_size, alpha=alpha, power=power, k_groups=len(columns))
        else:
            raise ValueError("Invalid test type selected.")

        return effect_size, sample_size
    except Exception as e:
        messagebox.showerror("Error", f"Failed to perform power analysis: {e}")
        return None, None

def generate_excel_report(columns, effect_size, alpha, power, test_type, sample_size):
    try:
        n_samples = len(data[columns[0]])
        biological_replicates = sample_size / n_samples

        # Check normality and transformation status for each column
        normality_status = []
        transformation_status = []

        for column in columns:
            is_norm = is_normal(original_data[column])
            normality_status.append("Yes" if is_norm else "No")

            if not is_norm:
                transformed_data = transform_data(original_data[column])
                is_transformed_norm = is_normal(transformed_data)
                transformation_status.append("Yes" if is_transformed_norm else "No")
            else:
                transformation_status.append("Not Applicable")

        # Prepare report data
        report_data = {
            "Columns": [", ".join(columns)],
            "Normally Distributed": [", ".join(normality_status)],
            "Transformation Successful": [", ".join(transformation_status)],
            "Effect Size (Cohen's d)": [effect_size],
            "Alpha": [alpha],
            "Power": [power],
            "Test Type": [test_type],
            "Required Sample Size": [sample_size],
            "Required Biological Replicates": [biological_replicates],
        }

        report_df = pd.DataFrame(report_data)

        default_report_name = os.path.splitext(os.path.basename(file_path))[0] + "_PowerAnalysis.xlsx"
        report_file_path = filedialog.asksaveasfilename(initialfile=default_report_name, defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if report_file_path:
            with pd.ExcelWriter(report_file_path) as writer:
                report_df.to_excel(writer, sheet_name="Power Analysis Report", index=False)
                messagebox.showinfo("Success", f"Excel report generated successfully at {report_file_path}!")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to generate Excel report: {e}")

def on_test_type_change(*args):
    clear_group_menus()
    test_type = test_type_var.get()

    if test_type == "Two-way ANOVA":
        create_group_menus(4, ["Factor 1 Level 1", "Factor 1 Level 2", "Factor 2 Level 1", "Factor 2 Level 2"])
    elif test_type == "One-way ANOVA":
        num_groups_label.grid(row=4, column=0, sticky="w", padx=10, pady=5)
        num_groups_entry.grid(row=4, column=1, sticky="ew", padx=10, pady=5)
        update_button.grid(row=5, column=0, columnspan=2, pady=5)
    else:  # Default to t-test
        create_group_menus(2, ["Group 1", "Group 2"])

def clear_group_menus():
    for menu in group_menus:
        menu.grid_forget()
    for label in group_labels:
        label.grid_forget()
    group_menus.clear()
    group_vars.clear()
    group_labels.clear()

def create_group_menus(num_groups, labels):
    for i in range(num_groups):
        label_text = f"{labels[i]} Column:"
        label = tk.Label(root, text=label_text)
        label.grid(row=2 + i, column=0, sticky="w", padx=10, pady=5)
        group_labels.append(label)

        var = tk.StringVar(root)
        menu = ttk.Combobox(root, textvariable=var)
        menu.grid(row=2 + i, column=1, sticky="ew", padx=10, pady=5)
        group_menus.append(menu)
        group_vars.append(var)

    update_column_dropdowns()

def update_group_menus():
    try:
        num_groups = int(num_groups_entry.get())
        if num_groups < 2:
            raise ValueError("Number of groups must be at least 2.")

        clear_group_menus()
        labels = [f"Group {i+1}" for i in range(num_groups)]
        create_group_menus(num_groups, labels)
    except ValueError as e:
        messagebox.showerror("Error", f"Invalid input: {e}")

def on_analyze():
    if data is None:
        messagebox.showerror("Error", "No data loaded. Please load a data file first.")
        return

    try:
        test_type = test_type_var.get()
        alpha = float(alpha_entry.get())
        power = float(power_entry.get())

        columns = [var.get() for var in group_vars]

        if test_type == "One-way ANOVA":
            num_groups = int(num_groups_entry.get())
            effect_size, sample_size = perform_power_analysis(columns, alpha, power, test_type, num_groups)
        else:
            effect_size, sample_size = perform_power_analysis(columns, alpha, power, test_type)

        if sample_size is not None:
            n_samples = len(data[columns[0]])
            biological_replicates = sample_size / n_samples
            messagebox.showinfo("Result", f"Effect Size (Cohen's d): {effect_size:.2f}\nRequired Sample Size per Group: {sample_size:.2f}\nRequired Biological Replicates: {biological_replicates:.2f}")
            generate_excel_report(columns, effect_size, alpha, power, test_type, sample_size)
    except ValueError as e:
        messagebox.showerror("Error", f"Invalid input: {e}")

# Create the main window
root = tk.Tk()
root.title(f"Power Analysis Tool v{VERSION}")

# Instructions
instructions = """
Data Organization Instructions:

- For t-test: Organize your data into two columns, each representing a different group.
- For One-way ANOVA: Each group should be in a separate column. Use descriptive column names.
- For Two-way ANOVA: Each combination of factor levels should be in a separate column.
  For example, if you have factors A and B with two levels each, use columns named A1B1, A1B2, A2B1, A2B2.
"""
instructions_label = tk.Label(root, text=instructions, justify=tk.CENTER)
instructions_label.grid(row=0, column=0, columnspan=2, pady=10, padx=10)

# File loader
load_button = tk.Button(root, text="Load Data File", command=load_data)
load_button.grid(row=1, column=0, columnspan=2, pady=10)

# Number of groups input for One-way ANOVA
num_groups_label = tk.Label(root, text="Number of Groups:")
num_groups_entry = tk.Entry(root)
update_button = tk.Button(root, text="Update", command=update_group_menus)

# Parameters input
tk.Label(root, text="Alpha:").grid(row=6, column=0, sticky="w", padx=10, pady=5)
alpha_entry = tk.Entry(root)
alpha_entry.insert(0, "0.05")
alpha_entry.grid(row=6, column=1, sticky="ew", padx=10, pady=5)

tk.Label(root, text="Power:").grid(row=7, column=0, sticky="w", padx=10, pady=5)
power_entry = tk.Entry(root)
power_entry.insert(0, "0.8")
power_entry.grid(row=7, column=1, sticky="ew", padx=10, pady=5)

# Dropdown for test type
tk.Label(root, text="Statistical Test:").grid(row=8, column=0, sticky="w", padx=10, pady=5)
test_type_var = tk.StringVar(root)
test_type_var.set("t-test")
test_type_var.trace_add("write", on_test_type_change)
test_type_menu = tk.OptionMenu(root, test_type_var, "t-test", "One-way ANOVA", "Two-way ANOVA")
test_type_menu.grid(row=8, column=1, sticky="ew", padx=10, pady=5)

# Button to perform analysis
analyze_button = tk.Button(root, text="Perform Power Analysis and Generate Report", command=on_analyze)
analyze_button.grid(row=9, column=0, columnspan=2, pady=20)

# Configure grid weights
root.grid_columnconfigure(1, weight=1)

# Initialize with t-test dropdowns
create_group_menus(2, ["Group 1", "Group 2"])

# Run the application
root.mainloop()







