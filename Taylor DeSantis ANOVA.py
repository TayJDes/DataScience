# -*- coding: utf-8 -*-
"""
Created on Thu Jun 27 10:49:19 2024
@author: taylor

This script performs ANOVA (Analysis of Variance) on selected test scores across
different categories from a dataset. It checks for normality and homogeneity of variances,
runs ANOVA regardless of the assumption checks, and saves the results along with any
issues found during the assumptions checks into a DataFrame. Finally, it exports these
results into an Excel file for further analysis.
"""
import pandas as pd
import scipy.stats as stats

# Define column names and categories for the ANOVA
column_names = ['Current Status', 'Term Type', 'Completed ATD Class?', 'Attendance']
categories = {
    'Current Status': ['Active', 'Prehire', 'Terminated'],
    'Term Type': ['RTS', 'Discharge', 'Voluntary Quit'],
    'Completed ATD Class?': ['No - Failed Class', 'Yes', 'Prehire'],
    'Attendance': ['Poor', 'Average', 'Excellent']
}

# Define test names for which the ANOVA will be performed
test_names = ['MultiChat Score', 'Insight Score']

def load_excel():
    """Load data from an Excel file."""
    directory = "C://Users//Taylo//OneDrive//Documents//GitHub//Data Analysis//Data Files//"
    file = "Client Cleaned Data File - Working Sheet Tab.xlsx"
    return pd.read_excel(directory + file, sheet_name="Working Sheet")

def check_assumptions(df, column_name, categories, test_column):
    """Check normality and homogeneity of variances for each category."""
    groups = [df[df[column_name] == category][test_column].dropna() for category in categories]
    assumption_issues = []
    
    for category, group in zip(categories, groups):
        stat, p = stats.shapiro(group)
        if p < 0.05:
            assumption_issues.append(f"Normality test failed for {category} (p={p:.3f})")
    
    if len(groups) > 1:
        stat, p = stats.levene(*groups)
        if p < 0.05:
            assumption_issues.append(f"Levene's test failed (p={p:.3f})")
    
    return groups, assumption_issues

def anova_test(groups, test_column, column_name):
    """Perform ANOVA and return results, handling insufficient data and exceptions."""
    if any(len(group) < 15 for group in groups):
        return float('nan'), float('nan'), "Insufficient data"
    try:
        anova_result = stats.f_oneway(*groups)
        return anova_result.statistic, anova_result.pvalue, "Success"
    except Exception as e:
        return float('nan'), float('nan'), str(e)

df = load_excel()

results = []
for column in column_names:
    for test_name in test_names:
        groups, issues = check_assumptions(df, column, categories[column], test_name)
        F, p, status = anova_test(groups, test_name, column)
        results.append({
            'Category': column,
            'Test': test_name,
            'F-Statistic': F,
            'p-Value': p,
            'Status': status,
            'Assumption Issues': "; ".join(issues)
        })

results_df = pd.DataFrame(results)
print(results_df)

# Save results to an Excel file
results_df.to_excel("anova_results.xlsx", index=False)
