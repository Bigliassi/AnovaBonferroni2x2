import pandas as pd
import pingouin as pg
from itertools import product

# Load the data from Excel
file_path = 'E:/ConnectivityMatrix/consolidated_output.xlsx'  # Adjust path accordingly
df = pd.read_excel(file_path)

# Initialize Excel writer for the output
output_file = "RM_ANOVA_results_with_interactions.xlsx"
writer = pd.ExcelWriter(output_file, engine='openpyxl')

# List of columns for repeated measures ANOVA
rm_columns = df.columns[3:]  # Starting from the 4th column, which are electrode combinations

# Function to check sphericity using Mauchly's W Test
def check_sphericity(data, id_column, within, dv):
    """Check sphericity using Mauchly’s W test"""
    return pg.sphericity(data, dv=dv, within=within, subject=id_column)

# Function for mixed ANOVA (Group = between-subject factor, Condition = within-subject factor)
def run_mixed_anova(data, id_column, group_col, condition_col, dv_col):
    """Run Mixed ANOVA on a given column"""
    aov_data = data[[id_column, group_col, condition_col, dv_col]].dropna()
    aov = pg.mixed_anova(dv=dv_col, within=condition_col, between=group_col, subject=id_column, data=aov_data, correction='auto')
    sphericity = check_sphericity(aov_data, id_column, within=condition_col, dv=dv_col)
    return aov, sphericity

# Function to calculate the means for each group and condition
def calculate_means(data, group_col, condition_col, dv_col):
    """Calculate means for each group and condition combination."""
    means = data.groupby([group_col, condition_col])[dv_col].mean().unstack()
    return means

# Function to generate pairwise comparisons between conditions within each group
def compare_conditions_within_groups(posthoc_results, group, conditions):
    """Generate a table of pairwise comparisons between conditions within each group."""
    comparisons = {}
    for condition1, condition2 in product(conditions, repeat=2):
        if condition1 != condition2:
            comparison = posthoc_results.loc[
                (posthoc_results['Contrast'] == 'Condition') &
                (((posthoc_results['A'] == condition1) & (posthoc_results['B'] == condition2)) | 
                ((posthoc_results['A'] == condition2) & (posthoc_results['B'] == condition1))), 'p-unc'
            ]
            p_value = comparison.values[0] if len(comparison) > 0 else 'ns'
            comparisons[f'{condition1} vs {condition2}'] = p_value
    return comparisons

# Function to generate pairwise comparisons between groups within each condition
def compare_groups_within_conditions(posthoc_results, condition, groups):
    """Generate a table of pairwise comparisons between groups within each condition."""
    comparisons = {}
    for group1, group2 in product(groups, repeat=2):
        if group1 != group2:
            comparison = posthoc_results.loc[
                (posthoc_results['Contrast'] == 'Group') &
                (((posthoc_results['A'] == group1) & (posthoc_results['B'] == group2)) | 
                ((posthoc_results['A'] == group2) & (posthoc_results['B'] == group1))), 'p-unc'
            ]
            p_value = comparison.values[0] if len(comparison) > 0 else 'ns'
            comparisons[f'{group1} vs {group2}'] = p_value
    return comparisons

# Function to generate pairwise comparisons for interaction effects (Condition * Group)
def compare_interaction(posthoc_results, means):
    """Generate comparisons for significant interaction effects."""
    interaction_comparisons = posthoc_results.loc[
        posthoc_results['Contrast'] == 'Condition * Group', ['A', 'B', 'p-unc', 'p-corr']
    ]
    return interaction_comparisons, means

# Function to run Bonferroni post-hoc tests for 2x2 ANOVA, only if significant interaction effect
def run_posthoc_tests_if_interaction_significant(data, id_column, group_col, condition_col, dv_col, interaction_p_value):
    """Run Bonferroni post-hoc comparisons only if a significant interaction effect is present."""
    
    # Only run post-hoc tests if the interaction p-value is less than 0.001
    if interaction_p_value < 0.001:
        print(f"Running post-hoc tests for {dv_col} due to significant interaction effect (p = {interaction_p_value})")
        
        # Clean and standardize group and condition values
        data[group_col] = data[group_col].str.strip().str.lower()
        data[condition_col] = data[condition_col].str.strip().str.lower()
        
        # Filter the relevant columns
        posthoc_data = data[[id_column, group_col, condition_col, dv_col]].dropna(subset=[dv_col])
        
        # Perform pairwise comparisons with Bonferroni correction
        posthoc_results = pg.pairwise_tests(
            dv=dv_col, within=condition_col, between=group_col, subject=id_column, data=posthoc_data, padjust='bonferroni'
        )
        
        # Calculate means for each group and condition
        means = calculate_means(posthoc_data, group_col, condition_col, dv_col)
        
        # Get pairwise comparisons for conditions and groups
        groups = posthoc_data[group_col].unique()
        conditions = posthoc_data[condition_col].unique()
        
        condition_comparisons = {}
        group_comparisons = {}
        for group in groups:
            condition_comparisons[group] = compare_conditions_within_groups(posthoc_results, group, conditions)
        
        for condition in conditions:
            group_comparisons[condition] = compare_groups_within_conditions(posthoc_results, condition, groups)
        
        interaction_comparisons, means = compare_interaction(posthoc_results, means)

        return posthoc_results, condition_comparisons, group_comparisons, interaction_comparisons, means
    else:
        print(f"Skipping post-hoc tests for {dv_col} (interaction p = {interaction_p_value})")
        return None, None, None, None, None

# Perform the analysis and save results to Excel
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    for dv_col in rm_columns:
        print(f"Running Mixed ANOVA for: {dv_col}")
        mixed_anova, sphericity = run_mixed_anova(df, 'ID', 'Group', 'Condition', dv_col)

        mixed_anova.to_excel(writer, sheet_name=f'{dv_col}_Mixed_ANOVA', index=False)

        # Check for interaction effect in ANOVA results
        if 'Interaction' in mixed_anova['Source'].values:  # Check for 'Interaction'
            interaction_p_value = mixed_anova.loc[mixed_anova['Source'] == 'Interaction', 'p-unc'].values[0]
        else:
            print(f"Interaction effect not found for {dv_col}")
            continue  # Skip to the next iteration if no interaction effect is present

        # Only run post-hoc tests if there is a significant interaction effect
        posthoc_results, condition_comparisons, group_comparisons, interaction_comparisons, means = run_posthoc_tests_if_interaction_significant(
            df, 'ID', 'Group', 'Condition', dv_col, interaction_p_value)
        
        # Save post-hoc results if available
        if posthoc_results is not None:
            posthoc_results.to_excel(writer, sheet_name=f'{dv_col}_Posthoc_Comparisons', index=False)
            # Save condition and group comparisons
            pd.DataFrame.from_dict(condition_comparisons).to_excel(writer, sheet_name=f'{dv_col}_Condition_Comparisons', index=True)
            pd.DataFrame.from_dict(group_comparisons).to_excel(writer, sheet_name=f'{dv_col}_Group_Comparisons', index=True)
            # Save interaction comparisons
            interaction_comparisons.to_excel(writer, sheet_name=f'{dv_col}_Interaction_Comparisons', index=False)
            # Save means
            means.to_excel(writer, sheet_name=f'{dv_col}_Means', index=True)

print(f"Mixed ANOVA analysis complete. Results saved to {output_file}")
