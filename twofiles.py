import pandas as pd

def load_verified_carrier_data(file_path):
    """Load the verified carrier table into a DataFrame."""
    return pd.read_excel(file_path)

def scenario_1(verified_carrier_table):
    """Identify the list of customer_id for each unique matching_platform_company_id."""
    scenario_1_df = verified_carrier_table.groupby([
        'matching_platform_company_id', 'matching_company_name', 
        'matching_platform_partner_id', 'matching_partner_name', 'has_aca_sku'
    ])['customer_id'].unique().explode().reset_index()
    return scenario_1_df

def scenario_2(verified_carrier_table):
    """Identify the list of matching_platform_company_id for each customer_id while keeping records distinct."""
    scenario_2_df = verified_carrier_table[[
        'customer_id', 'group_name', 'broker_agency_name', 
        'matching_platform_company_id', 'matching_company_name', 
        'matching_platform_partner_id', 'matching_partner_name', 'has_aca_sku'
    ]].drop_duplicates()
    return scenario_2_df

def main():
    file_path_2 = "verified_carrier_table.xlsx"
    
    verified_carrier_table = load_verified_carrier_data(file_path_2)
    
    scenario_1_df = scenario_1(verified_carrier_table)
    scenario_2_df = scenario_2(verified_carrier_table)
    
    # Save the results
    scenario_1_df.to_excel("scenario_1_output.xlsx", index=False)
    scenario_2_df.to_excel("scenario_2_output.xlsx", index=False)
    
    print("Scenario 1 output saved to scenario_1_output.xlsx")
    print("Scenario 2 output saved to scenario_2_output.xlsx")

if __name__ == "__main__":
    main()
