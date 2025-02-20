import pandas as pd

def load_data(file_path):
    """Load the verified carrier table or company additional data into a DataFrame."""
    return pd.read_excel(file_path)

def scenario_1(verified_carrier_table):
    """Identify the list of customer_id, group_name, and broker_agency_name for each unique matching_platform_company_id and count occurrences."""
    grouped = verified_carrier_table.groupby([
        'matching_platform_company_id', 'matching_company_name', 
        'matching_platform_partner_id', 'matching_partner_name', 'has_aca_sku'
    ]).agg({
        'customer_id': list,
        'group_name': list,
        'broker_agency_name': list
    }).reset_index()
    grouped['customer_count'] = grouped['customer_id'].apply(len)
    return grouped

def scenario_2(verified_carrier_table):
    """Identify the list of matching_platform_company_id for each customer_id while keeping group_name and broker_agency_name grouped similarly to scenario 1."""
    grouped = verified_carrier_table.groupby([
        'customer_id', 'group_name', 'broker_agency_name'
    ]).agg({
        'matching_platform_company_id': list,
        'matching_company_name': list,
        'matching_platform_partner_id': list,
        'matching_partner_name': list,
        'has_aca_sku': list
    }).reset_index()
    grouped['matching_count'] = grouped['matching_platform_company_id'].apply(len)
    return grouped

def analyze_overlap(scenario_1_df, scenario_2_df):
    """Analyze overlapping customer_id between scenario 1 and scenario 2."""
    overlap_results = []
    
    for _, row in scenario_1_df.iterrows():
        matching_company_id = row['matching_platform_company_id']
        customers_in_scenario_1 = set(row['customer_id'])
        
        for _, row2 in scenario_2_df.iterrows():
            if row2['customer_id'] in customers_in_scenario_1:
                overlap_results.append({
                    'customer_id': row2['customer_id'],
                    'group_name': row2['group_name'],
                    'broker_agency_name': row2['broker_agency_name'],
                    'matching_platform_company_id': matching_company_id,
                    'overlap_with_scenario_1': True
                })
    
    overlap_df = pd.DataFrame(overlap_results)
    return overlap_df

def scenario_3(company_additional_data):
    """Identify the list of company_id for each unique MNL Company ID and count occurrences."""
    grouped = company_additional_data.groupby([
        'MNL Company ID', 'MNL Company Name', 'MNL Partner ID', 'MNL Partner Name', 'has_aca_sku'
    ]).agg({
        'company_id': list
    }).reset_index()
    grouped['company_count'] = grouped['company_id'].apply(len)
    return grouped

def scenario_4(company_additional_data):
    """Identify the list of MNL Company details for each company_id while keeping details grouped similarly to scenario 1."""
    grouped = company_additional_data.groupby([
        'company_id'
    ]).agg({
        'MNL Company ID': list,
        'MNL Company Name': list,
        'MNL Partner ID': list,
        'MNL Partner Name': list,
        'has_aca_sku': list
    }).reset_index()
    grouped['mnl_count'] = grouped['MNL Company ID'].apply(len)
    return grouped

def main():
    file_path_2 = "verified_carrier_table.xlsx"
    file_path_1 = "company_additional_data.xlsx"
    
    verified_carrier_table = load_data(file_path_2)
    company_additional_data = load_data(file_path_1)
    
    scenario_1_df = scenario_1(verified_carrier_table)
    scenario_2_df = scenario_2(verified_carrier_table)
    scenario_3_df = scenario_3(company_additional_data)
    scenario_4_df = scenario_4(company_additional_data)
    
    overlap_df = analyze_overlap(scenario_1_df, scenario_2_df)
    
    # Save the results
    scenario_1_df.to_excel("scenario_1_output.xlsx", index=False)
    scenario_2_df.to_excel("scenario_2_output.xlsx", index=False)
    scenario_3_df.to_excel("scenario_3_output.xlsx", index=False)
    scenario_4_df.to_excel("scenario_4_output.xlsx", index=False)
    overlap_df.to_excel("overlap_analysis.xlsx", index=False)
    
    print("Scenario 1 output saved to scenario_1_output.xlsx")
    print("Scenario 2 output saved to scenario_2_output.xlsx")
    print("Scenario 3 output saved to scenario_3_output.xlsx")
    print("Scenario 4 output saved to scenario_4_output.xlsx")
    print("Overlap analysis saved to overlap_analysis.xlsx")

if __name__ == "__main__":
    main()
