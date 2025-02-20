import pandas as pd

def load_data(file_path):
    """Load an Excel file into a DataFrame."""
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
    """Keep only the latest record per customer_id based on process_date and save duplicates separately."""
    verified_carrier_table = verified_carrier_table.sort_values(by=['customer_id', 'process_date'], ascending=[True, False])
    latest_records = verified_carrier_table.drop_duplicates(subset=['customer_id'], keep='first')
    duplicate_records = verified_carrier_table[verified_carrier_table.duplicated(subset=['customer_id'], keep=False)]
    
    grouped = latest_records.groupby([
        'customer_id', 'group_name', 'broker_agency_name'
    ]).agg({
        'matching_platform_company_id': list,
        'matching_company_name': list,
        'matching_platform_partner_id': list,
        'matching_partner_name': list,
        'has_aca_sku': list
    }).reset_index()
    grouped['matching_count'] = grouped['matching_platform_company_id'].apply(len)
    
    return grouped, duplicate_records

def scenario_3(company_additional_data):
    """Identify the list of company_id for each unique MNL Company ID and count occurrences."""
    grouped = company_additional_data.groupby([
        'MNL Company ID', 'MNL Company Name', 'MNL Partner ID', 'MNL Partner Name'
    ]).agg({
        'company_id': list
    }).reset_index()
    grouped['company_count'] = grouped['company_id'].apply(len)
    return grouped

def scenario_4(company_additional_data):
    """Identify the list of MNL Company details for each company_id while keeping details grouped similarly to scenario 1."""
    grouped = company_additional_data.groupby([
        'UHC Company ID'
    ]).agg({
        'MNL Company ID': list,
        'MNL Company Name': list,
        'MNL Partner ID': list,
        'MNL Partner Name': list,
        'has_aca_sku': list
    }).reset_index()
    grouped['mnl_count'] = grouped['MNL Company ID'].apply(len)
    return grouped

def merge_master_file(file_master, scenario_2_df, scenario_4_df):
    """Merge master file with scenario 2 and scenario 4 data."""
    file_master = file_master.merge(scenario_2_df, left_on='Customer ID', right_on='customer_id', how='left')
    file_master = file_master.merge(scenario_4_df, left_on='matching_platform_company_id', right_on='UHC Company ID', how='left')
    
    # Rename columns to final format
    file_master.rename(columns={
        'matching_platform_company_id': 'matching_platform_company_id',
        
        'matching_company_name': 'matching_company_name',
        'matching_platform_partner_id': 'matching_platform_partner_id',
        'matching_partner_name': 'matching_partner_name',
        'has_aca_sku_x': 'has_aca_sku',
        'has_aca_sku_y': 'has_aca_sku (Additional_comp)',
        'MNL Company ID': 'MNL Company ID (Additional_comp)',
        'MNL Company Name': 'MNL Company Name (Additional_comp)',
        'MNL Partner ID': 'MNL Partner ID (Additional_comp)',
        'MNL Partner Name': 'MNL Partner Name (Additional_comp)'
    }, inplace=True)
    
    return file_master

def main():
    file_path_master = "file_master.xlsx"
    file_path_2 = "verified_carrier_table.xlsx"
    file_path_1 = "company_additional_data.xlsx"
    
    file_master = load_data(file_path_master)
    verified_carrier_table = load_data(file_path_2)
    company_additional_data = load_data(file_path_1)
    
    scenario_1_df = scenario_1(verified_carrier_table)
    scenario_2_df = scenario_2(verified_carrier_table)
    scenario_3_df = scenario_3(company_additional_data)
    scenario_4_df = scenario_4(company_additional_data)
    
    file_master = merge_master_file(file_master, scenario_2_df, scenario_4_df)
    
    # Save the results
    file_master.to_excel("file_master_output.xlsx", index=False)
    scenario_1_df.to_excel("scenario_1_output.xlsx", index=False)
    scenario_2_df.to_excel("scenario_2_output.xlsx", index=False)
    scenario_3_df.to_excel("scenario_3_output.xlsx", index=False)
    scenario_4_df.to_excel("scenario_4_output.xlsx", index=False)
    
    print("Master file output saved to file_master_output.xlsx")
    print("Scenario 1 output saved to scenario_1_output.xlsx")
    print("Scenario 2 output saved to scenario_2_output.xlsx")
    print("Scenario 3 output saved to scenario_3_output.xlsx")
    print("Scenario 4 output saved to scenario_4_output.xlsx")

if __name__ == "__main__":
    main()
