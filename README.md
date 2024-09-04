import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

# Load the Excel files
min_groups = pd.read_excel('Min_Groups.xlsx')
uhc_newgroups = pd.read_excel('UHC_Aug_Newgroups.xlsx')

# Ensure all relevant columns are strings and fill NaN values with empty strings
min_groups['name'] = min_groups['name'].fillna('').astype(str)
min_groups['l_address1'] = min_groups['l_address1'].fillna('').astype(str)
min_groups['l_city'] = min_groups['l_city'].fillna('').astype(str)
min_groups['l_state'] = min_groups['l_state'].fillna('').astype(str)
min_groups['l_zip'] = min_groups['l_zip'].fillna('').astype(str)

uhc_newgroups['Group Name'] = uhc_newgroups['Group Name'].fillna('').astype(str)
uhc_newgroups['Group Address 1'] = uhc_newgroups['Group Address 1'].fillna('').astype(str)
uhc_newgroups['Group City'] = uhc_newgroups['Group City'].fillna('').astype(str)
uhc_newgroups['Group State'] = uhc_newgroups['Group State'].fillna('').astype(str)
uhc_newgroups['Group Zip'] = uhc_newgroups['Group Zip'].fillna('').astype(str)

# Define similarity thresholds (adjust as needed)
name_threshold = 0.8
address_threshold = 0.75
city_threshold = 0.85

# Helper function for cosine similarity
def compute_similarity(series1, series2):
    vectorizer = TfidfVectorizer().fit(pd.concat([series1, series2]))
    series1_tfidf = vectorizer.transform(series1)
    series2_tfidf = vectorizer.transform(series2)
    return cosine_similarity(series1_tfidf, series2_tfidf)

# Compute similarities for each field
name_similarity = compute_similarity(uhc_newgroups['Group Name'], min_groups['name'])
address_similarity = compute_similarity(uhc_newgroups['Group Address 1'], min_groups['l_address1'])
city_similarity = compute_similarity(uhc_newgroups['Group City'], min_groups['l_city'])

# Initialize a list to store all matches
all_matches = []

# Iterate over UHC records
for i in range(len(uhc_newgroups)):
    # For each record in UHC_Aug_Newgroups, find all matches in Min_Groups
    for j in range(len(min_groups)):
        # Check individual similarities against thresholds
        if name_similarity[i, j] > name_threshold and address_similarity[i, j] > address_threshold and city_similarity[i, j] > city_threshold:
            # Ensure state and zip are exact matches
            state_match = uhc_newgroups.iloc[i]['Group State'] == min_groups.iloc[j]['l_state']
            zip_match = uhc_newgroups.iloc[i]['Group Zip'] == min_groups.iloc[j]['l_zip']

            # If state and zip match, calculate total score
            if state_match and zip_match:
                # Calculate individual similarity score
                similarity_score = {
                    'name': name_similarity[i, j],
                    'address': address_similarity[i, j],
                    'city': city_similarity[i, j],
                    'state': 1.0 if state_match else 0.0,  # Exact match for state
                    'zip': 1.0 if zip_match else 0.0       # Exact match for zip
                }
                # Calculate overall score (weighted)
                overall_score = (similarity_score['name'] * 0.4 +
                                 similarity_score['address'] * 0.3 +
                                 similarity_score['city'] * 0.2 +
                                 similarity_score['state'] * 0.05 +
                                 similarity_score['zip'] * 0.05)

                all_matches.append({
                    'UHC_Group_Name': uhc_newgroups.iloc[i]['Group Name'],
                    'UHC_Group_Address': uhc_newgroups.iloc[i]['Group Address 1'],
                    'UHC_Group_City': uhc_newgroups.iloc[i]['Group City'],
                    'UHC_Group_State': uhc_newgroups.iloc[i]['Group State'],
                    'UHC_Group_Zip': uhc_newgroups.iloc[i]['Group Zip'],
                    'Best_Match_NAME': min_groups.iloc[j]['name'],
                    'Best_Match_Address': min_groups.iloc[j]['l_address1'],
                    'Best_Match_City': min_groups.iloc[j]['l_city'],
                    'Best_Match_State': min_groups.iloc[j]['l_state'],
                    'Best_Match_Zip': min_groups.iloc[j]['l_zip'],
                    'platform_id': min_groups.iloc[j]['platform_id'],  # From min_groups
                    'p_name': min_groups.iloc[j]['p_name'],  # From min_groups
                    'ultimate_parent_name': min_groups.iloc[j]['ultimate_parent_name'],  # From min_groups
                    'Similarity_Score': overall_score,  # Overall similarity score
                    'Individual_Similarity': similarity_score  # To track individual scores if needed
                })

# Convert the matches into a DataFrame
matches_df = pd.DataFrame(all_matches)

# Merge the best matches and overall similarity score into the UHC_Aug_Newgroups file
uhc_newgroups_with_matches = uhc_newgroups.merge(
    matches_df[['UHC_Group_Name', 'Best_Match_NAME', 'Best_Match_Address', 'Best_Match_City', 'Best_Match_State', 'Best_Match_Zip', 'platform_id', 'p_name', 'ultimate_parent_name', 'Similarity_Score']],
    left_on='Group Name', right_on='UHC_Group_Name',
    how='left'
)

# Drop the extra 'UHC_Group_Name' column as it's now redundant
uhc_newgroups_with_matches.drop(columns=['UHC_Group_Name'], inplace=True)

# Save the updated UHC_Aug_Newgroups file with matches and scores
uhc_newgroups_with_matches.to_excel('UHC_Aug_Newgroups_with_matches.xlsx', index=False)

# Optionally, copy to clipboard
uhc_newgroups_with_matches.to_clipboard(index=False)

print("Matching complete. Results saved to 'UHC_Aug_Newgroups_with_matches.xlsx' and copied to clipboard.")
