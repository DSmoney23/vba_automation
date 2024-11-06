```python
import pandas as pd
import numpy as np
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import re
import os
from datetime import datetime


# Define file paths and create timestamp for backup file
FILE_PATH = r"C:\Users\johnsmith\Documents\UHC\Mass_Termination"
MASS_TERM_FILE = os.path.join(FILE_PATH, "UHC_Mass_Term.xlsx")
NOV_LF_FILE = os.path.join(FILE_PATH, "UHC_Nov_LF.xlsx")
TIMESTAMP = datetime.now().strftime("%Y%m%d_%H%M%S")
OUTPUT_FILE = os.path.join(FILE_PATH, f"matching_results_{TIMESTAMP}.csv")


def preprocess_text(text):
    """Clean and standardize text for better matching"""
    if pd.isna(text):
        return ""
    # Convert to string if not already
    text = str(text).strip()
    # Convert to lowercase and remove special characters
    text = re.sub(r'[^a-zA-Z0-9\s]', '', text.lower())
    # Remove extra whitespace
    text = ' '.join(text.split())
    return text


def calculate_similarity_scores(series1, series2):
    """Calculate cosine similarity between two series of text"""
    try:
        # Preprocess all texts
        processed_series1 = series1.apply(preprocess_text)
        processed_series2 = series2.apply(preprocess_text)
        
        # Combine all texts for vectorization
        all_texts = pd.concat([processed_series1, processed_series2])
        
        # Create TF-IDF vectors
        vectorizer = TfidfVectorizer(analyzer='char', ngram_range=(2,3))
        tfidf_matrix = vectorizer.fit_transform(all_texts)
        
        # Split the matrix back into two parts
        matrix1 = tfidf_matrix[:len(series1)]
        matrix2 = tfidf_matrix[len(series1):]
        
        # Calculate cosine similarity
        similarities = cosine_similarity(matrix1, matrix2)
        
        return similarities
    except Exception as e:
        print(f"Error in calculate_similarity_scores: {str(e)}")
        raise


def find_best_matches(mass_term_df, nov_lf_df):
    """Find best matches between the two dataframes"""
    try:
        print("\nStarting similarity calculations...")
        
        # Calculate similarity scores for each field
        print("Calculating name similarities...")
        name_similarities = calculate_similarity_scores(
            mass_term_df['name'], 
            nov_lf_df['Group Name']
        )
        
        print("Calculating city similarities...")
        city_similarities = calculate_similarity_scores(
            mass_term_df['l_city'], 
            nov_lf_df['Group City']
        )
        
        print("Calculating state similarities...")
        state_similarities = calculate_similarity_scores(
            mass_term_df['l_state'], 
            nov_lf_df['Group State']
        )
        
        print("Calculating zip similarities...")
        zip_similarities = calculate_similarity_scores(
            mass_term_df['l_zip'], 
            nov_lf_df['Group Zip Code']
        )
        
        # Initialize lists to store best matches
        best_matches = []
        
        print("\nFinding best matches for each group...")
        total_rows = len(mass_term_df)
        
        # For each row in mass_term_df
        for i in range(total_rows):
            # Print progress every 1000 rows
            if (i + 1) % 1000 == 0 or (i + 1) == total_rows:
                print(f"Processing row {i + 1} of {total_rows}")
            
            # Get indices of best matches based on name similarity
            best_match_idx = np.argmax(name_similarities[i])
            
            # Get all similarity scores for this match
            match = {
                'Matched_Group_Name': nov_lf_df['Group Name'].iloc[best_match_idx],
                'Matched_Group_City': nov_lf_df['Group City'].iloc[best_match_idx],
                'Matched_Group_State': nov_lf_df['Group State'].iloc[best_match_idx],
                'Matched_Group_Zipcode': nov_lf_df['Group Zip Code'].iloc[best_match_idx],
                'Name_Similarity': round(name_similarities[i][best_match_idx], 4),
                'City_Similarity': round(city_similarities[i][best_match_idx], 4),
                'State_Similarity': round(state_similarities[i][best_match_idx], 4),
                'Zip_Similarity': round(zip_similarities[i][best_match_idx], 4)
            }
            best_matches.append(match)
        
        # Convert matches to DataFrame
        matches_df = pd.DataFrame(best_matches)
        
        # Combine with original mass_term_df
        result_df = pd.concat([mass_term_df.reset_index(drop=True), matches_df], axis=1)
        
        # Sort by Name_Similarity in descending order
        result_df = result_df.sort_values('Name_Similarity', ascending=False)
        
        return result_df
    
    except Exception as e:
        print(f"Error in find_best_matches: {str(e)}")
        raise


def main():
    try:
        print("\n=== Group Matching Process Started ===")
        print(f"Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"Working Directory: {FILE_PATH}")
        
        # Check if directory exists
        if not os.path.exists(FILE_PATH):
            raise FileNotFoundError(f"Directory not found: {FILE_PATH}")
        
        # Check if input files exist
        if not os.path.exists(MASS_TERM_FILE):
            raise FileNotFoundError(f"File not found: {MASS_TERM_FILE}")
        if not os.path.exists(NOV_LF_FILE):
            raise FileNotFoundError(f"File not found: {NOV_LF_FILE}")
        
        # Read the input files
        print("\nReading input files...")
        mass_term_df = pd.read_excel(MASS_TERM_FILE)
        nov_lf_df = pd.read_excel(NOV_LF_FILE)
        
        print(f"Number of records in Mass Term file: {len(mass_term_df)}")
        print(f"Number of records in November LF file: {len(nov_lf_df)}")
        
        # Process the data and get matches
        result_df = find_best_matches(mass_term_df, nov_lf_df)
        
        # Copy to clipboard
        print("\nCopying results to clipboard...")
        result_df.to_clipboard(excel=True, index=False)
        print("Results have been copied to clipboard successfully!")
        
        # Save as CSV for backup
        print(f"\nSaving backup file...")
        result_df.to_csv(OUTPUT_FILE, index=False)
        print(f"Backup saved as: {OUTPUT_FILE}")
        
        print("\n=== Process Completed Successfully ===")
        print(f"Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
    except Exception as e:
        print(f"\nERROR: {str(e)}")
        print("\n=== Process Failed ===")
        raise
    

if __name__ == "__main__":
    main()
```
