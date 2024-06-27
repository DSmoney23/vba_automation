import streamlit as st
import pandas as pd
import base64

# Load existing data
file_path = r'C:\Users\john.smith\documents\streamlit_file.xlsx'

@st.cache_resource
def load_data(file_path):
    return pd.read_excel(file_path)

# Save updated data
def save_data(data, file_path):
    data.to_excel(file_path, index=False)

# Function to get download link
def get_table_download_link(file_path):
    with open(file_path, 'rb') as f:
        b64 = base64.b64encode(f.read()).decode()
    href = f'<a href="data:application/octet-stream;base64,{b64}" download="streamlit_file.xlsx">Download Updated Excel File</a>'
    return href

# Streamlit UI
st.set_page_config(page_title='UHC Brokers & MNL Partners', layout='wide')

# Custom CSS for cool background and modern styling
st.markdown(
    """
    <style>
    body {
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
        font-family: 'Arial', sans-serif;
    }
    .main {
        background-color: rgba(255, 255, 255, 0.8);
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 4px 8px rgba(0, 0, 0, 0.2);
    }
    .stButton>button {
        background-color: #4CAF50;
        color: white;
        border-radius: 12px;
        padding: 10px 20px;
        font-size: 16px;
        margin: 10px 0px;
    }
    .stTextInput>div>input {
        padding: 10px;
        border: 2px solid #ddd;
        border-radius: 8px;
        margin-bottom: 10px;
    }
    .stDataFrame {
        border: 2px solid #ddd;
        border-radius: 8px;
        padding: 10px;
        background-color: white;
    }
    .title {
        font-size: 32px;
        font-weight: bold;
        color: #333;
        text-align: center;
        margin-top: 20px;
    }
    .subtitle {
        font-size: 24px;
        font-weight: bold;
        color: #4CAF50;
        margin-bottom: 10px;
        text-align: center;
    }
    </style>
    """,
    unsafe_allow_html=False
)

# Main Content
st.markdown('<div class="title">🔄 UHC Brokers and MNL Partners Update Portal</div>', unsafe_allow_html=False)

# Display existing data
df = load_data(file_path)
st.markdown('<div class="subtitle">📊 Current Data:</div>', unsafe_allow_html=False)
st.dataframe(df[['MNL Partner ID', 'MNL Partner', 'UHC Agency', 'Month', 'SF Partner ID']], use_container_width=True, height=400)

# Input form for new data
st.markdown('<div class="subtitle">📝 Add New Data:</div>', unsafe_allow_html=False)
with st.form(key='data_form'):
    col1, col2 = st.columns(2)
    with col1:
        mnl_partner_id = st.text_input('MNL Partner ID')
        mnl_partner = st.text_input('MNL Partner')
        uhc_agency = st.text_input('UHC Agency')
    with col2:
        month = st.text_input('Month')
        sf_partner_id = st.text_input('SF Partner ID')
    submit_button = st.form_submit_button(label='Add Row')

# Update data if form is submitted
if submit_button:
    if mnl_partner_id and mnl_partner and uhc_agency and month and sf_partner_id:
        new_row = pd.DataFrame({
            'MNL Partner ID': [mnl_partner_id],
            'MNL Partner': [mnl_partner],
            'UHC Agency': [uhc_agency],
            'Month': [month],
            'SF Partner ID': [sf_partner_id]
        })
        df = pd.concat([df, new_row], ignore_index=True)
        save_data(df, file_path)
        st.success('✅ New row added successfully!')
        # Reload the data to display the updated DataFrame
        df = load_data(file_path)
        st.markdown('<div class="subtitle">🔄 Updated Data:</div>', unsafe_allow_html=False)
        st.dataframe(df[['MNL Partner ID', 'MNL Partner', 'UHC Agency', 'Month', 'SF Partner ID']], use_container_width=True, height=400)
    else:
        st.error('⚠️ Please fill out all fields.')

# Display download link for updated file
st.markdown('<div class="subtitle">📥 Download the updated file:</div>', unsafe_allow_html=False)
st.markdown(get_table_download_link(file_path), unsafe_allow_html=False)
