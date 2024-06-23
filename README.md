import streamlit as st
import pandas as pd
import base64

# Load existing data
file_path = 'UHC_Brokers_MNL_Partners.xlsx'

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
    href = f'<a href="data:application/octet-stream;base64,{b64}" download="UHC_Brokers_MNL_Partners.xlsx">Download Updated Excel File</a>'
    return href

# Streamlit UI
st.set_page_config(page_title='UHC Brokers & MNL Partners', layout='wide')
st.markdown("""
<style>
    .main {
        background-color: #f0f2f6;
        padding: 20px;
        border-radius: 10px;
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
    }
    .stDataFrame {
        border: 2px solid #ddd;
        border-radius: 8px;
        padding: 10px;
        background-color: white;
    }
    .sidebar .sidebar-content {
        padding: 20px;
    }
    .sidebar .sidebar-content img {
        max-width: 100%;
        border-radius: 10px;
    }
    .sidebar .sidebar-content h1 {
        font-size: 24px;
        margin-top: 10px;
    }
    .sidebar .sidebar-content a {
        display: block;
        padding: 10px 0;
        color: #333;
        text-decoration: none;
    }
    .sidebar .sidebar-content a:hover {
        color: #4CAF50;
    }
    .title {
        font-size: 32px;
        font-weight: bold;
        margin-bottom: 20px;
        color: #333;
    }
    .subtitle {
        font-size: 24px;
        font-weight: bold;
        margin-bottom: 10px;
        color: #4CAF50;
    }
</style>
""", unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    st.image('https://via.placeholder.com/150', use_column_width=True)
    st.title("Navigation")
    st.markdown("""
    - [Current Data](#current-data)
    - [Add New Data](#add-new-data)
    - [Download File](#download-the-updated-file)
    """)

# Main Content
st.markdown('<div class="title">🔄 UHC Brokers and MNL Partners Update Portal</div>', unsafe_allow_html=True)

# Display existing data
df = load_data(file_path)
st.markdown('<div class="subtitle">📊 Current Data:</div>', unsafe_allow_html=True)
st.dataframe(df, use_container_width=True, height=400)

# Input form for new data
st.markdown('<div class="subtitle">📝 Add New Data:</div>', unsafe_allow_html=True)
with st.form(key='data_form'):
    col1, col2 = st.columns(2)
    with col1:
        uhc_broker = st.text_input('UHC Broker')
    with col2:
        mnl_partner = st.text_input('MNL Partner')
    submit_button = st.form_submit_button(label='Add Row')

# Update data if form is submitted
if submit_button:
    if uhc_broker and mnl_partner:
        new_row = pd.DataFrame({'UHC Brokers': [uhc_broker], 'MNL Partner': [mnl_partner]})
        df = pd.concat([df, new_row], ignore_index=True)
        save_data(df, file_path)
        st.success('✅ New row added successfully!')
        # Reload the data to display the updated DataFrame
        df = load_data(file_path)
        st.markdown('<div class="subtitle">🔄 Updated Data:</div>', unsafe_allow_html=True)
        st.dataframe(df, use_container_width=True, height=400)
    else:
        st.error('⚠️ Please fill out both fields.')

# Display download link for updated file
st.markdown('<div class="subtitle">📥 Download the updated file:</div>', unsafe_allow_html=True)
st.markdown(get_table_download_link(file_path), unsafe_allow_html=True)
