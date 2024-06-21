
Copy code
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

# Streamlit UI
st.set_page_config(page_title='UHC Brokers & MNL Partners', layout='wide')
st.title('🔄 UHC Brokers and MNL Partners Update Portal')
st.markdown("""
<style>
    .main {
        background-color: #F5F5F5;
    }
    .stButton>button {
        background-color: #4CAF50;
        color: white;
        border-radius: 12px;
        padding: 10px;
        font-size: 16px;
        margin: 10px 0px;
    }
    .stTextInput>div>input {
        padding: 10px;
        border: 2px solid #ddd;
        border-radius: 8px;
    }
</style>
""", unsafe_allow_html=True)

# Display existing data
df = load_data(file_path)
st.subheader('📊 Current Data:')
st.dataframe(df, use_container_width=True)

# Input form for new data
st.subheader('📝 Add New Data:')
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
        st.experimental_rerun()
    else:
        st.error('⚠️ Please fill out both fields.')

# Display download link for updated file
def get_table_download_link(file_path):
    with open(file_path, 'rb') as f:
        b64 = base64.b64encode(f.read()).decode()
    href = f'<a href="data:application/octet-stream;base64,{b64}" download="UHC_Brokers_MNL_Partners.xlsx">Download Updated Excel File</a>'
    return href

st.markdown('### 📥 Download the updated file:')
st.markdown(get_table_download_link(file_path), unsafe_allow_html=True)
