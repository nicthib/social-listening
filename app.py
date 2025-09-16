import pandas as pd
import streamlit as st

st.title("Excel File Processor")

# Help section
with st.expander("Help?"):
    st.markdown("""
    Navigate to https://chat.igi.uw.nikon.com/  
    Attach the Excel file and input this prompt:

    ```
    Read the following excel document attached, and reply with the following:

    "Institution City" estimates the city and state of the institution this person will be working at. Respond with "International" if outside the US.

    "Institution ZIP" estimates the ZIP code of the institution this person will be working at. Respond with "N/A" if outside the US or Canada.

    "Relevance" is a score from 1-10 that scores the column C on how likely the tweet is about someone doing work related to microscopy, and that someone in the tweet is likely to receive funding in the near future (i.e. won an award, new lab, new faculty etc) (i.e. if they are getting a new position as a professor in the dept of neuroscience that's a 10. If they are getting a position as a postdoc in a CS lab that's a 1)

    Format the response as follows :
    New York, NY: 10027: 1|Des Moines, Iowa: 50047: 8|...

    and so on...please put a pipe between each row, as I need to add it back into my excel file and will use these to convert into line breaks. Don't respond with anything else, just the results.
    ```

    Copy the output and paste below.
    """)

# Upload Excel file
uploaded_file = st.file_uploader("Upload the Excel file (.xlsx)", type=["xlsx"])

# Input string
data = st.text_area("Paste the formatted location data string:")

# Slider for minimum score threshold
min_score = st.slider("Minimum Score Threshold", min_value=1, max_value=10, value=3)

# Run button
if st.button("Go!"):
    if uploaded_file and data:
        try:
            raw_df = pd.read_excel(uploaded_file, engine='openpyxl')

            raw_df = raw_df.drop([
                'Network', 'Author', 'Message ID', 'Profile', 'Potential Impressions',
                'Comments', 'Shares', 'Likes', 'Sentiment', 'Location', 'Hashtags',
                'Images', 'Language'
            ], axis=1, errors='ignore')

            entries = data.split('|')
            parsed = [entry.split(': ') for entry in entries]
            df = pd.DataFrame(parsed, columns=['City', 'ZIP', 'Score'])
            df['Score'] = df['Score'].astype(int)

            df = pd.concat([raw_df, df], axis=1)

            soql_df = pd.read_csv('SOQL.csv')
            soql_df = soql_df.drop_duplicates(subset='BillingPostalCode')
            users_df = pd.read_csv('Users.csv')

            df = df.merge(soql_df[['BillingPostalCode', 'OwnerId']], left_on='ZIP', right_on='BillingPostalCode', how='left')
            df = df.merge(users_df[['Id', 'Name']], left_on='OwnerId', right_on='Id', how='left')

            df.rename(columns={'Name': 'Account Owner'}, inplace=True)
            df.drop(columns=['BillingPostalCode', 'OwnerId', 'Id'], inplace=True)
            df = df[df['Score'] >= min_score]
            df = df[df['ZIP'] != 'N/A']
            df.drop_duplicates(subset=['Message'], inplace=True)

            output_filename = uploaded_file.name.replace('.xlsx', '_processed.xlsx')
            df.to_excel(output_filename, index=False)

            st.success(f"File processed and saved as {output_filename}")
        except Exception as e:
            st.error(f"An error occurred: {e}")
    else:
        st.warning("Please upload a file and enter the location data string.")
