import pandas as pd
import numpy as np
import os
import re

import PyPDF2
from urllib.parse import quote

import pprint

def get_file_basename(file_path):
    # Get the base name of the file (including extension)
    base_name_with_extension = os.path.basename(file_path)

    # Split the base name and extension
    base_name, _ = os.path.splitext(base_name_with_extension)

    return base_name

def clean_text(text):
    """
    Cleans text by replacing line breaks with spaces if the text is a string.
    
    Parameters:
    - text: The text to clean, which may not necessarily be a string.
    
    Returns:
    - Cleaned text with line breaks replaced by spaces if it was a string,
      or the original input if it was not a string.
    """
    if isinstance(text, str):
        return text.replace('\n', ' ').replace('\r', ' ')
        
    return text

def format_datetime_to_iso8601(datetime_obj):
    """
    Formats a datetime object to a string in ISO 8601 format.
    
    Parameters:
    - datetime_obj: datetime.datetime, the datetime object to format.
    
    Returns:
    - str, the formatted datetime string in ISO 8601 format.
    """
    if pd.isnull(datetime_obj):
        return None
    return datetime_obj.strftime('%Y-%m-%dT%H:%M:%S')

def load_excel_into_dataframe(file_path):
    """
    Loads an Excel file into a pandas DataFrame, applies text cleaning, formats datetime fields to ISO 8601,
    and filters rows based on NaN value thresholds.
    
    Parameters:
    - file_path: str, path to the Excel file.
    
    Returns:
    - DataFrame with cleaned and processed data.
    """
    try:
        df = pd.read_excel(file_path, header=0)  # Assuming header is in the first row

        # Define the expected column names after skipping the initial rows
        df.columns = [
            'RegistrationNumber', 'NameOfMEP', 'Capacity', 'NameOfDonor',
            'DescriptionOfGift', 'EstimatedValue', 'LinkToPhoto',
            'DateOfReception', 'DateOfNotification', 'Location', 'Miscellaneous'
        ]

        # Remove rows that contain the document title and contain merged columns
        # Count NaN values in each row
        nan_counts = df.isna().sum(axis=1)
        
        # Skip merged rows, which contain the value only in the first column
        total_columns = len(df.columns)
        nan_threshold = total_columns - 2

        # Identify rows with multiple NaN values based on the threshold
        rows_skip = nan_counts > nan_threshold
        
        # Remove rows with multiple NaN values
        df = df[~rows_skip]


        # Remove rows that exactly match the column headers
        # This step assumes that the DataFrame's columns are named according to the headers
        header_row = df.iloc[0]
        rows_skip = df.apply(lambda x: all(x == header_row), axis=1)
        
        # Remove rows with multiple NaN values
        df = df[~rows_skip]


        # Explicitly convert known datetime columns from object to datetime
        datetime_columns = ['DateOfReception', 'DateOfNotification']  # Update with your actual datetime column names
        for col in datetime_columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')  # Converts to datetime, making invalid parsing 'NaT'
        
        # Regular expression to match the ID pattern
        # Adapt this pattern to match your specific ID format within the RegistrationNumber column
        pattern = r'G(\d+)[\-_]\d\d'

        # Extract the ID into a new column
        # The extract method applies the regular expression and captures the first matching group
        df['Id'] = df['RegistrationNumber'].str.extract(pattern, expand=False)

        
        # Clean text fields and format datetime fields
        for col in df.columns:
            if df[col].dtype == 'object':
                df[col] = df[col].apply(clean_text)
            
            elif np.issubdtype(df[col].dtype, np.datetime64):
                df[col] = df[col].apply(format_datetime_to_iso8601)

        return df

    except Exception as e:
        print(f"Error loading {file_path}: {e}")
        return pd.DataFrame()

def generate_markdown_gifts(row, output_directory):
    """
    Generates a Markdown file for a given row of data, saving it in a subdirectory based on the year extracted from the RegistrationNumber.
    
    Parameters:
    - row: Series, a row of data from the DataFrame.
    - output_directory: str, the directory where the Markdown files will be saved.
    """
    # Extract the year from the RegistrationNumber (assuming the format is Gxx-YY)
    year = "20" + row['RegistrationNumber'].split('-')[-1]  # Prefix "20" to get the full year
    year_directory = os.path.join(output_directory, year)
    
    # Create the year subdirectory if it doesn't exist
    if not os.path.exists(year_directory):
        os.makedirs(year_directory)
    
    # Construct the full path for the new Markdown file within the year subdirectory
    filename = f"{row['RegistrationNumber']}.md"
    full_path = os.path.join(year_directory, filename)
    
    # Write the Markdown file content
    with open(full_path, 'w', encoding='utf-8') as file:

        file.write('---\n')
        for col_name, value in row.items():
            # Check if the value is a string and format accordingly
            if col_name == "NameOfMEP" or \
               col_name == "NameOfDonor":
                # Escape existing double quotes in the string
                value = value.replace('"', '\\"')
                file.write(f'{col_name}: "[[{value}]]"\n')

            elif isinstance(value, str):
                # Escape existing double quotes in the string
                value = value.replace('"', '\\"')
                file.write(f'{col_name}: "{value}"\n')

            else:
                # For non-string values, write as is
                file.write(f'{col_name}: {value}\n')
        file.write('---\n\n')

        file.write(f"# {row['DescriptionOfGift']}\n\n")
        file.write(f"Received by: {row['NameOfMEP']}\n")
        file.write(f"From: {row['NameOfDonor']}\n")
    
    print(f"File '{filename}' has been created in {year_directory}.")

def generate_markdown_for_column_values(unique_values, column_name, df, output_directory):
    """
    Generates Markdown files for each unique value in a specified column,
    including all rows from the DataFrame that match each unique value.
    
    Parameters:
    - unique_values: List of unique values in the column.
    - column_name: Name of the column being processed.
    - df: The main DataFrame with all data.
    - output_directory: The directory where the Markdown files will be saved.
    """
    for value in unique_values:
        # Define a valid filename based on the column value
        filename = f"{value}.md".replace('/', '-').replace('\\', '-')
        full_path = os.path.join(output_directory, filename)
        
        # Generate the Markdown content
        with open(full_path, 'w', encoding='utf-8') as file:
            file.write(f"# {value}\n\n")
            file.write('\n')
        
        print(f"Markdown file created for \"{filename}\" in {column_name}")

def process_excel_files(directory_path, output_directory, pdf_urls=None):
    """
    Processes all Excel files in the given directory, generating Markdown files.
    
    Parameters:
    - directory_path: str, the path to the directory containing Excel files.
    - output_directory: str, the directory where Markdown files will be saved.
    """
    for filename in os.listdir(directory_path):

        basename = get_file_basename(filename)
        if filename.endswith('.xlsx'):

            file_path = os.path.join(directory_path, filename)
            print(f"Processing {filename}...")

            df = load_excel_into_dataframe(file_path)
            if not df.empty:

                if basename in pdf_urls:
                    photo_links = pdf_urls[basename]
                    print(f"Processing {len(photo_links)}...")
                    print(f"Processing {len(df)}...")

                    df = df.drop(columns=['LinkToPhoto'])
                    df = df.merge(photo_links, left_on='Id', right_on='Id', suffixes=('_xlsx', '_pdf'))

                # Generate the gifts markdown files
                for index, row in df.iterrows():
                    generate_markdown_gifts(row, output_directory)

                # Generate the MEPs markdown files via NameOfMEP
                unique_meps = df['NameOfMEP'].dropna().unique()
                output_directory_mep = 'meps'  # Update this path
                generate_markdown_for_column_values(unique_meps, 'NameOfMEP', df, output_directory_mep)

                # Generate the Donors markdown files via NameOfDonor
                unique_donors = df['NameOfDonor'].dropna().unique()
                output_directory_donor = 'donors'  # Update this path
                generate_markdown_for_column_values(unique_donors, 'NameOfDonor', df, output_directory_donor)

            else:
                print(f"Skipping {filename} due to loading issues.")

# Function to apply multiple regex patterns and return the first match
def extract_id(url):
    # List of regex patterns to try
    patterns = [
        r'G(\d+)[_\-]\d\d',
        r'G(\d+)',
        r'(\d+)[_\-]\d\d',
        r'(\d+)\.jpg'
    ]
    # Try each pattern until a match is found
    for pattern in patterns:
        match = re.search(pattern, url, flags=re.IGNORECASE)
        if match:
            return match.group(1)  # Return the first matching group

    return None  # No match found

def extract_urls_from_pdf(pdf_path):
    urls = []

    pdfFile = PyPDF2.PdfReader(pdf_path)
    pages = len(pdfFile.pages)
    key = '/Annots'
    uri = '/URI'
    ank = '/A'

    for page_number in range(pages):
        pageSliced = pdfFile.pages[page_number]
        pageObject = pageSliced.get_object()
        if key in pageObject.keys():
            ann = pageObject[key]
            for a in ann:
                u = a.get_object()
                if uri in u[ank].keys():
                    link = u[ank][uri]
                    encoded_uri = quote(link, safe='#:/?&=')
                    urls.append(encoded_uri)

    return urls
    
def process_pdf_files(directory_path):
    """
    Processes all Excel files in the given directory, generating Markdown files.
    
    Parameters:
    - directory_path: str, the path to the directory containing Excel files.
    - output_directory: str, the directory where Markdown files will be saved.
    """
    pdf_urls = {}
    for filename in os.listdir(directory_path):

        basename = get_file_basename(filename)
        if filename.endswith('.pdf'):
            file_path = os.path.join(directory_path, filename)
            print(f"Processing {filename}...")

            urls = extract_urls_from_pdf(file_path)
    
            urls_series = pd.Series(urls)
            # Apply the function to each URL
            ids_series = urls_series.apply(extract_id)
            
            df = pd.DataFrame({'Id': ids_series, 'LinkToPhoto': urls_series})
            #print(df.to_string())
            # if ids1 and ids2 and ids3 are []
            #   take ids4

            pdf_urls[basename] = df

    return pdf_urls


if __name__ == "__main__":
    # Specify the directory containing Excel files
    input_path = 'gifts_register/'  # Update this path
    # Specify the output directory for Markdown files
    output_directory = 'gifts/'  # Update this path

    pdf_urls = process_pdf_files(input_path)
    #print(pdf_urls.keys())

    process_excel_files(input_path, output_directory, pdf_urls)
    print("All files have been processed.")
