import pandas as pd
import gspread
import imgkit
import base64

import configs

def connect_gcp():
    gc = gspread.service_account(filename=configs.gcp_filename)
    print("Google Cloud Platform authenticated successfully")
    return gc

def fetch_worksheet_df(gc, sheet_id, worksheet_name):
    wks = gc.open_by_key(sheet_id)
    worksheet = wks.worksheet(worksheet_name)
    # Load all sheet records
    df = pd.DataFrame(worksheet.get_all_records())
    print("Daily Report loaded successfully")
    return df

def generate_td_classes(df):
    classes = [column_to_class(column) for column in df.columns]
    td_classes = pd.DataFrame([classes], index=df.index, columns=df.columns)
    return td_classes

def df_to_html(df, table_styles, td_classes, extra_css=""):    
    # Apply styling to the table
    styled = df.style.\
        hide(axis=0).\
        set_table_styles(table_styles).\
        set_td_classes(td_classes)

    # Define CSS styles
    styles = """
        <meta charset="UTF-8">
        <style>
            {{extra_css}}
        </style>
    """
    
    styles = styles.replace("{{extra_css}}", extra_css)

    # Concatenate styles with DataFrame HTML
    html_content = styles + styled.to_html()
        
    return html_content

def html_to_img(input_filename, ouput_filename, options=None):
    if options is None:
        # Configuration options for wkhtmltoimage
        options = {
            'format': 'jpg',  # Output image format
            'width': 1920,     # Width of the output image
        }
    # Convert HTML to image
    imgkit.from_file(input_filename, ouput_filename, options)
    print(f"HTML file '{input_filename}' has been converted to '{ouput_filename}'.")
    return True

def write_text_to_file(filename, content):
    try:
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(content)
            return True
    except Exception as e:
        print("Error", e)
        return False
    
def read_text_from_file(filename):
    try:
        with open(filename, 'r', encoding='utf-8') as f:
            return f.read()
    except Exception as e:
        print("Error", e)
        return ""
    
def file_to_base64(filename):
    if filename == "":
        return ""
    try:
        with open(filename, 'rb') as f:
            file_binary = f.read()
            image_data = (base64.b64encode(file_binary)).decode('ascii')
        return image_data
    except Exception as e:
        print("Error", e)
        return ""

# Function to create class name from column label
def column_to_class(column_label):
    # Convert column label to lowercase and decode as UTF-8
    class_name = column_label.lower().encode('utf-8').decode('utf-8')
    # Replace spaces with underscores
    class_name = class_name.replace(' ', '_')
    return class_name
