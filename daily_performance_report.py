# importing the required libraries
import requests
import os
import datetime
import pandas as pd

# import user defined libraries
import configs
import utils

# Update current working directory
abspath = os.path.abspath(__file__)
dname = os.path.dirname(abspath)
os.chdir(dname)

def load_sheet_data_frame(sheet_id, worksheet_name="Daily Report") -> pd.DataFrame:
    gc = utils.connect_gcp()
    df = utils.fetch_worksheet_df(gc, sheet_id, worksheet_name)
    
    # Remove all invalid characters in table column headers
    df.columns = [col.replace("\n", "") for col in df.columns]
    
    return df

def generate_daily_performance_report_message(df: pd.DataFrame, source_name: str = "default"):
    # Prepare some initial variables
    today = datetime.date.today()
    today_b2 = today - datetime.timedelta(days=2)
    today_b2_str = today_b2.isoformat()
    
    # Message heading content
    message_heading = f"[Central SOCs] Thống kê các chỉ số KPI dưới target ngày {today_b2_str}"
    
    # Process ontime invalid message
    ontime_invalid_heading = "- Các SOC có Ontime dưới 99%: "
    ontime_invalid_values = df[df['KPI type'] == 'Ontime']
    ontime_invalid_body = ", ".join(ontime_invalid_values['Station'])
    
    # Process productivity invalid message
    productivity_invalid_heading = "- Các SOC có Pro dưới 95%: "
    productivity_invalid_values = df[df['KPI type'] == 'Productivity']
    productivity_invalid_body = ", ".join(productivity_invalid_values['Station'])
    
    # Process overdue invalid message
    overdue_invalid_heading = "- Các SOC có Overdue dưới 95%: "
    overdue_invalid_values = df[df['KPI type'] == 'Overdue']
    overdue_invalid_body = ", ".join(overdue_invalid_values['Station'])
    
    # Process didn't feedback yet message
    no_feedback_invalid_heading = "- Các SOC chưa feedback: "
    no_feedback_query = "`Group Feedback`.str.strip() == '' or `Feedback detail`.str.strip() == '' or `Action`.str.strip() == ''"
    no_feedback_values = df.query(no_feedback_query)
    no_feedback_body = ", ".join(no_feedback_values['Station'])
    
    # Process call to action message
    call_to_action = f"Nhờ các Sup bổ sung đầy đủ feedback tại link:\nhttps://docs.google.com/spreadsheets/d/{configs.gcp_sheet_id_daily_report}"
    
    # Compile message content
    message_content = message_heading
    
    if len(ontime_invalid_body) > 0:
        message_content += "\n" + ontime_invalid_heading + ontime_invalid_body
        
    if len(productivity_invalid_body) > 0:
        message_content += "\n" + productivity_invalid_heading + productivity_invalid_body
        
    if len(overdue_invalid_body) > 0:
        message_content += "\n" + overdue_invalid_heading + overdue_invalid_body
        
    if len(no_feedback_body) > 0:
        message_content += "\n" + no_feedback_invalid_heading + no_feedback_body
        
    message_content += "\n" + call_to_action
    
    utils.write_text_to_file(f"temp/daily_performance_report_{source_name}.txt", message_content)
    
    return message_content

def generate_daily_performance_report_img(df: pd.DataFrame, source_name: str = "default"):
    # Define style for each class in the data frame
    table_styles = {
        "Station": [
            {
                "selector": "",
                "props": "white-space: nowrap",
            }
        ],
        "Rate(%)": [
            {
                "selector": "",
                "props": "color: red",
            }
        ],
        "% After Remove": [
            {
                "selector": "",
                "props": "color: red",
            }
        ],
        "Group Feedback": [
            {
                "selector": "",
                "props": "width: 12%",
            }
        ],
        "Feedback detail": [
            {
                "selector": "",
                "props": "width: 25%",
            }
        ],
        "Action": [
            {
                "selector": "",
                "props": "width: 12%",
            }
        ],
    }
    
    # Generate td classes
    td_classes = utils.generate_td_classes(df)
    
    extra_css = """
        table {
            font-family: Arial, sans-serif;
            border-collapse: collapse;
            width: 100%;
        }
        
        th, td {
            border: 1px solid #dddddd;
            text-align: left;
            padding: 8px;
        }
        
        table th {
            background-color: 2f5d62;
            color: white !important;
        }
    """
        
    # Generate html content
    html_content = utils.df_to_html(df, table_styles, td_classes, extra_css)

    # Export DataFrame to HTML with custom styling
    html_file = f"temp/daily_performance_report_{source_name}.html"
    
    utils.write_text_to_file(html_file, html_content)
    
    print(f"DataFrame has been exported to '{html_file}' with custom styling.")

    # Path to save the output image
    output_image = f"temp/daily_performance_report_{source_name}.jpg"
    
    utils.html_to_img(html_file, output_image)

    return output_image

def send_daily_performance_report(message="", report_img=""):
    if message.strip() != "":
        data = {
            "tag": "text",
            "text": {
                "content": message,
            }
        }

        # sending post request and saving response as response object
        r = requests.post(url=configs.st_webhook_group_kpi, json=data)

        # extracting response text
        pastebin_url = r.text
        print("The pastebin URL is:%s" % pastebin_url)
    
    # Check output and send report
    report_img_base64 = utils.file_to_base64(report_img)
    
    if report_img_base64 != "":
        data = {
            "tag": "image",
            "image_base64": {
                "content": report_img_base64
            }
        }

        # sending post request and saving response as response object
        r = requests.post(url=configs.st_webhook_group_kpi, json=data)

        # extracting response text
        pastebin_url = r.text
        print("The pastebin URL is:%s" % pastebin_url)

if __name__ == "__main__": # Check if this file is executed or imported
    # Sheet id list for looping and creating report
    sheet_ids = {
        "south": configs.gcp_sheet_id_south,
        "central": configs.gcp_sheet_id_central,
    }
    
    for sheet_name, sheet_id in sheet_ids.items():
        df = load_sheet_data_frame(sheet_id)
        report_message = generate_daily_performance_report_message(df, sheet_name)
        report_img = generate_daily_performance_report_img(df, sheet_name)
        print(report_message)
        # send_daily_performance_report(report_message, report_img)
       