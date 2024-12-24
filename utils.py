#%%
import re
import pandas as pd
import io
from io import BytesIO

def process_data(data):
    # Your data processing logic goes here
    # Remove unwanted content
    data = data.replace("Yichang Liu(Kevin)\n订单问题请留Q&A会有专人处理。将不接受lark沟通", "")
    data = data.replace("Kim Shitong.Jiang\nKevin's Assistant | 订单问题请留Q&A会有专人处理。将不接受lark沟通", "")
    data = data.replace("Suixin.Lu\nIris - Kevin's Assistant | 订单问题请留Q&A会有专人处理。将不接受Lark沟通", "")
    data = data.replace("Yiting.Wang", "")
    data = data.replace("Jianan.Sheng", "")
    # ... [rest of your code]
    # Extract and pair dates with corresponding data
    date_pattern = r'\(DELAWARE\)--(\d{2}/\d{2}/\d{4})'
    dates = re.findall(date_pattern, data)
    sections = re.split(date_pattern, data)[1:]
    
    # Consolidate sections by dates
    consolidated_data = {}
    for i in range(0, len(sections), 2):
        date = sections[i]
        section = sections[i + 1]
        if date in consolidated_data:
            consolidated_data[date] += section  # Append section to existing date
        else:
            consolidated_data[date] = section  # Start new section for the date


    # Extract information from consolidated data
    entries = []

    for date, section in consolidated_data.items():
        # Define your regex patterns for each data type
        unfinished_order = re.search(r'当日未完成订单共计[:：]\s*(\d+)', section)
        general_unfinished = re.search(r'常规单未完成：(\d+)', section)
        coding_unfinished = re.search(r'改码单未完成：(\d+)', section)
        stock_rejection = re.search(r'因库存不准驳回：(\d+)', section)
        coding_order = re.search(r'Coding Orders\s*#\s*(\d+)', section)
        codes_used = re.search(r'(?i)Codes Used\s*#\s*(\d+)', section)
        finished_codes = re.search(r'(\d+) finisar codes', section)
        total_order = re.search(r'TOTAL ORDERS# (\d+)', section)
        general_order = re.search(r'GENERAL ORDER# (\d+)', section)
        transshipment_order = re.search(r'TOTAL TRANSSHIPMENT ORDERS# (\d+)', section)

        entries.append([
            date,
            unfinished_order.group(1) if unfinished_order else None,
            general_unfinished.group(1) if general_unfinished else None,
            coding_unfinished.group(1) if coding_unfinished else None,
            stock_rejection.group(1) if stock_rejection else None,
            coding_order.group(1) if coding_order else None,
            codes_used.group(1) if codes_used else None,
            finished_codes.group(1) if finished_codes else None,
            total_order.group(1) if total_order else None,
            general_order.group(1) if general_order else None,
            transshipment_order.group(1) if transshipment_order else None
        ])

    df = pd.DataFrame(entries, columns=["Date", "Unfinished Orders", "Regular Unfinished", "Coding Unfinished", "Stock Rejections", "Coding Orders", "Codes Used", "Finished Codes", "Total Orders", "General Orders", "Transshipment Orders"])

    # Instead of saving the file to disk, save it to an in-memory BytesIO buffer
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name="Sheet1", index=False)
        writer.close() # Close the Excel writer

    output.seek(0)  # Go to the start of the BytesIO buffer
    
    return output

# %%
