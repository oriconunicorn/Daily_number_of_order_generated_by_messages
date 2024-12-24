#%%
import re
import pandas as pd

# Read txt file of messages
with open('messages.txt', 'r', encoding="utf8") as f:
    data = f.read()

# Remove unwanted content
data = data.replace("Kim Shitong.Jiang\nKevin's Assistant | 订单问题请留Q&A会有专人处理。将不接受lark沟通", "")

# Extract and pair dates with corresponding data
date_pattern = r'\(DELAWARE\)--(\d{2}/\d{2}/\d{4})'
dates = re.findall(date_pattern, data)
sections = re.split(date_pattern, data)[1:]
date_data_pairs = [(sections[i], sections[i + 1]) for i in range(0, len(sections), 2)]

entries = []

for date, section in date_data_pairs:
    unfinished_order = re.search(r'当日未完成订单共计：(\d+)', section)
    general_unfinished = re.search(r'常规单未完成：(\d+)', section)
    coding_unfinished = re.search(r'改码单未完成：(\d+)', section)
    stock_rejection = re.search(r'因库存不准驳回：(\d+)', section)
    coding_order = re.search(r'(\d+) Coding orders', section)
    codes_used = re.search(r'(\d+)  codes used', section)
    finisar_codes = re.search(r'(\d+)  finisar codes', section)
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
        finisar_codes.group(1) if finisar_codes else None,
        total_order.group(1) if total_order else None,
        general_order.group(1) if general_order else None,
        transshipment_order.group(1) if transshipment_order else None
    ])

df = pd.DataFrame(entries, columns=["Date", "Unfinished Orders", "Regular Unfinished", "Coding Unfinished", "Stock Rejections", "Coding Orders", "Codes Used", "Finisar Codes", "Total Orders", "General Orders", "Transshipment Orders"])
df.to_excel("general_orders_data_sep.xlsx", index=False)
#%%