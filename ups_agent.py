import pandas as pd
import cloudscraper
import os
import io
import smtplib
from datetime import datetime
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.drawing.image import Image as OpenpyxlImage
import matplotlib.pyplot as plt
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import warnings

warnings.filterwarnings('ignore')

# --- CONFIGURATION ---
UPS_URL = "https://www.ups.com/in/en/support/shipping-support/shipping-costs-rates/fuel-surcharges"
EXCEL_FILE = "ups_fuel_history.xlsx"
TEMP_GRAPH = "ups_plot.png"

# --- EMAIL SECRETS (Passed securely from Cloud) ---
SMTP_SERVER = "smtp.gmail.com"  # Change if using O365
SMTP_PORT = 587
EMAIL_SENDER = os.environ.get("EMAIL_SENDER")
EMAIL_PASSWORD = os.environ.get("EMAIL_PASSWORD")
EMAIL_RECEIVERS = [EMAIL_SENDER] # Sending to yourself

def get_date_with_suffix():
    now = datetime.now()
    day = now.day
    suffix = 'th' if 11 <= day <= 13 else {1: 'st', 2: 'nd', 3: 'rd'}.get(day % 10, 'th')
    return now.strftime(f"{day}{suffix} %B %Y")

def extrapolate_and_expand(df):
    col_start, col_till = 'At Least (USD)', 'But Less Than (USD)'
    step_down = df.iloc[0]['Steps']
    rate_change_down = round(df.iloc[1]['Surcharge'] - df.iloc[0]['Surcharge'], 2)
    down_rows = []
    curr_s = round(df.iloc[0][col_start] - step_down, 2)
    curr_sur = round(df.iloc[0]['Surcharge'] - rate_change_down, 2)
    while curr_s >= 0.99:
        down_rows.insert(0, {col_start: curr_s, col_till: round(curr_s + step_down, 2), 'Surcharge': max(0, curr_sur), 'Steps': step_down})
        curr_s = round(curr_s - step_down, 2); curr_sur = round(curr_sur - rate_change_down, 2)

    step_up = df.iloc[-1]['Steps']
    rate_change_up = round(df.iloc[-1]['Surcharge'] - df.iloc[-2]['Surcharge'], 2)
    up_rows = []
    curr_s = round(df.iloc[-1][col_till], 2)
    curr_sur = round(df.iloc[-1]['Surcharge'] + rate_change_up, 2)
    while curr_s < 5.00:
        up_rows.append({col_start: curr_s, col_till: round(curr_s + step_up, 2), 'Surcharge': curr_sur, 'Steps': step_up})
        curr_s = round(curr_s + step_up, 2); curr_sur = round(curr_sur + rate_change_up, 2)

    full_df = pd.concat([pd.DataFrame(down_rows), df, pd.DataFrame(up_rows)], ignore_index=True)
    full_df = full_df[(full_df[col_start] >= 1.00) & (full_df[col_start] < 5.00)]

    exp_rows = []
    for _, row in full_df.iterrows():
        s, t = row[col_start], row[col_till]
        while round(s, 2) < round(t, 2) and round(s, 2) < 5.00:
            exp_rows.append({col_start: round(s, 2), col_till: t, 'Surcharge': row['Surcharge'], 'Steps': row['Steps']})
            s += 0.01
    return pd.DataFrame(exp_rows)

def get_live_data():
    scraper = cloudscraper.create_scraper(browser={'browser': 'chrome', 'platform': 'windows', 'mobile': False})
    try:
        response = scraper.get(UPS_URL, timeout=45)
        all_tables = pd.read_html(io.StringIO(response.text), match="Gulf Coast")
        if all_tables:
            df = all_tables[0]
            df.columns = ['At Least (USD)', 'But Less Than (USD)', 'Surcharge']
            for col in ['At Least (USD)', 'But Less Than (USD)']:
                df[col] = pd.to_numeric(df[col].astype(str).str.replace('USD', '').str.strip(), errors='coerce').round(2)
            df['Surcharge'] = pd.to_numeric(df['Surcharge'].astype(str).str.replace('%', '').str.replace(',', '.').str.strip(), errors='coerce').round(2)
            df['Steps'] = (df['But Less Than (USD)'] - df['At Least (USD)']).round(2)
            return extrapolate_and_expand(df.dropna(subset=['At Least (USD)']))
    except Exception as e:
        print(f"Scrape failed: {e}")
        return None

def send_email(subject, body):
    msg = MIMEMultipart()
    msg['From'] = EMAIL_SENDER
    msg['To'] = ", ".join(EMAIL_RECEIVERS)
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    with open(EXCEL_FILE, "rb") as f:
        part = MIMEApplication(f.read(), Name=os.path.basename(EXCEL_FILE))
    part['Content-Disposition'] = f'attachment; filename="{os.path.basename(EXCEL_FILE)}"'
    msg.attach(part)

    server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    server.starttls()
    server.login(EMAIL_SENDER, EMAIL_PASSWORD)
    server.send_message(msg)
    server.quit()
    print("Agent successfully dispatched email.")

def run_agent():
    today_df = get_live_data()
    if today_df is None: return

    today_date_str = get_date_with_suffix()
    changed_prices = set()
    email_body = f"Automated UPS Fuel Surcharge Report for {today_date_str}.\n\n"
    has_changes = False

    if os.path.exists(EXCEL_FILE):
        with pd.ExcelFile(EXCEL_FILE) as xls:
            last_sheet = xls.sheet_names[-1]
            yesterday_df = pd.read_excel(xls, sheet_name=last_sheet)
        
        merged = yesterday_df.merge(today_df, on=['At Least (USD)'], suffixes=('_old', '_new'))
        s_changes = merged[abs(merged['Surcharge_old'] - merged['Surcharge_new']) > 0.001]
        st_changes = merged[abs(merged['Steps_old'] - merged['Steps_new']) > 0.001]
        
        changed_prices.update(s_changes['At Least (USD)'].tolist())
        changed_prices.update(st_changes['At Least (USD)'].tolist())

        if not s_changes.empty or not st_changes.empty:
            has_changes = True
            if not s_changes.empty:
                email_body += f"[ALERT] Surcharge changed for {len(s_changes)} brackets.\n"
            if not st_changes.empty:
                email_body += f"[ALERT] Inflection Points changed for {len(st_changes)} brackets.\n"
            email_body += "Please see the attached Excel file. Changed rows are highlighted in YELLOW.\n"
        else:
            email_body += "Status: No changes detected since yesterday.\n"

        plt.figure(figsize=(10, 6)); plt.gca().set_facecolor('#FFFEE0')
        plt.plot(yesterday_df['At Least (USD)'], yesterday_df['Surcharge'], color='grey', linestyle='--', label='Previous')
        plt.plot(today_df['At Least (USD)'], today_df['Surcharge'], color='#351C15', linewidth=2.5, label='Current')
        plt.title('UPS Fuel Surcharge Index', fontweight='bold')
        plt.xlabel('Fuel Price (USD)'); plt.ylabel('Surcharge (%)'); plt.legend(); plt.savefig(TEMP_GRAPH); plt.close()
    else:
        email_body += "Initial baseline established.\n"

    # Save Excel
    yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
    mode = 'a' if os.path.exists(EXCEL_FILE) else 'w'
    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode=mode if mode == 'a' else 'w', if_sheet_exists='replace' if mode=='a' else None) as writer:
        today_df.to_excel(writer, sheet_name=today_date_str, index=False)
        worksheet = writer.sheets[today_date_str]
        for col_idx in range(1, 5):
            cell = worksheet.cell(row=1, column=col_idx)
            cell.fill = yellow_fill; cell.font = Font(bold=True); cell.alignment = Alignment(horizontal='center')
        for row_idx, start_val in enumerate(today_df['At Least (USD)'], start=2):
            worksheet.cell(row=row_idx, column=3).number_format = '0.00"%"'
            if start_val in changed_prices:
                for col_idx in range(1, 5): worksheet.cell(row=row_idx, column=col_idx).fill = yellow_fill
        if os.path.exists(TEMP_GRAPH):
            worksheet.add_image(OpenpyxlImage(TEMP_GRAPH), 'F2')

    if os.path.exists(TEMP_GRAPH): os.remove(TEMP_GRAPH)
    
    # Email logic
    subject = f"UPS Surcharge Alert - {today_date_str}" if has_changes else f"UPS Surcharge Log - {today_date_str}"
    send_email(subject, email_body)

if __name__ == "__main__":
    run_agent()
