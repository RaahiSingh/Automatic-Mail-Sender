from flask import Flask, request, render_template
import pandas as pd
import ssl
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import re
import os
from datetime import datetime, timedelta
import schedule
import threading
import time

app = Flask(__name__)  # Fixed from _name

# ========== CONFIGURATION ==========
EMAIL_SENDER = "raahisingh2005@gmail.com"
EMAIL_RECEIVER = "raahisingh2005@gmail.com"
EMAIL_PASSWORD = "qlxq ryqz cxxl ousy"
DIRECTORY_PATH = r"C:\\Users\\raahi\\Python"

# ========== UTILITY FUNCTIONS ==========
def load_latest_file(dir_path):
    """Get the most recent .xlsx file (within last 31 days)."""
    latest_file = None
    latest_mtime = 0
    for filename in os.listdir(dir_path):
        if filename.endswith(".xlsx"):
            filepath = os.path.join(dir_path, filename)
            if os.path.isfile(filepath):
                mtime = os.path.getmtime(filepath)
                file_date = datetime.fromtimestamp(mtime)
                if mtime > latest_mtime and datetime.now() - file_date < timedelta(days=31):
                    latest_mtime = mtime
                    latest_file = filepath
    return latest_file


def process_excel_and_send_email(filepath):
    df_raw = pd.read_excel(filepath, header=None)
    df = pd.read_excel(filepath, skiprows=3, header=None)

    data = []

    for _, row in df.iterrows():
        signum = row[1]
        name = row[4]

        if pd.isna(signum) or pd.isna(name):
            continue

        oc_days = []
        for col in range(6, len(row)):
            cell_value = str(row[col]).strip().upper()
            if cell_value == "OC":
                date = df_raw.iloc[2, col]
                oc_days.append(str(date))

        if oc_days:
            title_cell = str(df_raw.iloc[0].values)
            match = re.search(r"\b([A-Za-z]+)\s+(\d{4})", title_cell)
            if match:
                month, year = match.groups()
                month = month.title()
            else:
                month = "Invalid"

            oc_days_with_month = [f"{day}-{month}" for day in oc_days]

            data.append({
                "Name": name,
                "Signum": signum,
                "OnCall Support Days": f"{len(oc_days)} ({', '.join(oc_days_with_month)})",
                "Amount": round(len(oc_days) * 714.29, 2)
            })

    if not data:
        return False

    summary = pd.DataFrame(data)

    html_content = f"""
    <p>Hi Swati,</p>
    <p>Please find below the On Call summary:</p>
    <table style="font-family: 'Times New Roman', Times, serif; border-collapse: collapse; width: 80%; text-align: center;">
        <thead>
            <tr style="background-color: #f2a365; color: white;">
                <th style="border: 1px solid #999; padding: 8px;">Name of Team member</th>
                <th style="border: 1px solid #999; padding: 8px;">Signum</th>
                <th style="border: 1px solid #999; padding: 8px;">On Call support days</th>
                <th style="border: 1px solid #999; padding: 8px;">Amount</th>
            </tr>
        </thead>
        <tbody>
    """

    for _, row in summary.iterrows():
        html_content += f"""
            <tr style="background-color: white;">
                <td style="border: 1px solid #999; padding: 8px;">{row['Name']}</td>
                <td style="border: 1px solid #999; padding: 8px;">{row['Signum']}</td>
                <td style="border: 1px solid #999; padding: 8px;">{row['OnCall Support Days']}</td>
                <td style="border: 1px solid #999; padding: 8px;">{row['Amount']:.2f}</td>
            </tr>
        """

    html_content += """
        </tbody>
    </table>
    """

    msg = MIMEMultipart()
    msg["From"] = EMAIL_SENDER
    msg["To"] = EMAIL_RECEIVER

    title_cell = str(df_raw.iloc[0].values)
    match = re.search(r"\b([A-Za-z]+)\s+(\d{4})", title_cell)
    if match:
        month, year = match.groups()
        subject = f"On Call Data – {month.title()}-{year}"
    else:
        subject = "On Call Data – Unknown"

    msg["Subject"] = subject
    msg.attach(MIMEText(html_content, 'html'))

    try:
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
            server.login(EMAIL_SENDER, EMAIL_PASSWORD)
            server.send_message(msg)
        return True
    except Exception as e:
        print(f"Failed to send email: {e}")
        return False


# ========== SCHEDULED JOB ==========
def job():
    if datetime.now().day == 17:
        print(" Running monthly job...")
        latest_file = load_latest_file(DIRECTORY_PATH)
        if latest_file:
            success = process_excel_and_send_email(latest_file)
            if success:
                print(" Email sent successfully from scheduler.")
            else:
                print(" Email sending failed or no OnCall data found.")
        else:
            print("⚠ No recent file found.")


def run_scheduler():
    schedule.every().day.at("11:45").do(job)
    while True:
        schedule.run_pending()
        time.sleep(60)


# ========== ROUTES ==========
@app.route('/')
def index():
    return render_template('mail.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return "No file part", 400

    file = request.files['file']
    if file.filename == '':
        return "No selected file", 400

    try:
        filepath = os.path.join(DIRECTORY_PATH, file.filename)
        file.save(filepath)
        success = process_excel_and_send_email(filepath)
        if success:
            return " Email sent successfully!"
        else:
            return " Email sending failed or no OnCall data found.", 500
    except Exception as e:
        return f"Error processing file: {str(e)}", 500


# ========== MAIN APP ==========
if __name__ == '__main__':
    # Only run scheduler in the actual Flask process, not the reloader
    if os.environ.get('WERKZEUG_RUN_MAIN') == 'true':
        scheduler_thread = threading.Thread(target=run_scheduler)
        scheduler_thread.daemon = True
        scheduler_thread.start()

    app.run(debug=True)
