import os
import re
import msal
import time
import requests
import asyncio
from telegram import Bot
from flask import Flask
from threading import Thread
from dotenv import load_dotenv

# === Load Environment Variables ===
load_dotenv()

EXCEL_FILE_PATH = os.getenv("EXCEL_FILE_PATH")
WORKSHEET_NAME = os.getenv("WORKSHEET_NAME")
TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")

bot = Bot(token=TELEGRAM_TOKEN)

# === Flask Web Server untuk Keep-Alive ===
app = Flask('')

@app.route('/')
def home():
    return "✅ Bot is running!"

def run():
    app.run(host='0.0.0.0', port=8080)

def keep_alive():
    t = Thread(target=run)
    t.start()

# === Microsoft Graph Token ===
def get_access_token():
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET
    )
    token_response = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    return token_response.get('access_token')

# === Ambil Data Excel dari Graph API ===
def fetch_excel_data(access_token):
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    url = f"https://graph.microsoft.com/v1.0/users/itadmin.iass@ias.id/drive/root:/{EXCEL_FILE_PATH}:/workbook/worksheets('{WORKSHEET_NAME}')/usedRange"
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        return response.json()
    elif response.status_code == 401:
        raise Exception('Unauthorized, token expired or invalid.')
    else:
        raise Exception(f"Error fetching Excel: {response.status_code}, {response.text}")

# === Fungsi Ekstrak Durasi Flood Control ===
def extract_retry_seconds(error_msg):
    match = re.search(r'Retry in (\d+) seconds', error_msg)
    return int(match.group(1)) if match else 30

# === Kirim Pesan ke Telegram ===
async def send_message(row):
    try:
        if len(row) < 8:
            print(f"Skipping row, not enough columns: {row}")
            return

        pesan = (
            f"🎫 *No.Tiket:* {row[1]}\n"
            f"👤 *Nama:* {row[2]}\n"
            f"📱 *No HP:* {row[4]}\n"
            f"ℹ️ *Perihal:* {row[5]}\n"
            f"📋 *Deskripsi:* {row[9]}"
        )

        print("Sending message to Telegram...")
        response = await bot.send_message(chat_id=TELEGRAM_CHAT_ID, text=pesan, parse_mode='Markdown')
        print("✔️ Message sent successfully:", response.message_id)

    except Exception as e:
        error_msg = str(e)
        print(f"❌ Failed to send message: {error_msg}")
        if "Flood control exceeded" in error_msg:
            retry_seconds = extract_retry_seconds(error_msg)
            print(f"⏳ Waiting {retry_seconds} seconds due to flood control...")
            await asyncio.sleep(retry_seconds)
            return await send_message(row)  # Retry
        else:
            print("⚠️ Unknown error, skipping this message.")

# === Loop Utama ===
async def main_loop():
    last_row_file = 'last_row.txt'

    while True:
        try:
            token = get_access_token()
            print("✅ Token obtained successfully")

            data = fetch_excel_data(token)
            print("✅ Excel data fetched successfully")

            try:
                with open(last_row_file, 'r') as f:
                    last_row = int(f.read().strip())
            except (FileNotFoundError, ValueError):
                print("Creating new last_row file")
                last_row = 0
                with open(last_row_file, 'w') as f:
                    f.write('0')

            if data.get('values') and len(data['values']) > last_row + 1:
                for i in range(last_row + 1, len(data['values'])):
                    print(f"📤 Sending row {i}...")
                    await send_message(data['values'][i])
                    with open(last_row_file, 'w') as f:
                        f.write(str(i))  # Update posisi terakhir
                    await asyncio.sleep(1.5)  # Delay untuk hindari flood
            else:
                print("ℹ️ No new rows to send.")

        except Exception as e:
            print(f"⚠️ Error in main loop: {str(e)}")
            await asyncio.sleep(30)
            continue

        await asyncio.sleep(10)

# === Main Entry Point ===
if __name__ == '__main__':
    keep_alive()
    asyncio.run(main_loop())
