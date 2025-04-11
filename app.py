from flask import Flask, render_template, request, send_file
import asyncio
from playwright.async_api import async_playwright
import pandas as pd
import os
import io
import telegram

app = Flask(__name__)

# Telegram 設定
TELEGRAM_TOKEN = '你的 Bot Token'
TELEGRAM_CHAT_ID = '你的 Chat ID'

async def extract_box_nos_from_page(url):
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page()
        await page.goto(url)
        content = await page.content()
        await browser.close()

        import re
        import json
        match = re.search(r'"setSkuList":(\[.*?\])', content)
        if not match:
            return []

        set_sku_list = json.loads(match.group(1))
        box_nos = [item["boxNo"] for item in set_sku_list]
        return box_nos

async def fetch_sku_list_by_box_no(box_no):
    import requests
    url = "https://m.popmart.com/sg/api/pop/boxProductDetail"
    params = {"boxNo": box_no}
    headers = {
        "User-Agent": "Mozilla/5.0"
    }
    response = requests.get(url, headers=headers, params=params)
    data = response.json()
    set_sku_list = data.get("data", {}).get("setSkuList", [])
    return set_sku_list

async def check_boxes(url, target_characters):
    box_nos = await extract_box_nos_from_page(url)
    results = []
    data_for_excel = []

    if not box_nos:
        results.append("未能取得 boxNo 列表，請檢查網址是否正確。")
        return results, None

    for idx, box_no in enumerate(box_nos, start=1):
        sku_list = await fetch_sku_list_by_box_no(box_no)
        characters = [sku.get("characterName").upper() for sku in sku_list]
        target_found = [char for char in target_characters if char in characters]

        if target_found:
            result_text = f"盒號 {idx}：包含角色 {', '.join(target_found)}！"
            # 發送 Telegram 通知
            await send_telegram_message(result_text)
        else:
            result_text = f"盒號 {idx}：不含目標角色"

        results.append(result_text)
        data_for_excel.append({
            "盒號": idx,
            "角色列表": ", ".join(characters),
            "匹配角色": ", ".join(target_found) if target_found else "無"
        })

    # 生成 Excel
    df = pd.DataFrame(data_for_excel)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)

    return results, output

async def send_telegram_message(message):
    bot = telegram.Bot(token=TELEGRAM_TOKEN)
    await bot.send_message(chat_id=TELEGRAM_CHAT_ID, text=message)

@app.route("/", methods=["GET", "POST"])
def index():
    results = []
    excel_file = None

    if request.method == "POST":
        activity_url = request.form["activity_url"]
        target_characters = [char.strip().upper() for char in request.form["target_character"].split(",")]

        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        results, excel_file = loop.run_until_complete(check_boxes(activity_url, target_characters))
        loop.close()

        if excel_file:
            excel_file_path = os.path.join("static", "查詢結果.xlsx")
            with open(excel_file_path, "wb") as f:
                f.write(excel_file.read())

    return render_template("index.html", results=results)

@app.route("/download")
def download_file():
    path = "static/查詢結果.xlsx"
    return send_file(path, as_attachment=True)

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5000)
