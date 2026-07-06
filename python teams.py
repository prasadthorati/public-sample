# teams_sender_ui_onebox.py
# One-box UI for Teams group-wise/general messages.
# Features:
# - Upload Excel
# - Paste Microsoft Graph token ONCE (Step 0) - shared automatically by both steps below, hidden by default, show/hide toggle, Clear Token button
# - Update Chat IDs in uploaded Excel
# - Write complete message in ONE box
# - Optional @Everyone mention highlighting
# - Optional App/CTASK table
# - Group-to-group delay
# - Dry run and actual send (with a confirmation popup before real sending)
# - Cancel buttons on every form to reset fields without submitting
#
# Install once:
#   pip install flask pandas openpyxl requests
# Run:
#   python teams_sender_ui_onebox.py
# Open:
#   http://127.0.0.1:5000

import html
import os
import shutil
import time
from datetime import datetime

import pandas as pd
import requests
from flask import Flask, request, render_template_string, send_file

GRAPH = "https://graph.microsoft.com/v1.0"
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

REQUIRED_COLUMNS = ["App ID", "App Name", "CTASK ID", "CTASK Group"]
PREFERRED_GROUP_COLUMNS = ["Teams Group Name", "Group Name", "CTASK Group"]

app = Flask(__name__)
LAST_EXCEL_FILE = None
LAST_DISCOVERED_FILE = None
GRAPH_TOKEN = None  # stored once, in memory only, shared by both steps

HTML = """
<!doctype html>
<html>
<head>
    <title>Teams Message Sender - One Box</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 30px; background: #f7f7f7; }
        .box { background: white; padding: 20px; border-radius: 8px; margin-bottom: 20px; box-shadow: 0 1px 4px #ccc; }
        h1, h2 { color: #333; }
        label { font-weight: bold; display:block; margin-top: 12px; }
        textarea, input[type=file], select { width: 100%; padding: 8px; margin-top: 5px; box-sizing: border-box; }
        textarea.token { height: 90px; }
        textarea.message { height: 220px; font-family: Consolas, monospace; font-size: 14px; }
        button { background: #6264A7; color: white; border: none; padding: 10px 18px; border-radius: 5px; margin-top: 15px; cursor: pointer; margin-right: 8px; }
        button:hover { background: #464775; }
        button.secondary { background: #888; }
        button.secondary:hover { background: #666; }
        button.danger { background: #b00020; }
        button.danger:hover { background: #7d0016; }
        pre { background: #111; color: #0f0; padding: 12px; white-space: pre-wrap; border-radius: 5px; }
        .hint { color: #555; font-size: 13px; }
        .warning { color: #b00020; font-weight: bold; }
        .token-status { padding: 8px 12px; border-radius: 5px; margin-top: 10px; font-size: 14px; }
        .token-ok { background: #e6f4ea; color: #1e7e34; border: 1px solid #b6e2c1; }
        .token-missing { background: #fdecea; color: #b00020; border: 1px solid #f5c6cb; }
        .token-row { display: flex; gap: 10px; align-items: center; }
        .token-row input[type=password], .token-row input[type=text] { flex: 1; }
    </style>
    <script>
        function toggleTokenVisibility() {
            var field = document.getElementById('token_input');
            field.type = (field.type === 'password') ? 'text' : 'password';
        }
        function confirmSend(form) {
            var mode = document.activeElement.value;
            if (mode === 'send') {
                return confirm('This will send REAL messages to all matched Teams groups. Are you sure you want to continue?');
            }
            return true;
        }
    </script>
</head>
<body>
    <h1>Teams Message Sender</h1>

    <div class="box">
        <h2>Step 0: Microsoft Graph Token</h2>
        <p class="hint">Paste your token once here. It is reused for Step 1 and Step 2 below, so you never paste it twice. Kept in memory only (never written to disk) and hidden by default.</p>
        <form method="post" action="/set-token">
            <label>Microsoft Graph token</label>
            <div class="token-row">
                <input type="password" id="token_input" name="token" placeholder="Paste token here" autocomplete="off">
                <button type="button" class="secondary" onclick="toggleTokenVisibility()">Show/Hide</button>
            </div>
            <button type="submit">Save Token</button>
            <button type="submit" formaction="/clear-token" class="danger">Clear Token</button>
            <button type="reset" class="secondary">Cancel</button>
        </form>
        {% if token_set %}
        <div class="token-status token-ok">Token is set (ending in ****{{ token_tail }}). Ready to use in Step 1 and Step 2.</div>
        {% else %}
        <div class="token-status token-missing">No token set yet. Paste your token above and click "Save Token" before running Step 1 or Step 2.</div>
        {% endif %}
    </div>

    <div class="box">
        <h2>Step 1: Upload Excel and Update Chat IDs</h2>
        <form method="post" enctype="multipart/form-data" action="/update-chat-ids">
            <label>Select Excel file</label>
            <input type="file" name="excel" accept=".xlsx" required>

            <button type="submit">Update Chat IDs in Excel</button>
            <button type="reset" class="secondary">Cancel</button>
        </form>
        <p class="hint">Excel should have App ID, App Name, CTASK ID, CTASK Group, Group Name/Teams Group Name, and Chat ID columns. Chat ID can be blank. Uses the token saved in Step 0.</p>
    </div>

    <div class="box">
        <h2>Step 2: Write Message in One Box</h2>
        <form method="post" action="/send" onsubmit="return confirmSend(this)">
            <label>Message to send</label>
            <textarea class="message" name="message_text" required>Hi @Everyone,

Please proceed with pre-piv, attach artefacts to CTASK and move to “In progress” once done.
- CHG0690343 – Sat 04 July</textarea>
            <div class="hint">Write full message here. Line breaks will be preserved in Teams. If @Everyone is present and checkbox is selected, app will try to send it as Teams mention.</div>

            <label><input type="checkbox" name="everyone" checked> Try actual Teams @Everyone mention</label>

            <label><input type="checkbox" name="include_table"> Include App / CTASK table below message</label>
            <div class="hint">Unchecked = general message only. Checked = message plus each group’s respective application table.</div>

            <label>Delay between each group message</label>
            <select name="delay_seconds">
                <option value="0">0 Seconds</option>
                <option value="3">3 Seconds</option>
                <option value="5" selected>5 Seconds</option>
                <option value="10">10 Seconds</option>
                <option value="15">15 Seconds</option>
                <option value="30">30 Seconds</option>
                <option value="60">60 Seconds</option>
            </select>

            <button type="submit" name="mode" value="dry_run">Dry Run Only</button>
            <button type="submit" name="mode" value="send">Send Actual Messages</button>
            <button type="reset" class="secondary">Cancel</button>
        </form>
        <p class="hint">Uses the token saved in Step 0. "Send Actual Messages" will ask you to confirm before it sends anything.</p>
    </div>

    {% if message %}
    <div class="box">
        <h2>Result</h2>
        <pre>{{ message }}</pre>
        {% if download_excel %}<p><a href="/download/excel">Download Updated Excel</a></p>{% endif %}
        {% if download_discovered %}<p><a href="/download/discovered">Download Discovered Chats</a></p>{% endif %}
    </div>
    {% endif %}
</body>
</html>
"""


def normalise(value):
    if pd.isna(value):
        return ""
    text = str(value).strip().lower()
    for ch in ["_", "-", "/", "\\", ",", ".", "(", ")", "[", "]", "&"]:
        text = text.replace(ch, " ")
    return " ".join(text.split())


def graph_get_all(url, token):
    headers = {"Authorization": f"Bearer {token}"}
    results = []
    while url:
        response = requests.get(url, headers=headers, timeout=60)
        if response.status_code >= 400:
            raise RuntimeError(f"Graph GET failed {response.status_code}: {response.text}")
        data = response.json()
        results.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return results


def get_my_group_chats(token):
    rows = []
    chats = graph_get_all(f"{GRAPH}/me/chats?$top=50", token)
    for chat in chats:
        if chat.get("chatType") == "group":
            rows.append({
                "Teams Group Name": chat.get("topic") or "",
                "Chat ID": chat.get("id") or "",
                "Chat Type": chat.get("chatType") or "",
                "Created DateTime": chat.get("createdDateTime") or ""
            })
    df = pd.DataFrame(rows) if rows else pd.DataFrame(columns=["Teams Group Name", "Chat ID", "Chat Type", "Created DateTime"])
    df["Match Key"] = df["Teams Group Name"].apply(normalise)
    return df


def read_all_valid_sheets(excel_file):
    workbook = pd.read_excel(excel_file, sheet_name=None, engine="openpyxl")
    valid = {}
    for sheet, df in workbook.items():
        df.columns = df.columns.astype(str).str.strip()
        if all(col in df.columns for col in REQUIRED_COLUMNS):
            valid[sheet] = df
    if not valid:
        raise ValueError(f"No valid sheet found. Required columns: {REQUIRED_COLUMNS}")
    return workbook, valid


def choose_group_column(df):
    for col in PREFERRED_GROUP_COLUMNS:
        if col in df.columns:
            return col
    raise ValueError("No group name column found. Add 'Teams Group Name' or 'Group Name'.")


def find_chat_id(group_name, chats_df):
    key = normalise(group_name)
    if not key:
        return "", "Blank group name"

    exact = chats_df[chats_df["Match Key"] == key]
    if len(exact) == 1:
        return exact.iloc[0]["Chat ID"], "Exact match"

    candidates = chats_df[chats_df["Match Key"].apply(lambda x: bool(x) and (key in x or x in key))]
    if len(candidates) == 1:
        return candidates.iloc[0]["Chat ID"], f"Partial match: {candidates.iloc[0]['Teams Group Name']}"
    if len(candidates) > 1:
        return "", "Multiple matches - fill manually"

    return "", "No match found"


def update_chat_ids(excel_file, token):
    workbook, valid_sheets = read_all_valid_sheets(excel_file)
    chats_df = get_my_group_chats(token)

    discovered_file = excel_file.replace(".xlsx", "_discovered_chats.xlsx")
    chats_df.drop(columns=["Match Key"], errors="ignore").to_excel(discovered_file, index=False, engine="openpyxl")

    total_updated = 0
    total_missing = 0
    total_existing = 0

    for sheet, df in valid_sheets.items():
        group_col = choose_group_column(df)

        for col in ["Chat ID", "Chat ID Match Status"]:
            if col not in df.columns:
                df[col] = ""
            df[col] = df[col].astype("object").where(pd.notna(df[col]), "").astype(str).replace({"nan": "", "NaN": ""})

        for idx, row in df.iterrows():
            existing = str(row.get("Chat ID", "")).strip()
            if existing:
                total_existing += 1
                df.loc[idx, "Chat ID Match Status"] = "Already available"
                continue

            chat_id, status = find_chat_id(row.get(group_col, ""), chats_df)
            df.loc[idx, "Chat ID"] = chat_id
            df.loc[idx, "Chat ID Match Status"] = status
            if chat_id:
                total_updated += 1
            else:
                total_missing += 1

        workbook[sheet] = df

    backup_file = excel_file.replace(".xlsx", f"_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    shutil.copy2(excel_file, backup_file)

    with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:
        for sheet, df in workbook.items():
            df.to_excel(writer, sheet_name=sheet[:31], index=False)

    return discovered_file, (
        f"Backup created: {backup_file}\n"
        f"Same Excel updated: {excel_file}\n"
        f"Already had Chat ID: {total_existing}\n"
        f"New Chat IDs updated: {total_updated}\n"
        f"Rows still missing Chat ID: {total_missing}"
    )


def collect_rows(excel_file):
    _, valid_sheets = read_all_valid_sheets(excel_file)
    frames = []
    for sheet, df in valid_sheets.items():
        if "Chat ID" not in df.columns:
            continue
        df = df.copy()
        df["Chat ID"] = df["Chat ID"].astype("object").where(pd.notna(df["Chat ID"]), "").astype(str).str.strip()
        df = df.dropna(subset=["App ID", "App Name", "CTASK ID", "CTASK Group"])
        df = df[df["Chat ID"].ne("")]
        frames.append(df)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def convert_message_to_html(message_text, use_everyone):
    safe_text = html.escape(message_text).replace("\n", "<br>")
    if use_everyone and "@Everyone" in message_text:
        safe_text = safe_text.replace("@Everyone", '<at id="0">Everyone</at>', 1)
    elif "@Everyone" in message_text:
        safe_text = safe_text.replace("@Everyone", "Everyone", 1)
    return safe_text


def build_app_table(rows):
    table_rows = ""
    for _, row in rows.iterrows():
        table_rows += (
            "<tr>"
            f"<td>{html.escape(str(row['App ID']))}</td>"
            f"<td>{html.escape(str(row['App Name']))}</td>"
            f"<td>{html.escape(str(row['CTASK ID']))}</td>"
            f"<td>{html.escape(str(row['CTASK Group']))}</td>"
            "</tr>"
        )
    return (
        "<br><br>"
        "<table border='1' style='border-collapse:collapse'>"
        "<tr><th>App ID</th><th>App Name</th><th>CTASK ID</th><th>Assignment Group</th></tr>"
        + table_rows +
        "</table>"
    )


def build_message(rows, message_text, use_everyone, include_table):
    message_html = convert_message_to_html(message_text, use_everyone)
    if include_table:
        message_html += build_app_table(rows)
    return message_html


def post_message(chat_id, message_html, token, use_everyone):
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    payload = {"body": {"contentType": "html", "content": message_html}}

    if use_everyone and '<at id="0">Everyone</at>' in message_html:
        payload["mentions"] = [{
            "id": 0,
            "mentionText": "Everyone",
            "mentioned": {
                "conversation": {
                    "id": chat_id,
                    "displayName": "Everyone",
                    "conversationIdentityType": "chat"
                }
            }
        }]

    response = requests.post(f"{GRAPH}/chats/{chat_id}/messages", headers=headers, json=payload, timeout=60)

    # If actual @Everyone mention fails due to tenant/API policy, retry once with plain text Everyone.
    if response.status_code >= 400 and "mentions" in payload:
        plain_html = message_html.replace('<at id="0">Everyone</at>', 'Everyone')
        response = requests.post(
            f"{GRAPH}/chats/{chat_id}/messages",
            headers=headers,
            json={"body": {"contentType": "html", "content": plain_html}},
            timeout=60
        )

    if response.status_code not in [200, 201, 202]:
        raise RuntimeError(f"{response.status_code}: {response.text}")


def render_page(message=None, download_excel=False, download_discovered=False):
    token_set = bool(GRAPH_TOKEN)
    token_tail = GRAPH_TOKEN[-4:] if token_set else ""
    return render_template_string(
        HTML,
        message=message,
        download_excel=download_excel,
        download_discovered=download_discovered,
        token_set=token_set,
        token_tail=token_tail,
    )


@app.route("/")
def home():
    return render_page()


@app.route("/set-token", methods=["POST"])
def set_token_route():
    global GRAPH_TOKEN
    token = request.form.get("token", "").strip()
    if not token:
        return render_page(message="ERROR:\nToken box was empty. Paste your token before clicking Save Token.")
    GRAPH_TOKEN = token
    return render_page(message="Token saved. It will now be used automatically for Step 1 and Step 2.")


@app.route("/clear-token", methods=["POST"])
def clear_token_route():
    global GRAPH_TOKEN
    GRAPH_TOKEN = None
    return render_page(message="Token cleared. Paste a new token in Step 0 before running Step 1 or Step 2.")


@app.route("/update-chat-ids", methods=["POST"])
def update_route():
    global LAST_EXCEL_FILE, LAST_DISCOVERED_FILE
    try:
        if not GRAPH_TOKEN:
            raise ValueError("No token set. Paste your token in Step 0 and click Save Token first.")

        file = request.files["excel"]
        excel_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(excel_path)
        LAST_EXCEL_FILE = excel_path

        discovered, msg = update_chat_ids(excel_path, GRAPH_TOKEN)
        LAST_DISCOVERED_FILE = discovered
        msg += "\n\nNext: write full message in one box, choose table/delay, then click Dry Run."
        return render_page(message=msg, download_excel=True, download_discovered=True)
    except Exception as e:
        return render_page(message=f"ERROR:\n{e}", download_excel=False, download_discovered=False)


@app.route("/send", methods=["POST"])
def send_route():
    global LAST_EXCEL_FILE
    try:
        if not GRAPH_TOKEN:
            raise ValueError("No token set. Paste your token in Step 0 and click Save Token first.")

        if not LAST_EXCEL_FILE or not os.path.exists(LAST_EXCEL_FILE):
            raise ValueError("Please upload Excel and update Chat IDs first.")

        token = GRAPH_TOKEN
        message_text = request.form.get("message_text", "").strip()
        include_table = "include_table" in request.form
        use_everyone = "everyone" in request.form and "@Everyone" in message_text
        dry_run = request.form.get("mode") == "dry_run"

        try:
            delay_seconds = int(request.form.get("delay_seconds", "5"))
        except ValueError:
            delay_seconds = 5
        if delay_seconds < 0:
            delay_seconds = 0

        if not message_text:
            raise ValueError("Message box is empty")

        df = collect_rows(LAST_EXCEL_FILE)
        if df.empty:
            return render_page(
                message="No rows with Chat ID found. Download updated Excel and check Chat ID column.",
                download_excel=True,
                download_discovered=True
            )

        output = [
            f"Message mode: {'Message + App Table' if include_table else 'General Message Only'}",
            f"Selected delay between group messages: {delay_seconds} seconds"
        ]
        success = 0
        failed = 0
        groups = list(df.groupby("Chat ID"))
        total = len(groups)

        for i, (chat_id, rows) in enumerate(groups, start=1):
            rows = rows[["App ID", "App Name", "CTASK ID", "CTASK Group"]].drop_duplicates()
            message_html = build_message(rows, message_text, use_everyone, include_table)

            output.append(f"\nGroup {i} of {total}")
            output.append(f"Chat ID: {chat_id}")
            output.append(f"Apps in Excel for this group: {len(rows)}")
            if include_table:
                output.append(rows.to_string(index=False))
            output.append("Message Preview HTML:\n" + message_html)

            if dry_run:
                output.append("DRY RUN ONLY - not sent")
            else:
                try:
                    post_message(chat_id, message_html, token, use_everyone)
                    success += 1
                    output.append("SUCCESS: Message sent")
                except Exception as e:
                    failed += 1
                    output.append(f"FAILED: {e}")

                if i < total and delay_seconds > 0:
                    output.append(f"Waiting {delay_seconds} seconds before next group...")
                    time.sleep(delay_seconds)

        output.append(f"\nCompleted. Success: {success}, Failed: {failed}")
        return render_page(message="\n".join(output), download_excel=True, download_discovered=True)

    except Exception as e:
        return render_page(message=f"ERROR:\n{e}", download_excel=True, download_discovered=True)


@app.route("/download/excel")
def download_excel():
    return send_file(LAST_EXCEL_FILE, as_attachment=True)


@app.route("/download/discovered")
def download_discovered():
    return send_file(LAST_DISCOVERED_FILE, as_attachment=True)


if __name__ == "__main__":
    print("Starting UI... open http://127.0.0.1:5000")
    app.run(host="127.0.0.1", port=5000, debug=False)
