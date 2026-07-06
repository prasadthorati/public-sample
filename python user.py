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
from urllib.parse import quote

import pandas as pd
import requests
from flask import Flask, request, render_template_string, send_file

GRAPH = "https://graph.microsoft.com/v1.0"
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

REQUIRED_COLUMNS = ["App ID", "App Name", "CTASK ID", "CTASK Group"]
PREFERRED_GROUP_COLUMNS = ["Teams Group Name", "Group Name", "CTASK Group"]
POC_NAME_COLUMNS = ["POC Name", "Poc Name", "POC"]
POC_EMAIL_COLUMNS = ["POC Email", "POC Mail ID", "POC Mail Id", "POC Email ID", "POC Mail"]

POC_ID_CACHE = {}  # email (lowercased) -> {"id":..., "displayName":...} or None if lookup failed


def find_column(df, candidates):
    lower_map = {col.lower(): col for col in df.columns}
    for candidate in candidates:
        if candidate.lower() in lower_map:
            return lower_map[candidate.lower()]
    return None

app = Flask(__name__)
LAST_EXCEL_FILE = None
LAST_DISCOVERED_FILE = None
GRAPH_TOKEN = None  # stored once, in memory only, shared by both steps

HTML = """
<!doctype html>
<html>
<head>
    <title>Teams Message Sender</title>
    <style>
        * { box-sizing: border-box; }
        body {
            font-family: 'Segoe UI', Arial, sans-serif;
            margin: 0;
            background: #eef0f5;
            color: #222;
        }
        .topbar {
            background: linear-gradient(135deg, #6264A7 0%, #464775 100%);
            color: white;
            padding: 26px 40px;
        }
        .topbar h1 { margin: 0; font-size: 26px; }
        .topbar p { margin: 6px 0 0; opacity: 0.85; font-size: 14px; }
        .wrap { max-width: 780px; margin: 0 auto; padding: 30px 20px 60px; }
        .box {
            background: white;
            padding: 22px 26px;
            border-radius: 10px;
            margin-bottom: 22px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.07);
            border-left: 5px solid #6264A7;
        }
        .box h2 { display: flex; align-items: center; gap: 10px; margin-top: 0; font-size: 18px; color: #333; }
        .step-badge {
            display: inline-flex; align-items: center; justify-content: center;
            width: 26px; height: 26px; border-radius: 50%;
            background: #6264A7; color: white; font-size: 13px; font-weight: bold;
            flex-shrink: 0;
        }
        label { font-weight: 600; display: block; margin-top: 14px; font-size: 14px; color: #333; }
        textarea, input[type=file], input[type=text], input[type=password], select {
            width: 100%; padding: 9px 10px; margin-top: 6px; box-sizing: border-box;
            border: 1px solid #d3d3d9; border-radius: 6px; font-size: 14px; background: #fbfbfd;
        }
        textarea:focus, input:focus, select:focus { outline: none; border-color: #6264A7; background: white; }
        textarea.message { height: 220px; font-family: Consolas, monospace; font-size: 14px; }
        button {
            background: #6264A7; color: white; border: none; padding: 10px 18px;
            border-radius: 6px; margin-top: 16px; cursor: pointer; margin-right: 8px;
            font-size: 14px; font-weight: 600; transition: background 0.15s ease;
        }
        button:hover { background: #464775; }
        button.secondary { background: #eee; color: #444; }
        button.secondary:hover { background: #ddd; }
        button.danger { background: #fdecea; color: #b00020; }
        button.danger:hover { background: #f6c9c9; }
        pre {
            background: #1e1e1e; color: #7ee787; padding: 14px; white-space: pre-wrap;
            border-radius: 8px; font-size: 13px; max-height: 480px; overflow-y: auto;
        }
        .hint { color: #666; font-size: 12.5px; margin-top: 6px; line-height: 1.4; }
        .warning { color: #b00020; font-weight: bold; }
        .token-status { padding: 9px 14px; border-radius: 6px; margin-top: 14px; font-size: 13.5px; }
        .token-ok { background: #e6f4ea; color: #1e7e34; border: 1px solid #b6e2c1; }
        .token-missing { background: #fdecea; color: #b00020; border: 1px solid #f5c6cb; }
        .token-row { display: flex; gap: 10px; align-items: center; }
        .token-row input[type=password], .token-row input[type=text] { flex: 1; }
        .checkbox-row { display: flex; align-items: center; gap: 8px; margin-top: 14px; }
        .checkbox-row label { margin: 0; font-weight: 500; }
        .checkbox-row input { width: auto; margin: 0; }
        .result-header { display: flex; justify-content: space-between; align-items: center; }
        .result-header a { color: #6264A7; font-weight: 600; font-size: 13.5px; text-decoration: none; margin-left: 14px; }
        .result-header a:hover { text-decoration: underline; }
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
    <div class="topbar">
        <h1>Teams Message Sender</h1>
        <p>Update chat IDs from Excel, then broadcast one message to every matched Teams group.</p>
    </div>
    <div class="wrap">

    <div class="box">
        <h2><span class="step-badge">0</span> Microsoft Graph Token</h2>
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
        <h2><span class="step-badge">1</span> Upload Excel and Update Chat IDs</h2>
        <form method="post" enctype="multipart/form-data" action="/update-chat-ids">
            <label>Select Excel file</label>
            <input type="file" name="excel" accept=".xlsx" required>

            <button type="submit">Update Chat IDs in Excel</button>
            <button type="reset" class="secondary">Cancel</button>
        </form>
        <p class="hint">Required columns: App ID, App Name, CTASK ID, CTASK Group, Group Name/Teams Group Name. Chat ID can be blank.
        Optional columns: <b>POC Name</b> and <b>POC Email</b> — add these if you want a specific person mentioned in that group's message alongside @Everyone. Uses the token saved in Step 0.</p>
    </div>

    <div class="box">
        <h2><span class="step-badge">2</span> Write Message in One Box</h2>
        <form method="post" action="/send" onsubmit="return confirmSend(this)">
            <label>Message to send</label>
            <textarea class="message" name="message_text" required>Hi @Everyone,

Please proceed with pre-piv, attach artefacts to CTASK and move to “In progress” once done.
- CHG0690343 – Sat 04 July</textarea>
            <div class="hint">Write full message here. Line breaks will be preserved in Teams. If @Everyone is present and checkbox is selected, app will try to send it as Teams mention.</div>

            <div class="checkbox-row">
                <input type="checkbox" id="everyone" name="everyone" checked>
                <label for="everyone">Try actual Teams @Everyone mention</label>
            </div>

            <div class="checkbox-row">
                <input type="checkbox" id="notify_poc" name="notify_poc" checked>
                <label for="notify_poc">Also @mention the POC from Excel (POC Name / POC Email columns) so both POC and Everyone are notified</label>
            </div>

            <div class="checkbox-row">
                <input type="checkbox" id="include_table" name="include_table">
                <label for="include_table">Include App / CTASK table below message</label>
            </div>
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
        <p class="hint">Uses the token saved in Step 0. "Send Actual Messages" will ask you to confirm before it sends anything. "Dry Run" now also checks (read-only) whether each POC email can be turned into a real highlighted mention, and tells you why if it can't — usually a wrong/typo'd email, or the token's account lacking permission to look up other users in the directory.</p>
    </div>

    {% if message %}
    <div class="box">
        <div class="result-header">
            <h2 style="margin-bottom:0;">Result</h2>
            <div>
                {% if download_excel %}<a href="/download/excel">Download Updated Excel</a>{% endif %}
                {% if download_discovered %}<a href="/download/discovered">Download Discovered Chats</a>{% endif %}
            </div>
        </div>
        <pre>{{ message }}</pre>
    </div>
    {% endif %}
    </div>
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


def get_group_pocs(rows_full):
    """Return a list of unique {'name':..., 'email':...} dicts found in this group's rows."""
    name_col = find_column(rows_full, POC_NAME_COLUMNS)
    email_col = find_column(rows_full, POC_EMAIL_COLUMNS)
    if not name_col and not email_col:
        return []

    pocs = []
    seen = set()
    for _, row in rows_full.iterrows():
        name = str(row.get(name_col, "")).strip() if name_col else ""
        email = str(row.get(email_col, "")).strip() if email_col else ""
        if name.lower() == "nan":
            name = ""
        if email.lower() == "nan":
            email = ""
        if not name and not email:
            continue
        key = (name.lower(), email.lower())
        if key in seen:
            continue
        seen.add(key)
        pocs.append({"name": name or email, "email": email})
    return pocs


def resolve_user(email, token):
    """Look up a user's Graph object id/displayName by email, so they can be @mentioned.
    Returns (result_or_None, reason). result is {'id':..., 'displayName':...}.
    reason explains why lookup failed, for diagnostics shown to the user. Cached per email."""
    if not email:
        return None, "No email provided for this POC"
    key = email.lower()
    if key in POC_ID_CACHE:
        return POC_ID_CACHE[key]

    headers = {"Authorization": f"Bearer {token}"}
    select = "id,displayName,mail,userPrincipalName"

    try:
        # Attempt 1: direct lookup, works when the email IS the userPrincipalName
        response = requests.get(
            f"{GRAPH}/users/{quote(email, safe='')}",
            headers=headers,
            params={"$select": select},
            timeout=30
        )
        if response.status_code == 200:
            data = response.json()
            result = ({"id": data.get("id"), "displayName": data.get("displayName") or email}, None)
            POC_ID_CACHE[key] = result
            return result

        first_error = f"{response.status_code}: {response.text[:200]}"

        # Attempt 2: search by mail OR userPrincipalName in case they differ from the email given
        safe_email = email.replace("'", "''")
        response2 = requests.get(
            f"{GRAPH}/users",
            headers=headers,
            params={"$filter": f"mail eq '{safe_email}' or userPrincipalName eq '{safe_email}'", "$select": select},
            timeout=30
        )
        if response2.status_code == 200:
            values = response2.json().get("value", [])
            if values:
                data = values[0]
                result = ({"id": data.get("id"), "displayName": data.get("displayName") or email}, None)
                POC_ID_CACHE[key] = result
                return result
            reason = f"No Graph user found matching '{email}' (direct lookup: {first_error})"
        else:
            reason = f"Direct lookup failed ({first_error}); search also failed ({response2.status_code}: {response2.text[:200]})"

    except requests.RequestException as e:
        reason = f"Network/Graph error while resolving '{email}': {e}"

    result = (None, reason)
    POC_ID_CACHE[key] = result
    return result


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


def build_message(rows, message_text, use_everyone, include_table, poc_list=None, notify_poc=False):
    """Returns (message_html, mention_map).
    mention_map: {mention_id: {"type": "everyone"}} or {"type": "user", "name":..., "email":...}
    """
    message_html = convert_message_to_html(message_text, use_everyone)
    mention_map = {}
    if use_everyone and '<at id="0">Everyone</at>' in message_html:
        mention_map[0] = {"type": "everyone"}

    if notify_poc and poc_list:
        next_id = max(mention_map.keys(), default=-1) + 1
        at_tags = []
        for poc in poc_list:
            display_name = poc["name"] or poc["email"]
            if not display_name:
                continue
            at_tags.append(f'<at id="{next_id}">{html.escape(display_name)}</at>')
            mention_map[next_id] = {"type": "user", "name": display_name, "email": poc["email"]}
            next_id += 1
        if at_tags:
            message_html += "<br><br>cc: " + ", ".join(at_tags)

    if include_table:
        message_html += build_app_table(rows)

    return message_html, mention_map


def post_message(chat_id, message_html, token, mention_map):
    """Sends the message. Returns a list of diagnostic strings (e.g. why a POC mention
    fell back to plain text) so the caller can show the user what happened."""
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    mentions_payload = []
    diagnostics = []

    for mention_id, info in mention_map.items():
        if info["type"] == "everyone":
            mentions_payload.append({
                "id": mention_id,
                "mentionText": "Everyone",
                "mentioned": {
                    "conversation": {
                        "id": chat_id,
                        "displayName": "Everyone",
                        "conversationIdentityType": "chat"
                    }
                }
            })
        elif info["type"] == "user":
            resolved, reason = resolve_user(info["email"], token) if info["email"] else (None, "No email given for this POC")
            if resolved and resolved.get("id"):
                mentions_payload.append({
                    "id": mention_id,
                    "mentionText": info["name"],
                    "mentioned": {
                        "user": {
                            "id": resolved["id"],
                            "displayName": resolved["displayName"],
                            "userIdentityType": "aadUser"
                        }
                    }
                })
                diagnostics.append(f"POC '{info['name']}' resolved OK - will show as a real highlighted mention.")
            else:
                # Could not resolve this POC to a Teams account - fall back to plain text so
                # the message still sends (with the name visible, just not a clickable/highlighted mention).
                message_html = message_html.replace(
                    f'<at id="{mention_id}">{html.escape(info["name"])}</at>', html.escape(info["name"])
                )
                diagnostics.append(f"POC '{info['name']}' could NOT be resolved to a Teams mention, sent as plain text. Reason: {reason}")

    payload = {"body": {"contentType": "html", "content": message_html}}
    if mentions_payload:
        payload["mentions"] = mentions_payload

    response = requests.post(f"{GRAPH}/chats/{chat_id}/messages", headers=headers, json=payload, timeout=60)

    # If mentions fail due to tenant/API policy, retry once as plain text (still sends message + Everyone/POC names).
    if response.status_code >= 400 and mentions_payload:
        diagnostics.append(f"Message with mentions was rejected ({response.status_code}: {response.text[:200]}), retrying as plain text.")
        plain_html = message_html
        for mention_id, info in mention_map.items():
            plain_html = plain_html.replace(f'<at id="{mention_id}">', "").replace("</at>", "")
        response = requests.post(
            f"{GRAPH}/chats/{chat_id}/messages",
            headers=headers,
            json={"body": {"contentType": "html", "content": plain_html}},
            timeout=60
        )

    if response.status_code not in [200, 201, 202]:
        raise RuntimeError(f"{response.status_code}: {response.text}")

    return diagnostics


def check_mentions(mention_map, token):
    """Read-only check (no message sent) of whether each POC mention would resolve to a
    real highlighted Teams mention. Used by Dry Run so you can fix bad emails before sending."""
    diagnostics = []
    for info in mention_map.values():
        if info["type"] != "user":
            continue
        resolved, reason = resolve_user(info["email"], token) if info["email"] else (None, "No email given for this POC")
        if resolved and resolved.get("id"):
            diagnostics.append(f"POC '{info['name']}' ({info['email']}) WOULD resolve to a real highlighted mention.")
        else:
            diagnostics.append(f"POC '{info['name']}' ({info['email']}) would NOT resolve - would send as plain text. Reason: {reason}")
    return diagnostics


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
        notify_poc = "notify_poc" in request.form
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
            f"POC mention: {'Enabled' if notify_poc else 'Disabled'}",
            f"Selected delay between group messages: {delay_seconds} seconds"
        ]
        success = 0
        failed = 0
        groups = list(df.groupby("Chat ID"))
        total = len(groups)

        for i, (chat_id, rows_full) in enumerate(groups, start=1):
            poc_list = get_group_pocs(rows_full) if notify_poc else []
            rows = rows_full[["App ID", "App Name", "CTASK ID", "CTASK Group"]].drop_duplicates()
            message_html, mention_map = build_message(rows, message_text, use_everyone, include_table, poc_list, notify_poc)

            output.append(f"\nGroup {i} of {total}")
            output.append(f"Chat ID: {chat_id}")
            output.append(f"Apps in Excel for this group: {len(rows)}")
            if notify_poc:
                if poc_list:
                    output.append("POC(s) in this group: " + ", ".join(p["name"] for p in poc_list))
                else:
                    output.append("POC(s) in this group: none found (add POC Name/POC Email columns in Excel)")
            if include_table:
                output.append(rows.to_string(index=False))
            output.append("Message Preview HTML:\n" + message_html)

            if dry_run:
                output.append("DRY RUN ONLY - not sent")
                if any(info["type"] == "user" for info in mention_map.values()):
                    output.append("Checking whether POC(s) can be resolved to a real highlighted mention (read-only check, nothing is sent):")
                    for line in check_mentions(mention_map, token):
                        output.append("  - " + line)
            else:
                try:
                    diagnostics = post_message(chat_id, message_html, token, mention_map)
                    success += 1
                    output.append("SUCCESS: Message sent")
                    for line in diagnostics:
                        output.append("  - " + line)
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
