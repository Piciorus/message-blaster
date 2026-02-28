import os
import time
import pandas as pd
import streamlit as st
import pywhatkit as pwk
import phonenumbers
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# ──────────────────────────────────────────────
# Page config
# ──────────────────────────────────────────────
st.set_page_config(
    page_title="Message Blaster",
    page_icon="💬",
    layout="wide",
)

st.title("💬 Message Blaster")
st.caption("Upload an Excel file, pick your contacts, write a message, and send via WhatsApp or Google Messages — 100% free.")

# ──────────────────────────────────────────────
# Sidebar
# ──────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Options")

    platform = st.radio(
        "Messaging platform",
        options=["WhatsApp", "Google Messages"],
        help="WhatsApp uses pywhatkit (WhatsApp Web). Google Messages uses Selenium (messages.google.com).",
    )

    st.divider()

    if platform == "WhatsApp":
        wait_time = st.slider(
            "Wait time per message (seconds)",
            min_value=10, max_value=40, value=15, step=1,
            help="Time to wait for WhatsApp Web to load. Increase on slow internet.",
        )
        close_tab = st.toggle(
            "Close WhatsApp tab after each message",
            value=False,
            help="Only enable if this app is open in a DIFFERENT browser than WhatsApp Web.",
        )

    else:  # Google Messages
        gm_wait = st.slider(
            "Wait time per message (seconds)",
            min_value=5, max_value=30, value=10, step=1,
            help="Time to wait for Google Messages to load each conversation.",
        )
        chrome_profile = st.text_input(
            "Chrome profile folder (optional)",
            value=os.path.join(os.path.expanduser("~"), "gm_blaster_profile"),
            help="A persistent Chrome profile so you only need to pair your phone once.",
        )

    st.divider()
    st.markdown("**Default country** — used when a number has no `+` prefix.")
    default_region = st.text_input(
        "Country code (ISO 3166-1, e.g. RO, US, GB)",
        value="RO",
        max_chars=2,
    ).upper()

# ──────────────────────────────────────────────
# How it works
# ──────────────────────────────────────────────
if platform == "WhatsApp":
    with st.expander("ℹ️ How WhatsApp sending works", expanded=False):
        st.markdown("""
1. Open **WhatsApp Web** in **Browser A** (e.g. Chrome) and log in via QR code.
2. Open **this app** in **Browser B** (e.g. Edge or Firefox).
3. Click Send — pywhatkit opens a WhatsApp Web tab per contact, types the message, and sends automatically.
4. Keep your mouse and keyboard free while sending.
        """)
else:
    with st.expander("ℹ️ How Google Messages sending works", expanded=False):
        st.markdown("""
1. Click **Pair phone** below — a Chrome window opens at messages.google.com.
2. Scan the QR code with your Android phone (Messages app → Device pairing).
3. Come back here and click **Send all messages**.
4. The app will open each conversation, type the message, and send automatically.
5. Your pairing is saved — you won't need to scan the QR code again next time.
        """)

# ──────────────────────────────────────────────
# Google Messages – Pair phone button
# ──────────────────────────────────────────────
if platform == "Google Messages":
    if st.button("📱 Pair phone (open Google Messages)"):
        opts = Options()
        opts.add_argument(f"--user-data-dir={chrome_profile}")
        opts.add_argument("--profile-directory=Default")
        driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=opts,
        )
        driver.get("https://messages.google.com/web")
        st.info("Chrome opened at messages.google.com. Scan the QR code with your phone, then close the window when paired.")

# ──────────────────────────────────────────────
# Helper – normalise phone number
# ──────────────────────────────────────────────
def normalise_phone(raw: str, region: str = "RO") -> str | None:
    try:
        parsed = phonenumbers.parse(str(raw).strip(), region)
        if phonenumbers.is_valid_number(parsed):
            return phonenumbers.format_number(parsed, phonenumbers.PhoneNumberFormat.E164)
    except phonenumbers.NumberParseException:
        pass
    return None


# ──────────────────────────────────────────────
# Google Messages sender (Selenium)
# ──────────────────────────────────────────────
def send_google_message(driver, phone: str, message: str, wait_sec: int) -> str:
    """Open a new conversation in Google Messages Web and send a message."""
    try:
        wait = WebDriverWait(driver, wait_sec)

        # Navigate to new conversation URL
        driver.get("https://messages.google.com/web/conversations/new")
        time.sleep(2)

        # Type phone number in the recipient field
        recipient = wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, "input[aria-label], input[placeholder*='name'], input[type='text']")
        ))
        recipient.click()
        recipient.clear()
        recipient.send_keys(phone)
        time.sleep(2)
        recipient.send_keys(Keys.RETURN)
        time.sleep(1.5)

        # Find message input and type
        msg_box = wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, "[contenteditable='true'], textarea[aria-label]")
        ))
        msg_box.click()
        msg_box.send_keys(message)
        time.sleep(0.5)
        msg_box.send_keys(Keys.RETURN)
        time.sleep(1.5)

        return "✅ sent"
    except Exception as e:
        return f"❌ {e}"


# ──────────────────────────────────────────────
# Step 1 – Upload Excel
# ──────────────────────────────────────────────
st.subheader("1️⃣  Upload your Excel file")
uploaded = st.file_uploader("Choose an .xlsx or .xls file", type=["xlsx", "xls"])

if not uploaded:
    st.info("Upload a file to get started.")
    st.stop()

try:
    df = pd.read_excel(uploaded, dtype=str)
    df.dropna(how="all", inplace=True)
    df.reset_index(drop=True, inplace=True)
except Exception as e:
    st.error(f"Could not read the file: {e}")
    st.stop()

st.success(f"Loaded **{len(df)} rows** and **{len(df.columns)} columns**.")
with st.expander("Preview data", expanded=True):
    st.dataframe(df.head(20), width="stretch")

# ──────────────────────────────────────────────
# Step 2 – Column mapping
# ──────────────────────────────────────────────
st.subheader("2️⃣  Map columns")

col1, col2 = st.columns(2)
with col1:
    phone_col = st.selectbox("Phone-number column *", options=df.columns.tolist())
with col2:
    name_col = st.selectbox(
        "Name column (optional — for {name} placeholder)",
        options=["— none —"] + df.columns.tolist(),
    )

df["_phone_e164"] = df[phone_col].apply(lambda x: normalise_phone(x, default_region))
valid_mask = df["_phone_e164"].notna()
valid_count = valid_mask.sum()
invalid_count = (~valid_mask).sum()

col_v, col_i = st.columns(2)
col_v.metric("Valid numbers", valid_count)
col_i.metric("Invalid / skipped", invalid_count)

if invalid_count:
    with st.expander(f"Show {invalid_count} invalid / unrecognised rows"):
        st.dataframe(df[~valid_mask][[phone_col]], width="stretch")

if valid_count == 0:
    st.error("No valid phone numbers found. Check the column and country code.")
    st.stop()

# ──────────────────────────────────────────────
# Step 3 – Compose message
# ──────────────────────────────────────────────
st.subheader("3️⃣  Compose your message")
st.caption("Use {name} or any other {ColumnName} as placeholders.")

message_template = st.text_area(
    "Message",
    value="Hello {name}, this is a custom message for you!",
    height=130,
)
st.caption(f"{len(message_template)} characters")


def render_message(template: str, row: pd.Series) -> str:
    values = {col: (str(row[col]) if pd.notna(row[col]) else "") for col in row.index if not col.startswith("_")}
    if name_col != "— none —":
        values["name"] = str(row[name_col]) if pd.notna(row[name_col]) else ""
    try:
        return template.format_map(values)
    except Exception:
        return template


first_valid_row = df[valid_mask].iloc[0]
with st.expander("Preview — first message"):
    st.info(render_message(message_template, first_valid_row))

# ──────────────────────────────────────────────
# Step 4 – Send
# ──────────────────────────────────────────────
st.subheader("4️⃣  Send messages")

if platform == "WhatsApp":
    st.warning(
        "**Before sending:**\n\n"
        "1. Open **WhatsApp Web** in **Browser A** (e.g. Chrome) and log in.\n"
        "2. Keep **this app** open in **Browser B** (e.g. Edge/Firefox).\n"
        "3. Do not move your mouse or type while messages are being sent."
    )
else:
    st.warning(
        "**Before sending:**\n\n"
        "1. Click **Pair phone** above and scan the QR code (first time only).\n"
        "2. Do not close the Chrome window that will open during sending.\n"
        "3. Do not move your mouse or type while messages are being sent."
    )

confirm = st.checkbox(
    f"I confirm I want to send **{valid_count} messages** via **{platform}**."
)

send_btn = st.button("🚀 Send all messages", disabled=not confirm, type="primary")

if send_btn:
    valid_rows = df[valid_mask].copy()
    progress = st.progress(0, text="Starting…")
    results = []

    # ── WhatsApp sending ──
    if platform == "WhatsApp":
        for i, (_, row) in enumerate(valid_rows.iterrows()):
            phone = row["_phone_e164"]
            body = render_message(message_template, row)
            try:
                pwk.sendwhatmsg_instantly(
                    phone_no=phone,
                    message=body,
                    wait_time=wait_time,
                    tab_close=close_tab,
                    close_time=3,
                )
                status = "✅ sent"
            except Exception as e:
                status = f"❌ {e}"

            results.append({"phone": phone, "platform": "WhatsApp", "status": status,
                            "message": body[:80] + ("…" if len(body) > 80 else "")})
            pct = int((i + 1) / len(valid_rows) * 100)
            progress.progress(pct, text=f"{i + 1}/{len(valid_rows)} — {phone}: {status}")

    # ── Google Messages sending ──
    else:
        opts = Options()
        opts.add_argument(f"--user-data-dir={chrome_profile}")
        opts.add_argument("--profile-directory=Default")
        opts.add_argument("--no-sandbox")
        opts.add_argument("--disable-dev-shm-usage")

        driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=opts,
        )
        driver.maximize_window()

        for i, (_, row) in enumerate(valid_rows.iterrows()):
            phone = row["_phone_e164"]
            body = render_message(message_template, row)
            status = send_google_message(driver, phone, body, gm_wait)

            results.append({"phone": phone, "platform": "Google Messages", "status": status,
                            "message": body[:80] + ("…" if len(body) > 80 else "")})
            pct = int((i + 1) / len(valid_rows) * 100)
            progress.progress(pct, text=f"{i + 1}/{len(valid_rows)} — {phone}: {status}")

        driver.quit()

    progress.empty()

    results_df = pd.DataFrame(results)
    sent_ok = (results_df["status"] == "✅ sent").sum()
    sent_fail = len(results_df) - sent_ok

    col_ok, col_fail = st.columns(2)
    col_ok.metric("Sent", sent_ok)
    col_fail.metric("Failed", sent_fail)

    st.dataframe(results_df, width="stretch")

    csv = results_df.to_csv(index=False).encode("utf-8")
    st.download_button(
        "📥 Download results as CSV",
        data=csv,
        file_name="message_results.csv",
        mime="text/csv",
    )
