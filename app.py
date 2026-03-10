from __future__ import annotations

import time
from dataclasses import asdict

import pandas as pd
import streamlit as st

from src.config import build_smtp_config, load_smtp_config
from src.data_loader import Recipient, load_recipients_from_upload
from src.email_sender import Attachment, build_message, render_template, send_via_gmail_smtp
from src.logging_utils import LogEvent, error, info, warn


def init_state() -> None:
    st.session_state.setdefault("step", 0)  # 0 = landing, 1..5 workflow
    st.session_state.setdefault("smtp", None)  # SmtpConfig | None
    st.session_state.setdefault("recipients", None)  # list[Recipient] | None
    st.session_state.setdefault("logs", [])  # list[LogEvent]
    st.session_state.setdefault("results", [])  # list[dict]
    st.session_state.setdefault("cancel_bulk", False)
    st.session_state.setdefault("attachments", [])  # list[Attachment]
    st.session_state.setdefault("subject_template", "")
    st.session_state.setdefault("body_template", "")


def push_log(evt: LogEvent) -> None:
    st.session_state.logs.append(evt)


def recipients_to_df(recipients: list[Recipient]) -> pd.DataFrame:
    rows: list[dict] = []
    for r in recipients:
        # Keep original columns if present for preview.
        row = dict(r.fields)
        row.setdefault("name", r.name)
        row.setdefault("email", r.email)
        rows.append(row)
    # Title case common columns.
    df = pd.DataFrame(rows)
    return df


def attachments_from_uploads(uploads: list) -> list[Attachment]:
    atts: list[Attachment] = []
    for up in uploads or []:
        atts.append(
            Attachment(
                filename=up.name,
                content_type=getattr(up, "type", None),
                data=up.getvalue(),
            )
        )
    return atts


def render_logs_table() -> None:
    if not st.session_state.logs:
        st.caption("No logs yet.")
        return
    df = pd.DataFrame([asdict(x) for x in st.session_state.logs])
    st.dataframe(df, use_container_width=True, hide_index=True)


def render_results_table() -> None:
    if not st.session_state.results:
        st.caption("No sending results yet.")
        return
    df = pd.DataFrame(st.session_state.results)
    st.dataframe(df, use_container_width=True, hide_index=True)


def inject_styles() -> None:
    st.markdown(
        """
<style>
  /* Make Streamlit use true full width */
  .main .block-container {
    max-width: 100% !important;
    padding-left: 3.2rem !important;
    padding-right: 3.2rem !important;
    padding-top: 2.2rem !important;
    padding-bottom: 3.0rem !important;
  }

  /* Reduce default vertical gaps a bit */
  div[data-testid="stVerticalBlock"] > div:has(> div[data-testid="stVerticalBlock"]) {
    gap: 0.9rem;
  }

  .stApp {
    background: radial-gradient(1200px 600px at 20% 0%, rgba(99,102,241,0.18), rgba(0,0,0,0) 60%),
                radial-gradient(900px 500px at 85% 10%, rgba(6,182,212,0.14), rgba(0,0,0,0) 55%),
                linear-gradient(180deg, #0F172A 0%, #111827 100%);
    color: #F8FAFC;
  }
  .card {
    background: #1E293B;
    border: 1px solid rgba(148,163,184,0.18);
    border-radius: 16px;
    padding: 22px 22px 16px 22px;
    margin: 14px 0;
    font-size: 16px;
  }
  .subtle { color: #94A3B8; }
  .kpi {
    background: rgba(30,41,59,0.7);
    border: 1px solid rgba(148,163,184,0.18);
    border-radius: 14px;
    padding: 18px;
  }
  .kpi-title { color: #94A3B8; font-size: 15px; }
  .kpi-value { font-size: 38px; font-weight: 900; letter-spacing: -0.02em; }

  .email-preview {
    border: 1px solid rgba(148,163,184,0.22);
    border-radius: 14px;
    padding: 16px;
    background: rgba(15,23,42,0.55);
    font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace;
    white-space: pre-wrap;
    line-height: 1.55;
    font-size: 15px;
  }
  /* Make text areas feel more like a compose window */
  textarea {
    font-size: 16px !important;
    line-height: 1.6 !important;
  }

  /* Buttons: larger, rounded, SaaS-like */
  .stButton > button {
    border-radius: 14px !important;
    padding: 0.80rem 1.15rem !important;
    font-size: 16px !important;
    font-weight: 700 !important;
    border: 1px solid rgba(148,163,184,0.18) !important;
  }
  button[data-testid="baseButton-primary"] {
    background: #6366F1 !important;
    border: 1px solid rgba(99,102,241,0.85) !important;
  }
  button[data-testid="baseButton-primary"]:hover {
    background: #4F46E5 !important;
  }
  button[data-testid="baseButton-secondary"] {
    background: rgba(148,163,184,0.10) !important;
  }

  /* Inputs */
  div[data-baseweb="input"] input {
    font-size: 16px !important;
  }
  div[data-baseweb="textarea"] textarea {
    border-radius: 14px !important;
  }
</style>
        """,
        unsafe_allow_html=True,
    )


def card(title: str, step_label: str, body_fn) -> None:
    st.markdown(f"<div class='card'><div><b>[ {step_label} ]</b> {title}</div>", unsafe_allow_html=True)
    body_fn()
    st.markdown("</div>", unsafe_allow_html=True)


def available_variables(recipients: list[Recipient] | None) -> list[str]:
    if not recipients:
        return []
    keys: set[str] = set()
    for r in recipients[:50]:
        keys.update((k or "").strip().lower() for k in (r.fields or {}).keys())
    keys.discard("")
    # Ensure common ones appear.
    keys.update(["name", "email"])
    return sorted(keys)


def render_email_like_preview(*, from_name: str, from_email: str, to_email: str, subject: str, body: str) -> str:
    header = (
        "----------------------------------\n"
        f"From: {from_name or from_email}\n"
        f"To: {to_email}\n"
        f"Subject: {subject}\n"
        "----------------------------------\n\n"
    )
    return header + (body or "") + "\n\n----------------------------------\n"


st.set_page_config(page_title="Bulk Email Automation", layout="wide")
init_state()
inject_styles()

st.markdown(
    """
<div style="padding: 6px 0 2px 0;">
  <div style="font-size: 56px; font-weight: 950; line-height: 1.05;
              background: linear-gradient(90deg, #6366F1, #06B6D4);
              -webkit-background-clip: text; -webkit-text-fill-color: transparent;">
    Send Swift
  </div>
  <div class="subtle" style="font-size: 19px; margin-top: 10px;">
    Send FREE bulk personalized campaigns directly from Gmail.
    &nbsp;&nbsp;Upload recipients → Compose email → Send campaign
  </div>
</div>
    """,
    unsafe_allow_html=True,
)

smtp = st.session_state.get("smtp")

recipients = st.session_state.get("recipients") or []
uploaded_count = len(recipients) if recipients else 0
sent_count = sum(1 for r in (st.session_state.results or []) if r.get("status") == "sent")
failed_count = sum(1 for r in (st.session_state.results or []) if r.get("status") == "failed")

k1, k2, k3 = st.columns(3)
with k1:
    st.markdown(f"<div class='kpi'><div class='kpi-title'>Emails Uploaded</div><div class='kpi-value'>{uploaded_count}</div></div>", unsafe_allow_html=True)
with k2:
    st.markdown(f"<div class='kpi'><div class='kpi-title'>Emails Sent</div><div class='kpi-value'>{sent_count}</div></div>", unsafe_allow_html=True)
with k3:
    st.markdown(f"<div class='kpi'><div class='kpi-title'>Failed</div><div class='kpi-value'>{failed_count}</div></div>", unsafe_allow_html=True)

st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

if st.session_state.step == 0:
    st.markdown(
        """
<div class="card" style="text-align:center; padding: 34px 22px;">
  <div style="font-size:30px; font-weight:950;">Send Swift</div>
  <div class="subtle" style="margin-top:10px; font-size:18px;">Your Mail friend making your life easier with bulk perosnalized mails for free; No need to worry anymore </div>
</div>
        """,
        unsafe_allow_html=True,
    )
    if st.button("Start Sending Emails", type="primary"):
        st.session_state.step = 1
        st.rerun()

step = st.session_state.step
if step == 0:
    st.stop()

st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)

container = st.container()
with container:
    def step1_body():
        env_prefill = None
        try:
            env_prefill = load_smtp_config()
        except Exception:
            env_prefill = None

        st.caption("Use a Gmail App Password (not your normal password). Credentials are kept only in this session.")
        gmail = st.text_input("Gmail address", value=st.session_state.get("gmail_input", env_prefill.email if env_prefill else ""), placeholder="you@gmail.com")
        app_pw = st.text_input("App password", value=st.session_state.get("app_pw_input", ""), type="password", placeholder="16-character app password")
        sender_name_input = st.text_input("Sender name (optional)", value=st.session_state.get("sender_name_input", env_prefill.sender_name if env_prefill else ""))

        st.session_state.gmail_input = gmail
        st.session_state.app_pw_input = app_pw
        st.session_state.sender_name_input = sender_name_input

        c1, c2 = st.columns([0.6, 0.4])
        with c1:
            if st.button("Connect Gmail", type="primary", use_container_width=True):
                try:
                    st.session_state.smtp = build_smtp_config(email=gmail, app_password=app_pw, sender_name=sender_name_input)
                    push_log(info(f"Connected Gmail: {gmail}"))
                    st.session_state.step = max(st.session_state.step, 2)
                    st.rerun()
                except Exception as e:
                    push_log(error(f"Gmail connect failed: {e}"))
                    st.error(str(e))
        with c2:
            if st.button("Clear", use_container_width=True):
                st.session_state.smtp = None
                st.session_state.gmail_input = ""
                st.session_state.app_pw_input = ""
                push_log(warn("Cleared Gmail credentials"))
                st.rerun()

        if st.session_state.smtp is not None:
            st.success(f"Connected as {st.session_state.smtp.email}")

    card("Connect Gmail", "STEP 1", step1_body)

    def step2_body():
        if st.session_state.smtp is None:
            st.info("Complete Step 1 first.")
            return

        upload = st.file_uploader("Drag & drop CSV / Excel here", type=["csv", "xlsx", "xls"])
        if upload is not None:
            try:
                loaded = load_recipients_from_upload(upload.name, upload.getvalue())
                st.session_state.recipients = loaded
                st.session_state.results = []
                push_log(info(f"Loaded {len(loaded)} recipients from {upload.name}"))
                st.session_state.step = max(st.session_state.step, 3)
            except Exception as e:
                st.session_state.recipients = None
                push_log(error(f"Failed to load recipients: {e}"))
                st.error(str(e))

        if st.session_state.recipients:
            st.write("Preview:")
            st.dataframe(recipients_to_df(st.session_state.recipients).head(25), use_container_width=True, hide_index=True)

    card("Upload Recipients", "STEP 2", step2_body)

    def step3_body():
        if not st.session_state.recipients:
            st.info("Complete Step 2 first.")
            return

        vars_list = available_variables(st.session_state.recipients)
        st.caption("Personalize using tags like {{name}}, {{email}}, {{company}}, etc.")
        if vars_list:
            st.markdown("**Available variables:** " + ", ".join([f"`{{{{{v}}}}}`" for v in vars_list[:12]]) + (" …" if len(vars_list) > 12 else ""))

        subject_template = st.text_input(
            "Subject",
            value=st.session_state.get("subject_template") or "Hello {{name}}",
        )
        body_template = st.text_area(
            "Email Body",
            value=st.session_state.get("body_template") or "Hi {{name}},\n\nThis is a personalized bulk email.\n\nRegards,\n",
            height=320,
        )
        st.session_state.subject_template = subject_template
        st.session_state.body_template = body_template

        uploads = st.file_uploader("Attachments (optional)", accept_multiple_files=True)
        st.session_state.attachments = attachments_from_uploads(uploads or [])
        if st.session_state.attachments:
            st.caption("Attachments: " + ", ".join(a.filename for a in st.session_state.attachments))

        if st.button("Continue to Preview", use_container_width=True):
            st.session_state.step = max(st.session_state.step, 4)
            st.rerun()

    card("Compose Email", "STEP 3", step3_body)

    def step4_body():
        if st.session_state.smtp is None or not st.session_state.recipients:
            st.info("Complete Steps 1–3 first.")
            return

        sample: Recipient = st.session_state.recipients[0]
        from_email = st.session_state.smtp.email
        from_name = st.session_state.smtp.sender_name or ""
        to_email = sample.email or "test@example.com"
        variables = dict(sample.fields)
        variables.setdefault("name", sample.name)
        variables.setdefault("email", sample.email)

        subj = render_template(st.session_state.subject_template, variables)
        body = render_template(st.session_state.body_template, variables)
        preview_text = render_email_like_preview(
            from_name=from_name,
            from_email=from_email,
            to_email=to_email,
            subject=subj,
            body=body,
        )
        st.markdown(f"<div class='email-preview'>{preview_text}</div>", unsafe_allow_html=True)

        if st.button("Continue to Send Campaign", type="primary", use_container_width=True):
            st.session_state.step = max(st.session_state.step, 5)
            st.rerun()

    card("Preview Email", "STEP 4", step4_body)
    def step5_body():
        if st.session_state.smtp is None or not st.session_state.recipients:
            st.info("Complete Steps 1–4 first.")
            return

        delay_s = st.number_input("Delay between sends (seconds)", min_value=0.0, max_value=60.0, value=2.0, step=0.5)
        test_to = st.text_input("Test recipient email", placeholder="test@example.com")

        c1, c2 = st.columns(2)
        with c1:
            send_test = st.button("📩 Send Test Email", use_container_width=True)
        with c2:
            start_bulk = st.button("🚀 Send Campaign", type="primary", use_container_width=True)

        if send_test:
            if not test_to.strip():
                st.error("Enter a test recipient email.")
            else:
                try:
                    variables = {"name": "Test", "email": test_to.strip()}
                    subj = render_template(st.session_state.subject_template, variables)
                    body = render_template(st.session_state.body_template, variables)
                    msg = build_message(
                        smtp=st.session_state.smtp,
                        to_email=test_to.strip(),
                        subject=subj,
                        body=body,
                        attachments=st.session_state.attachments,
                        from_email=st.session_state.smtp.email,
                        from_name=st.session_state.smtp.sender_name,
                    )
                    with st.spinner("Sending test email..."):
                        send_via_gmail_smtp(smtp=st.session_state.smtp, message=msg)
                    push_log(info(f"Test email sent to {test_to.strip()}"))
                    st.success("Test email sent.")
                except Exception as e:
                    message = str(e)
                    push_log(error(f"Test email failed: {message}"))
                    st.error(f"Failed to send test email: {message}")

        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
        progress = st.progress(0, text="Ready to send…")
        counter = st.empty()
        status_box = st.empty()

        if start_bulk:
            st.session_state.cancel_bulk = False
            st.session_state.results = []

            recipients_local = st.session_state.recipients
            total = len(recipients_local)
            sent_ok = 0
            sent_fail = 0
            push_log(info(f"Starting campaign to {total} recipient(s)"))

            for idx, r in enumerate(recipients_local, start=1):
                if st.session_state.cancel_bulk:
                    push_log(warn("Campaign stopped by user"))
                    status_box.warning("Stopped.")
                    break

                progress.progress((idx - 1) / total, text="Sending Emails…")
                counter.markdown(f"**{idx - 1} / {total} emails sent**")
                status_box.info(f"Sending to {r.email} …")

                try:
                    variables = dict(r.fields)
                    variables.setdefault("name", r.name)
                    variables.setdefault("email", r.email)
                    subj = render_template(st.session_state.subject_template, variables)
                    body = render_template(st.session_state.body_template, variables)
                    msg = build_message(
                        smtp=st.session_state.smtp,
                        to_email=r.email,
                        subject=subj,
                        body=body,
                        attachments=st.session_state.attachments,
                        from_email=st.session_state.smtp.email,
                        from_name=st.session_state.smtp.sender_name,
                    )
                    send_via_gmail_smtp(smtp=st.session_state.smtp, message=msg)
                    sent_ok += 1
                    st.session_state.results.append({"email": r.email, "status": "sent", "error": ""})
                    push_log(info(f"Sent to {r.email}"))
                except Exception as e:
                    sent_fail += 1
                    st.session_state.results.append({"email": r.email, "status": "failed", "error": str(e)})
                    push_log(error(f"Failed to send to {r.email}: {e}"))

                progress.progress(idx / total, text="Sending Emails…")
                counter.markdown(f"**{idx} / {total} emails sent**")
                if idx < total and delay_s > 0 and not st.session_state.cancel_bulk:
                    time.sleep(float(delay_s))

            done = sent_ok + sent_fail
            if done == total:
                status_box.success(f"Done. Sent: {sent_ok}, Failed: {sent_fail}")
                push_log(info(f"Campaign complete. Sent={sent_ok} Failed={sent_fail}"))
            else:
                status_box.warning(f"Stopped. Processed: {done}/{total}")

        st.divider()
        st.subheader("Sent / Failed logs")
        render_results_table()

    card("Send Campaign", "STEP 5", step5_body)

st.subheader("Logs")
render_logs_table()

