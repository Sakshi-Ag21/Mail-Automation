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
    st.session_state.setdefault("recipients", None)  # list[Recipient] | None
    st.session_state.setdefault("logs", [])  # list[LogEvent]
    st.session_state.setdefault("results", [])  # list[dict]
    st.session_state.setdefault("cancel_bulk", False)


def push_log(evt: LogEvent) -> None:
    st.session_state.logs.append(evt)


def recipients_to_df(recipients: list[Recipient]) -> pd.DataFrame:
    return pd.DataFrame([{"Name": r.name, "Email": r.email} for r in recipients])


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


st.set_page_config(page_title="Bulk Email Automation", layout="wide")
init_state()

st.title("Bulk Email Automation Dashboard")
st.caption("Upload recipients, compose email, attach files, send test email, then send bulk one-by-one via Gmail SMTP.")

with st.sidebar:
    st.subheader("SMTP configuration")
    st.caption("Enter your Gmail + App Password. These are kept in memory for this session only.")

    # Optional: prefill from env/.env (useful for local dev / admin usage).
    env_smtp = None
    try:
        env_smtp = load_smtp_config()
    except Exception:
        env_smtp = None

    smtp_email_input = st.text_input(
        "Gmail address",
        value=st.session_state.get("smtp_email_input", env_smtp.email if env_smtp else ""),
        placeholder="you@gmail.com",
    )
    smtp_app_password_input = st.text_input(
        "Gmail App Password",
        value=st.session_state.get("smtp_app_password_input", ""),
        type="password",
        placeholder="16-character app password",
        help="Use a Google App Password (not your normal Gmail password). Spaces are OK.",
    )
    smtp_sender_name_input = st.text_input(
        "Sender name (optional)",
        value=st.session_state.get("smtp_sender_name_input", env_smtp.sender_name if env_smtp else ""),
        placeholder="Your Name",
    )

    st.session_state.smtp_email_input = smtp_email_input
    st.session_state.smtp_app_password_input = smtp_app_password_input
    st.session_state.smtp_sender_name_input = smtp_sender_name_input

    smtp = None
    try:
        if smtp_email_input.strip() and smtp_app_password_input.strip():
            smtp = build_smtp_config(
                email=smtp_email_input,
                app_password=smtp_app_password_input,
                sender_name=smtp_sender_name_input,
            )
            st.success(f"Ready to send as {smtp.email}")
        else:
            st.warning("Enter Gmail address + App Password to enable sending.")
    except Exception as e:
        st.error(str(e))

col_left, col_right = st.columns([1.15, 0.85], gap="large")

with col_left:
    st.subheader("1) Choose sender")
    if smtp is None:
        st.info("Enter Gmail credentials in the sidebar first.")
        sender_email = ""
        sender_name = ""
    else:
        sender_email = st.text_input(
            "Sender email address (From)",
            value=st.session_state.get("sender_email", smtp.email),
            help="This is the address that appears in the From field.",
        )
        sender_name = st.text_input(
            "Sender display name (optional)",
            value=st.session_state.get("sender_name", smtp.sender_name or ""),
        )
        st.session_state.sender_email = sender_email
        st.session_state.sender_name = sender_name
        if sender_email and sender_email != smtp.email:
            st.caption(
                "Note: Gmail may reject sending if the From address does not match the authenticated Gmail."
            )

    st.subheader("2) Upload recipients")
    upload = st.file_uploader("Upload CSV/Excel", type=["csv", "xlsx", "xls"])

    if upload is not None:
        try:
            recipients = load_recipients_from_upload(upload.name, upload.getvalue())
            st.session_state.recipients = recipients
            push_log(info(f"Loaded {len(recipients)} recipients from {upload.name}"))
        except Exception as e:
            st.session_state.recipients = None
            push_log(error(f"Failed to load recipients: {e}"))
            st.error(str(e))

    if st.session_state.recipients:
        st.write("Preview:")
        st.dataframe(recipients_to_df(st.session_state.recipients), use_container_width=True, hide_index=True)

    st.subheader("3) Compose email")
    subject_template = st.text_input(
        "Subject",
        value=st.session_state.get("subject_template", "Hello {name}"),
        help="Use {name} and {email} placeholders.",
    )
    body_template = st.text_area(
        "Body (plain text)",
        value=st.session_state.get(
            "body_template",
            "Hi {name},\n\nThis is a personalized bulk email.\n\nRegards,\n",
        ),
        height=220,
        help="Use {name} and {email} placeholders. Plain text is used for compatibility.",
    )

    st.session_state.subject_template = subject_template
    st.session_state.body_template = body_template

    st.subheader("4) Attachments")
    uploads = st.file_uploader(
        "Upload one or more attachments",
        accept_multiple_files=True,
    )
    attachments = attachments_from_uploads(uploads or [])
    if attachments:
        st.write("Selected attachments:")
        st.write([a.filename for a in attachments])

with col_right:
    st.subheader("5) Send test email")
    test_to = st.text_input("Test recipient email", placeholder="test@example.com")
    test_name = st.text_input("Test recipient name (for {name})", value="Test")

    test_preview = st.checkbox("Preview rendered test email", value=True)
    if test_preview:
        st.markdown("**Preview (rendered):**")
        st.code(
            "Subject:\n"
            + render_template(subject_template, name=test_name, email=test_to or "test@example.com")
            + "\n\nBody:\n"
            + render_template(body_template, name=test_name, email=test_to or "test@example.com"),
            language="text",
        )

    if st.button("Send test email", type="primary", use_container_width=True):
        if smtp is None:
            push_log(error("Test send blocked: SMTP not configured"))
            st.error("SMTP is not configured. Set environment variables in the sidebar instructions.")
        elif not sender_email.strip():
            st.error("Enter a sender email in step 1.")
        elif not test_to.strip():
            st.error("Enter a test recipient email.")
        else:
            try:
                subj = render_template(subject_template, name=test_name, email=test_to.strip())
                body = render_template(body_template, name=test_name, email=test_to.strip())
                msg = build_message(
                    smtp=smtp,
                    to_email=test_to.strip(),
                    subject=subj,
                    body=body,
                    attachments=attachments,
                    from_email=sender_email.strip(),
                    from_name=sender_name.strip() or None,
                )
                with st.spinner("Sending test email..."):
                    send_via_gmail_smtp(smtp=smtp, message=msg)
                push_log(info(f"Test email sent to {test_to.strip()}"))
                st.success("Test email sent.")
            except Exception as e:
                message = str(e)
                push_log(error(f"Test email failed: {message}"))
                if "Username and Password not accepted" in message or "535" in message:
                    st.error(
                        "Failed to send test email: Gmail rejected the username/password.\n\n"
                        "Make sure SMTP_EMAIL and SMTP_APP_PASSWORD in your .env match the same Google account "
                        "and that SMTP_APP_PASSWORD is a valid App Password."
                    )
                else:
                    st.error(f"Failed to send test email: {message}")

    st.divider()
    st.subheader("6) Send bulk emails")

    delay_s = st.number_input("Delay between sends (seconds)", min_value=0.0, max_value=60.0, value=2.0, step=0.5)

    c1, c2 = st.columns(2)
    with c1:
        start = st.button("Start bulk send", use_container_width=True)
    with c2:
        stop = st.button("Stop", use_container_width=True)

    if stop:
        st.session_state.cancel_bulk = True
        push_log(warn("Bulk send cancellation requested"))

    progress = st.progress(0, text="Waiting to start…")
    status_box = st.empty()

    if start:
        st.session_state.cancel_bulk = False
        st.session_state.results = []

        if smtp is None:
            push_log(error("Bulk send blocked: SMTP not configured"))
            st.error("SMTP is not configured. Set environment variables in the sidebar instructions.")
        elif not sender_email.strip():
            st.error("Enter a sender email in step 1.")
        elif not st.session_state.recipients:
            push_log(error("Bulk send blocked: no recipients uploaded"))
            st.error("Upload a CSV/Excel with Name and Email columns first.")
        elif not subject_template.strip():
            st.error("Subject cannot be empty.")
        elif not body_template.strip():
            st.error("Body cannot be empty.")
        else:
            recipients = st.session_state.recipients
            total = len(recipients)
            sent_ok = 0
            sent_fail = 0

            push_log(info(f"Starting bulk send to {total} recipient(s)"))

            for idx, r in enumerate(recipients, start=1):
                if st.session_state.cancel_bulk:
                    push_log(warn("Bulk send stopped by user"))
                    status_box.warning("Stopped.")
                    break

                progress.progress((idx - 1) / total, text=f"Preparing {idx}/{total}: {r.email}")
                status_box.info(f"Sending {idx}/{total} to {r.name} <{r.email}> …")

                try:
                    subj = render_template(subject_template, name=r.name, email=r.email)
                    body = render_template(body_template, name=r.name, email=r.email)
                    msg = build_message(
                        smtp=smtp,
                        to_email=r.email,
                        subject=subj,
                        body=body,
                        attachments=attachments,
                        from_email=sender_email.strip(),
                        from_name=sender_name.strip() or None,
                    )
                    send_via_gmail_smtp(smtp=smtp, message=msg)
                    sent_ok += 1
                    st.session_state.results.append(
                        {"name": r.name, "email": r.email, "status": "sent", "error": ""}
                    )
                    push_log(info(f"Sent to {r.email}"))
                except Exception as e:
                    sent_fail += 1
                    st.session_state.results.append(
                        {"name": r.name, "email": r.email, "status": "failed", "error": str(e)}
                    )
                    push_log(error(f"Failed to send to {r.email}: {e}"))

                progress.progress(idx / total, text=f"Done {idx}/{total}")
                if idx < total and delay_s > 0 and not st.session_state.cancel_bulk:
                    time.sleep(float(delay_s))

            done = sent_ok + sent_fail
            if done == total and not st.session_state.cancel_bulk:
                status_box.success(f"Bulk send complete. Sent: {sent_ok}, Failed: {sent_fail}")
                push_log(info(f"Bulk send complete. Sent={sent_ok} Failed={sent_fail}"))
            else:
                status_box.warning(f"Bulk send stopped. Processed: {done}/{total}")

    st.divider()
    st.subheader("Sending results")
    render_results_table()

st.subheader("Logs")
render_logs_table()

