import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()

GMAIL_SENDER  = os.getenv("GMAIL_SENDER")
GMAIL_APP_PASS = os.getenv("GMAIL_APP_PASS")
ANALYST_EMAIL  = os.getenv("ANALYST_EMAIL")

def send_review_notification(deal: dict, excel_path: str):

    d = deal["extracted_data"]
    flags = d.get("flags", [])

    flag_html = ""
    if flags:
        flag_items = "".join(f"<li style='margin-bottom:6px;color:#92400E;'>{f}</li>" for f in flags)
        flag_html = f"""
        <div style='background:#FEF3C7;border-left:4px solid #F59E0B;padding:12px 16px;margin:16px 0;border-radius:4px;'>
            <strong style='color:#92400E;'>⚠ AI Flags & Warnings</strong>
            <ul style='margin-top:8px;'>{flag_items}</ul>
        </div>"""
    else:
        flag_html = """
        <div style='background:#DCFCE7;border-left:4px solid #16A34A;padding:12px 16px;margin:16px 0;border-radius:4px;'>
            <strong style='color:#166534;'>✅ No flags raised by AI</strong>
        </div>"""

    def fmt(val, prefix="£"):
        if isinstance(val, (int, float)):
            return f"{prefix}{val:,.0f}" if prefix == "£" else f"{val}%"
        return "— Not provided"

    def conf_badge(conf):
        colors = {"high": "#16A34A", "medium": "#D97706", "low": "#DC2626"}
        return f"<span style='color:{colors.get(conf,'#000')};font-weight:bold;'>● {str(conf).upper()}</span>"

    site = d.get("site_address", {}).get("value", "New Deal")

    html = f"""
    <html><body style='font-family:Calibri,Arial,sans-serif;color:#1E293B;max-width:700px;margin:0 auto;'>
      <div style='background:#1A3C2E;padding:24px;border-radius:8px 8px 0 0;'>
        <h1 style='color:white;margin:0;font-size:22px;'>⚡ Electric Land</h1>
        <p style='color:#D8F3DC;margin:6px 0 0;font-size:14px;'>Finance Valuation Agent — New Draft Ready for Review</p>
      </div>
      <div style='background:#FEF3C7;padding:12px 24px;border-bottom:2px solid #F59E0B;'>
        <strong style='color:#92400E;'>⚠ DRAFT — Requires your review and approval before use</strong>
      </div>
      <div style='padding:24px;background:#F8FAFC;'>
        <h2 style='color:#1A3C2E;font-size:16px;border-bottom:2px solid #2D6A4F;padding-bottom:8px;'>📍 Site Information</h2>
        <table style='width:100%;border-collapse:collapse;margin-bottom:16px;'>
          <tr style='background:white;'>
            <td style='padding:8px 12px;border:1px solid #E2E8F0;width:40%;color:#475569;'>Site Address</td>
            <td style='padding:8px 12px;border:1px solid #E2E8F0;font-weight:bold;'>{d.get('site_address',{}).get('value','—')}</td>
            <td style='padding:8px 12px;border:1px solid #E2E8F0;text-align:center;'>{conf_badge(d.get('site_address',{}).get('confidence','low'))}</td>
          </tr>
          <tr style='background:#F8FAFC;'>
            <td style='padding:8px 12px;border:1px solid #E2E8F0;color:#475569;'>Planning Status</td>
            <td style='padding:8px 12px;border:1px solid #E2E8F0;'>{d.get('planning_status',{}).get('value','—')}</td>
            <td style='padding:8px 12px;border:1px solid #E2E8F0;text-align:center;'>{conf_badge(d.get('planning_status',{}).get('confidence','low'))}</td>
          </tr>
          <tr style='background:white;'>
            <td style='padding:8px 12px;border:1px solid #E2E8F0;color:#475569;'>Development Timeline</td>
            <td style='padding:8px 12px;border:1px solid #E2E8F0;'>{d.get('development_timeline_months',{}).get('value','—')} months</td>
            <td style='padding:8px 12px;border:1px solid #E2E8F0;text-align:center;'>{conf_badge(d.get('development_timeline_months',{}).get('confidence','low'))}</td>
          </tr>
        </table>
        <h2 style='color:#1A3C2E;font-size:16px;border-bottom:2px solid #2D6A4F;padding-bottom:8px;'>💷 Financial Summary</h2>
        <table style='width:100%;border-collapse:collapse;margin-bottom:16px;'>
          <tr style='background:white;'>
            <td style='padding:8px 12px;border:1px solid #E2E8F0;color:#475569;'>Purchase Price</td>
            <td style='padding:8px 12px;border:1px solid #E2E8F0;font-weight:bold;'>{fmt(d.get('purchase_price_gbp',{}).get('value'))}</td>
            <td style='padding:8px 12px;border:1px solid #E2E8F0;text-align:center;'>{conf_badge(d.get('purchase_price_gbp',{}).get('confidence','low'))}</td>
          </tr>
          <tr style='background:#F8FAFC;'>
            <td style='padding:8px 12px;border:1px solid #E2E8F0;color:#475569;'>Build Cost</td>
            <td style='padding:8px 12px;border:1px solid #E2E8F0;'>{fmt(d.get('build_cost_gbp',{}).get('value'))}</td>
            <td style='padding:8px 12px;border:1px solid #E2E8F0;text-align:center;'>{conf_badge(d.get('build_cost_gbp',{}).get('confidence','low'))}</td>
          </tr>
          <tr style='background:white;'>
            <td style='padding:8px 12px;border:1px solid #E2E8F0;color:#475569;'>Total Development Cost</td>
            <td style='padding:8px 12px;border:1px solid #E2E8F0;'>{fmt(d.get('total_development_cost_gbp',{}).get('value'))}</td>
            <td style='padding:8px 12px;border:1px solid #E2E8F0;text-align:center;'>{conf_badge(d.get('total_development_cost_gbp',{}).get('confidence','low'))}</td>
          </tr>
          <tr style='background:#F8FAFC;'>
            <td style='padding:8px 12px;border:1px solid #E2E8F0;color:#475569;'>GDV</td>
            <td style='padding:8px 12px;border:1px solid #E2E8F0;'>{fmt(d.get('gdv_gbp',{}).get('value'))}</td>
            <td style='padding:8px 12px;border:1px solid #E2E8F0;text-align:center;'>{conf_badge(d.get('gdv_gbp',{}).get('confidence','low'))}</td>
          </tr>
          <tr style='background:white;'>
            <td style='padding:8px 12px;border:1px solid #E2E8F0;color:#475569;'>Total Rent Roll p.a.</td>
            <td style='padding:8px 12px;border:1px solid #E2E8F0;'>{fmt(d.get('total_rent_roll_pa_gbp',{}).get('value'))}</td>
            <td style='padding:8px 12px;border:1px solid #E2E8F0;text-align:center;'>{conf_badge(d.get('total_rent_roll_pa_gbp',{}).get('confidence','low'))}</td>
          </tr>
          <tr style='background:#F8FAFC;'>
            <td style='padding:8px 12px;border:1px solid #E2E8F0;color:#475569;'>Profit on Cost</td>
            <td style='padding:8px 12px;border:1px solid #E2E8F0;'>{fmt(d.get('profit_on_cost_percent',{}).get('value'), prefix="%")}</td>
            <td style='padding:8px 12px;border:1px solid #E2E8F0;text-align:center;'>{conf_badge(d.get('profit_on_cost_percent',{}).get('confidence','low'))}</td>
          </tr>
        </table>
        {flag_html}
        <div style='background:#1A3C2E;padding:16px;border-radius:6px;margin-top:20px;'>
          <p style='color:white;margin:0;font-size:13px;'>
            📎 Full draft valuation model attached as Excel.<br><br>
            Deal ID: <strong style='color:#D8F3DC;'>{deal['id']}</strong> —
            please review, complete the approval section, and return the signed-off file.
          </p>
        </div>
      </div>
      <div style='background:#E2E8F0;padding:12px 24px;border-radius:0 0 8px 8px;text-align:center;'>
        <p style='color:#64748B;font-size:11px;margin:0;'>
          Auto-generated by Electric Land Finance Valuation Agent. Do not distribute without approval.
        </p>
      </div>
    </body></html>
    """

    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"[DRAFT VALUATION] {site} — Awaiting Review (Deal #{deal['id']})"
    msg["From"]    = GMAIL_SENDER
    msg["To"]      = ANALYST_EMAIL

    msg.attach(MIMEText(html, "html"))

    if os.path.exists(excel_path):
        with open(excel_path, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition",
                            f"attachment; filename={os.path.basename(excel_path)}")
            msg.attach(part)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(GMAIL_SENDER, GMAIL_APP_PASS)
        server.sendmail(GMAIL_SENDER, ANALYST_EMAIL, msg.as_string())

    print(f"✅ Notification sent to {ANALYST_EMAIL}")