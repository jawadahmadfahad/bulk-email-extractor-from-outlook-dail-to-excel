import datetime
import os
import re
import sys
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
import pandas as pd
import pythoncom
import win32com.client

# ============================================================
# Paths & Settings
# ============================================================
RADIO_SHEET_PATH = os.path.expandvars(
    r"%USERPROFILE%\Desktop\Radio Link Sheet updated.xlsx"
)

# Authorized Solace Identifiers (Emails / Names)
AUTHORIZED_IDENTIFIERS = [
    "safeerkhan@solacetelecom.com.pk",
    "farhanraja@solacetelecom.com.pk",
    "support@solacetelecom.com.pk",
    "safeer khan",
    "farhan raja",
    "support team",
    "solace",
]

# Body / Tag Mentions (Agar email body me tag kiya gaya ho)
TAG_MENTION_PATTERNS = [
    r"@safeer",
    r"@farhan",
    r"@solace",
    r"@support",
    r"\bsolace\b",
    r"\bsafeer\b",
    r"\bfarhan\b",
    r"solace\s*team",
    r"team\s*solace",
    r"support\s*team",
    r"dear\s*solace",
    r"hi\s*solace",
]

# Odoo UI Headers
ODOO_HEADERS = [
    "Complaint No.",
    "Complaint Date & Time",
    "Customer",
    "Customer Site",
    "Complaint Type",
    "Field Engineer",
    "Status",
    "Complaint Details",
    "Resolved Date & Time",
    "Resolution Time",
]

# Active Regional Engineers (Clean names for Odoo)
REGION_FE_MAPPING = {
    "Islamabad": "Safeer Khan",
    "Rawalpindi": "Safeer Khan",
    "Attock": "Abbas Ahmed",
    "Lahore": "Muzammil Hassam",
    "Gujranwala": "Muhammad Adnan",
    "Sargodha": "Fayyaz Ul Hassan",
    "Khanewal": "Muhammad Bilal",
    "Multan": "Muhammad Bilal",
    "Bahawalpur": "Ali Faizan Bukhari",
    "DG-Khan": "Fiaz Hussain",
    "Kohat": "Danish Khan",
    "Mansehra": "Firdous Jamal",
    "Karachi": "Mashooque Ali",
    "Hyderabad": "Arbaz Ali",
    "Larkana": "Aziz Ullah",
    "Gawadar": "Nisar",
}


def load_sites_inventory():
  """416 sites کی لسٹ لوڈ کرتا ہے تاکہ درست سائٹ کا نام نکالا جا سکے۔"""
  sites_list = []
  if not os.path.exists(RADIO_SHEET_PATH):
    return sites_list

  try:
    xls = pd.ExcelFile(RADIO_SHEET_PATH)
    for sheet in ["CCC-UPDATED", "SME", "Wateen"]:
      if sheet in xls.sheet_names:
        df = pd.read_excel(RADIO_SHEET_PATH, sheet_name=sheet)
        for _, row in df.iterrows():
          site = str(
              row.get("Site Name", row.get("Customer Site Name", ""))
          ).strip()
          addr = str(row.get("Customer Site Address ", "")).strip()
          exch = str(row.get("Exchange Name", row.get("POP ID", ""))).strip()
          city = str(row.get("City", row.get("City Name", ""))).strip()
          reg = str(
              row.get(
                  "FE Name/Region",
                  row.get("FE Name", row.get("Installed team supervisor", "")),
              )
          ).strip()

          if site and site.lower() != "nan" and site.lower() != "dismantling":
            sites_list.append({
                "site_name": site,
                "address": "" if addr.lower() == "nan" else addr,
                "exchange": "" if exch.lower() == "nan" else exch,
                "city": "" if city.lower() == "nan" else city,
                "region": "" if reg.lower() == "nan" else reg,
            })
  except Exception as e:
    print(f"[!] Error loading inventory: {e}")

  return sites_list


def clean_subject_base(subject):
  """سبجیکٹ سے RE, FW ہٹا کر بنیادی تھریڈ نکالتا ہے۔"""
  clean = re.sub(
      r"^(re:\s*|fw:\s*|fwd:\s*|urgent:\s*|automatic reply:\s*)",
      "",
      subject,
      flags=re.I,
  )
  clean = re.sub(
      r"^(re:\s*|fw:\s*|fwd:\s*|urgent:\s*)", "", clean, flags=re.I
  ).strip()
  normalized = re.sub(r"[^a-zA-Z0-9\s]", " ", clean).strip().lower()
  return normalized, clean


def match_clean_site(clean_subj, body, inventory):
  """ای میل سے سائٹ اور متعلقہ فیلڈ انجینئر میچ کرتا ہے۔"""
  combined_raw = f"{clean_subj} {body}".lower()

  best_site = None
  highest_score = 0

  for item in inventory:
    score = 0
    site_clean = re.sub(r"[^a-zA-Z0-9\s]", " ", item["site_name"].lower())
    tokens = [
        t
        for t in site_clean.split()
        if len(t) >= 4 and t not in ["bank", "limited", "branch", "ptcl"]
    ]

    for t in tokens:
      if t in combined_raw:
        score += 3

    if item["exchange"]:
      ex_clean = re.sub(r"[^a-zA-Z0-9\s]", " ", item["exchange"].lower())
      for et in ex_clean.split():
        if len(et) >= 4 and et not in ["exchange", "ptcl"] and et in combined_raw:
          score += 2

    if score > highest_score and score >= 3:
      highest_score = score
      best_site = item

  if best_site:
    matched_site_name = best_site["site_name"]
    fe_region = best_site["region"] or best_site["city"]
  else:
    matched_site_name = clean_subj.split("||")[0].split("-")[-1].strip()
    fe_region = ""

  assigned_fe = "Safeer Khan"
  for city, fe in REGION_FE_MAPPING.items():
    if city.lower() in combined_raw or (
        fe_region and city.lower() in fe_region.lower()
    ):
      assigned_fe = fe
      break

  return matched_site_name, assigned_fe


def get_odoo_complaint_type(subject, body):
  text = f"{subject} {body}".lower()
  if "packet loss" in text or "packet drop" in text:
    return "Packet Loss"
  elif "latency" in text or "delay" in text or "ping high" in text:
    return "High Latency"
  elif "intermittent" in text or "flapping" in text or "fluctuat" in text:
    return "Radio Link Intermittent"
  elif "power" in text or "electricity" in text:
    return "Power Issue"
  elif "hardware" in text or "faulty" in text or "burnt" in text:
    return "Hardware Fault"
  elif "rssi" in text or "signal" in text:
    return "Low RSSI"
  elif "throughput" in text or "bandwidth" in text:
    return "Low Throughput"
  else:
    return "Radio Link Down"


def is_solace_tagged_or_addressed(item, subject, body):
  """
  سخت فلٹر: صرف وہی ای میل قبول کرے گا جس میں:
  1. Solace / Safeer / Farhan / Support 'To' یا 'CC' میں ہو۔
  2. یا ای میل کی باڈی میں باقاعدہ ٹیگ (@Safeer, @Farhan, Solace Team) کیا گیا ہو۔
  """
  # 1. Check Recipients (To + CC)
  try:
    recipients = item.Recipients
    for i in range(1, recipients.Count + 1):
      recip = recipients.Item(i)
      name = recip.Name.lower() if recip.Name else ""
      address = recip.Address.lower() if recip.Address else ""
      try:
        prop = recip.PropertyAccessor.GetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
        )
        smtp_addr = str(prop).lower()
      except Exception:
        smtp_addr = ""

      all_recip_text = f"{name} {address} {smtp_addr}"
      for ident in AUTHORIZED_IDENTIFIERS:
        if ident in all_recip_text:
          return True
  except Exception:
    pass

  # Fallback to To / CC string properties
  raw_to = getattr(item, "To", "").lower()
  raw_cc = getattr(item, "CC", "").lower()
  for ident in AUTHORIZED_IDENTIFIERS:
    if ident in raw_to or ident in raw_cc:
      return True

  # 2. Check Body Mentions & Tags
  body_lower = body.lower()
  for pattern in TAG_MENTION_PATTERNS:
    if re.search(pattern, body_lower, re.IGNORECASE):
      return True

  # 3. Check Subject for Solace mentions
  subject_lower = subject.lower()
  if "solace" in subject_lower:
    return True

  return False


def generate_odoo_daily_report(target_date=None):
  pythoncom.CoInitialize()
  if target_date is None:
    target_date = datetime.date.today()

  output_excel_path = os.path.expandvars(
      rf"%USERPROFILE%\Desktop\Odoo_Complaints_Report_{target_date}.xlsx"
  )

  inventory = load_sites_inventory()

  print("=" * 75)
  print(f"Solace Telecom - Filtered Odoo Complaints Generator [{target_date}]")
  print("Filter: Only emails where Solace/Safeer/Farhan/Support is tagged/addressed")
  print("=" * 75)

  try:
    try:
      outlook = win32com.client.Dispatch("Outlook.Application")
    except Exception:
      outlook = win32com.client.GetActiveObject("Outlook.Application")

    inbox = outlook.GetNamespace("MAPI").GetDefaultFolder(6)
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)

    records = []
    seen_threads = set()
    skipped_duplicates = 0
    skipped_irrelevant = 0

    print("[*] Scanning Outlook Inbox...")
    for item in messages:
      try:
        received_time = getattr(item, "ReceivedTime", None)
        if not received_time:
          continue

        item_date = datetime.date(
            received_time.year, received_time.month, received_time.day
        )
        if item_date < target_date:
          break

        if item_date == target_date:
          subject = getattr(item, "Subject", "No Subject")
          body = getattr(item, "Body", "")

          # --- STRICT SOLACE TAG / MENTION FILTER ---
          if not is_solace_tagged_or_addressed(item, subject, body):
            skipped_irrelevant += 1
            continue

          norm_sub, clean_sub = clean_subject_base(subject)

          # --- DEDUPLICATION ---
          conv_id = getattr(item, "ConversationID", "")
          unique_key = conv_id if conv_id else " ".join(norm_sub.split()[:6])

          if unique_key in seen_threads:
            skipped_duplicates += 1
            continue

          seen_threads.add(unique_key)

          rec_time_str = received_time.strftime("%Y-%m-%d %H:%M:%S")
          site_name, assigned_fe = match_clean_site(clean_sub, body, inventory)
          comp_type = get_odoo_complaint_type(subject, body)
          clean_details = body.strip().replace("\r\n", "\n")[:1500]

          records.append([
              "",  # Complaint No. (Odoo auto-generates)
              rec_time_str,  # Complaint Date & Time
              "PTCL",  # Customer
              site_name,  # Customer Site
              comp_type,  # Complaint Type
              assigned_fe,  # Field Engineer
              "New",  # Status
              clean_details,  # Complaint Details
              "",  # Resolved Date & Time
              "",  # Resolution Time
          ])
          print(
              f"  [✓] Solace Complaint: {site_name[:30]} | FE: {assigned_fe} |"
              f" Type: {comp_type}"
          )

      except Exception:
        continue

    if not records:
      print(
          f"\n[!] تاریخ ({target_date}) کی کوئی سولیس سے متعلقہ ای میل نہیں"
          " ملی۔"
      )
      print(f"    (غیر متعلقہ چھوڑی گئیں: {skipped_irrelevant})")
      return

    wb = Workbook()
    ws = wb.active
    ws.title = "NOC Complaints"

    ws.append(ODOO_HEADERS)
    header_fill = PatternFill(
        start_color="4F81BD", end_color="4F81BD", fill_type="solid"
    )
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

    for col_num in range(1, len(ODOO_HEADERS) + 1):
      cell = ws.cell(row=1, column=col_num)
      cell.fill = header_fill
      cell.font = header_font
      cell.alignment = Alignment(horizontal="center", vertical="center")

    for r in records:
      ws.append(r)

    for col in ws.columns:
      max_len = max(len(str(cell.value or "")) for cell in col)
      col_letter = col[0].column_letter
      ws.column_dimensions[col_letter].width = min(max(max_len + 3, 14), 50)

    wb.save(output_excel_path)
    print("\n" + "=" * 75)
    print(
        f"[✓] کامیاب! کل {len(records)} اصلی سولیس کمپلینز محفوظ ہوئیں:"
    )
    print(
        f"    (ڈپلیکیٹس ہٹائی گئیں: {skipped_duplicates} | فالتو ای میلز فلٹر"
        f" ہوئیں: {skipped_irrelevant})"
    )
    print(f"[📁] {output_excel_path}")
    print("=" * 75)

  except Exception as e:
    print(f"[!] Error: {e}")


if __name__ == "__main__":
  print("\nکس دن کی رپورٹ بنانی ہے؟")
  print("1. آج کی رپورٹ -> Enter دبائیں")
  print("2. کل کی رپورٹ -> '1' لکھ کر Enter دبائیں")
  print("3. مخصوص تاریخ -> YYYY-MM-DD (مثال: 2026-09-28)")

  user_choice = input("\nتاریخ منتخب کریں (Default: آج): ").strip().lower()

  if not user_choice or user_choice == "today":
    selected_date = datetime.date.today()
  elif user_choice in ["1", "kal", "yesterday"]:
    selected_date = datetime.date.today() - datetime.timedelta(days=1)
  else:
    try:
      selected_date = datetime.datetime.strptime(
          user_choice, "%Y-%m-%d"
      ).date()
    except ValueError:
      print("[!] غلط تاریخ فارمیٹ۔ آج کی تاریخ پر پروسیس کیا جا رہا ہے۔")
      selected_date = datetime.date.today()

  generate_odoo_daily_report(target_date=selected_date)