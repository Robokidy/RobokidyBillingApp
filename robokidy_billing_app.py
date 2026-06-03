"""
Robokidy Innovative Centre - Billing Software v2.0
Data storage: Excel workbook (robokidy_billing_data.xlsx)
Run: pip install openpyxl pillow && python robokidy_billing_app.py
"""

from __future__ import annotations

import base64
import os
import webbrowser
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import tkinter.font as tkfont

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Brand Colours ─────────────────────────────────────────────────────────────
PRIMARY   = "00186D"
SECONDARY = "22BF96"
ACCENT    = "FF6B35"
BG_LIGHT  = "#F0F4FF"
BG_WHITE  = "#FFFFFF"
TEXT_DARK = "#1A1A2E"

DATA_FILE   = Path("robokidy_billing_data.xlsx")
INVOICE_DIR = Path("invoices")
LOGO_PATH   = Path("logo.jpeg")

COURSES = [
    "Python", "Scratch", "Robotics", "Arduino", "IoT", "AI/ML",
    "Web Dev", "App Dev", "3D Printing", "Drone Tech", "Electronics", "Other"
]
COURSE_CODES = {
    "Python": "PY", "Scratch": "SC", "Robotics": "RB", "Arduino": "AR",
    "IoT": "IT", "AI/ML": "AI", "Web Dev": "WD", "App Dev": "AD",
    "3D Printing": "3D", "Drone Tech": "DR", "Electronics": "EL", "Other": "OT"
}
LEVELS    = ["Level 1", "Level 2", "Level 3", "Level 4", "Advanced", "Beginner"]
BATCHES   = ["Morning", "Afternoon", "Evening", "Weekend", "Online"]
PAY_MODES = ["Cash", "UPI", "Bank Transfer", "Cheque", "Card", "Online"]

STUDENT_HEADERS = [
    "Student ID", "Student Name", "Parent Name", "Phone", "Email",
    "Grade/Age", "Course", "Course Code", "Level", "Batch",
    "Joining Date", "Status", "Address", "Notes"
]
INVOICE_HEADERS = [
    "Invoice No", "Date", "Student ID", "Student Name", "Course", "Level",
    "Batch", "Package/Fee", "Discount Type", "Discount Value", "Discount Amount",
    "Tax Type", "Tax Value", "Tax Amount", "Total Amount",
    "Paid Amount", "Balance", "Payment Status", "Payment Mode", "Due Date", "Notes"
]
PAYMENT_HEADERS = [
    "Receipt No", "Date", "Invoice No", "Student ID", "Student Name",
    "Payment Mode", "Paid Amount", "Received By", "Notes"
]
SETTINGS_HEADERS = ["Key", "Value"]


def money(v) -> float:
    try: return round(float(v), 2)
    except: return 0.0

def safe_save(wb, path: Path):
    """Save workbook; if Excel has the file open, show a clear message."""
    import time
    for attempt in range(3):
        try:
            wb.save(path)
            return
        except PermissionError:
            if attempt < 2:
                time.sleep(0.5)
            else:
                raise PermissionError(
                    f"Cannot save — the file is open in Excel or another program.\n\n"
                    f"Please CLOSE the file:\n{path.resolve()}\n\n"
                    f"Then try again."
                )

def logo_base64() -> str | None:
    if LOGO_PATH.exists():
        with open(LOGO_PATH, "rb") as f:
            return base64.b64encode(f.read()).decode()
    return None


# ── Excel Backend ──────────────────────────────────────────────────────────────
class ExcelStore:
    def __init__(self, path: Path = DATA_FILE):
        self.path = path
        self.ensure_workbook()

    def ensure_workbook(self):
        if self.path.exists():
            return
        wb = Workbook()
        ws = wb.active; ws.title = "Dashboard"
        for name, headers in {
            "Students": STUDENT_HEADERS,
            "Invoices": INVOICE_HEADERS,
            "Payments": PAYMENT_HEADERS,
            "Settings": SETTINGS_HEADERS,
        }.items():
            w = wb.create_sheet(name)
            w.append(headers)
            self._style_header(w)
        ws.append(["ROBOKIDY INNOVATIVE CENTRE - BILLING DASHBOARD"])
        ws.append(["Open the app to update dashboard figures automatically."])
        settings = wb["Settings"]
        for k, v in [
            ("Company Name", "Robokidy Innovative Centre"),
            ("Address", "91/3 Ramachandra Nagar, Kallakurichi"),
            ("Phone", "+91 8300967241"),
            ("Email", "info@robokidy.com"),
            ("Website", "www.robokidy.com"),
            ("GSTIN", "33AAOCR3798M1Z1"),
            ("Invoice Prefix", "RKI-INV"),
            ("Receipt Prefix", "RKI-RCP"),
        ]: settings.append([k, v])
        safe_save(wb, self.path)

    def wb(self): return load_workbook(self.path)

    @staticmethod
    def _style_header(ws):
        fill = PatternFill("solid", fgColor=PRIMARY)
        thin = Side(style="thin", color="D9E2F3")
        for cell in ws[1]:
            cell.fill = fill
            cell.font = Font(color="FFFFFF", bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
        ws.freeze_panes = "A2"
        for i, _ in enumerate(ws[1], 1):
            ws.column_dimensions[get_column_letter(i)].width = 20

    def read_rows(self, sheet: str):
        wb = self.wb(); ws = wb[sheet]
        headers = [c.value for c in ws[1]]
        return [
            dict(zip(headers, row))
            for row in ws.iter_rows(min_row=2, values_only=True)
            if any(v is not None for v in row)
        ]

    def append(self, sheet: str, values: list):
        wb = self.wb(); ws = wb[sheet]
        ws.append(values)
        self._style_header(ws)
        safe_save(wb, self.path)

    def settings(self):
        return {r["Key"]: r["Value"] for r in self.read_rows("Settings")}

    def update_dashboard(self):
        wb = self.wb(); ws = wb["Dashboard"]
        ws.delete_rows(1, ws.max_row)
        inv = self.read_rows("Invoices")
        stu = self.read_rows("Students")
        pay = self.read_rows("Payments")
        ws.append(["ROBOKIDY INNOVATIVE CENTRE - BILLING DASHBOARD"])
        ws.append(["Last Updated", datetime.now().strftime("%d-%m-%Y %I:%M %p")])
        ws.append([])
        ws.append(["Total Students", len(stu)])
        ws.append(["Total Invoices", len(inv)])
        ws.append(["Total Receipts", len(pay)])
        ws.append(["Total Billing", sum(money(r.get("Total Amount")) for r in inv)])
        ws.append(["Total Received", sum(money(r.get("Paid Amount")) for r in inv)])
        ws.append(["Total Balance", sum(money(r.get("Balance")) for r in inv)])
        ws["A1"].font = Font(size=16, bold=True, color=PRIMARY)
        ws.column_dimensions["A"].width = 30; ws.column_dimensions["B"].width = 25
        safe_save(wb, self.path)

    def next_student_id(self, course: str, level: str) -> str:
        code  = COURSE_CODES.get(course, course[:2].upper())
        lvl   = level.replace("Level ", "L").replace(" ", "").upper()
        year  = str(datetime.now().year)
        prefix = f"RK-{code}-{lvl}-{year}-"
        existing = [r.get("Student ID","") for r in self.read_rows("Students")]
        nums = []
        for sid in existing:
            if str(sid).startswith(prefix):
                try: nums.append(int(str(sid).split("-")[-1]))
                except: pass
        n = max(nums) + 1 if nums else 1
        return f"{prefix}{n:03d}"

    def next_invoice_no(self) -> str:
        prefix = self.settings().get("Invoice Prefix", "RKI-INV")
        year   = datetime.now().year
        p = f"{prefix}-{year}-"
        existing = [r.get("Invoice No","") for r in self.read_rows("Invoices")]
        nums = [int(str(v).split("-")[-1]) for v in existing if str(v).startswith(p)]
        return f"{p}{(max(nums)+1 if nums else 1):04d}"

    def next_receipt_no(self) -> str:
        prefix = self.settings().get("Receipt Prefix", "RKI-RCP")
        year   = datetime.now().year
        p = f"{prefix}-{year}-"
        existing = [r.get("Receipt No","") for r in self.read_rows("Payments")]
        nums = [int(str(v).split("-")[-1]) for v in existing if str(v).startswith(p)]
        return f"{p}{(max(nums)+1 if nums else 1):04d}"

    def add_student(self, d: dict) -> str:
        sid = self.next_student_id(d["course"], d["level"])
        code = COURSE_CODES.get(d["course"], d["course"][:2].upper())
        self.append("Students", [
            sid, d["name"], d["parent"], d["phone"], d["email"],
            d["grade"], d["course"], code, d["level"], d["batch"],
            d["joining"], "Active", d["address"], d["notes"]
        ])
        self.update_dashboard()
        return sid

    def create_invoice(self, d: dict) -> str:
        inv_no = self.next_invoice_no()
        fee  = money(d["fee"])
        # Discount
        if d["disc_type"] == "%":
            disc_amt = round(fee * money(d["disc_val"]) / 100, 2)
        else:
            disc_amt = money(d["disc_val"])
        subtotal = fee - disc_amt
        # Tax
        if d["tax_type"] == "%":
            tax_amt = round(subtotal * money(d["tax_val"]) / 100, 2)
        else:
            tax_amt = money(d["tax_val"])
        total  = round(subtotal + tax_amt, 2)
        paid   = money(d["paid"])
        bal    = round(total - paid, 2)
        status = "Paid" if bal <= 0 else ("Partial" if paid > 0 else "Unpaid")
        self.append("Invoices", [
            inv_no, d["date"], d["student_id"], d["student_name"],
            d["course"], d["level"], d["batch"],
            fee, d["disc_type"], money(d["disc_val"]), disc_amt,
            d["tax_type"], money(d["tax_val"]), tax_amt,
            total, paid, bal, status, d["mode"], d["due_date"], d["notes"]
        ])
        if paid > 0:
            self._save_payment(inv_no, d["student_id"], d["student_name"],
                               d["mode"], paid, d["received_by"], "Initial payment", update=False)
        self.update_dashboard()
        return inv_no

    def add_payment(self, d: dict) -> str:
        r = self._save_payment(d["invoice_no"], d["student_id"], d["student_name"],
                                d["mode"], d["paid"], d["received_by"], d["notes"], update=True)
        self.update_dashboard()
        return r

    def _save_payment(self, inv_no, sid, sname, mode, paid, rcvd, notes, update=True) -> str:
        rno = self.next_receipt_no()
        self.append("Payments", [
            rno, datetime.now().strftime("%d-%m-%Y"), inv_no, sid, sname,
            mode, money(paid), rcvd, notes
        ])
        if update:
            wb = self.wb(); ws = wb["Invoices"]
            hdrs = [c.value for c in ws[1]]
            ic = hdrs.index("Invoice No")+1
            pc = hdrs.index("Paid Amount")+1
            bc = hdrs.index("Balance")+1
            tc = hdrs.index("Total Amount")+1
            sc = hdrs.index("Payment Status")+1
            for r in range(2, ws.max_row+1):
                if ws.cell(r, ic).value == inv_no:
                    new_p = money(ws.cell(r, pc).value) + money(paid)
                    tot   = money(ws.cell(r, tc).value)
                    new_b = round(tot - new_p, 2)
                    ws.cell(r, pc).value = new_p
                    ws.cell(r, bc).value = new_b
                    ws.cell(r, sc).value = "Paid" if new_b <= 0 else "Partial"
                    break
            safe_save(wb, self.path)
        return rno

    def get_invoice(self, inv_no: str):
        for r in self.read_rows("Invoices"):
            if str(r.get("Invoice No")) == inv_no:
                return r
        return None

    def get_students(self): return self.read_rows("Students")
    def get_invoices(self): return self.read_rows("Invoices")
    def get_payments(self): return self.read_rows("Payments")


# ── Invoice / Receipt HTML ─────────────────────────────────────────────────────
def make_html(company: dict, inv: dict, receipt_no: str | None = None) -> str:
    logo_b64 = logo_base64()
    logo_html = (f'<img src="data:image/jpeg;base64,{logo_b64}" style="height:72px;object-fit:contain;">'
                 if logo_b64 else
                 f'<div style="font-size:28px;font-weight:900;color:#{PRIMARY};">🤖 Robokidy</div>')

    is_receipt = receipt_no is not None
    badge_text = "RECEIPT" if is_receipt else "INVOICE"
    badge_color = SECONDARY if is_receipt else PRIMARY

    def row(label, val):
        return f"""
        <tr>
          <td style="padding:10px 14px;background:#f8f9ff;font-weight:600;color:#555;width:42%;border-bottom:1px solid #e8ecf4;">{label}</td>
          <td style="padding:10px 14px;border-bottom:1px solid #e8ecf4;color:#111;">{val}</td>
        </tr>"""

    disc_type = inv.get("Discount Type", "")
    disc_val  = money(inv.get("Discount Value", 0))
    disc_amt  = money(inv.get("Discount Amount", 0))
    tax_type  = inv.get("Tax Type", "")
    tax_val   = money(inv.get("Tax Value", 0))
    tax_amt   = money(inv.get("Tax Amount", 0))

    disc_str = f"₹ {disc_amt:,.2f}" + (f" ({disc_val}%)" if disc_type == "%" else "")
    tax_str  = f"₹ {tax_amt:,.2f}" + (f" ({tax_val}% GST)" if tax_type == "%" else " (GST)")

    status = inv.get("Payment Status","")
    status_color = {"Paid": "#16a34a", "Partial": "#d97706", "Unpaid": "#dc2626"}.get(status, "#555")

    gstin = company.get("GSTIN","")
    gstin_line = f'<div style="font-size:11px;color:#888;">GSTIN: {gstin}</div>' if gstin else ""

    return f"""<!doctype html>
<html><head><meta charset="utf-8">
<title>{badge_text} - {inv.get('Invoice No','')}</title>
<link href="https://fonts.googleapis.com/css2?family=Poppins:wght@400;600;700;900&display=swap" rel="stylesheet">
<style>
  * {{ box-sizing: border-box; margin:0; padding:0; }}
  body {{ font-family:'Poppins',sans-serif; background:#eef2ff; color:#222; }}
  .page {{ max-width:820px; margin:30px auto; background:#fff; border-radius:18px;
            box-shadow:0 8px 40px rgba(0,24,109,.13); overflow:hidden; }}
  .header {{ background:linear-gradient(135deg, #{PRIMARY} 0%, #0a2d9c 100%);
             padding:28px 36px; display:flex; justify-content:space-between; align-items:center; }}
  .badge {{ background:#{badge_color}; color:#fff; font-weight:700; font-size:15px;
             padding:10px 24px; border-radius:30px; letter-spacing:2px; }}
  .body {{ padding:32px 36px; }}
  .meta-grid {{ display:grid; grid-template-columns:1fr 1fr; gap:20px; margin-bottom:28px; }}
  .meta-box {{ background:#f8f9ff; border:1.5px solid #dde3f8; border-radius:12px; padding:16px 20px; }}
  .meta-box h4 {{ font-size:11px; text-transform:uppercase; letter-spacing:1.5px;
                  color:#{SECONDARY}; margin-bottom:8px; font-weight:700; }}
  .meta-box p {{ font-size:13px; color:#333; line-height:1.7; }}
  .meta-box strong {{ color:#{PRIMARY}; }}
  table {{ width:100%; border-collapse:collapse; border-radius:12px; overflow:hidden;
           border:1.5px solid #dde3f8; }}
  .total-row td {{ background:#{PRIMARY} !important; color:#fff !important;
                   font-weight:700; font-size:16px; padding:14px !important; }}
  .status-badge {{ display:inline-block; padding:4px 14px; border-radius:20px;
                   font-size:12px; font-weight:700; color:{status_color};
                   background:{status_color}18; border:1.5px solid {status_color}44; }}
  .footer {{ margin-top:36px; display:flex; justify-content:space-between;
             padding-top:20px; border-top:2px dashed #dde3f8; }}
  .sig {{ text-align:center; }}
  .sig-line {{ width:160px; border-top:2px solid #555; margin:0 auto 6px; padding-top:6px;
               font-size:12px; color:#666; }}
  .print-btn {{ display:flex; gap:12px; margin-bottom:24px; }}
  .btn {{ padding:10px 24px; border:none; border-radius:8px; font-family:'Poppins',sans-serif;
          font-weight:600; font-size:13px; cursor:pointer; transition:.2s; }}
  .btn-primary {{ background:#{PRIMARY}; color:#fff; }}
  .btn-secondary {{ background:#{SECONDARY}; color:#fff; }}
  @media print {{ .print-btn,.no-print {{ display:none!important; }}
    body {{ background:#fff; }} .page {{ margin:0; box-shadow:none; border-radius:0; }} }}
</style>
</head><body>
<div class="page">
  <div class="header">
    <div style="display:flex;align-items:center;gap:18px;">
      {logo_html}
      <div>
        <div style="color:#fff;font-size:20px;font-weight:700;">{company.get('Company Name','Robokidy Innovative Centre')}</div>
        <div style="color:#a8b8ff;font-size:12px;">{company.get('Address','')}</div>
        <div style="color:#a8b8ff;font-size:12px;">{company.get('Phone','')} &nbsp;|&nbsp; {company.get('Email','')}</div>
        {gstin_line}
      </div>
    </div>
    <div class="badge">{badge_text}</div>
  </div>

  <div class="body">
    <div class="print-btn no-print">
      <button class="btn btn-primary" onclick="window.print()">🖨️ Print / Save PDF</button>
      <button class="btn btn-secondary" onclick="window.close()">✕ Close</button>
    </div>

    <div class="meta-grid">
      <div class="meta-box">
        <h4>{'Receipt' if is_receipt else 'Invoice'} Details</h4>
        <p><strong>{'Receipt' if is_receipt else 'Invoice'} No:</strong> {receipt_no if is_receipt else inv.get('Invoice No','')}</p>
        {'<p><strong>Invoice No:</strong> '+inv.get('Invoice No','')+'</p>' if is_receipt else ''}
        <p><strong>Date:</strong> {inv.get('Date','')}</p>
        {'<p><strong>Due Date:</strong> '+str(inv.get('Due Date',''))+'</p>' if not is_receipt else ''}
      </div>
      <div class="meta-box">
        <h4>Student Details</h4>
        <p><strong>ID:</strong> {inv.get('Student ID','')}</p>
        <p><strong>Name:</strong> {inv.get('Student Name','')}</p>
        <p><strong>Course:</strong> {inv.get('Course','')} &nbsp;|&nbsp; {inv.get('Level','')}</p>
        <p><strong>Batch:</strong> {inv.get('Batch','')}</p>
      </div>
    </div>

    <table>
      {row("Course / Program", f"{inv.get('Course','')} — {inv.get('Level','')}")}
      {row("Package / Fee", f"₹ {money(inv.get('Package/Fee')):,.2f}")}
      {row("Discount", disc_str if disc_amt else "–")}
      {row("Tax (GST)", tax_str if tax_amt else "–")}
      <tr class="total-row">
        <td style="padding:10px 14px;">TOTAL AMOUNT</td>
        <td style="padding:10px 14px;">₹ {money(inv.get('Total Amount')):,.2f}</td>
      </tr>
      {row("Paid Amount", f"₹ {money(inv.get('Paid Amount')):,.2f}")}
      {row("Balance Due", f"₹ {money(inv.get('Balance')):,.2f}")}
      {row("Payment Mode", inv.get('Payment Mode',''))}
      {row("Payment Status", f'<span class="status-badge">{status}</span>')}
    </table>

    {'<p style="margin-top:16px;font-size:13px;color:#666;"><b>Notes:</b> ' + str(inv.get('Notes','')) + '</p>' if inv.get('Notes') else ''}

    <div class="footer">
      <div class="sig">
        <div class="sig-line">Parent / Student Signature</div>
      </div>
      <div style="text-align:center;font-size:12px;color:#888;">
        <div>Thank you for choosing Robokidy!</div>
        <div style="color:#{SECONDARY};font-weight:600;">{company.get('Website','')}</div>
      </div>
      <div class="sig">
        <div class="sig-line">Authorised Signature</div>
      </div>
    </div>
  </div>
</div>
</body></html>"""


# ── UI Helpers ─────────────────────────────────────────────────────────────────
def styled_btn(parent, text, cmd, bg=f"#{PRIMARY}", fg="white", pad=(8,18)):
    b = tk.Button(parent, text=text, command=cmd,
                  bg=bg, fg=fg, font=("Segoe UI", 9, "bold"),
                  relief="flat", cursor="hand2",
                  padx=pad[1], pady=pad[0], bd=0, activebackground=bg)
    return b

def section_label(parent, text):
    return tk.Label(parent, text=text, bg=BG_WHITE,
                    font=("Segoe UI", 11, "bold"), fg=f"#{PRIMARY}")


# ── Main App ───────────────────────────────────────────────────────────────────
class BillingApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Robokidy Innovative Centre — Billing Software v2.0")
        self.geometry("1200x780")
        self.minsize(900, 600)
        self.store  = ExcelStore()
        self.configure(bg=BG_LIGHT)
        INVOICE_DIR.mkdir(exist_ok=True)
        self._build_ui()
        self.refresh_all()

    # ── Build UI ───────────────────────────────────────────────────────────────
    def _build_ui(self):
        # Top bar
        topbar = tk.Frame(self, bg=f"#{PRIMARY}", height=58)
        topbar.pack(fill="x"); topbar.pack_propagate(False)
        tk.Label(topbar, text="  🤖  ROBOKIDY INNOVATIVE CENTRE",
                 bg=f"#{PRIMARY}", fg="white",
                 font=("Segoe UI", 16, "bold")).pack(side="left", padx=14)
        tk.Label(topbar, text="Billing Software v2.0",
                 bg=f"#{PRIMARY}", fg=f"#{SECONDARY}",
                 font=("Segoe UI", 10)).pack(side="left")
        self._clock_lbl = tk.Label(topbar, bg=f"#{PRIMARY}", fg="#aabbff",
                                   font=("Segoe UI", 10))
        self._clock_lbl.pack(side="right", padx=18)
        self._tick()

        # Notebook
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TNotebook",       background=BG_LIGHT, borderwidth=0)
        style.configure("TNotebook.Tab",   background="#dde3f8", foreground=TEXT_DARK,
                        padding=[16,9], font=("Segoe UI", 10, "bold"))
        style.map("TNotebook.Tab",
                  background=[("selected", f"#{SECONDARY}")],
                  foreground=[("selected", "white")])
        style.configure("Treeview",        rowheight=26, font=("Segoe UI", 9))
        style.configure("Treeview.Heading",font=("Segoe UI", 9, "bold"),
                        background=f"#{PRIMARY}", foreground="white")
        style.map("Treeview", background=[("selected", f"#{SECONDARY}")])

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=10, pady=8)

        self._tab_dashboard()
        self._tab_students()
        self._tab_invoice()
        self._tab_payment()
        self._tab_reports()

    def _tick(self):
        self._clock_lbl.config(text=datetime.now().strftime("  %d %b %Y   %I:%M:%S %p  "))
        self.after(1000, self._tick)

    # ── Shared widgets ─────────────────────────────────────────────────────────
    def _field(self, parent, label, row, col, width=22, colspan=1):
        tk.Label(parent, text=label, bg=BG_WHITE, fg="#444",
                 font=("Segoe UI", 9)).grid(row=row, column=col, sticky="w",
                                             padx=(10,4), pady=5)
        e = ttk.Entry(parent, width=width)
        e.grid(row=row, column=col+1, columnspan=colspan, sticky="ew", padx=(0,10), pady=5)
        return e

    def _combo(self, parent, label, values, row, col, width=22):
        tk.Label(parent, text=label, bg=BG_WHITE, fg="#444",
                 font=("Segoe UI", 9)).grid(row=row, column=col, sticky="w",
                                             padx=(10,4), pady=5)
        c = ttk.Combobox(parent, values=values, state="readonly", width=width)
        c.grid(row=row, column=col+1, sticky="ew", padx=(0,10), pady=5)
        return c

    def _tree(self, parent, cols):
        fr = tk.Frame(parent, bg=BG_WHITE)
        fr.pack(fill="both", expand=True, padx=10, pady=(0,10))
        t = ttk.Treeview(fr, columns=cols, show="headings", height=11,
                         selectmode="browse")
        vs = ttk.Scrollbar(fr, orient="vertical", command=t.yview)
        hs = ttk.Scrollbar(fr, orient="horizontal", command=t.xview)
        t.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        for c in cols:
            t.heading(c, text=c)
            t.column(c, width=120, anchor="center", minwidth=80)
        t.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        fr.grid_rowconfigure(0, weight=1); fr.grid_columnconfigure(0, weight=1)
        # alternating row colours
        t.tag_configure("odd",  background="#f5f7ff")
        t.tag_configure("even", background=BG_WHITE)
        return t

    def _insert_rows(self, tree, rows, headers):
        for i in tree.get_children(): tree.delete(i)
        for idx, r in enumerate(rows):
            tag = "odd" if idx % 2 else "even"
            tree.insert("", "end", values=[r.get(h,"") for h in headers], tags=(tag,))

    def _form_frame(self, parent):
        f = tk.Frame(parent, bg=BG_WHITE, padx=10, pady=12,
                     relief="flat", bd=0)
        f.pack(fill="x", padx=10, pady=(10,4))
        return f

    # ── Dashboard tab ──────────────────────────────────────────────────────────
    def _tab_dashboard(self):
        f = tk.Frame(self.nb, bg=BG_WHITE); self.nb.add(f, text="📊  Dashboard")
        cards_row = tk.Frame(f, bg=BG_WHITE); cards_row.pack(fill="x", padx=20, pady=18)
        self._dash_cards = {}
        for key in ["Students","Invoices","Receipts","Total Billing","Received","Balance"]:
            card = tk.Frame(cards_row, bg=BG_LIGHT, bd=0, relief="flat",
                            width=160, height=90)
            card.pack(side="left", padx=8, pady=4, expand=True, fill="both")
            card.pack_propagate(False)
            tk.Label(card, text=key, bg=BG_LIGHT, font=("Segoe UI", 9),
                     fg="#666").pack(pady=(14,2))
            lbl = tk.Label(card, text="—", bg=BG_LIGHT,
                           font=("Segoe UI", 16, "bold"), fg=f"#{PRIMARY}")
            lbl.pack()
            self._dash_cards[key] = lbl

        mid = tk.Frame(f, bg=BG_WHITE); mid.pack(fill="x", padx=20, pady=4)
        styled_btn(mid, "🔄  Refresh Dashboard", self.refresh_all,
                   bg=f"#{SECONDARY}").pack(side="left", padx=4)
        styled_btn(mid, "📂  Open Excel File",
                   lambda: webbrowser.open(DATA_FILE.resolve().as_uri())).pack(side="left", padx=4)

        sep = tk.Frame(f, bg="#dde3f8", height=2); sep.pack(fill="x", padx=20, pady=12)

        # Recent invoices
        tk.Label(f, text="Recent Invoices", bg=BG_WHITE,
                 font=("Segoe UI", 11, "bold"), fg=f"#{PRIMARY}").pack(anchor="w", padx=20)
        short_cols = ["Invoice No","Date","Student Name","Total Amount","Paid Amount","Balance","Payment Status"]
        self._dash_inv_tree = self._tree(f, short_cols)

    # ── Students tab ──────────────────────────────────────────────────────────
    def _tab_students(self):
        f = tk.Frame(self.nb, bg=BG_WHITE); self.nb.add(f, text="👤  Students")
        section_label(f, "  Add New Student").pack(anchor="w", padx=14, pady=(12,0))
        form = self._form_frame(f)
        form.columnconfigure((1,3,5), weight=1)

        self.s_name    = self._field(form, "Student Name *", 0, 0)
        self.s_parent  = self._field(form, "Parent Name",    0, 2)
        self.s_phone   = self._field(form, "Phone *",        0, 4)

        self.s_email   = self._field(form, "Email",     1, 0)
        self.s_grade   = self._field(form, "Grade / Age", 1, 2)
        self.s_address = self._field(form, "Address",   1, 4)

        self.s_course  = self._combo(form, "Course *",  COURSES,          2, 0)
        self.s_level   = self._combo(form, "Level *",   LEVELS,           2, 2)
        self.s_batch   = self._combo(form, "Batch",     BATCHES,          2, 4)

        self.s_joining = self._field(form, "Joining Date", 3, 0)
        self.s_joining.insert(0, datetime.now().strftime("%d-%m-%Y"))
        self.s_notes   = self._field(form, "Notes", 3, 2, width=50, colspan=3)

        self.s_id_preview = tk.Label(form, text="Student ID: (auto-generated)",
                                     bg=BG_WHITE, font=("Segoe UI", 9, "italic"),
                                     fg=f"#{SECONDARY}")
        self.s_id_preview.grid(row=4, column=0, columnspan=4, sticky="w", padx=10, pady=4)

        self.s_course.bind("<<ComboboxSelected>>", self._preview_id)
        self.s_level.bind("<<ComboboxSelected>>",  self._preview_id)

        btn_row = tk.Frame(form, bg=BG_WHITE); btn_row.grid(row=5, column=0, columnspan=6, pady=8)
        styled_btn(btn_row, "➕  Add Student", self._add_student,
                   bg=f"#{SECONDARY}").pack(side="left", padx=6)
        styled_btn(btn_row, "🔄  Clear Form", self._clear_student_form).pack(side="left", padx=6)

        tk.Frame(f, bg="#dde3f8", height=2).pack(fill="x", padx=10, pady=4)
        tk.Label(f, text="  All Students", bg=BG_WHITE,
                 font=("Segoe UI",10,"bold"), fg=f"#{PRIMARY}").pack(anchor="w")
        self.student_tree = self._tree(f, STUDENT_HEADERS)

    def _preview_id(self, _=None):
        c = self.s_course.get(); l = self.s_level.get()
        if c and l:
            code = COURSE_CODES.get(c, c[:2].upper())
            lvl  = l.replace("Level ","L").replace(" ","").upper()
            self.s_id_preview.config(
                text=f"Student ID preview: RK-{code}-{lvl}-{datetime.now().year}-###")

    # ── Invoice tab ────────────────────────────────────────────────────────────
    def _tab_invoice(self):
        f = tk.Frame(self.nb, bg=BG_WHITE); self.nb.add(f, text="🧾  Create Invoice")
        section_label(f, "  New Invoice").pack(anchor="w", padx=14, pady=(12,0))
        form = self._form_frame(f)
        form.columnconfigure((1,3,5), weight=1)

        tk.Label(form, text="Select Student *", bg=BG_WHITE, fg="#444",
                 font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w", padx=(10,4), pady=5)
        self.i_student = ttk.Combobox(form, width=32, state="readonly")
        self.i_student.grid(row=0, column=1, columnspan=2, sticky="ew", padx=(0,10), pady=5)
        self.i_student.bind("<<ComboboxSelected>>", self._fill_student_invoice)

        self.i_date = self._field(form, "Invoice Date", 0, 3)
        self.i_date.insert(0, datetime.now().strftime("%d-%m-%Y"))
        self.i_due  = self._field(form, "Due Date",     0, 5)

        self.i_course = self._field(form, "Course",  1, 0)
        self.i_level  = self._field(form, "Level",   1, 2)
        self.i_batch  = self._field(form, "Batch",   1, 4)

        self.i_fee    = self._field(form, "Package Fee (₹) *", 2, 0)

        # Discount row with type toggle
        tk.Label(form, text="Discount", bg=BG_WHITE, fg="#444",
                 font=("Segoe UI",9)).grid(row=2, column=2, sticky="w", padx=(10,4), pady=5)
        disc_inner = tk.Frame(form, bg=BG_WHITE)
        disc_inner.grid(row=2, column=3, sticky="ew", padx=(0,6), pady=5)
        self.i_disc_val = ttk.Entry(disc_inner, width=12); self.i_disc_val.pack(side="left")
        self.i_disc_val.insert(0, "0")
        self.i_disc_type = tk.StringVar(value="%")
        tk.Radiobutton(disc_inner, text="%", variable=self.i_disc_type, value="%",
                       bg=BG_WHITE, font=("Segoe UI",9), fg=f"#{PRIMARY}").pack(side="left", padx=3)
        tk.Radiobutton(disc_inner, text="₹", variable=self.i_disc_type, value="₹",
                       bg=BG_WHITE, font=("Segoe UI",9), fg=f"#{PRIMARY}").pack(side="left")

        # Tax row with type toggle
        tk.Label(form, text="Tax (GST)", bg=BG_WHITE, fg="#444",
                 font=("Segoe UI",9)).grid(row=2, column=4, sticky="w", padx=(10,4), pady=5)
        tax_inner = tk.Frame(form, bg=BG_WHITE)
        tax_inner.grid(row=2, column=5, sticky="ew", padx=(0,6), pady=5)
        self.i_tax_val = ttk.Entry(tax_inner, width=12); self.i_tax_val.pack(side="left")
        self.i_tax_val.insert(0, "0")
        self.i_tax_type = tk.StringVar(value="%")
        tk.Radiobutton(tax_inner, text="%", variable=self.i_tax_type, value="%",
                       bg=BG_WHITE, font=("Segoe UI",9), fg=f"#{PRIMARY}").pack(side="left", padx=3)
        tk.Radiobutton(tax_inner, text="₹", variable=self.i_tax_type, value="₹",
                       bg=BG_WHITE, font=("Segoe UI",9), fg=f"#{PRIMARY}").pack(side="left")

        self.i_paid     = self._field(form, "Paid Now (₹)",   3, 0); self.i_paid.insert(0, "0")
        self.i_mode     = self._combo(form, "Payment Mode",   PAY_MODES, 3, 2)
        self.i_received = self._field(form, "Received By",    3, 4); self.i_received.insert(0, "Admin")
        self.i_mode.set("Cash")

        self.i_notes = self._field(form, "Notes", 4, 0, width=70, colspan=5)

        # Live total preview
        self._total_lbl = tk.Label(form, text="",
                                   bg=BG_WHITE, font=("Segoe UI",10,"bold"), fg=f"#{ACCENT}")
        self._total_lbl.grid(row=5, column=0, columnspan=4, sticky="w", padx=10, pady=4)
        for w in [self.i_fee, self.i_disc_val, self.i_tax_val, self.i_paid]:
            w.bind("<KeyRelease>", self._preview_total)
        self.i_disc_type.trace_add("write", lambda *_: self._preview_total())
        self.i_tax_type.trace_add("write",  lambda *_: self._preview_total())

        btn_row = tk.Frame(form, bg=BG_WHITE); btn_row.grid(row=6, column=0, columnspan=6, pady=8)
        styled_btn(btn_row, "🧾  Create Invoice",   self._create_invoice,
                   bg=f"#{SECONDARY}").pack(side="left", padx=6)
        styled_btn(btn_row, "🖨️  Print Selected",   self._print_invoice).pack(side="left", padx=6)
        styled_btn(btn_row, "🔄  Clear Form",        self._clear_invoice_form,
                   bg="#888").pack(side="left", padx=6)

        tk.Frame(f, bg="#dde3f8", height=2).pack(fill="x", padx=10, pady=4)
        tk.Label(f, text="  All Invoices", bg=BG_WHITE,
                 font=("Segoe UI",10,"bold"), fg=f"#{PRIMARY}").pack(anchor="w")
        self.invoice_tree = self._tree(f, INVOICE_HEADERS)

    def _preview_total(self, _=None):
        try:
            fee = money(self.i_fee.get())
            dv  = money(self.i_disc_val.get())
            tv  = money(self.i_tax_val.get())
            da  = round(fee * dv/100, 2) if self.i_disc_type.get()=="%" else dv
            sub = fee - da
            ta  = round(sub * tv/100, 2) if self.i_tax_type.get()=="%" else tv
            tot = sub + ta
            pd  = money(self.i_paid.get())
            bal = tot - pd
            self._total_lbl.config(
                text=f"  Fee ₹{fee:,.2f}  –  Disc ₹{da:,.2f}  +  Tax ₹{ta:,.2f}"
                     f"  =  Total ₹{tot:,.2f}   |   Paid ₹{pd:,.2f}   |   Balance ₹{bal:,.2f}")
        except: pass

    # ── Payment tab ────────────────────────────────────────────────────────────
    def _tab_payment(self):
        f = tk.Frame(self.nb, bg=BG_WHITE); self.nb.add(f, text="💳  Payments")
        section_label(f, "  Add Payment / Generate Receipt").pack(anchor="w", padx=14, pady=(12,0))
        form = self._form_frame(f)
        form.columnconfigure((1,3,5), weight=1)

        self.p_invoice  = self._field(form, "Invoice No *",   0, 0)
        self.p_amount   = self._field(form, "Paid Amount (₹)*", 0, 2)
        self.p_mode     = self._combo(form, "Payment Mode", PAY_MODES, 0, 4)
        self.p_mode.set("Cash")
        self.p_received = self._field(form, "Received By",  1, 0); self.p_received.insert(0, "Admin")
        self.p_notes    = self._field(form, "Notes",        1, 2, width=50, colspan=3)

        # Invoice lookup
        self._inv_info_lbl = tk.Label(form, text="", bg=BG_WHITE,
                                      font=("Segoe UI",9,"italic"), fg="#666")
        self._inv_info_lbl.grid(row=2, column=0, columnspan=6, sticky="w", padx=10, pady=2)
        self.p_invoice.bind("<FocusOut>", self._lookup_invoice)
        self.p_invoice.bind("<Return>",   self._lookup_invoice)

        btn_row = tk.Frame(form, bg=BG_WHITE); btn_row.grid(row=3, column=0, columnspan=6, pady=8)
        styled_btn(btn_row, "💳  Add Payment & Print Receipt",
                   self._add_payment, bg=f"#{SECONDARY}").pack(side="left", padx=6)

        tk.Frame(f, bg="#dde3f8", height=2).pack(fill="x", padx=10, pady=4)
        tk.Label(f, text="  Payment History", bg=BG_WHITE,
                 font=("Segoe UI",10,"bold"), fg=f"#{PRIMARY}").pack(anchor="w")
        self.payment_tree = self._tree(f, PAYMENT_HEADERS)

    def _lookup_invoice(self, _=None):
        inv = self.store.get_invoice(self.p_invoice.get().strip())
        if inv:
            self._inv_info_lbl.config(
                fg=f"#{SECONDARY}",
                text=f"✓  {inv.get('Student Name','')}  |  Total: ₹{money(inv.get('Total Amount')):,.2f}"
                     f"  |  Paid: ₹{money(inv.get('Paid Amount')):,.2f}"
                     f"  |  Balance: ₹{money(inv.get('Balance')):,.2f}"
                     f"  |  Status: {inv.get('Payment Status','')}")
        else:
            self._inv_info_lbl.config(fg="#dc2626", text="✗  Invoice not found")

    # ── Reports tab ────────────────────────────────────────────────────────────
    def _tab_reports(self):
        f = tk.Frame(self.nb, bg=BG_WHITE); self.nb.add(f, text="📁  Reports & Settings")
        section_label(f, "  Quick Actions").pack(anchor="w", padx=14, pady=(14,6))
        btns = tk.Frame(f, bg=BG_WHITE); btns.pack(fill="x", padx=14, pady=4)
        for txt, cmd, col in [
            ("🔄  Refresh All",           self.refresh_all,  f"#{SECONDARY}"),
            ("📂  Open Excel File",        lambda: webbrowser.open(DATA_FILE.resolve().as_uri()), f"#{PRIMARY}"),
            ("🗂️  Open Invoice Folder",   lambda: webbrowser.open(INVOICE_DIR.resolve().as_uri()), f"#{PRIMARY}"),
            ("📁  Switch Data File",       self._choose_file, "#555"),
        ]:
            styled_btn(btns, txt, cmd, bg=col).pack(side="left", padx=6, pady=4)

        tk.Frame(f, bg="#dde3f8", height=2).pack(fill="x", padx=14, pady=12)
        section_label(f, "  Pending Balance Report").pack(anchor="w", padx=14, pady=(0,6))
        pending_cols = ["Invoice No","Date","Student Name","Total Amount","Paid Amount","Balance","Due Date","Payment Status"]
        self.pending_tree = self._tree(f, pending_cols)

        tk.Label(f, text="💡  Tip: Always take a regular backup of robokidy_billing_data.xlsx",
                 bg=BG_WHITE, fg="#888", font=("Segoe UI", 10, "italic")).pack(pady=8)

    # ── Actions ────────────────────────────────────────────────────────────────
    def _add_student(self):
        if not self.s_name.get().strip():
            return messagebox.showerror("Missing", "Student name is required")
        if not self.s_course.get():
            return messagebox.showerror("Missing", "Please select a course")
        if not self.s_level.get():
            return messagebox.showerror("Missing", "Please select a level")
        if not self.s_phone.get().strip():
            return messagebox.showerror("Missing", "Phone number is required")
        try:
            sid = self.store.add_student({
                "name": self.s_name.get().strip(), "parent": self.s_parent.get().strip(),
                "phone": self.s_phone.get().strip(), "email": self.s_email.get().strip(),
                "grade": self.s_grade.get().strip(), "course": self.s_course.get(),
                "level": self.s_level.get(), "batch": self.s_batch.get(),
                "joining": self.s_joining.get(), "address": self.s_address.get().strip(),
                "notes": self.s_notes.get().strip()
            })
        except PermissionError as e:
            return messagebox.showerror("File In Use ⚠️", str(e))
        messagebox.showinfo("Student Added ✓", f"Student registered successfully!\n\nStudent ID: {sid}")
        self._clear_student_form()
        self.refresh_all()

    def _clear_student_form(self):
        for w in [self.s_name, self.s_parent, self.s_phone, self.s_email,
                  self.s_grade, self.s_address, self.s_notes]:
            w.delete(0, tk.END)
        self.s_course.set(""); self.s_level.set(""); self.s_batch.set("")
        self.s_joining.delete(0, tk.END)
        self.s_joining.insert(0, datetime.now().strftime("%d-%m-%Y"))
        self.s_id_preview.config(text="Student ID: (auto-generated)")

    def _fill_student_invoice(self, _=None):
        s = self._student_lookup.get(self.i_student.get())
        if not s: return
        for e, k in [(self.i_course, "Course"), (self.i_level, "Level"), (self.i_batch, "Batch")]:
            e.config(state="normal"); e.delete(0, tk.END)
            e.insert(0, s.get(k,""))

    def _create_invoice(self):
        s = self._student_lookup.get(self.i_student.get())
        if not s:
            return messagebox.showerror("Missing", "Please select a student")
        if not self.i_fee.get().strip():
            return messagebox.showerror("Missing", "Package fee is required")
        try:
            inv_no = self.store.create_invoice({
                "date": self.i_date.get(), "student_id": s.get("Student ID"),
                "student_name": s.get("Student Name"), "course": self.i_course.get(),
                "level": self.i_level.get(), "batch": self.i_batch.get(),
                "fee": self.i_fee.get(), "disc_type": self.i_disc_type.get(),
                "disc_val": self.i_disc_val.get(), "tax_type": self.i_tax_type.get(),
                "tax_val": self.i_tax_val.get(), "paid": self.i_paid.get(),
                "mode": self.i_mode.get(), "due_date": self.i_due.get(),
                "received_by": self.i_received.get(), "notes": self.i_notes.get()
            })
        except PermissionError as e:
            return messagebox.showerror("File In Use ⚠️", str(e))
        # auto-print
        self._open_invoice_html(inv_no)
        messagebox.showinfo("Invoice Created ✓", f"Invoice created: {inv_no}\n\nThe invoice has been opened for printing.")
        self._clear_invoice_form()
        self.refresh_all()

    def _clear_invoice_form(self):
        for w in [self.i_fee, self.i_notes, self.i_course, self.i_level, self.i_batch, self.i_due]:
            w.config(state="normal"); w.delete(0, tk.END)
        self.i_disc_val.delete(0,tk.END); self.i_disc_val.insert(0,"0")
        self.i_tax_val.delete(0,tk.END);  self.i_tax_val.insert(0,"0")
        self.i_paid.delete(0,tk.END);     self.i_paid.insert(0,"0")
        self.i_disc_type.set("%"); self.i_tax_type.set("%")
        self.i_student.set(""); self.i_mode.set("Cash")
        self._total_lbl.config(text="")

    def _print_invoice(self):
        sel = self.invoice_tree.selection()
        if sel:
            inv_no = self.invoice_tree.item(sel[0], "values")[0]
        else:
            inv_no = simpledialog_ask(self, "Invoice No", "Enter Invoice Number:")
        if not inv_no: return
        self._open_invoice_html(inv_no)

    def _open_invoice_html(self, inv_no: str):
        inv = self.store.get_invoice(str(inv_no))
        if not inv:
            messagebox.showerror("Not Found", f"Invoice {inv_no} not found")
            return
        html = make_html(self.store.settings(), inv)
        path = INVOICE_DIR / f"{inv_no}.html"
        path.write_text(html, encoding="utf-8")
        webbrowser.open(path.resolve().as_uri())

    def _add_payment(self):
        inv = self.store.get_invoice(self.p_invoice.get().strip())
        if not inv:
            return messagebox.showerror("Not Found", "Invoice not found. Check the invoice number.")
        amt = self.p_amount.get().strip()
        if not amt:
            return messagebox.showerror("Missing", "Enter the paid amount")
        try:
            rno = self.store.add_payment({
                "invoice_no": inv["Invoice No"], "student_id": inv["Student ID"],
                "student_name": inv["Student Name"], "mode": self.p_mode.get(),
                "paid": amt, "received_by": self.p_received.get(),
                "notes": self.p_notes.get()
            })
        except PermissionError as e:
            return messagebox.showerror("File In Use ⚠️", str(e))
        updated = self.store.get_invoice(inv["Invoice No"])
        html = make_html(self.store.settings(), updated, receipt_no=rno)
        path = INVOICE_DIR / f"{rno}.html"
        path.write_text(html, encoding="utf-8")
        webbrowser.open(path.resolve().as_uri())
        messagebox.showinfo("Receipt Generated ✓", f"Receipt: {rno}\n\nReceipt opened for printing.")
        for w in [self.p_invoice, self.p_amount, self.p_notes]:
            w.delete(0,tk.END)
        self._inv_info_lbl.config(text="")
        self.refresh_all()

    def _choose_file(self):
        global DATA_FILE
        sel = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
            title="Choose / Create Data File"
        )
        if sel:
            DATA_FILE = Path(sel)
            self.store = ExcelStore(DATA_FILE)
            self.refresh_all()

    # ── Refresh ────────────────────────────────────────────────────────────────
    def refresh_all(self):
        try:
            self.store.update_dashboard()
        except PermissionError:
            pass  # silently skip dashboard update if file is locked
        students = self.store.get_students()
        invoices = self.store.get_invoices()
        payments = self.store.get_payments()

        # Cards
        total_b = sum(money(r.get("Total Amount")) for r in invoices)
        total_r = sum(money(r.get("Paid Amount")) for r in invoices)
        total_bl= sum(money(r.get("Balance")) for r in invoices)
        for key, val in [
            ("Students",     str(len(students))),
            ("Invoices",     str(len(invoices))),
            ("Receipts",     str(len(payments))),
            ("Total Billing",f"₹{total_b:,.0f}"),
            ("Received",     f"₹{total_r:,.0f}"),
            ("Balance",      f"₹{total_bl:,.0f}"),
        ]:
            self._dash_cards[key].config(text=val)

        # Trees
        self._insert_rows(self.student_tree, students, STUDENT_HEADERS)
        self._insert_rows(self.invoice_tree, invoices, INVOICE_HEADERS)
        self._insert_rows(self.payment_tree, payments, PAYMENT_HEADERS)

        # Dashboard recent invoices (last 10)
        short_cols = ["Invoice No","Date","Student Name","Total Amount","Paid Amount","Balance","Payment Status"]
        self._insert_rows(self._dash_inv_tree, invoices[-10:][::-1], short_cols)

        # Pending tree
        pending = [r for r in invoices if r.get("Payment Status") != "Paid"]
        pending_cols = ["Invoice No","Date","Student Name","Total Amount","Paid Amount","Balance","Due Date","Payment Status"]
        self._insert_rows(self.pending_tree, pending, pending_cols)

        # Student dropdown for invoice tab
        self._student_lookup = {
            f"{s.get('Student ID')} — {s.get('Student Name')} ({s.get('Course','')})": s
            for s in students
        }
        if hasattr(self, "i_student"):
            self.i_student["values"] = list(self._student_lookup.keys())


def simpledialog_ask(parent, title, prompt):
    """Simple modal input dialog."""
    result = [None]
    d = tk.Toplevel(parent)
    d.title(title); d.grab_set()
    d.resizable(False, False)
    tk.Label(d, text=prompt, font=("Segoe UI", 10), padx=18, pady=12).pack()
    e = ttk.Entry(d, width=28); e.pack(padx=18, pady=4); e.focus()
    def ok():
        result[0] = e.get().strip(); d.destroy()
    e.bind("<Return>", lambda _: ok())
    styled_btn(d, "OK", ok, bg=f"#{PRIMARY}").pack(pady=10)
    parent.wait_window(d)
    return result[0]


if __name__ == "__main__":
    app = BillingApp()
    app.mainloop()
