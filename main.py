"""
D365 Case Bot  (Python + Playwright)  -- New Case / Resolve Case จากไฟล์ Excel

สร้างจากการวิเคราะห์ recordings ของ session จริง (selector ทั้งหมดมาจาก data-id ที่อัดไว้)

วิธีใช้ (ปิด d365_recorder.py และปิดไฟล์ใน Excel ก่อน เพราะใช้ browser_profile/ และไฟล์เดียวกัน):

  # ทดลองกรอกฟอร์มแถวแรกโดยไม่บันทึก
  python main.py new --xlsx data.xlsx --dry-run

  # สร้างเคสจริง: ค่าเริ่มต้นทำ 1 แถว ถามยืนยันก่อน Save
  python main.py new --xlsx data.xlsx
  # หลายเคสต่อเนื่อง (0 = ทุกแถว) ไม่ถามยืนยัน ถ้าเคสไหนพลาดจะข้ามไปเคสถัดไป
  python main.py new --xlsx data.xlsx --limit 0 --yes

  # Resolve จาก Excel (แถวที่มี Record Case ID + Resolution และยังไม่ "Resolve complete")
  python main.py resolve --xlsx data.xlsx --limit 0 --yes
  # หรือเคสเดียว
  python main.py resolve --case CAS-07898-DCAQP --type Resolved --text "ดำเนินการเรียบร้อย"

คอลัมน์ "Resolve Status" ใน Excel (มี dropdown)
  Resolve complete = Resolve สำเร็จ (หรือเคสนั้น Resolved อยู่แล้ว)
  Can't Resolve    = Resolve ไม่สำเร็จ รันซ้ำได้ สคริปต์จะลองแถวนี้ใหม่
ถ้าไฟล์ยังไม่มีคอลัมน์นี้ สคริปต์จะเพิ่มให้เองตอนเริ่มรัน

หมายเหตุ: ระบบจริงเป็น PROD ค่าเริ่มต้นจึงจำกัด 1 แถวต่อครั้งและถามยืนยันก่อนบันทึก/Resolve
"""
import atexit
import argparse
import re
import sys
import csv
import time
import zipfile
import xml.etree.ElementTree as ET
from copy import copy
from pathlib import Path

import openpyxl
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from playwright.sync_api import expect, sync_playwright

# ---------------- ตั้งค่า ----------------
BASE = "https://operations-dc3-rth-prod01.crm5.dynamics.com"
APP_ID = "02d373cd-b255-ec11-8f8f-000d3ac7aa53"
LIST_URL = (
    f"{BASE}/main.aspx?appid={APP_ID}&forceUCI=1&pagetype=entitylist&etn=incident"
    "&viewid=a8ce8755-787a-f111-ab0e-6045bd219565&viewType=1039"
)
OPEN_CASE_FROM_LIST = True   # วิธี Resolve แบบกดหน้าจอ: เปิดเคสจากรายการ (ค้นหาเลขเคสแล้วคลิก) ไม่เปิดด้วยลิงก์ GUID
RECORD_URL = f"{BASE}/main.aspx?appid={APP_ID}&forceUCI=1&pagetype=entityrecord&etn=incident&id="
PROFILE_DIR = Path("browser_profile")          # ใช้ login เดียวกับ d365_recorder.py
# โฟลเดอร์เก็บภาพ/ไฟล์ผลลัพธ์ (สร้างให้อัตโนมัติ):
#   results/form_row/      form_row*.png      ภาพฟอร์ม New Case ก่อนกด Save
#   results/resolve/       resolve_*.png      ภาพหน้าต่าง Resolve ก่อนกด Resolve
#   results/resolve_fail/  resolve_fail_*.png + resolve_fail_*.txt  เมื่อ Resolve ไม่สำเร็จ
RESULTS_DIR = Path("results")


def result_path(sub, filename):
    d = RESULTS_DIR / sub
    d.mkdir(parents=True, exist_ok=True)
    return d / filename


SHEET = "Cases"                                 # ชื่อชีตที่ใช้ (ถ้าไม่พบชื่อนี้จะใช้ชีตที่เปิดค้างอยู่)

# ชื่อหัวคอลัมน์ใน Excel
COL = {
    "title": "Case Title",
    "device": "Device Number/Config ID",
    "problem": "Problem Code",
    "priority": "Priority",
    "status": "Status Reason",
    "channel": "Channel",
    "summary": "Work Order Summary",       # สมมติว่าใส่ลงช่อง Description ของเคส
    "case_id": "Record Case ID",
    "rtype": "Resolution Type",
    "resolution": "Resolution",
    "rstatus": "Resolve Status",
}
OPTIONAL_COLS = {"rstatus"}   # ไม่มีในไฟล์ก็ทำงานได้ (rstatus จะถูกเพิ่มให้เอง)

STATUS_OK = "Resolve complete"
STATUS_FAIL = "Can't Resolve"
# ใส่ในช่อง Record Case ID เมื่อกด Save แล้วแต่ยืนยันเลขเคสไม่ได้ กันสร้างซ้ำตอนรันรอบหน้า
UNCONFIRMED = "ตรวจเอง: บันทึกแล้วแต่ไม่ได้เลขเคส"

# data-id ที่ได้จากการอัด
NEW_CASE_BTN = "incident|NoRelationship|HomePageGrid|Mscrm.HomepageGrid.incident.NewRecord"
SAVE_BTN = "incident|NoRelationship|Form|Mscrm.Form.incident.Save"
RESOLVE_BTN = "incident|NoRelationship|Form|Mscrm.Form.incident.Resolve"
TITLE = "title.fieldControl-text-box-text"
DESC = "description.fieldControl-text-box-text"
TICKET = "ticketnumber.fieldControl-text-box-text"
STATUS_SELECT = "statuscode.fieldControl-option-set-select"
CHANNEL_SELECT = "caseorigincode.fieldControl-option-set-select"
RTYPE_SELECT = "resolutionType_id.fieldControl-option-set-select"
RESOLUTION_INPUT = "resolution_id.fieldControl-text-box-text"
RESOLVE_OK = "ok_id"

# ปุ่ม dialog ที่ยอมให้กดอัตโนมัติ (ตามที่เห็นในการอัด) -- ปุ่มอื่นจะไม่ถูกกด
SAFE_DIALOG_SELECTORS = ['[data-id="ignore_save"]', '[data-id="confirmButton"]', '[data-id="okButton"]']
SAFE_DIALOG_LABELS = {"Ignore and save", "Save and continue", "Continue anyway", "OK"}


def valid_case_no(x):
    return bool(re.fullmatch(r"CAS-\d+-[A-Z0-9]+", x or ""))


def short_err(e):
    lines = str(e).strip().splitlines()
    return f"{type(e).__name__}: {lines[0][:300] if lines else ''}"


# ---------------- helpers ----------------
def vis(page, data_id):
    return page.locator(f'[data-id="{data_id}"]:visible').first


DIALOG_LOG = []      # ข้อความ dialog ที่เจอ (ใช้บอกสาเหตุตอน Resolve ไม่ได้)


def dialog_text(page):
    """ข้อความใน dialog ที่เปิดอยู่ (ถ้ามี)"""
    try:
        d = page.locator('[role="alertdialog"]:visible, [role="dialog"]:visible').last
        if d.count():
            return " ".join((d.inner_text(timeout=1500) or "").split())[:200]
    except Exception:
        pass
    return ""


def clear_dialogs(page, rounds=8):
    """กด dialog ที่รู้จัก (duplicate / unsaved / memo) จนไม่เหลือ"""
    for _ in range(rounds):
        clicked = False
        for sel in SAFE_DIALOG_SELECTORS:
            btn = page.locator(sel + ":visible")
            if btn.count():
                label = (btn.first.inner_text() or "").strip()
                if label in SAFE_DIALOG_LABELS:
                    note = dialog_text(page)
                    DIALOG_LOG.append(f"{label}: {note}")
                    print(f"   dialog -> {label}" + (f"  [{note[:90]}]" if note else ""))
                    btn.first.click()
                    page.wait_for_timeout(900)
                    clicked = True
                    break
                print(f"   (พบ dialog ที่ไม่รู้จัก '{label}' ไม่กดให้ ต้องกดเอง)")
        if not clicked:
            page.wait_for_timeout(1200)
            if not any(page.locator(s + ":visible").count() for s in SAFE_DIALOG_SELECTORS):
                return


def set_text(page, data_id, value):
    el = vis(page, data_id)
    el.wait_for(timeout=20000)
    el.click()
    el.fill(value)
    el.press("Tab")           # ให้ D365 รับ change event


def set_option(page, data_id, value):
    vis(page, data_id).click()
    opt = page.get_by_role("option", name=value, exact=True).first
    try:
        opt.wait_for(timeout=10000)
    except Exception:
        names = [t.strip() for t in page.get_by_role("option").all_inner_texts() if t.strip()]
        page.keyboard.press("Escape")
        raise RuntimeError(f"ไม่พบตัวเลือก '{value}' ตัวเลือกที่มี: {names[:12]}")
    opt.click()


def _norm(t):
    return " ".join((t or "").replace("*", " ").split()).casefold()


def lookup_matches(cur, want):
    """ค่าใน lookup ตรงกับที่ต้องการไหม: ต้องเท่ากัน หรือขึ้นต้นเหมือนกัน
    ไม่ใช้ 'มีคำนี้อยู่ในข้อความ' เพราะ 'Office Hour' จะไปตรงกับ 'Non-Office Hour'"""
    c, w = _norm(cur), _norm(want)
    if not c or c in ("---", "--") or not w:
        return False
    return c == w or c.startswith(w) or w.startswith(c)


def pick_item(items, text):
    """เลือกผลการค้นหาที่ข้อความตรงที่สุด: เท่ากันก่อน แล้วขึ้นต้นเหมือนกัน แล้วตัวแรก"""
    w = _norm(text)
    texts = []
    for i in range(items.count()):
        try:
            texts.append(_norm(items.nth(i).inner_text(timeout=2000)))
        except Exception:
            texts.append("")
    for i, t in enumerate(texts):
        if t == w:
            return items.nth(i)
    for i, t in enumerate(texts):
        if t.startswith(w):
            return items.nth(i)
    return items.first


def set_lookup(page, lk, text):
    """lookup เช่น ricoh_customerasset / ricoh_problemcode / ricoh_priority"""
    tag = page.locator(f'[data-id="{lk}.fieldControl-LookupResultsDropdown_{lk}_selected_tag_text"]:visible')
    if tag.count():
        if lookup_matches(tag.first.inner_text() or "", text):
            return                                   # มีค่านี้อยู่แล้ว ไปต่อ (ไม่ลบ)
        cancel = page.locator(f'[data-id="{lk}.fieldControl-LookupResultsDropdown_{lk}_microsoftIcon_cancelButton"]:visible')
        if cancel.count():
            cancel.first.click()
    inp = vis(page, f"{lk}.fieldControl-LookupResultsDropdown_{lk}_textInputBox_with_filter_new")
    inp.wait_for(timeout=15000)
    inp.click()
    inp.fill(text)
    items = page.locator(f'span[data-id^="{lk}.fieldControl-"]:visible', has_text=text)
    items.first.wait_for(timeout=15000)
    page.wait_for_timeout(400)                       # ให้รายการผลการค้นหาโหลดครบก่อนเลือก
    pick_item(items, text).click()
    page.wait_for_timeout(500)


def ensure_login(page):
    page.goto(LIST_URL)
    if "login.microsoftonline.com" in page.url or "login.live.com" in page.url:
        print("ยังไม่ได้ล็อกอิน: ล็อกอินในหน้าต่างเบราว์เซอร์ให้เสร็จ (รอสูงสุด 5 นาที)...")
        page.wait_for_url(re.compile(r"dynamics\.com/main\.aspx"), timeout=300000)
    page.locator(f'[data-id="{NEW_CASE_BTN}"]').first.wait_for(timeout=90000)
    clear_dialogs(page, rounds=2)


def recover(page):
    """หลังเคสพลาด: กลับหน้ารายการ (ทิ้งฟอร์มที่ค้าง ไม่กดบันทึกอะไรเพิ่ม) เพื่อไปเคสถัดไป"""
    try:
        page.goto(LIST_URL)
        page.locator(f'[data-id="{NEW_CASE_BTN}"]').first.wait_for(timeout=60000)
    except Exception as e:
        print(f"   (กลับหน้ารายการไม่สำเร็จ: {short_err(e)})")


def confirm(msg, yes):
    if yes:
        return True
    return input(f"{msg} [y/N]: ").strip().lower() == "y"


# ---------------- Excel ----------------
class Sheet:
    def __init__(self, path):
        self.path = path
        try:
            open(path, "ab").close()
        except PermissionError:
            sys.exit("เขียนไฟล์ Excel ไม่ได้: ปิดไฟล์นี้ใน Excel ก่อนแล้วรันใหม่")
        self.wb = openpyxl.load_workbook(path)
        self.ws = self.wb[SHEET] if SHEET in self.wb.sheetnames else self.wb.active
        self._restore_ext_validations()
        self._index()
        missing = [v for k, v in COL.items() if k not in OPTIONAL_COLS and v not in self.head]
        if missing:
            sys.exit(f"ไม่พบคอลัมน์ในชีต '{self.ws.title}': {missing} (ยังไม่ได้แก้ไฟล์ Excel)")
        if "rstatus" not in self.idx:
            self._add_status_column()
            self._index()

    def _restore_ext_validations(self):
        """dropdown ที่ดึงรายการจากชีตอื่น (เช่น Lists!A2:A4) Excel เก็บแบบ x14 ซึ่ง openpyxl จะทิ้งตอนเซฟ
        จึงอ่านจาก XML เดิมแล้วใส่กลับเป็น dropdown ปกติ เพื่อไม่ให้ dropdown หายหลังบอทเขียนไฟล์"""
        try:
            ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
                  "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                  "x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
                  "xm": "http://schemas.microsoft.com/office/excel/2006/main"}
            with zipfile.ZipFile(self.path) as z:
                wbx = ET.fromstring(z.read("xl/workbook.xml"))
                rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
                rid = next(sh.get("{%s}id" % ns["r"]) for sh in wbx.find("m:sheets", ns)
                           if sh.get("name") == self.ws.title)
                target = next(r.get("Target") for r in rels if r.get("Id") == rid)
                part = target.lstrip("/") if target.startswith("/") else "xl/" + target
                sheet = ET.fromstring(z.read(part))
            have = {str(dv.sqref) for dv in self.ws.data_validations.dataValidation}
            restored = 0
            for x in sheet.iter("{%s}dataValidation" % ns["x14"]):
                f = x.find("x14:formula1/xm:f", ns)
                rng = x.find("xm:sqref", ns)
                if f is None or rng is None or not f.text or not rng.text or rng.text in have:
                    continue
                dv = DataValidation(type=x.get("type") or "list", formula1=f.text,
                                    allow_blank=x.get("allowBlank") == "1",
                                    showErrorMessage=x.get("showErrorMessage") == "1",
                                    showInputMessage=x.get("showInputMessage") == "1")
                dv.add(rng.text)
                self.ws.add_data_validation(dv)
                restored += 1
            if restored:
                print(f"กู้ dropdown ที่อ้างอิงชีตอื่นกลับมา {restored} รายการ (กันหายตอนบอทเขียนไฟล์)")
        except Exception as e:
            print(f"(กู้ dropdown ไม่สำเร็จ: {short_err(e)})")

    def _index(self):
        self.head = {str(c.value).strip(): c.column for c in self.ws[1] if c.value}
        self.idx = {k: self.head[v] for k, v in COL.items() if v in self.head}

    def _save(self):
        try:
            self.wb.save(self.path)
            return True
        except PermissionError:
            return False

    def _add_status_column(self):
        """เพิ่มคอลัมน์ Resolve Status ต่อท้าย พร้อม dropdown และสีเขียว/แดง"""
        ws = self.ws
        last_col = max(self.head.values())
        col = last_col + 1
        letter = get_column_letter(col)
        max_row = max(ws.max_row, 2)
        end_row = max(max_row, 500)
        for r in range(1, max_row + 1):
            src, dst = ws.cell(r, last_col), ws.cell(r, col)
            dst.font = copy(src.font)
            dst.fill = copy(src.fill)
            dst.border = copy(src.border)
            dst.alignment = copy(src.alignment)
            dst.number_format = src.number_format
        ws.cell(1, col).value = COL["rstatus"]
        ws.column_dimensions[letter].width = 20

        dv = DataValidation(type="list", formula1=f'"{STATUS_OK},{STATUS_FAIL}"', allow_blank=True,
                            showErrorMessage=True, errorTitle="ค่าไม่ถูกต้อง",
                            error=f"เลือกจากรายการ: {STATUS_OK} / {STATUS_FAIL}")
        rng = f"{letter}2:{letter}{end_row}"
        dv.add(rng)
        ws.add_data_validation(dv)
        green = PatternFill("solid", start_color="C6EFCE", end_color="C6EFCE")
        red = PatternFill("solid", start_color="FFC7CE", end_color="FFC7CE")
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{STATUS_OK}"'], fill=green))
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{STATUS_FAIL}"'], fill=red))
        if ws.auto_filter.ref:
            ws.auto_filter.ref = f"A1:{letter}{max_row}"
        if not self._save():
            sys.exit("เขียนไฟล์ Excel ไม่ได้: ปิดไฟล์นี้ใน Excel ก่อนแล้วรันใหม่")
        print(f"เพิ่มคอลัมน์ '{COL['rstatus']}' (คอลัมน์ {letter}) พร้อม dropdown ให้แล้ว")

    def get(self, row, key):
        if key not in self.idx:
            return ""
        v = self.ws.cell(row, self.idx[key]).value
        return "" if v is None else str(v).strip()

    def set(self, row, key, value):
        self.ws.cell(row, self.idx[key]).value = value
        if not self._save():     # เซฟทุกแถว กันข้อมูลหาย
            print(f"   !!! เขียน Excel ไม่ได้ (เปิดไฟล์ค้างอยู่?) จดไว้เอง: แถว {row} {COL[key]} = {value}")

    def rows(self):
        return range(2, self.ws.max_row + 1)


# ---------------- New Case ----------------
def new_case(page, sh, r, args):
    """คืน 'done' | 'skipped'  (ถ้าพลาดจะ raise เพื่อให้ผู้เรียกข้ามไปแถวถัดไป)"""
    d = {k: sh.get(r, k) for k in COL}
    print(f"\n[แถว {r}] {d['title'][:70]}")
    page.goto(LIST_URL)
    clear_dialogs(page, rounds=2)
    page.locator(f'[data-id="{NEW_CASE_BTN}"]').first.click()
    vis(page, TITLE).wait_for(timeout=60000)

    set_text(page, TITLE, d["title"])
    if d["device"]:
        set_lookup(page, "ricoh_customerasset", d["device"])
    if d["problem"]:
        set_lookup(page, "ricoh_problemcode", d["problem"])
    if d["priority"]:
        set_lookup(page, "ricoh_priority", d["priority"])
    if d["status"]:
        set_option(page, STATUS_SELECT, d["status"])
    if d["channel"]:
        set_option(page, CHANNEL_SELECT, d["channel"])
    if d["summary"]:
        set_text(page, DESC, d["summary"])

    page.screenshot(path=str(result_path("form_row", f"form_row{r}.png")))
    if args.dry_run:
        input("   dry-run: กรอกครบแล้ว (ยังไม่บันทึก) ตรวจหน้าจอ แล้วกด Enter เพื่อจบ...")
        return "skipped"
    if not confirm("   บันทึกเคสนี้ลง PROD?", args.yes):
        print("   ข้าม")
        return "skipped"

    clicked_save = False
    try:
        vis(page, SAVE_BTN).click()
        clicked_save = True
        clear_dialogs(page)
        ticket = vis(page, TICKET)
        expect(ticket).to_have_value(re.compile(r"^CAS-"), timeout=40000)
        case_no = ticket.input_value()
    except Exception:
        if clicked_save:
            # กด Save ไปแล้ว เคสอาจถูกสร้างจริง จึงทำเครื่องหมายไว้ กันสร้างซ้ำตอนรันรอบหน้า
            sh.set(r, "case_id", UNCONFIRMED)
            print(f"   !!! กด Save แล้วแต่ยืนยันเลขเคสไม่ได้ ตรวจใน D365 ว่าเคสถูกสร้างหรือไม่ "
                  f"(ใส่ '{UNCONFIRMED}' ในช่อง Record Case ID แล้ว)")
        raise
    print(f"   สร้างแล้ว: {case_no}")
    sh.set(r, "case_id", case_no)
    clear_dialogs(page, rounds=3)
    return "done"


# ---------------- Resolve ----------------
def _case_guid_once(page, ticket):
    return page.evaluate(
        """async (t) => {
            const q = `/api/data/v9.0/incidents?$select=incidentid&$filter=ticketnumber eq '${t}'`;
            const r = await fetch(q, {headers: {'Accept': 'application/json'}});
            if (!r.ok) return null;
            const j = await r.json();
            return j.value && j.value[0] ? j.value[0].incidentid : null;
        }""",
        ticket,
    )


def case_guid(page, ticket):
    """หา GUID ของเคส (ลอง 3 รอบ เพราะ API อ่านพลาดเป็นครั้งคราวได้)"""
    for i in range(3):
        try:
            g = _case_guid_once(page, ticket)
        except Exception:
            g = None
        if g:
            return g
        page.wait_for_timeout(2000)
    return None


def case_state(page, guid):
    """statecode ของเคส: 0=Active, 1=Resolved, 2=Cancelled (อ่านอย่างเดียว)"""
    return page.evaluate(
        """async (id) => {
            const q = `/api/data/v9.0/incidents(${id})?$select=statecode`;
            const r = await fetch(q, {headers: {'Accept': 'application/json'}});
            if (!r.ok) return null;
            const j = await r.json();
            return j.statecode;
        }""",
        guid,
    )


RESOLVE_SELECTORS = [
    f'[data-id="{RESOLVE_BTN}"]',
    '[data-id^="incident|NoRelationship|Form|"][data-id$=".Resolve"]',
]


def wait_not_processing(page, secs=60):
    """รอให้ overlay 'Processing...' ของ D365 หาย (ระหว่างนั้นหน้าจอถูกบัง กดอะไรไม่ได้)"""
    pat = re.compile(r"^\s*Processing\.\.\.\s*$", re.I)
    announced = False
    for _ in range(int(secs)):
        try:
            loc = page.get_by_text(pat)
            if loc.count() == 0 or not loc.first.is_visible():
                return True
        except Exception:
            return True
        if not announced:
            print("   รอระบบประมวลผล (Processing...)")
            announced = True
        page.wait_for_timeout(1000)
    return False


def wait_form_ready(page):
    """รอ command bar โหลด (ปุ่ม Save ขึ้น) แล้วเคลียร์ dialog เช่น Active Memo ที่มักโผล่หลังโหลดเสร็จ"""
    try:
        vis(page, SAVE_BTN).wait_for(timeout=45000)
    except Exception:
        pass
    page.wait_for_timeout(1500)
    wait_not_processing(page, 45)
    clear_dialogs(page, rounds=4)


def _visible_resolve_button(page):
    for sel in RESOLVE_SELECTORS:
        loc = page.locator(sel + ":visible")
        if loc.count():
            return loc.first
    by_name = page.get_by_role("button", name=re.compile(r"^Resolve Case"))
    if by_name.count():
        return by_name.first
    return None


def find_resolve_button(page, timeout=30):
    """หาปุ่ม Resolve Case หลายวิธี: data-id -> ชื่อปุ่ม -> เปิดเมนู ⋮ (More commands)"""
    end = time.time() + timeout
    while time.time() < end:
        b = _visible_resolve_button(page)
        if b is not None:
            return b
        page.wait_for_timeout(1000)
    try:
        page.locator('button[aria-label^="More commands"]:visible').first.click(timeout=3000)
        page.wait_for_timeout(800)
        b = _visible_resolve_button(page)
        if b is not None:
            print("   พบปุ่ม Resolve Case ในเมนู ⋮")
            return b
        page.keyboard.press("Escape")
    except Exception:
        pass
    return None


def wait_resolve_dialog(page, timeout=30):
    """คืน None เมื่อหน้าต่าง Resolve ขึ้นแล้ว หรือข้อความสาเหตุ ถ้าระบบไม่ยอมเปิด (เช่น activity ยังไม่ปิด)"""
    end = time.time() + timeout
    while time.time() < end:
        if page.locator(f'[data-id="{RTYPE_SELECT}"]:visible').count():
            return None
        clear_dialogs(page, rounds=1)
        if page.locator(f'[data-id="{RTYPE_SELECT}"]:visible').count():
            return None
    msgs = [m for m in DIALOG_LOG if m and "Ignore and save" not in m]
    left = dialog_text(page)
    if left:
        msgs.append(left)
    return "ระบบไม่เปิดหน้าต่าง Resolve: " + (" | ".join(msgs)[:250] if msgs else "ไม่มีข้อความจากระบบ")


def dump_resolve_debug(page, ticket, extra=""):
    """เซฟภาพ + รายการปุ่มบน command bar ตอน Resolve พลาด เพื่อดูสาเหตุจริง"""
    try:
        page.screenshot(path=str(result_path("resolve_fail", f"resolve_fail_{ticket}.png")))
        items = page.evaluate(
            r"""() => Array.from(document.querySelectorAll('[data-id^="incident|NoRelationship|Form|"]'))
                .filter(e => e.offsetParent !== null)
                .map(e => e.getAttribute('data-id') + ' | ' +
                          (e.getAttribute('aria-label') || e.innerText || '').replace(/\s+/g, ' ').slice(0, 80))"""
        )
        body = "\n".join(items) + f"\n\nURL: {page.url}\nDialogs: " + " || ".join(DIALOG_LOG) + (f"\n\nสาเหตุ: {extra}" if extra else "")
        result_path("resolve_fail", f"resolve_fail_{ticket}.txt").write_text(body, encoding="utf-8")
        print(f"   เซฟรายละเอียดไว้ที่ {RESULTS_DIR / 'resolve_fail'} (resolve_fail_{ticket}.png / .txt)")
    except Exception as e:
        print(f"   (เซฟรายละเอียดไม่สำเร็จ: {short_err(e)})")


def wait_enabled(page, btn, secs, stable=2):
    """รอให้ปุ่มกดได้ (ไม่เทา) ต่อเนื่อง `stable` ครั้ง สูงสุด secs วินาที"""
    ok = 0
    for _ in range(int(secs)):
        try:
            ok = ok + 1 if btn.is_enabled(timeout=1500) else 0
        except Exception:
            ok = 0
        if ok >= stable:
            return True
        page.wait_for_timeout(1000)
    return False


def form_is_dirty(page):
    """ฟอร์มมีค่าที่ยังไม่ Save หรือไม่ (ผ่าน Xrm และ/หรือหัวฟอร์มที่ขึ้น '- Unsaved')"""
    dirty = False
    try:
        v = page.evaluate("() => { try { return Xrm.Page.data.entity.getIsDirty(); } catch (e) { return null; } }")
        dirty = bool(v)
    except Exception:
        pass
    if not dirty:
        try:
            dirty = page.locator(':text-matches("Unsaved", "i"):visible').count() > 0
        except Exception:
            pass
    return dirty


def save_form(page, timeout=40):
    """กด Save แล้วรอจนฟอร์มไม่ Unsaved (เคลียร์ dialog เคสซ้ำ/Unsaved changes ระหว่างทาง) คืน True เมื่อ Saved"""
    vis(page, SAVE_BTN).click(timeout=10000)
    for _ in range(int(timeout // 2)):
        clear_dialogs(page, rounds=1)
        page.wait_for_timeout(1000)
        if not form_is_dirty(page):
            page.wait_for_timeout(1500)           # รอให้ command bar อัปเดตหลัง Saved
            wait_not_processing(page, 45)
            return True
    return False


def ensure_resolve_enabled(page, btn, allow_save=True):
    """ปุ่ม Resolve Case เทา = ฟอร์ม Unsaved/โหลดไม่เสร็จ: รอ -> Save -> รีโหลด แล้วเช็กอีกครั้ง"""
    if wait_enabled(page, btn, 10):
        return btn
    if not allow_save:
        return None
    print("   ปุ่ม Resolve Case ยังเทา (ฟอร์มอาจ Unsaved) -> กด Save ก่อน")
    vis(page, SAVE_BTN).click()
    clear_dialogs(page, rounds=6)             # Ignore and save / Save and continue
    if wait_enabled(page, btn, 20):
        return btn
    print("   ยังเทาอยู่ -> รีโหลดหน้าแล้วลองใหม่")
    page.reload()
    vis(page, TITLE).wait_for(timeout=60000)
    wait_form_ready(page)
    btn = find_resolve_button(page)
    if btn is not None and wait_enabled(page, btn, 20):
        return btn
    return None


def error_texts(page):
    """ข้อความ error/แจ้งเตือนที่ระบบแสดง (dialog ที่เคยกดผ่าน + alert ที่ค้างอยู่)"""
    out = list(DIALOG_LOG)
    try:
        out += [" ".join(t.split())[:200] for t in
                page.locator('[role="alert"]:visible, [role="alertdialog"]:visible').all_inner_texts()]
    except Exception:
        pass
    return [t for t in dict.fromkeys(out) if t]


def open_activities(page, guid):
    """activity ที่ยังเปิดอยู่ของเคส (อ่านอย่างเดียว) ใช้บอกสาเหตุเมื่อ Resolve ไม่ผ่าน"""
    try:
        return page.evaluate(
            r"""async (id) => {
                const q = `/api/data/v9.0/activitypointers?$select=subject,activitytypecode` +
                          `&$filter=_regardingobjectid_value eq ${id} and statecode eq 0&$top=10`;
                const r = await fetch(q, {headers: {'Accept': 'application/json'}});
                if (!r.ok) return null;
                const j = await r.json();
                return (j.value || []).map(a => `${a.activitytypecode}: ${(a.subject || '').slice(0, 50)}`);
            }""",
            guid,
        )
    except Exception:
        return None


def wait_resolved(page, guid, secs=120):
    """Resolve ใน D365 ใช้เวลาประมวลผลหลายวินาที (ขึ้น 'Processing...') จึงถามสถานะซ้ำจนเป็น Resolved หรือหมดเวลา"""
    state, none_count = None, 0
    for i in range(int(secs // 2)):
        page.wait_for_timeout(2000)
        state = case_state(page, guid)
        if state == 1:
            return 1
        if state not in (0, None):
            return state                      # เช่น 2 = Cancelled
        none_count = none_count + 1 if state is None else 0
        if none_count >= 3:
            return None                       # อ่านสถานะจาก API ไม่ได้
        if i == 0:
            print("   รอระบบประมวลผล Resolve...")
    return state


def fail_with_diagnostics(page, ticket, guid, headline):
    """Resolve ไม่ผ่าน: รวมข้อความจากระบบ + activity ที่ยังเปิด เป็นเหตุผล แล้วเซฟภาพ/ไฟล์ไว้"""
    msgs = error_texts(page)
    acts = open_activities(page, guid)
    reason = headline
    if msgs:
        reason += " | ข้อความจากระบบ: " + " / ".join(msgs)[:250]
    if acts:
        reason += f" | มี activity ที่ยังเปิด {len(acts)} รายการ: " + "; ".join(acts)[:150]
    dump_resolve_debug(page, ticket, reason)
    return "failed", reason


def open_case(page, ticket, guid):
    """เปิดเคสแบบที่คนทำ: ค้นหาเลขเคสในหน้ารายการแล้วคลิกเปิด ถ้าไม่สำเร็จค่อยถอยไปใช้ลิงก์ GUID
    คืน 'list' หรือ 'guid' (บอกว่าใช้ทางไหนเปิด)"""
    if OPEN_CASE_FROM_LIST:
        try:
            page.goto(LIST_URL)
            page.locator(f'[data-id="{NEW_CASE_BTN}"]').first.wait_for(timeout=60000)
            clear_dialogs(page, rounds=2)
            box = page.locator('input[placeholder*="Filter by keyword" i], '
                               'input[aria-label*="Filter by keyword" i]').first
            box.wait_for(timeout=15000)
            box.click()
            box.fill(ticket)
            box.press("Enter")
            row = page.locator('[role="row"]', has_text=ticket).first
            row.wait_for(timeout=20000)
            links = row.locator('a, [role="link"]')
            (links.first if links.count() else row.locator('[role="gridcell"]', has_text=ticket).first).click()
            page.wait_for_url(re.compile(re.escape(guid), re.I), timeout=20000)   # ยืนยันว่าเปิดถูกเคส
            vis(page, TITLE).wait_for(timeout=60000)
            print("   เปิดเคสจากรายการ")
            return "list"
        except Exception as e:
            print(f"   (เปิดจากรายการไม่ได้: {short_err(e)} -> ใช้ลิงก์ GUID แทน)")
    page.goto(RECORD_URL + guid)
    vis(page, TITLE).wait_for(timeout=60000)
    return "guid"


def resolve_case(page, ticket, rtype, text, args):
    """คืน ('done'|'failed'|'skipped', เหตุผล)  ถ้ามี exception ผู้เรียกจะนับเป็น failed"""
    print(f"\n[Resolve] {ticket}")
    if not valid_case_no(ticket):
        return "failed", "รูปแบบเลขเคสไม่ถูกต้อง"
    if BASE.split("//", 1)[1] not in page.url:        # หน้าเบราว์เซอร์ต้องอยู่ใน D365 ก่อนอ่าน API
        ensure_login(page)
    guid = case_guid(page, ticket)            # อ่านอย่างเดียว ผ่าน session ของหน้าเว็บเอง
    if not guid:
        return "failed", "ไม่พบเคสนี้ (หรืออ่าน API ไม่ได้)"
    st = case_state(page, guid)
    if st == 1:
        print("   เคสนี้ Resolved อยู่แล้ว")
        return "done", "Resolved อยู่แล้ว"
    if st == 2:
        return "failed", "เคสถูก Cancel ไปแล้ว"

    open_case(page, ticket, guid)             # เปิดจากรายการ (สำรอง: ลิงก์ GUID)
    wait_form_ready(page)                     # รอ command bar + เคลียร์ Active Memo

    btn = None
    for attempt in range(2):
        btn = find_resolve_button(page)
        if btn is not None:
            break
        if attempt == 0:
            print("   ยังไม่เห็นปุ่ม Resolve Case ลองรีโหลดหน้าอีกรอบ...")
            page.reload()
            vis(page, TITLE).wait_for(timeout=60000)
            wait_form_ready(page)
    if btn is None:
        dump_resolve_debug(page, ticket)
        return "failed", f"ไม่พบปุ่ม Resolve Case (ดูใน {RESULTS_DIR / 'resolve_fail'})"

    if not args.dry_run and not confirm(f"   Resolve {ticket} ด้วย [{rtype}] '{text[:40]}'?", args.yes):
        print("   ข้าม")
        return "skipped", ""

    # ฟอร์ม Unsaved (กฎ/สคริปต์ในฟอร์มแก้ค่าตอนโหลด) ทำให้ Resolve ไม่ทำงาน: Save ก่อนเสมอ
    if form_is_dirty(page):
        if args.dry_run:
            print("   dry-run: ฟอร์ม Unsaved (ตอนรันจริงจะกด Save ก่อน Resolve)")
        else:
            print("   ฟอร์ม Unsaved -> กด Save ก่อน Resolve")
            if not save_form(page):
                print("   ! Save แล้วฟอร์มยังเป็น Unsaved จะลองกด Resolve ต่อ")
            nb = find_resolve_button(page)
            if nb is not None:
                btn = nb

    btn = ensure_resolve_enabled(page, btn, allow_save=not args.dry_run)
    if btn is None:
        if args.dry_run:
            print("   dry-run: ปุ่ม Resolve Case เทา (ไม่ได้กด Save ให้)")
            return "skipped", ""
        return fail_with_diagnostics(page, ticket, guid, "ปุ่ม Resolve Case เทา (กดไม่ได้) แม้ Save/รีโหลดแล้ว")

    DIALOG_LOG.clear()
    reason = None
    for attempt in range(2):
        clicked = False
        wait_not_processing(page, 30)         # ถ้า overlay Processing... บังอยู่ รอให้หายก่อน
        for _ in range(2):                    # ถ้ามี dialog (เช่น Active Memo) โผล่มาบัง ให้เคลียร์แล้วกดใหม่
            try:
                btn.click(timeout=8000)
                clicked = True
                break
            except Exception:
                clear_dialogs(page, rounds=3)
        if clicked:
            reason = wait_resolve_dialog(page, timeout=20)   # Ignore and save -> หน้าต่าง Resolve
        else:
            reason = "กดปุ่ม Resolve Case ไม่ได้ (ปุ่มเทาหรือมีหน้าต่างบังอยู่)"
        if reason is None:
            break
        if attempt == 0 and not args.dry_run and form_is_dirty(page):   # dry-run ห้าม Save
            print("   หน้าต่าง Resolve ไม่เปิด และฟอร์ม Unsaved -> Save แล้วกดใหม่")
            save_form(page)
            nb = find_resolve_button(page)
            if nb is not None:
                btn = nb
                continue
        break
    if reason and args.dry_run and form_is_dirty(page):
        print("   dry-run: หน้าต่าง Resolve ไม่เปิดเพราะฟอร์ม Unsaved (ตอนรันจริงจะกด Save ก่อน)")
        return "skipped", ""
    if reason:
        return fail_with_diagnostics(page, ticket, guid, reason)
    set_option(page, RTYPE_SELECT, rtype)
    set_text(page, RESOLUTION_INPUT, text)
    page.screenshot(path=str(result_path("resolve", f"resolve_{ticket}.png")))
    if args.dry_run:
        input("   dry-run: กรอกครบแล้ว (ยังไม่กด Resolve) กด Enter เพื่อจบ...")
        return "skipped", ""

    DIALOG_LOG.clear()
    vis(page, RESOLVE_OK).click()
    clear_dialogs(page, rounds=3)
    try:
        page.locator(f'[data-id="{RESOLVE_OK}"]:visible').first.wait_for(state="hidden", timeout=30000)
    except Exception:
        return fail_with_diagnostics(page, ticket, guid, "กด Resolve แล้วแต่หน้าต่างไม่ปิด (ระบบน่าจะปฏิเสธ)")
    state = wait_resolved(page, guid)
    if state == 1:
        print("   Resolve แล้ว (ยืนยันจากระบบ: สถานะ Resolved)")
        wait_not_processing(page, 30)         # ให้ระบบเสร็จจริงก่อนไปเคสถัดไป
        return "done", ""
    if state is None:
        print("   กด Resolve แล้ว แต่อ่านสถานะจากระบบไม่ได้ ตรวจในหน้าเว็บอีกครั้ง")
        return "done", "ยืนยันสถานะจากระบบไม่ได้"
    return fail_with_diagnostics(page, ticket, guid, f"กด Resolve แล้วรอ 120 วินาที แต่ระบบยังรายงาน statecode={state} (ไม่ใช่ Resolved) รอบถัดไปจะเช็กสถานะซ้ำ")


# ---------------- Resolve ผ่าน API (ทางเลือก: ใช้ด้วย --api) ----------------
# เรียก action CloseIncident ของ Dataverse (ตัวเดียวกับที่ปุ่ม Resolve Case ในหน้าเว็บเรียก) ผ่าน session ของหน้าเว็บเอง
# ไม่ต้องคลิกปุ่ม จึงไม่ติดปัญหาฟอร์ม Unsaved / Processing... / ปุ่มเทา และได้ข้อความ error จริงจากเซิร์ฟเวอร์
# ข้อจำกัด: ข้ามกฎ/สคริปต์ฝั่งหน้าฟอร์ม (กฎฝั่งเซิร์ฟเวอร์ เช่น activity ที่ยังไม่ปิด ยังทำงานตามปกติ)
SUBJECT_MAX = 200          # ความยาวสูงสุดของ subject ใน incidentresolution

STATUS_OPTIONS_JS = r"""async () => {
    const q = `${location.origin}/api/data/v9.2/EntityDefinitions(LogicalName='incident')` +
              `/Attributes(LogicalName='statuscode')/Microsoft.Dynamics.CRM.StatusAttributeMetadata` +
              `?$select=LogicalName&$expand=OptionSet($select=Options)`;
    const r = await fetch(q, {headers: {'Accept': 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0'},
                              credentials: 'include'});
    if (!r.ok) return {error: 'HTTP ' + r.status};
    const j = await r.json();
    const opts = (j.OptionSet && j.OptionSet.Options) || [];
    return {options: opts.map(o => ({
        value: o.Value, state: o.State,
        label: (o.Label && o.Label.UserLocalizedLabel && o.Label.UserLocalizedLabel.Label) || ''}))};
}"""

CLOSE_JS = r"""async ([guid, status, subject]) => {
    const body = {
        IncidentResolution: {
            "@odata.type": "Microsoft.Dynamics.CRM.incidentresolution",
            "incidentid@odata.bind": `/incidents(${guid})`,
            subject: subject
        },
        Status: status
    };
    const r = await fetch(`${location.origin}/api/data/v9.2/CloseIncident`, {
        method: 'POST',
        headers: {'Accept': 'application/json', 'Content-Type': 'application/json; charset=utf-8',
                  'OData-MaxVersion': '4.0', 'OData-Version': '4.0'},
        credentials: 'include',
        body: JSON.stringify(body)
    });
    let msg = '';
    if (!r.ok) {
        try { const j = await r.json(); msg = (j.error && j.error.message) || JSON.stringify(j); }
        catch (e) { try { msg = await r.text(); } catch (e2) { msg = ''; } }
    }
    return {ok: r.ok, status: r.status, message: msg};
}"""

_STATUS_CACHE = {}


def resolved_status_value(page, label):
    """แปลงชื่อ Resolution Type (เช่น 'Resolved') เป็นค่า statuscode ของสถานะ Resolved อ่านจาก metadata ของระบบ
    คืน (ค่า, None) หรือ (None, ข้อความ error)"""
    if "opts" not in _STATUS_CACHE:
        res = page.evaluate(STATUS_OPTIONS_JS)
        if not res or res.get("error"):
            return None, f"อ่านรายการ Resolution Type จากระบบไม่ได้ ({(res or {}).get('error', 'ไม่มีผลลัพธ์')})"
        _STATUS_CACHE["opts"] = res.get("options", [])
    resolved = [o for o in _STATUS_CACHE["opts"] if o.get("state") == 1]     # statecode 1 = Resolved
    want = " ".join((label or "").split()).casefold()
    for o in resolved:
        if " ".join((o.get("label") or "").split()).casefold() == want:
            return o["value"], None
    names = [o.get("label") for o in resolved]
    return None, f"ไม่พบ Resolution Type '{label}' ในระบบ ตัวเลือกที่มี: {names}"


def resolve_case_api(page, ticket, rtype, text, args):
    """คืน ('done'|'failed'|'skipped', เหตุผล)  ทำงานเหมือน resolve_case แต่ไม่คลิกหน้าเว็บ"""
    print(f"\n[Resolve/API] {ticket}")
    if not valid_case_no(ticket):
        return "failed", "รูปแบบเลขเคสไม่ถูกต้อง"
    if BASE.split("//", 1)[1] not in page.url:
        ensure_login(page)
    guid = case_guid(page, ticket)
    if not guid:
        return "failed", "ไม่พบเคสนี้ (หรืออ่าน API ไม่ได้)"
    st = case_state(page, guid)
    if st == 1:
        print("   เคสนี้ Resolved อยู่แล้ว")
        return "done", "Resolved อยู่แล้ว"
    if st == 2:
        return "failed", "เคสถูก Cancel ไปแล้ว"

    subject = (text or "").strip()
    if not subject:
        return "failed", "ช่อง Resolution ว่าง"
    if len(subject) > SUBJECT_MAX:
        return "failed", f"Resolution ยาว {len(subject)} ตัวอักษร เกิน {SUBJECT_MAX} (ระบบไม่รับ) ตัดให้สั้นลงใน Excel"
    value, err = resolved_status_value(page, rtype)
    if err:
        return "failed", err

    print(f"   Resolution Type '{rtype}' -> statuscode {value} | Resolution: {subject[:60]}")
    if args.dry_run:
        print("   dry-run: ยังไม่ส่งคำสั่ง Resolve (ตรวจค่าข้างบนว่าถูกต้อง)")
        return "skipped", ""
    if not confirm(f"   Resolve {ticket} ผ่าน API ด้วย [{rtype}] '{subject[:40]}'?", args.yes):
        print("   ข้าม")
        return "skipped", ""

    res = page.evaluate(CLOSE_JS, [guid, value, subject])
    if not res or not res.get("ok"):
        status = (res or {}).get("status")
        msg = " ".join(((res or {}).get("message") or "").split())[:300]
        return "failed", f"ระบบปฏิเสธ (HTTP {status}): {msg or 'ไม่มีข้อความ'}"
    state = wait_resolved(page, guid)
    if state == 1:
        print("   Resolve แล้ว (ยืนยันจากระบบ: สถานะ Resolved)")
        return "done", ""
    if state is None:
        print("   ส่งคำสั่งแล้ว แต่อ่านสถานะจากระบบไม่ได้ ตรวจในหน้าเว็บอีกครั้ง")
        return "done", "ยืนยันสถานะจากระบบไม่ได้"
    return "failed", f"ส่งคำสั่งแล้วแต่ระบบยังรายงาน statecode={state} (ไม่ใช่ Resolved)"


def log_result(ticket, outcome, why):
    """บันทึกผล Resolve ทุกเคสลง results/resolve_log.csv (ไว้ดูย้อนหลังว่าพลาดเพราะอะไร)"""
    try:
        f = result_path("", "resolve_log.csv")
        new = not f.exists()
        with open(f, "a", newline="", encoding="utf-8-sig") as fh:
            w = csv.writer(fh)
            if new:
                w.writerow(["เวลา", "เลขเคส", "ผล", "เหตุผล"])
            w.writerow([time.strftime("%Y-%m-%d %H:%M:%S"), ticket, outcome, why])
    except Exception as e:
        print(f"   (บันทึก log ไม่สำเร็จ: {short_err(e)})")


def resolve_with_retry(page, ticket, rtype, text, args, tries=2):
    """Resolve 1 เคส ถ้าพลาดจะโหลดหน้าใหม่แล้วลองอีก (รอบถัดไปเช็กสถานะก่อน
    ถ้าเคสถูก Resolve ไปแล้วจะนับเป็นสำเร็จ) dry-run และ skipped ไม่ลองซ้ำ"""
    outcome, why = "failed", ""
    for i in range(tries):
        try:
            fn = resolve_case_api if getattr(args, "api", False) else resolve_case
            outcome, why = fn(page, ticket, rtype, text, args)
        except KeyboardInterrupt:
            raise
        except Exception as e:
            outcome, why = "failed", short_err(e)
            if not getattr(args, "api", False):
                dump_resolve_debug(page, ticket, why)
        if outcome != "failed" or args.dry_run:
            return outcome, why
        print(f"   ไม่สำเร็จ ({why})")
        recover(page)                               # กลับหน้ารายการ เริ่มรอบใหม่จากหน้าสะอาด
        if i + 1 < tries:
            print(f"   ลองใหม่ครั้งที่ {i + 2}/{tries} ...")
    return outcome, why


# ---------------- main ----------------
_menu_mode = False
_next_menu_command = None
_menu_xlsx = None


def select_command():
    """คืน 'new' | 'resolve' (2.1 กดผ่านหน้าจอ) | 'resolve_api' (2.2 API) | None (ออก)"""
    print("\nเลือกการทำงาน")
    print("1. new case")
    print("2. resolve case")
    print("   2.1 Resolve (กดผ่านหน้าจอ)")
    print("   2.2 Resolve API")
    print("0. ออกจากโปรแกรม")
    while True:
        choice = input("เลือกเมนูดังกล่าว: ").strip()
        if choice == "1":
            return "new"
        if choice == "2.1":
            return "resolve"
        if choice == "2.2":
            return "resolve_api"
        if choice == "2":
            while True:
                sub = input("เลือกวิธี Resolve  1 = กดผ่านหน้าจอ (2.1)  2 = API (2.2): ").strip()
                if sub in ("1", "2.1"):
                    return "resolve"
                if sub in ("2", "2.2"):
                    return "resolve_api"
                print("กรุณาเลือก 1 หรือ 2")
        if choice == "0":
            return None
        print("กรุณาเลือกตัวเลือก")


def apply_menu_choice(args, choice):
    """แปลงตัวเลือกจากเมนูเป็น args.cmd และ args.api"""
    if choice == "resolve_api":
        args.cmd, args.api = "resolve", True
    elif choice == "resolve":
        args.cmd, args.api = "resolve", False
    else:
        args.cmd = choice


def wait_for_exit():
    """รอจนกว่าผู้ใช้จะกด q ก่อนออกจากโปรแกรม"""
    while True:
        choice = input("\nกด q แล้ว Enter เพื่อออกจากโปรแกรม: ").strip().lower()
        if choice == "q":
            return
        print("กรุณากด q เพื่อออกจากโปรแกรม")


def main():
    global _menu_mode, _next_menu_command, _menu_xlsx

    ap = argparse.ArgumentParser(description="D365 Case Bot")
    ap.add_argument("cmd", nargs="?", choices=["new", "resolve"],
                    help="คำสั่งที่ต้องการทำ ถ้าไม่ระบุจะแสดงเมนู")
    ap.add_argument("--xlsx")
    ap.add_argument("--case")
    ap.add_argument("--type", default="Resolved")
    ap.add_argument("--text")
    ap.add_argument("--limit", type=int, default=0,
                    help="จำนวนแถวสูงสุดที่ลอง (0 = ทุกแถว) ค่าเริ่มต้นทำทุกแถว")
    ap.add_argument("--row", type=int, help="ทำเฉพาะแถวนี้")
    ap.add_argument("--dry-run", action="store_true", help="กรอกแต่ไม่บันทึก")
    ap.add_argument("--yes", action="store_true", help="ใช้ต่อเนื่องเพื่อความเข้ากันได้ (ค่าเริ่มต้นไม่ถามอยู่แล้ว)")
    ap.add_argument("--api", action="store_true",
                    help="Resolve ผ่าน API (CloseIncident) แทนการคลิกหน้าเว็บ ใช้ --dry-run ดูค่าก่อนได้")
    ap.add_argument("--ask", action="store_true", help="ถามยืนยันก่อนบันทึกหรือ Resolve")
    args = ap.parse_args()
    args.yes = not args.ask

    if _next_menu_command is not None:
        apply_menu_choice(args, _next_menu_command)
        _next_menu_command = None
        args.xlsx = _menu_xlsx
    elif args.cmd is None:
        _menu_mode = True
        choice = select_command()
        if choice is None:
            return
        apply_menu_choice(args, choice)
        if not args.xlsx:
            path = input("ไฟล์ Excel [data.xlsx]: ").strip()
            args.xlsx = path or "data.xlsx"
        _menu_xlsx = args.xlsx

    if args.cmd == "resolve" and not (args.xlsx or (args.case and args.text)):
        sys.exit("resolve ต้องมี --xlsx หรือ --case ... --text ...")
    if args.cmd == "new" and not args.xlsx:
        sys.exit("new ต้องมี --xlsx")

    sh = Sheet(args.xlsx) if args.xlsx else None
    results = []          # (แถว, ผล, เหตุผล)

    def on_dialog(d):     # ตอนข้ามเคสที่พลาด ยอมทิ้งฟอร์มที่ค้างโดยไม่บันทึก
        try:
            d.accept() if d.type == "beforeunload" else d.dismiss()
        except Exception:
            pass

    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            str(PROFILE_DIR), headless=False, no_viewport=True, args=["--start-maximized"])
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        page.set_default_timeout(20000)
        page.on("dialog", on_dialog)
        try:
            ensure_login(page)

            if args.cmd == "resolve" and args.case:
                outcome, why = resolve_with_retry(page, args.case, args.type, args.text, args)
                log_result(args.case, outcome, why)
                results.append((0, outcome, why))
            else:
                attempted = 0
                for r in ([args.row] if args.row else sh.rows()):
                    if args.limit and attempted >= args.limit:
                        break

                    if args.cmd == "new":
                        if not (sh.get(r, "title") and not sh.get(r, "case_id")):
                            continue
                        attempted += 1
                        try:
                            outcome, why = new_case(page, sh, r, args), ""
                        except KeyboardInterrupt:
                            raise
                        except Exception as e:
                            outcome, why = "failed", short_err(e)
                            print(f"   ไม่สำเร็จ: {why}  -> ข้ามไปเคสถัดไป")
                            recover(page)
                        results.append((r, outcome, why))
                        if args.dry_run:
                            break

                    else:  # resolve จาก Excel
                        no, res = sh.get(r, "case_id"), sh.get(r, "resolution")
                        if not (no and res) or sh.get(r, "rstatus") == STATUS_OK:
                            continue
                        if not valid_case_no(no):
                            print(f"\n[แถว {r}] ข้าม: Record Case ID '{no}' ไม่ใช่เลขเคส")
                            results.append((r, "skipped", f"Record Case ID '{no}' ไม่ใช่เลขเคส"))
                            continue
                        attempted += 1
                        rtype = sh.get(r, "rtype") or "Resolved"
                        outcome, why = resolve_with_retry(page, no, rtype, res, args)
                        if not args.dry_run:
                            log_result(no, outcome, why)
                        if args.dry_run:
                            print(f"   dry-run: ผล = {outcome} {why}  (ไม่เขียนลง Excel)")
                        elif outcome == "done":
                            sh.set(r, "rstatus", STATUS_OK)
                        elif outcome == "failed":
                            sh.set(r, "rstatus", STATUS_FAIL)
                            print(f"   ไม่สำเร็จ: {why}  -> ทำเครื่องหมาย '{STATUS_FAIL}' แล้วข้ามไปเคสถัดไป")
                        results.append((r, outcome, why))
                        if args.dry_run:
                            break
        finally:
            ctx.close()

    ok = [x for x in results if x[1] == "done"]
    bad = [x for x in results if x[1] == "failed"]
    skip = [x for x in results if x[1] == "skipped"]
    print("\n=== สรุป ===")
    print(f"สำเร็จ {len(ok)} | ไม่สำเร็จ {len(bad)} | ข้าม {len(skip)}")
    for r, _, why in bad:
        print(f"  ไม่สำเร็จ แถว {r}: {why}")
    if not results and sh is not None:
        print("ไม่มีแถวที่ต้องทำ: new = ต้องมี Case Title และเว้น Record Case ID ว่าง | "
              f"resolve = ต้องมี Record Case ID และ Resolution และสถานะยังไม่ใช่ '{STATUS_OK}'")
    if _menu_mode:
        print()
        _next_menu_command = select_command()
        if _next_menu_command is not None:
            main()


if __name__ == "__main__":
    main()