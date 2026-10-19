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
import argparse
import re
import sys
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
RECORD_URL = f"{BASE}/main.aspx?appid={APP_ID}&forceUCI=1&pagetype=entityrecord&etn=incident&id="
PROFILE_DIR = Path("browser_profile")          # ใช้ login เดียวกับ d365_recorder.py
SHEET = None                                    # None = ชีตแรก

# ชื่อหัวคอลัมน์ใน Excel
COL = {
    "title": "Case Title",
    "device": "Device Number/Config ID",
    "problem": "Problem Code",
    "job": "Job Category",
    "priority": "Priority",
    "status": "Status Reason",
    "channel": "Channel",
    "summary": "Work Order Summary",       # สมมติว่าใส่ลงช่อง Description ของเคส
    "case_id": "Record Case ID",
    "rtype": "Resolution Type",
    "resolution": "Resolution",
    "rstatus": "Resolve Status",
}
OPTIONAL_COLS = {"job", "rstatus"}   # ไม่มีในไฟล์ก็ทำงานได้ (rstatus จะถูกเพิ่มให้เอง)

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
    return f"{type(e).__name__}: {lines[0][:150] if lines else ''}"


# ---------------- helpers ----------------
def vis(page, data_id):
    return page.locator(f'[data-id="{data_id}"]:visible').first


def clear_dialogs(page, rounds=8):
    """กด dialog ที่รู้จัก (duplicate / unsaved / memo) จนไม่เหลือ"""
    for _ in range(rounds):
        clicked = False
        for sel in SAFE_DIALOG_SELECTORS:
            btn = page.locator(sel + ":visible")
            if btn.count():
                label = (btn.first.inner_text() or "").strip()
                if label in SAFE_DIALOG_LABELS:
                    print(f"   dialog -> {label}")
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
    opt.wait_for(timeout=10000)
    opt.click()


def set_lookup(page, lk, text):
    """lookup เช่น ricoh_customerasset / ricoh_problemcode / ricoh_priority"""
    tag = page.locator(f'[data-id="{lk}.fieldControl-LookupResultsDropdown_{lk}_selected_tag_text"]:visible')
    if tag.count():
        if text.lower() in (tag.first.inner_text() or "").lower():
            return                                   # มีค่านี้อยู่แล้ว
        cancel = page.locator(f'[data-id="{lk}.fieldControl-LookupResultsDropdown_{lk}_microsoftIcon_cancelButton"]:visible')
        if cancel.count():
            cancel.first.click()
    inp = vis(page, f"{lk}.fieldControl-LookupResultsDropdown_{lk}_textInputBox_with_filter_new")
    inp.wait_for(timeout=15000)
    inp.click()
    inp.fill(text)
    item = page.locator(f'span[data-id^="{lk}.fieldControl-"]:visible', has_text=text).first
    item.wait_for(timeout=15000)
    item.click()
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
        self.ws = self.wb[SHEET] if SHEET else self.wb.active
        self._index()
        if "rstatus" not in self.idx:
            self._add_status_column()
            self._index()
        missing = [v for k, v in COL.items() if k not in OPTIONAL_COLS and v not in self.head]
        if missing:
            sys.exit(f"ไม่พบคอลัมน์ใน Excel: {missing}")

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
    if d["job"]:
        print(f"   ! Job Category '{d['job']}' ยังไม่ได้ทำอัตโนมัติ ต้องกรอกเอง")

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

    page.screenshot(path=f"form_row{r}.png")
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
def case_guid(page, ticket):
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


def resolve_case(page, ticket, rtype, text, args):
    """คืน ('done'|'failed'|'skipped', เหตุผล)  ถ้ามี exception ผู้เรียกจะนับเป็น failed"""
    print(f"\n[Resolve] {ticket}")
    if not valid_case_no(ticket):
        return "failed", "รูปแบบเลขเคสไม่ถูกต้อง"
    guid = case_guid(page, ticket)            # อ่านอย่างเดียว ผ่าน session ของหน้าเว็บเอง
    if not guid:
        return "failed", "ไม่พบเคสนี้ (หรืออ่าน API ไม่ได้)"
    st = case_state(page, guid)
    if st == 1:
        print("   เคสนี้ Resolved อยู่แล้ว")
        return "done", "Resolved อยู่แล้ว"
    if st == 2:
        return "failed", "เคสถูก Cancel ไปแล้ว"

    page.goto(RECORD_URL + guid)
    vis(page, TITLE).wait_for(timeout=60000)
    clear_dialogs(page, rounds=3)             # Active Memo -> OK

    btn = page.locator(f'[data-id="{RESOLVE_BTN}"]:visible')
    found = False
    for attempt in range(2):
        try:
            btn.first.wait_for(timeout=30000)
            found = True
            break
        except Exception:
            if attempt == 0:
                print("   ยังไม่เห็นปุ่ม Resolve Case ลองรีโหลดหน้าอีกรอบ...")
                page.reload()
                vis(page, TITLE).wait_for(timeout=60000)
                clear_dialogs(page, rounds=3)
    if not found:
        return "failed", "ไม่พบปุ่ม Resolve Case หลังลอง 2 รอบ"

    if not args.dry_run and not confirm(f"   Resolve {ticket} ด้วย [{rtype}] '{text[:40]}'?", args.yes):
        print("   ข้าม")
        return "skipped", ""

    btn.first.click()
    clear_dialogs(page, rounds=3)             # Ignore and save
    vis(page, RTYPE_SELECT).wait_for(timeout=30000)
    set_option(page, RTYPE_SELECT, rtype)
    set_text(page, RESOLUTION_INPUT, text)
    page.screenshot(path=f"resolve_{ticket}.png")
    if args.dry_run:
        input("   dry-run: กรอกครบแล้ว (ยังไม่กด Resolve) กด Enter เพื่อจบ...")
        return "skipped", ""

    vis(page, RESOLVE_OK).click()
    clear_dialogs(page, rounds=3)
    page.locator(f'[data-id="{RESOLVE_OK}"]:visible').first.wait_for(state="hidden", timeout=30000)
    page.wait_for_timeout(2500)
    state = case_state(page, guid)
    if state == 1:
        print("   Resolve แล้ว (ยืนยันจากระบบ: สถานะ Resolved)")
        return "done", ""
    if state is None:
        print("   กด Resolve แล้ว แต่อ่านสถานะจากระบบไม่ได้ ตรวจในหน้าเว็บอีกครั้ง")
        return "done", "ยืนยันสถานะจากระบบไม่ได้"
    return "failed", f"กด Resolve แล้วแต่ระบบรายงาน statecode={state} (ไม่ใช่ Resolved)"


# ---------------- main ----------------
def main():
    ap = argparse.ArgumentParser(description="D365 Case Bot")
    ap.add_argument("cmd", choices=["new", "resolve"])
    ap.add_argument("--xlsx")
    ap.add_argument("--case")
    ap.add_argument("--type", default="Resolved")
    ap.add_argument("--text")
    ap.add_argument("--limit", type=int, default=1,
                    help="จำนวนแถวสูงสุดที่ลอง (0 = ทุกแถว) ค่าเริ่มต้น 1 นับรวมแถวที่พลาดด้วย")
    ap.add_argument("--row", type=int, help="ทำเฉพาะแถวนี้")
    ap.add_argument("--dry-run", action="store_true", help="กรอกแต่ไม่บันทึก")
    ap.add_argument("--yes", action="store_true", help="ไม่ถามยืนยัน")
    args = ap.parse_args()

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
                outcome, why = resolve_case(page, args.case, args.type, args.text, args)
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
                        try:
                            outcome, why = resolve_case(page, no, rtype, res, args)
                        except KeyboardInterrupt:
                            raise
                        except Exception as e:
                            outcome, why = "failed", short_err(e)
                            recover(page)
                        if outcome == "done":
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


if __name__ == "__main__":
    main()
