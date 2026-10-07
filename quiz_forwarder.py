import os
import sys
import re
import time
import logging
from datetime import datetime
import pandas as pd
import openpyxl
from playwright.sync_api import sync_playwright

# ================= SYSTEM & LOGGING SETUP =================
BASE_DIR = r"C:\EnglishCourseAutomation"
os.makedirs(BASE_DIR, exist_ok=True)

LOG_FILE = os.path.join(BASE_DIR, "quiz_automation.log")
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

# ================= CONFIGURATION =================
EXCEL_FILENAME = "Hashims Community Group.xlsx"
SHEET_NAME = "Fluence(09Sep)"

# Shared staging group where either you OR a teammate forwards the 3 daily posts:
SOURCE_CHAT_NAME = "Community automation"

PROFILE_DIR = os.path.join(BASE_DIR, "whatsapp_profile")
BATCH_SIZE = 5

# Constant footer line on the daily video post
VIDEO_ANCHOR_LINE = "Do the Shadow Writing"

# Set to None to automatically use today's date (1 to 31),
# or set a specific number like TARGET_DAY = 2
TARGET_DAY = None
MARK_TEXT = "SWAP"

# Weekly send limits (checked across the last 7 day columns in Excel)
WEEKLY_LIMITS = {
    "test": None,
    "hot": None,
    "funnel 2": None,
    "warm": 2,
    "fluencer": 2,
    "fluvencer": 2,
    "cool": 1,
    "a_fluencer": 1,
    "staff fluencer": 1,
    "staff_fluencer": 1,
}
# =================================================


def resolve_excel_path():
    """Finds the Excel file next to the .exe launcher first, then falls back to C:\\EnglishCourseAutomation."""
    exe_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    candidate_paths = [
        os.path.join(exe_dir, EXCEL_FILENAME),
        os.path.join(os.getcwd(), EXCEL_FILENAME),
        os.path.join(BASE_DIR, EXCEL_FILENAME)
    ]
    for path in candidate_paths:
        if os.path.exists(path):
            return path
    return os.path.join(BASE_DIR, EXCEL_FILENAME)


def get_target_day():
    return TARGET_DAY if TARGET_DAY is not None else datetime.now().day


def _find_day_col_name(df_columns, d):
    if d in df_columns:
        return d
    if str(d) in df_columns:
        return str(d)
    return None


def count_recent_sends(row, df_columns, current_day, window_days=7):
    """Counts how many times a contact was marked with 'SWAP' or 'SWCP' in the last 7 days."""
    start_day = max(1, current_day - (window_days - 1))
    count = 0
    for d in range(start_day, current_day + 1):
        col = _find_day_col_name(df_columns, d)
        if col is not None:
            cell_val = str(row.get(col, "")).strip().upper()
            if MARK_TEXT.upper() in cell_val or "SWCP" in cell_val:
                count += 1
    return count


def load_and_sort_contacts(excel_path, sheet_name, day_num):
    if not os.path.exists(excel_path):
        raise FileNotFoundError(
            f"Could not find '{EXCEL_FILENAME}'. Please place it in {BASE_DIR} or next to QuizForwarder.exe."
        )

    print(f"📊 Reading Excel file: {excel_path} (Sheet: {sheet_name}, Day Column: {day_num})...")
    df = pd.read_excel(excel_path, sheet_name=sheet_name)

    df["Excel_Row"] = df.index + 2
    df["Type_Clean"] = df["Type"].astype(str).str.strip().str.lower()
    df["Funnel_Clean"] = df.get("Funnel 2 State", pd.Series("", index=df.index)).astype(str).str.strip().str.lower()

    # Check if ANY rows in the sheet are marked as 'Test' in Column D ('Type') or Column E ('Funnel 2 State')
    test_rows = df[(df["Type_Clean"] == "test") | (df["Funnel_Clean"] == "test")].copy()
    if len(test_rows) > 0:
        print("\n" + "🧪" * 28)
        print(f"🧪 TEST MODE ACTIVE: Found {len(test_rows)} row(s) marked 'Test'!")
        print("🧪 Only sending to 'Test' contacts (All other contacts are safely ignored).")
        print("🧪" * 28 + "\n")
        valid_df = test_rows.reset_index(drop=True)
        valid_df["Priority"] = 0
        is_test_mode = True
    else:
        is_test_mode = False

        def assign_priority(type_val):
            if type_val in ["hot", "funnel 2"]:
                return 1
            elif type_val in ["warm", "fluencer", "fluvencer"]:
                return 2
            elif type_val in ["cool", "a_fluencer", "staff fluencer", "staff_fluencer"]:
                return 3
            return 99

        df["Priority"] = df["Type_Clean"].apply(assign_priority)
        valid_df = df[df["Priority"] < 99].copy()
        valid_df = valid_df.sort_values(by="Priority", kind="mergesort").reset_index(drop=True)

    day_col = _find_day_col_name(valid_df.columns, day_num)
    if day_col is None:
        raise Exception(f"Could not find day column '{day_num}' in sheet '{sheet_name}'.")

    contacts = []
    skipped_today = 0
    skipped_weekly_cap = 0

    for _, row in valid_df.iterrows():
        type_clean = row["Type_Clean"]
        current_mark = str(row.get(day_col, "")).strip()

        if not is_test_mode and MARK_TEXT.lower() in current_mark.lower():
            skipped_today += 1
            continue

        weekly_cap = WEEKLY_LIMITS.get(type_clean, None)
        if not is_test_mode and weekly_cap is not None:
            sends_this_week = count_recent_sends(row, valid_df.columns, day_num, window_days=7)
            if sends_this_week >= weekly_cap:
                skipped_weekly_cap += 1
                continue

        raw_name = str(row.get("Name", "")).strip()
        raw_num = str(row.get("Number", "")).strip()

        first_num_line = raw_num.split("\n")[0].strip()
        digits_only = re.sub(r"\D", "", first_num_line)

        contacts.append({
            "excel_row": int(row["Excel_Row"]),
            "name": raw_name,
            "raw_number": first_num_line,
            "digits": digits_only,
            "type": str(row.get("Type", "")).strip(),
            "priority": row["Priority"]
        })

    print(
        f"✅ Loaded {len(contacts)} eligible contact(s) for Day {day_num} "
        f"(Skipped {skipped_today} already marked '{MARK_TEXT}' today, "
        f"and {skipped_weekly_cap} at their Warm/Cool weekly limit)."
    )
    return contacts


def mark_contacts_as_swap(excel_path, sheet_name, day_num, completed_contacts):
    """Writes 'SWAP' into the specific date column for each completed contact."""
    if not completed_contacts:
        return

    try:
        wb = openpyxl.load_workbook(excel_path)
        ws = wb[sheet_name]

        target_col_idx = None
        for col_idx in range(1, ws.max_column + 1):
            cell_val = ws.cell(row=1, column=col_idx).value
            if str(cell_val).strip() == str(day_num):
                target_col_idx = col_idx
                break

        if target_col_idx is None:
            print(f"⚠️ Could not find Day {day_num} header in Excel to mark '{MARK_TEXT}'.")
            return

        for contact in completed_contacts:
            row_idx = contact["excel_row"]
            existing = ws.cell(row=row_idx, column=target_col_idx).value
            if existing and str(existing).strip():
                if MARK_TEXT.lower() not in str(existing).lower():
                    ws.cell(row=row_idx, column=target_col_idx).value = f"{existing}\n{MARK_TEXT}"
            else:
                ws.cell(row=row_idx, column=target_col_idx).value = MARK_TEXT

        wb.save(excel_path)
        wb.close()
        print(f"💾 Marked '{MARK_TEXT}' in Day {day_num} column for {len(completed_contacts)} contact(s).")
    except PermissionError:
        print("❌ Could not save Excel file! Please close 'Hashims Community Group.xlsx' in Excel so Python can update it.")
    except Exception as e:
        print(f"⚠️ Error saving '{MARK_TEXT}' to Excel: {e}")


class WhatsAppQuizForwarder:
    def __init__(self, profile_dir):
        self.profile_dir = profile_dir
        os.makedirs(self.profile_dir, exist_ok=True)
        self.playwright = sync_playwright().start()

        try:
            self.browser = self.playwright.chromium.launch_persistent_context(
                user_data_dir=self.profile_dir,
                headless=False,
                channel="chrome",
                args=["--start-maximized"]
            )
        except Exception:
            self.browser = self.playwright.chromium.launch_persistent_context(
                user_data_dir=self.profile_dir,
                headless=False,
                channel="msedge",
                args=["--start-maximized"]
            )

        self.page = self.browser.pages[0]
        self.page.goto("https://web.whatsapp.com/", timeout=0)
        print("Waiting for WhatsApp Web to load...")

        try:
            self.page.wait_for_selector("#pane-side, header", timeout=20000)
        except Exception:
            print("Please scan the WhatsApp QR code if prompted. Waiting up to 2 minutes...")
            self.page.wait_for_selector("#pane-side, header", timeout=120000)

        print("WhatsApp Web is ready.")
        time.sleep(2)

    def _force_scroll_to_very_bottom(self):
        """Scrolls the chat message pane all the way to the bottom."""
        for _ in range(2):
            self.page.mouse.move(850, 500)
            self.page.mouse.wheel(0, 4000)
            self.page.evaluate("""() => {
                const main = document.querySelector('#main');
                if (!main) return;
                const divs = Array.from(main.querySelectorAll('div'));
                for (const d of divs) {
                    if (d.scrollHeight > d.clientHeight + 50) {
                        d.scrollTop = d.scrollHeight;
                    }
                }
            }""")
            time.sleep(0.4)

    def _is_chat_actually_open(self, chat_name):
        """Checks if the right-hand #main chat pane is open and matches the target group/chat."""
        return self.page.evaluate("""(targetName) => {
            const header = document.querySelector('#main header');
            if (!header) return false;
            const hText = (header.innerText || '').toLowerCase();
            if (!hText.includes(targetName.toLowerCase())) return false;
            const msgCount = document.querySelectorAll('#main .message-in, #main .message-out, #main [data-id]').length;
            return msgCount > 0;
        }""", chat_name)

    def open_source_chat(self, chat_name):
        """Forcefully ensures the source chat is open and settled, preventing stale DOM reads."""
        # Wait 2 seconds before checking, so if WhatsApp just navigated to a recipient's chat,
        # the DOM animation finishes and we don't accidentally read the old header!
        time.sleep(2.0)

        if self._is_chat_actually_open(chat_name):
            print(f"✅ Source group '{chat_name}' is open on screen!")
            return

        print(f"🔍 Navigating back to source group: '{chat_name}'...")

        # Clear sidebar search box explicitly
        search_box = self.page.locator('#side [contenteditable="true"], #side input[type="text"], #side [role="textbox"]').first
        if search_box.is_visible():
            search_box.click()
            self.page.keyboard.press("Control+A")
            self.page.keyboard.press("Backspace")
            time.sleep(0.5)

            self.page.keyboard.insert_text(chat_name)
            time.sleep(2.0)

            # Look for the chat row in the search results
            matched_row = self.page.locator("#pane-side [role='row']").get_by_text(chat_name, exact=False).first
            if matched_row.is_visible():
                box = matched_row.bounding_box()
                if box:
                    self.page.mouse.click(box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
                else:
                    matched_row.click(force=True)
            else:
                self.page.keyboard.press("ArrowDown")
                time.sleep(0.3)
                self.page.keyboard.press("Enter")

            time.sleep(2.5)
            if self._is_chat_actually_open(chat_name):
                print(f"✅ Re-opened '{chat_name}' via search!")
                return

        print("\n" + "=" * 60)
        print(f"👉 Please manually click on '{chat_name}' in the left chat list of WhatsApp Web!")
        print("=" * 60)
        for _ in range(60):
            if self._is_chat_actually_open(chat_name):
                print(f"✅ '{chat_name}' is now open! Resuming...")
                time.sleep(1)
                return
            time.sleep(2)

        raise Exception(f"Could not return to '{chat_name}'.")

    def _get_forward_modal_and_search(self):
        """Verifies the Forward modal dialog is actually open."""
        dialog_inputs = self.page.locator(
            'div[role="dialog"] div[contenteditable="true"], '
            'div[role="dialog"] input[type="text"], '
            '[data-animate-modal-popup="true"] div[contenteditable="true"], '
            '[data-animate-modal-popup="true"] input[type="text"]'
        ).all()
        for inp in dialog_inputs:
            if inp.is_visible():
                return inp

        has_forward_title = self.page.evaluate("""() => {
            const bodyTxt = document.body.innerText || '';
            return bodyTxt.includes('Forward message to') ||
                   bodyTxt.includes('Forward update to') ||
                   bodyTxt.includes('Forward messages to');
        }""")
        if has_forward_title:
            all_inputs = self.page.locator('div[contenteditable="true"], input[type="text"]').all()
            for inp in all_inputs:
                if not inp.is_visible():
                    continue
                box = inp.bounding_box()
                if box and 300 < box["x"] < 700 and 140 < box["y"] < 360:
                    return inp

        return None

    def _close_modal_if_open(self):
        """Only closes the Forward popup if it is actually open."""
        if self._get_forward_modal_and_search() is not None:
            close_btn = self.page.locator(
                'div[role="dialog"] [aria-label="Close"], '
                'div[role="dialog"] [data-icon="x"], '
                '[data-animate-modal-popup="true"] [data-icon="x"]'
            ).first
            if close_btn.is_visible():
                close_btn.click(force=True)
            else:
                self.page.keyboard.press("Escape")
            time.sleep(0.6)

    def _check_and_advance_to_modal(self):
        """Handles the multi-select Forward flow."""
        if self._get_forward_modal_and_search() is not None:
            return True

        fwd_menu = self.page.locator(
            'li:has-text("Forward"), '
            'div[role="button"]:has-text("Forward"), '
            '[data-animate-dropdown-item="true"]:has-text("Forward"), '
            'span:has-text("Forward")'
        ).first
        if fwd_menu.is_visible():
            fwd_menu.click(force=True)
            time.sleep(1.2)
            if self._get_forward_modal_and_search() is not None:
                return True

        clicked_bottom_bar_fwd = self.page.evaluate("""() => {
            const main = document.querySelector('#main');
            if (!main) return false;

            const candidates = Array.from(main.querySelectorAll(
                '[data-icon*="forward"], button[title*="Forward"], [aria-label*="Forward"], button, [role="button"]'
            ));
            for (let i = candidates.length - 1; i >= 0; i--) {
                const el = candidates[i];
                const r = el.getBoundingClientRect();
                if (r.width === 0 || r.height === 0) continue;

                const icon = (el.getAttribute('data-icon') || (el.querySelector('[data-icon]') ? el.querySelector('[data-icon]').getAttribute('data-icon') : '') || '').toLowerCase();
                const aria = (el.getAttribute('aria-label') || el.getAttribute('title') || '').toLowerCase();

                if ((icon.includes('forward') || aria.includes('forward')) && r.top > 500) {
                    const btn = el.closest('button') || el.closest('[role="button"]') || el;
                    btn.click();
                    return true;
                }
            }
            return false;
        }""")
        if clicked_bottom_bar_fwd:
            time.sleep(1.5)
            if self._get_forward_modal_and_search() is not None:
                return True

        return self._get_forward_modal_and_search() is not None

    def _find_and_click_forward_in_chat(self, item_type, quiz_index):
        # 1. Double check we are explicitly inside the source chat
        time.sleep(1.0)
        if not self._is_chat_actually_open(SOURCE_CHAT_NAME):
            self.open_source_chat(SOURCE_CHAT_NAME)

        # 2. Strict Fail-Safe: If it STILL isn't the source chat, ABORT immediately!
        if not self._is_chat_actually_open(SOURCE_CHAT_NAME):
            raise Exception(f"❌ CRITICAL: Chat pane is NOT '{SOURCE_CHAT_NAME}'. Aborting to prevent sending from a stranger's chat.")

        self._force_scroll_to_very_bottom()

        if item_type == "video":
            for _ in range(4):
                has_video_post = self.page.evaluate("""(anchor) => {
                    const main = document.querySelector('#main');
                    if (!main) return false;
                    const txt = (main.innerText || '').toLowerCase();
                    const hasKw = txt.includes(anchor.toLowerCase()) || txt.includes('shadow writing') || txt.includes('rule: use');
                    const hasVidTag = main.querySelector('video, [data-icon*="video"], [data-icon*="media-play"]') !== null;
                    return hasKw || hasVidTag;
                }""", VIDEO_ANCHOR_LINE)
                if has_video_post:
                    break
                self.page.mouse.move(850, 450)
                self.page.mouse.wheel(0, -350)
                time.sleep(0.7)

        target_box = self.page.evaluate("""({ itemType, quizIdx, videoAnchor }) => {
            const rawNodes = Array.from(document.querySelectorAll(
                '#main .message-in, #main .message-out, #main [data-id], #main [role="row"]'
            ));
            const posts = [];

            for (const r of rawNodes) {
                const bubble = (r.classList.contains('message-in') || r.classList.contains('message-out'))
                    ? r
                    : (r.querySelector('.message-in, .message-out') || r);

                const rect = bubble.getBoundingClientRect();
                if (rect.width < 140 || rect.height < 55 || rect.height > 1200) continue;

                const dup = posts.some(p => Math.abs(p.el.getBoundingClientRect().top - rect.top) < 25);
                if (dup) continue;

                const txt = (bubble.innerText || '').trim();
                if (!txt && !bubble.querySelector('video, img')) continue;

                const isQuiz = (txt.includes('A)') && txt.includes('B)')) ||
                               txt.includes('View responses') ||
                               (bubble.querySelector('[data-icon*="poll"]') !== null);

                const isDailyVideo = !isQuiz && (
                    txt.toLowerCase().includes(videoAnchor.toLowerCase()) ||
                    txt.toLowerCase().includes('shadow writing') ||
                    txt.toLowerCase().includes('rule: use') ||
                    bubble.querySelector('video, [data-icon*="video"], [data-icon*="play"]') !== null
                );

                posts.push({
                    el: bubble,
                    rowEl: bubble.closest('[role="row"]') || bubble.closest('[data-id]') || bubble.parentElement,
                    top: rect.top,
                    isQuiz: isQuiz,
                    isDailyVideo: isDailyVideo,
                    text: txt
                });
            }

            posts.sort((a, b) => a.top - b.top);

            if (posts.length === 0) {
                const headerTitle = document.querySelector('#main header') ? document.querySelector('#main header').innerText.split('\\n')[0] : 'Unknown';
                return { error: 'No message bubbles detected inside open chat "' + headerTitle + '".' };
            }

            const last3 = posts.slice(-3);
            let chosen = null;
            let matchedBy = '';

            if (itemType === 'video') {
                const videoMatches = posts.filter(p => p.isDailyVideo && !p.isQuiz);
                if (videoMatches.length > 0) {
                    chosen = videoMatches[videoMatches.length - 1];
                    matchedBy = 'VIDEO ("Do the Shadow Writing")';
                } else {
                    const nonQuizzes = last3.filter(p => !p.isQuiz);
                    chosen = nonQuizzes.length > 0 ? nonQuizzes[0] : last3[0];
                    matchedBy = 'VIDEO (1st of 3 messages)';
                }
            } else {
                const quizzes = posts.filter(p => p.isQuiz);
                if (quizzes.length >= 2) {
                    chosen = (quizIdx === 0) ? quizzes[quizzes.length - 2] : quizzes[quizzes.length - 1];
                    matchedBy = (quizIdx === 0) ? 'QUIZ 1 (Older)' : 'QUIZ 2 (Latest)';
                } else if (quizzes.length === 1) {
                    chosen = quizzes[0];
                    matchedBy = 'QUIZ (Single Quiz)';
                } else if (last3.length >= 2) {
                    chosen = (quizIdx === 0) ? last3[last3.length - 2] : last3[last3.length - 1];
                    matchedBy = 'QUIZ (By position)';
                } else {
                    chosen = last3[last3.length - 1];
                    matchedBy = 'QUIZ (Fallback)';
                }
            }

            const childSpans = Array.from(chosen.el.querySelectorAll('span, div'));
            let anchorNode = chosen.el;
            for (const s of childSpans) {
                const st = (s.innerText || '').trim().toLowerCase();
                if (itemType === 'video' && st.includes('do the shadow writing') && s.children.length <= 2) {
                    anchorNode = s;
                } else if (itemType !== 'video' && (st.includes('view channel') || st.includes('d)')) && s.children.length <= 2) {
                    anchorNode = s;
                }
            }

            anchorNode.scrollIntoView({ block: 'center', behavior: 'instant' });

            const finalRect = chosen.el.getBoundingClientRect();

            const scope = chosen.rowEl || document.querySelector('#main');
            const btns = Array.from(scope.querySelectorAll('button, [role="button"], [data-icon]'));
            let sideForwardBtn = null;

            for (const btn of btns) {
                const r = btn.getBoundingClientRect();
                if (r.width === 0 || r.height === 0 || r.top < 100) continue;
                const icon = (btn.getAttribute('data-icon') || '').toLowerCase();
                const aria = (btn.getAttribute('aria-label') || btn.getAttribute('title') || '').toLowerCase();

                if (icon.includes('smiley') || icon.includes('react') || aria.includes('react')) continue;

                const isLeftOfBubble = (r.right <= finalRect.left + 15 && r.left >= finalRect.left - 95);
                const isRightOfBubble = (r.left >= finalRect.right - 15 && r.right <= finalRect.right + 95);
                const isVerticallyAligned = (r.top >= finalRect.top - 10 && r.bottom <= finalRect.bottom + 10);

                if (icon.includes('forward') || aria.includes('forward') ||
                    ((isLeftOfBubble || isRightOfBubble) && isVerticallyAligned)) {
                    sideForwardBtn = {
                        x: Math.round(r.left + r.width / 2),
                        y: Math.round(r.top + r.height / 2)
                    };
                    if (icon.includes('forward') || aria.includes('forward')) break;
                }
            }

            const preview = chosen.text.replace(/\\n+/g, ' ').substring(0, 50);
            return {
                matchedBy: matchedBy,
                preview: preview,
                left: Math.round(finalRect.left),
                right: Math.round(finalRect.right),
                top: Math.round(finalRect.top),
                bottom: Math.round(finalRect.bottom),
                sideForwardBtn: sideForwardBtn
            };
        }""", {"itemType": item_type, "quizIdx": quiz_index, "videoAnchor": VIDEO_ANCHOR_LINE})

        time.sleep(0.8)

        if "error" in target_box:
            raise Exception(target_box["error"])

        print(f"  🎯 Locked onto {target_box['matchedBy']}: '{target_box['preview']}...'")

        if target_box.get("sideForwardBtn"):
            fx = target_box["sideForwardBtn"]["x"]
            fy = target_box["sideForwardBtn"]["y"]
            self.page.mouse.move(fx, fy)
            time.sleep(0.2)
            self.page.mouse.click(fx, fy)
            time.sleep(1.2)
            if self._check_and_advance_to_modal():
                return

        safe_x = target_box["left"] + 80
        safe_y = min(max(target_box["bottom"] - 55, 140), 780)
        self.page.mouse.move(safe_x, safe_y)
        time.sleep(0.3)
        self.page.mouse.click(safe_x, safe_y, button="right")
        time.sleep(1.0)
        if self._check_and_advance_to_modal():
            return

        tr_x = target_box["right"] - 16
        tr_y = max(target_box["top"] + 16, 130)
        self.page.mouse.move(target_box["left"] + 100, min(max(target_box["bottom"] - 60, 150), 750))
        time.sleep(0.4)
        self.page.mouse.move(tr_x, tr_y)
        time.sleep(0.5)
        self.page.mouse.click(tr_x, tr_y)
        time.sleep(1.0)
        if self._check_and_advance_to_modal():
            return

        raise Exception(f"Could not open Forward modal for {target_box['matchedBy']}.")

    def _select_contact_in_modal(self, contact):
        queries_to_try = []
        if contact["raw_number"] and contact["raw_number"].lower() != "nan":
            queries_to_try.append(contact["raw_number"])
        if len(contact["digits"]) >= 10:
            last_10 = contact["digits"][-10:]
            if last_10 not in queries_to_try:
                queries_to_try.append(last_10)
        if contact["name"] and contact["name"].lower() != "nan" and contact["name"] not in queries_to_try:
            queries_to_try.append(contact["name"])

        for q in queries_to_try:
            search_input = self._get_forward_modal_and_search()
            if search_input is None:
                return "MODAL_CLOSED"

            try:
                s_box = search_input.bounding_box()
                if not s_box:
                    return "MODAL_CLOSED"
                self.page.mouse.click(s_box["x"] + 30, s_box["y"] + s_box["height"] / 2)
            except Exception:
                return "MODAL_CLOSED"

            self.page.keyboard.press("Control+A")
            self.page.keyboard.press("Backspace")
            time.sleep(0.3)

            self.page.keyboard.insert_text(q)
            time.sleep(1.6)

            row_info = self.page.evaluate("""(sRect) => {
                const dialog = document.querySelector('div[role="dialog"], [data-animate-modal-popup="true"]') || document.body;
                const dText = (dialog.innerText || '').toLowerCase();
                if (dText.includes('no results') || dText.includes('no chats') || dText.includes('no contacts')) {
                    return { found: false, reason: 'No results text visible' };
                }

                const minRowY = sRect.y + sRect.height + 50;
                const maxRowY = sRect.y + sRect.height + 155;

                const rowCandidates = [];
                const allNodes = Array.from(dialog.querySelectorAll('span[dir="auto"], span[title], div[role="checkbox"], input[type="checkbox"], [role="listitem"], [role="button"]'));

                for (const n of allNodes) {
                    const r = n.getBoundingClientRect();
                    if (r.width === 0 || r.height === 0 || r.height > 110) continue;
                    const cy = r.top + r.height / 2;
                    if (cy >= minRowY && cy <= maxRowY && r.left >= sRect.x - 60 && r.right <= sRect.x + sRect.width + 80) {
                        const label = (n.innerText || n.getAttribute('title') || '').trim().split('\\n')[0];
                        const low = label.toLowerCase();
                        if (!label || low === 'recent chats' || low === 'contacts' || low === 'groups' || low === 'frequently contacted') continue;

                        rowCandidates.push({
                            y: Math.round(cy),
                            x: Math.round(r.left + Math.min(r.width / 2, 80)),
                            label: label
                        });
                    }
                }

                if (rowCandidates.length > 0) {
                    rowCandidates.sort((a, b) => a.y - b.y);
                    const topHit = rowCandidates[0];
                    return {
                        found: true,
                        rowY: topHit.y,
                        rowX: topHit.x,
                        label: topHit.label
                    };
                }

                return { found: false, reason: 'No contact elements in row band' };
            }""", s_box)

            if not row_info.get("found"):
                continue

            click_x = int(row_info["rowX"])
            click_y = int(row_info["rowY"])
            print(f"     ✅ Clicking match '{row_info.get('label')}' ONCE at ({click_x}, {click_y})...")
            self.page.mouse.move(click_x, click_y)
            time.sleep(0.2)
            self.page.mouse.click(click_x, click_y)
            time.sleep(0.8)
            return "SELECTED"

        search_input = self._get_forward_modal_and_search()
        if search_input is not None:
            try:
                s_box = search_input.bounding_box()
                if s_box:
                    self.page.mouse.click(s_box["x"] + 30, s_box["y"] + s_box["height"] / 2)
                    self.page.keyboard.press("Control+A")
                    self.page.keyboard.press("Backspace")
                    time.sleep(0.3)
            except Exception:
                pass
        return "NOT_FOUND"

    def _click_modal_send_button(self):
        """Clicks the green Send button in the Forward modal and EXPLICITLY WAITS for the screen transition."""
        send_btn = self.page.locator(
            'div[role="dialog"] [data-icon*="send"], '
            'div[role="dialog"] [aria-label="Send"], '
            '[data-animate-modal-popup="true"] [data-icon*="send"], '
            '[data-animate-modal-popup="true"] [aria-label="Send"], '
            '[data-icon*="send"], [aria-label="Send"]'
        ).last
        
        button_clicked = False
        if send_btn.is_visible():
            box = send_btn.bounding_box()
            if box:
                self.page.mouse.click(box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
                button_clicked = True
            else:
                send_btn.click(force=True)
                button_clicked = True

        if not button_clicked:
            clicked_green_btn = self.page.evaluate("""() => {
                const dialog = document.querySelector('div[role="dialog"], [data-animate-modal-popup="true"]') || document.body;
                const btns = Array.from(dialog.querySelectorAll('button, [role="button"], span[data-icon]'));
                for (let i = btns.length - 1; i >= 0; i--) {
                    const r = btns[i].getBoundingClientRect();
                    if (r.width >= 28 && r.width <= 75 && r.top > 420 && r.left > 420 && r.left < 780) {
                        btns[i].click();
                        return true;
                    }
                }
                return false;
            }""")
            if not clicked_green_btn:
                self.page.keyboard.press("Enter")

        # MANDATORY WAIT: WhatsApp instantly navigates to the recipient's chat.
        # We must freeze the bot to let the DOM settle, or the next batch will read a stale screen!
        print("  ⏳ Waiting for UI to transition to the forwarded chat...")
        time.sleep(3.5)

    def forward_item_to_batch(self, item_type, quiz_index, batch_contacts):
        self._close_modal_if_open()

        label = "Video ('Do the Shadow Writing')" if item_type == "video" else ("Quiz 1 (Older)" if quiz_index == 0 else "Quiz 2 (Latest)")
        print(f"\n🔄 Selecting {label} from '{SOURCE_CHAT_NAME}' for batch of {len(batch_contacts)} contact(s)...")

        self._find_and_click_forward_in_chat(item_type, quiz_index)

        if self._get_forward_modal_and_search() is None:
            raise Exception("Forward popup did not open.")

        succeeded_contacts = []

        for contact in batch_contacts:
            print(f"  -> Selecting: {contact['name']} ({contact['raw_number']}) [Row {contact['excel_row']} | {contact['type']}]")
            status = self._select_contact_in_modal(contact)
            if status == "SELECTED":
                succeeded_contacts.append(contact)
            elif status == "MODAL_CLOSED":
                print("     ⚠️ Forward popup closed early; Re-opening popup for remaining contacts...")
                self._find_and_click_forward_in_chat(item_type, quiz_index)
                if self._get_forward_modal_and_search() is not None:
                    retry_status = self._select_contact_in_modal(contact)
                    if retry_status == "SELECTED":
                        succeeded_contacts.append(contact)
            else:
                print(f"     ❌ Could not find '{contact['name']}' in this WhatsApp account.")
                logging.warning(f"Contact not found: {contact['name']} ({contact['raw_number']})")

        if succeeded_contacts:
            if self._get_forward_modal_and_search() is not None:
                self._click_modal_send_button()
            print(f"✅ Forwarded {label} to {len(succeeded_contacts)} contact(s)!")
            return succeeded_contacts
        else:
            print("  ⚠️ None of the contacts in this batch matched. Closing popup...")
            self._close_modal_if_open()
            return []

    def close(self):
        self.browser.close()
        self.playwright.stop()


def run_quiz_automation():
    print("=" * 60)
    print(f"🚀 FLUENCE VIDEO & QUIZ FORWARDER (SOURCE: '{SOURCE_CHAT_NAME}')")
    print("=" * 60)

    excel_path = resolve_excel_path()
    day_num = get_target_day()
    contacts = load_and_sort_contacts(excel_path, SHEET_NAME, day_num)

    if not contacts:
        print(f"🎉 All eligible contacts are already completed for Day {day_num}! Nothing to send.")
        return

    bot = WhatsAppQuizForwarder(PROFILE_DIR)

    try:
        bot.open_source_chat(SOURCE_CHAT_NAME)

        total_batches = (len(contacts) + BATCH_SIZE - 1) // BATCH_SIZE
        quiz_turn_counter = 0

        for batch_idx in range(total_batches):
            start = batch_idx * BATCH_SIZE
            end = start + BATCH_SIZE
            batch = contacts[start:end]

            # Alternates every batch: 5 Video -> 5 Quiz 1 -> 5 Video -> 5 Quiz 2
            if batch_idx % 2 == 0:
                item_type = "video"
                quiz_to_send = 0
            else:
                item_type = "quiz"
                quiz_to_send = quiz_turn_counter % 2
                quiz_turn_counter += 1

            print(f"\n--- Batch {batch_idx + 1} of {total_batches} (Day {day_num} | Mode: {item_type.upper()}) ---")
            
            sent_contacts = bot.forward_item_to_batch(item_type, quiz_to_send, batch)

            if sent_contacts:
                mark_contacts_as_swap(excel_path, SHEET_NAME, day_num, sent_contacts)

        print(f"\n🎉 All batches for Day {day_num} have been forwarded and marked '{MARK_TEXT}'!")
        logging.info(f"Completed all {total_batches} batches for Day {day_num}.")
    finally:
        bot.close()


if __name__ == "__main__":
    try:
        run_quiz_automation()
    except Exception as e:
        logging.exception("Critical error in Quiz Forwarder:")
        print(f"\n❌ Critical Error: {e}")
    input("\nPress Enter to exit...")
