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
CHANNEL_NAME = "in_Fluence_y (English Reading & Writing)"
CHANNEL_SHORT_KEY = "in_Fluence_y"
PROFILE_DIR = os.path.join(BASE_DIR, "whatsapp_profile")
BATCH_SIZE = 5

# Set to None to automatically use today's date (1 to 31),
# or set a specific number like TARGET_DAY = 2
TARGET_DAY = None
MARK_TEXT = "SWAP"

# Weekly send limits (checked across the last 7 day columns in Excel)
# None = Unlimited (sent every day)
WEEKLY_LIMITS = {
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

        if MARK_TEXT.lower() in current_mark.lower():
            skipped_today += 1
            continue

        weekly_cap = WEEKLY_LIMITS.get(type_clean, None)
        if weekly_cap is not None:
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
        f"✅ Loaded {len(contacts)} eligible contacts for Day {day_num} "
        f"(Skipped {skipped_today} already marked '{MARK_TEXT}' today, "
        f"and {skipped_weekly_cap} at their Warm/Cool weekly limit)."
    )
    logging.info(
        f"Day {day_num}: {len(contacts)} eligible, {skipped_today} already SWAP, {skipped_weekly_cap} weekly capped."
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
        print(f"💾 Marked '{MARK_TEXT}' in Day {day_num} column for {len(completed_contacts)} contacts.")
        logging.info(f"Marked '{MARK_TEXT}' in Day {day_num} for {len(completed_contacts)} contacts.")
    except PermissionError:
        print("❌ Could not save Excel file! Please close 'Hashims Community Group.xlsx' in Excel so Python can update it.")
        logging.error("PermissionError saving Excel file (file open in Excel).")
    except Exception as e:
        print(f"⚠️ Error saving '{MARK_TEXT}' to Excel: {e}")
        logging.error(f"Error saving SWAP to Excel: {e}")


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

    def _find_channel_targets_js(self):
        """Locates both the Quizzes and the Video/Top Post sitting above the quizzes inside #main."""
        return self.page.evaluate("""() => {
            const rawNodes = Array.from(document.querySelectorAll('#main .message-in, #main .message-out, #main [role="row"]'));
            const items = [];

            for (const r of rawNodes) {
                const bubble = (r.classList.contains('message-in') || r.classList.contains('message-out'))
                    ? r
                    : (r.querySelector('.message-in, .message-out') || r);

                const rect = bubble.getBoundingClientRect();
                if (rect.width < 120 || rect.height < 45) continue;

                const duplicate = items.some(existing => Math.abs(existing.top - rect.top) < 20);
                if (duplicate) continue;

                const txt = (bubble.innerText || '').trim();
                const isQuiz = txt.includes('View responses') || (txt.includes('A)') && txt.includes('B)'));
                const hasVideoTag = bubble.querySelector('video, [data-icon*="video"], [data-icon*="media"], [data-icon*="play"], iframe, a[href*="youtu"]') !== null;

                items.push({
                    top: Math.round(rect.top),
                    bottom: Math.round(rect.bottom),
                    left: Math.round(rect.left),
                    right: Math.round(rect.right),
                    width: Math.round(rect.width),
                    height: Math.round(rect.height),
                    isQuiz: isQuiz,
                    hasVideoTag: hasVideoTag,
                    preview: txt.split('\\n')[0].substring(0, 45) || (hasVideoTag ? '[Video Post]' : '[Channel Post]')
                });
            }

            items.sort((a, b) => a.top - b.top);

            const quizzes = items.filter(i => i.isQuiz);
            const nonQuizzes = items.filter(i => !i.isQuiz);

            let videoTarget = null;
            if (quizzes.length > 0) {
                const firstQuizTop = quizzes[0].top;
                const aboveQuizzes = nonQuizzes.filter(i => i.top < firstQuizTop - 10);
                const videoPostsAbove = aboveQuizzes.filter(i => i.hasVideoTag);

                if (videoPostsAbove.length > 0) {
                    videoTarget = videoPostsAbove[videoPostsAbove.length - 1];
                } else if (aboveQuizzes.length > 0) {
                    videoTarget = aboveQuizzes[aboveQuizzes.length - 1];
                }
            }
            if (!videoTarget && nonQuizzes.length > 0) {
                videoTarget = nonQuizzes[nonQuizzes.length - 1];
            }

            return {
                quizzes: quizzes,
                video: videoTarget
            };
        }""")

    def open_channel(self, full_name, short_key):
        print(f"🔍 Opening Channels tab (below Status) to find: {full_name}")

        header = self.page.locator("#main header").first
        if header.is_visible() and short_key.lower() in header.inner_text().lower():
            if len(self._find_channel_targets_js().get("quizzes", [])) >= 1:
                print("✅ Channel is already open on screen!")
                return

        clicked_channels_tab = False
        channel_tab_selectors = [
            '[aria-label="Channels"]',
            '[title="Channels"]',
            '[data-icon="newsletter-tab"]',
            '[data-icon="channels"]',
            '[data-navbar-item-index="2"]'
        ]
        for sel in channel_tab_selectors:
            loc = self.page.locator(sel).first
            if loc.is_visible():
                loc.click(force=True)
                clicked_channels_tab = True
                time.sleep(2)
                break

        if not clicked_channels_tab:
            status_btn = self.page.locator('[aria-label="Status"], [title="Status"], [data-icon*="status"]').first
            if status_btn.is_visible():
                box = status_btn.bounding_box()
                if box:
                    self.page.mouse.click(box["x"] + box["width"] / 2, box["y"] + box["height"] + 28)
                    time.sleep(2)

        channel_item = self.page.get_by_text(short_key, exact=False).first
        if channel_item.is_visible():
            channel_item.click(force=True)
            time.sleep(3)
            if len(self._find_channel_targets_js().get("quizzes", [])) >= 1:
                print("✅ Opened channel from the Channels list!")
                return

        search_box = self.page.locator('[placeholder="Search"], [aria-label*="Search"], [contenteditable="true"]').first
        if search_box.is_visible():
            search_box.click()
            self.page.keyboard.press("Control+A")
            self.page.keyboard.press("Backspace")
            time.sleep(0.5)
            search_box.fill(short_key)
            time.sleep(2)

            match = self.page.get_by_text(short_key, exact=False).first
            if match.is_visible():
                match.click(force=True)
                time.sleep(3)
                if len(self._find_channel_targets_js().get("quizzes", [])) >= 1:
                    print("✅ Opened channel via Channels search bar!")
                    return

        print("\n" + "=" * 60)
        print(f"👉 Please click on the '{full_name}' channel once in WhatsApp Web.")
        print("=" * 60)
        for _ in range(60):
            if len(self._find_channel_targets_js().get("quizzes", [])) >= 1:
                print("✅ Channel detected! Starting batch forwarding...")
                time.sleep(1)
                return
            time.sleep(2)

        raise Exception("Could not find quizzes in the open Channel.")

    def _get_forward_search_box(self):
        """Locates the search input inside the 'Forward message to' popup window."""
        modal_inputs = self.page.locator(
            'div[role="dialog"] div[contenteditable="true"], '
            'div[role="dialog"] input[type="text"], '
            'div[role="dialog"] [role="textbox"]'
        )
        if modal_inputs.count() > 0 and modal_inputs.first.is_visible():
            return modal_inputs.first

        all_inputs = self.page.locator('div[contenteditable="true"], input[type="text"], [role="textbox"]').all()
        for inp in all_inputs:
            if not inp.is_visible():
                continue
            box = inp.bounding_box()
            # Center popup search bar sits at x > 350 and y < 420
            if box and box["x"] > 350 and box["y"] < 420:
                return inp

        return None

    def _check_and_advance_to_modal(self):
        """Checks if the Forward modal is open, or clicks the Forward menu/button to open it."""
        if self._get_forward_search_box() is not None:
            return True

        fwd_menu = self.page.locator(
            'li:has-text("Forward"), div[role="button"]:has-text("Forward"), span:has-text("Forward")'
        ).first
        if fwd_menu.is_visible():
            fwd_menu.click(force=True)
            time.sleep(1.2)
            if self._get_forward_search_box() is not None:
                return True

        bottom_fwd = self.page.locator(
            '#main [data-icon*="forward"], button[title*="Forward"], [aria-label*="Forward"]'
        ).last
        if bottom_fwd.is_visible():
            bottom_fwd.click(force=True)
            time.sleep(1.2)
            if self._get_forward_search_box() is not None:
                return True

        return self._get_forward_search_box() is not None

    def _scroll_target_into_comfort_zone(self, item_type, quiz_index):
        """Scrolls the Channel view so the target Video or Quiz is cleanly on screen."""
        for _ in range(6):
            targets = self._find_channel_targets_js()
            quizzes = targets.get("quizzes", [])
            video = targets.get("video")

            chosen = None
            if item_type == "video":
                chosen = video if video else (quizzes[0] if quizzes else None)
            else:
                if len(quizzes) >= 2:
                    chosen = quizzes[-2:][quiz_index]
                elif len(quizzes) == 1:
                    chosen = quizzes[0]

            if chosen is None:
                self.page.mouse.move(850, 450)
                self.page.mouse.wheel(0, -350 if item_type == "video" else 300)
                time.sleep(0.8)
                continue

            if chosen["top"] < 100:
                self.page.mouse.move(850, 450)
                self.page.mouse.wheel(0, -280)
                time.sleep(0.8)
                continue

            if chosen["bottom"] > 880:
                self.page.mouse.move(850, 450)
                self.page.mouse.wheel(0, 280)
                time.sleep(0.8)
                continue

            return chosen

        targets = self._find_channel_targets_js()
        quizzes = targets.get("quizzes", [])
        if item_type == "video":
            return targets.get("video") or (quizzes[0] if quizzes else None)
        return quizzes[-2:][quiz_index] if len(quizzes) >= 2 else (quizzes[0] if quizzes else None)

    def _click_forward_on_bubble(self, target_bubble):
        """Triggers the Forward modal for either a Video post or a Quiz post."""
        print(f"  🎯 Target Post: '{target_bubble['preview']}...' (top={target_bubble['top']}, bottom={target_bubble['bottom']})")

        safe_hover_x = target_bubble["left"] + 80
        safe_hover_y = target_bubble["top"] + 22
        self.page.mouse.move(safe_hover_x, safe_hover_y)
        time.sleep(0.6)

        # Method 1: Hover top-right of the green card (x ≈ 1148) & trigger Forward
        for tr_x in [1148, target_bubble["right"] - 18]:
            tr_y = target_bubble["top"] + 18
            print(f"  🖱️ Opening Forward from top-right ({tr_x}, {tr_y})...")
            self.page.mouse.move(tr_x, tr_y)
            time.sleep(0.6)
            self.page.mouse.click(tr_x, tr_y)
            time.sleep(1.0)
            if self._check_and_advance_to_modal():
                return

        # Method 2: Click the circular Forward icon outside/beside the bubble
        side_buttons = self.page.evaluate("""(qRect) => {
            const allBtns = Array.from(document.querySelectorAll('#main button, #main [role="button"], #main [data-icon]'));
            const results = [];
            for (const b of allBtns) {
                const r = b.getBoundingClientRect();
                if (r.width === 0 || r.height === 0 || r.top < 105) continue;
                if (r.top < qRect.top - 15 || r.bottom > qRect.bottom + 65) continue;

                const cx = Math.round(r.left + r.width / 2);
                const cy = Math.round(r.top + r.height / 2);
                if (cx > 730 && cx < 1120) continue;

                const aria = (b.getAttribute('aria-label') || b.getAttribute('title') || '').trim();
                const text = (b.innerText || '').trim();
                if (text.includes('View responses') || aria.includes('View poll voters') || aria.toLowerCase().includes('react')) continue;

                results.push({ x: cx, y: cy, aria: aria });
            }
            return results;
        }""", target_bubble)

        for btn in side_buttons:
            print(f"  🖱️ Trying side/bottom-left button at ({btn['x']}, {btn['y']})...")
            self.page.mouse.click(btn["x"], btn["y"])
            time.sleep(1.2)
            if self._check_and_advance_to_modal():
                return

        # Method 3: Right-click context menu on the top edge of the bubble
        print(f"  🖱️ Right-clicking post header at ({safe_hover_x}, {safe_hover_y})...")
        self.page.mouse.click(safe_hover_x, safe_hover_y, button="right")
        time.sleep(1.2)
        if self._check_and_advance_to_modal():
            return

        raise Exception(f"Could not open Forward modal for post: {target_bubble['preview']}")

    def _select_contact_in_modal(self, search_input, contact):
        """
        Searches the contact inside the Forward popup and physically clicks the first matching
        contact row/checkbox directly below the search bar.
        """
        queries_to_try = [contact["raw_number"]]
        if len(contact["digits"]) >= 10:
            last_10 = contact["digits"][-10:]
            if last_10 not in queries_to_try:
                queries_to_try.append(last_10)
        if contact["name"] and contact["name"] not in queries_to_try:
            queries_to_try.append(contact["name"])

        s_box = search_input.bounding_box()
        if not s_box:
            return False

        for q in queries_to_try:
            search_input.click()
            self.page.keyboard.press("Control+A")
            self.page.keyboard.press("Backspace")
            time.sleep(0.3)

            search_input.fill(q)
            time.sleep(1.5)

            # Inspect the area directly below the search box inside the Forward modal
            result_info = self.page.evaluate("""(sRect) => {
                // Check if 'No results' is visible directly below the search box
                const allNodes = Array.from(document.querySelectorAll('div, span'));
                for (const n of allNodes) {
                    const r = n.getBoundingClientRect();
                    if (r.width === 0 || r.height === 0) continue;
                    if (r.left >= sRect.x - 60 && r.right <= sRect.x + sRect.width + 60 &&
                        r.top >= sRect.y + sRect.height && r.top <= sRect.y + 250) {
                        const t = (n.innerText || '').trim().toLowerCase();
                        if (t === 'no results found' || t.startsWith('no results') || t.startsWith('no chats')) {
                            return { empty: true };
                        }
                    }
                }

                // Look for clickable contact rows or checkboxes directly below the search box
                const candidates = Array.from(document.querySelectorAll(
                    '[role="checkbox"], [role="listitem"], [role="row"], [data-animate-modal-body="true"] [role="button"]'
                ));
                const validRows = [];

                for (const c of candidates) {
                    const r = c.getBoundingClientRect();
                    if (r.width < 15 || r.height < 15) continue;
                    // Must be horizontally aligned with the search box and vertically below it
                    if (r.left >= sRect.x - 80 && r.left <= sRect.x + sRect.width + 50 &&
                        r.top >= sRect.y + sRect.height + 10 && r.top <= sRect.y + 320) {
                        validRows.push({
                            x: Math.round(r.left + Math.min(r.width / 2, 120)),
                            y: Math.round(r.top + r.height / 2),
                            top: r.top
                        });
                    }
                }

                validRows.sort((a, b) => a.top - b.top);
                if (validRows.length > 0) {
                    return { empty: false, foundRow: true, x: validRows[0].x, y: validRows[0].y };
                }

                // Fallback: Check if any contact name text is rendered in the first row spot (+85px below search box)
                const fallbackX = Math.round(sRect.x + 120);
                const fallbackY = Math.round(sRect.y + sRect.height + 85);
                const elAtPoint = document.elementFromPoint(fallbackX, fallbackY);
                if (elAtPoint) {
                    const rowText = (elAtPoint.closest('[role="listitem"], [role="row"], div') || elAtPoint).innerText || '';
                    if (rowText.trim().length > 1 && !rowText.toLowerCase().includes('no results')) {
                        return { empty: false, foundRow: true, x: fallbackX, y: fallbackY };
                    }
                }

                return { empty: true };
            }""", s_box)

            if result_info.get("empty"):
                continue

            if result_info.get("foundRow"):
                click_x = result_info["x"]
                click_y = result_info["y"]
                print(f"     ✅ Found match for '{q}'! Clicking checkbox row at ({click_x}, {click_y})...")
                self.page.mouse.click(click_x, click_y)
                time.sleep(0.6)
                return True

        return False

    def forward_item_to_batch(self, item_type, quiz_index, batch_contacts):
        self.page.keyboard.press("Escape")
        time.sleep(0.5)

        label = "Video / Top Post" if item_type == "video" else ("Quiz 1 (Older)" if quiz_index == 0 else "Quiz 2 (Latest)")
        print(f"\n🔄 Selecting {label} for batch of {len(batch_contacts)} contacts...")

        target_bubble = self._scroll_target_into_comfort_zone(item_type, quiz_index)
        if not target_bubble:
            raise Exception(f"Could not locate {label} inside the Channel.")

        self._click_forward_on_bubble(target_bubble)

        search_input = self._get_forward_search_box()
        if search_input is None:
            raise Exception("Forward popup did not open.")

        succeeded_contacts = []

        for contact in batch_contacts:
            print(f"  -> Selecting: {contact['name']} ({contact['raw_number']}) [Row {contact['excel_row']} | {contact['type']}]")
            if self._select_contact_in_modal(search_input, contact):
                succeeded_contacts.append(contact)
            else:
                print(f"     ❌ Could not find '{contact['name']}' in this WhatsApp account.")
                logging.warning(f"Contact not found in WhatsApp: {contact['name']} ({contact['raw_number']})")

        if succeeded_contacts:
            # Find and click the green Send button at the bottom-right of the Forward popup
            send_btn = self.page.locator('[data-icon="send"], [aria-label="Send"], [data-icon*="send"]').last
            if send_btn.is_visible():
                box = send_btn.bounding_box()
                if box:
                    self.page.mouse.click(box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
                else:
                    send_btn.click(force=True)
                print(f"✅ Forwarded {label} to {len(succeeded_contacts)} people!")
                time.sleep(3)
                return succeeded_contacts
            else:
                self.page.keyboard.press("Enter")
                print(f"✅ Forwarded {label} to {len(succeeded_contacts)} people!")
                time.sleep(3)
                return succeeded_contacts
        else:
            print("  ⚠️ None of the 5 contacts in this batch exist on this WhatsApp. Closing popup...")
            self.page.keyboard.press("Escape")
            time.sleep(1)
            return []

    def close(self):
        self.browser.close()
        self.playwright.stop()


def run_quiz_automation():
    print("=" * 60)
    print("🚀 FLUENCE CHANNEL VIDEO & QUIZ FORWARDER")
    print("=" * 60)

    excel_path = resolve_excel_path()
    day_num = get_target_day()
    contacts = load_and_sort_contacts(excel_path, SHEET_NAME, day_num)

    if not contacts:
        print(f"🎉 All eligible contacts are already completed for Day {day_num}! Nothing to send.")
        return

    bot = WhatsAppQuizForwarder(PROFILE_DIR)

    try:
        bot.open_channel(CHANNEL_NAME, CHANNEL_SHORT_KEY)

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
            bot.open_channel(CHANNEL_NAME, CHANNEL_SHORT_KEY)
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
