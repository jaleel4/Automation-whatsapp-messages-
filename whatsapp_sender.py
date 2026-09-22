from playwright.sync_api import sync_playwright
import time

class WhatsAppSender:
    def __init__(self, profile_dir):
        self.profile_dir = profile_dir
        self.playwright = sync_playwright().start()
        self.browser = self.playwright.chromium.launch_persistent_context(
            user_data_dir=self.profile_dir,
            headless=False,
            args=["--start-maximized"]
        )
        self.page = self.browser.pages[0]
        self.page.goto("https://web.whatsapp.com/")
        print("Waiting for WhatsApp Web to load...")
        
        try:
            self.page.wait_for_selector('#pane-side', timeout=20000)
        except:
            print("Please scan the WhatsApp QR code if prompted. Waiting up to 2 minutes...")
            self.page.wait_for_selector('#pane-side', timeout=120000)
            
        print("WhatsApp Web is ready.")
        time.sleep(2)

    def open_chat(self, chat_name):
        print(f"Opening chat: {chat_name}")
        search_box = self.page.locator('#side [contenteditable="true"], #side input[type="text"]').first
        search_box.wait_for(state="visible", timeout=15000)
        
        search_box.click()
        self.page.keyboard.press("Control+A")
        self.page.keyboard.press("Backspace")
        time.sleep(1)
        
        search_box.fill(chat_name)
        time.sleep(2)
        
        chat = self.page.locator('#pane-side').get_by_text(chat_name, exact=True).first
        
        if chat.is_visible():
            chat.click(force=True)
        else:
            print("Exact text match not found, pressing Enter to open top result...")
            self.page.keyboard.press("Enter")
            
        time.sleep(2)

    def get_composer(self):
        try:
            composer = self.page.locator('div[contenteditable="true"][role="textbox"][data-tab="10"]').last
            composer.wait_for(state="visible", timeout=5000)
            return composer
        except:
            composer = self.page.locator('footer div[contenteditable="true"]').last
            composer.wait_for(state="visible", timeout=5000)
            return composer

    def _open_attachment_menu(self):
        composer = self.get_composer()
        composer.click()
        time.sleep(1)

        footer = self.page.locator('#main footer').first
        attach_btn = footer.locator(
            '[title="Attach"], '
            '[aria-label="Attach"], '
            '[data-icon="plus"], '
            '[data-icon="attach-menu-plus"]'
        ).first
        
        attach_btn.wait_for(state="visible", timeout=10000)
        attach_btn.click(force=True)
        time.sleep(1.5)

    def _click_menu_item(self, text_labels, data_icons):
        for _ in range(4):
            clicked = self.page.evaluate('''([texts, icons]) => {
                const elements = Array.from(document.querySelectorAll('span, div, li, svg'));
                for (let i = elements.length - 1; i >= 0; i--) {
                    const el = elements[i];
                    const txt = (el.innerText || el.textContent || '').trim();
                    const iconAttr = el.getAttribute('data-icon');
                    
                    if (texts.includes(txt) || icons.includes(iconAttr)) {
                        const target = el.closest('li') || el.closest('[role="button"]') || el;
                        target.click();
                        return true;
                    }
                }
                return false;
            }''', [text_labels, data_icons])
            
            if clicked:
                return True
            time.sleep(1)
            
        return False

    def send_text(self, text):
        composer = self.get_composer()
        composer.click()
        composer.focus()
        self.page.keyboard.insert_text(text)
        time.sleep(1)
        self.page.keyboard.press("Enter")
        time.sleep(1)

    def send_image(self, image_path):
        print("Attaching photo...")
        self._open_attachment_menu()

        clicked = self._click_menu_item(["Photos & videos", "Photos and videos", "Photos"], ["image"])
        if not clicked:
            print("⚠️ Could not click Photos button visually, attempting direct input injection...")
        
        time.sleep(1)
        
        image_input = self.page.locator('input[type="file"][accept*="image"]').last
        image_input.wait_for(state="attached", timeout=10000)
        
        print("Selecting photo file...")
        image_input.set_input_files(image_path)
        time.sleep(3)
        
        print("Photo preview loaded.")
        
        # Explicitly click the Send button on the photo preview
        send_btn = self.page.locator('[data-icon="send"], [aria-label="Send"]').last
        if send_btn.is_visible():
            send_btn.click(force=True)
        else:
            self.page.keyboard.press("Enter")
            
        time.sleep(4)
        
        # Safety Check: If the photo preview got stuck open, forcefully close it so the script doesn't crash!
        add_file_btn = self.page.locator('[aria-label="Add file"]').first
        if add_file_btn.is_visible():
            print("⚠️ Photo preview stuck open! Attempting force send...")
            self.page.keyboard.press("Enter")
            time.sleep(2)
            if add_file_btn.is_visible():
                print("❌ Could not send photo. Closing preview to save the rest of the day.")
                self.page.keyboard.press("Escape")
                time.sleep(1)

        print("✅ Photo sent.")

    def send_poll(self, question, options):
        print("Poll window opening...")
        
        clean_question = question.strip()
        if '\n' in clean_question:
            parts = [p.strip() for p in clean_question.split('\n') if p.strip()]
            clean_question = parts[0]
            if not options and len(parts) > 1:
                options = parts[1:]
                
        self._open_attachment_menu()

        clicked = self._click_menu_item(["Poll", "Create poll"], ["poll", "poll-outline"])
        if not clicked:
            raise Exception("Could not find the Poll icon in the attachment menu.")

        modal = self.page.locator('div[role="dialog"][aria-modal="true"]').last
        modal.wait_for(state="visible", timeout=10000)
        time.sleep(1)

        inputs = modal.locator('div[contenteditable="true"], input[type="text"]')
        inputs.first.wait_for(state="visible", timeout=5000)
        
        input_elements = inputs.all()
        if len(input_elements) < 2:
            raise Exception("Could not find enough inputs inside the poll modal.")

        # Fill Question
        input_elements[0].click()
        self.page.keyboard.insert_text(clean_question)
        time.sleep(0.5)

        # Fill Options dynamically
        for i, opt in enumerate(options):
            clean_opt = opt.strip().replace('\n', ' ')
            if not clean_opt:
                continue
                
            current_inputs = modal.locator('div[contenteditable="true"], input[type="text"]').all()
            target_idx = i + 1
            
            if target_idx < len(current_inputs):
                current_inputs[target_idx].click()
            else:
                current_inputs[-1].click()
                
            self.page.keyboard.insert_text(clean_opt)
            time.sleep(0.5)

        # Disable Multiple Answers
        switch = modal.locator('[role="switch"], [role="checkbox"], input[type="checkbox"]').last
        if switch.is_visible():
            try:
                if switch.get_attribute('aria-checked') == 'true':
                    switch.click(force=True)
            except:
                pass
        time.sleep(1)

        print("Sending poll...")
        send_btn = self.page.locator('[data-icon="send"], [aria-label="Send"]').last
        
        if send_btn.is_visible():
            send_btn.click(force=True)
        else:
            current_inputs = modal.locator('div[contenteditable="true"], input[type="text"]').all()
            if current_inputs:
                current_inputs[-1].click(force=True)
                time.sleep(0.5)
            self.page.keyboard.press("Enter")

        try:
            modal.wait_for(state="hidden", timeout=5000)
            print("✅ WhatsApp Poll sent!")
        except:
            print("⚠️ Poll modal stuck open. Forcing close...")
            cancel_btn = self.page.locator('[data-icon="x"], [aria-label="Cancel"], [aria-label="Close"], [data-icon="back"]').first
            if cancel_btn.is_visible():
                cancel_btn.click(force=True)
                
            self.page.keyboard.press("Escape")
            time.sleep(1)
            self.page.keyboard.press("Escape")

    def close(self):
        self.browser.close()
        self.playwright.stop()