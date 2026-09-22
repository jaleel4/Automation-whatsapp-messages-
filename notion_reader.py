from notion_client import Client
import os
import requests
import uuid
import time

class NotionCourseReader:
    def __init__(self, token):
        self.client = Client(auth=token)
        self.downloads_dir = r"C:\EnglishCourseAutomation\downloads"
        os.makedirs(self.downloads_dir, exist_ok=True)

    def get_blocks(self, block_id):
        blocks = []
        cursor = None
        while True:
            response = self.client.blocks.children.list(block_id=block_id, start_cursor=cursor)
            blocks.extend(response.get("results", []))
            cursor = response.get("next_cursor")
            if not cursor:
                break
        return blocks

    def _get_title(self, block):
        b_type = block.get("type", "")
        title = ""
        if b_type == "child_page":
            title = block["child_page"].get("title", "")
        elif b_type == "child_database":
            title = block["child_database"].get("title", "")
        elif b_type == "toggle" and block.get("toggle", {}).get("rich_text"):
            title = "".join([t.get("plain_text", "") for t in block["toggle"]["rich_text"]])
        elif b_type.startswith("heading_") and block.get(b_type, {}).get("rich_text"):
            title = "".join([t.get("plain_text", "") for t in block[b_type]["rich_text"]])
            
        return title.lower()

    def _deep_scan_for_keywords(self, block_id, keywords):
        blocks = self.get_blocks(block_id)
        for b in blocks:
            title = self._get_title(b)
            if title and any(kw in title for kw in keywords):
                return b["id"]
            
            if b.get("has_children"):
                found = self._deep_scan_for_keywords(b["id"], keywords)
                if found:
                    return found
        return None

    def find_course_day_page(self, module_num, day_num):
        module_kws = [f"module {module_num}", f"module {module_num:02d}"]
        
        course_day = (module_num - 1) * 2 + day_num
        day_kws = [
            f"day {day_num:02d}", f"day {day_num}", 
            f"day {course_day:02d}", f"day {course_day}"
        ]
        
        try:
            mod1_block = self.client.blocks.retrieve("3dede9ba-0af7-8056-9873-db8257f4b67e")
            parent = mod1_block.get("parent", {})
            parent_type = parent.get("type", "")
            parent_id = parent.get(parent_type, "").replace("-", "")
        except Exception as e:
            raise Exception(f"Failed to access the anchor block. Ensure token is correct. Error: {e}")

        target_module_id = self._deep_scan_for_keywords(parent_id, module_kws)
        
        if not target_module_id:
            if module_num == 1: 
                target_module_id = "3dede9ba-0af7-8056-9873-db8257f4b67e"
            else: 
                raise Exception(f"Could not find '{module_kws[0]}' inside the course root.")

        target_day_id = self._deep_scan_for_keywords(target_module_id, day_kws)

        if not target_day_id:
            if module_num == 1 and day_num == 2: 
                target_day_id = "92ede9ba-0af7-8367-a6f3-8117553aa985"
            else: 
                raise Exception(f"Could not find '{day_kws[0]}' or '{day_kws[2]}' strictly inside Module {module_num}.")

        return target_day_id

    def parse_rich_text(self, rich_text_list):
        text = ""
        for rt in rich_text_list:
            content = rt.get("plain_text", "")
            if rt.get("annotations", {}).get("bold"):
                content = f"*{content}*"
            text += content
        return text

    def get_emoji_number(self, num):
        emoji_map = {1: "1️⃣", 2: "2️⃣", 3: "3️⃣", 4: "4️⃣", 5: "5️⃣", 
                     6: "6️⃣", 7: "7️⃣", 8: "8️⃣", 9: "9️⃣"}
        return emoji_map.get(num, f"{num}.")

    def parse_blocks(self, block_id, list_counter=1, is_root=True, inside_toggle=False):
        parsed = []
        blocks = self.get_blocks(block_id)

        for b in blocks:
            b_type = b["type"]
            
            if b_type == "paragraph":
                text = self.parse_rich_text(b["paragraph"]["rich_text"]).strip()
                if not text:
                    continue
                
                force_break = False
                
                if "Todays Challenges" in text or "Today's Challenges" in text:
                    text = "*📚 Today's Challenges:*"
                    force_break = True
                elif "Zoom URL" in text and "https://" not in text:
                    continue 
                elif "zoom.us" in text.lower():
                    clean_text = text.replace('Zoom URL:', '').replace('Zoom URL', '').strip()
                    text = f"*🔗 Zoom Meeting:*\n{clean_text}"
                    force_break = True

                # Single-block poll detection (Only run if NOT inside a toggle)
                lines = [l.strip() for l in text.split('\n') if l.strip()]
                if len(lines) >= 3 and any("?" in l for l in lines[:2]) and not inside_toggle:
                    q_idx = 0 if "?" in lines[0] else 1
                    question = " ".join(lines[:q_idx+1])
                    options = lines[q_idx+1:]
                    if len(options) >= 2:
                        parsed.append({"type": "poll", "question": question, "options": options})
                        continue

                parsed.append({"type": "text", "content": text, "force_break": force_break, "inside_toggle": inside_toggle})
                list_counter = 1

            elif b_type == "numbered_list_item":
                text = self.parse_rich_text(b["numbered_list_item"]["rich_text"])
                prefix = self.get_emoji_number(list_counter)
                parsed.append({"type": "text", "content": f"{prefix} {text}", "inside_toggle": inside_toggle})
                list_counter += 1

            elif b_type == "bulleted_list_item":
                text = self.parse_rich_text(b["bulleted_list_item"]["rich_text"])
                parsed.append({"type": "text", "content": f"• {text}", "inside_toggle": inside_toggle})
                list_counter = 1

            elif b_type == "toggle":
                text = self.parse_rich_text(b["toggle"]["rich_text"])
                parsed.append({"type": "text", "content": f"*{text}*", "force_break": True, "inside_toggle": inside_toggle})
                if b["has_children"]:
                    # Pass inside_toggle=True so children never become polls
                    parsed.extend(self.parse_blocks(b["id"], is_root=False, inside_toggle=True))
                list_counter = 1

            elif b_type == "image":
                url = b["image"].get("file", {}).get("url") or b["image"].get("external", {}).get("url")
                if url:
                    ext = url.split("?")[0].split(".")[-1]
                    if len(ext) > 4: ext = "jpeg"
                    filename = f"{uuid.uuid4().hex}.{ext}"
                    filepath = os.path.join(self.downloads_dir, filename)
                    
                    success = False
                    for attempt in range(3):
                        try:
                            response = requests.get(url, timeout=20)
                            response.raise_for_status()
                            with open(filepath, 'wb') as f:
                                f.write(response.content)
                            success = True
                            break
                        except Exception as e:
                            print(f"⚠️ Image download attempt {attempt+1} failed ({e}). Retrying...")
                            time.sleep(2)
                            
                    if success:
                        parsed.append({"type": "image", "path": filepath})
                    else:
                        print("❌ Failed to download image after 3 attempts.")
                list_counter = 1

            elif b_type == "video":
                url = b["video"].get("external", {}).get("url", "")
                if url:
                    parsed.append({"type": "text", "content": f"*🎥 Video:* {url}", "force_break": True})
                list_counter = 1

            elif b_type == "link_preview" or b_type == "bookmark":
                url = b.get(b_type, {}).get("url", "")
                if url:
                    parsed.append({"type": "text", "content": f"*🔗 Link:* {url}", "force_break": True})
                list_counter = 1

        if is_root:
            # --- MULTI-BLOCK POLL DETECTOR ---
            optimized_parsed = []
            i = 0
            while i < len(parsed):
                item = parsed[i]
                
                # ONLY create polls if it's NOT inside a toggle
                if item["type"] == "text" and "?" in item["content"] and not item.get("force_break") and not item.get("inside_toggle"):
                    options = []
                    j = i + 1
                    while j < len(parsed):
                        next_item = parsed[j]
                        # Don't grab options that are inside toggles either
                        if next_item["type"] == "text" and not next_item.get("force_break") and not next_item.get("inside_toggle"):
                            content = next_item["content"]
                            if len(content) < 60 and "?" not in content and "http" not in content and not content.endswith('.'):
                                options.append(content)
                                j += 1
                                continue
                        break
                        
                    if len(options) >= 2:
                        optimized_parsed.append({
                            "type": "poll",
                            "question": item["content"],
                            "options": options
                        })
                        i = j
                        continue
                        
                optimized_parsed.append(item)
                i += 1
                
            parsed = optimized_parsed

            # --- SMART TEXT MERGING ENGINE ---
            merged_results = []
            current_text_group = []
            
            for item in parsed:
                if item["type"] == "text":
                    if item.get("force_break") and current_text_group:
                        merged_results.append({
                            "type": "text", 
                            "content": "\n\n".join(current_text_group)
                        })
                        current_text_group = []
                    
                    current_text_group.append(item["content"])
                else:
                    if current_text_group:
                        merged_results.append({
                            "type": "text", 
                            "content": "\n\n".join(current_text_group)
                        })
                        current_text_group = []
                    
                    merged_results.append(item)
            
            if current_text_group:
                merged_results.append({
                    "type": "text", 
                    "content": "\n\n".join(current_text_group)
                })
                
            return merged_results

        return parsed