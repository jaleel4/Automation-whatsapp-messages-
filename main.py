import os
import time
import datetime
import logging
import requests
from dotenv import load_dotenv
from course_scheduler import get_course_day, get_module_and_day
from notion_reader import NotionCourseReader
from whatsapp_sender import WhatsAppSender

# Set up logging to track everything quietly in a file
logging.basicConfig(
    filename=r"C:\EnglishCourseAutomation\automation.log", 
    level=logging.INFO, 
    format='%(asctime)s - %(message)s'
)

def run_daily_automation():
    load_dotenv(dotenv_path=r"C:\EnglishCourseAutomation\.env")
    token = os.getenv("NOTION_TOKEN")
    db_id = "4a36c0fd-2ef7-4eef-a400-94e409490d94"

    print("Initializing Notion Reader...")
    reader = NotionCourseReader(token)
    
    print("Initializing WhatsApp Sender (Keep the browser window open!)...")
    profile_dir = r"C:\EnglishCourseAutomation\whatsapp_profile"
    whatsapp = WhatsAppSender(profile_dir)

    print("\n🚀 Automation System is LIVE.")
    print("The system is now checking your Notion database every minute.")
    print("Do not close this terminal or the WhatsApp browser.")
    print("============================================================\n")

    # Bulletproof Headers for direct API calls
    notion_headers = {
        "Authorization": f"Bearer {token}",
        "Notion-Version": "2022-06-28",
        "Content-Type": "application/json"
    }

    while True:
        now = datetime.datetime.now()
        today_str = now.strftime("%Y-%m-%d")
        current_time_str = now.strftime("%H:%M") 
        
        # 1. Check if it's the weekend. If so, do nothing!
        if now.weekday() >= 5: # 5 = Saturday, 6 = Sunday
            print(f"[{current_time_str}] It's the weekend. Automation is resting...")
            time.sleep(3600) # Sleep for an hour
            continue

        try:
            # 2. Pull all Active students from Notion using a Direct HTTP Request
            query_payload = {
                "filter": {
                    "property": "Active",
                    "checkbox": {"equals": True}
                }
            }
            
            response = requests.post(
                f"https://api.notion.com/v1/databases/{db_id}/query", 
                headers=notion_headers, 
                json=query_payload
            )
            
            if response.status_code != 200:
                print(f"⚠️ API Error querying students: {response.text}")
                time.sleep(60)
                continue
                
            students = response.json().get("results", [])

            for student in students:
                props = student.get("properties", {})
                
                # Extract Student Data Safely
                name = props["Name"]["title"][0]["plain_text"] if props.get("Name", {}).get("title") else "Unknown"
                
                # Automatically handles whether the column is set to Number (#) or Text (Aa) in Notion
                number_prop = props.get("WhatsApp Number", {})
                if number_prop.get("type") == "number":
                    number = str(number_prop.get("number")).replace('.0', '') if number_prop.get("number") else ""
                else:
                    rt = number_prop.get("rich_text", [])
                    number = rt[0]["plain_text"].strip() if rt else ""
                
                start_date_prop = props.get("Start Date", {}).get("date")
                start_date = start_date_prop["start"] if start_date_prop else ""
                
                last_sent_prop = props.get("Last Sent", {}).get("date")
                last_sent = last_sent_prop["start"] if last_sent_prop else ""

                status_prop = props.get("Status", {}).get("select")
                current_status = status_prop["name"] if status_prop else ""

                # Handles lowercase 'send time' or uppercase 'Send Time'
                send_time_prop = props.get("send time") or props.get("Send Time", {})
                send_time = send_time_prop.get("rich_text", [])[0]["plain_text"].strip() if send_time_prop.get("rich_text") else "09:00"

                # Skip if missing crucial contact info
                if not number or not start_date:
                    continue 

                # ==========================================
                # 🌟 FREE MIDNIGHT RESET LOGIC 🌟
                # ==========================================
                # If it's a new day and they aren't marked as Pending, reset them in Notion!
                if last_sent != today_str and current_status != "Pending":
                    print(f"🔄 Midnight Reset: Changing {name}'s status back to 'Pending'.")
                    try:
                        reset_payload = {
                            "properties": {
                                "Status": {"select": {"name": "Pending"}},
                                "Last Sent": None  # This wipes the old date completely
                            }
                        }
                        requests.patch(
                            f"https://api.notion.com/v1/pages/{student['id']}", 
                            headers=notion_headers, 
                            json=reset_payload
                        )
                        # Update local variables so the script knows they are ready for today
                        last_sent = ""
                        current_status = "Pending"
                    except Exception as e:
                        print(f"⚠️ Could not reset status for {name}: {e}")
                # ==========================================

                # 3. Prevent duplicate sends on the same day
                if last_sent == today_str:
                    continue 

                # 4. Check if it is time to send!
                if current_time_str >= send_time:
                    
                    course_day = get_course_day(start_date, today_str)
                    
                    if course_day < 1 or course_day > 40:
                        continue 

                    module, day = get_module_and_day(course_day)
                    
                    print(f"\n⏰ [{current_time_str}] It is time for {name}!")
                    print(f"📚 Calculated Course Day: {course_day} -> (Module {module}, Day {day})")

                    try:
                        # Find page in Notion
                        page_id = reader.find_course_day_page(module, day)
                        blocks = reader.parse_blocks(page_id)

                        # Send via WhatsApp
                        whatsapp.open_chat(number)

                        for block in blocks:
                            if block["type"] == "text":
                                whatsapp.send_text(block["content"])
                            elif block["type"] == "image":
                                whatsapp.send_image(block["path"])
                            elif block["type"] == "poll":
                                whatsapp.send_poll(block["question"], block["options"])

                        # 5. Update Notion Database (Direct HTTP Request)
                        update_payload = {
                            "properties": {
                                "Last Sent": {"date": {"start": today_str}},
                                "Status": {"select": {"name": "Sent"}} 
                            }
                        }
                        update_response = requests.patch(
                            f"https://api.notion.com/v1/pages/{student['id']}", 
                            headers=notion_headers, 
                            json=update_payload
                        )
                        
                        if update_response.status_code != 200:
                            print(f"⚠️ Message sent, but Notion update failed: {update_response.text}")
                        else:
                            print(f"✅ Successfully completed daily automation for {name}.")
                        
                        logging.info(f"Success: {name} | Course Day {course_day}")

                    except Exception as e:
                        print(f"❌ Failed for {name}: {str(e)}")
                        logging.error(f"Failed: {name} | Error: {str(e)}")
                        
                        # Mark as Failed in Notion (Direct HTTP Request)
                        try:
                            fail_payload = {
                                "properties": {
                                    "Status": {"select": {"name": "Failed"}} 
                                }
                            }
                            requests.patch(
                                f"https://api.notion.com/v1/pages/{student['id']}", 
                                headers=notion_headers, 
                                json=fail_payload
                            )
                        except:
                            pass

        except Exception as e:
            print(f"⚠️ Network error checking Notion database: {e}")

        # Wait exactly 60 seconds before checking the clock again
        time.sleep(60)

if __name__ == "__main__":
    run_daily_automation()
