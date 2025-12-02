# IIMK Class Schedule → Google Calendar Automation

A simple Google Sheets + Apps Script setup that enables IIM Kozhikode students to automatically sync their class timetable into **Google Calendar** using their selected subjects.  
No duplicates. No manual event creation. Works individually for every student.

---

## 🌟 Brief Value Proposition
- Automatically create Google Calendar events for all your classes  
- Always stay updated when the official schedule changes  
- Save time by avoiding manual entry of 50–100 class events  
- Prevent missed classes and scheduling conflicts  
- Personalized to each student through the **Subjects** sheet  
- Safe to re-run anytime — no duplicate events  

---

## 📌 Description
This automation uses a combination of:
- **Main Sheet** — the official IIMK schedule (imported or pasted)
- **Subjects** — your personal course list (Track + Course Code)
- **Apps Script** — to filter your schedule and create calendar events

The script:
1. Builds your personalized **My Schedule**
2. Creates Google Calendar events for each class
3. Stores event IDs to avoid duplicates on future runs

---

## 📚 Steps for Students (Complete Setup Guide)

### **1. Create Your Google Sheet**
Create a new Google Sheet with two tabs:

- `Main Sheet`
- `Subjects`
  <img width="1070" height="81" alt="image" src="https://github.com/user-attachments/assets/9c472175-9834-4b4b-9ed6-b7401e830efe" />


---

### **2. Add Your Subjects**
Open the `Subjects` tab and fill in:

| Course Code | Track |
|-------------|-------|
| EECO-002    | 1     |
| EHLAM-002   | 2     |

Both values **must match** Main Sheet for the script to detect your classes.
<img width="979" height="67" alt="image" src="https://github.com/user-attachments/assets/eaa0af94-ed8e-4daf-abb6-b8851db3d4ee" />


---

### **3. Add the Automation Script**
1. Go to **Extensions → Apps Script**
2. Delete the default code
3. Paste the complete `Code.gs` from this repository
4. Save the script

---

### **4. Authorize the Script**
1. In Apps Script, select the function `runAll_`
2. Press **Run**
3. Grant Sheets + Calendar permissions when prompted
4. Refresh your Google Sheet

You will now see a menu:  
**Schedule Automation**

---

### **5. Run the Automation**
Use:

**Schedule Automation → Run All (1 → 2)**

This will:
1. Generate/update **My Schedule** from Main Sheet + Subjects  
2. Create Google Calendar events (only where event IDs are empty)

Run this whenever:
- Your subjects change  
- IIMK updates the official schedule  
- You want to refresh your calendar  

Event IDs ensure events are **never duplicated**.

---

## 📝 Notes
- The script auto-adds a `Calendar Event ID` column in `Main Sheet`.
- You can re-run anytime — only new class rows create new events.
- Existing events are not touched (no updates/deletes yet).

---

## 📄 Files
- `Code.gs` — the Apps Script code  
- `README.md` — this document  

---

## 📜 License
MIT License
