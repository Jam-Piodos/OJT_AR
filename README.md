# NBSC OJT Weekly Report System
## Form 9 — Northern Bukidnon State College

---

## 📁 FILE STRUCTURE

```
ojt-system/
├── index.html     ← Main web app (open this in browser)
├── style.css      ← All styling
├── script.js      ← Frontend logic (modals, export, API calls)
├── Code.gs        ← Google Apps Script backend
└── README.md      ← This file
```

---

## 🚀 QUICK START (Frontend Only — No Backend)

1. Open `index.html` in any modern browser.
2. Fill in the form fields.
3. Use **Add Objective / Activity / Reflection** buttons.
4. Click **Upload Images** to add documentation photos.
5. Export via **PDF, Print, or DOCX** buttons in the sidebar.
6. Data auto-saves to browser localStorage every 30 seconds.

---

## ☁️ GOOGLE SHEETS + DRIVE SETUP

### Step 1 — Create Google Sheet

1. Go to [sheets.google.com](https://sheets.google.com)
2. Create a new spreadsheet named **OJT Reports**
3. Note the **Sheet ID** from the URL:
   ```
   https://docs.google.com/spreadsheets/d/SHEET_ID_HERE/edit
   ```

### Step 2 — Create Google Drive Folder

1. Go to [drive.google.com](https://drive.google.com)
2. Create a folder named **OJT Documentation**
3. Note the **Folder ID** from the URL:
   ```
   https://drive.google.com/drive/folders/FOLDER_ID_HERE
   ```

### Step 3 — Set Up Google Apps Script

1. Open your Google Sheet → **Extensions → Apps Script**
2. Delete all default code
3. Paste the entire contents of `Code.gs`
4. Update these two lines:
   ```javascript
   const SHEET_ID  = 'YOUR_GOOGLE_SHEET_ID_HERE';
   const FOLDER_ID = 'YOUR_DRIVE_FOLDER_ID_HERE';
   ```
5. Click **Save** (Ctrl+S)

### Step 4 — Deploy as Web App

1. Click **Deploy → New Deployment**
2. Click the gear icon ⚙️ → Select type: **Web app**
3. Set:
   - **Description**: OJT Report API
   - **Execute as**: Me (your Google account)
   - **Who has access**: Anyone
4. Click **Deploy**
5. Authorize permissions when prompted
6. Copy the **Web App URL** (starts with `https://script.google.com/macros/s/…/exec`)

### Step 5 — Connect to Frontend

1. Open `index.html` in your browser
2. Find the **Google Sheets Backend** panel in the right sidebar
3. Paste the Web App URL
4. Click **Connect**

---

## 🗃️ GOOGLE SHEET STRUCTURE

The sheet will be auto-created with these columns:

| Column | Description |
|--------|-------------|
| Timestamp | ISO date of last save |
| Name | Trainee's full name |
| Company | HTE / Company name |
| Week | Week 1–12 |
| Objectives | JSON array of objectives |
| Activities | JSON array of activities |
| Reflections | JSON array of reflections |
| Image URLs | JSON array of Drive file URLs |
| Sig_Trainee | Student signature name |
| Sig_Supervisor | HTE Supervisor name |
| Sig_Coordinator | OJT Coordinator name |

---

## 📡 API ENDPOINTS

### Save Report
```
POST {SCRIPT_URL}
Body: {
  "action": "saveReport",
  "name": "Juan Dela Cruz",
  "company": "Acme Corp",
  "week": "Week 1",
  "objectives": "[{\"text\":\"...\",\"status\":\"In Progress\"}]",
  "activities": "[{\"title\":\"...\",\"desc\":\"...\"}]",
  "reflections": "[{\"type\":\"Learning\",\"content\":\"...\"}]",
  "imageUrls": "[\"https://...\"]",
  "sig_trainee": "Juan Dela Cruz",
  "sig_supervisor": "Maria Santos",
  "sig_coordinator": "Dr. Pedro Reyes",
  "timestamp": "2025-01-15T08:00:00.000Z"
}
```

### Get Report
```
GET {SCRIPT_URL}?action=getReport&name=Juan+Dela+Cruz&week=Week+1&company=Acme+Corp
```

### Upload File
```
POST {SCRIPT_URL}
Body: {
  "action": "uploadFile",
  "fileName": "photo.jpg",
  "mimeType": "image/jpeg",
  "data": "<base64 encoded file>"
}
```

---

## 📤 EXPORT FEATURES

| Button | Output |
|--------|--------|
| Download PDF | html2pdf.js renders form → PDF |
| Print | Browser print dialog (optimized CSS) |
| Download DOCX | RTF format (opens in Word/LibreOffice) |

---

## 💡 FEATURES

- ✅ Exact replica of NBSC Form 9 layout
- ✅ Modal popups for all data entry
- ✅ Auto-save to localStorage (30-second intervals)
- ✅ Image gallery with lightbox preview
- ✅ Drag & drop file upload
- ✅ Form completion progress ring
- ✅ Live stats (objectives / activities / reflections / files)
- ✅ Google Sheets cloud save & load
- ✅ Google Drive image storage
- ✅ PDF, Print, DOCX export
- ✅ Responsive (mobile + desktop)
- ✅ Toast notifications
- ✅ Input validation with shake animation

---

## ⚙️ REQUIREMENTS

- Modern browser (Chrome, Firefox, Edge, Safari)
- Google account (for Sheets/Drive backend)
- Internet connection (for Google APIs and fonts)

No server or installation required — just open `index.html`!

---

## 🏫 NORTHERN BUKIDNON STATE COLLEGE
Manolo Fortich, 8703 Bukidnon  
*Creando futura, Transformationis vitae, Ductae a Deo*
