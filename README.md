<p align="center">
  <img src="icons/icon-192.png" alt="Smart Attendance Logo" width="120">
</p>

<h1 align="center">Smart Attendance — Digital Classroom Register</h1>

<p align="center">
  <strong>A comprehensive, offline-first Progressive Web App for teachers and educational institutions to take digital classroom attendance with Google Sheets sync, offline support, and Excel downloads.</strong>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/PWA-Offline%20Ready-34d399?style=for-the-badge&logo=pwa&logoColor=white" alt="PWA">
  <img src="https://img.shields.io/badge/JavaScript-ES6+-fbbf24?style=for-the-badge&logo=javascript&logoColor=black" alt="JavaScript">
  <img src="https://img.shields.io/badge/Google%20Sheets-Synced-38bdf8?style=for-the-badge&logo=google-sheets&logoColor=white" alt="Google Sheets">
  <img src="https://img.shields.io/badge/SheetJS-0.20-a78bfa?style=for-the-badge" alt="SheetJS">
  <img src="https://img.shields.io/badge/License-Proprietary-f87171?style=for-the-badge" alt="License">
</p>

---

## ✨ Features

### ⏱️ Session & Subject Selection
- Fast setup: Choose the session date and subject from dynamically loaded dropdowns.
- Automated retrieval of teachers, courses, academic years, semesters, and subjects.
- Quick indicators highlighting active session info.

### 📋 Interactive Roll Call Register
- Fast entry: Tap rolls to mark students as **Present** or **Absent** instantly.
- Live dynamic counter tracks total Present, Absent, and Total students in real-time.
- Visual cues (Neon colors: green for Present, red for Absent) optimized for mobile-first views.
- Supports class-wise and batch-wise filtering with easy toggle buttons.

### 🧠 VibeMantra Schema Learner & Translator
- Custom header detection learns your spreadsheet structure dynamically (column order does not matter!).
- Normalizes variations like `roll no`, `student id`, `student name`, `batch`, `section`, etc.
- Translates array-of-array CSV entries into clean normalized JavaScript objects on the fly.

### ☁️ Offline-First Queue System
- Full offline support: Take attendance even when the internet drops or in remote classrooms.
- Auto queue: Saves records locally using `localStorage` under a pending sync status.
- **Automatic Sync:** Syncs local records with Google Sheets instantly when an active connection is restored.
- In-app feedback: Uses toasts to notify users of successful offline captures and pending records sync.

### 📊 Smart Reports & Analytics
- Visual overview of attendance stats per subject and batch.
- Detailed present/absent ratios and full roll sheets.
- **One-click Excel Export:** Generate and download attendance registers as clean sheets powered by SheetJS (XLSX).

### 🔐 Secure License Verification
- HWID-bound license key authentication.
- Expiry date validation and institution name binding.
- Full-screen lock out screen for unauthorized instances.
- Base-64 encoded license validation keys with local keystore backup.

### 🌙 Cyberpunk Dark Theme
- Sleek dark aesthetic designed to reduce eye strain for teachers in low-light classrooms.
- Responsive, mobile-first design built with modern CSS custom variables and glassmorphic cards.
- Powered by clean, custom layout tokens.

---

## 🏗️ Tech Stack

| Technology | Purpose |
|---|---|
| **HTML5 + CSS3** | Structure & styling using modern CSS variables |
| **Vanilla JavaScript (ES6+)** | Dynamic application core — zero heavy frameworks |
| **PWA / Service Worker** | Offline caching, background asset delivery, install capability |
| **SheetJS (XLSX)** | Fast client-side Excel exporting |
| **Google Apps Script** | Cloud synchronization, API handler, database communication |
| **Google Sheets** | Central database storage |
| **Phosphor Icons** | Clean UI iconography |
| **Inter** | Typography (Google Fonts) |

---

## 📁 Project Structure

```
smart-attendance/
├── index.html                  # Landing page
├── app.html                    # Main PWA application
├── manifest.json               # PWA manifest
├── sw.js                       # Service worker
├── version.json                # App version tracking
├── sitemap.xml                 # SEO Sitemap
├── robots.txt                  # Search engine rules
├── css/
│   └── app.css                 # Application styles
├── js/
│   ├── app.js                  # Core app logic
│   └── api.js                  # API communication & offline queue
├── appstart/
│   ├── appstart.js             # Boot sequence loader
│   ├── appstart.css            # Boot screen styles
│   ├── config.js               # App configuration (theme, API endpoint)
│   ├── license.js              # License key logic
│   ├── keystore.js             # Local storage key validation
│   ├── schema.js               # Dynamic Schema learner
│   └── translator.js           # Fuzzy headers matcher
├── icons/
│   ├── icon-192.png            # PWA icon (192x192)
│   └── icon-512.png            # PWA icon (512x512)
└── google_apps_script/
    └── Central_API.gs          # Google Apps Script backend
```

---

## 🚀 Getting Started

### 1. Deploy Google Apps Script
1. Open [Google Apps Script](https://script.google.com).
2. Create a new project and paste the contents of `google_apps_script/Central_API.gs`.
3. Deploy as a **Web App** (Execute as: "Me", Who has access: "Anyone").
4. Copy the deployment Web App URL.

### 2. Configure the App
1. Open `appstart/config.js` in your text editor.
2. Update the `CENTRAL_API_URL` variable with your deployed Web App URL:
   ```javascript
   CENTRAL_API_URL: "https://script.google.com/macros/s/YOUR_API_DEPLOYMENT_ID/exec"
   ```

### 3. Host the PWA
Host the root directory on any static web host (e.g. Cloudflare Pages, GitHub Pages, Netlify, Vercel):
- The app operates fully in the browser (client-side).
- No server-side database configuration required — Google Sheets acts as your database.
- Caching logic inside `sw.js` takes care of offline capabilities.

### 4. Activate License
1. Launch the app in your browser.
2. Enter your activation license key (contact developer for generating verification keys).
3. The license gets verified and cached in your browser's secure keystore.

---

## 📱 Install as PWA

Smart Attendance can be installed as a native app on any mobile or desktop device:

1. Open the hosted landing page in **Google Chrome** (Android/Desktop) or **Safari** (iOS).
2. For Chrome, tap the install prompt in the address bar, or select **"Install App"** from the browser options menu.
3. For Safari, tap the **Share** button and select **"Add to Home Screen"**.
4. Launch the app directly from your home screen. It will operate full-screen with no browser address bar.

---

## 👨‍💻 Author

**VibeMantra Studio**
Email: pranavparekhcontent@gmail.com

---

## 📄 License

This project is proprietary software. Unauthorized distribution, copying, or modification is prohibited.
