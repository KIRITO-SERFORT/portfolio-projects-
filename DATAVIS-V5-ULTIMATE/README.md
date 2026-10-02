# DATAVIS v5 ULTIMATE — Web App & Portal 📊

An end-to-end data analysis & student performance visualization web app built with **Streamlit**, **Plotly**, **Pandas**, and **SQLite Auth**.

Designed specifically for analyzing student marks, subject-wise scores, head-to-head student comparisons, multi-semester performance growth, data cleaning, and automated reports.

---

## ✨ Features & Architecture

### 🔑 1. User Authentication & Authorization
- **User Registration**: Register with Name, Email, and Password.
- **User Login**: Secure password authentication using PBKDF2 SHA-256 password hashing with salt.
- **Session Protection**: Protects visualization features behind authorized user sessions.

### 🎁 2. Freemium & Credit System (Turns Metering)
- **3 Free Analysis Turns**: Every new user automatically gets **3 FREE turns** upon registering.
- **$3 / Analysis Monetization**: Once free turns are used, users can purchase additional turns at **$3 per analysis turn** ($3 for 1 turn, $15 for 5 turns, $30 for 10 turns).
- **Paywall & Usage Tracking**: Automatically meters dataset uploads/analyses and prompts users to top up when balance reaches 0.

### 💳 3. Account & Billing Management Portal
- View account profile, email, and turns balance.
- **Instant Test Checkout**: Demo mode for testing credit top-ups instantly.
- **Stripe / Online Payment Integration**: Ready for connecting Stripe Payment Links or Webhooks.
- Full activity log for previous analyses and payment transaction history.

### 📊 4. Core Visualization Suite
- **🏠 Dashboard**: Overview KPIs, auto-generated data insights, quick bar/histogram previews.
- **📊 Explore & Customize**: 9 interactive chart types, radar (hex) charts, and 3D scatter explorer.
- **🧮 Stats & Correlation**: Summary stats, missing values detector, correlation heatmaps, distribution plots.
- **🎯 Personal Analysis**: In-depth breakdown of a student's score vs. average of higher performers and gap to topper.
- **⚔️ Head-to-Head Comparison**: Compare two students side-by-side with radar charts and leader scorecards.
- **📅 Multi-File Semester Growth**: Track student performance trends across semesters with scale-change detection and top-10 growth leaderboards.
- **📁 Raw Data**: Searchable, downloadable multi-table viewer.
- **🧹 Data Cleaning**: Remove duplicates, trim whitespace, drop empty columns, fill missing values, and IQR outlier detection.
- **📤 Combined Excel Report**: Export multi-sheet workbook containing raw tables, stats, and correlation matrices.
- **🤖 AI Insights**: Optional Claude API integration for narrative analysis.

---

## 🚀 How to Run the Website

### Option 1: Direct Command Line
```bash
pip install -r requirements.txt
streamlit run app.py
```
App will open automatically in your browser at `http://localhost:8501`.

### Option 2: Desktop Launcher (PyQt5 GUI)
```bash
python launcher.py
```

---

## 🛠️ Requirements & Setup on Your Side

All core features, SQLite database, authentication, free 3 turns, and test checkout work **out of the box**!

### Optional Steps for Production Deployment & Real Stripe Payments:
1. **Real Online Payments (Stripe)**:
   - Create a payment link in your [Stripe Dashboard](https://dashboard.stripe.com/) for $3 per turn.
   - Insert your Stripe Payment Link URL in `app.py` under the Account tab checkout section.
   - Use `auth_db.add_credits(user_id, turns, amount, "Stripe", payment_id)` in your webhook handler to automatically credit paid turns.

2. **Web Hosting Deployment**:
   - **Streamlit Community Cloud**: Connect your GitHub repository to [share.streamlit.io](https://share.streamlit.io/) for free web hosting.
   - **Render / Railway / Cloud Run**: Deploy as a standard Docker or Python web app with `streamlit run app.py --server.port=$PORT`.

---

## 👨‍💻 Credits
Coded by **Arun** (a.k.a. **KIRITO_SERFORT**).
