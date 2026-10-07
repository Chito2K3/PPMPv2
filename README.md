# 🏥 PPMP Mobile Pharmacy Product Evaluation System

An enterprise-grade, AppSheet-powered mobile pharmacy product evaluation system backed by Google Apps Script and Google Sheets. Designed for hospital pharmacy departments to streamline product evaluation across **3 evaluators** with automated score calculations, tab mirroring, and real-time consensus dashboards.

---

## 📁 Repository Structure

| File | Description |
|---|---|
| [`Code.gs`](file:///c:/Users/chito/OneDrive/Desktop/PPMP%20V2/Code.gs) | Google Apps Script backend engine containing initialization, synchronization, score calculation, sheet mirroring, and consensus aggregation logic |
| [`index.html`](file:///c:/Users/chito/OneDrive/Desktop/PPMP%20V2/index.html) | Native 100% Free Web Application frontend with 5-step stepper, digital signature pad, role-lockdown, and live summary dashboard |
| [`AppSheet_Setup_Guide.md`](file:///c:/Users/chito/OneDrive/Desktop/PPMP%20V2/AppSheet_Setup_Guide.md) | Step-by-step AppSheet configuration blueprint (tables, slices, 5-page form UX, format rules, virtual columns) |
| [`walkthrough.md`](file:///c:/Users/chito/OneDrive/Desktop/PPMP%20V2/walkthrough.md) | Deployment guide and system walkthrough |
| [`README.md`](file:///c:/Users/chito/OneDrive/Desktop/PPMP%20V2/README.md) | Project deployment and setup guide |

---

## 🚀 Complete Step-by-Step Setup (From a Blank Google Sheet)

### Phase 1: Creating & Setting Up the Google Sheet

#### Step 1: Create a New Google Sheet
1. Open your browser and go to [sheets.google.com](https://sheets.google.com).
2. Click **+ Blank Spreadsheet**.
3. Name your spreadsheet at the top left (e.g., `PPMP Pharmacy Evaluation System`).

#### Step 2: Create the Source Drug List Tab (`indicative 1`)
1. At the bottom of your sheet, double-click the tab named `Sheet1` and rename it to: `indicative 1` (exact spelling).
2. Fill in Column A of `indicative 1` starting at Row 4:
   - **Cell A1**: `PPMP MASTER MEDICINE LIST` (Optional Title)
   - **Cell A2**: *(Leave Blank)*
   - **Cell A3**: `GENERIC NAME` (Column Header)
   - **Cell A4 onwards**: Type or paste your generic drug list (one item per row).

*Example Layout of `indicative 1`:*

| Row | Column A |
|---|---|
| 1 | PPMP MASTER MEDICINE LIST |
| 2 | |
| 3 | **GENERIC NAME** |
| 4 | Paracetamol 500mg Tablet |
| 5 | Amoxicillin 500mg Capsule |
| 6 | Cefuroxime 500mg Tablet |
| 7 | Sodium Chloride 0.9% 1L Bottle |

---

### Phase 2: Adding the Apps Script Engine (`Code.gs`)

1. In your Google Sheet top menu, click **Extensions** → **Apps Script**.
2. The Apps Script code editor will open in a new tab displaying a default `Code.gs` file.
3. Select all existing text in `Code.gs` and delete it.
4. Copy **everything** from [`Code.gs`](file:///c:/Users/chito/OneDrive/Desktop/PPMP%20V2/Code.gs) and paste it into the editor window.
5. Click **Save** (💾 icon) or press `Ctrl + S`.

---

### Phase 3: One-Click Automated Sheet Setup

1. Go back to your Google Sheet tab in your browser and **Refresh the page** (`F5` or `Ctrl + R`).
2. A new menu item titled **🏥 PPMP Evaluation** will appear in the top toolbar.
3. **Run Initialization**:
   - Click **🏥 PPMP Evaluation** → **🚀 Initialize Evaluation Workbook**.
   - *Authorization*: When prompted, click **Continue** → Select your Google Account → Click **Advanced** → Click **Go to PPMP Evaluation Script (unsafe)** → Click **Allow**.
   - The script will automatically create and format the operational sheets: `Medicine_Master`, `Suppliers`, `Evaluator_Accounts`, `Evaluations_Master`, `End-user`, `Nurse`, `Pharmacist`, and `Consolidated_Summary`.
4. **Sync Drug List**:
   - Click **🏥 PPMP Evaluation** → **🔄 Sync Master Drug List**.
   - This imports all generic drug names from `indicative 1` (Row 4 down) into `Medicine_Master`.
5. **Install Triggers**:
   - Click **🏥 PPMP Evaluation** → **⚡ Setup Auto-Sync Triggers**.
   - This sets up automated edit sync and score calculations.

---

### Phase 4: AppSheet Mobile App Connection

1. In your Google Sheet top menu, click **Extensions** → **AppSheet** → **Create an app**.
2. Follow the comprehensive setup blueprint in [`AppSheet_Setup_Guide.md`](file:///c:/Users/chito/OneDrive/Desktop/PPMP/AppSheet_Setup_Guide.md) to configure your 5-page mobile evaluation form, views, and format rules.

---

## 📊 Sheet Schemas & Features

- **Automated Score Calculations**:
  - **Part I Score**: $\frac{\text{Yes Count}}{19 - \text{N/A Count}} \times 100\%$
  - **Part II Score**: $\frac{\text{Yes Count}}{6 - \text{N/A Count}} \times 100\%$
- **Consolidated Consensus Dashboard**:
  - 🟩 **Unanimous Recommended**: All 3 evaluators voted Recommended
  - 🟥 **Unanimous Not Recommended**: All 3 evaluators voted Not Recommended
  - 🟨 **Split Decision**: Mixed recommendations
  - ⬜ **Pending**: Evaluation incomplete (< 3 evaluators submitted)
