# PPMP V2 — Complete System Walkthrough & Deployment Guide

This document explains the fully integrated **100% Free Pharmacy Product Evaluation System** built directly with **Google Sheets** and **Google Apps Script** (`Code.gs` + `index.html`), replacing AppSheet with zero licensing fees.

---

## 🌟 Key Architecture & Capabilities

### 1. Web App Frontend (`index.html`)
- **Native Web Application**: Runs directly inside any desktop or mobile browser without requiring any AppSheet app or subscription.
- **Role Lockdown & Security**:
  - Automatically matches logged-in user email against the `Evaluator_Accounts` sheet (Columns: `Account_ID`, `Evaluator_Role`, `Evaluator_Name`, `Email`).
  - Locks evaluator cards and form submissions strictly to the user's assigned role (`End-user`, `Nurse`, or `Pharmacist`).
  - Provides a built-in profile switcher modal for hospital desktop or tablet devices shared by multiple clinicians.
- **5-Step Mobile Evaluation Stepper**:
  - **Step 1: Product & Supplier Identification**: Autocompletes drugs from `Medicine_Master` and suppliers from `Supplier`.
  - **Step 2: Part I (19 Regulatory & Labeling Criteria)**: Segmented Yes / No / N/A buttons with live score pill.
  - **Step 3: Part II (6 Container & Physical Integrity Criteria)**: Immediate scoring and visual feedback.
  - **Step 4: Part III (Reconstitution & Dilution Criteria)**: Conditionally appears when "Requires Reconstitution" is toggled ON.
  - **Step 5: Recommendation & Digital Signature**: Choose Recommended / Not Recommended, write remarks, and sign using an HTML5 high-resolution canvas signature pad.

### 2. High-Concurrency Backend (`Code.gs`)
- **Simultaneous Submission Protection**: Uses Google Apps Script `LockService.getScriptLock()` with a 30-second queuing timeout to prevent row write collisions when multiple clinicians submit at the exact same second.
- **AppSheet `UNIQUEID()` Equivalent**: Every evaluation receives a true RFC 4122 Version 4 UUID (`Utilities.getUuid()`).
- **Cloud Digital Signature Storage**: Canvas base64 signatures are automatically converted into PNG images and saved to a dedicated Google Drive folder (`PPMP_Evaluation_Signatures`), generating public direct URLs for `=IMAGE(...)` formula rendering.
- **Zero-Click Automation**:
  - `submitEvaluationFromApp()` atomically saves to `Evaluations_Master`, computes Part I and II scores, mirrors the record to the respective evaluator sheet, updates `Consolidated_Summary`, and refreshes `Checklist_Report_Horizontal` **silently and automatically**.
  - Clinicians and administrators never need to click menu buttons after submitting evaluations.

### 3. Printable Checklist Viewer (`Checklist_Report`)
- **Cascading Dropdowns**:
  - **Row 3**: `🏢 Select Supplier:` (`C3:E3`, populated from `Supplier` tab `A2:A`) and `👤 Role:` (`G3`, dropdown: Pharmacist, Nurse, End-user, All).
  - **Row 4**: `🔍 Select Generic Drug:` (`C4:E4`, populated from `Medicine_Master` tab `A2:A`) and `📄 Ref #:` (`G4`).
- **Dynamic Live Lookup (Cell `H3`)**:
  - Matches Supplier (`C3`) + Generic Drug (`C4`) + Evaluator Role (`G3`), or jump directly via Ref # (`G4`).
  - **If Evaluated**: Populates Offered Price, Brand, Manufacturer, all 19 Part I checkmarks (`☑` / `☐`), all 6 Part II checkmarks (`☑` / `☐`), scores, recommendation, remarks, evaluator name/role, date, and **actual digital signature image** via `=IMAGE(...)`.
  - **If Not Evaluated**: Displays clean `"Pending Evaluation"` and empty boxes (`☐`) without formula errors or broken chips.
- **Cleaned Data Validation**: Residual validation rules on formula cells are wiped programmatically during initialization to eliminate phantom dropdown chip errors.

---

## 🚀 Step-by-Step Deployment Instructions

### Step 1: Update `Code.gs` in Apps Script
1. Open your Google Spreadsheet (`PPMP v2`).
2. Go to **Extensions** → **Apps Script**.
3. Replace the entire contents of `Code.gs` with the updated [`Code.gs`](file:///c:/Users/chito/OneDrive/Desktop/PPMP%20V2/Code.gs).
4. Save the project (<kbd>Ctrl</kbd> + <kbd>S</kbd>).

### Step 2: Deploy / Update Web App
1. In Apps Script, click the blue **Deploy** button at the top right.
2. Select **Manage deployments**.
3. Click the **Edit (pencil icon)** next to your active deployment.
4. Under **Version**, select **New version**.
5. Ensure:
   - **Execute as**: `Me (your account)`
   - **Who has access**: `Anyone` (or `Anyone with Google account` within your hospital workspace)
6. Click **Deploy**.
7. Copy the Web App URL (ending with `/exec`) to share with evaluators or open on tablets/phones.

### Step 3: Initialize Checklist Viewer in Google Sheets (One Time)
1. Switch back to your Google Spreadsheet tab.
2. Click the top menu **`🏥 PPMP Evaluation`** → **`📑 Initialize / Reset Printable Checklist Viewer`**.
3. Verify the layout:
   - Cell `C3`: Select any Supplier from your `Supplier` tab.
   - Cell `G3`: Select the Evaluator Role (`Pharmacist`, `Nurse`, `End-user`, or `All`).
   - Cell `C4`: Select any Generic Drug from `Medicine_Master`.
4. The checklist immediately reflects evaluations, scores, and signature!

---

## 🔍 Verification Checklist

| Feature | Expected Result | Status |
| :--- | :--- | :---: |
| **Supplier Dropdown** | Populates from `Supplier` tab (Col A) in cell `C3` | ✅ |
| **Drug Dropdown** | Populates from `Medicine_Master` tab (Col A) in cell `C4` | ✅ |
| **Role Selector** | Filters by Pharmacist, Nurse, End-user, or All in cell `G3` | ✅ |
| **Digital Signature** | Stored in Drive and rendered in row 43 via `=IMAGE(...)` | ✅ |
| **Evaluator Lockdown** | Web app restricts submissions to logged-in clinician role | ✅ |
| **Atomic Concurrency** | Protected by `LockService` with RFC4122 v4 UUIDs | ✅ |
| **Zero-Click Pipeline** | Submissions from Web App instantly sync to spreadsheet without manual menu clicks | ✅ |
