# PPMP V2 — UI/UX Modernization & Mobile-First Redesign Plan

## 1. Overview & Objective
Transform the PPMP Pharmacy Evaluation Portal from a light-mode, emoji-centric interface into an **executive, clinical-grade, mobile-first dark application** inspired by modern high-contrast design systems (e.g., USTET Master Mock reference design).

---

## 2. Diagnosis & Comparison

| Design Dimension | Current UI (Issue) | Modern Target (Solution) |
| :--- | :--- | :--- |
| **Aesthetic / Theme** | Basic gray/white background (`#F1F5F9`) with pastel cards. Feels casual/cartoonish. | **Obsidian & Elevated Slate Dark Mode** (`#0B0F19` base, `#131B2A` card surfaces) with luminous accent borders. |
| **Iconography** | Native OS cartoon emojis (`🏥`, `👨‍⚕️`, `👩‍⚕️`, `💊`, `📝`, `📊`, `🔒`). Inconsistent across Android/iOS/Windows. | **100% Inline SVG Vector Icons** (Crisp stroke icons: Medical Cross, Stethoscope, Syringe, Capsule, Shield, Clipboard, Chart). |
| **Evaluator Role Card** | Wide cards overflowing mobile screens; cartoon doctor emojis. | **Single Centered Clinical Credential Badge** with glowing border, security pill (`VERIFIED CLINICIAN`), and SVG medical insignia. |
| **Form Stepper** | Numbered circles in a basic horizontal line. | **Illuminated Progress Pipeline** with glowing active step pill, completed check badges, and responsive touch layout. |
| **Mobile Bottom Dock** | Row of emoji buttons at the bottom. | **Floating Glassmorphic Dock** (`backdrop-filter: blur(16px)`), SVG stroke icons, and active indicator pills. |
| **Inputs & Pickers** | Plain white browser inputs. | **Dark Slate Glass Inputs** (`#0F172A`) with subtle border glows and smooth animated bottom-sheet drawers. |

---

## 3. Design System Tokens

```css
:root {
  /* Surface & Background */
  --bg-base: #0B0F19;              /* Deep Obsidian */
  --bg-surface: #131B2A;           /* Elevated Slate Card */
  --bg-surface-elevated: #1A2333;  /* Hover / Active Card */
  --bg-glass: rgba(15, 23, 42, 0.88); /* Blur Glassmorphism */

  /* Borders & Glows */
  --border-subtle: rgba(255, 255, 255, 0.08);
  --border-focus: rgba(13, 148, 136, 0.40);
  --border-gold: rgba(245, 158, 11, 0.35);

  /* Primary Clinical Accents */
  --accent-teal: #0D9488;          /* Modern Medical Cyan */
  --accent-teal-glow: rgba(13, 148, 136, 0.20);
  --accent-gold: #F59E0B;          /* Security & Warning Gold */
  --accent-gold-glow: rgba(245, 158, 11, 0.20);
  --accent-emerald: #10B981;       /* Approval Green */

  /* High-Contrast Typography */
  --text-primary: #F8FAFC;         /* Crisp White */
  --text-secondary: #94A3B8;       /* Cool Slate Muted */
  --text-tertiary: #64748B;        /* Dimmed Slate */

  /* Radii & Shadows */
  --radius-card: 14px;
  --radius-pill: 9999px;
  --glow-card: 0 4px 20px -2px rgba(0, 0, 0, 0.50);
  --glow-active: 0 0 20px rgba(13, 148, 136, 0.25);
}
```

---

## 4. Visual Layout Mockup

### Mobile View (Step 1 — Basic Info & Locked Role Badge)

```
+-------------------------------------------------------------------+
| [CROSS] PPMP Pharmacy Evaluation         [👤 Chito • END-USER]    |
|         Clinical Product Evaluation Portal                        |
+-------------------------------------------------------------------+
|  ( 1 Basic Info )  ---  2 Part I  ---  3 Part II  ---  4 Verdict   |
+-------------------------------------------------------------------+
|                                                                   |
|  +-------------------------------------------------------------+  |
|  | [SHIELD] Authorized Session: Chito Saba (Clinical End-user) |  |
|  +-------------------------------------------------------------+  |
|                                                                   |
|  EVALUATOR CREDENTIAL                                             |
|                                                                   |
|                +---------------------------------+                |
|                |  [STETHOSCOPE SVG]              |                |
|                |  End-user Evaluator             |                |
|                |  [🛡️ VERIFIED & LOCKED]         |                |
|                +---------------------------------+                |
|                                                                   |
|  GENERIC NAME / DRUG *                                            |
|  [ Select or search generic medicine...                    [v] ]  |
|                                                                   |
|  BRAND NAME                     PRICE PER UNIT (P) *              |
|  [ e.g., Rocephin           ]   [ 0.00                         ]  |
|                                                                   |
+-------------------------------------------------------------------+
|  [+] New Eval   [STETH] End-user   [CARE] Nurse   [PILL] Pharma   |
+-------------------------------------------------------------------+
```

---

## 5. Inline SVG Icon Library Specification
All cartoon emojis will be replaced with lightweight, crisp inline SVGs:
1. **Medical Cross Logo:** Minimal clinical cross inside rounded square badge.
2. **End-user Evaluator:** Clinical stethoscope vector icon.
3. **Nurse Evaluator:** Clinical syringe / care vector icon.
4. **Pharmacist Evaluator:** Medical pill / capsule vector icon.
5. **New Evaluation:** Document with plus badge vector icon.
6. **Consolidated Summary:** Analytics bar chart vector icon.
7. **Security / Lock:** Security shield with checkmark / lock vector.
8. **Navigation & Inputs:** Minimalist chevrons, search magnifying glass, and checkmarks.

---

## 6. Phased Implementation Strategy

### Phase 1: Design Tokens & SVG Vector Foundation
- Replace light theme CSS custom properties with the dark obsidian palette.
- Embed inline SVG helper icons for fast, zero-dependency rendering in Google Apps Script.

### Phase 2: Top Header, Stepper & Centered Role Badge
- Redesign the app header with dark glassmorphism and verified user chip.
- Implement the illuminated wizard progress track.
- Style the Step 1 centered assigned role card into a clinical credential badge.

### Phase 3: Form Fields & Searchable Bottom-Sheet Drawer
- Restyle text inputs, number fields, and select wrappers with dark slate glass.
- Upgrade the Generic Drug and Supplier picker into an ultra-responsive slide-up bottom sheet.

### Phase 4: Dynamic Questionnaire Cards & Rating Controls
- Restyle Part I, II, and III criteria cards inspired by the USTET subtest cards.
- Upgrade Yes/No/N/A rating options into modern segmented pill sliders.
- Modernize the Digital Signature pad canvas.

### Phase 5: Mobile Bottom Dock & Summary Dashboard
- Upgrade mobile navigation into a floating glassmorphic dock.
- Redesign the Consolidated Summary tab with dark data cards and glowing score progress meters.

### Phase 6: Mobile Ergonomics & Verification
- Test across mobile screen widths (360px, 390px, 412px, 768px, desktop).
- Ensure 48px touch targets and zero horizontal viewport overflows.
- Validate JavaScript syntax and deploy.

---

## 7. Questions & Alignment Before We Begin

### A. Accent Color Palette Preference
- **Option A (Image 2 & 3 Style):** Obsidian Dark (`#0B0F19`) + **Amber Gold** (`#F59E0B`) highlights and glowing badges.
- **Option B (Clinical Tech Fusion - Recommended):** Obsidian Dark (`#0B0F19`) + **Emerald/Teal** (`#0D9488`) primary accents with **Amber Gold** (`#F59E0B`) security/lock badges.

### B. Execution Strategy
- **Option 1 (Phased Review):** Execute and review Phase 1 & 2 first before proceeding to questionnaire cards and summary tables.
- **Option 2 (Full Pipeline - Selected & Completed):** Executed the transformation across the entire application in a single unified sweep.

---

## 8. Implementation & Verification Status (Completed)

- **Selected Theme:** Option B — Clinical Tech Fusion (Obsidian Dark `#0B0F19`, Emerald/Teal `#0D9488` primary accents, Amber Gold `#F59E0B` security/lock insignia).
- **Execution Mode:** Full Pipeline (Unified Sweep).
- **Implementation Highlights in [index.html](file:///c:/Users/chito/OneDrive/Desktop/PPMP%20V2/index.html):**
  - **100% Inline SVG Vector Icons:** Zero cartoon emojis across top branding, navigation tabs, credential badges, drawers, and mobile floating dock.
  - **Illuminated Stepper Pipeline:** Gradient track, glowing active rings, completed checkmark badges.
  - **Centered Clinical Credential Badge:** High-security Amber Gold glowing container with verified clinician pill.
  - **Dark Slate Glass Inputs & Bottom-Sheet Search:** Slide-up drawer on mobile, custom scrollbars, and glowing search field.
  - **Interactive Questionnaires:** Segmented Yes / No / N/A glowing pill toggles and live score metrics.
  - **Floating Glassmorphic Mobile Dock:** Floating rounded capsule with 5 icon tabs, 48px+ touch targets, and blur glassmorphism.
- **Verification Result:**
  - Full visual browser testing completed on mobile (`390px x 844px`) and desktop viewports.
  - Zero JavaScript errors, zero runtime warnings (`0 Errors / 0 Warnings`).

