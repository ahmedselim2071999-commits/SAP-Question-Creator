# SAP LMS Bulk Question Automator — GitHub Pages

## What this version is
A static web app made from HTML, CSS and JavaScript. It does not need Flask, Python, Node, Render, or a backend server.

## Files
- `index.html`
- `styles.css`
- `app.js`

## Deploy to GitHub Pages
1. Create a GitHub repository, e.g. `sap-lms-bulk-question-automator`.
2. Upload the three files to the repository root.
3. Open **Settings → Pages**.
4. Under **Build and deployment**, select **Deploy from a branch**.
5. Select branch `main` and folder `/ (root)`.
6. Click **Save**.
7. GitHub will publish a Pages URL.

## How it works
The user uploads:
- Word DOCX, AIKEN or TXT questions
- SAP LMS XLSX template

Then selects:
- Course code
- Pre Assessment / Knowledge Check / Post Assessment
- Locale
- Domain
- Default question type

The app parses and validates questions, previews them, then generates an XLSX file in the browser.

## ID convention
Pre Assessment = `PA`
Knowledge Check = `KC`
Post Assessment = `POST`

Example:
`26Q3_vid_c_SOWA2_JB_KC_001`

## AIKEN
Question
A. choice
B. choice
C. choice
D. choice
ANSWER: B

Multiple:
`ANSWER: A,C`

Optional per-question type:
`TYPE: MULTI_CHOICE_MULTIPLE_ANSWER`

## Browser libraries
ExcelJS and Mammoth.js are loaded from jsDelivr. The app source is static, but an internet connection is required to load those libraries unless you vendor them into the repository.
