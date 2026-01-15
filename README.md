# Moodle Rubric - A4 Export + Quick Grade (Userscript)

Generate a clean A4 landscape rubric report and apply rubric selections quickly from token input on the Moodle assignment grading page.

## Features
- Export a full rubric breakdown to A4 landscape with grid-fit layout
- Optional "Fit to 1 page" scaling before print
- Quick grade panel for token-based selection (A, A-, B+, NS, etc.)
- Includes per-criterion remarks, overall feedback, and current grade
- Strips timestamps and due date text near the student name

## Install (Chrome + Tampermonkey)
1. Install Tampermonkey from the Chrome Web Store:
   https://chromewebstore.google.com/detail/tampermonkey/dhdgffkkebhmkfjojejmpbldmpobfkfo
2. Open the userscript install link:
   https://raw.githubusercontent.com/raffitch/moodle-rubric-a4-export-userscript/main/moodle-rubric-a4-export.user.js
3. Tampermonkey will open an install page. Click "Install".
4. Visit your Moodle assignment grading page and confirm the script is enabled.

## Usage
1. Open an assignment grading page with a rubric:
   https://moodle.didi.ac.ae/mod/assign/*
2. The "Quick grade" panel appears in the bottom-left corner.
3. Enter one token per rubric criterion (comma or space separated) and click "Apply".
4. Click "Export A4" to open the report, then optionally check "Fit to 1 page" and print or save to PDF.

## Customize for another Moodle domain
If your Moodle domain is different, edit these lines in the script and re-install:
- `@match https://moodle.didi.ac.ae/mod/assign/*`
- `@downloadURL` and `@updateURL` (if you fork the repo)

## Files
- `moodle-rubric-a4-export.user.js` - main userscript
- `moodle-rubric-a4-export.meta.js` - metadata for auto-updates

## Troubleshooting
- No panel? Make sure the page is the grader view and refresh once.
- Tokens not applying? The token count must match the number of criteria.
- Print overflow? Use "Fit to 1 page" before printing.
