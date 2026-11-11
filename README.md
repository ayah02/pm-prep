# PM Prep — Practice. Track. Pass.

PMP-style practice and mock-exam system built for **CIS 386 Project Management** (Group G5, Al Yamamah University).

**Live site:** https://ayah02.github.io/pm-prep/  
**Instructor dashboard:** https://ayah02.github.io/pm-prep/admin.html

## Features
- Practice by domain / lecture / difficulty with instant feedback and explanations
- Timed mock exams: Course (60 Q / 75 min), Full-length (150 Q / 180 min), Quick quiz (20 Q / 25 min)
- Score report by domain, review of wrong answers, personal progress page
- English / Arabic interface (RTL)
- Instructor dashboard: average score, readiness rate, domain and lecture accuracy, question p-values, top distractors, CSV export

## How it works
| Part | Where it lives |
|---|---|
| Website (this repo: `index.html`, `admin.html`) | GitHub Pages |
| Backend (`apps-script/`) | Google Apps Script web app bound to a private Google Sheet |
| Question bank, answer keys, student responses | The private Google Sheet (never in this repo) |

The website sends requests to the Apps Script web app. Answer keys stay on the server and are only revealed after a question is answered (practice) or after the exam is submitted.

## Run your own copy
1. Create a Google Sheet, open **Extensions → Apps Script**, and add the files in `apps-script/`.
2. Add your question bank to the sheet (same columns as the `Questions` tab).
3. In **Project Settings → Script properties** add `ADMIN_PASSWORD`.
4. **Deploy → New deployment → Web app** (Execute as: Me, Who has access: Anyone), then copy the `/exec` URL.
5. Paste that URL into `API_URL` near the top of the `<script>` in `index.html` and `admin.html`.
6. Enable **GitHub Pages** (Settings → Pages → Deploy from branch → `main` / root).

---
Student project — CIS 386 Project Management, Group G5, Al Yamamah University.
