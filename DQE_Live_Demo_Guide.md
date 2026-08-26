# DQE — Live Demo Guide
*A 10–15 minute run-of-show for presenting the Data Quality Engine to DQE / Deployment-Migrations stakeholders*

---

## 1. Demo Objective

By the end of this demo, the audience should understand:
- What DQE is and the problem it replaces (manual Postman pulls + manual file diffing)
- How the core workflow works, end to end
- Where the biggest time savings are for their specific role (DQE root-cause work vs. Deployment/Migrations pre-go-live checks)
- How to actually get started using it themselves

Not the goal: covering every feature. Depth kills demos — pick the flagship workflow, show 2–3 "wow" features, leave time for questions.

---

## 2. Recommended 10–15 Minute Demo Flow

| # | Section | Time |
|---|---------|------|
| 1 | Opening | 1 min |
| 2 | Quick orientation | 1 min |
| 3 | Main workflow: Upload DVA & Compare | 5–7 min |
| 4 | "This is the good part" — 3 features | 2–3 min |
| 5 | Real-world example (woven through 3–4) | — |
| 6 | Q&A | remaining time |
| 7 | Close | 30 sec |

---

## 3. Presenter Talking Points (say these naturally, don't read them)

1. This replaces a manual Postman-and-spreadsheet process with one page.
2. Credentials are entered once, at the top, and every action below uses them — no re-typing.
3. You can pull live data straight from Opera/OHIP without ever opening Postman.
4. The core value is root-cause, not just "numbers don't match" — it tells you *why* and *what to do about it*.
5. Every result exports to Excel and can go straight to Monday.com — no separate write-up step.
6. This was built based on real feedback from the DQE team already testing it.
7. Nothing here is a black box — every comparison shows its raw data, not just a pass/fail.

---

## 4. Step-by-Step: Main Workflow (Upload DVA & Compare)

This is the flagship workflow — it's the reason the tool exists. Everything else supports it.

### Step 0 — Credentials (do this before the demo starts, or fast, on camera)
**WHAT TO DO:** Enter OHIP credentials in Step 1, or select a Saved Connection from the dropdown, then click Validate.
**WHAT TO SAY:** "Everything below shares these credentials — I only enter this once, and I can save it as a named connection so next time it's one click."
**WHY IT MATTERS:** Removes the single biggest friction point of doing this in Postman — re-entering tokens per environment, per hotel.

### Step 1 — Open "Upload DVA & Compare"
**WHAT TO DO:** Scroll to the Upload DVA & Compare card.
**WHAT TO SAY:** "This is where the actual root-cause analysis happens. Every file here is optional — I'll show you the full picture, but you can run this with just one or two files if that's all you have."
**WHY IT MATTERS:** Sets expectation that this is flexible, not an all-or-nothing form.

### Step 2 — Attach the DVA file
**WHAT TO DO:** Drop in the Duetto Data Validation export (.xlsx).
**WHAT TO SAY:** "This is Duetto's own validation export — it's what unlocks the full accuracy scoring and discrepancy engine."
**WHY IT MATTERS:** This is the anchor file; without it, you still get partial comparisons, but this is what turns on the full report.

### Step 3 — Attach Bookings Report + Folio Report
**WHAT TO DO:** Drop in the two flat exports.
**WHAT TO SAY:** "These are Duetto's own flat exports — bookings and folio. Feeding these in unlocks reservation-level root-cause instead of just a daily number."
**WHY IT MATTERS:** This is what lets the tool point at *which booking* or *which transaction code* caused a mismatch, not just that one existed.

### Step 4 — Click Run Analysis
**WHAT TO DO:** Click the Run Analysis button.
**WHAT TO SAY:** "And that's it — one click." *(let it run, keep talking through the wait)*
**WHY IT MATTERS:** This single click is replacing what used to be a manual side-by-side spreadsheet exercise.

### Step 5 — Walk the Summary tab
**WHAT TO DO:** Point at Rooms Commit Accuracy and Room Revenue Accuracy cards.
**WHAT TO SAY:** "At a glance — are we matching Opera on rooms and revenue for this date range."
**WHY IT MATTERS:** This is the number a stakeholder actually wants first.

### Step 6 — Walk the Discrepancies tab
**WHAT TO DO:** Click into Discrepancies, scroll to a failing date.
**WHAT TO SAY:** "For every date that's off, it shows Duetto's number, Opera's number, the difference, a root-cause code, and a plain-English explanation — this used to be the part that took hours."
**WHY IT MATTERS:** This is the actual "wow" moment — root cause, not just a red flag.

### Step 7 — Show the Excel download
**WHAT TO DO:** Click "Download Excel Report."
**WHAT TO SAY:** "Same report, formatted and ready to send — no rebuilding this in a spreadsheet afterward."
**WHY IT MATTERS:** Closes the loop — this is a deliverable, not just an on-screen dashboard.

---

## 5. Top Features to Highlight ("This Is the Good Part")

Pick 2–4 depending on time and audience. Recommended order:

### A. Live API pulls (no more Postman)
**DEMONSTRATE:** Pull Bookings/Blocks/Folio card — pick Folio, pick a date range, Fetch Data.
**SAY:** "This is a live call straight to Opera's API — no Postman collection, no manual token refresh."
**WHY IT MATTERS / PROBLEM SOLVED:** Removes the Postman dependency entirely for quick data checks.

### B. Config Snapshot — one-click pre-go-live check (aim this at Deployment/Migrations)
**DEMONSTRATE:** Pull Deployment Info card → Run Config Snapshot → point at the 6 result cards → click Download All.
**SAY:** "Six checks that used to be six separate Postman calls, one click, and one Excel file with all of it if you want to hand it off."
**WHY IT MATTERS:** This is Deployment/Migrations' single most relevant feature — a pre-go-live sanity check in seconds.

### C. API-vs-Duetto Folio/Bookings Comparison (if time allows)
**DEMONSTRATE:** Pull live folio, then show the "API Folio Comparison" tab after running analysis.
**SAY:** "This compares what the API says live, right now, against Duetto's own report — side by side, including the ones that match, not just the problems."
**WHY IT MATTERS:** This is brand new and directly answers "can I trust what Duetto's showing me" without waiting for an export.

### D. Recommendations tab
**DEMONSTRATE:** Click Recommendations tab.
**SAY:** "Every root-cause code maps to an actual next step — this isn't just diagnosis, it's what to do about it."
**WHY IT MATTERS:** Turns the tool from "reporting" into "actionable."

---

## 6. Real-World Example (use this as your throughline)

> "Let's say a hotel just went live last month, and Duetto's showing 94% room accuracy for last week — somebody needs to know why before it becomes a bigger problem. Normally that means pulling three or four reports by hand and eyeballing them side by side. Instead, I drop in the DVA, the bookings export, and the folio export, hit Run, and in a few seconds I've got the exact dates that are off, the root cause — say, a batch of reservations still sitting in RESERVED status that never got picked up — and a recommendation for what to ask the client to fix."

Reuse this story across steps 4–7 of the main workflow and feature A/B above ("this is the same hotel from our example — let's check its Config Snapshot before we ever got to this point").

---

## 7. Likely Questions & Answers

1. **"Does this replace Postman entirely?"** — For most day-to-day pulls and checks, yes. A few advanced/one-off calls may still need Postman, but the common workflows are covered here.
2. **"Where does this data go — is it stored anywhere?"** — It runs locally on my machine. History of past runs stays in a local database on my computer; nothing is sent to a shared server.
3. **"Can I use this for my own hotel/portfolio?"** — Yes — enter your own OHIP credentials in Step 1, or ask me to add you as a Saved Connection.
4. **"Is this replacing the DVA process, or working with it?"** — Working with it. The DVA export is still what unlocks full accuracy scoring; this tool consumes it, it doesn't replace it.
5. **"What if I don't have all the files — DVA, bookings, folio?"** — Every file is optional. You'll get whatever comparison the files you do have support — even one file gives you something.
6. **"Can Support use this too?"** — That's being discussed — right now it's DQE and Deployment/Migrations; open to expanding the audience.
7. **"How long did this take to build?"** — Built and iterated by me directly, based on real feedback from the team already testing it.
8. **"Can it tell me if a fix actually worked?"** — Not automatically today — you'd re-run the analysis after the fix to confirm. Flag this as a fair future ask, don't oversell it.
9. **"Does it work for every PMS, or just Opera/OHIP?"** — Currently built specifically around Opera Cloud / OHIP.
10. **"What happens to feedback I give during testing?"** — There's a Feedback button in the top nav that posts straight to the DQE board on Monday.com — it's read and acted on, like the last few updates in this app.

---

## 8. What NOT to Demo

- **The Pseudo Room filter toggle on API Bookings Comparison** — too in-the-weeds for a first-look audience; only bring it up if someone specifically asks why bookings counts look off.
- **Saved Connections CRUD (save/delete flow)** — useful, but a UI-mechanics detail; mention it exists, don't walk through creating one live unless asked.
- **Hotel Stats' async polling wait** — it can take up to a minute; don't let dead air kill your pacing. Either pre-run it before the demo and show the result, or skip it live and describe it instead.
- **Raw JSON/table dumps of unusual API responses** — if a "Show all fields" toggle surfaces raw dot-notation field names, don't linger — it looks technical and adds no value to this audience.
- **Anything requiring you to type a client secret on screen** — use a Saved Connection instead so no credentials are visible.

---

## 9. Live Demo Recovery Plan

| Problem | What to say |
|---|---|
| App doesn't load | "Looks like my local server needs a restart — one sec." *(restart, keep talking about the workflow while it comes up)* |
| Data doesn't appear / pull returns nothing | "That's actually a good example of why this matters — let's check the date range or credentials." *(don't panic — reframe as realistic behavior)* |
| API/integration call fails | "That's a live call to Opera's API, so occasionally it's a timing thing on their end — let me retry." |
| A button doesn't work | "Let's come back to that one — I want to keep us moving through the main flow." |
| You click the wrong thing | "Let's back up — that's not actually where I meant to go." *(navigate back, keep tone light)* |
| Someone asks a tangential question | "Great question — let me park that for a second and finish the workflow, then I'll come back to it." |
| You don't know the answer | "I don't want to guess on that one — let me confirm and follow up after." |

---

## 10. One-Page Presenter Cheat Sheet

### BEFORE THE DEMO
- [ ] Have OHIP credentials for a known demo hotel validated and saved as a Connection beforehand
- [ ] Have a real DVA export + Bookings Report + Folio Report staged and ready to drag in (don't hunt for files live)
- [ ] Pre-run one Config Snapshot and one Hotel Stats check earlier so you have a "before" result ready if live pulls are slow
- [ ] Confirm the local server is running and the page is loaded before people join
- [ ] Close any unrelated browser tabs / notifications

### DEMO FLOW
1. Opening (1 min) — what it is, what problem it solves
2. Orientation (1 min) — Step 1 credentials, three action cards below
3. Main workflow (5–7 min) — Upload DVA & Compare, end to end
4. Key features (2–3 min) — Live pulls, Config Snapshot, API Comparison
5. Real-world example — woven throughout, don't do it as a separate block
6. Q&A
7. Close — how to get access / who to talk to next

### KEY TALKING POINTS
- One credential entry unlocks everything below
- Live API pulls remove the Postman dependency
- Root cause, not just red flags — codes + explanations + recommendations
- Every file input is optional — partial data still gives partial value
- Config Snapshot = one-click pre-go-live check for Deployment/Migrations
- Everything exports to Excel and can post to Monday.com directly

### KEY CLICKS (main workflow path)
Step 1 Credentials → Validate → scroll to **Upload DVA & Compare** → attach DVA + Bookings + Folio → **Run Analysis** → **Summary tab** → **Discrepancies tab** → **Download Excel Report**

### DON'T FORGET
- Mention the Feedback button exists — it signals this is actively maintained, not a one-off
- If asked "what's next," you have a real answer: Deployment/Migrations rollout, then ongoing enhancement on request
- Don't apologize for things that are optional/not built yet — just say "not yet, that's a fair future ask" and move on

---

## Future Enhancement Notes (not in the app today — do not demo as if they exist)
- Automatic "did the fix work" re-check after a discrepancy is addressed
- Broader PMS support beyond Opera Cloud/OHIP
- Support team access/workflow (currently under discussion, not yet built)
