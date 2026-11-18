---
name: webyith-sprint-report
description: >
  Generates the Webyith Sprint Report for AMB/OSA from DevOps daily export files.
  Use this skill when the user asks to create, generate, or update a sprint report.
  Trigger phrases: "tao sprint report", "lam bao cao sprint", "generate the sprint report",
  "sprint hien tai", "bao cao sprint", or when user names a Sprint folder.
---

# Webyith Sprint Report Skill

## Purpose
Generate the AMB/OSA Sprint Report from Webyith DevOps daily HTML export files
uploaded to SharePoint under OSA - Reports/{Sprint-YYYYMMDD_YYYYMMDD}/.

## What this skill does
1. Ask the user which sprint to report on
2. Ask for Status (Completed / Partial / Delayed)
3. Ask for Key Highlight (optional)
4. Ask for Timeline Issue (optional)
5. Ask for Prepared By (default: Gerald Nguyen)
6. Ask for Submission Date (default: today)
7. Ask which developers to include (multi-select)
8. Trigger the Power Automate flow to generate Sprint_Report.xlsx

## Input fields
- sprintFolder: Sprint folder name (e.g. Sprint-20260901_20260912)
- status: Completed | Partial | Delayed
- keyHighlight: Free text, optional
- timelineIssue: Free text, optional
- preparedBy: Free text, default "Gerald Nguyen"
- submissionDate: Date in YYYY-MM-DD format
- developers: Comma-separated list of developer names

## Developer list
- Gerald Nguyen Nhat - ONESECURE
- Zidane Ly - ONESECURE
- Cuong Thinh - ONESECURE
- Huu Thai - ONESECURE
- Alec Van Hai Son - ONESECURE
- Wen Xin - ONESECURE
- Jeffery Tay - ONESECURE
- David Gan - ONESECURE

## Sprint folder naming convention
Sprint-YYYYMMDD_YYYYMMDD (e.g. Sprint-20260901_20260912)
"sprint hien tai" / "current sprint" = most recent sprint folder
"sprint truoc" / "last sprint" = the one before current

## Output
Sprint_Report.xlsx saved to SharePoint under OSA - Reports/{sprintFolder}/

## Behavior rules
- Always ask for all inputs before triggering the flow
- Never generate the report without confirming inputs first
- If user says "mac dinh" or "skip" for optional fields, pass empty string
- If user selects no developers or says "tat ca", include all developers
