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

## Sprint folder selection
- Default: automatically use the MOST RECENT sprint folder in OSA - Reports
  (the folder with the latest end date in Sprint-YYYYMMDD_YYYYMMDD format)
- If user says "sprint hien tai" / "current sprint" / nothing → use most recent folder
- If user says "sprint truoc" / "last sprint" → use the second most recent folder
- If user explicitly names a folder (e.g. "Sprint-20260901_20260912") → use that
- DO NOT ask the user for sprint folder unless they want a specific one that is not current

## What this skill does
1. Confirm which sprint folder to use (default: most recent, no need to ask)
2. Ask for Status (Completed / Partial / Delayed)
3. Ask for Key Highlight (optional)
4. Ask for Timeline Issue (optional)
5. Ask for Prepared By (default: Gerald Nguyen, no need to ask unless user wants to change)
6. Ask for Submission Date (default: today, no need to ask unless user wants to change)
7. Ask which developers to include (multi-select)
8. Confirm all inputs then trigger the Power Automate flow

## Input fields
- sprintFolder: Sprint folder name (e.g. Sprint-20260901_20260912)
- status: Completed | Partial | Delayed
- keyHighlight: Free text, optional
- timelineIssue: Free text, optional
- preparedBy: Free text, default "Gerald Nguyen"
- submissionDate: Date in YYYY-MM-DD format, default today
- developers: Comma-separated list of developer names

## Developer list (for multi-select)
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
Sort by end date (second date) to find most recent.

## Output
Sprint_Report.xlsx saved to SharePoint under OSA - Reports/{sprintFolder}/

## Behavior rules
- Never ask for sprintFolder unless user wants a non-current sprint
- Never ask for preparedBy unless user wants to change from default
- Never ask for submissionDate unless user wants to change from today
- Always ask for status and developers
- Ask for keyHighlight and timelineIssue together in one message
- If user selects no developers or says "tat ca" / "all", include all developers
- Confirm all collected inputs before triggering the flow
- After flow completes, show link to Sprint_Report.xlsx
