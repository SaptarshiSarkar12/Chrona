# Chrona

Chrona is a lightweight Google Apps Script application that automates the tedious process of decoding IIIT Dharwad’s slot-based CSE branch's timetable and syncing it with Google Calendar.
Students configure their semester dates and course selections once, and it ensures their calendar stays updated even when the timetable changes.

## Features
- **Google Account Integration**: Authenticates with the student's Google account which has read access to timetable in Google Sheets.
- **Semester Config**: Enter semester details like sem start, mid-sem, and end-sem dates once.
- **Course Selection**: Compulsory courses auto-selected; electives can be checked manually.
- **Flexible Durations**: Choose *Full Sem*, *Pre‑Mid*, or *Post‑Mid* for each course.
- **Daily Sync**: Automatically schedules a sync job that runs every morning at 7 AM to keep events updated.
- **Instant Calendar Events**: Classes appear in Google Calendar with correct time slots and classroom numbers.

## Demo
Watch the demo video to see how instantly events are created after saving config.

https://github.com/user-attachments/assets/f0ae0baf-6f07-4ff1-a8a3-384f205d75f8

## Setup Guide
1. Open [Google Apps Script](https://script.google.com/).
2. Create a new project and paste the contents of [`code.gs`](./code.gs) and [`sidebar.html`](./sidebar.html) after modifying the decoding logic and required inputs in the respective files.
3. Replace the hardcoded **Sheet ID** with your institution’s timetable sheet.
4. Enable the **Google Calendar API** in Apps Script services. Or, you can use the [`appsscript.json`](./appsscript.json) file for the app configuration.
5. Deploy the project as a web app.
6. Log in with your Google account, configure semester dates and courses, then click **Save Config & Sync**.

## Notes
- Currently tailored for **IIIT Dharwad CSE Timetable** (slot decoding logic is specific).
- Other colleges can adapt the parsing logic for their own timetable format.
- Settings are stored under [`PropertiesService`](https://developers.google.com/apps-script/reference/properties/) per user account, so configuration is one‑time.

## Inspiration
Chrona was built to solve a personal pain point: manually decoding slots and mapping them to class times was error-prone and time-consuming.
By open-sourcing this, the hope is that other institutions can adapt the idea and save students countless hours.
