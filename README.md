# EcoOils Project – JMS & JHA Builder

Single-file web app (`index.html`) for Work Method Statements (JMS) and Job Hazard Analyses (JHA, form QF/HS/01-02), with automatic revisions and lessons learned attached.

- `Code.gs` – Google Apps Script backend (Sheets index + Drive storage). Run `setup()` once, then deploy as a Web app (Execute as: Me, Access: Anyone). Every data call needs a login; PINs are set only through the email reset flow.
- App works offline on its own; paste the web app URL under Sync (⚙) to connect.
