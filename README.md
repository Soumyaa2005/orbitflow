# OrbitFlow

A full project-management web app backed entirely by `OrbitFlow_PM_Tracker_v3.xlsx` — no separate database. Every screen reads the workbook live; every create/edit/delete writes straight back into it.

## Run it

```
cd orbitflow_app
pip install flask openpyxl
python server.py
```

Open **http://127.0.0.1:5000**. Keep Excel closed while using the app (Excel locks the file); reopen it afterward to see everything reflected, fully recalculated.

## What's included

- **Dashboard** — project/resource/task/bug/risk/CR counters, budget utilization, project health & risk-severity charts, workload distribution, upcoming deadlines, recent activity, pending approvals, quick actions.
- **Projects** — searchable/filterable/sortable/paginated table; create, edit, duplicate, archive, delete; a detail page with Overview/Tasks/Bugs/Risks/CRs/Financials tabs.
- **Resources** — full CRUD with role, department, billing rate, skills, availability, etc.
- **Work Items** — one unified table for Tasks/Bugs/Risks/CRs with type/status/priority/project filters, bulk status update, and type-specific fields (e.g. Probability/Impact for Risks).
- **Risks** and **Change Requests** — read-only views filtered from Work Items (Type=Risk / Type=CR) — nothing is ever written to these directly.
- **Financial** and **Billing** — read views of Financial Tracker / Billing Milestones.
- **Reports** — one-click Excel downloads (project, risk, financial, resource, task reports).
- **Settings** — workbook info, reload, theme toggle, download the full workbook.
- Global search (`Ctrl+K`), light/dark mode, toasts, confirm dialogs, loading skeletons, collapsible sidebar, responsive layout.

## How the Excel integrity is preserved

The workbook is modeled relationally:

- **Projects Master, Resources, Work Items** are the only sheets ever written to.
- **Risk Register / Change Requests** are pure array-formula views of Work Items — the app never writes to them; they update themselves the moment Work Items changes.
- **Project Dashboard, Financial Tracker, Resource Cost, CONFIG's project registry, Billing Milestones' summary** are formula-driven reporting sheets that pull from the source sheets automatically (pre-extended ~150 rows ahead) — also never written to directly.
- Column positions for every field are resolved at request time by reading each sheet's header row (not hardcoded indices), so new columns can be appended without code changes.
- **Delete behavior differs by sheet on purpose**: Financial Tracker and Project Dashboard address Projects Master *positionally* (safe to shift rows), but CONFIG's registry, Billing Milestones' summary, and Resource Cost's project cost-benefit block address it with a *hardcoded row number* baked into each formula — shifting rows would silently misalign those three. So deleting a Project or Resource blanks that row in place rather than removing it (position preserved, every dependent sheet stays aligned); "Archive" just flips Status. Work Items has no such dependency (everything reads it via whole-column formulas), so work items really are deleted.

This was verified end-to-end with a real recalculation engine (not just formula inspection): create/edit/duplicate/archive/delete/bulk-update were all exercised and the resulting workbook was recalculated to confirm every dependent sheet still computes correctly with no corruption.

## Notes / scope

- Documents and a Gantt/timeline view aren't implemented — the Details page covers Overview/Tasks/Bugs/Risks/CRs/Financials.
- Billing Milestones' milestone/team-pricing detail blocks (the actual billing schedule line items) are still filled in directly in Excel — those are genuinely bespoke per contract type and stay a deliberate manual-input exception.
- If you want the app to point at a different copy of the file, set an `XLSX_PATH` environment variable before running.
