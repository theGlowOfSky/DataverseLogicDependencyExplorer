# Dataverse Logic Dependency Explorer

**Maintainer:** Gagan Deep Singh  
**LinkedIn:** [Gagan Deep Singh](https://www.linkedin.com/in/glowofsky/)  
**Host:** XrmToolBox  
**Data platform:** Microsoft Dataverse  
**Export format:** Excel workbook (`.xlsx`)  
**Version:** 1.2026.11.11  
**Project:** [GitHub repository](https://github.com/theGlowOfSky/DataverseLogicDependencyExplorer)

Before you rename a column, change a table, or touch a form, business rule, plugin step or flow, you need to know **what else uses it**.
Dataverse Logic Dependency Explorer answers that question from inside XrmToolBox: pick a table (or a column), tick the kinds of components to
check, and get a list of everything that references it - with the evidence, how strong that evidence is, and an Excel report you can share.

It is **read-only**. It never creates, changes or deletes anything in your environment.

> It tells you exactly what it was able to inspect. It does **not** promise to find every dependency, and a category it could not look at
> is never shown as "no dependencies".

## What you can do with it

- Search for a **table** (what is attached to it?) or a **column** (what reads, sets or reacts to it?).
- Check any mix of: **Business Rules**, **JavaScript web resources**, **Forms and form events**, **Plugin registrations**,
  **Classic workflows and processes**, **Power Automate cloud flows**.
- Watch the scan progress live, and see findings appear in the **Impact Report** tab **while the scan is still running**.
- Filter, sort and search the findings, open any one to read its evidence, and **export everything to Excel**.

## Quick start

1. Open the tool in XrmToolBox and connect to your environment.
2. Click **Load Tables**. Search the list (you can show custom or standard tables only) and select a table.
3. Choose what to analyze:
   - **Table** - the table and its associated components, or
   - **Column** - the selected column and the components that reference it (pick the column from the list below the tables).
4. Tick the component scopes you care about (**Select All** / **Clear All** are available) and click **Find Dependencies**.
5. Open the **Impact Report** tab to filter the findings, click a row to see its evidence, and click **Export Impact Report to Excel**.

You can click **Cancel Scan** at any time; findings found so far are kept, and unfinished categories are marked.

## Reading the results

Every finding carries a confidence label so you can tell a fact from a hint:

| Label | What it means | What to do |
|---|---|---|
| **Confirmed metadata reference** | Dataverse itself says so: a plugin step registered on the table, a column in a step's filtering attributes or image, a business rule or form that belongs to the table. | Treat as real. |
| **Strong definition reference** | The table/column was found inside a parsed definition: a form's fields and events, a workflow step, a business rule's conditions and actions, a flow action's parameters. | Very likely real; check the evidence. |
| **Text match - review required** | The name appears in source text (mostly JavaScript). A real runtime dependency is **not** proven. | Open it and read the line. |
| **Possible or dynamic dependency** | A pattern suggests a link that cannot be resolved without running the code: dynamic field names, an Update step with no filter, a flow whose table is an expression. | Investigate before changing anything. |

The detail panel of each finding shows: the component, its type and ID, parent table, the matched column, where in the definition it was found
(XML node, JSON path, attribute, line number), the evidence snippet, why it was included, and the limits of that finding.

### Empty results - four different situations
- **Nothing run yet** - start a scan.
- **Filtered out** - reset the filters.
- **Clean result** - every selected category was inspected and none references the target.
- **"No matches, but not a clean result"** - at least one selected category failed, was denied, or could only be partly inspected. Check the status table.

## What each scope looks at

| Scope | What is inspected | Coverage |
|---|---|---|
| Business Rules | Rules of the selected table: conditions and actions in the rule definition. | Full |
| JavaScript Web Resources | Script text, with comments, strings and code told apart; matches are always "review required". | Text analysis |
| Forms and Form Events | Forms of the selected table: fields, registered libraries, OnLoad / OnSave / OnChange handlers, and whether a handler's code mentions the column. | Full |
| Plugin Registrations | Steps, stage, mode, filtering attributes, pre/post images, plugin class and assembly. **The compiled plugin code is not inspected.** | Registration only |
| Classic Workflows and Processes | Workflows, dialogs, actions and business process flows (their XAML definitions). | Full for readable definitions |
| Cloud Flows | Flows stored in Dataverse (solution-aware flows), read with your existing connection. | **Partial** |

Categories you leave unticked appear as **Not inspected**.

## Faster JavaScript scans

Searching **all script web resources** reads every JavaScript library in the environment, which can take a while in large environments.
The tool warns you before it starts, and then:

- downloads libraries **in parallel** and analyzes them **while later ones are still downloading**;
- shows live progress (downloaded / analyzed / matches so far) and streams findings into the Impact Report as they are found;
- skips files that cannot contain the name without analyzing them.

To make it faster still:
- tick **Skip managed web resources** (code from installed solutions - usually not yours to change), or
- choose **Only libraries registered on the selected table's forms** for a targeted scan of just the scripts that run on that table.

## The Excel report

**Export Impact Report to Excel** always exports the complete analysis (not only the filtered rows) as a real `.xlsx` workbook:

| Sheet | Contents |
|---|---|
| **Summary** | Environment, scan time, search mode, target, totals, counts by component type and by confidence, scan status per category, important limitations and warnings. |
| **Dependencies** | One row per finding with component, IDs, matched value, confidence, evidence, location (node / path / line), solution, method and limitation. |
| **Scan Diagnostics** | Per category: status, records examined, matches, errors, warnings, permission limits and what that scan cannot see. |
| **Scan Configuration** | The table, column, scopes and options used, plus the version and the definition of each confidence label. |

The workbook has frozen header rows, filters and readable widths. It never contains passwords, tokens or secrets.

## Permissions

The tool uses your XrmToolBox connection and can only see what **your user** can read. If a category needs data you cannot read, that category shows
**Permission denied** and the rest of the scan continues. Typical read access needed: business rules and processes, forms, web resources, plugin registrations and solutions.

## Good to know (limitations)

- **Plugins:** a plugin can use a column in its code without any trace in its registration, so registration-based results can never be complete.
- **JavaScript:** dynamic names (`getAttribute(name)`, `"new_" + x`), aliases and minified code cannot be resolved. Generic column names without a prefix (such as `name`) are matched only in API calls, query text and exact strings, and cannot be tied to one table.
- **Cloud flows:** only flows stored in Dataverse are visible. The Power Automate Management API is not used and no extra sign-in is requested. Environment variables, child flows and custom connectors are not resolved.
- **Forms:** only forms of the selected table; forms of other tables that embed it (subgrids, quick views) are not searched.
- **Classic processes:** by default only processes of the selected table are read. Tick **Cross-table process scan** to also find processes of other tables (this reads every process definition and is slower).
- **Business rules:** line numbers refer to the rule's generated client code.
- Solution names exclude the Default/Active solutions; Excel cells are limited to 32,767 characters (longer text is cut and marked `[TRUNCATED]`).

## Troubleshooting

| Symptom | What to try |
|---|---|
| The table list is empty | Click **Load Tables** after connecting; check your user can read table metadata. |
| A category says **Permission denied** | Ask an administrator to grant read access to the data named in the message, or untick that category. |
| The JavaScript scan is slow | Tick **Skip managed web resources**, or use the targeted scope; you can also **Cancel Scan** and keep partial results. |
| "No matches" but you expected some | Open the status table: a partial or failed category means the search was incomplete. |
| Excel export fails | Close the file if it is open in Excel, or save to another folder. |

---
**Dataverse Logic Dependency Explorer** v1.2026.11.11 - maintained by Gagan Deep Singh ([LinkedIn](https://www.linkedin.com/in/glowofsky/)) - [GitHub](https://github.com/theGlowOfSky/DataverseLogicDependencyExplorer) - for XrmToolBox and Microsoft Dataverse.
