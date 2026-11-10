# Approved-File Code and No. Reconciliation

## Primary Workflow (Current)

The main workspace now prioritizes approved-file reconciliation, not ERP code generation or discussion discovery. Load the approved Excel sheet, choose full target code and target No. columns, confirm the project and file, then identify each row by an existing 1Office internal ID. The reader still requires the existing `No.`, `Ma goi thau/Ma CV`, and `Noi dung cong viec` headers with their Vietnamese accents. Select additional ID/old-code columns when available, or paste `Excel row / existing ID` pairs. For newly added tasks, create them manually or import them in 1Office first, recollect the project, then map their IDs. No task creation is performed by this tool.

- ID is stable identity; target code and target No. may both differ from their current values. A unique old code can establish an initial match, or the target code can match an unchanged existing code. Name and No. never establish identity. Unmatched rows remain unresolved, not automatically new. Manual IDs conflicting with file IDs, old-code mismatches, duplicate row IDs, duplicate targets and out-of-project IDs block planning.
- Approved full codes are used literally, not regenerated from row position or ERP suggestions. Blank target code/No. is a review error, not a deletion or implicit keep. The No. field must be explicitly confirmed as the detail custom No., not Gantt/tree outline numbering. The latter has no supported update path here.
- The table shows ID, matching method, current/approved code and current/approved No. on the same row. Unknown current fields remain unknown. Tasks outside the file are untouched and cannot silently relinquish a target value. Missing outside-file field coverage prevents full-project collision validation.
- For test project 203 at port 9443 only, an explicit read button retrieves missing code/No. evidence in batches of up to 20 tasks, mapped IDs first. It uses the existing two-widget reader, does not read discussions, does not change the API inventory, and disposes late results on unmount. Each batch requires a click; errors remain visible/unknown, with no write/navigation fallback. Detail reads supplement the snapshot, not an atomic live inventory refresh.
- Only fully resolved, conflict-free snapshots yield an operation preview. Free targets are allocated first; cycles use a reserved temporary value for the affected field. Code and No. are planned independently by ID. A 1,000-task cycle test checks the planner, not server throughput. The temporary values and operation order are NOT verified against the server's code-edit contract or business constraints.
- The local JSON download includes project, sheet, column/ID choices, timestamps, detail evidence and the operation preview, with `writeEnabled: false`. Changing choices invalidates the preview. Clearing/changing file, sheet, project context or temporary data discards local mappings/evidence. No write calls, approval state or production write permissions are added. Never feed these operations to the old No.-only writer.

The existing code-generation/ERP/date comparison tools are collapsed as auxiliary tools. The No.-only writer remains explicitly labeled as the old separate test flow, with its original allowlists and confirmations unchanged. Test baseline/discussion preservation tools are collapsed below the main workspace and are no longer a prerequisite for preparing a code/No. comparison.

Still required before enabling saves: verify the actual code-edit request and custom-No. field binding, fresh project-wide identity/value checks, persistent journal/recovery and concurrency handling for both fields, scoped test execution/readback, and proportionate preservation checks. The new flow does not claim live updates, reparenting, report/dependency synchronization or preserved protected data. Automated tests validate preview semantics and UI safety; a live approved-file reconciliation run is not yet confirmed.

## Earlier Auxiliary File Review

The standalone `import-review.html` workspace opens directly from the persistent
"Chuẩn hóa mã dự án" launcher in the popup footer, or the floating launcher on the
No. review page. The popup project-inventory card no longer has a separate No.
entry. The workspace groups two tools: file/codes/ERP and bulk No. comparison.
The file tab has its own three-step header; the No. tab lazily mounts the existing
No. review component with the already validated snapshot, not another storage read.
Both tabs stay mounted after opening so switching does not discard file inputs or
preview state. Switching and workspace clear are disabled during No. operations.
All existing test-only write allowlists, confirmations, journal recovery and
before-unload protections remain unchanged. File previews do not invoke the writer.
The original embedded panel remains available for compatibility. No new permissions
or injected controls on the 1Office site are added.

Opening without an inventory permits local file review and deliberately does not
read a previously stored project snapshot. When opened with a collected inventory,
the URL binds the existing shared session snapshot to its project ID and capture
timestamp. A replaced/cleared snapshot is rejected rather than attached silently.
Files and choices are not transferred between open pages. The update/first-time
mode defaults and coding rules are unchanged. Temporary-data clearing remounts the
local file workflow and clears optional in-memory inventory context.

It does not create tasks, generate task codes, prepare write commands, or invoke
1Office APIs. It uses the already collected project inventory if available.

## Scope

- Select an `.xlsx` file up to 5 MB and choose its sheet.
- Parse in a dedicated worker, terminated after 30 seconds or on clear/unmount.
- Limit parsed workbooks to 20 sheets, 10,000 data rows and 100 columns per sheet.
- Required first-row headers: `No.`, `Mã gói thầu/Mã CV`, `Nội dung công việc`.
- Report levels and parent candidates from the file outline only. These are not
  confirmed 1Office relationships. Accept legacy letter suffixes for inspection,
  but flag them for numbering review; do not strip suffixes or renumber yet.
- Flag blank/numeric-local/duplicate codes, duplicate No., missing or later
  parent candidates, and duplicate matching codes in the system inventory.
- Only propose exact, case-sensitive unique code matches against a complete
  inventory. An unmatched code is not proof of a new business task.
- Preserve source cell values in memory; do not evaluate spreadsheet formulas.
  The Excel reader returns cached formula results and can return null for formula
  errors/missing caches. Actual dates and dependencies require separate validation.
- Paginate results in groups of 50 rows. Render source strings as text, not HTML.

## Data and boundaries

The workbook is held in page memory only. Clearing the file, clearing temporary
data from T-eraClip, or closing the page removes this review from the UI. No file
contents are stored in extension storage or sent to a service. Existing recovery
journals and the test-only No. writer are unchanged and remain separate.

File-size and timeout controls are not a complete defense against hostile ZIP
archives or memory exhaustion. Expanded-size validation is not implemented.
This initial internal workflow is for known-source project workbooks, not a
general untrusted-document processing service.

## Verification

Automated tests cover classification, legacy outlines, source completeness,
limits, field preservation, safe text rendering, worker cancellation/timeout,
sheet selection and clearing. Set `TERA_IMPORT_TEST_FILE` to the local source file
to run the opt-in real-workbook reader test; it is not copied into the repository.

The supplied KCN Hung Yen file was read using the installed reader: 773 data
rows, 143 missing codes, 229 missing original WBS values, maximum outline depth 9.
This is not a live import or proof of field compatibility with 1Office.

## Step 2: local normalization proposal

The CNTT HH1 workbook is a structure reference only, not a counterpart of the KCN
Hung Yen source. Its hierarchy and code-generation formulas inform configurable
proposal rules; its actual task codes/data are not copied into another project.

- Default to update mode. Only exact unique existing codes or explicit row-to-code
  mappings can identify old tasks. Complete system inventory is required.
- First-time coding mode must be selected deliberately. Add the user-supplied
  project prefix to item codes, then append the package code to the full item
  parent code (not directly to the project prefix). Allocate child suffixes from
  the full parent code: project.item.package.parent.child and deeper descendants.
  ERP components containing dots are opaque strings, not extra hierarchy levels.
  Intermediate configured levels also extend their immediate parent's code.
  Explicit full codes for new rows must extend the parent's full code, or the
  project prefix at item level. Inconsistent full codes are blocked, not repaired.
  Existing uniquely matched codes retain their old format and IDs in update mode.
  Never use regenerated suffixes to identify old tasks. Catalog membership is not
  yet checked; examples demonstrate composition, not approved ERP codes.
- In update mode, only explicitly marked new Excel rows can receive proposed new
  codes. Reserve existing codes and explicit file/mapping codes before allocation.
- Configure the first task level and package level; phase/project rows above the
  task level are excluded. Group names are suggestions, not confirmed business
  classifications. Missing parents, duplicates, invalid codes and identity
  conflicts block affected rows and dependent descendants.
- No. and all original cells are preserved. No actual dates, status, assignment
  or dependency conversion is performed. No write/export approval is provided.
- Changing files/sheets or clearing temporary data discards the proposal setup.
  Configuration changes invalidate the shown preview. No persistence/network use.

### ERP catalog within step 2

- Select a separate known-source `.xlsx` catalog up to 10 MB, with the exact sheet
  name (default `TDR`). A dedicated worker reads only this sheet using `readSheet`.
  Parsing has a 30-second deadline including the file read, a 100,000-row/20-column
  selected-sheet limit and a 20,000 retained-entry limit. Other sheets are not parsed.
- Recognize numbered `Danh mục gói thầu` and `Danh mục hạng mục` sections with
  `Mã`, `Tên` headers. The latter is the ERP project catalog in the supplied file,
  not a source of task-outline item codes. Retain only code/name/source-row entries
  from these two catalogs. Personnel/bank/customer/other sections are not posted
  back from the worker or persisted. They can still transiently exist in the
  selected-sheet parser's worker memory; ZIP unpacking is not section selective.
- Exact duplicate code/name entries are deduplicated. One code with conflicting
  names is ambiguous and cannot be selected. Duplicate names with different codes
  remain distinct options. Invalid rows warn about incomplete catalog coverage.
- Case-sensitive code membership is automatic; accent-insensitive token/name search
  only ranks up to 20 suggestions. No suggestion, even a unique name, is applied
  without a selection. A catalog code does not prove the task identity, project
  applicability or name equivalence. No prefix/suffix guessing across ERP catalogs.
- Search gives content terms more weight than generic consultancy/design wording,
  expands `QH` for search only, and requires explicit numeric pairs such as `1/2000`
  to match. Exact code/prefix search remains available. These are retrieval rules,
  not business equivalence or ERP identity rules. Same-description codes remain
  distinct and are annotated for confirmation with the chief accountant. Price,
  contractor and contract scope are not inferred from a two-column catalog. Leaving
  a code unselected keeps its row/descendants blocked, without preventing review of
  other rows. No pending confirmation is persisted after closing/clearing the page.
- Source package components, explicit full codes with known parent boundaries, and
  chosen package codes must match one catalog entry. Missing/ambiguous codes block
  the affected row and descendants. If project entries are present, the project
  prefix must match exactly one; absent project entries explicitly remain unverified.
- A chosen package component is composed under the full item parent code. It never
  replaces source cells. It can be used only for initial coding or explicitly new
  rows. Update-mode old records retain their code/ID; ERP choices cannot rename them.
  Conflicts with manually mapped full codes require correction instead of precedence
  guessing. Item/segment codes are not claimed to be validated by this catalog.
- Loading/replacing/removing a catalog invalidates the preview and clears choices;
  mode/level changes also clear choices. Sheet changes cancel parsing. Late results
  after clear, timeout or unmount are ignored. File/sheet/global temporary-data
  clearing remounts this panel. No network requests, storage or writes are involved.
- Reader/worker UI tests use mocks; the opt-in `TERA_ERP_TEST_FILE` test reads the real
  supplied workbook with the Node reader: TDR has 35 project entries and 2,243 unique
  package entries. This is not an in-Chrome end-to-end or actual import test.

The existing hostile-ZIP/memory-exhaustion limitation applies to this larger
catalog upload too. These controls are for internal, known-source workbooks and
do not implement expanded-archive-size validation.

## Step 3: selected-field before/after comparison

Available below the code proposal. All field checkboxes start off. The user must
select a field and its Excel column. The only permitted target fields are name,
raw No., planned start/end, and baseline start/end. ID, code, parent, actual dates,
progress, discussions, attachments, status, assignment and MQH are not update fields.

- Existing rows require a complete snapshot and explicit project confirmation.
  ID/code uniqueness is checked again against the current in-memory inventory.
  This is not a live re-read and does not establish concurrency safety for writes.
- File-parent and snapshot-parent differences require review; no reparenting is
  proposed. No. targets currently held by other tasks are flagged for later planning.
- Plan-date mappings require a separate semantic confirmation. Known actual-date
  column names are rejected as plan-date sources, even if the checkbox is checked.
- Date comparison accepts validated ISO or day/month/year calendar dates, including
  midnight Date objects. Nonzero times, Excel serial numbers and invalid dates are
  not silently converted. Resulting start/end order is checked for selected pairs.
- Blank/NA source means keep, not clear. Missing snapshot fields remain unknown.
  No. comparison uses sourceNo only, never generated outline numbering as fallback.
- Each field is changed/same/keep/unknown/invalid/draft. Rows are grouped as changed,
  same, review, draft, blocked, or excluded, with 25 rows per page and a status filter.
- First-time/new rows are drafts only, never assigned existing IDs. Other tasks
  absent from the file are not deletion candidates. Values are rendered as text.
- Mapping/selection/confirmation changes invalidate the preview. Changing code
  settings, file, sheet or snapshot identity discards the nested comparison setup.

No write commands, approval state, exports or API calls are created by step 3.
Preserving protected data after an actual update still requires later test-project
verification. Keep MS Project as the schedule calculation source until dependency
behavior is tested. Do not feed proposals into the test No. writer.
