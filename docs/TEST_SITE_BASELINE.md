# Test-site baseline, read only

Scope: project 203 on `https://testoffice.meygroup.vn:9443` only. This is preparation for a later update test, not authorization to save changes.

## Run

1. Open the test project and invoke T-eraClip on that tab.
2. Choose `Tuy chon khac` and the inventory-only action. Wait for the complete inventory.
3. Open the project normalization workspace from the same popup.
4. In `Chuan bi kiem thu tren site test`, create the baseline and five samples.
5. Review the table and coverage gaps. Download the JSON locally for a before/after comparison.

The runner re-reads the inventory, validates environment, project, counts, unique task IDs and hierarchy, then reads only five discussions using the existing read API. Version 3 reads two detail widgets per sample in the background. It never opens an edit form, submits a save, navigates the source tab as a fallback, or persists a write journal. Selection uses parent IDs, not task names or No. values. Duplicate No. values are not duplicate task identities.

## Evidence Limits

- All inventory fields and the original hierarchy are retained, with collection timestamps and source URLs.
- Missing original No. values remain unknown; blank strings remain blank. Displayed outline numbers are not substituted into the original field.
- Version 2 stores HTML-detail evidence separately in each sample's `details`, without overwriting the API inventory. The HTML must identify the expected task through its discussion container and link to the test project. Only unambiguous `Ma cong viec` / `No.` labels (with their actual Vietnamese accents in code) are read. Placeholder `--` remains display-only evidence, never a verified raw blank. Login pages, client-side shells, and conflicting identities/labels stay unverified. No site scripts are evaluated and no guessed field IDs or follow-up endpoints are used.
- Version 3 replaces the live HTML reader with the observed GET `/features/core-app-features-board-objects-widgets/data` endpoint. Widget IDs 23 (general information) and 1915 (custom fields), tab 3 / `visual-3`, are explicitly pinned to the inspected project-203 layout. They are not universal widget or custom-field IDs. Changed layout, widget name, URL or task identity fails closed, with no discovery scan or retry. Each request has a 5 MB streaming limit; both share a 30-second deadline, and the source page must remain unchanged.
- Widget JSON contains inert `elem` markup, encoded bodies and template children. The parser uses JSON.parse and DOMParser only, never executes callbacks, and checks both widgets' task IDs plus the unique project link in the general-information `Du an` field. Only the two requested labels are retained, not raw responses, headers or callbacks. Explicit `content: ""` is recorded as `empty`; missing content remains `missing`. Neither state overwrites inventory `sourceNo` or authorizes an edit. The former rendered-HTML parser remains for existing fixtures, not as a live fallback.
- Discussion read failures and incomplete responses are retained per sample. No automatic retry through an edit/detail workflow is used.
- Zero comments in these five samples does not prove there are no comments elsewhere in the project.
- Completion percentage, report-result content and all attachments are not covered by this snapshot. Do not claim that they are preserved by a later update based only on this data.
- The read spans multiple requests and is not an atomic snapshot or full backup.
- JSON stays in page memory until downloaded. Clearing temporary data invalidates the panel; closing the tab loses undownloaded results.

## Current Verification

Automated fixture tests cover selection, scope validation, missing parents/cycles, original No. semantics, read failures, cancellation, explicit user initiation, duplicate-click prevention and result disposal. A live inventory-only run reported 864 tasks and eight levels on the test project. Direct browser-tool access to the extension workspace was blocked by the browser URL policy; completion of the five-sample live run requires the user's UI result or downloaded JSON. No live writes were executed.

The user's baseline ending `2026-10-21T10-51-34-331Z` confirms 864 tasks but no API `sourceNo` values or codes. Earlier direct UI inspection of task 7698 showed code `.18902.18908`, No. placeholder `--`, progress 0, and no discussions. This establishes at least one code omission in the list reader, not an empty code on the site. Version 2's detail parser has automated tests using that observed DOM shape; live background-HTML extraction has not yet been confirmed. Do not infer the API field mapping or claim a fully repaired inventory from those parser tests.

On 2026-10-22, read-only Chrome Network inspection of task 7698 confirmed that quickview returns declarative JSON, then loads detail widgets. Widget 23 returned task 7698, project 203 and code `.18902.18908`; widget 1915 returned the explicit `No.` field with empty content, rendered as `--` in the UI. The prior raw-page parser could not read these nested responses. Version 3 fixtures model this wire format, including encoded layers and template children. A successful live extension run of the new reader is still required; inspecting one task does not establish field coverage for 864 tasks or verify preservation of discussions/progress/results. No live writes were performed.

The user's subsequent v3 screenshot still reported five unverified detail reads. Follow-up Network inspection exposed a mismatch in the original fixtures: they HTML-escaped the entire JSON attribute, while the observed response uses a raw JSON attribute with an encoded body. DOMParser already decodes the body's entities when reading that attribute. Decoding the body again corrupts quoted nested field attributes. Version 3.1 preserves literal `<elem ...>` bodies and decodes an `&lt;elem ...>` envelope once only, never recursively. A new regression test failed before this change and passed afterward; wrong task/project and missing-versus-empty cases are also covered with this envelope. Failures now record a fixed diagnostic stage, not raw payloads or exception text. This fixes a reproduced parser defect, but still requires a new live extension result before declaring all five reads successful.

## Version 4: Find a Preservation Sample

The user's v3.1 screenshots show codes read for all five samples and explicit empty No. responses, with zero saves. This verifies those sample reads, not the 864-task inventory field mapping or a preservation test.

After creating a baseline, choose `Tim mau co thao luan` (the UI uses Vietnamese accents). Each explicit click reads a fresh project-203 inventory, validates its complete ID/parent set against the baseline, then tries up to 20 discussions sequentially. It stops at the first complete nonempty discussion and reads only that candidate's two detail widgets. Further batches require another click. A manual ID must exist in the validated test inventory. Failed, incomplete or mismatched discussion reads remain `unverified`, not empty; a manual ID read can revisit them. No edit/navigation/write fallback or automatic creation of test data is used.

The detail parser also reads numeric progress from the previously observed `Tien do` field's `progress` component (`type: progresstask`, finite numeric `percent` from 0 to 100). Zero is valid. Task status is not a percentage. Report results remain explicitly `not-collected`; the similarly named expected-results custom field is not treated as a report. Discussion attachment metadata is not a full file backup. Even a found candidate is not approved for a write or sufficient to claim preservation.

Search state stays in page memory. The existing local JSON download includes an additive `preservation` record with attempted IDs/statuses, timestamps and the last found sample. Searching again retains the previous candidate if no new one is found, but does not refresh its evidence. Cancellation stops between requests; an in-flight read can take up to its 30-second deadline, and the incomplete batch is discarded. Closing/clearing the page disposes late results. Automated tests cover the bounds, continuation, scope/tree changes, incomplete reads, progress semantics, cancellation, explicit UI initiation and export. Live discovery of a discussion-bearing test sample and report-result collection are still unverified.
