# T-eraClip Security Follow-up, 2026-10-12

## Approved scope

Keep 1Office responsible for login and authorization. No separate extension
login, backend, new permissions or account-based report binding was added.
The operator reports that the site's inactivity logout period is three days;
this configuration was not independently inspected and is not a reason to assume
every request must succeed.

## Changes

- Added **Xóa dữ liệu tạm** to discussion, project and access-history review
  pages, with confirmation. It removes only the two report session keys and
  invalidates in-memory reports in listening extension pages using a Chrome
  session-storage revision event. Capture and unrelated stored data are retained.
- Clearing does not delete downloaded files, change 1Office records or provide
  forensic erasure of browser/OS memory. An export already in flight in another
  tab cannot be revoked. Clear buttons are disabled during collection/export in
  their own page; another page's clear event cancels or discards late results.
- Comment, quickview and Gantt requests now have a 30-second AbortController
  deadline covering fetch and body reading. History/personnel deadlines were
  aligned from 20 to 30 seconds. The deadline is per request, not per project.
- Timeout/HTTP failures are reported as failed tasks rather than silently
  starting another quickview/DOM request for the same failed fast-path request.
  Project review offers **Thử lại thu thập**, which reloads and recollects the
  stored inventory. This retries the batch, not only failed tasks. No automated
  retry storm or concurrency increase was added.
- Data fetches refuse HTTP redirects. No legitimate special redirect was
  observed or verified on a live 1Office session in this turn. Failed requests
  ask the operator to check login and retry; no redirect exception was invented.
  DOM fallback can still navigate the source tab as before, with origin checks.
- Progress messages require the extension sender ID, main frame, expected source
  tab, approved source origin, per-collection ID and nonnegative safe-integer
  counts. Popup subscriptions exist only during their own collection and are
  removed in finally. Messages for other tabs, frames or old runs are ignored.

## Dependency check

The authorized npm audit on 2026-10-12 still reported 16 affected package entries:
2 critical, 10 high, 4 moderate. These counts include dependent packages, not
16 distinct proven exploits. No dependency upgrades were made in this change.
Major toolchain upgrades suggested by npm need a separate compatibility review;
colleagues using only discussion export does not remove developer-machine risk.

The production-only audit was repeated separately and reported zero known
advisories. This is not a guarantee that runtime code has no vulnerabilities.

## Verification and limits

- Tests use local fixtures, mocked Chrome APIs, mocked HTTP responses and fake
  timers. No production load test, account-switch exercise or backend permission
  attack was performed.
- Tests cover 30-second expiry on all five request helpers; 401/403/429 rejection;
  a failed task not entering fallback; progress sender/run validation; cross-page
  clear events; cancelled confirmation; unrelated storage events; storage errors.
- Existing active-account, company-scope, XLSX and Capture tests remain in scope.
- Historical audit finding F1 is mitigated by manual clearing, not fully resolved:
  automatic expiry and account binding remain absent by the approved scope.
- F3 and F4 have defensive tests now. F5 (unknown fields in review JSON) remains
  unchanged. The 2026-10-09 audit is a historical baseline, not current results.
- Live Chrome/1Office validation of the new button, messages and redirect policy
  remains outstanding. Automated tests and a successful build are not security
  certification.

## Final verification

- `npm test`: 21 files, 133 tests passed. Two characterization tests still
  document intentionally unchanged account-binding/expiry and JSON-field limits.
- `npm run typecheck`: passed.
- `npm run build`: passed; updated dist is ready for local extension reload.
- `git diff --check`: passed.
- Heuristic secret scan: no pattern matches in the working files, dist and
  155 reachable local Git blobs across 23 commits. Scanner limitations from
  the original audit still apply.
- No package/lockfile, manifest permissions or 1Office server settings were
  changed in this follow-up. Existing unrelated working-tree changes were kept.
