# T-eraClip Security Review, 2026-10-09

## Scope and conclusion

This is a local source, dependency, build and mocked-test assessment. It is not
penetration testing of 1Office or a security certification. No live corporate
records, passwords, cookies or tokens were sent to a scanning service. npm audit
sent dependency information to the npm advisory service only.

Application source, manifest, package files and existing user changes were not
modified during this assessment. Added files are a test suite, a heuristic secret
scanner and this report. No dependencies were upgraded or credentials rotated.

Existing controls passed the tested scenarios, but the findings below remain
unfixed. Passing tests must not be interpreted as proving the absence of leaks.

## Findings and suggested order

### F1: Temporary reports are not bound to the authenticated account (high priority)

- `src/oneoffice/reviewSession.ts` stores the last discussion report under one
  shared session key. It has no account binding, expiry or consume-on-read.
- The reproduced test reads a report captured in 2020 repeatedly, and confirms
  that saving another report replaces the first report under that same key.
- `src/review/App.tsx` loads the saved report and exports it without checking the
  current source account. Access-history and project-review pages also retain
  collected data in page memory for local export.
- Risk: previously collected administrative data can remain available after a
  1Office logout/account change on the same Chrome profile. Concurrent review
  tabs can load the wrong last-saved report. This is not evidence that unrelated
  remote users can read extension storage.
- Account switching on a live Chrome/1Office session was not tested. The lack of
  account checks is verified from source, not a demonstrated backend bypass.
- Suggested repair: separate review IDs, bounded retention, explicit clear,
  trustworthy source-account binding and invalidation on account/session change.
  Do not treat an active-employee filter as authorization for the operator.

### F2: Development/test dependencies have published advisories (high priority)

`npm audit --json` reported 16 affected package entries: 2 critical, 10 high and
4 moderate. Entries include inherited dependency effects; they are not 16
independent exploitable application vulnerabilities.

Installed versions inspected include Vite 5.4.21, Vitest 3.2.6, Tinypool 1.1.1,
PostCSS 8.5.15 and Tailwind CSS 3.4.19. The two critical package entries are
Tinypool and its dependent Vitest. These are development/test dependencies, not
the extension's declared runtime dependencies.

`npm audit --omit=dev --json` reported zero advisories in the production dependency
graph. This is limited to known registry advisories and does not prove that
runtime libraries or the compiled artifact have no vulnerabilities.

The Tinypool maintainer describes additional exploit conditions, including an
upstream prototype-pollution primitive and affected worker options. The opened
maintainer advisory labels that issue High, while this npm audit categorizes the
affected package Critical. Retain both descriptions; no exploit was attempted.

- [Tinypool maintainer advisory](https://github.com/tinylibs/tinypool/security/advisories/GHSA-85c8-ppgw-ccpr)
- [Vite maintainer advisory](https://github.com/vitejs/vite/security/advisories/GHSA-4w7w-66w2-5vf9)

Suggested repair: review compatible patched versions and transitive dependencies,
then rerun tests/build. Do not run `npm audit fix --force` blindly or expose a dev
server beyond its intended local development scope.

### F3: Discussion/project fetches lack redirect and deadline controls (medium priority)

The test reproduces missing `redirect: 'error'` and AbortSignal on comment API
and quickview fetches. The Gantt fetch helper also lacks these controls in source.
By comparison, personnel/history requests block redirects and use a 20-second
deadline, verified by mocked tests.

Requests initially use the approved source origin in normal callers. The missing
redirect control does not itself prove exfiltration, and does not mean 1Office
cookies would automatically be sent to another domain. It permits redirect
following instead of explicitly failing closed. Missing deadlines can stall a
batch. Suggested repair: consistent destination checks, redirect refusal,
bounded cancellation/timeouts and bounded handling of 429 responses.

### F4: Progress messages lack sender and payload validation (medium priority)

`subscribeToOneOfficeProgress()` accepts a matching message type and any truthy
progress object. The test supplies malformed progress and an unrelated sender;
the listener receives it. Sender tab/session is not checked at this boundary.

The manifest has no external messaging entry point. Therefore this is not proof
that any website can send these messages or steal data. Suggested repair: check
extension sender, expected source tab, collection ID and runtime field schema;
ignore unrelated, malformed and expired messages.

### F5: Unknown fields survive review validation and JSON export (lower priority)

The test adds a synthetic private field to an otherwise valid review object.
Validation accepts it and JSON serialization retains it. Normalized collectors
generally construct known fields, so this test does not prove a real credential
is currently retained or exported.

Suggested repair: project review/export inputs to an explicit allowlist and
validate size/structure. Do not reject harmless new fields throughout every raw
1Office response if selecting known fields is sufficient.

## Verified controls

- Manifest permissions are storage, contextMenus, activeTab and scripting; no
  host_permissions, content_scripts or externally_connectable declaration.
- Source-origin checks reject HTTP, lookalike hosts, userinfo-host tricks,
  unrelated domains and javascript URLs.
- History/personnel requests do not fetch off the approved origin; invalid
  pagination/lookup IDs are rejected before fetching.
- Mocked HTTP 401, 403 and 429 do not become usable history/personnel responses.
- Existing tests cover active-account fail-closed behavior, exclusion of locked
  and unknown accounts, exact account matching, six shared network slots,
  cancellation and removal of stale UI results after retry failure.
- Personnel HTML/icon markup is converted to text in the tested examples.
- Generated discussion XLSX was unzipped and inspected: tested values beginning
  with =, +, -, @ and tab-plus-= did not create formula elements or external
  relationships. Access export account/name cells use explicit String type.
- The built HTML pages contain no external script tags; dist/assets contains no
  source-map files. The production build completed. These static checks do not
  rule out every dynamic resource load and are not a network-traffic capture.

## Secret scan

`node scripts/security-scan.mjs` checked non-ignored tracked/untracked worktree
files, dist, and 152 reachable blobs across 23 commits in local Git refs. It
found no matches for its private-key, GitHub-token, AWS-access-ID, OpenAI-key and
literal-credential patterns. It prints only locations/types, never matched values.

Limits: this is not a comprehensive secret detector. Binary files, ignored files
outside dist, reflogs and remote-only Git history are not covered. No entropy
analysis, credential-validity test or exhaustive corporate-data classification
was performed. No finding means no pattern matches in this scope, not no secrets.

## Test results

- Baseline: 18 test files, 102 tests passed.
- Added security audit: 17 tests passed, comprising 13 control checks and 4
  characterization tests that reproduce unresolved findings F1/F3/F4/F5.
- A passing characterization test confirms the gap exists; it does not certify
  the behavior as secure. Replace these assertions with defensive expectations
  when the findings are repaired.
- Typecheck passed; production build passed.
- Final full suite: 19 test files, 119 tests passed (including the 4 reproduced
  findings, not 119 independent security guarantees).
- `git diff --check` passed. Existing worktree changes were preserved.

## Not tested and residual risks

- Live backend authorization with a second, lower-privilege account. No guarantee
  is made about 1Office's resource-level permissions or response semantics.
- Actual account switching, browser resource loads from parsed HTML, malicious
  extensions, browser compromise or OS-level access to local storage.
- Full network capture proving there are no unexpected outbound destinations.
- Load/abuse testing of the corporate server, or exploit attempts against it.
- Captured text persists locally until cleared. Exported Excel/JSON files are
  not application-encrypted or revocable after download. Handling and sharing
  those files remains an operational data-protection responsibility.
