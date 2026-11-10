# T-eraClip

**Clip anything. Turn it into action.**

T-eraClip is a local-first Chrome Extension for collecting and exporting 1Office work discussions and capturing selected webpage text. The current interface is Vietnamese. Supported 1Office origins are `https://office.meygroup.vn`, `https://testoffice.meygroup.vn`, and `https://testoffice.meygroup.vn:9443` (the observed test deployment).

Collection requests and exported task links retain the source origin, including its port. Project inventories cannot mix origins. Access-history collection and personnel refresh remain pinned to the same environment. No persistent host permissions are added. The No. write experiment remains restricted to the existing allowlisted project/tasks on the production origin; adding test-site reading does not enable test-site writes. After rebuilding, reload the unpacked extension and invoke it again from the intended source tab to grant `activeTab` access. Automated checks use fixtures/mocks, not live-server compatibility guarantees.

## Cài Cho Người Dùng (Windows / macOS)

**Không cần cài Node.js, npm, Git hoặc chạy lệnh.** Người phụ trách build một lần, rồi gửi bản cài đã đóng gói cho người dùng.

1. Nhận file `T-eraClip-Thao-luan-Noi-bo.zip` từ người phụ trách. Đây là bản **chỉ thu thập thảo luận**, không có Capture hoặc công cụ thay đổi mã/No.
2. Giải nén ZIP vào nơi cố định, ví dụ `Documents/T-eraClip`. Windows: nhấp chuột phải file ZIP, chọn **Extract All / Giải nén tất cả**.
3. Mở Chrome, nhập `chrome://extensions` vào thanh địa chỉ.
4. Bật **Developer mode / Chế độ dành cho nhà phát triển** và giữ chế độ này bật khi dùng bản unpacked.
5. Bấm **Load unpacked / Tải tiện ích đã giải nén**, chọn thư mục `T-eraClip-Thao-luan` có file `manifest.json` ngay bên trong.
6. Ghim tiện ích, đăng nhập 1Office bằng tài khoản của mình và bắt đầu thu thập thảo luận.

Không chọn file ZIP, không chọn thư mục mã nguồn GitHub, không mở file HTML để cài. **Code > Download ZIP trên GitHub là mã nguồn, không phải bản cài đã build.** Không xóa/di chuyển thư mục đã nạp. Bản này không tự cập nhật; nhận bản mới, thay file trong cùng thư mục rồi bấm Reload trên thẻ tiện ích.

Nếu máy công ty chặn Developer mode hoặc Load unpacked, liên hệ CNTT, không vượt chính sách. Cài unpacked không giới hạn được người sử dụng theo danh sách chỉ định: ai có gói cài đều có thể sao chép; quyền xem dữ liệu vẫn do tài khoản 1Office quyết định.

Xem [hướng dẫn cài bản thảo luận](docs/INSTALL_DISCUSSIONS.md). Đây là bản pilot, chưa được Google xét duyệt và chưa xác minh trực tiếp trên máy Windows của key user. Quy trình nạp unpacked tham khảo [Chrome Developers](https://developer.chrome.com/docs/extensions/get-started/tutorial/hello-world#load-unpacked).

## MVP Status

The current working version is `0.1.0`, intended for internal pilot testing. It is not published to or approved by the Chrome Web Store. The earlier local MVP has since been extended with 1Office collection and export.

Typecheck, automated tests and production build are available. Passing these checks does not establish production readiness, security certification, or compatibility with colleagues' accounts and devices.

## Features

- Collect discussions on the current supported 1Office page
- Collect discussions across project tasks and the current filtered task list, including pagination
- Review collection progress and warnings on a separate page
- Export original data to Excel/JSON, with work codes, No., hierarchy, comments and replies
- Retain attachment metadata; original attachment download is not implemented
- Open the earlier local text-capture tools through the Capture button and return to the main menu
- Capture selected webpage text from the right-click context menu
- Preserve the selected text, source URL, source page title, and timestamps
- Store work items in `chrome.storage.local`
- View captured items in a newest-first popup list
- Edit item title, status, priority, and notes
- Filter items by status and priority
- Generate daily, follow-up, waiting, and completed reports
- Generate vendor follow-up, internal follow-up, leadership summary, and meeting action templates
- Copy generated reports and templates to the clipboard
- Save the default report-language preference for future report localization
- Export items and settings as a versioned JSON backup
- Validate and replace local data from a JSON backup
- Clear all local items and settings with confirmation
- Safely ignore invalid stored items during reads
- Reset all Capture items with confirmation, without removing settings or original 1Office data

## Local-First Data

T-eraClip stores Capture items and settings in the current Chrome profile using `chrome.storage.local`. 1Office review/collection sessions use page memory and/or `chrome.storage.session`. There is no T-eraClip account, backend, cloud sync, external database, or cross-device synchronization. 1Office collection uses the user's existing signed-in session on the supported site.

Export important data before clearing extension data, removing the extension, or resetting the browser profile.

## Privacy And Security

- Capture occurs only after the user selects text and chooses **Save to T-eraClip**.
- The extension uses `storage`, `contextMenus`, `activeTab`, and `scripting` permissions.
- It has no persistent host permissions or automatically registered content scripts. User-invoked collectors are injected on the supported 1Office tab, read DOM/API data, and load additional items within the chosen scope.
- Collection uses the user's 1Office permissions, not the developer's admin privileges.
- Selected text and item contents are not logged by a webpage content script.
- Backup imports are validated before confirmation and replacement.
- Items and settings are written together in one Chrome Storage operation during import.
- Invalid stored items are skipped without logging their contents or rewriting storage.

## Tech Stack

- Chrome Extension Manifest V3
- React 18
- TypeScript
- Vite
- Tailwind CSS
- Chrome Storage API
- Chrome Context Menus API
- Clipboard API

## Project Structure

```text
public/manifest.json             Extension manifest
src/background/index.ts         Context-menu registration and capture
src/popup/                      Popup application
src/options/                    Settings and local data management
src/shared/storage.ts           Chrome Storage access and validation
src/shared/reportGenerator.ts   Report generation
src/shared/templateGenerator.ts Message-template generation
src/shared/types.ts             Shared data types
docs/                           Product, architecture, and release documents
```

## Developer Setup (Not Required For End Users)

This section is only for maintainers building from source. End users receiving a prebuilt package should follow the installation section above instead. These npm commands also apply on Windows with Node.js installed, but recipients do not need to run them.

Requirements:

- Node.js and npm
- Google Chrome or another Chromium-based browser that supports unpacked extensions

Install dependencies:

```bash
npm install
```

Check TypeScript:

```bash
npm run typecheck
```

Build the extension:

```bash
npm run build
```

The unpacked extension is generated in `dist/`.

## Test A Local Developer Build In Chrome

1. Run `npm run build`.
2. Open `chrome://extensions`.
3. Enable **Developer mode**.
4. Select **Load unpacked**.
5. Choose the repository's `dist` directory.
6. Confirm T-eraClip loads without manifest errors.

After rebuilding, return to `chrome://extensions` and reload the extension.

## Internal Distribution

For the discussions-only profile, maintainers run `npm run build:discussions` and `node scripts/package-discussions.mjs`. Send the generated `T-eraClip-Thao-luan/` folder (zipped for transport) and [installation guide](docs/INSTALL_DISCUSSIONS.md), not the source repository. The generated Store ZIP can also be extracted and loaded by selecting the folder containing its root `manifest.json`; its filename does not mean it has been approved or published by Google.

After validation and build, run `node scripts/package-release.mjs` to create a dated pilot ZIP, Chrome Web Store ZIP, file inventory and SHA-256 checksums under ignored `releases/`.

- [Pilot installation and test checklist](docs/INSTALL_INTERNAL_TEST.md)
- [Private Chrome Web Store release plan](docs/CHROME_WEB_STORE_RELEASE.md)
- [Privacy policy draft, requires approval and a hosted URL](docs/PRIVACY_POLICY_DRAFT.md)

The pilot is loaded unpacked. The Store ZIP is only a prepared upload package; publishing requires a registered developer account, truthful privacy disclosures, reviewer test access and Google's approval. Do not send the whole repository or browser profile to testers. Backup Capture before switching from unpacked to Store installation; extension IDs and storage may differ.

## Basic Usage

1. Select text on a normal webpage.
2. Right-click the selection.
3. Choose **Save to T-eraClip**.
4. Open the extension popup to review the captured item.
5. Edit its title, status, priority, or notes as needed.
6. Use filters to narrow the visible list.
7. Generate and copy a report or message template.
8. Open the Options page to manage language preference, backups, or local data.

Reports and templates use all saved items, not only the currently filtered popup results.

## Smoke Checklist

- The unpacked extension loads from `dist` without manifest errors.
- **Save to T-eraClip** appears for selected webpage text.
- A captured item includes its original text, source URL, and source title.
- Popup edits persist after closing and reopening the popup.
- Status and priority filters work, including filtered-empty behavior.
- At least one report and one message template generate and copy successfully.
- Options opens and saves the report-language preference.
- Export, clear, and valid replace-only import complete successfully.
- Invalid JSON or invalid backup schema does not replace existing data.
- Data persists after extension reload or browser restart.

## Known Limitations

- Version 0.1 is distributed only as an unpacked extension.
- Data exists only in the current browser profile unless manually exported and imported.
- Import replaces existing local data; merge import is not supported.
- The saved report-language preference is not yet connected to generated output.
- Item category, due date, owner, tags, and type are present in the data model but are not editable in the popup.
- Items cannot be deleted individually.
- Stored items missing required fields are skipped rather than migrated or repaired.
- Chrome Storage operations are hardened but are not claimed to provide full database-style ACID transactions.
- Automated tests exist; real-device/account pilot testing is still required before wider distribution.

## Future Backlog

Potential work after v0.1 stabilization:

- Wire report-language preference into report generation
- Add carefully scoped editing for additional item fields
- Add individual item deletion and recovery safeguards
- Add automated tests for storage validation, imports, reports, and templates
- Improve storage concurrency handling
- Evaluate packaging and Chrome Web Store readiness
- Consider optional synchronization or AI only as separately scoped future products

Cloud sync, a T-eraClip backend/login, analytics, and AI are not part of the current pilot. The manual AI experiment was removed.

## Development Principles

- Keep the core product local-first and usable without an account.
- Request the minimum Chrome permissions required for current behavior.
- Do not log captured user content.
- Validate external and stored data at runtime.
- Prefer small, reviewable changes that preserve existing behavior.
- Keep build and typecheck commands passing.
- Treat new features as explicit scope decisions rather than incidental additions.

## Compact Project Discussion Entry

On a project page, the popup's main action checks the task inventory first, then
opens the existing discussion collection page automatically only when the list is
complete and nonempty. Partial lists stop for explicit continuation; failures and
empty lists never launch collection. A check-only action remains under the extra
options for the code-normalization workflow. Inventory warnings remain available
in both the popup and collection page. Collectors, request limits, retries, export
filters and incomplete-export confirmation are unchanged. Filtered-list collection
keeps its existing separate flow. This reduces clicks, not server request volume.

## Project Documentation

- `docs/PRD_v0.1.md`
- `docs/MVP_SCOPE_v0.1.md`
- `docs/ARCHITECTURE_v0.1.md`
- `docs/CODEX_PROMPTS.md`
- `docs/RELEASE_NOTES_v0.1.md`
