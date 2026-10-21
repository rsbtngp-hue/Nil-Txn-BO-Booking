# Apps Script upgrade — v4.2

The repository now contains a complete, integrated Apps Script file. Use that file rather than manually merging the older add-on snippets.

## 1. Replace the existing Apps Script code

Open the Google Apps Script project currently used by the dashboard's `APPS_SCRIPT_URL`.

1. Open the existing main script file (usually `Code.gs`).
2. Copy the complete contents of [`Nil-Txn-BO-Booking-API-v4.2.gs`](https://github.com/rsbtngp-hue/Nil-Txn-BO-Booking/blob/main/Nil-Txn-BO-Booking-API-v4.2.gs).
3. Replace the old main script's contents with this complete code and save. Do not leave the old `doGet`, `doPost`, or helper functions in another active script file, or Apps Script will report duplicate function/constant declarations.
4. The code preserves the existing sheet names and upload actions. It does not clear your existing sheet data merely by being saved or deployed.

## 2. Deploy a new version

1. In Apps Script choose **Deploy → Manage deployments**.
2. Edit the current web-app deployment.
3. Choose **New version**, then deploy.
4. Keep the same existing `/exec` URL in the dashboard HTML if you edited the existing deployment.

## 3. What v4.2 adds

- `GET ?action=getDailySince&since=YYYY-MM-DD`: returns Daily Nil Transactions from the requested date onward.
- `GET ?action=getBookingSince&since=YYYY-MM-DD`: returns booking rows from the requested booking date onward.
- `GET ?action=getExclusions`: returns the shared excluded office IDs.
- `POST {"action":"saveExclusions","excludedOfficeIds":["..."]}`: updates the shared exclusion list.
- A dedicated `Dashboard Office Exclusions` sheet is created automatically on first use; its columns are `office_id, office_name, excluded, updated_at`.
- Existing `health`, `getMaster`, `getDaily`, `getBooking`, `status`, `beginSync`, `saveBatch`, and `commitBooking` actions remain available.

## 4. Verify after deployment

Open these using the current deployed `/exec` URL:

- `?action=health` — should report version `4.2-incremental-exclusions`.
- `?action=getDailySince&since=2026-10-15`
- `?action=getBookingSince&since=2026-10-15`
- `?action=getExclusions`

The exclusions sheet is created when `getExclusions` is first called. Test by excluding one office, then open the dashboard in another browser and confirm that the exclusion appears there too.

## Important limitations

- A browser with no local IndexedDB data must download the full dataset once.
- Incremental transfer requests rows on or after the local latest date, so additions for the same date are included. Client-side merging must use the dashboard's stable row keys to avoid duplicates.
- The date-based endpoint scans the sheet's date column to locate matching rows, but transfers only matching rows to the browser. This reduces JSON transfer substantially; it does not eliminate the Apps Script work of scanning dates.
- The script cannot be deployed to your Google account by a GitHub commit. You must paste/save and deploy it in Apps Script as described above.
