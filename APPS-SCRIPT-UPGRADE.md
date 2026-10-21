# Apps Script upgrade for incremental sync and shared office exclusions

The dashboard HTML now checks remote dataset metadata first. When metadata is unchanged, it skips the download. When rows have been appended, it requests only records whose date is on/after the locally cached latest date. If the remote row count did not increase (for example, a month was replaced), the HTML intentionally performs a full refresh to avoid retaining stale historical rows.

## 1. Add the helper file

In the same Google Apps Script project used by the dashboard's current `APPS_SCRIPT_URL`, create a new script file named `IncrementalSyncAndExclusions.gs` and paste the contents of this repository's `IncrementalSyncAndExclusions.gs`.

## 2. Add three GET routes in the existing doGet(e)

Inside the existing `doGet(e)`, add these branches before the final unsupported-action `else`:

```javascript
} else if (action === 'getDailySince') {
  result = getDailySince_(params.since);
} else if (action === 'getBookingSince') {
  result = getBookingSince_(params.since);
} else if (action === 'getExclusions') {
  result = getExclusions_();
}
```

Keep the existing actions and their current branches unchanged.

## 3. Add one POST route in the existing doPost(e)

Add this branch to the existing action chain:

```javascript
else if (request.action === 'saveExclusions') result = saveExclusions_(request.excludedOfficeIds);
```

For example, place it after the existing `commitBooking` branch and before the unsupported-action error.

## 4. Deploy

1. Save the Apps Script project.
2. Choose **Deploy → Manage deployments**.
3. Edit the existing web-app deployment, select **New version**, and deploy it. Keep the existing /exec URL if you edited that deployment.
4. Run the dashboard once as an administrator while online. The new `Dashboard Office Exclusions` sheet is created automatically when the endpoint is first used.
5. Verify these URLs in a browser (append the query string to your existing /exec URL):
   - `?action=getDailySince&since=2026-10-15`
   - `?action=getBookingSince&since=2026-10-15`
   - `?action=getExclusions`

The shared exclusion sheet stores only excluded offices; unchecking an office removes it from the shared list. The HTML retains a local cache for offline viewing, but the shared Google Sheet is the online source of truth.

## Notes

- A first visit on a browser with no local database must download the full datasets once. Later visits use the local IndexedDB first and skip downloads when remote metadata has not changed.
- If the latest row count increases, only rows from the local latest date onward are transferred; the latest date is included to capture additional records for that same day. Records are merged by `office_id + transac_date`.
- If a dataset's row count falls or remains unchanged while its write metadata changes, the client falls back to a full refresh so replaced/deleted historical data is not left stale.
- Office exclusions are applied through the same `excludedOfficeIds` set already used by both Nil and Booking dashboard calculations.
