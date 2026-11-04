# Decisions

A running log of why things are done the way they are. Newest at the top.

Template:

```
## YYYY-MM-DD – Short title
**Decision:** what we chose
**Why:** the reason
**Decided by:** who
```

---

## 2026-11-05 – Store the script in GitHub, keep client data out
**Decision:** The script, tag lists and docs go in a private GitHub repo. `Sites.csv`, the cache, logs and all reports stay local only (see `.gitignore`).
**Why:** We get a full version history and a shared copy of the script without uploading client site names or scan results.
**Decided by:** Vida

## 2026-10 (between 23 and 26 Oct) – Only reconnect when the session is actually lost
**Decision:** On a failed call, reconnect only for auth/session errors (token, 401/403, role definition). Other server errors get a pause and a retry without reconnecting.
**Why:** Reconnecting can open a sign-in prompt that waits forever on an unattended overnight run.
**Decided by:** Vida

## 2026-10 (between 20 and 23 Oct) – Build the report in one streaming pass
**Decision:** The report step uses hashtables in one pass instead of `Group-Object` / `Select-Object -Unique`.
**Why:** Those cmdlets are quadratic in PowerShell 5.1 and took hours (and GBs of memory) on sites with around 1M detail rows.
**Decided by:** Vida

## 2026-10 – Count of Hits = largest confidence bucket, not the sum
**Decision:** A file's hit count is the largest of its High / Medium / Low counts.
**Why:** Content Explorer's buckets are cumulative (a High match is also counted in Medium and Low), so adding them double-counts.
**Decided by:** Vida

## 2026-10 – De-duplicate files on full URL, never file name
**Why:** The same file name repeats across folders, so de-duplicating on name under-counts.

## 2026-10 – Probe each tag tenant-wide before going site by site
**Decision:** One cheap call per tag first (PageSize 1); tags with zero hits tenant-wide are skipped for every site.
**Why:** Cuts the number of calls from sites × all tags down to sites × tags that have hits.

## 2026-10 – Cache each finished (site, tag) pair
**Why:** Runs take hours and get interrupted (token expiry, RDP drops). Re-running the same command resumes instead of starting over.

## 2026-10 – Run in Windows PowerShell 5.1, not PowerShell 7
**Why:** `Connect-IPPSSession` breaks in PS7 (WAM broker sign-in).
