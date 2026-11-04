# CE Site Export

Counts, per SharePoint site, how many files have sensitive information type (SIT) or trainable classifier hits, grouped by document type (Word, Excel, PDF, etc.). The data comes from Microsoft Purview Content Explorer (`Export-ContentExplorerData`).

## What you need

- **Windows PowerShell 5.1** (`powershell.exe`), not PowerShell 7. `Connect-IPPSSession` is unreliable in PS7.
- The **Content Explorer List Viewer** role group in the tenant (needed for item-level export).
- The ExchangeOnlineManagement module.

## Files

| File | What it is |
|---|---|
| `Export-CESiteSummary.ps1` | The script |
| `Sites.example.csv` | Template for the site list. Copy it to `Sites.csv` and add the real sites. `Sites.csv` is never uploaded. |
| `Tags.csv` | Full list of trainable classifiers / SITs to sweep |
| `TestTags.csv` | Three tags for a quick test run |
| `DECISIONS.md` | Why the script works the way it does |

## How to run

```powershell
# In Windows PowerShell 5.1
Connect-IPPSSession -UserPrincipalName you@client.com
.\Export-CESiteSummary.ps1 -SitesCsv .\Sites.csv -TagsCsv .\Tags.csv `
    -UserPrincipalName you@client.com -OutDir C:\MDPScripts\CE_SiteExport
```

Quick test on one site:

```powershell
.\Export-CESiteSummary.ps1 -SiteUrl https://client.sharepoint.com/sites/it-team -TagsCsv .\TestTags.csv
```

**If it stops** (token expiry, RDP drop), run the same command again. Finished site/tag pairs are cached, so it picks up where it left off. Use `-Force` to start fresh, or `-RefreshProbe` to re-check which tags have hits.

## Output (in `-OutDir`)

- `SiteSummary_<date>.csv`: one row per site / tag / doc type, with file count, hit count and highest confidence
- `SiteDetail_<date>.csv`: one row per file
- `cache\`, `probe_*.csv`, `run.log`: working files used for resuming

All output contains client data and is excluded from Git by `.gitignore`.
