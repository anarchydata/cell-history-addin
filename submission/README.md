# Cell History — AppSource submission kit

Everything you need for Partner Center is in this folder.

## Upload these

| File | Where it goes |
|------|----------------|
| `manifest.xml` | Partner Center → Packages (upload this file) |
| `listing-copy.txt` | Partner Center → Store listing fields |
| `test-notes.txt` | Partner Center → Notes for certification |
| `screenshots/README.txt` | Take 1–2 screenshots as described, then upload them in Partner Center |

Hosted web files (already pointed to by the manifest) live at:

https://anarchydata.github.io/cell-history-addin/

After any code/icon change, redeploy `dist/web/` to that GitHub Pages site, then re-validate:

```
npx office-addin-manifest validate manifest.xml -p
```

## Partner Center account

1. Enroll in **Microsoft 365 and Copilot** program.
2. Publisher name must match manifest ProviderName: **Goldmeier Consulting Co LLC**
3. New offer → Office Add-in → name **Cell History** → upload `manifest.xml`
