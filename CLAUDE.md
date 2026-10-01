# Project notes

PowerShell script (`GooglePhotosSync.ps1`) that automates Google Takeout for Google Photos: it
downloads the export from Google Drive and copies it into a local or SMB folder. See `README.md`.

## Keep the README comparison table up to date

`README.md` has a section "Why use this instead of doing it by hand?" with a table comparing each
manual Takeout job against what the script does. After any feature update (new option, changed
behaviour, new limitation, or a limitation that is lifted), update that table in the same change so
it still shows how the script differs from the manual steps. Add a row for a new capability, edit
rows whose behaviour changed, and revise the "Limits to be aware of" list. Don't leave the README
describing old behaviour.

## Other conventions

- Must run on Windows PowerShell 5.1 and PowerShell 7, with no extra installs. Keep the script
  ASCII-only.
- Never commit `config.json` or `client_secret*.json` (both are in `.gitignore`).
