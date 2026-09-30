# Google Photos Sync (Takeout helper)

A PowerShell script that does the tedious part of Google Takeout for you. Normally you'd download
many large archives by hand, unzip them and move the photos into the right folders. This script
fetches your Google Photos Takeout exports from Google Drive and copies the photos into a local or
network (SMB) folder, keeping Takeout's folder layout.

It needs nothing beyond what ships with Windows 10/11: Windows PowerShell 5.1 or PowerShell 7.

> **Why Takeout?** Since March 2025, Google's Photos API no longer lets third-party tools read a
> whole Google Photos library. Google Takeout is the practical way to get a full copy, and this
> script automates everything around it.

## What it does

1. Signs in to Google with **read-only** access to Google Drive. The first time, your browser
   opens for you to approve it. After that, runs are unattended.
2. Finds the Takeout archives (`takeout-*.zip`) that Takeout delivered to your Drive.
3. Downloads them one part at a time. Downloads resume if interrupted and are checked against
   Google's checksum.
4. Copies the photos and videos into your destination folder, for example
   `Destination\Photos from 2023\IMG_1234.JPG` or `Destination\<Album name>\...`.
5. Sets each file's date to when the photo was **taken**, using Takeout's `.json` metadata files.
6. Remembers which exports it has imported, so a scheduled run only acts on new ones.

**One-way only:** files go from Google to your folder. The script never deletes anything, either
in your folder or in Google.

**Name collisions:** if a file with the same name is already there:

- **Same content:** it's skipped.
- **Different content:** both are kept, and the new one is saved as
  `IMG_1234 (gphotos-collision-1).JPG`.

Re-running the script on the same export is therefore safe.

## One-time setup

### 1. Create a Google OAuth client (about 5 minutes)

Google requires every app, including a personal script, to have its own client ID.

1. Go to <https://console.cloud.google.com/> and create a project, for example "Photos Sync".
2. Go to **APIs & Services → Library**, search for **Google Drive API** and click **Enable**.
3. Go to **Google Auth Platform** (older consoles call it "OAuth consent screen"):
   - **Branding:** enter an app name and your email.
   - **Audience:** choose **External**, then click **Publish app** so the status is
     **In production**.

     *Why publish:* in "Testing" status, Google expires your sign-in every 7 days. Publishing
     doesn't make the app public. You'll just see a "Google hasn't verified this app" warning
     at sign-in. Click **Advanced → Go to (app name)** to continue, since it's your own app.
4. Go to **Clients → Create client → Application type: Desktop app**, then click
   **Create → Download JSON**.
5. Save the downloaded file next to the script as **`client_secret.json`**.

### 2. Schedule the Takeout export

1. Go to <https://takeout.google.com/>, click **Deselect all**, then tick only **Google Photos**.
2. Click **Next step** and choose these options:
   - **Destination:** Add to Drive.
   - **Frequency:** Export every 2 months for 1 year. You can also export once.
   - **File type:** .zip.
   - **Size:** 50 GB, which means fewer parts.
3. Click **Create export**. Google emails you when it's ready, which can take hours or days.

Each export is a full copy of your library, and it takes up Google Drive storage. The script only
has read access, so it can't delete old exports. Delete them from the `Takeout` folder in Drive
after they've been imported. The script logs when an export is safe to delete.

### 3. Configure

Copy `config.example.json` to `config.json` and set `Destination`. Backslashes must be doubled in
JSON: `"\\\\nas\\photos\\Google Photos"` means `\\nas\photos\Google Photos`.

| Setting | Default | Meaning |
|---|---|---|
| `Destination` | *(required)* | Local folder or UNC path the photos are copied into. |
| `ClientSecretFile` | `client_secret.json` | The OAuth client JSON from step 1. Relative paths are relative to the script. |
| `StagingDirectory` | `%LOCALAPPDATA%\GooglePhotosSync\staging` | Where archive parts are downloaded. Needs free space for one part (up to 50 GB). |
| `StateDirectory` | `%LOCALAPPDATA%\GooglePhotosSync` | Saved sign-in, import history, hash cache and logs. |
| `ExportSelection` | `Latest` | `Latest` imports only the newest export, since each one is a full copy. `AllUnprocessed` imports every export not yet imported. |
| `ProductFolder` | *(auto)* | Name of the Photos folder inside Takeout. Only needed if detection fails, for example when an export includes other Google products. |
| `CollisionTag` | `gphotos-collision` | The text used when renaming a colliding file. |
| `SetFileTimes` | `true` | Set file dates to when the photo was taken. |
| `CopyJsonSidecars` | `false` | Also copy Takeout's `.json` metadata files. |
| `DeleteStagedArchives` | `true` | Delete each downloaded part once it has been imported without errors. |

## Running it

```bash
powershell -ExecutionPolicy Bypass -File .\GooglePhotosSync.ps1
```

Useful options:

- `-DryRun` shows what would be copied without writing anything.
- `-SourcePath <path>` imports Takeout archives you already downloaded (a `.zip`/`.tgz`, a folder
  of them, or an extracted `Takeout` folder). No Google sign-in is needed.
- `-Destination <path>` overrides the config.
- `-ReAuthenticate` signs in to Google again.
- `-Reprocess` imports an export again even if it was already imported.
- `-Verbose` shows every copied file. Every file is always recorded in the log.

Exit codes: `0` means OK, `1` means some files failed (see the log), `2` means a fatal error.

### Running on a schedule

In Task Scheduler, create a task that runs **as your user**. It has to be your user because the
saved Google sign-in is encrypted for that Windows user.

- **Program:** `powershell.exe`
- **Arguments:** `-NoProfile -ExecutionPolicy Bypass -File "D:\path\to\GooglePhotosSync.ps1"`
- **Trigger:** weekly is plenty, since Takeout exports arrive every 2 months.

### Network (SMB) destinations

The script uses whatever access your Windows user already has to the share. If the NAS needs a
different login, open the share once in Explorer and tick "Remember my credentials", or run
`cmdkey /add:nas /user:NAS\you /pass`.

## Things to know

- **Album folders contain duplicates.** Takeout puts a photo both in `Photos from YYYY` and in
  each album it belongs to. The script mirrors that layout, so albums use extra space.
- **Edited photos:** Takeout exports both `IMG_1234.jpg` and `IMG_1234-edited.jpg`, and both are
  kept.
- **Live Photos** arrive as a still image plus a separate `.MP4`. Both are kept, with the same date.
- **EXIF isn't modified.** Only the file's created and modified dates are set.
- **Very long paths** (over 260 characters) can fail on Windows PowerShell 5.1. Failures are
  logged, and the run continues.
- The saved Google sign-in is encrypted with Windows DPAPI in `StateDirectory\refresh-token.dat`.
  To revoke access, go to <https://myaccount.google.com/permissions>.
