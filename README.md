# Google Photos Sync (Takeout helper)

A PowerShell script that does the tedious part of Google Takeout for you. Normally you'd download
many large archives by hand, unzip them and move the photos into the right folders. This script
fetches your Google Photos Takeout exports from Google Drive and copies the photos into a local or
network (SMB) folder, keeping Takeout's folder layout.

It needs nothing beyond what ships with Windows 10/11: Windows PowerShell 5.1 or PowerShell 7.

> **Why Takeout?** Since March 2025, Google's Photos API no longer lets third-party tools read a
> whole Google Photos library. Google Takeout is the practical way to get a full copy, and this
> script automates everything around it.

## Why use this instead of doing it by hand?

You still have to create the Takeout export yourself, so that step is the same either way. The
script takes over the work after that:

| Manual job | What the script does |
|---|---|
| Download each archive part in a browser. For a big library that's dozens of 50 GB files, and Takeout download links expire after 7 days. | Pulls the parts straight from Drive. Downloads resume if interrupted and are checked against Google's checksum. |
| Unzip every part and merge the folders. A photo's `.json` metadata file can land in a different part from the photo. | Reads the zips directly and merges all parts into one folder layout, with no manual unzipping. Metadata that sits in a different part from its photo is still matched. |
| File dates: Takeout stamps files with the export date, not when the photo was taken, so everything sorts wrongly. Fixing this by hand means reading the `.json` files, which isn't realistic. | Sets each file's date from the "photo taken" time in its `.json` file. |
| Merging folders in Explorer asks you to overwrite or skip each clash. | Skips files that are identical, and keeps both versions of different files, renaming the new one `(gphotos-collision-N)`. Nothing is overwritten. |
| Finding out how many photos there are and how much disk space you need. Takeout only tells you the archive size, and you find out the unzipped size by downloading and unzipping everything. | `-Review` reads each archive's table of contents without downloading it, and reports the photo and video counts, the unzipped size, and the space needed against what's free. |
| Choosing where files end up: browsers save to Downloads, and you unzip and move things by hand. | Unzips straight into one folder you choose with `-ExtractTo` (default: `Takeout` in the current directory), including network (SMB) folders. |
| Downloading everything first and dealing with it later. | `-DownloadOnly` fetches and verifies the archives and stops. Unzip them later with `-SourcePath`. |
| Watching progress, or guessing whether it's stuck. | Progress bars for downloading and unzipping show the amount done, speed and time left. When output is captured, a line is logged every 10%. |
| Repeating all this for each new export. | Imports only exports it hasn't seen, so re-running is safe. It can run on a schedule, and it works with network (SMB) folders. |

<!-- MAINTAINER NOTE: when you add, change or remove a script feature, update this table so it
     still shows how the script differs from doing the same job by hand. Add a row for a new
     capability, edit a row if behaviour changes, and update "Limits to be aware of" below if a
     limitation is lifted or a new one appears, including the "Tested so far" entry when more is
     tested. -->

**Limits to be aware of:**

- **It can't read your library directly.** Since March 2025, Google's Photos API only lets an app
  see photos that the app itself uploaded. Takeout is the only practical way to get a whole
  library, so the script depends on it.
- **New photos arrive only with a new export.** The script can't see photos added to Google Photos
  since the last export. Takeout can schedule an export at most every 2 months, so your folder can
  be up to 2 months behind. You can also create an export by hand at any time.
- **Each export is a full copy of the library, not just the new photos.** Every new export is
  downloaded in full (identical files are then skipped when copying). For a large library that
  means a long download every time.
- **One-way only (Google to folder).** There's no upload back to Google Photos and no true
  two-way sync. The script never deletes anything, so a photo you delete in Google Photos stays in
  your folder.
- **Not covered:** creating the Takeout export, and deleting old exports from Drive. The script
  has read-only access to Drive, so it can't remove them. Delete them yourself in Drive's
  `Takeout` folder.
- **Drive storage:** the "Add to Drive" delivery option uses your Drive space until you delete the
  export. Downloading straight from Takeout in a browser doesn't. That's the cost of automating
  the download.
- **Only file dates are restored.** Takeout's metadata (descriptions, locations, people, album
  membership) isn't written into the photos. The `.json` files are discarded unless you set
  `CopyJsonSidecars`.
- **Small libraries:** for a few dozen photos, doing it by hand takes about as long, and the
  one-time Google Cloud setup below is more work than the manual job. The script pays off with big
  libraries, many archive parts, or repeated exports.
- **Windows only.** The saved sign-in is encrypted with Windows DPAPI, so the script needs
  Windows PowerShell 5.1 or PowerShell 7 on Windows.
- **`-Review` can't read `.tgz` exports.** It leaves them out of the report. Use the `.zip` file
  type when creating the export.
- **Tested so far** on a small export (62 photos, one album) plus synthetic test archives, on both
  PowerShell versions. That covered review, download-only and unzipping, with the zip64 table of
  contents used by archives over 4 GB tested on hand-built data. It hasn't been run on a library of
  100 GB or more, on `.tgz` exports, or against a network (SMB) destination.

## What it does

1. Signs in to Google with **read-only** access to Google Drive. The first time, your browser
   opens for you to approve it. After that, runs are unattended.
2. Finds the Takeout archives (`takeout-*.zip`) that Takeout delivered to your Drive.
3. Downloads them one part at a time. Downloads resume if interrupted and are checked against
   Google's checksum.
4. Unzips the photos and videos into your destination folder (`Takeout` in the current directory by
   default), for example `Takeout\Photos from 2023\IMG_1234.JPG` or `Takeout\<Album name>\...`.
   Takeout's extra `Takeout\Google Photos\` folders are left out, and all archive parts are merged
   into the one folder.
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
   - **Audience:** choose **External**. Then do **one** of these. Without it, sign-in fails with
     `Error 403: access_denied` ("has not completed the Google verification process").
     - **Add yourself as a test user (quickest):** under **Test users**, click **Add users**,
       enter the Google email you'll sign in with, and save. The app stays in **Testing**
       status, so Google expires your sign-in every 7 days and you have to sign in again.
     - **Publish the app (recommended):** click **Publish app** so the status is
       **In production**. Sign-ins then no longer expire after 7 days. Publishing doesn't make
       the app public or need a review for personal use. You'll just see a "Google hasn't
       verified this app" warning at sign-in. Click **Advanced → Go to (app name)** to continue,
       since it's your own app.
4. Go to **Clients → Create client → Application type: Desktop app**, then click
   **Create → Download JSON**.
5. Save the downloaded file in the same folder as the script. The name Google gives it
   (`client_secret_<numbers>.apps.googleusercontent.com.json`) works as it is, or you can rename
   it to **`client_secret.json`**. Keep it private: it's excluded from git, so don't commit or
   share it.
6. Run `GooglePhotosSync.ps1 -SignIn`. Your browser opens, you sign in with the account from
   step 3 and approve read-only access to Google Drive, and the script confirms which account it
   is signed in as.

**If sign-in fails:**

| Message | Cause and fix |
|---|---|
| `Error 403: access_denied` ... "can only be accessed by developer-approved testers" | The app is in Testing status and your account isn't a test user. Add it under **Audience → Test users**, or publish the app (step 3). |
| `Google Drive API has not been used in project ...` | Enable the Drive API (step 2) and wait a minute. |
| `Google OAuth client file not found` | `client_secret*.json` isn't in the script's folder. Set `ClientSecretFile` in `config.json` if it's elsewhere. |
| `Several client_secret*.json files found` | Leave only one in the folder, or set `ClientSecretFile`. |
| Browser doesn't open | Copy the link the script prints into your browser. |

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

A config file is optional. Without one, the photos are unzipped into a folder named `Takeout` in the
folder you run the script from. To change that, either pass `-ExtractTo <folder>` on the command
line, or copy `config.example.json` to `config.json` and set `Destination`. Backslashes must be
doubled in JSON: `"\\\\nas\\photos\\Google Photos"` means `\\nas\photos\Google Photos`.

| Setting | Default | Meaning |
|---|---|---|
| `Destination` | `Takeout` in the current directory | The folder the photos are unzipped into. Local folder or UNC path. A relative path in `config.json` is relative to the script; one given with `-ExtractTo` is relative to the current directory. |
| `ClientSecretFile` | `client_secret.json` | The OAuth client JSON from step 1. Relative paths are relative to the script. |
| `StagingDirectory` | `%LOCALAPPDATA%\GooglePhotosSync\staging` | Where archive parts are downloaded. Needs free space for one part (up to 50 GB), or for all parts with `-DownloadOnly`. |
| `StateDirectory` | `%LOCALAPPDATA%\GooglePhotosSync` | Saved sign-in, import history, hash cache and logs. |
| `ExportSelection` | `Latest` | `Latest` imports only the newest export, since each one is a full copy. `AllUnprocessed` imports every export not yet imported. |
| `ProductFolder` | *(auto)* | Name of the Photos folder inside Takeout. Only needed if detection fails, for example when an export includes other Google products. |
| `CollisionTag` | `gphotos-collision` | The text used when renaming a colliding file. |
| `SetFileTimes` | `true` | Set file dates to when the photo was taken. |
| `CopyJsonSidecars` | `false` | Also copy Takeout's `.json` metadata files. |
| `DeleteStagedArchives` | `true` | Delete each downloaded part once it has been imported without errors. |
| `SignOutWhenDone` | `false` | Sign out of Google after every Drive run (see [Your Google sign-in](#your-google-sign-in)). |
| `DownloadOnly` | `false` | Download the archives and stop, without unzipping (see [Download only](#download-only)). |

## Running it

```bash
powershell -ExecutionPolicy Bypass -File .\GooglePhotosSync.ps1
```

Downloading and unzipping each show a progress bar with the amount done, speed and time left. When
the output is being captured instead of shown in a console, such as in a scheduled task, a line is
logged at every 10% instead.

Useful options:

- `-ExtractTo <folder>` (also `-Destination`) is the folder to unzip into. The default is a folder
  named `Takeout` in the current directory. It can be a network path such as `\\nas\photos`.
- `-Review` reports how many photos and videos there are and how much disk space is needed, and
  stops there (see below).
- `-DownloadOnly` downloads the archives and doesn't unzip them (see below).
- `-StagingDirectory <folder>` is where archives are downloaded to.
- `-DryRun` shows what would be copied without writing anything to the destination. It still
  downloads the archives; use `-Review` to avoid that.
- `-SourcePath <path>` imports Takeout archives you already downloaded (a `.zip`/`.tgz`, a folder
  of them, or an extracted `Takeout` folder). No Google sign-in is needed.
- `-SignIn` / `-SignOut` / `-SignOutWhenDone` / `-Unattended` manage the saved Google sign-in
  (see below).
- `-ReAuthenticate` signs in to Google again, for example to switch accounts.
- `-Reprocess` imports an export again even if it was already imported.
- `-Verbose` shows every copied file. Every file is always recorded in the log.

Exit codes: `0` means OK, `1` means some files failed (see the log), `2` means a fatal error.

### Reviewing size and disk space first

```bash
powershell -ExecutionPolicy Bypass -File .\GooglePhotosSync.ps1 -Review
```

This shows, for the export(s) a normal run would import:

- the number of archive parts and how much has to be downloaded,
- how many files there are, split into photos, videos and other, and how much space they take once
  unzipped,
- roughly how many *unique* photos that is (a photo that is also in an album is counted once),
- how much disk space is needed in the download folder and in the destination, against how much
  is free on each (this works for network folders too).

Nothing is downloaded or written. For exports in Google Drive, the script reads only the table of
contents at the end of each archive, so even a 50 GB part takes a few seconds. Add `-Reprocess` to
review an export that was already imported. `.tgz` archives can't be reviewed this way and are
left out of the report. You can also review archives you already have, with `-SourcePath`.

The numbers are for the whole export. Files already in the destination are skipped, so a repeat
import needs less space.

### Download only

```bash
powershell -ExecutionPolicy Bypass -File .\GooglePhotosSync.ps1 -DownloadOnly
```

Downloads the archives and stops. Nothing is unzipped, and the export isn't marked as imported.
The archives stay in the staging folder (`-StagingDirectory`), and the script prints where. Use
this to fetch everything first, for example overnight or onto a large drive, and unzip later:

```bash
powershell -ExecutionPolicy Bypass -File .\GooglePhotosSync.ps1 -SourcePath "<staging folder>" -ExtractTo D:\Photos
```

A normal run afterwards, with the same staging folder, finds the downloaded archives and uses them
instead of downloading again. Download only is off by default.

### Your Google sign-in

By default the script **keeps you signed in**. After you approve access in the browser once, it
saves Google's long-lived authorisation, encrypted with Windows DPAPI so only your Windows user on
this computer can use it. Later runs reuse it without opening a browser, and each run logs which
Google account it's using.

| Command | What it does |
|---|---|
| `-SignIn` | Signs in and saves the authorisation without importing anything. If you're already signed in, it just shows the account. Do this once before scheduling. |
| `-SignIn -ReAuthenticate` | Signs in again, for example with a different Google account. |
| `-SignOut` | Revokes the script's access at Google and deletes the saved authorisation. |
| `-SignOutWhenDone` | Runs the backup, then signs out, even if the run failed part-way, so no credentials stay on the PC. The next run asks you to sign in again. To make this permanent, set `"SignOutWhenDone": true` in `config.json`. |
| `-Unattended` | Never opens a browser. If the saved sign-in is missing or has expired, the run stops immediately with exit code 2, instead of waiting for someone to sign in. |

The saved sign-in lasts until you sign out or revoke access at
<https://myaccount.google.com/permissions>. The exception is an OAuth app left in **Testing**
status, whose sign-ins expire after 7 days. See setup step 1.

`SignOutWhenDone` and scheduled `-Unattended` runs don't mix: once the script signs out, the next
unattended run can't sign back in. Use `SignOutWhenDone` for manual runs.

### Running on a schedule

1. Run `GooglePhotosSync.ps1 -SignIn` once, by hand.
2. In Task Scheduler, create a task that runs **as the same Windows user**. It has to be the same
   user because the saved Google sign-in is encrypted for that user.

- **Program:** `powershell.exe`
- **Arguments:** `-NoProfile -ExecutionPolicy Bypass -File "D:\path\to\GooglePhotosSync.ps1" -Unattended`
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
- The saved Google sign-in is stored in `StateDirectory\refresh-token.dat`. It can't be copied to
  another PC or Windows user; sign in there separately.
