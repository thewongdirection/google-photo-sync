#Requires -Version 5.1
<#
.SYNOPSIS
    Automates getting your Google Photos out of Google Takeout and into a local or network (SMB) folder.

.DESCRIPTION
    A convenience tool around Google Takeout. Instead of downloading many Takeout archives by hand,
    unzipping them and moving the photos into place, this script:

      1. Signs in to Google (read-only Google Drive access) and finds the Takeout exports that
         Takeout delivered to your Drive (schedule them with "Add to Drive" - see README.md).
      2. Downloads each archive part (resuming if interrupted, verifying checksums).
      3. Copies the photos into the destination folder, mirroring Takeout's folder layout, and sets
         each file's date to when the photo was taken (from Takeout's .json metadata files).
      4. Remembers which exports it has imported, so scheduled runs only act on new ones.

    Background: Google's Photos API no longer lets third-party tools read a whole library (since
    March 2025), so Takeout is the practical route for a full copy. Copying is one-way
    (Google -> folder). Nothing is ever deleted from the destination or from Google.

    Name collisions: if a file with the same name already exists and has the same content, it is
    skipped. If the content differs, both are kept and the new one is saved as
    "name (gphotos-collision-N).ext".

.PARAMETER ConfigPath
    Path to the JSON config file. Defaults to config.json next to this script.

.PARAMETER SourcePath
    Import Takeout data you already have locally instead of downloading from Google Drive: .zip/.tgz
    archives, a folder containing them, or an extracted "Takeout" folder. No Google sign-in needed.

.PARAMETER Destination
    The folder to unzip the photos into (alias: -ExtractTo). Defaults to a folder named "Takeout" in
    the current directory. A relative path is relative to the current directory. Overrides
    Destination in the config file. Can be a network (UNC) path.

.PARAMETER Review
    Only report how many photos and videos the export(s) contain and how much disk space the
    download and the unzipped files need (alias: -Estimate). Downloads and writes nothing: for
    Google Drive exports it reads just the table of contents of each archive.

.PARAMETER DownloadOnly
    Download the archives from Google Drive and stop: nothing is unzipped or copied to the
    destination, and the export is not marked as imported. The archives are kept in the staging
    folder (see -StagingDirectory); unzip them later with -SourcePath. Off by default. Can also be
    set with "DownloadOnly": true in config.json.

.PARAMETER StagingDirectory
    Folder the archives are downloaded to. Overrides StagingDirectory in the config file.

.PARAMETER DryRun
    Report what would be copied without writing anything to the destination. Archives are still
    downloaded; use -Review to avoid that.

.PARAMETER ReAuthenticate
    Ignore the saved Google authorisation and sign in again.

.PARAMETER Reprocess
    Import the selected Drive export(s) again even if they were imported before.

.PARAMETER SignIn
    Only sign in to Google and save the authorisation (no import). If a saved sign-in still works,
    shows which account it is for. Combine with -ReAuthenticate to switch accounts.

.PARAMETER SignOut
    Revoke the saved Google authorisation and delete it from this computer.

.PARAMETER SignOutWhenDone
    After a Google Drive run finishes, revoke and delete the saved sign-in, so no Google credentials
    stay on this computer. The next run will ask you to sign in again in the browser.
    Can also be set permanently with "SignOutWhenDone": true in config.json.

.PARAMETER Unattended
    Never open a browser. If the saved sign-in is missing or no longer valid, fail immediately
    (exit code 2) instead of waiting for someone to sign in. Use this for scheduled tasks.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\GooglePhotosSync.ps1

.EXAMPLE
    .\GooglePhotosSync.ps1 -SignIn

.EXAMPLE
    .\GooglePhotosSync.ps1 -Review

.EXAMPLE
    .\GooglePhotosSync.ps1 -ExtractTo D:\Photos\Google

.EXAMPLE
    .\GooglePhotosSync.ps1 -DownloadOnly -StagingDirectory D:\TakeoutZips

.EXAMPLE
    .\GooglePhotosSync.ps1 -SourcePath D:\Downloads\takeout -Destination \\nas\photos\Google -DryRun
#>
[CmdletBinding()]
param(
    [string]$ConfigPath,
    [string[]]$SourcePath,
    [Alias('ExtractTo')]
    [string]$Destination,
    [string]$StagingDirectory,
    [Alias('Estimate')]
    [switch]$Review,
    [switch]$DownloadOnly,
    [switch]$DryRun,
    [switch]$ReAuthenticate,
    [switch]$Reprocess,
    [switch]$SignIn,
    [switch]$SignOut,
    [switch]$SignOutWhenDone,
    [switch]$Unattended
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem

if (-not $ConfigPath) { $ConfigPath = Join-Path $PSScriptRoot 'config.json' }

$script:OAuthScope     = 'https://www.googleapis.com/auth/drive.readonly'
$script:AuthEndpoint   = 'https://accounts.google.com/o/oauth2/v2/auth'
$script:TokenEndpoint  = 'https://oauth2.googleapis.com/token'
$script:RevokeEndpoint = 'https://oauth2.googleapis.com/revoke'
$script:DriveAboutUri  = 'https://www.googleapis.com/drive/v3/about?fields=user(emailAddress,displayName)'
$script:DriveFilesUri  = 'https://www.googleapis.com/drive/v3/files'
$script:InvalidChars   = [IO.Path]::GetInvalidFileNameChars()
$script:CopyBuffer     = New-Object byte[] (1MB)
$script:Stats          = @{ Copied = 0; CollisionRenamed = 0; SkippedIdentical = 0; Errors = 0; TimesSet = 0; TimesMissing = 0 }
$script:HashCache      = @{}
$script:Access         = $null
$script:ForceConsent   = [bool]$ReAuthenticate
$script:SignOutPending = $false
$script:ImpTracker     = $null   # progress for the archive currently being unzipped
$script:ImpBase        = 0L
$script:Downloaded     = @{ Parts = 0; Bytes = 0L }
$script:U32Max         = [long]4294967295
try { $script:ConsoleRedirected = [Console]::IsOutputRedirected } catch { $script:ConsoleRedirected = $true }
$script:LogWriter      = $null

#region Logging and helpers

function Write-Log {
    param([string]$Message, [ValidateSet('INFO', 'WARN', 'ERROR', 'DEBUG')][string]$Level = 'INFO')
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    if ($script:LogWriter) { $script:LogWriter.WriteLine($line) }
    switch ($Level) {
        'ERROR' { Write-Host $line -ForegroundColor Red }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'DEBUG' { Write-Verbose $line }
        default { Write-Host $line }
    }
}

function Get-HttpStatus($ErrorRecord) {
    $ex = $ErrorRecord.Exception
    while ($ex) {
        $prop = $ex.PSObject.Properties['Response']
        if ($prop -and $prop.Value) {
            try { return [int]$prop.Value.StatusCode } catch { }
        }
        $ex = $ex.InnerException
    }
    return 0
}

function Get-ErrorText($ErrorRecord) {
    if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) { return $ErrorRecord.ErrorDetails.Message }
    return $ErrorRecord.Exception.Message
}

function Resolve-ConfiguredPath([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return $Path }
    $expanded = [Environment]::ExpandEnvironmentVariables($Path)
    if ([IO.Path]::IsPathRooted($expanded)) { return $expanded }
    return [IO.Path]::GetFullPath((Join-Path $PSScriptRoot $expanded))
}

function Format-GB([long]$Bytes) {
    if ($Bytes -ge 1GB) { return '{0:N2} GB' -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return '{0:N1} MB' -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return '{0:N0} KB' -f ($Bytes / 1KB) }
    return "$Bytes bytes"
}

# Relative paths given on the command line are relative to the current directory.
function Resolve-CurrentDirPath([string]$Path) {
    $expanded = [Environment]::ExpandEnvironmentVariables($Path)
    return [IO.Path]::GetFullPath([IO.Path]::Combine((Get-Location).ProviderPath, $expanded))
}

#endregion

#region Progress bars

# Interactive console: a Write-Progress bar with speed and time left.
# Output redirected (scheduled task, log capture): a log line every 10% instead.
function New-ProgressTracker([string]$Activity, [long]$Total, [long]$Start = 0, [int]$Id = 1) {
    return @{ Activity = $Activity; Total = $Total; Start = $Start; Id = $Id; Watch = [Diagnostics.Stopwatch]::StartNew(); LastTick = -1000L; LastDecile = 0 }
}

function Update-ProgressTracker($T, [long]$Done) {
    $pct = 0
    if ($T.Total -gt 0) { $pct = [int][Math]::Min(100, [Math]::Floor($Done * 100 / $T.Total)) }
    $ms = $T.Watch.ElapsedMilliseconds
    if ($script:ConsoleRedirected) {
        $decile = [int][Math]::Floor($pct / 10)
        if ($decile -gt $T.LastDecile) {
            $T.LastDecile = $decile
            Write-Log ('{0}: {1}% ({2} of {3})' -f $T.Activity, $pct, (Format-GB $Done), (Format-GB $T.Total))
        }
        return
    }
    if ($ms - $T.LastTick -lt 250) { return }
    $T.LastTick = $ms
    $status = '{0} of {1} ({2}%)' -f (Format-GB $Done), (Format-GB $T.Total), $pct
    $moved = $Done - $T.Start
    if ($ms -gt 2000 -and $moved -gt 0) {
        $rate = $moved / ($ms / 1000.0)
        $status += ' - {0}/s' -f (Format-GB ([long]$rate))
        if ($T.Total -gt $Done) { $status += ', about {0} left' -f [TimeSpan]::FromSeconds([Math]::Min(359999, ($T.Total - $Done) / $rate)).ToString('hh\:mm\:ss') }
    }
    Write-Progress -Id $T.Id -Activity $T.Activity -Status $status -PercentComplete $pct
}

function Complete-ProgressTracker($T) {
    if ($script:ConsoleRedirected) {
        if ($T.LastDecile -lt 10 -and $T.Total -gt 0) { Write-Log ('{0}: 100% ({1})' -f $T.Activity, (Format-GB $T.Total)) }
        return
    }
    Write-Progress -Id $T.Id -Activity $T.Activity -Completed
}

#endregion

#region Config and state

function Read-Config {
    $cfg = [ordered]@{
        Destination          = $null
        ClientSecretFile     = 'client_secret.json'
        ClientId             = $null
        ClientSecret         = $null
        StateDirectory       = '%LOCALAPPDATA%\GooglePhotosSync'
        StagingDirectory     = '%LOCALAPPDATA%\GooglePhotosSync\staging'
        ExportSelection      = 'Latest'
        ArchiveNamePattern   = '^takeout-(\d{8}T\d{6}Z)(-\d+)*\.(zip|tgz)$'
        ProductFolder        = ''
        CollisionTag         = 'gphotos-collision'
        SetFileTimes         = $true
        CopyJsonSidecars     = $false
        DeleteStagedArchives = $true
        SignOutWhenDone      = $false
        DownloadOnly         = $false
    }
    # The config file is optional: without one the defaults apply (and -Destination can be given).
    if (Test-Path -LiteralPath $ConfigPath) {
        $json = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json
        foreach ($p in $json.PSObject.Properties) {
            if ($p.Name.StartsWith('_')) { continue }   # "_comment" style keys
            $cfg[$p.Name] = $p.Value
        }
    }
    if ($SignOutWhenDone) { $cfg.SignOutWhenDone = $true }
    if ($DownloadOnly) { $cfg.DownloadOnly = $true }
    if ($cfg.ExportSelection -notin 'Latest', 'AllUnprocessed') { throw "ExportSelection must be 'Latest' or 'AllUnprocessed'." }

    # Paths in the config file are relative to the script; paths from the command line are
    # relative to the current directory. With neither, unzip into .\Takeout in the current directory.
    foreach ($k in 'Destination', 'StateDirectory', 'StagingDirectory', 'ClientSecretFile') {
        $cfg[$k] = Resolve-ConfiguredPath $cfg[$k]
    }
    if ($Destination) { $cfg.Destination = Resolve-CurrentDirPath $Destination }
    elseif ([string]::IsNullOrWhiteSpace($cfg.Destination)) { $cfg.Destination = Join-Path (Get-Location).ProviderPath 'Takeout' }
    if ($StagingDirectory) { $cfg.StagingDirectory = Resolve-CurrentDirPath $StagingDirectory }
    if ($cfg.Destination.Length -gt 3) { $cfg.Destination = $cfg.Destination.TrimEnd('\') }
    return [pscustomobject]$cfg
}

function Get-StatePath { Join-Path $Config.StateDirectory 'state.json' }

function Read-State {
    $state = @{ ProcessedExports = @{} }
    $p = Get-StatePath
    if (Test-Path -LiteralPath $p) {
        $j = Get-Content -LiteralPath $p -Raw | ConvertFrom-Json
        if ($j.PSObject.Properties['ProcessedExports'] -and $j.ProcessedExports) {
            foreach ($prop in $j.ProcessedExports.PSObject.Properties) { $state.ProcessedExports[$prop.Name] = $prop.Value }
        }
    }
    return $state
}

function Save-State($State) {
    $State | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Get-StatePath) -Encoding UTF8
}

#endregion

#region Google sign-in (OAuth 2.0, loopback redirect + PKCE)

function ConvertTo-Base64Url([byte[]]$Bytes) {
    [Convert]::ToBase64String($Bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')
}

function New-RandomBytes([int]$Count) {
    $bytes = New-Object byte[] $Count
    $rng = [Security.Cryptography.RandomNumberGenerator]::Create()
    try { $rng.GetBytes($bytes) } finally { $rng.Dispose() }
    return , $bytes
}

function Get-ClientCredentials {
    if ($Config.ClientId -and $Config.ClientSecret) { return @{ Id = $Config.ClientId; Secret = $Config.ClientSecret } }
    $file = $Config.ClientSecretFile
    if (-not (Test-Path -LiteralPath $file)) {
        # Google's download is named "client_secret_<id>.apps.googleusercontent.com.json"; accept it as-is.
        $found = @(Get-ChildItem -LiteralPath (Split-Path -Parent $file) -Filter 'client_secret*.json' -File -ErrorAction SilentlyContinue)
        if ($found.Count -eq 1) { $file = $found[0].FullName }
        elseif ($found.Count -gt 1) { throw "Several client_secret*.json files found next to $file; set ClientSecretFile in config.json." }
    }
    if (-not (Test-Path -LiteralPath $file)) {
        throw "Google OAuth client file not found: $file. Create a Desktop OAuth client in Google Cloud and save its JSON there (see README.md)."
    }
    $j = Get-Content -LiteralPath $file -Raw | ConvertFrom-Json
    $node = $j
    if ($j.PSObject.Properties['installed']) { $node = $j.installed } elseif ($j.PSObject.Properties['web']) { $node = $j.web }
    if (-not $node.client_id -or -not $node.client_secret) { throw "$file does not contain client_id/client_secret." }
    return @{ Id = $node.client_id; Secret = $node.client_secret }
}

function Wait-OAuthRedirect($Listener, [TimeSpan]$Timeout) {
    $deadline = (Get-Date) + $Timeout
    while ((Get-Date) -lt $deadline) {
        if (-not $Listener.Pending()) { Start-Sleep -Milliseconds 200; continue }
        $client = $Listener.AcceptTcpClient()
        try {
            $stream = $client.GetStream()
            $stream.ReadTimeout = 5000
            $reader = New-Object IO.StreamReader -ArgumentList $stream, ([Text.Encoding]::ASCII), $false, 1024, $true
            $requestLine = $reader.ReadLine()
            while ($true) { $h = $reader.ReadLine(); if ([string]::IsNullOrEmpty($h)) { break } }

            $params = @{}
            if ($requestLine -match '^GET\s+(\S+)\s') {
                $target = $Matches[1]
                $q = $target.IndexOf('?')
                if ($q -ge 0) {
                    foreach ($pair in $target.Substring($q + 1).Split('&')) {
                        if (-not $pair) { continue }
                        $kv = $pair.Split([char[]]@('='), 2)
                        $value = ''
                        if ($kv.Count -gt 1) { $value = [Uri]::UnescapeDataString($kv[1].Replace('+', ' ')) }
                        $params[[Uri]::UnescapeDataString($kv[0])] = $value
                    }
                }
            }
            $done = $params.ContainsKey('code') -or $params.ContainsKey('error')
            $html = '<html><body style="font-family:sans-serif"><h3>Google Photos Sync</h3><p>Waiting for Google sign-in...</p></body></html>'
            if ($done) { $html = '<html><body style="font-family:sans-serif"><h3>Google Photos Sync</h3><p>Sign-in received. You can close this tab and return to PowerShell.</p></body></html>' }
            $body = [Text.Encoding]::UTF8.GetBytes($html)
            $head = [Text.Encoding]::ASCII.GetBytes("HTTP/1.1 200 OK`r`nContent-Type: text/html; charset=utf-8`r`nContent-Length: $($body.Length)`r`nConnection: close`r`n`r`n")
            $stream.Write($head, 0, $head.Length)
            $stream.Write($body, 0, $body.Length)
            $stream.Flush()
            if ($done) { return $params }
        } catch {
            Write-Log "Ignoring unexpected request on sign-in port: $($_.Exception.Message)" DEBUG
        } finally {
            $client.Close()
        }
    }
    throw 'Timed out waiting for Google sign-in (5 minutes).'
}

function Invoke-BrowserConsent($Client) {
    if ($Unattended -or -not [Environment]::UserInteractive) {
        throw ('Google sign-in is required but this is an unattended run, so no browser can be opened. ' +
               'Run "GooglePhotosSync.ps1 -SignIn" once as this Windows user. If this keeps happening every ~7 days, ' +
               'your OAuth app is still in "Testing" status; publish it (see README).')
    }
    $verifier = ConvertTo-Base64Url (New-RandomBytes 48)
    $sha = [Security.Cryptography.SHA256]::Create()
    try { $challenge = ConvertTo-Base64Url ($sha.ComputeHash([Text.Encoding]::ASCII.GetBytes($verifier))) } finally { $sha.Dispose() }
    $state = ConvertTo-Base64Url (New-RandomBytes 16)

    $listener = [Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback, 0)
    $listener.Start()
    try {
        $redirectUri = 'http://127.0.0.1:{0}/' -f $listener.LocalEndpoint.Port
        $query = [ordered]@{
            client_id             = $Client.Id
            redirect_uri          = $redirectUri
            response_type         = 'code'
            scope                 = $script:OAuthScope
            code_challenge        = $challenge
            code_challenge_method = 'S256'
            state                 = $state
            access_type           = 'offline'
            prompt                = 'consent'
        }
        $url = $script:AuthEndpoint + '?' + (($query.GetEnumerator() | ForEach-Object { '{0}={1}' -f $_.Key, [Uri]::EscapeDataString($_.Value) }) -join '&')
        Write-Log 'Opening your browser to sign in to Google and grant read-only access to Google Drive...'
        Write-Host "If the browser does not open, paste this link into it:`n$url`n"
        Start-Process $url
        $result = Wait-OAuthRedirect $listener ([TimeSpan]::FromMinutes(5))
    } finally {
        $listener.Stop()
    }

    if ($result['error']) { throw "Google sign-in failed: $($result['error'])" }
    if ($result['state'] -ne $state) { throw 'Google sign-in failed: state mismatch (possible stale or forged redirect). Try again.' }

    $tok = Invoke-RestMethod -Method Post -Uri $script:TokenEndpoint -Body @{
        code          = $result['code']
        client_id     = $Client.Id
        client_secret = $Client.Secret
        redirect_uri  = $redirectUri
        grant_type    = 'authorization_code'
        code_verifier = $verifier
    }
    if ($tok.scope -notmatch [regex]::Escape($script:OAuthScope)) {
        throw 'Google Drive read access was not granted. Run again with -ReAuthenticate and tick the Google Drive permission.'
    }
    if (-not $tok.refresh_token) { throw 'Google did not return a refresh token. Run again with -ReAuthenticate.' }
    return $tok
}

function Get-TokenFilePath { Join-Path $Config.StateDirectory 'refresh-token.dat' }

function Save-RefreshToken([string]$Token) {
    # Encrypted with Windows DPAPI: only this Windows user on this machine can decrypt it.
    $secure = ConvertTo-SecureString -String $Token -AsPlainText -Force
    ConvertFrom-SecureString -SecureString $secure | Set-Content -LiteralPath (Get-TokenFilePath) -Encoding ASCII
}

function Read-RefreshToken {
    $p = Get-TokenFilePath
    if (-not (Test-Path -LiteralPath $p)) { return $null }
    try {
        $secure = (Get-Content -LiteralPath $p -Raw).Trim() | ConvertTo-SecureString
        return (New-Object PSCredential 'x', $secure).GetNetworkCredential().Password
    } catch {
        Write-Log "Saved Google authorisation could not be read ($($_.Exception.Message)); you will need to sign in again." WARN
        return $null
    }
}

function Set-AccessToken($TokenResponse) {
    $script:Access = @{ Token = $TokenResponse.access_token; ExpiresAt = (Get-Date).AddSeconds([int]$TokenResponse.expires_in - 120) }
}

function Get-AccessToken {
    if ($script:Access -and (Get-Date) -lt $script:Access.ExpiresAt) { return $script:Access.Token }
    $refresh = $null
    if (-not $script:ForceConsent) { $refresh = Read-RefreshToken }
    if ($refresh) {
        try {
            $tok = Invoke-RestMethod -Method Post -Uri $script:TokenEndpoint -Body @{
                client_id     = $script:Client.Id
                client_secret = $script:Client.Secret
                refresh_token = $refresh
                grant_type    = 'refresh_token'
            }
            Set-AccessToken $tok
            return $script:Access.Token
        } catch {
            $detail = Get-ErrorText $_
            if ($detail -notmatch 'invalid_grant') { throw "Refreshing Google authorisation failed: $detail" }
            Write-Log 'Saved Google authorisation has expired or was revoked; signing in again.' WARN
        }
    }
    $tok = Invoke-BrowserConsent $script:Client
    Save-RefreshToken $tok.refresh_token
    $script:ForceConsent = $false
    Set-AccessToken $tok
    return $script:Access.Token
}

#endregion

#region Google Drive

function Invoke-GoogleApi([string]$Uri) {
    $attempt = 0
    while ($true) {
        $attempt++
        $bearer = Get-AccessToken   # outside the try: sign-in problems must not be retried
        try {
            return Invoke-RestMethod -Uri $Uri -Headers @{ Authorization = "Bearer $bearer" }
        } catch {
            $status = Get-HttpStatus $_
            $text = Get-ErrorText $_
            if ($status -eq 401 -and $attempt -le 2) { $script:Access = $null; continue }
            $retryable = $status -eq 0 -or $status -eq 429 -or $status -ge 500 -or ($status -eq 403 -and $text -match 'rateLimitExceeded')
            if ($retryable -and $attempt -le 5) {
                $wait = [Math]::Pow(2, $attempt)
                Write-Log "Google API request failed ($status), retrying in $wait s..." WARN
                Start-Sleep -Seconds $wait
                continue
            }
            throw "Google API request failed ($status): $text"
        }
    }
}

function Get-SignedInAccount {
    $about = Invoke-GoogleApi $script:DriveAboutUri
    $user = $about.user
    if ($user.displayName) { return '{0} <{1}>' -f $user.displayName, $user.emailAddress }
    return $user.emailAddress
}

function Invoke-SignIn {
    $script:Client = Get-ClientCredentials
    $account = Get-SignedInAccount
    Write-Log "Signed in to Google as $account. The sign-in is saved for this Windows user in $(Get-TokenFilePath)."
}

function Invoke-SignOut {
    $p = Get-TokenFilePath
    $refresh = Read-RefreshToken
    if (-not $refresh) {
        if (Test-Path -LiteralPath $p) { Remove-Item -LiteralPath $p -Force }
        Write-Log 'No saved Google sign-in on this computer.'
        return
    }
    try {
        Invoke-RestMethod -Method Post -Uri $script:RevokeEndpoint -Body @{ token = $refresh } | Out-Null
        Write-Log 'Google access revoked.'
    } catch {
        Write-Log "Could not revoke access with Google ($(Get-ErrorText $_)). You can remove it manually at https://myaccount.google.com/permissions" WARN
    }
    Remove-Item -LiteralPath $p -Force
    Write-Log 'Saved Google sign-in deleted from this computer.'
}

function Get-DriveTakeoutExports {
    $q = "name contains 'takeout-' and trashed = false and mimeType != 'application/vnd.google-apps.folder'"
    $fields = 'nextPageToken,files(id,name,size,md5Checksum,createdTime)'
    $files = New-Object Collections.Generic.List[object]
    $pageToken = $null
    do {
        $uri = '{0}?pageSize=1000&spaces=drive&fields={1}&q={2}' -f $script:DriveFilesUri, [Uri]::EscapeDataString($fields), [Uri]::EscapeDataString($q)
        if ($pageToken) { $uri += '&pageToken=' + [Uri]::EscapeDataString($pageToken) }
        $resp = Invoke-GoogleApi $uri
        foreach ($f in @($resp.files)) { if ($f) { $files.Add($f) } }
        $pageToken = $null
        if ($resp.PSObject.Properties['nextPageToken']) { $pageToken = $resp.nextPageToken }
    } while ($pageToken)

    $exports = @{}
    foreach ($f in $files) {
        if ($f.name -notmatch $Config.ArchiveNamePattern) { continue }
        $id = $Matches[1]
        if (-not $exports.ContainsKey($id)) { $exports[$id] = New-Object Collections.Generic.List[object] }
        $exports[$id].Add($f)
    }
    foreach ($key in ($exports.Keys | Sort-Object)) {
        $parts = @($exports[$key] | Sort-Object name)
        [pscustomobject]@{
            Id    = $key
            Files = $parts
            Size  = [long](($parts | ForEach-Object { [long]$_.size } | Measure-Object -Sum).Sum)
        }
    }
}

function Copy-StreamWithProgress($In, $Out, [long]$Done, [long]$Total, [string]$Activity) {
    $buf = $script:CopyBuffer
    $t = New-ProgressTracker $Activity $Total $Done 1
    while (($n = $In.Read($buf, 0, $buf.Length)) -gt 0) {
        $Out.Write($buf, 0, $n)
        $Done += $n
        Update-ProgressTracker $t $Done
    }
    Complete-ProgressTracker $t
}

function Save-DriveFile($File, [string]$TargetPath, [string]$Label = '') {
    $expected = [long]$File.size
    if (Test-Path -LiteralPath $TargetPath) {
        if ((Get-Item -LiteralPath $TargetPath).Length -eq $expected) { Write-Log "Already downloaded: $($File.name)"; return }
        Remove-Item -LiteralPath $TargetPath -Force
    }
    $part = "$TargetPath.part"
    $uri = '{0}/{1}?alt=media' -f $script:DriveFilesUri, $File.id
    $failures = 0
    while ($true) {
        $have = 0L
        if (Test-Path -LiteralPath $part) { $have = (Get-Item -LiteralPath $part).Length }
        if ($expected -gt 0 -and $have -ge $expected) { break }
        $bearer = Get-AccessToken
        try {
            $req = [Net.HttpWebRequest][Net.WebRequest]::Create($uri)
            $req.Headers['Authorization'] = "Bearer $bearer"
            $req.Timeout = 120000
            $req.ReadWriteTimeout = 300000
            if ($have -gt 0) { $req.AddRange($have) }
            $resp = $req.GetResponse()
            try {
                $append = $have -gt 0 -and [int]$resp.StatusCode -eq 206
                $mode = [IO.FileMode]::Create
                if ($append) { $mode = [IO.FileMode]::Append } else { $have = 0 }
                $in = $resp.GetResponseStream()
                $out = [IO.File]::Open($part, $mode, [IO.FileAccess]::Write)
                try { Copy-StreamWithProgress $in $out $have $expected "Downloading $($File.name)$Label" }
                finally { $out.Dispose(); $in.Dispose() }
            } finally { $resp.Close() }
            if ($expected -le 0) { break }
            $failures = 0
        } catch {
            $failures++
            if ((Get-HttpStatus $_) -eq 401) { $script:Access = $null }
            if ($failures -ge 8) { throw "Download of $($File.name) failed after $failures attempts: $($_.Exception.Message)" }
            $wait = [Math]::Min(60, [Math]::Pow(2, $failures))
            Write-Log "Download interrupted ($($_.Exception.Message)); resuming in $wait s (attempt $failures)..." WARN
            Start-Sleep -Seconds $wait
        }
    }

    $actual = (Get-Item -LiteralPath $part).Length
    if ($expected -gt 0 -and $actual -ne $expected) { throw "Downloaded size of $($File.name) is $actual bytes, expected $expected." }
    if ($File.PSObject.Properties['md5Checksum'] -and $File.md5Checksum) {
        Write-Log "Verifying checksum of $($File.name)..."
        $md5 = (Get-FileHash -LiteralPath $part -Algorithm MD5).Hash
        if ($md5 -ne $File.md5Checksum.ToUpperInvariant()) {
            Remove-Item -LiteralPath $part -Force
            throw "Checksum mismatch for $($File.name); the download was discarded. Re-run to try again."
        }
    }
    Move-Item -LiteralPath $part -Destination $TargetPath -Force
}

# Free bytes on the volume holding $Path (works for network shares too), or $null if unknown.
# $Path may not exist yet; the nearest existing parent folder is used.
function Get-FreeSpace([string]$Path) {
    $p = $Path
    while ($p -and -not (Test-Path -LiteralPath $p)) { $p = Split-Path -Parent $p }
    if (-not $p) { return $null }
    try {
        if (-not ('GooglePhotosSync.Native' -as [type])) {
            Add-Type -Namespace GooglePhotosSync -Name Native -MemberDefinition '[System.Runtime.InteropServices.DllImport("kernel32.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode, SetLastError = true)] public static extern bool GetDiskFreeSpaceEx(string dir, out ulong avail, out ulong total, out ulong free);'
        }
        [uint64]$avail = 0; [uint64]$total = 0; [uint64]$free = 0
        if ([GooglePhotosSync.Native]::GetDiskFreeSpaceEx($p, [ref]$avail, [ref]$total, [ref]$free)) { return [long]$avail }
    } catch { }
    return $null
}

function Assert-FreeSpace([string]$Directory, [long]$Needed) {
    $free = Get-FreeSpace $Directory
    if ($null -ne $free -and $free -lt $Needed + 1GB) {
        throw ('Not enough free space in {0}: need {1}, have {2}. Change StagingDirectory in config.json or free up space.' -f $Directory, (Format-GB $Needed), (Format-GB $free))
    }
}

#endregion

#region Review: how many photos and how much disk space

$script:MediaKind = @{}
foreach ($e in 'jpg', 'jpeg', 'jfif', 'png', 'gif', 'heic', 'heif', 'webp', 'avif', 'tif', 'tiff', 'bmp', 'dng', 'raw', 'cr2', 'cr3', 'nef', 'arw', 'orf', 'rw2') { $script:MediaKind[".$e"] = 'Photos' }
foreach ($e in 'mp4', 'mov', 'avi', 'mkv', 'm4v', '3gp', 'wmv', 'mpg', 'mpeg', 'webm', 'mts', 'm2ts') { $script:MediaKind[".$e"] = 'Videos' }

# Reads $From..$To (inclusive) of a Drive file without downloading the rest of it.
function Read-DriveRange($File, [long]$From, [long]$To) {
    $uri = '{0}/{1}?alt=media' -f $script:DriveFilesUri, $File.id
    $failures = 0
    while ($true) {
        $bearer = Get-AccessToken
        try {
            $req = [Net.HttpWebRequest][Net.WebRequest]::Create($uri)
            $req.Headers['Authorization'] = "Bearer $bearer"
            $req.Timeout = 120000
            $req.ReadWriteTimeout = 300000
            $req.AddRange($From, $To)
            $resp = $req.GetResponse()
            try {
                if ([int]$resp.StatusCode -ne 206) { throw 'Google Drive did not honour the byte-range request.' }
                $ms = New-Object IO.MemoryStream
                $s = $resp.GetResponseStream()
                try { $s.CopyTo($ms) } finally { $s.Dispose() }
                return , $ms.ToArray()
            } finally { $resp.Close() }
        } catch {
            $failures++
            if ((Get-HttpStatus $_) -eq 401) { $script:Access = $null }
            if ($failures -ge 5 -or $_.Exception.Message -like '*did not honour*') { throw }
            Start-Sleep -Seconds ([Math]::Pow(2, $failures))
        }
    }
}

# Lists the files in a zip (path and unzipped size) from its table of contents, which sits at the end
# of the archive. $ReadRange is called as ($from, $to) and returns those bytes. Handles zip64.
function Read-ZipCentralDirectory([scriptblock]$ReadRange, [long]$Size) {
    $tailLen = [Math]::Min($Size, 131072L)
    [byte[]]$tail = & $ReadRange ($Size - $tailLen) ($Size - 1)
    $eocd = -1
    for ($i = $tail.Length - 22; $i -ge 0; $i--) {
        if ($tail[$i] -eq 0x50 -and $tail[$i + 1] -eq 0x4B -and $tail[$i + 2] -eq 5 -and $tail[$i + 3] -eq 6) { $eocd = $i; break }
    }
    if ($eocd -lt 0) { throw 'Not a zip archive (no end-of-directory record found).' }
    $total = [long][BitConverter]::ToUInt16($tail, $eocd + 10)
    $cdSize = [long][BitConverter]::ToUInt32($tail, $eocd + 12)
    $cdOffset = [long][BitConverter]::ToUInt32($tail, $eocd + 16)
    if ($total -eq 65535 -or $cdSize -eq $script:U32Max -or $cdOffset -eq $script:U32Max) {
        $loc = $eocd - 20
        if ($loc -lt 0 -or $tail[$loc] -ne 0x50 -or $tail[$loc + 1] -ne 0x4B -or $tail[$loc + 2] -ne 6 -or $tail[$loc + 3] -ne 7) {
            throw 'Zip64 end-of-directory locator not found.'
        }
        $recOffset = [long][BitConverter]::ToUInt64($tail, $loc + 8)
        [byte[]]$rec = & $ReadRange $recOffset ($recOffset + 55)
        if ($rec.Length -lt 56 -or $rec[2] -ne 6 -or $rec[3] -ne 6) { throw 'Zip64 end-of-directory record not found.' }
        $total = [long][BitConverter]::ToUInt64($rec, 32)
        $cdSize = [long][BitConverter]::ToUInt64($rec, 40)
        $cdOffset = [long][BitConverter]::ToUInt64($rec, 48)
    }
    if ($cdSize -gt 1GB) { throw 'Zip table of contents is unexpectedly large.' }
    [byte[]]$cd = @()
    if ($cdSize -gt 0) { $cd = & $ReadRange $cdOffset ($cdOffset + $cdSize - 1) }

    $entries = New-Object Collections.Generic.List[object]
    $utf8 = [Text.Encoding]::UTF8
    $p = 0
    for ($n = 0; $n -lt $total -and $p + 46 -le $cd.Length; $n++) {
        if ($cd[$p] -ne 0x50 -or $cd[$p + 1] -ne 0x4B -or $cd[$p + 2] -ne 1 -or $cd[$p + 3] -ne 2) { throw 'Corrupt zip table of contents.' }
        $comp = [long][BitConverter]::ToUInt32($cd, $p + 20)
        $uncomp = [long][BitConverter]::ToUInt32($cd, $p + 24)
        $nameLen = [int][BitConverter]::ToUInt16($cd, $p + 28)
        $extraLen = [int][BitConverter]::ToUInt16($cd, $p + 30)
        $commentLen = [int][BitConverter]::ToUInt16($cd, $p + 32)
        $localOffset = [long][BitConverter]::ToUInt32($cd, $p + 42)
        $name = $utf8.GetString($cd, $p + 46, $nameLen)
        if ($uncomp -eq $script:U32Max -or $comp -eq $script:U32Max -or $localOffset -eq $script:U32Max) {
            $e = $p + 46 + $nameLen
            $end = $e + $extraLen
            while ($e + 4 -le $end) {
                $id = [BitConverter]::ToUInt16($cd, $e)
                $len = [int][BitConverter]::ToUInt16($cd, $e + 2)
                if ($id -eq 1) {   # zip64 extra field: only the fields that were 0xFFFFFFFF, in this order
                    $q = $e + 4
                    if ($uncomp -eq $script:U32Max) { $uncomp = [long][BitConverter]::ToUInt64($cd, $q); $q += 8 }
                    break
                }
                $e += 4 + $len
            }
        }
        if (-not $name.EndsWith('/')) { $entries.Add([pscustomobject]@{ Path = $name.Replace('\', '/'); Length = $uncomp }) }
        $p += 46 + $nameLen + $extraLen + $commentLen
    }
    return , $entries
}

function Get-RemoteZipEntries($File) {
    $reader = { param($From, $To) Read-DriveRange $File $From $To }.GetNewClosure()
    return Read-ZipCentralDirectory $reader ([long]$File.size)
}

function New-ReviewAccumulator {
    return @{ Files = 0; Photos = 0; Videos = 0; Other = 0; Bytes = 0L; Seen = @{}; UniqueFiles = 0; Problems = (New-Object Collections.Generic.List[string]) }
}

# Counts the files from one archive part that the import would copy.
function Add-ReviewItems($Acc, $Items, [string]$PartName) {
    try { $product = Get-ProductFolder $Items } catch { $Acc.Problems.Add("${PartName}: $($_.Exception.Message)"); return }
    if (-not $product) { return }
    foreach ($it in $Items) {
        $rel = Get-MirrorPath $it.Path $product
        if (-not $rel) { continue }
        if ($rel.EndsWith('.json', [StringComparison]::OrdinalIgnoreCase) -and -not $Config.CopyJsonSidecars) { continue }
        $kind = $script:MediaKind[[IO.Path]::GetExtension($rel).ToLowerInvariant()]
        if (-not $kind) { $kind = 'Other' }
        $Acc[$kind]++
        $Acc.Files++
        $Acc.Bytes += $it.Length
        $key = $rel.Substring($rel.LastIndexOf('\') + 1).ToLowerInvariant() + '|' + $it.Length
        if (-not $Acc.Seen.ContainsKey($key)) { $Acc.Seen[$key] = $true; $Acc.UniqueFiles++ }
    }
}

function Write-Report([string]$Text = '') {
    Write-Host $Text
    if ($script:LogWriter) { $script:LogWriter.WriteLine($Text) }
}

function Get-VolumeKey([string]$Path) { return [IO.Path]::GetPathRoot($Path).ToUpperInvariant() }

function Write-SpaceLine([string]$Label, [string]$Path, [long]$Needed, $Free) {
    $verdict = 'free space unknown'
    $freeText = '?'
    if ($null -ne $Free) {
        $freeText = Format-GB $Free
        if ($Free -ge $Needed) { $verdict = 'enough' } else { $verdict = 'NOT ENOUGH' }
    }
    Write-Report ('  {0,-14} {1,10} needed, {2,10} free  [{3}]  {4}' -f $Label, (Format-GB $Needed), $freeText, $verdict, $Path)
}

function Show-ReviewReport($Reviews) {
    $stagingNeed = 0L
    $destNeed = 0L
    $downloadTotal = 0L
    $partSizes = New-Object Collections.Generic.List[long]
    Write-Report
    Write-Report 'Review of the Takeout export(s)'
    Write-Report '-------------------------------'
    foreach ($r in $Reviews) {
        $a = $r.Acc
        Write-Report ('{0}' -f $r.Name)
        if ($r.PartSizes.Count -gt 0) {
            $size = [long](($r.PartSizes | Measure-Object -Sum).Sum)
            $downloadTotal += $size
            foreach ($s in $r.PartSizes) { $partSizes.Add([long]$s) }
            Write-Report ('  Archive parts : {0} ({1} to download)' -f $r.PartSizes.Count, (Format-GB $size))
        }
        Write-Report ('  Files         : {0:N0} ({1:N0} photos, {2:N0} videos, {3:N0} other), {4} once unzipped' -f $a.Files, $a.Photos, $a.Videos, $a.Other, (Format-GB $a.Bytes))
        Write-Report ('  Unique files  : about {0:N0} (a photo that is also in an album is counted once)' -f $a.UniqueFiles)
        foreach ($problem in $a.Problems) { Write-Report "  Note          : $problem" }
        if ($a.Bytes -gt $destNeed) { $destNeed = $a.Bytes }
    }

    if ($Config.DownloadOnly -or -not $Config.DeleteStagedArchives) { $stagingNeed = $downloadTotal }
    elseif ($partSizes.Count -gt 0) { $stagingNeed = [long](($partSizes | Measure-Object -Maximum).Maximum) }
    if ($Config.DownloadOnly) { $destNeed = 0L }

    Write-Report
    Write-Report 'Disk space needed'
    $stagingFree = Get-FreeSpace $Config.StagingDirectory
    $destFree = Get-FreeSpace $Config.Destination
    $sameVolume = (Get-VolumeKey $Config.StagingDirectory) -eq (Get-VolumeKey $Config.Destination)
    if ($stagingNeed -gt 0) {
        $need = $stagingNeed
        if ($sameVolume) { $need += $destNeed }
        Write-SpaceLine 'Downloads' $Config.StagingDirectory $need $stagingFree
    }
    if ($destNeed -gt 0 -and -not ($stagingNeed -gt 0 -and $sameVolume)) {
        Write-SpaceLine 'Unzipped files' $Config.Destination $destNeed $destFree
    } elseif ($destNeed -gt 0) {
        Write-Report ('  Unzipped files: {0}, on the same drive as the downloads (counted above)  {1}' -f (Format-GB $destNeed), $Config.Destination)
    }
    if ($stagingNeed -le 0 -and $destNeed -le 0) { Write-Report '  Nothing to download or unzip.' }
    Write-Report
    if ($stagingNeed -gt 0) {
        if ($Config.DownloadOnly) { Write-Report '  Note on downloads: with -DownloadOnly every archive is kept, so all parts are counted.' }
        elseif ($Config.DeleteStagedArchives) { Write-Report '  Note on downloads: one part at a time; each is deleted once it has been unzipped, so the largest part is counted.' }
        else { Write-Report '  Note on downloads: archives are kept (DeleteStagedArchives is off), so all parts are counted.' }
    }
    if ($destNeed -gt 0) {
        Write-Report '  Note on unzipped files: every file is copied, including the album copies Takeout makes of photos. Files already'
        Write-Report '  in the destination (identical content) are skipped, so a repeat import needs less. If you import several'
        Write-Report '  different exports, the space needed can grow up to the sum of their sizes.'
    }
    Write-Report
}

#endregion

#region Takeout sources

function Open-TakeoutSource([string]$Path) {
    $items = New-Object Collections.Generic.List[object]
    if (Test-Path -LiteralPath $Path -PathType Container) {
        $dir = New-Object IO.DirectoryInfo $Path
        $root = $dir.FullName.TrimEnd('\')
        $prefix = ''
        if ($dir.Name -ieq 'Takeout') { $prefix = 'Takeout/' }
        foreach ($f in $dir.EnumerateFiles('*', [IO.SearchOption]::AllDirectories)) {
            $rel = $prefix + $f.FullName.Substring($root.Length + 1).Replace('\', '/')
            $items.Add([pscustomobject]@{ Path = $rel; Length = $f.Length; File = $f.FullName; Entry = $null })
        }
        return [pscustomobject]@{ Items = $items; Archive = $null; ExtractedDir = $null }
    }
    if ($Path -match '\.zip$') {
        $zip = [IO.Compression.ZipFile]::OpenRead($Path)
        foreach ($e in $zip.Entries) {
            if (-not $e.Name) { continue }   # directory entry
            $items.Add([pscustomobject]@{ Path = $e.FullName.Replace('\', '/'); Length = $e.Length; File = $null; Entry = $e })
        }
        return [pscustomobject]@{ Items = $items; Archive = $zip; ExtractedDir = $null }
    }
    if ($Path -match '\.(tgz|tar\.gz)$') {
        $dir = Join-Path $Config.StagingDirectory ('extract-' + [IO.Path]::GetFileNameWithoutExtension($Path))
        if (Test-Path -LiteralPath $dir) { Remove-Item -LiteralPath $dir -Recurse -Force }
        [void][IO.Directory]::CreateDirectory($dir)
        Write-Log "Extracting $([IO.Path]::GetFileName($Path)) with tar..."
        & tar.exe -xzf $Path -C $dir
        if ($LASTEXITCODE -ne 0) { throw "tar failed (exit code $LASTEXITCODE) extracting $Path" }
        $src = Open-TakeoutSource $dir
        $src.ExtractedDir = $dir
        return $src
    }
    throw "Unsupported source (expected .zip, .tgz or a folder): $Path"
}

function Close-TakeoutSource($Source) {
    if ($Source.Archive) { $Source.Archive.Dispose() }
    if ($Source.ExtractedDir -and (Test-Path -LiteralPath $Source.ExtractedDir)) { Remove-Item -LiteralPath $Source.ExtractedDir -Recurse -Force }
}

function Open-ItemStream($Item) {
    if ($Item.Entry) { return $Item.Entry.Open() }
    return [IO.File]::OpenRead($Item.File)
}

function Read-ItemText($Item) {
    $s = Open-ItemStream $Item
    try { return (New-Object IO.StreamReader -ArgumentList $s, ([Text.Encoding]::UTF8)).ReadToEnd() } finally { $s.Dispose() }
}

function Get-PathSegments([string]$ArchivePath) {
    $segs = @($ArchivePath.Split([char[]]@('/'), [StringSplitOptions]::RemoveEmptyEntries))
    $start = 0
    if ($segs.Count -gt 0 -and $segs[0] -ieq 'Takeout') { $start = 1 }
    return @{ Segments = $segs; Start = $start }
}

function Get-ProductFolder($Items) {
    if ($Config.ProductFolder) { return $Config.ProductFolder }
    $names = @{}
    foreach ($it in $Items) {
        $p = Get-PathSegments $it.Path
        if ($p.Segments.Count - $p.Start -ge 2) { $names[$p.Segments[$p.Start]] = $true }
    }
    $list = @($names.Keys)
    if ($list.Count -le 1) { return $list | Select-Object -First 1 }
    $photos = @($list | Where-Object { $_ -match 'Photo|Foto' })
    if ($photos.Count -eq 1) { return $photos[0] }
    if ($photos.Count -eq 0) { return $null }   # export has no Google Photos folder
    throw ("Cannot tell which Takeout folder holds Google Photos (found: {0}). Set ProductFolder in config.json." -f ($list -join ', '))
}

function ConvertTo-SafeName([string]$Name) {
    $s = $Name
    if ($s.IndexOfAny($script:InvalidChars) -ge 0) {
        $sb = New-Object Text.StringBuilder $s.Length
        foreach ($c in $s.ToCharArray()) {
            if ([Array]::IndexOf($script:InvalidChars, $c) -ge 0) { [void]$sb.Append('_') } else { [void]$sb.Append($c) }
        }
        $s = $sb.ToString()
    }
    $s = $s.TrimEnd(' ', '.')
    if (-not $s) { $s = '_' }
    if ($s -match '^(CON|PRN|AUX|NUL|COM\d|LPT\d)(\..*)?$') { $s = '_' + $s }
    return $s
}

# Path of an archive item relative to the destination (inside the product folder), or $null to skip it.
function Get-MirrorPath([string]$ArchivePath, [string]$Product) {
    $p = Get-PathSegments $ArchivePath
    $segs = $p.Segments
    $i = $p.Start
    if ($segs.Count - $i -lt 2 -or $segs[$i] -ine $Product) { return $null }
    $parts = for ($k = $i + 1; $k -lt $segs.Count; $k++) { ConvertTo-SafeName $segs[$k] }
    return (@($parts) -join '\')
}

function Split-MirrorPath([string]$Rel) {
    $i = $Rel.LastIndexOf('\')
    if ($i -lt 0) { return @('', $Rel) }
    return @($Rel.Substring(0, $i), $Rel.Substring($i + 1))
}

#endregion

#region Takeout metadata (.json sidecars -> "photo taken" time)

function Reset-Metadata {
    $script:Meta       = @{}   # "dir|media name"  -> unix seconds
    $script:MetaTitle  = @{}   # "dir|title"       -> unix seconds
    $script:MetaStem   = @{}   # "dir|name w/o ext" -> unix seconds (Live Photo .MP4 companions)
    $script:MetaPrefix = @{}   # dir -> list of truncated names
}

function Add-MetaKey($Table, [string]$Key, [long]$Ts) {
    if (-not $Table.ContainsKey($Key)) { $Table[$Key] = $Ts }
}

# "IMG_1.JPG.supplemental-metadata(1).json" / "IMG_1.JPG(1).json" -> "IMG_1(1).JPG"
function Get-MediaNameFromSidecar([string]$JsonName) {
    $base = $JsonName -replace '\.json$', ''
    $n = $null
    if ($base -match '^(.*)\((\d+)\)$') { $base = $Matches[1]; $n = $Matches[2] }
    if ($base -match '^(.+\..+)\.([^.]+)$' -and 'supplemental-metadata'.StartsWith($Matches[2], [StringComparison]::OrdinalIgnoreCase)) {
        $base = $Matches[1]
    }
    if ($n) {
        $ext = [IO.Path]::GetExtension($base)
        $base = $base.Substring(0, $base.Length - $ext.Length) + "($n)" + $ext
    }
    return $base
}

function Register-Sidecar([string]$MirrorRel, [string]$JsonText) {
    if ($JsonText -notmatch '"photoTakenTime"\s*:\s*\{[^}]*?"timestamp"\s*:\s*"?(\d+)') { return }
    $ts = [long]$Matches[1]
    if ($ts -le 0) { return }
    $dir, $jsonName = Split-MirrorPath $MirrorRel
    $media = Get-MediaNameFromSidecar $jsonName

    Add-MetaKey $script:Meta "$dir|$media" $ts
    Add-MetaKey $script:MetaStem ("$dir|" + [IO.Path]::GetFileNameWithoutExtension($media)) $ts
    if ($JsonText -match '"title"\s*:\s*"((?:[^"\\]|\\.)*)"') {
        $title = ConvertTo-SafeName ([regex]::Unescape($Matches[1]))
        Add-MetaKey $script:MetaTitle "$dir|$title" $ts
    }
    if ($jsonName.Length -ge 40) {
        # Takeout truncates long sidecar names; keep these for prefix matching.
        if (-not $script:MetaPrefix.ContainsKey($dir)) { $script:MetaPrefix[$dir] = New-Object Collections.Generic.List[object] }
        $script:MetaPrefix[$dir].Add(@($media, $ts))
    }
}

function Find-TakenTime([string]$Dir, [string]$Name) {
    $names = @($Name)
    $unedited = $Name -replace '-edited(?=(\(\d+\))?(\.[^.]*)?$)', ''
    if ($unedited -ne $Name) { $names += $unedited }
    foreach ($n in $names) {
        $k = "$Dir|$n"
        if ($script:Meta.ContainsKey($k)) { return $script:Meta[$k] }
        if ($script:MetaTitle.ContainsKey($k)) { return $script:MetaTitle[$k] }
    }
    $stemKey = "$Dir|" + [IO.Path]::GetFileNameWithoutExtension($unedited)
    if ($script:MetaStem.ContainsKey($stemKey)) { return $script:MetaStem[$stemKey] }
    if ($script:MetaPrefix.ContainsKey($Dir)) {
        $best = $null; $bestLen = 0
        foreach ($p in $script:MetaPrefix[$Dir]) {
            $prefix = $p[0]
            if ($prefix.Length -gt $bestLen -and $unedited.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) { $best = $p[1]; $bestLen = $prefix.Length }
        }
        if ($null -ne $best) { return $best }
    }
    return $null
}

#endregion

#region Destination writes, hashing, collisions

function Get-HashCachePath { Join-Path $Config.StateDirectory 'hash-cache.tsv' }

function Import-HashCache {
    $p = Get-HashCachePath
    if (-not (Test-Path -LiteralPath $p)) { return }
    foreach ($line in [IO.File]::ReadLines($p)) {
        $f = $line.Split("`t")
        if ($f.Count -eq 4) { $script:HashCache[$f[0]] = @([long]$f[1], [long]$f[2], $f[3]) }
    }
}

function Save-HashCache {
    $p = Get-HashCachePath
    $tmp = "$p.tmp"
    $w = New-Object IO.StreamWriter -ArgumentList $tmp, $false, (New-Object Text.UTF8Encoding $false)
    try {
        foreach ($k in $script:HashCache.Keys) {
            $v = $script:HashCache[$k]
            $w.WriteLine("$k`t$($v[0])`t$($v[1])`t$($v[2])")
        }
    } finally { $w.Dispose() }
    Move-Item -LiteralPath $tmp -Destination $p -Force
}

function Get-StreamHash($Stream) {
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return [BitConverter]::ToString($sha.ComputeHash($Stream)).Replace('-', '') } finally { $sha.Dispose() }
}

function Get-FileHashCached([string]$Path) {
    $fi = New-Object IO.FileInfo $Path
    $e = $script:HashCache[$Path]
    if ($e -and $e[0] -eq $fi.Length -and $e[1] -eq $fi.LastWriteTimeUtc.Ticks) { return $e[2] }
    $s = [IO.File]::OpenRead($Path)
    try { $h = Get-StreamHash $s } finally { $s.Dispose() }
    $script:HashCache[$Path] = @($fi.Length, $fi.LastWriteTimeUtc.Ticks, $h)
    return $h
}

function Get-ItemHash($Item) {
    $s = Open-ItemStream $Item
    try { return Get-StreamHash $s } finally { $s.Dispose() }
}

# Copies an item to $Path via a temp file; returns its SHA-256.
function Write-ImportedFile($Item, [string]$Path) {
    $tmp = "$Path.gpsync-tmp"
    try {
        $in = Open-ItemStream $Item
        try {
            $out = [IO.File]::Create($tmp)
            $sha = [Security.Cryptography.SHA256]::Create()
            try {
                $buf = $script:CopyBuffer
                $written = 0L
                while (($n = $in.Read($buf, 0, $buf.Length)) -gt 0) {
                    $out.Write($buf, 0, $n)
                    [void]$sha.TransformBlock($buf, 0, $n, $null, 0)
                    $written += $n
                    if ($script:ImpTracker) { Update-ProgressTracker $script:ImpTracker ($script:ImpBase + $written) }
                }
                [void]$sha.TransformFinalBlock($buf, 0, 0)
                $hash = [BitConverter]::ToString($sha.Hash).Replace('-', '')
            } finally { $out.Dispose(); $sha.Dispose() }
        } finally { $in.Dispose() }
        [IO.File]::Move($tmp, $Path)
    } catch {
        if ([IO.File]::Exists($tmp)) { [IO.File]::Delete($tmp) }
        throw
    }
    $fi = New-Object IO.FileInfo $Path
    $script:HashCache[$Path] = @($fi.Length, $fi.LastWriteTimeUtc.Ticks, $hash)
    return $hash
}

function Set-TakenTime([string]$Path, [long]$UnixSeconds, [string]$Hash) {
    $t = [DateTimeOffset]::FromUnixTimeSeconds($UnixSeconds).UtcDateTime
    [IO.File]::SetCreationTimeUtc($Path, $t)
    [IO.File]::SetLastWriteTimeUtc($Path, $t)
    $fi = New-Object IO.FileInfo $Path
    $script:HashCache[$Path] = @($fi.Length, $fi.LastWriteTimeUtc.Ticks, $Hash)
    $script:Stats.TimesSet++
}

function Set-ImportedFileTime([string]$Path, [string]$MirrorRel, [string]$Hash, $Pending) {
    if (-not $Config.SetFileTimes -or $DryRun) { return }
    $dir, $name = Split-MirrorPath $MirrorRel
    $ts = Find-TakenTime $dir $name
    if ($null -ne $ts) { Set-TakenTime $Path $ts $Hash }
    else { $Pending.Add([pscustomobject]@{ Path = $Path; Dir = $dir; Name = $name; Hash = $Hash }) }
}

function Complete-PendingTimes($Pending) {
    foreach ($p in $Pending) {
        $ts = Find-TakenTime $p.Dir $p.Name
        if ($null -ne $ts) { Set-TakenTime $p.Path $ts $p.Hash }
        else {
            $script:Stats.TimesMissing++
            Write-Log "No 'photo taken' date found for $($p.Dir)\$($p.Name); file date left as copy time." DEBUG
        }
    }
    $Pending.Clear()
}

function Import-MediaItem($Item, [string]$MirrorRel, $Pending) {
    $dest = [IO.Path]::Combine($Config.Destination, $MirrorRel)
    $dir = [IO.Path]::GetDirectoryName($dest)
    if (-not $DryRun -and -not [IO.Directory]::Exists($dir)) { [void][IO.Directory]::CreateDirectory($dir) }

    if (-not [IO.File]::Exists($dest)) {
        if (-not $DryRun) {
            $hash = Write-ImportedFile $Item $dest
            Set-ImportedFileTime $dest $MirrorRel $hash $Pending
        }
        return [pscustomobject]@{ Result = 'Copied'; Path = $dest }
    }

    # Same name exists: skip if any same-named copy has identical content, otherwise keep both.
    $leaf = [IO.Path]::GetFileName($dest)
    $ext = [IO.Path]::GetExtension($leaf)
    $stem = $leaf.Substring(0, $leaf.Length - $ext.Length)
    $tag = $Config.CollisionTag
    $pattern = '^' + [regex]::Escape("$stem ($tag-") + '(\d+)' + [regex]::Escape(")$ext") + '$'
    $candidates = New-Object Collections.Generic.List[string]
    $candidates.Add($dest)
    $maxN = 0
    foreach ($f in [IO.Directory]::EnumerateFiles($dir, "$stem ($tag-*)$ext")) {
        if ([IO.Path]::GetFileName($f) -match $pattern) {
            $candidates.Add($f)
            if ([int]$Matches[1] -gt $maxN) { $maxN = [int]$Matches[1] }
        }
    }
    $srcHash = $null
    foreach ($c in $candidates) {
        if ((New-Object IO.FileInfo $c).Length -ne $Item.Length) { continue }
        if (-not $srcHash) { $srcHash = Get-ItemHash $Item }
        if ((Get-FileHashCached $c) -eq $srcHash) { return [pscustomobject]@{ Result = 'SkippedIdentical'; Path = $c } }
    }

    $newPath = [IO.Path]::Combine($dir, "$stem ($tag-$($maxN + 1))$ext")
    if (-not $DryRun) {
        $hash = Write-ImportedFile $Item $newPath
        Set-ImportedFileTime $newPath $MirrorRel $hash $Pending
    }
    return [pscustomobject]@{ Result = 'CollisionRenamed'; Path = $newPath }
}

function Import-TakeoutPart([string]$Path, $Pending) {
    $partName = [IO.Path]::GetFileName($Path.TrimEnd('\'))
    $src = Open-TakeoutSource $Path
    try {
        $product = Get-ProductFolder $src.Items
        if (-not $product) { Write-Log "No Google Photos content found in $partName." WARN; return }

        $work = New-Object Collections.Generic.List[object]
        foreach ($it in $src.Items) {
            $rel = Get-MirrorPath $it.Path $product
            if (-not $rel) { continue }
            if ($rel.EndsWith('.json', [StringComparison]::OrdinalIgnoreCase)) {
                try { Register-Sidecar $rel (Read-ItemText $it) } catch { Write-Log "Could not read metadata $rel : $($_.Exception.Message)" WARN }
                if (-not $Config.CopyJsonSidecars) { continue }
            }
            $work.Add(@($it, $rel))
        }
        Write-Log ("{0}: {1} files from '{2}'" -f $partName, $work.Count, $product)

        $totalBytes = 0L
        foreach ($w in $work) { $totalBytes += $w[0].Length }
        $script:ImpBase = 0L
        $script:ImpTracker = New-ProgressTracker "Unzipping $partName" $totalBytes 0 2
        foreach ($w in $work) {
            $rel = $w[1]
            try {
                $r = Import-MediaItem $w[0] $rel $Pending
                $script:Stats[$r.Result]++
                switch ($r.Result) {
                    'Copied'           { Write-Log "Copied $rel" DEBUG }
                    'CollisionRenamed' { Write-Log "Name collision, kept both: $rel -> $([IO.Path]::GetFileName($r.Path))" }
                }
            } catch {
                $script:Stats.Errors++
                Write-Log "Failed to import $rel : $($_.Exception.Message)" ERROR
            }
            $script:ImpBase += $w[0].Length
            Update-ProgressTracker $script:ImpTracker $script:ImpBase
        }
        Complete-ProgressTracker $script:ImpTracker
    } finally {
        $script:ImpTracker = $null
        Close-TakeoutSource $src
    }
}

function Assert-Destination {
    $d = $Config.Destination
    if (-not [IO.Directory]::Exists($d)) {
        if ($DryRun) { Write-Log "Destination does not exist yet: $d" WARN; return }
        try { [void][IO.Directory]::CreateDirectory($d) }
        catch { throw "Cannot create or reach destination '$d': $($_.Exception.Message) For a network share, check the path and that this Windows user can write to it (open it in Explorer once to save credentials)." }
    }
    if (-not $DryRun) {
        $probe = Join-Path $d ('.gpsync-write-test-' + [guid]::NewGuid().ToString('N'))
        try { [IO.File]::WriteAllText($probe, 'test'); [IO.File]::Delete($probe) }
        catch { throw "Destination '$d' is not writable: $($_.Exception.Message)" }
    }
}

#endregion

#region Import modes

function Expand-SourcePaths($Paths) {
    foreach ($p in $Paths) {
        $full = (Resolve-Path -LiteralPath $p).ProviderPath
        if (Test-Path -LiteralPath $full -PathType Container) {
            $archives = @(Get-ChildItem -LiteralPath $full -File | Where-Object { $_.Name -match '\.(zip|tgz|tar\.gz)$' } | Sort-Object Name)
            if ($archives.Count -gt 0) { $archives | ForEach-Object { $_.FullName } } else { $full }
        } else {
            $full
        }
    }
}

function Invoke-LocalImport {
    $parts = @(Expand-SourcePaths $SourcePath)
    Reset-Metadata
    $pending = New-Object Collections.Generic.List[object]
    $n = 0
    foreach ($p in $parts) {
        $n++
        Write-Log "[$n/$($parts.Count)] Importing $p"
        Import-TakeoutPart $p $pending
    }
    Complete-PendingTimes $pending
}

# Signs in, lists the Takeout exports in Drive and returns the ones this run should handle.
function Get-ExportsToProcess($State) {
    $script:Client = Get-ClientCredentials
    Write-Log "Signed in to Google as $(Get-SignedInAccount)."
    Write-Log 'Looking for Google Takeout archives in Google Drive...'
    $exports = @(Get-DriveTakeoutExports)
    if ($exports.Count -eq 0) {
        Write-Log 'No Takeout archives (takeout-*.zip / .tgz) found in Google Drive. Create a Google Photos export with delivery "Add to Drive" - see README.md.' WARN
        return @()
    }
    foreach ($e in $exports) {
        $note = ''
        if ($State.ProcessedExports.ContainsKey($e.Id)) { $note = ' (already imported)' }
        Write-Log ('Found export {0}: {1} part(s), {2}{3}' -f $e.Id, $e.Files.Count, (Format-GB $e.Size), $note)
    }

    if ($Config.ExportSelection -eq 'Latest') {
        $toDo = @($exports[-1])
    } else {
        $toDo = @($exports)
    }
    if (-not $Reprocess) { $toDo = @($toDo | Where-Object { -not $State.ProcessedExports.ContainsKey($_.Id) }) }
    if ($toDo.Count -eq 0) { Write-Log 'Nothing new to import. (Use -Reprocess to include exports that were already imported.)' }
    return $toDo
}

function Invoke-LocalReview {
    $acc = New-ReviewAccumulator
    foreach ($p in @(Expand-SourcePaths $SourcePath)) {
        $name = [IO.Path]::GetFileName($p.TrimEnd('\'))
        if ($p -match '\.(tgz|tar\.gz)$') { $acc.Problems.Add("$name is a .tgz archive, which can't be reviewed without unpacking it; it was left out."); continue }
        Write-Log "Reading the contents of $name..."
        $src = Open-TakeoutSource $p
        try { Add-ReviewItems $acc $src.Items $name } finally { Close-TakeoutSource $src }
    }
    Show-ReviewReport @([pscustomobject]@{ Name = 'Local files'; Acc = $acc; PartSizes = @() })
}

function Invoke-DriveReview {
    $toDo = @(Get-ExportsToProcess (Read-State))
    if ($toDo.Count -eq 0) { return }
    $reviews = foreach ($export in $toDo) {
        $acc = New-ReviewAccumulator
        $n = 0
        foreach ($file in $export.Files) {
            $n++
            if ($file.name -match '\.(tgz|tar\.gz)$') { $acc.Problems.Add("$($file.name) is a .tgz archive, which can't be reviewed without downloading it; it was left out."); continue }
            Write-Log ('[{0}/{1}] Reading the contents of {2} ({3}) without downloading it...' -f $n, $export.Files.Count, $file.name, (Format-GB ([long]$file.size)))
            Add-ReviewItems $acc (Get-RemoteZipEntries $file) $file.name
        }
        [pscustomobject]@{ Name = "Export $($export.Id)"; Acc = $acc; PartSizes = @($export.Files | ForEach-Object { [long]$_.size }) }
    }
    Show-ReviewReport @($reviews)
}

function Invoke-DriveImport {
    $state = Read-State
    $toDo = @(Get-ExportsToProcess $state)
    if ($toDo.Count -eq 0) { return }

    [void][IO.Directory]::CreateDirectory($Config.StagingDirectory)
    foreach ($export in $toDo) {
        if ($Config.DownloadOnly) { Write-Log "Downloading export $($export.Id) (download only)..." } else { Write-Log "Importing export $($export.Id)..." }
        $errorsBefore = $script:Stats.Errors
        Reset-Metadata
        $pending = New-Object Collections.Generic.List[object]
        $n = 0
        foreach ($file in $export.Files) {
            $n++
            $local = Join-Path $Config.StagingDirectory $file.name
            $label = ' (part {0} of {1})' -f $n, $export.Files.Count
            if (-not (Test-Path -LiteralPath $local)) {
                $partial = 0L
                if (Test-Path -LiteralPath "$local.part") { $partial = (Get-Item -LiteralPath "$local.part").Length }
                Assert-FreeSpace $Config.StagingDirectory ([long]$file.size - $partial)
                Write-Log ('[{0}/{1}] Downloading {2} ({3})...' -f $n, $export.Files.Count, $file.name, (Format-GB ([long]$file.size)))
            }
            Save-DriveFile $file $local $label
            $script:Downloaded.Parts++
            $script:Downloaded.Bytes += [long]$file.size
            if ($Config.DownloadOnly) { continue }
            $partErrors = $script:Stats.Errors
            Import-TakeoutPart $local $pending
            if ($Config.DeleteStagedArchives -and $script:Stats.Errors -eq $partErrors) { Remove-Item -LiteralPath $local -Force }
        }
        if ($Config.DownloadOnly) { continue }
        Complete-PendingTimes $pending

        if ($DryRun) {
            Write-Log 'Dry run: export not marked as imported.'
        } elseif ($script:Stats.Errors -eq $errorsBefore) {
            $state.ProcessedExports[$export.Id] = [ordered]@{
                ImportedAt = (Get-Date).ToString('o')
                Files      = @($export.Files | ForEach-Object { $_.name })
            }
            Save-State $state
            Write-Log "Export $($export.Id) imported. You can delete its archives from Google Drive to free up storage."
        } else {
            Write-Log "Export $($export.Id) had errors; it will be retried on the next run." WARN
        }
    }
}

#endregion

#region Main

$exitCode = 0
$started = $false
try {
    $Config = Read-Config
    [void][IO.Directory]::CreateDirectory($Config.StateDirectory)
    $logDir = Join-Path $Config.StateDirectory 'logs'
    [void][IO.Directory]::CreateDirectory($logDir)
    $logPath = Join-Path $logDir ('gphotos-sync-{0}.log' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
    $script:LogWriter = New-Object IO.StreamWriter -ArgumentList $logPath, $true, (New-Object Text.UTF8Encoding $false)
    $script:LogWriter.AutoFlush = $true

    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

    if ($SignIn -and $SignOut) { throw 'Use either -SignIn or -SignOut, not both.' }
    if ($SignOut) { Invoke-SignOut; exit 0 }
    if ($SignIn) { Invoke-SignIn; exit 0 }
    if ($Review -and $Config.DownloadOnly) { throw 'Use either -Review or -DownloadOnly, not both.' }
    if ($Config.DownloadOnly -and $SourcePath) { throw '-DownloadOnly downloads from Google Drive, so it cannot be combined with -SourcePath.' }

    $mode = 'Google Drive'
    if ($SourcePath) { $mode = 'local files' }
    $what = "Destination: $($Config.Destination)"
    if ($Review) { $what = 'REVIEW ONLY - nothing will be downloaded or written' }
    elseif ($Config.DownloadOnly) { $what = "DOWNLOAD ONLY - archives go to $($Config.StagingDirectory)" }
    elseif ($DryRun) { $what += ' [DRY RUN - nothing will be written to it]' }
    Write-Log "Google Photos Sync starting. Source: $mode. $what"
    Write-Log "Log file: $logPath"

    if ($SourcePath -and $Review) {
        Invoke-LocalReview
    } elseif ($Review) {
        if ($Config.SignOutWhenDone) { $script:SignOutPending = $true }
        Invoke-DriveReview
    } elseif ($SourcePath) {
        Assert-Destination
        Import-HashCache
        $started = $true
        Invoke-LocalImport
    } else {
        if ($Config.SignOutWhenDone -and $Unattended) {
            Write-Log 'SignOutWhenDone is on, so the next unattended run will not be able to sign in. Run with -SignIn before it.' WARN
        }
        $script:SignOutPending = [bool]$Config.SignOutWhenDone
        if (-not $Config.DownloadOnly) {
            Assert-Destination
            Import-HashCache
            $started = $true
        }
        Invoke-DriveImport
        if ($Config.DownloadOnly -and $script:Downloaded.Parts -gt 0) {
            Write-Log ('Download complete: {0} archive part(s), {1}, saved in {2}' -f $script:Downloaded.Parts, (Format-GB $script:Downloaded.Bytes), $Config.StagingDirectory)
            Write-Log "Nothing was unzipped. To unzip them later: GooglePhotosSync.ps1 -SourcePath `"$($Config.StagingDirectory)`""
        }
    }
} catch {
    $exitCode = 2
    Write-Log $_.Exception.Message ERROR
    Write-Log $_.ScriptStackTrace DEBUG
} finally {
    if ($started) {
        try { Save-HashCache } catch { Write-Log "Could not save hash cache: $($_.Exception.Message)" WARN }
        $s = $script:Stats
        Write-Log ('Summary: {0} copied, {1} kept both (renamed), {2} already present, {3} errors. File dates set: {4}, no date found: {5}.' -f `
                $s.Copied, $s.CollisionRenamed, $s.SkippedIdentical, $s.Errors, $s.TimesSet, $s.TimesMissing)
    }
    if ($script:SignOutPending) {
        # Runs even if the import failed part-way, so credentials never stay behind when this option is on.
        try { Write-Log 'SignOutWhenDone: signing out of Google...'; Invoke-SignOut }
        catch { Write-Log "Sign-out failed: $($_.Exception.Message)" ERROR }
    }
    if ($script:LogWriter) { $script:LogWriter.Dispose() }
}
if ($exitCode -eq 0 -and $script:Stats.Errors -gt 0) { $exitCode = 1 }
exit $exitCode

#endregion
