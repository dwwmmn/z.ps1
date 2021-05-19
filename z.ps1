<#

Maintains a jump-list of the directories you actually use

INSTALL:
    * put something like this in your $profile:
        . /path/to/z.ps1
    * cd around for a while to build up the db
    * PROFIT!!
    * optionally:
        set $env:_Z_CMD in $profile to change the command (default z).
        set $env:_Z_DATA in $profile to change the datafile (default ~/.z).
        set $env:_Z_NO_RESOLVE_SYMLINKS to prevent symlink resolution.
        set $env:_Z_NO_PROMPT_COMMAND if you're handling PROMPT_COMMAND yourself.
        set $env:_Z_EXCLUDE_DIRS to an array of directories to exclude.
        set $env:_Z_OWNER to your username if you want use z while sudo with $HOME kept

USE:
    * z foo               # cd to most frecent dir matching foo
    * z foo bar           # cd to most frecent dir matching foo and bar
    * z -Rank foo         # cd to highest ranked dir matching foo
    * z -Recent foo       # cd to most recently accessed dir matching foo
    * z -List foo         # list matches instead of cd
    * z -OnlySubdirs foo  # restrict matches to subdirs of $PWD
#>

function ZFrecent {
    param($record, $now)

    $delta = $now - $time;
    if ($delta -lt 3600) { $record.rank * 4 }
    elseif ($delta -lt 86400) { $record.rank * 2 }
    elseif ($delta -lt 604800) { $record.rank / 2 }
    else { $record.rank / 4 }
}

# Find the common root of a list of matches, if it exists.
function ZCommon {
    param($matches);

    $SortedKeys = (Sort-Object -InputObject $matches.Keys -Property length);
    $Short = $SortedKeys | Select-Object -First 1

    if ($Short -match "[A-Z]:$") { return $null; } # $Short is a drive letter
    foreach ($x in $matches.Keys) {
        if (($null -ne $x) -and ($x.indexOf($Short) -ne 0)) {
            return $null;
        }
    }

    return $Short;
}

function ZOutput {
    param( $_matches, $bestMatch, $common );
    if ($null -ne $common) { $bestMatch = $common; }
    return $bestMatch;

}

function Get-EpochTimeNow {
    $date1 = (Get-Date -Date "01/01/1970").ToUniversalTime()
    $date2 = (Get-Date).ToUniversalTime()
    (New-TimeSpan -Start $date1 -End $date2).TotalSeconds
}

function Add-ToZDatabase {
    $ZDataFile = if ($env:_Z_DATA) { $env:_Z_DATA } else { "$env:USERPROFILE\.z" };
    $Add = $args[0];

    if (-not (Test-Path -Path $Add -IsValid)) {
        throw "Invalid path specified"
    }
    $Add = (Resolve-Path $Add.TrimEnd("\"));
    if ($Add -match "^[A-Z]:$") {
        # NOTE: Fixes weird behavior where `Resolve-Path C:` =>
        # C:\Users\Username instead of C:\. Not sure why it does this.
        #
        # NOTE 2: Apparently the default behavior for 'Set-Location C:' is to cd
        # to C:\Users\Username, or whatever the 'home' folder is for that drive.
        # This is really stupid and unintuitive IMHO so I'm leaving this code
        # here out of spite even though it's a NO-OP.
        $Add += "\";
    }

    # Check so we don't match $HOME. The logic here (I think) is that it's
    # already easy to jump to home with ~, we don't need another shortcut.
    if (($Add -eq $env:HOME) -or ($Add -eq $env:USERPROFILE)) {
        Set-Location -Path $Add; return;
    }

    # Don't track excluded directories
    if (($null -ne $env:_Z_EXCLUDE_DIRS) -and ($env:_Z_EXCLUDE_DIRS -contains $Add)) {
        Set-Location -Path $Add;
        return;
    }

    # Maintain the data file
    $TmpFile = New-TemporaryFile;
    try {
        $Data = Import-Csv -Path $ZDataFile;
    } catch [FileNotFoundException] {
        Set-Location -Path $Add;
        throw "Couldn't read from $ZDataFile";
    }
    $Rank = 1;
    $Time = Get-EpochTimeNow;
    $Sum = 0;
    $Added = $false; $Appended = $false;
    foreach ($j in $Data) { 
        if ($j.path -eq $Add) {
            $j.rank = $Rank + $j.rank;
            $j.time = $Time;
            $Added = $true;
        }
        $Sum += $j.rank;
    }
    if (-not $Added) {
        $record = New-Object -TypeName PsObject -Property @{
            "Path" = $Add;
            "Rank" = $rank;
            "Time" = $Time
        }
        $record | Export-Csv -Path $TmpFile -NoTypeInformation;
        $Appended = $true;
    }
    if ($Sum -gt 9000) {
        $Data = foreach ($i in $Data) {
            $i.rank *= 0.99;
            $i
        };
        $Data = Where-Object -InputObject $Data { $_.rank -ge 1 }
    }

    if ($Data -and $Appended) { 
        $Data | Export-Csv -Path $TmpFile -Append -NoTypeInformation;
    }
    elseif ($Data) {
        $Data | Export-Csv -Path $TmpFile -NoTypeInformation;
    }

    Get-Content $TmpFile | Set-Content $ZDataFile;
    Set-Location -Path $Add
}

function Set-ZLocation {
    param(
        [parameter(Mandatory = $false)][switch]$List,
        [parameter(Mandatory = $false)][switch]$Rank,
        [parameter(Mandatory = $false)][switch]$Recent,
        [parameter(Mandatory = $false)][switch]$OnlySubdirs,
        [String[]]$Patterns
    )
    $ZDataFile = if ($env:_Z_DATA) { $env:_Z_DATA } else { "$env:USERPROFILE\.z" };
    New-ZDataFile;

    $HighRank = $null;
    $IHighRank = $null;
    $BestMatch = $null;
    $IBestMatch = $null;
    $NormalMatches = New-Object System.Collections.Hashtable;
    $IMatches = New-Object System.Collections.Hashtable;
    $Data = Import-Csv -Path $ZDataFile;
    $QuotedPatterns = $Patterns | ForEach-Object { [Regex]::Escape($_) };
    $BuiltRegexString = $QuotedPatterns -join ".*"
    $Now = Get-EpochTimeNow;

    if ($Patterns[0] -eq "-") {
        Set-Location -
        return
    }

    if ($OnlySubdirs) {
        $BuiltRegexString = "$([Regex]::Escape((Get-Location))).*$BuiltRegexString"
    }

    foreach ($record in $Data) {
        $r = if ($Rank) { $record.rank }
        elseif ($Recent) { $Now - $record.time  } 
        else { ZFrecent $record $Now }

        if ($record.path -match $BuiltRegexString) {
            $NormalMatches[$record.path] = $r;

            if ($r -gt $HighRank) {
                $HighRank = $r
                $BestMatch = $record.path;
            }
        }
        elseif ($record.path -imatch $BuiltRegexString) {
            $IMatches[$record.path] = $r;

            if ($r -gt $IHighRank) {
                $IHighRank = $r;
                $IBestMatch = $record.path;
            }
        }
    }

    $ChosenMatch = $null; $ChosenMatches = $null
    if ($null -ne $BestMatch) {
        $ChosenMatch = $BestMatch; $ChosenMatches = $NormalMatches
    }
    elseif ($null -ne $IBestMatch) {
        $ChosenMatch = $IBestMatch; $ChosenMatches = $IMatches
    }

    if ($List) {
        $ChosenMatches | Format-Table;
        return
    }
    
    $cd = ZOutput $ChosenMatches $ChosenMatch (ZCommon $ChosenMatches);
    if ($null -ne $cd) {
        Set-Location -Path $cd;
        return
    }

    throw "Unable to find appropriate directory."
}

function New-ZDataFile {
    if (Test-Path -PathType Container "$env:USERPROFILE\.z") {
        throw "ERROR: z.ps1's datafile ($env:USERPROFILE\.z) is a directory!"
    } elseif (-not (Test-Path -PathType Leaf "$env:USERPROFILE\.z")) {
        New-Item -Path "$env:USERPROFILE\.z"
    }
}

function Set-ZAlias {
    $AliasName = if ($env:_Z_CMD) { $env:_Z_CMD } else { "z" }
    Set-Alias -Name $AliasName -Value Set-ZLocation -Option AllScope
    Set-Alias -Name "cd" -Value Add-ToZDatabase -Option AllScope
}

New-ZDataFile
Set-ZAlias