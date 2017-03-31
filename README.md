# iTunes
To use this module, copy the folder iTunes to a folder designated by the environment variable PSModulePath, e. g. WindowsPowerShell\Modules in your personal documents folder.

To use the commands, in a Windows PowerShell window, run
```
import-module iTunes
```

To list the available cmdlet, run
```
Get-Command -Module iTunes
```

To get help and examples for a cmdlet, run
```
Get-Help <cmdlet-name> -Full
```

## Note
The module was developed for the German version of iTunes. Unfortunately the 'kinds' of tracks in iTunes are not language neutral. To make the Convert-\* functions work, you need to customize the values in the string array $supportedKinds in line 3 of iTunes.psm1.

## Examples
Beside the examples provided in the help, the script **itunesSync.ps1** contains a more sophisticated example on how to convert several playlists at once, while renaming them at destination.

**iTunesResetPlayCounts.ps1** shows how to use the module to reset play counts for several playlists based on the last play date and play count.
