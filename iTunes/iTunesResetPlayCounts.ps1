# This script shows how to use the Reset-ITunesPlayCount cmdlet.
# It resets the play count for tracks in several playlists,
# which have not been played for more than one year,
# and have a play count greater than 5, 4, 2, or 1.

Import-Module iTunes
Reset-ITunesPlayCount -PlayListName 'Beschwingt s�mtliche' -PlayedDateOnOrBefore (Get-Date).AddYears(-1) -PlayedCountGreaterThan 5
Reset-ITunesPlayCount -PlayListName 'Vormittag s�mtliche' -PlayedDateOnOrBefore (Get-Date).AddYears(-1) -PlayedCountGreaterThan 4
Reset-ITunesPlayCount -PlayListName 'Nachmittag s�mtliche' -PlayedDateOnOrBefore (Get-Date).AddYears(-1) -PlayedCountGreaterThan 2
Reset-ITunesPlayCount -PlayListName 'Abend s�mtliche' -PlayedDateOnOrBefore (Get-Date).AddYears(-1) -PlayedCountGreaterThan 1

