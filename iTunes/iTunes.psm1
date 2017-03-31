# Customize: The following variable must contain the supported kinds as strings, other tracks will not be converted by Convert-ITunesTrack

$supportedKinds = 'AAC-Audiodatei', 'Abgeglichene AAC-Audiodatei', 'Gekaufte AAC-Audiodatei', 'Gekaufte MPEG-4 Videodatei', 'MPEG-4 Videodatei', 'MPEG-Audiodatei', 'MPEG-4-Audiodatei'
$ConvertiTunesPlaylistProgressionId = 1
$ConvertiTunesTrackProgressionId = 2
$hashtableFileName = "ConvertedFileList.xml"
$convertedTracks = @{}

<#
	.SYNOPSIS
	Gets an COM object of the iTunes application.
	
	.EXAMPLE
	$itunes = Get-ITunesApplication
#>
function Get-ITunesApplication {
	New-Object -ComObject iTunes.Application
}

<#
	.SYNOPSIS
	Changes the encoder in iTunes and returns the previous encoder.
	
	.DESCRIPTION
	Changes the encoder in iTunes and returns the previous encoder. The return value can be saved, so that Set-ITunesEncoder can be called with it later to restore the original encoder.
	
	.PARAMETER FormatName
	The format the new encoder should support. Examples are MP3 and AAC. Check iTunes documentation for supported formats. If no encoder supports the format, nothing is changed.
	
	.PARAMETER Encoder
	A COM object of the encoder to be set to.
	
	.EXAMPLE
	$originalEncoder = Set-ITunesEncoder 'MP3'

	.EXAMPLE
	Set-ITunesEncoder $originalEncoder
#>
function Set-ITunesEncoder {
	[CmdletBinding()]
	param(
		[Parameter(ParameterSetName='FormatName', Position=1)]
		[string]$Format = 'MP3',

		[Parameter(ParameterSetName='EncoderObject', Position=1, Mandatory = $true)]
		[System.__ComObject]$Encoder
	)

	$iTunesApplication = Get-ItunesApplication

	# save current encoder

	$originalEncoder = $iTunesApplication.CurrentEncoder

	if ($PSCmdlet.ParameterSetName -eq 'FormatName') {
		# find MP3 encoder

		$Encoder = $iTunesApplication.Encoders | Where-Object {$PSItem.Format -eq $Format}
	}


	# set encoder

	if ($Encoder -ne $null) {
		$iTunesApplication.CurrentEncoder = $Encoder
	}

	# return original encoder
	$originalEncoder
}

<#
	.SYNOPSIS
	Gets iTunes playlists by name.

	.DESCRIPTION
	Finds playlists matching the name in the iTunes library and returns the COM objecst. If multiple playlists have the same name, all are returned.

	.PARAMETER Name
	Names of the playlist to retrieve.

	.EXAMPLE
	This example retrieves all playlists with the name 'My Playlist' and stores them in the variable $playlist.

	$playlist = Get-ITunesPlaylist 'My Playlist'

	.EXAMPLE
	This example retrieves the playlists "List1" and "List2" and stores them in the variable $playlists.

	$playlists = @("List1", "List2") | Get-ITunesPlaylist
#>
function Get-ITunesPlaylist {
	[CmdletBinding()]
	param(
		[Parameter(Position=0, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true,  Mandatory=$true)]
		[string[]]$Name
	)

	begin {
		$iTunesApplication = Get-ItunesApplication
	}

	process {
		$Name | ForEach-Object {
			$iTunesApplication.LibrarySource.Playlists | Where-Object { $Name -contains $PSItem.Name }
		}
	}
}

<#
	.SYNOPSIS
	Copies converted files playlists to destination and creates an M3U file.

	.DESCRIPTION
	The files in the playlists are converted to the defined format (default: MP3). The converted files are copied to destination and removed from the iTunes library. An extended M3U file is built for each playlist at destination. 
	
	.PARAMETER PlaylistObject
	COM objects of iTunes playlists containing the files to convert.

	.PARAMETER Name
	Names of iTunes playlists containing the files to convert. If multiple playlists match a name, consider using the NewPlaylistFileName parameter, otherwise the M3U files of the playlists might overwrite each other. 

	.PARAMETER Destination
	A full path where the converted files and the M3U file are stored. A folder structure of artists and albums is automatically created.

	.PARAMETER Playlisttable
	Hashtable with the keys playlist or pl containing the name of a playlist or a playlist object (see Name and PlaylistObject parameters for more details), and newname or n for the filename of the playlist at destination.

	.PARAMETER Format
	The file format for the converted files. The default is MP3. See iTunes documentation for supported file formats.

	.NOTE
	The function tries to load ConvertedFileList.xml from the destination and match tracks. If a track is found in that file, they are not converted and copied again. They are only added to the M3U file. If the files must be converted again, delete ConvertedFileList.xml from destination.
	Converted tracks are written to ConvertedFileList.xml at destination containing ids and paths to all copied files. This file can be used to match tracks later.

	.EXAMPLE
	Convert-ITunesPlaylist "My Playlist" D:\ -NewPlaylistName "Playlist1" -Format 'MP3'
#>
function Convert-ITunesPlaylist {
	[CmdletBinding()]
	param(
		[Parameter(Position=0, ParameterSetName='PlaylistObject', ValueFromPipeline=$true, Mandatory=$true)]
		[System.__ComObject[]]$PlaylistObject,

		[Parameter(Position=0, ParameterSetName='PlaylistName', ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Mandatory=$true)]
		[string[]]$Name,

		[Parameter(Position=0, ParameterSetName='Hashtable', ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Mandatory=$true)]
		[System.Collections.Hashtable[]]$Playlisttable,


		[Parameter(Position=1, Mandatory=$true)]
		[string]$Destination,

		[string]$Format = 'MP3'
	)

	begin {

		$iTunesApplication = Get-ITunesApplication

		# Load hash table with converted files from file, if available, otherwise create an empty one

		$hashtableFile = "$Destination\$hashtableFileName"

		if (Test-Path -LiteralPath $hashtableFile) {
			$convertedTracks = Import-Clixml $hashtableFile
		}
		else {
			$convertedTracks = @{}
		}

		$index = 0
	}

	process {

		# Build hashtable for playlists

		switch ($PSCmdlet.ParameterSetName) {
			'Hashtable' {
				$playlistsToConvert = $Playlisttable | ForEach-Object {
					# Get playlist keys from hashtable
					$playlist = $PSItem['playlist']
					if ($playlist -eq $null) {
						$playlist = $Playlisttable['pl']
					}
				
					# If playlist is string, find playlist object
					if ($playlist -is [string]) {
						$playlist = Get-ITunesPlaylist -Name $playlist
					}

					# If a playlist was found, build a hashtable
					if ($playlist -is [System.__ComObject]) {
						

						# Add new name to the hashtable
						$newName = $PSItem['newname']
						if ($newName -eq $null) {
							$newName = $PSItem['n']
						}
						if ($newName -eq $null) {
							$newName = $playlist.Name
						}

						# return hash table
						@{playlist=$playlist; name=$newName}
					}
					
				}
				 
			}
			'PlaylistObject' {
				$playlistsToConvert = $PlaylistObject | ForEach-Object {
					@{playlist = $PSItem; name = $PSItem.Name}
				}
			}
			'PlaylistName' {
				$playlistsToConvert = $Name | ForEach-Object {
					$playlist = Get-ITunesPlaylist $PSItem
					if ($playlist -ne $null) {
						@{playlist=$playlist; name = $playlist.Name}
					}
				}
			}
		}

		# Process playlists

		foreach ($playListToConvert in $playlistsToConvert) {
	
			$playlist = $playListToConvert['playlist']

			# Start converting the playlist

			$activity = "Converting playlist $($playlist.Name)"
			$tracksCompleted = 0

			Write-Progress -Id $ConvertiTunesPlaylistProgressionId -Activity $activity

			# Build filename for playlist file

			$playlistFileName = "$($playlistToConvert['name']).m3u"
			$playlistFileName = Remove-InvalidFileNameChars $playlistFileName
			$playlistPath = "$Destination\$playlistFileName"

			Write-Verbose "Converting playlist $($playlist.Name) to $playlistPath"

			# Write header
			'#EXTM3U' | Out-File $playlistPath -Encoding default
				

			# Convert each track

			$playList.Tracks | ForEach-Object {
					
				# Update progress

				$percentComplete = '{0:N0}' -f ($tracksCompleted / $playlist.Tracks.Count * 100)
				Write-Progress -Id $ConvertiTunesPlaylistProgressionId -Activity $activity -CurrentOperation "adding $($PSItem.Name)" -Status "$percentComplete % complete" -PercentComplete $percentComplete

				# Get persistent id for track

				$id = Get-ITunesPersistentId -ITObject $PSItem

				# Find track, if it was already converted
				 
				$convertedTrack = $convertedTracks[$id]

				if ($convertedTrack -ne $null) {
					
					# If the file is not found, the track must be converted again
					if (-not (Test-Path -LiteralPath "$Destination\$convertedTrack" -PathType Leaf)) {
						$convertedTrack = $null
						$convertedTracks.Remove($id)
					}
				}

				if ($convertedTrack -eq $null) {

					# Convert track
					$convertedITunesTracks = Convert-ITunesTrack $PSItem -Format $Format

					# If track was converted, copy it to destination and delete it from iTunes
					if ($convertedITunesTracks -ne $null) {

						# The relative path of the file will be put into the pipeline
						$convertedTrack = $convertedITunesTracks | Copy-ITunesTrackFile -Destination $Destination
					   
						# Add track to hash table, so that we can find it later
						$convertedTracks.Add($id, $convertedTrack)

						# Write hash table to file, in case the script gets aborted
						$convertedTracks | Export-Clixml -LiteralPath $hashtableFile

						# Remove track from iTunes library and from file system
						Remove-ITunesTrack $convertedITunesTracks
					}
				}

				# If track was converted or found, add it to playlist file

				if ($convertedTrack -ne $null) {

					# Write extended info into playlist file
					"#EXTINF:$($PSItem.Duration),$($PSItem.Name)"

					# Write track path into playlist file
					$convertedTrack
				}

				$tracksCompleted++
			} | 
			Out-File $playlistPath -Encoding default -Append
			$index++
			Write-Progress -Id $ConvertiTunesPlaylistProgressionId -Activity $activity -Completed
		}
	}
}

<#
	.SYNOPSIS
	Gets the persistent ID of an iTunes object.
	.DESCRIPTION
	Get the persistent ID of an iTunes object. The persistent ID remains the same through different iTunes sessions.
	.PARAMETER ITObject
	An iTunes object such as a track or playlist for which the persistent ID needs to be retrieved.
	.EXAMPLE
	(Get-ITunesPlayList "My Playlist").Tracks | Get-ITunesPeristentId

#>
function Get-ITunesPersistentId {
	[CmdletBinding()]
	param(
		[Parameter(Position=0, ValueFromPipeline=$true, Mandatory=$true)]
		[System.__ComObject[]]$ITObject
	)

	begin {
			$iTunesApplication = Get-ItunesApplication
	}
	process {
		$ITObject | ForEach-Object {
			$idLow = [int64]$iTunesApplication.ITObjectPersistentIDLow($PSItem)
			$idHigh = [int64]$iTunesApplication.ITObjectPersistentIDHigh($PSItem)
			($idHigh -shl 32) -bor $idLow
		}
	}
}

<#
	.SYNOPSIS
	Converts iTunes tracks into the defined format.

	.DESCRIPTION
	Converts iTunes tracks into the defined format. The new tracks are added to the iTunes library. The output are the converted tracks as COM objects.

	.PARAMETER Track
	iTunes track COM objects of the tracks to convert.

	.PARAMETER Format
	The file format for the converted files. The default is MP3. See iTunes documentation for supported file formats.

	.EXAMPLE
	(Get-ITunesPlayList "My Playlist").Tracks | Convert-ITunesTrack -Format 'MP3'
#>
function Convert-ITunesTrack {
	[CmdletBinding()]
	param(
		[Parameter(Position=0, ValueFromPipeline=$true, Mandatory=$true)]
		[System.__ComObject[]]$Track,

		[string]$Format='MP3'
	)
	begin {
		$iTunesApplication = Get-ItunesApplication
		$originalEncoder = Set-ITunesEncoder $Format
	}
	process {
		$Track | ForEach-Object {
			Write-Verbose "$($PSItem.Name) is of kind $($PSItem.KindAsString)"

			# Check if the track file is present

			if ($PSItem.Location -ne $null) {
				if (Test-Path -LiteralPath $PSItem.Location -PathType Leaf) {
					
					# Check if track is of supported format

					if ($supportedKinds -contains $PSItem.KindAsString `
						-and (Test-Path -LiteralPath $PSItem.Location -PathType Leaf)) {

						# Start conversion
						$status = $iTunesApplication.ConvertTrack2($PSItem)

						# Wait for conversion to complete and update progress
						$activity = "Converting $($PSItem.Name)"
						while ($status.InProgress -eq $true) {
							$MaxProgressValue = $status.MaxProgressValue
							if ($MaxProgressValue -ne 0) {
								$percentComplete = '{0:N0}' -f ($status.ProgressValue / $MaxProgressValue * 100)
							} else {
								$percentComplete = 100
							}
							Write-Progress -id $ConvertiTunesTrackProgressionId -Activity $activity -Status "$percentComplete % complete" -PercentComplete $percentComplete
						}

						Write-Progress -id $ConvertiTunesTrackProgressionId -Activity $activity -Completed

						# Return converted tracks
	
						$status.Tracks
					}
					else {
						Write-Verbose "$($PSItem.Name) was not converted, because of an unsupported format."
					}
				}

				else {
					Write-Verbose "$($PSItem.Name) cannot be converted, because $($PSItem.Location) was not found."
				}
			}
			else {
				Write-Verbose "$($PSItem.Name) cannot be converted, there is no local location defined. Is this a cloud file?"
			}
			

		}
	}
	end {
		$temporaryEncoder = Set-ITunesEncoder -encoder $originalEncoder
	}
}

<#
	.SYNOPSIS
	Copies iTunes tracks to the destination, creating a folder structure.

	.DESCRIPTION
	Copies iTunes tracks to the destination. At the destination a folder structure is created. The first level contains either a folder named 'Compilations' for compilation albums, or the album artist name, If the album artist name is not available, the artist name is used. The second level contains the album name, if available. Output is the relative path of the destination file.

	.PARAMETER Track
	iTunes track COM objects of the tracks to convert.

	.PARAMETER Destination
	A full path where the converted files and the M3U file are stored. A folder structure of artists and albums is automatically created.
	
	.EXAMPLE
	(Get-ITunesPlayList "My Playlist").Tracks | Convert-ITunesTrack -Format 'MP3' | CopyITunesTrackFile D:\

#>
function Copy-ITunesTrackFile {
	[CmdletBinding()]
	param(
		[Parameter(Position=0, ValueFromPipeline=$true, Mandatory=$true)]
		[System.__ComObject[]]$Track, 

		[Parameter(Position=1, Mandatory=$true)]
		[string]$Destination
	)
	process {
		$Track | ForEach-Object {
			$location = $PSItem.Location

			Write-Verbose "Start copying $location"

			# Build destination path

			if ($PSItem.Compilation) {
				$artist = "Compilations"
			}
			if ($artist -eq '' -or $artist -eq $null) {    
				$artist = $PSItem.AlbumArtist.Trim()
			}
			if ($artist -eq '' -or $artist -eq $null) {
				$artist = $PSItem.Artist.Trim()
			}
			if ($artist -eq '' -or $artist -eq $null) {
				$artist = "Unknown"
			}
			$artist = Remove-InvalidFileNameChars $artist

			if ($Track.Album -ne '' -and $Track.Album -ne $null) {
				$album = Remove-InvalidFileNameChars $Track.Album.Trim()
				$trackDirectory = "$artist\$album"
			} else {
				$trackDirectory = "$artist"
			}

			# Replace a dot at the end with an underscore

			if ($trackDirectory.EndsWith('.')) {
				$trackDirectory = "$($trackDirectory.Substring(0, $trackDirectory.Length -1))_"
			}

			$relativePath = $trackDirectory

			$trackDirectory = "$Destination\$trackDirectory"

			Write-Verbose "Destination is $trackDirectory"

			# Create destination path if necessary

			if (-not (Test-Path -LiteralPath $trackDirectory -PathType Container)) {
				Write-Verbose "Creating destination directory"
				$newFolder = New-Item $trackDirectory -ItemType directory
			}

			# Copy file to destination

			Write-Verbose "Copying $location to $trackDirectory"
			$copiedFile = Copy-Item -LiteralPath $location -Destination $trackDirectory -PassThru
			$convertedTrack = "$relativePath\$($copiedFile.BaseName)$($copiedFile.Extension)"

			# Push relative path to file into pipeline, so that a playlist can be built

			$convertedTrack
		}
	}
}

<#
	.SYNOPSIS
	Removes tracks from the iTunes library and deletes the files.

	.PARAMETER Track
	iTunes track COM objects of the tracks to convert.

	.EXAMPLE
	(Get-ITunesPlayList "My Playlist").Tracks | Remove-ITunesTrack
#>
function Remove-ITunesTrack {
	[CmdletBinding()]
	param(
		[Parameter(ValueFromPipeline=$true, Mandatory=$true)]
		[System.__ComObject[]]$Track
	)
	process {
		$Track | ForEach-Object {
			$location = $PSItem.Location
			Write-Verbose "Removing $location from iTunes Library"
			$PSItem.Delete()
			Remove-Item -LiteralPath $location
		}
	}
}

<#
	.SYNOPSIS
	Resets the play count for tracks in a playlist

	.PARAMETER PlayListName
	The name of the iTunes playlist with the tracks to reset the play count.

	.PARAMETER PlayedDateOnOrBefore
	If this parameter is provided, only tracks with a last played date on or before the given date will be reset.

	.PARAMETER $PlayedCountGreaterThan
	If this parameter is provided, only tracks with a play count greater than the provided value will be reset.

	.EXAMPLE
	Resets the play count for all tracks in the playlist 'Evergeens'
	Reset-ITunesPlayCount 'Evergreens'

	.EXAMPLE
	Resets the play count for all tracks in the playlist 'All morning songs' that where not played for more than one year and which have a play count greater than 5.
	Reset-ITunesPlayCount 'All morning songs' -PlayedDateOnOrBefore (Get-Date).AddYears(-1) -PlayedCountGreaterThan 5 
#>
function Reset-ITunesPlayCount {
	[CmdletBinding()]
	Param(
		[Parameter(Position=0, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Mandatory=$true)]
		[String[]] $PlayListName,
		[datetime] $PlayedDateOnOrBefore = (Get-Date),
		[int] $PlayedCountGreaterThan = 0
	)
	process {
		(Get-ITunesPlaylist -Name $PlayListName).Tracks |
		Where-Object {
			$PSItem.PlayedDate -le $PlayedDateOnOrBefore `
			-and $PSItem.PlayedCount -gt 1
		} |
		ForEach-Object {
			$PSItem.PlayedCount = 0
		}
	}
}

# Function from: https://gallery.technet.microsoft.com/scriptcenter/Remove-Invalid-Characters-39fa17b1
<# 
	.SYNOPSIS 
	Removes characters from a string that are not valid in Windows file names. 
 
	.DESCRIPTION 
	Remove-InvalidFileNameChars accepts a string and removes characters that are invalid in Windows file names.  It then outputs the cleaned string.  It accepts value from the pipeline by the property Name.  By default the space character is ignored, but can be included using the IncludeSpace parameter. 
 
	.PARAMETER Name 
	Specifies the file name to strip of invalid characters. 
 
	.PARAMETER IncludeSpace 
	The IncludeSpace parameter will include the space character (U+0032) in the removal process. 
 
	.INPUTS 
	System.String 
	Accepts the property Name from the pipeline 
 
	.OUTPUTS 
	System.String 
 
	.EXAMPLE 
	PS C:\> Remove-InvalidFileNameChars -Name "<This /name \is* an :illegal ?filename>.txt" 
	PS C:\> This name is an illegal filename.txt 
 
	This command will strip the invalid characters from the string and output a clean string. 
 
	.EXAMPLE 
	PS C:\> Remove-InvalidFileNameChars -Name "<This /name \is* an :illegal ?filename>.txt" -IncludeSpace 
	PS C:\> Thisnameisanillegalfilename.txt 
 
	This command will strip the invalid characters from the string and output a clean string, removing the space character (U+0032) as well. 
 
	.NOTES 
	Author:  Chris Carter 
	Version: 1.1 
	Last Updated: August 28, 2014 
 
	.Link 
	System.RegEx 
	about_Join 
	about_Operators 
#> 
 
#Requires -Version 2.0 
function Remove-InvalidFileNameChars { 
	[CmdletBinding()] 
	Param( 
		[Parameter( 
			Mandatory=$true, 
			Position=0,  
			ValueFromPipelineByPropertyName=$true 
		)] 
		[String]$Name, 
		[switch]$IncludeSpace 
	) 
 
	if ($IncludeSpace) { 
		[RegEx]::Replace($Name, "[{0}]" -f ([RegEx]::Escape([String][System.IO.Path]::GetInvalidFileNameChars())), '') 
	} 
	else { 
		[RegEx]::Replace($Name, "[{0}]" -f ([RegEx]::Escape(-join [System.IO.Path]::GetInvalidFileNameChars())), '') 
	}
}

