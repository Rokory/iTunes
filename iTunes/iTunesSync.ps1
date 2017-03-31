# This script shows, how to use iTunes.psm1 
# to convert several playlists at once 
# and renaming them at destination at the same time.
# 
# It pipes a list of hashtables, 
# each containing a playlist name and a new name, 
# into the Convert-ItunesPlaylist cmdlet.
# 
# You can use this script, 
# by customizing the $destination variable at the beginning, 
# e. g. with the drive letter of your USB stick,
# and customizing the playlist names and new names in the hash table.
# Moreover, you can delete lines from the list or add new ones.

$destination = 'G:\'

@{ playlist = "Beschwingt alle";       newname="Ford1" } ,
@{ playlist = "Ab geht's alle";        newname="Ford2" },
@{ playlist = "Vormittag sämtliche";   newname="Ford3" },
@{ playlist = "Aufgeregt alle";        newname="Ford4" },
@{ playlist = "Nachmittag sämtliche";  newname="Ford6" },
@{ playlist = "Lässig alle";           newname="Ford7" },
@{ playlist = "Beruhigend alle";       newname="Ford8" },
@{ playlist = "Abend sämtliche";       newname="Ford9"} |
Convert-ITunesPlaylist -Destination $destination -Verbose
