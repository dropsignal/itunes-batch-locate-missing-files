# itunes-batch-locate-missing-files
If you downgrade your iTunes library from ALAC to MP3 using an external transcoder (like xrecode), this utility will update your iTunes library so the MP3 files are used without losing your play counts or ratings. This project also serves an as an example of using the iTunes COM Type Library.

Note: If you do not know what you are doing, please do not attempt to follow these instructions as you will more than likely destroy your iTunes library.

To downgrade an iTunes library from ALAC to MP3:
Shutdown iTunes
Make a backup copy of your iTunes folder (iTunes -> iTunes Backup 2021-02-07.zip)
Move the contents of iTunes\iTunes Media\Music to a temporary folder (D:\iTunes\iTunes Media\Music\* -> D:\For Conversion)
Use a third-party tool to transcode your music. I used Xrecode. Use the temporary folder as the source and iTunes\iTunes Media\Music folder as the destination.
Run itunes-batch-locate-missing-files. Select iTunes\iTunes Media\Music as the source folder. Click scan.
If everything works out, iTunes will start and the utility will update iTunes pointing your library to the MP3 files. Your playlists, ratings, and play counts should remain in tact.
