Public Class FormMain
    Private Function BuildLocation(BaseFolder As String, AlbumArtist As String, Album As String, DiscNumber As Integer, DiscCount As Integer, TrackNumber As Integer, TrackCount As Integer, Name As String, Kind As String) As String

        Dim sPath As String
        Dim sNewAlbumArtist As String
        Dim sNewAlbum As String
        Dim sNewName As String

        sNewAlbumArtist = Truncate(AlbumArtist, 33)
        sNewAlbum = Truncate(Album, 40)
        sNewName = Truncate(Name, 33)

        sPath = BaseFolder

        sPath = sPath & Sanitize(sNewAlbumArtist) & "\"
        sPath = sPath & Sanitize(sNewAlbum) & "\"
        If DiscCount > 1 Then
            sPath = sPath & DiscNumber.ToString & "-"
        End If
        sPath = sPath & TrackNumber.ToString("00") & " "
        sPath = sPath & Sanitize(sNewName)

        Select Case Kind
            Case "Purchased AAC audio file"
                sPath = sPath & ".m4a"
            Case "MPEG audio file"
                sPath = sPath & ".mp3"
            Case "Protected AAC audio file"
                sPath = sPath & ".m4a"
            Case "Purchased MPEG-4 video file"
                sPath = sPath & ".m4v"
            Case "iTunes LP"
                sPath = sPath & ".itlp"
            Case "Protected MPEG-4 video file"
                sPath = sPath & ".m4v"
            Case "MPEG-4 video file"
                sPath = sPath & ".mp4"
            Case "Apple Lossless audio file"
                sPath = sPath & ".m4a"
        End Select

        BuildLocation = sPath

        sPath = ""

    End Function
    Private Function Sanitize(Value As String) As String

        Dim sValue As String

        sValue = Value

        sValue = sValue.Replace("\", "_")
        sValue = sValue.Replace("/", "_")
        sValue = sValue.Replace(":", "_")
        sValue = sValue.Replace("*", "_")
        sValue = sValue.Replace("?", "_")
        sValue = sValue.Replace(Chr(34), "_")
        sValue = sValue.Replace("<", "_")
        sValue = sValue.Replace(">", "_")
        sValue = sValue.Replace("|", "_")

        Sanitize = sValue

        sValue = ""

    End Function
    Private Function Truncate(Value As String, Length As Integer) As String

        Dim sValue As String
        Dim iLength As Integer

        sValue = Value

        iLength = Len(sValue)
        If iLength > Length Then iLength = Length
        sValue = sValue.Substring(0, iLength)
        sValue = sValue.Trim

        Truncate = sValue

        sValue = ""

    End Function
    Private Sub ButtonScan_Click(sender As Object, e As EventArgs) Handles ButtonScan.Click

        Dim FSO As Scripting.FileSystemObject
        Dim oTunes As iTunesLib.iTunesApp
        Dim oTracks As iTunesLib.IITTrackCollection
        Dim lTrackCount As Long
        Dim lTrackIndex As Long
        Dim oTrack As iTunesLib.IITFileOrCDTrack
        Dim sKind As String
        Dim sAlbumArtist As String
        Dim sAlbum As String
        Dim sName As String
        Dim iTrackNumber As Integer
        Dim iTrackCount As Integer
        Dim iDiscNumber As Integer
        Dim iDiscCount As Integer
        Dim sLocation As String
        Dim sPath As String
        Dim sMessage As String
        Dim sNewLocation As String

        TextBoxMediaFolder.Enabled = False
        ButtonBrowse.Enabled = False
        TextBoxLog.Enabled = False
        ButtonScan.Enabled = False

        FSO = New Scripting.FileSystemObject

        oTunes = New iTunesLib.iTunesApp
        oTracks = oTunes.LibraryPlaylist.Tracks

        lTrackCount = oTracks.Count

        For lTrackIndex = 1 To lTrackCount

            If oTracks.Item(lTrackIndex).Kind = iTunesLib.ITTrackKind.ITTrackKindFile Then

                oTrack = oTracks.Item(lTrackIndex)

                sKind = oTrack.KindAsString
                sAlbumArtist = oTrack.AlbumArtist
                sAlbum = oTrack.Album
                sName = oTrack.Name
                iTrackNumber = oTrack.TrackNumber
                iTrackCount = oTrack.TrackCount
                iDiscNumber = oTrack.DiscNumber
                iDiscCount = oTrack.DiscCount
                sLocation = oTrack.Location

                Select Case sKind
                    Case "Purchased AAC audio file", "MPEG audio file", "Protected AAC audio file", "Purchased MPEG-4 video file", "Protected MPEG-4 video file", "MPEG-4 video file"

                        If Not FSO.FileExists(sLocation) Then

                            sPath = BuildLocation(TextBoxMediaFolder.Text, sAlbumArtist, sAlbum, iDiscNumber, iDiscCount, iTrackNumber, iTrackCount, sName, sKind)

                            If FSO.FileExists(sPath) Then
                                sMessage = "Found: "
                                sMessage = sMessage & sPath & vbCrLf
                                TextBoxLog.AppendText(sMessage)
                                sMessage = ""
                            Else
                                sMessage = "Missing: "
                                sMessage = sMessage & sPath & vbCrLf
                                TextBoxLog.AppendText(sMessage)
                                sMessage = ""
                            End If

                            sPath = ""

                        End If

                    Case "iTunes LP"

                        If Not FSO.FolderExists(sLocation) Then

                            sPath = BuildLocation(TextBoxMediaFolder.Text, sAlbumArtist, sAlbum, iDiscNumber, iDiscCount, iTrackNumber, iTrackCount, sName, sKind)

                            If FSO.FileExists(sPath) Then
                                sMessage = "Found: "
                                sMessage = sMessage & sPath & vbCrLf
                                TextBoxLog.AppendText(sMessage)
                                sMessage = ""
                            Else
                                sMessage = "Missing: "
                                sMessage = sMessage & sPath & vbCrLf
                                TextBoxLog.AppendText(sMessage)
                                sMessage = ""
                            End If

                            sPath = ""

                        End If

                    Case "Apple Lossless audio file"

                        If sLocation = "" Then

                            sPath = BuildLocation(TextBoxMediaFolder.Text, sAlbumArtist, sAlbum, iDiscNumber, iDiscCount, iTrackNumber, iTrackCount, sName, sKind)

                            sMessage = "Missing: "
                            sMessage = sMessage & sPath & vbCrLf
                            TextBoxLog.AppendText(sMessage)
                            sMessage = ""

                            sNewLocation = BuildLocation(TextBoxMediaFolder.Text, sAlbumArtist, sAlbum, iDiscNumber, iDiscCount, iTrackNumber, iTrackCount, sName, "MPEG audio file")

                            If FSO.FileExists(sNewLocation) Then

                                sMessage = "Updating "
                                sMessage = sMessage & sPath
                                sMessage = sMessage & " to "
                                sMessage = sMessage & sNewLocation & vbCrLf
                                TextBoxLog.AppendText(sMessage)
                                sMessage = ""

                                'oTrack.Location = sNewLocation
                                'oTrack.UpdateInfoFromFile()

                            Else

                                sMessage = "Missing: "
                                sMessage = sMessage & sNewLocation & vbCrLf
                                TextBoxLog.AppendText(sMessage)
                                sMessage = ""

                            End If

                            sNewLocation = ""

                            sPath = Nothing

                        End If

                End Select

                sLocation = ""
                iDiscCount = 0
                iDiscNumber = 0
                iTrackCount = 0
                iTrackNumber = 0
                sName = ""
                sAlbum = ""
                sAlbumArtist = ""
                sKind = ""

                oTrack = Nothing

            Else

                sMessage = "Unsupported: "

                Select Case oTracks.Item(lTrackIndex).Kind
                    Case iTunesLib.ITTrackKind.ITTrackKindCD
                        sMessage = sMessage & "ITTrackKindCD" & vbCrLf
                    Case iTunesLib.ITTrackKind.ITTrackKindDevice
                        sMessage = sMessage & "ITTrackKindDevice" & vbCrLf
                    Case iTunesLib.ITTrackKind.ITTrackKindSharedLibrary
                        sMessage = sMessage & "ITTrackKindSharedLibrary" & vbCrLf
                    Case iTunesLib.ITTrackKind.ITTrackKindUnknown
                        sMessage = sMessage & "ITTrackKindUnknown" & vbCrLf
                    Case iTunesLib.ITTrackKind.ITTrackKindURL
                        sMessage = sMessage & "ITTrackKindURL" & vbCrLf
                End Select

                TextBoxLog.AppendText(sMessage)
                sMessage = ""

            End If
        Next

        lTrackCount = 0

        oTracks = Nothing
        oTunes = Nothing

        FSO = Nothing

        ButtonScan.Enabled = True
        TextBoxLog.Enabled = True
        ButtonBrowse.Enabled = True
        TextBoxMediaFolder.Enabled = True

    End Sub

    Private Sub ButtonBrowse_Click(sender As Object, e As EventArgs) Handles ButtonBrowse.Click

        Dim FBD As FolderBrowserDialog

        FBD = New FolderBrowserDialog
        If FBD.ShowDialog = DialogResult.OK Then TextBoxMediaFolder.Text = FBD.SelectedPath
        FBD = Nothing

    End Sub

    Private Sub TextBoxMediaFolder_TextChanged(sender As Object, e As EventArgs) Handles TextBoxMediaFolder.TextChanged

        Dim FSO As Scripting.FileSystemObject

        FSO = New Scripting.FileSystemObject

        TextBoxMediaFolder.Text = Trim(TextBoxMediaFolder.Text)
        If TextBoxMediaFolder.Text <> "" Then
            If TextBoxMediaFolder.Text.EndsWith("\") = False Then TextBoxMediaFolder.Text = TextBoxMediaFolder.Text & "\"
        End If

        Select Case TextBoxMediaFolder.Text
            Case = ""
                ButtonScan.Enabled = False
            Case <> ""
                If FSO.FolderExists(TextBoxMediaFolder.Text) Then
                    ButtonScan.Enabled = True
                Else
                    ButtonScan.Enabled = False
                End If
        End Select

        FSO = Nothing

    End Sub
End Class
