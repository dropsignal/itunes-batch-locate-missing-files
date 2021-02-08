Public Class FormMain
    Private Function BuildLocation(BaseFolder As String, AlbumArtist As String, Album As String, DiscNumber As Integer, DiscCount As Integer, TrackNumber As Integer, TrackCount As Integer, Name As String, Kind As String) As String

        Dim sPath As String

        sPath = BaseFolder

        sPath = sPath & AlbumArtist & "\"
        sPath = sPath & Album & "\"
        If DiscCount > 1 Then
            sPath = sPath & DiscNumber.ToString & "-"
        End If
        sPath = sPath & TrackNumber.ToString("00") & " "
        sPath = sPath & Name

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
    Private Function TestLocation(Path As String) As Boolean

        Dim FSO As Scripting.FileSystemObject

        If Path <> "" Then

            FSO = New Scripting.FileSystemObject

            If FSO.FileExists(Path) Or FSO.FolderExists(Path) Then
                TestLocation = True
            Else
                TestLocation = False
            End If

            FSO = Nothing

        Else
            TestLocation = False
        End If

    End Function
    Private Function Sanitizer(Value As String) As String

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

        Sanitizer = sValue

        sValue = ""

    End Function
    Private Function SanitizeLeadingPeriod(Value As String) As String

        Dim sValue As String
        Dim iLength As Integer
        Dim cCharacter As Char

        sValue = Value
        iLength = Len(sValue)

        cCharacter = sValue.Substring(0, 1)
        If cCharacter = "." Then
            sValue = sValue.Substring(1, iLength - 1)
            sValue = "_" & sValue
        End If

        SanitizeLeadingPeriod = sValue

        sValue = ""

    End Function
    Private Function SanitizeTrailingPeriod(Value As String) As String

        Dim sValue As String
        Dim iLength As Integer
        Dim cCharacter As Char

        sValue = Value
        iLength = Len(sValue)

        cCharacter = sValue.Substring(iLength - 1, 1)
        If cCharacter = "." Then
            sValue = sValue.Substring(0, iLength - 1)
            sValue = sValue & "_"
        End If

        SanitizeTrailingPeriod = sValue

        sValue = ""

    End Function
    Private Function SanitizeLeadingSingleQuote(Value As String) As String

        Dim sValue As String
        Dim iLength As Integer
        Dim cCharacter As Char

        sValue = Value
        iLength = Len(sValue)

        cCharacter = sValue.Substring(0, 1)
        If cCharacter = "'" Then
            sValue = sValue.Substring(1, iLength - 1)
            sValue = "_" & sValue
        End If

        SanitizeLeadingSingleQuote = sValue

        sValue = ""

    End Function
    Private Function Truncator(Value As String, Length As Integer) As String

        Dim sValue As String
        Dim iLength As Integer

        sValue = Value

        If Length > 0 Then
            iLength = Len(sValue)
            If iLength > Length Then iLength = Length
            sValue = sValue.Substring(0, iLength)
            sValue = sValue.Trim
        End If

        Truncator = sValue

        sValue = ""

    End Function
    Private Function CombinationTests(Kind As String, AlbumArtist As String, Album As String, Name As String, TrackNumber As Integer, TrackCount As Integer, DiscNumber As Integer, DiscCount As Integer, Location As String) As String

        Dim sKind As String
        Dim sAlbumArtist As String
        Dim sAlbum As String
        Dim sName As String
        Dim iTrackNumber As Integer
        Dim iTrackCount As Integer
        Dim iDiscNumber As Integer
        Dim iDiscCount As Integer
        Dim sLocation As String
        Dim bPathFound As Boolean
        Dim sPath As String

        sKind = Kind
        sAlbumArtist = AlbumArtist
        sAlbum = Album
        sName = Name
        iTrackNumber = TrackNumber
        iTrackCount = TrackCount
        iDiscNumber = DiscNumber
        iDiscCount = DiscCount
        sLocation = Location

        bPathFound = False
        sPath = sLocation

        bPathFound = TestLocation(sPath)
        If Not bPathFound Then
            sAlbumArtist = Sanitizer(sAlbumArtist) : sAlbum = Sanitizer(sAlbum) : sName = Sanitizer(sName)
            sName = SanitizeLeadingSingleQuote(sName)
            sAlbum = SanitizeLeadingPeriod(sAlbum)
            sAlbumArtist = SanitizeTrailingPeriod(sAlbumArtist) : sAlbum = SanitizeTrailingPeriod(sAlbum)
            sAlbumArtist = Truncator(sAlbumArtist, 0) : sAlbum = Truncator(sAlbum, 0) : sName = Truncator(sName, 0)
            sPath = BuildLocation(TextBoxMediaFolder.Text, sAlbumArtist, sAlbum, iDiscNumber, iDiscCount, iTrackNumber, iTrackCount, sName, sKind)
            bPathFound = TestLocation(sPath)
        End If

        If Not bPathFound Then
            sAlbumArtist = Truncator(sAlbumArtist, 40) : sAlbum = Truncator(sAlbum, 40) : sName = Truncator(sName, 40)
            sPath = BuildLocation(TextBoxMediaFolder.Text, sAlbumArtist, sAlbum, iDiscNumber, iDiscCount, iTrackNumber, iTrackCount, sName, sKind)
            bPathFound = TestLocation(sPath)
        End If

        If Not bPathFound Then
            sAlbumArtist = Truncator(sAlbumArtist, 33) : sAlbum = Truncator(sAlbum, 40) : sName = Truncator(sName, 40)
            sPath = BuildLocation(TextBoxMediaFolder.Text, sAlbumArtist, sAlbum, iDiscNumber, iDiscCount, iTrackNumber, iTrackCount, sName, sKind)
            bPathFound = TestLocation(sPath)
        End If

        sAlbumArtist = AlbumArtist : sAlbum = Album : sName = Name

        If Not bPathFound Then
            sAlbumArtist = Truncator(sAlbumArtist, 40) : sAlbum = Truncator(sAlbum, 33) : sName = Truncator(sName, 40)
            sPath = BuildLocation(TextBoxMediaFolder.Text, sAlbumArtist, sAlbum, iDiscNumber, iDiscCount, iTrackNumber, iTrackCount, sName, sKind)
            bPathFound = TestLocation(sPath)
        End If

        sAlbumArtist = AlbumArtist : sAlbum = Album : sName = Name

        If Not bPathFound Then
            sAlbumArtist = Truncator(sAlbumArtist, 40) : sAlbum = Truncator(sAlbum, 40) : sName = Truncator(sName, 33)
            sPath = BuildLocation(TextBoxMediaFolder.Text, sAlbumArtist, sAlbum, iDiscNumber, iDiscCount, iTrackNumber, iTrackCount, sName, sKind)
            bPathFound = TestLocation(sPath)
        End If

        sAlbumArtist = AlbumArtist : sAlbum = Album : sName = Name

        If Not bPathFound Then
            sAlbumArtist = Truncator(sAlbumArtist, 33) : sAlbum = Truncator(sAlbum, 33) : sName = Truncator(sName, 40)
            sPath = BuildLocation(TextBoxMediaFolder.Text, sAlbumArtist, sAlbum, iDiscNumber, iDiscCount, iTrackNumber, iTrackCount, sName, sKind)
            bPathFound = TestLocation(sPath)
        End If

        sAlbumArtist = AlbumArtist : sAlbum = Album : sName = Name

        If Not bPathFound Then
            sAlbumArtist = Truncator(sAlbumArtist, 33) : sAlbum = Truncator(sAlbum, 40) : sName = Truncator(sName, 33)
            sPath = BuildLocation(TextBoxMediaFolder.Text, sAlbumArtist, sAlbum, iDiscNumber, iDiscCount, iTrackNumber, iTrackCount, sName, sKind)
            bPathFound = TestLocation(sPath)
        End If

        sAlbumArtist = AlbumArtist : sAlbum = Album : sName = Name

        If Not bPathFound Then
            sAlbumArtist = Truncator(sAlbumArtist, 40) : sAlbum = Truncator(sAlbum, 33) : sName = Truncator(sName, 33)
            sPath = BuildLocation(TextBoxMediaFolder.Text, sAlbumArtist, sAlbum, iDiscNumber, iDiscCount, iTrackNumber, iTrackCount, sName, sKind)
            bPathFound = TestLocation(sPath)
        End If

        If Not bPathFound Then
            sAlbumArtist = Truncator(sAlbumArtist, 33) : sAlbum = Truncator(sAlbum, 33) : sName = Truncator(sName, 33)
            sPath = BuildLocation(TextBoxMediaFolder.Text, sAlbumArtist, sAlbum, iDiscNumber, iDiscCount, iTrackNumber, iTrackCount, sName, sKind)
            bPathFound = TestLocation(sPath)
        End If

        sAlbumArtist = AlbumArtist : sAlbum = Album : sName = Name

        If Not bPathFound Then
            sAlbumArtist = Truncator(sAlbumArtist, 31) : sAlbum = Truncator(sAlbum, 40) : sName = Truncator(sName, 40)
            sPath = BuildLocation(TextBoxMediaFolder.Text, sAlbumArtist, sAlbum, iDiscNumber, iDiscCount, iTrackNumber, iTrackCount, sName, sKind)
            bPathFound = TestLocation(sPath)
        End If

        sAlbumArtist = AlbumArtist : sAlbum = Album : sName = Name

        If Not bPathFound Then
            sAlbumArtist = Truncator(sAlbumArtist, 40) : sAlbum = Truncator(sAlbum, 31) : sName = Truncator(sName, 40)
            sPath = BuildLocation(TextBoxMediaFolder.Text, sAlbumArtist, sAlbum, iDiscNumber, iDiscCount, iTrackNumber, iTrackCount, sName, sKind)
            bPathFound = TestLocation(sPath)
        End If

        sAlbumArtist = AlbumArtist : sAlbum = Album : sName = Name

        If Not bPathFound Then
            sAlbumArtist = Truncator(sAlbumArtist, 40) : sAlbum = Truncator(sAlbum, 40) : sName = Truncator(sName, 31)
            sPath = BuildLocation(TextBoxMediaFolder.Text, sAlbumArtist, sAlbum, iDiscNumber, iDiscCount, iTrackNumber, iTrackCount, sName, sKind)
            bPathFound = TestLocation(sPath)
        End If

        sAlbumArtist = AlbumArtist : sAlbum = Album : sName = Name

        If Not bPathFound Then
            sAlbumArtist = Truncator(sAlbumArtist, 31) : sAlbum = Truncator(sAlbum, 31) : sName = Truncator(sName, 40)
            sPath = BuildLocation(TextBoxMediaFolder.Text, sAlbumArtist, sAlbum, iDiscNumber, iDiscCount, iTrackNumber, iTrackCount, sName, sKind)
            bPathFound = TestLocation(sPath)
        End If

        sAlbumArtist = AlbumArtist : sAlbum = Album : sName = Name

        If Not bPathFound Then
            sAlbumArtist = Truncator(sAlbumArtist, 31) : sAlbum = Truncator(sAlbum, 40) : sName = Truncator(sName, 31)
            sPath = BuildLocation(TextBoxMediaFolder.Text, sAlbumArtist, sAlbum, iDiscNumber, iDiscCount, iTrackNumber, iTrackCount, sName, sKind)
            bPathFound = TestLocation(sPath)
        End If

        sAlbumArtist = AlbumArtist : sAlbum = Album : sName = Name

        If Not bPathFound Then
            sAlbumArtist = Truncator(sAlbumArtist, 40) : sAlbum = Truncator(sAlbum, 31) : sName = Truncator(sName, 31)
            sPath = BuildLocation(TextBoxMediaFolder.Text, sAlbumArtist, sAlbum, iDiscNumber, iDiscCount, iTrackNumber, iTrackCount, sName, sKind)
            bPathFound = TestLocation(sPath)
        End If

        If Not bPathFound Then
            sAlbumArtist = Truncator(sAlbumArtist, 31) : sAlbum = Truncator(sAlbum, 31) : sName = Truncator(sName, 31)
            sPath = BuildLocation(TextBoxMediaFolder.Text, sAlbumArtist, sAlbum, iDiscNumber, iDiscCount, iTrackNumber, iTrackCount, sName, sKind)
            bPathFound = TestLocation(sPath)
        End If

        sAlbumArtist = AlbumArtist : sAlbum = Album : sName = Name

        If Not bPathFound Then
            sAlbumArtist = Truncator(sAlbumArtist, 31) : sAlbum = Truncator(sAlbum, 33) : sName = Truncator(sName, 33)
            sPath = BuildLocation(TextBoxMediaFolder.Text, sAlbumArtist, sAlbum, iDiscNumber, iDiscCount, iTrackNumber, iTrackCount, sName, sKind)
            bPathFound = TestLocation(sPath)
        End If

        sAlbumArtist = AlbumArtist : sAlbum = Album : sName = Name

        If Not bPathFound Then
            sAlbumArtist = Truncator(sAlbumArtist, 33) : sAlbum = Truncator(sAlbum, 31) : sName = Truncator(sName, 33)
            sPath = BuildLocation(TextBoxMediaFolder.Text, sAlbumArtist, sAlbum, iDiscNumber, iDiscCount, iTrackNumber, iTrackCount, sName, sKind)
            bPathFound = TestLocation(sPath)
        End If

        sAlbumArtist = AlbumArtist : sAlbum = Album : sName = Name

        If Not bPathFound Then
            sAlbumArtist = Truncator(sAlbumArtist, 33) : sAlbum = Truncator(sAlbum, 33) : sName = Truncator(sName, 31)
            sPath = BuildLocation(TextBoxMediaFolder.Text, sAlbumArtist, sAlbum, iDiscNumber, iDiscCount, iTrackNumber, iTrackCount, sName, sKind)
            bPathFound = TestLocation(sPath)
        End If

        sAlbumArtist = AlbumArtist : sAlbum = Album : sName = Name

        If Not bPathFound Then
            sAlbumArtist = Truncator(sAlbumArtist, 31) : sAlbum = Truncator(sAlbum, 31) : sName = Truncator(sName, 33)
            sPath = BuildLocation(TextBoxMediaFolder.Text, sAlbumArtist, sAlbum, iDiscNumber, iDiscCount, iTrackNumber, iTrackCount, sName, sKind)
            bPathFound = TestLocation(sPath)
        End If

        sAlbumArtist = AlbumArtist : sAlbum = Album : sName = Name

        If Not bPathFound Then
            sAlbumArtist = Truncator(sAlbumArtist, 31) : sAlbum = Truncator(sAlbum, 33) : sName = Truncator(sName, 31)
            sPath = BuildLocation(TextBoxMediaFolder.Text, sAlbumArtist, sAlbum, iDiscNumber, iDiscCount, iTrackNumber, iTrackCount, sName, sKind)
            bPathFound = TestLocation(sPath)
        End If

        sAlbumArtist = AlbumArtist : sAlbum = Album : sName = Name

        If Not bPathFound Then
            sAlbumArtist = Truncator(sAlbumArtist, 33) : sAlbum = Truncator(sAlbum, 31) : sName = Truncator(sName, 31)
            sPath = BuildLocation(TextBoxMediaFolder.Text, sAlbumArtist, sAlbum, iDiscNumber, iDiscCount, iTrackNumber, iTrackCount, sName, sKind)
            bPathFound = TestLocation(sPath)
        End If

        If bPathFound Then
            CombinationTests = sPath
        Else
            CombinationTests = ""
        End If

        sPath = ""
        bPathFound = False

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
        Dim sCalculatedPath As String
        Dim bPathFound As Boolean
        Dim sPath As String
        Dim sMessage As String

        TextBoxMediaFolder.Enabled = False
        ButtonBrowse.Enabled = False
        TextBoxLog.Enabled = False
        ButtonScan.Enabled = False

        FSO = New Scripting.FileSystemObject

        oTunes = New iTunesLib.iTunesApp
        oTracks = oTunes.LibraryPlaylist.Tracks

        lTrackCount = oTracks.Count

        ProgressBar1.Maximum = 1
        ProgressBar1.Maximum = lTrackCount

        For lTrackIndex = 1 To lTrackCount

            ProgressBar1.Value = lTrackIndex

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
                    Case "iTunes LP", "MPEG audio file", "MPEG-4 video file", "Protected AAC audio file", "Protected MPEG-4 video file", "Purchased AAC audio file", "Purchased MPEG-4 video file"

                        sPath = CombinationTests(sKind, sAlbumArtist, sAlbum, sName, iTrackNumber, iTrackCount, iDiscNumber, iDiscCount, sLocation)

                    Case "Apple Lossless audio file"

                        sCalculatedPath = BuildLocation(TextBoxMediaFolder.Text, sAlbumArtist, sAlbum, iDiscNumber, iDiscCount, iTrackNumber, iTrackCount, sName, sKind)
                        bPathFound = False
                        sPath = ""

                        sPath = CombinationTests("MPEG audio file", sAlbumArtist, sAlbum, sName, iTrackNumber, iTrackCount, iDiscNumber, iDiscCount, sLocation)

                        If sPath <> "" Then

                            sMessage = "Updating "
                            sMessage = sMessage & sCalculatedPath
                            sMessage = sMessage & " to "
                            sMessage = sMessage & sPath & vbCrLf
                            TextBoxLog.AppendText(sMessage)
                            sMessage = ""

                            If lTrackIndex <> 4536 Then
                                oTrack.Location = sPath
                                oTrack.UpdateInfoFromFile()
                            End If

                        Else
                            sMessage = "Failed to find alternative to "
                            sMessage = sMessage & sCalculatedPath & vbCrLf
                            TextBoxLog.AppendText(sMessage)
                            sMessage = ""
                        End If

                        sPath = ""
                        bPathFound = False
                        sCalculatedPath = ""

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
