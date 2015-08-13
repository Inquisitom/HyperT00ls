Imports System.IO

Public Class FrmCopyMedia
    Private OrigSystem As String
    Private DestSystem As String
    Private OrigRom As String
    Private DestRom As String

    Public Sub New(ByVal System As String, ByVal Rom As String)
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.

        OrigSystem = System
        OrigRom = Rom

        'Default config
        Label1.Text = "From: [" & OrigSystem & "]/[" & OrigRom & "]"
        RadioButton1.Checked = True
        RadioButton2.Checked = False

        CheckBox1.Checked = True
        CheckBox2.Checked = True

        For Each Entry In gHS.Systems
            If Entry.IsEnabled = True And Entry.IsEXE = False Then
                ComboBox1.Items.Add(Entry.Name)
            End If
        Next
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If DestRom = OrigRom And DestSystem = OrigSystem Then
                MsgBox("Nothing to do -> same system & same rom !", MsgBoxStyle.Information, "St00pid?")
            Else
                CopyMedia()
            End If
        Catch ex As Exception
            LogEntry(LogType._Error, "{0}", ex.Message.ToString)
        End Try
    End Sub

    Private Sub CopyMedia()
        Dim CopyMedia As Boolean = RadioButton1.Checked
        Dim MoveMedia As Boolean = RadioButton2.Checked

        Dim IncludeHyperSpinMedia As Boolean = CheckBox1.Checked
        Dim IncludeHyperPauseMedia As Boolean = CheckBox2.Checked
        Dim Overwrite As Boolean = CheckBox4.Checked

        If ComboBox1.SelectedItem.ToString = "" Or ComboBox2.SelectedItem.ToString = "" Then
            MsgBox("Please select a System & Rom", MsgBoxStyle.Information, "Info")
            Exit Sub
        End If

        'copy data or not
        If CopyMedia = True Or MoveMedia = True Then
            Dim Orig As String = ""
            Dim Dest As String = ""

            If IncludeHyperSpinMedia = True Then
                'then copy theme, wheel, video, art1, art2, art3, art4
                LogEntry(LogType._Info, "{0}", "Copy started (overwrite=" & Overwrite.ToString & ") --> ")

                'Theme
                Orig = gHS.Systems(OrigSystem).Roms(OrigRom).ThemePath
                Dest = gHS.Systems(DestSystem).Roms(DestRom).ThemePath
                Try
                    If File.Exists(Orig) Then
                        LogEntry(LogType._Info, "{0}", "     Source : " & Orig & "  ---> Dest : " & Dest)
                        If Directory.Exists(Path.GetDirectoryName(Dest)) = False Then
                            Directory.CreateDirectory(Path.GetDirectoryName(Dest))
                            LogEntry(LogType._Info, "{0}", "Creating folder : " & Path.GetDirectoryName(Dest))
                        End If
                        File.Copy(Orig, Dest, Overwrite)
                        Application.DoEvents()
                        If MoveMedia = True Then
                            File.Delete(Orig)
                        End If
                    Else
                        LogEntry(LogType._Info, "{0}", "     Theme not existing : skipping")
                    End If
                Catch ex As Exception
                    LogEntry(LogType._Error, "{0}", "     Cannot perform copy : " & ex.Message.ToString)
                End Try

                'Video
                Orig = gHS.Systems(OrigSystem).Roms(OrigRom).VideoPath
                Dest = gHS.Systems(DestSystem).Roms(DestRom).VideoPath
                If Dest = "" Then
                    Dest = gHSPath & "\Media\" & DestSystem & "\Video\" & DestRom & Path.GetExtension(gHS.Systems(OrigSystem).Roms(OrigRom).VideoPath)
                End If
                Try
                    If File.Exists(Orig) Then
                        LogEntry(LogType._Info, "{0}", "     Source : " & Orig & "  ---> Dest : " & Dest)
                        If Directory.Exists(Path.GetDirectoryName(Dest)) = False Then
                            Directory.CreateDirectory(Path.GetDirectoryName(Dest))
                            LogEntry(LogType._Info, "{0}", "Creating folder : " & Path.GetDirectoryName(Dest))
                        End If
                        File.Copy(Orig, Dest, Overwrite)
                        Application.DoEvents()
                        If MoveMedia = True Then
                            File.Delete(Orig)
                        End If
                    Else
                        LogEntry(LogType._Info, "{0}", "     Video not existing : skipping")
                    End If
                Catch ex As Exception
                    LogEntry(LogType._Error, "{0}", "     Cannot perform copy : " & ex.Message.ToString)
                End Try

                'Wheel
                Orig = gHS.Systems(OrigSystem).Roms(OrigRom).WheelPath
                Dest = gHS.Systems(DestSystem).Roms(DestRom).WheelPath
                Try
                    If File.Exists(Orig) Then
                        LogEntry(LogType._Info, "{0}", "     Source : " & Orig & "  ---> Dest : " & Dest)
                        If Directory.Exists(Path.GetDirectoryName(Dest)) = False Then
                            Directory.CreateDirectory(Path.GetDirectoryName(Dest))
                            LogEntry(LogType._Info, "{0}", "Creating folder : " & Path.GetDirectoryName(Dest))
                        End If
                        File.Copy(Orig, Dest, Overwrite)
                        Application.DoEvents()
                        If MoveMedia = True Then
                            File.Delete(Orig)
                        End If
                    Else
                        LogEntry(LogType._Info, "{0}", "     Wheel not existing : skipping")
                    End If
                Catch ex As Exception
                    LogEntry(LogType._Error, "{0}", "     Cannot perform copy : " & ex.Message.ToString)
                End Try

                'Art1
                Orig = gHS.Systems(OrigSystem).Roms(OrigRom).Art1Path
                Dest = gHS.Systems(DestSystem).Roms(DestRom).Art1Path
                Try
                    If File.Exists(Orig) Then
                        LogEntry(LogType._Info, "{0}", "     Source : " & Orig & "  ---> Dest : " & Dest)
                        If Directory.Exists(Path.GetDirectoryName(Dest)) = False Then
                            Directory.CreateDirectory(Path.GetDirectoryName(Dest))
                            LogEntry(LogType._Info, "{0}", "Creating folder : " & Path.GetDirectoryName(Dest))
                        End If
                        File.Copy(Orig, Dest, Overwrite)
                        Application.DoEvents()
                        If MoveMedia = True Then
                            File.Delete(Orig)
                        End If
                    Else
                        LogEntry(LogType._Info, "{0}", "     Art1 not existing : skipping")
                    End If
                Catch ex As Exception
                    LogEntry(LogType._Error, "{0}", "     Cannot perform copy : " & ex.Message.ToString)
                End Try

                'Art2
                Orig = gHS.Systems(OrigSystem).Roms(OrigRom).Art2Path
                Dest = gHS.Systems(DestSystem).Roms(DestRom).Art2Path
                Try
                    If File.Exists(Orig) Then
                        LogEntry(LogType._Info, "{0}", "     Source : " & Orig & "  ---> Dest : " & Dest)
                        If Directory.Exists(Path.GetDirectoryName(Dest)) = False Then
                            Directory.CreateDirectory(Path.GetDirectoryName(Dest))
                            LogEntry(LogType._Info, "{0}", "Creating folder : " & Path.GetDirectoryName(Dest))
                        End If
                        File.Copy(Orig, Dest, Overwrite)
                        Application.DoEvents()
                        If MoveMedia = True Then
                            File.Delete(Orig)
                        End If
                    Else
                        LogEntry(LogType._Info, "{0}", "     Art2 not existing : skipping")
                    End If
                Catch ex As Exception
                    LogEntry(LogType._Error, "{0}", "     Cannot perform copy : " & ex.Message.ToString)
                End Try

                'Art3
                Orig = gHS.Systems(OrigSystem).Roms(OrigRom).Art3Path
                Dest = gHS.Systems(DestSystem).Roms(DestRom).Art3Path
                Try
                    If File.Exists(Orig) Then
                        LogEntry(LogType._Info, "{0}", "     Source : " & Orig & "  ---> Dest : " & Dest)
                        If Directory.Exists(Path.GetDirectoryName(Dest)) = False Then
                            Directory.CreateDirectory(Path.GetDirectoryName(Dest))
                            LogEntry(LogType._Info, "{0}", "Creating folder : " & Path.GetDirectoryName(Dest))
                        End If
                        File.Copy(Orig, Dest, Overwrite)
                        Application.DoEvents()
                        If MoveMedia = True Then
                            File.Delete(Orig)
                        End If
                    Else
                        LogEntry(LogType._Info, "{0}", "     Art3 not existing : skipping")
                    End If
                Catch ex As Exception
                    LogEntry(LogType._Error, "{0}", "     Cannot perform copy : " & ex.Message.ToString)
                End Try

                'Art4
                Orig = gHS.Systems(OrigSystem).Roms(OrigRom).Art4Path
                Dest = gHS.Systems(DestSystem).Roms(DestRom).Art4Path
                Try
                    If File.Exists(Orig) Then
                        LogEntry(LogType._Info, "{0}", "     Source : " & Orig & "  ---> Dest : " & Dest)
                        If Directory.Exists(Path.GetDirectoryName(Dest)) = False Then
                            Directory.CreateDirectory(Path.GetDirectoryName(Dest))
                            LogEntry(LogType._Info, "{0}", "Creating folder : " & Path.GetDirectoryName(Dest))
                        End If
                        File.Copy(Orig, Dest, Overwrite)
                        Application.DoEvents()
                        If MoveMedia = True Then
                            File.Delete(Orig)
                        End If
                    Else
                        LogEntry(LogType._Info, "{0}", "     Art4 not existing : skipping")
                    End If
                Catch ex As Exception
                    LogEntry(LogType._Error, "{0}", "     Cannot perform copy : " & ex.Message.ToString)
                End Try

            End If

            If IncludeHyperPauseMedia = True Then
                'Copy :
                'Artwork
                Orig = gHS.Systems(OrigSystem).Roms(OrigRom).ArtworkPath
                If Orig.EndsWith("\") = True Then
                    Orig = Directory.GetParent(Orig).FullName
                End If
                Dest = gHS.Systems(DestSystem).Roms(DestRom).ArtworkPath
                Try
                    If Directory.Exists(Orig) Then
                        LogEntry(LogType._Info, "{0}", "     Source : " & Orig & "  ---> Dest : " & Dest)
                        My.Computer.FileSystem.CopyDirectory(Orig, Dest, Overwrite)
                        Application.DoEvents()
                        If MoveMedia = True Then
                            Directory.Delete(Orig, True)
                        End If
                    End If
                Catch ex As Exception
                    LogEntry(LogType._Error, "{0}", "     Cannot perform copy : " & ex.Message.ToString)
                End Try

                'Backgrounds
                Orig = gHS.Systems(OrigSystem).Roms(OrigRom).BackgroundsPath
                If Orig.EndsWith("\") = True Then
                    Orig = Directory.GetParent(Orig).FullName
                End If
                Dest = gHS.Systems(DestSystem).Roms(DestRom).BackgroundsPath
                Try
                    If Directory.Exists(Orig) Then
                        LogEntry(LogType._Info, "{0}", "     Source : " & Orig & "  ---> Dest : " & Dest)
                        My.Computer.FileSystem.CopyDirectory(Orig, Dest, Overwrite)
                        Application.DoEvents()
                        If MoveMedia = True Then
                            Directory.Delete(Orig, True)
                        End If
                    End If
                Catch ex As Exception
                    LogEntry(LogType._Error, "{0}", "     Cannot perform copy : " & ex.Message.ToString)
                End Try

                'Bezels
                Orig = gHS.Systems(OrigSystem).Roms(OrigRom).BezelsPath
                If Orig.EndsWith("\") = True Then
                    Orig = Directory.GetParent(Orig).FullName
                End If
                Dest = gHS.Systems(DestSystem).Roms(DestRom).BezelsPath
                Try
                    If Directory.Exists(Orig) Then
                        LogEntry(LogType._Info, "{0}", "     Source : " & Orig & "  ---> Dest : " & Dest)
                        My.Computer.FileSystem.CopyDirectory(Orig, Dest, Overwrite)
                        Application.DoEvents()
                        If MoveMedia = True Then
                            Directory.Delete(Orig, True)
                        End If
                    End If
                Catch ex As Exception
                    LogEntry(LogType._Error, "{0}", "     Cannot perform copy : " & ex.Message.ToString)
                End Try

                'Controller
                Orig = gHS.Systems(OrigSystem).Roms(OrigRom).ControllerPath
                If Orig.EndsWith("\") = True Then
                    Orig = Directory.GetParent(Orig).FullName
                End If
                Dest = gHS.Systems(DestSystem).Roms(DestRom).ControllerPath
                Try
                    If Directory.Exists(Orig) Then
                        LogEntry(LogType._Info, "{0}", "     Source : " & Orig & "  ---> Dest : " & Dest)
                        My.Computer.FileSystem.CopyDirectory(Orig, Dest, Overwrite)
                        Application.DoEvents()
                        If MoveMedia = True Then
                            Directory.Delete(Orig, True)
                        End If
                    End If
                Catch ex As Exception
                    LogEntry(LogType._Error, "{0}", "     Cannot perform copy : " & ex.Message.ToString)
                End Try

                'Fade
                Orig = gHS.Systems(OrigSystem).Roms(OrigRom).FadePath
                If Orig.EndsWith("\") = True Then
                    Orig = Directory.GetParent(Orig).FullName
                End If
                Dest = gHS.Systems(DestSystem).Roms(DestRom).FadePath
                Try
                    If Directory.Exists(Orig) Then
                        LogEntry(LogType._Info, "{0}", "     Source : " & Orig & "  ---> Dest : " & Dest)
                        My.Computer.FileSystem.CopyDirectory(Orig, Dest, Overwrite)
                        Application.DoEvents()
                        If MoveMedia = True Then
                            Directory.Delete(Orig, True)
                        End If
                    End If
                Catch ex As Exception
                    LogEntry(LogType._Error, "{0}", "     Cannot perform copy : " & ex.Message.ToString)
                End Try

                'Fonts
                Orig = gHS.Systems(OrigSystem).Roms(OrigRom).FontsPath
                If Orig.EndsWith("\") = True Then
                    Orig = Directory.GetParent(Orig).FullName
                End If
                Dest = gHS.Systems(DestSystem).Roms(DestRom).FontsPath
                Try
                    If Directory.Exists(Orig) Then
                        LogEntry(LogType._Info, "{0}", "     Source : " & Orig & "  ---> Dest : " & Dest)
                        My.Computer.FileSystem.CopyDirectory(Orig, Dest, Overwrite)
                        Application.DoEvents()
                        If MoveMedia = True Then
                            Directory.Delete(Orig, True)
                        End If
                    End If
                Catch ex As Exception
                    LogEntry(LogType._Error, "{0}", "     Cannot perform copy : " & ex.Message.ToString)
                End Try

                'Guides
                Orig = gHS.Systems(OrigSystem).Roms(OrigRom).GuidesPath
                If Orig.EndsWith("\") = True Then
                    Orig = Directory.GetParent(Orig).FullName
                End If
                Dest = gHS.Systems(DestSystem).Roms(DestRom).GuidesPath
                Try
                    If Directory.Exists(Orig) Then
                        LogEntry(LogType._Info, "{0}", "     Source : " & Orig & "  ---> Dest : " & Dest)
                        My.Computer.FileSystem.CopyDirectory(Orig, Dest, Overwrite)
                        Application.DoEvents()
                        If MoveMedia = True Then
                            Directory.Delete(Orig, True)
                        End If
                    End If
                Catch ex As Exception
                    LogEntry(LogType._Error, "{0}", "     Cannot perform copy : " & ex.Message.ToString)
                End Try

                'Manuals
                Orig = gHS.Systems(OrigSystem).Roms(OrigRom).ManualsPath
                If Orig.EndsWith("\") = True Then
                    Orig = Directory.GetParent(Orig).FullName
                End If
                Dest = gHS.Systems(DestSystem).Roms(DestRom).ManualsPath
                Try
                    If Directory.Exists(Orig) Then
                        LogEntry(LogType._Info, "{0}", "     Source : " & Orig & "  ---> Dest : " & Dest)
                        My.Computer.FileSystem.CopyDirectory(Orig, Dest, Overwrite)
                        Application.DoEvents()
                        If MoveMedia = True Then
                            Directory.Delete(Orig, True)
                        End If
                    End If
                Catch ex As Exception
                    LogEntry(LogType._Error, "{0}", "     Cannot perform copy : " & ex.Message.ToString)
                End Try

                'Menu Images
                Orig = gHS.Systems(OrigSystem).Roms(OrigRom).MenuImagesPath
                If Orig.EndsWith("\") = True Then
                    Orig = Directory.GetParent(Orig).FullName
                End If
                Dest = gHS.Systems(DestSystem).Roms(DestRom).MenuImagesPath
                Try
                    If Directory.Exists(Orig) Then
                        LogEntry(LogType._Info, "{0}", "     Source : " & Orig & "  ---> Dest : " & Dest)
                        My.Computer.FileSystem.CopyDirectory(Orig, Dest, Overwrite)
                        Application.DoEvents()
                        If MoveMedia = True Then
                            Directory.Delete(Orig, True)
                        End If
                    End If
                Catch ex As Exception
                    LogEntry(LogType._Error, "{0}", "     Cannot perform copy : " & ex.Message.ToString)
                End Try

                'MultiGame
                Orig = gHS.Systems(OrigSystem).Roms(OrigRom).MultiGamePath
                If Orig.EndsWith("\") = True Then
                    Orig = Directory.GetParent(Orig).FullName
                End If
                Dest = gHS.Systems(DestSystem).Roms(DestRom).MultiGamePath
                Try
                    If Directory.Exists(Orig) Then
                        LogEntry(LogType._Info, "{0}", "     Source : " & Orig & "  ---> Dest : " & Dest)
                        My.Computer.FileSystem.CopyDirectory(Orig, Dest, Overwrite)
                        Application.DoEvents()
                        If MoveMedia = True Then
                            Directory.Delete(Orig, True)
                        End If
                    End If
                Catch ex As Exception
                    LogEntry(LogType._Error, "{0}", "     Cannot perform copy : " & ex.Message.ToString)
                End Try

                'Music
                Orig = gHS.Systems(OrigSystem).Roms(OrigRom).MusicPath
                If Orig.EndsWith("\") = True Then
                    Orig = Directory.GetParent(Orig).FullName
                End If
                Dest = gHS.Systems(DestSystem).Roms(DestRom).MusicPath
                Try
                    If Directory.Exists(Orig) Then
                        LogEntry(LogType._Info, "{0}", "     Source : " & Orig & "  ---> Dest : " & Dest)
                        My.Computer.FileSystem.CopyDirectory(Orig, Dest, Overwrite)
                        Application.DoEvents()
                        If MoveMedia = True Then
                            Directory.Delete(Orig, True)
                        End If
                    End If
                Catch ex As Exception
                    LogEntry(LogType._Error, "{0}", "     Cannot perform copy : " & ex.Message.ToString)
                End Try

                'Sounds
                Orig = gHS.Systems(OrigSystem).Roms(OrigRom).SoundsPath
                If Orig.EndsWith("\") = True Then
                    Orig = Directory.GetParent(Orig).FullName
                End If
                Dest = gHS.Systems(DestSystem).Roms(DestRom).SoundsPath
                Try
                    If Directory.Exists(Orig) Then
                        LogEntry(LogType._Info, "{0}", "     Source : " & Orig & "  ---> Dest : " & Dest)
                        My.Computer.FileSystem.CopyDirectory(Orig, Dest, Overwrite)
                        Application.DoEvents()
                        If MoveMedia = True Then
                            Directory.Delete(Orig, True)
                        End If
                    End If
                Catch ex As Exception
                    LogEntry(LogType._Error, "{0}", "     Cannot perform copy : " & ex.Message.ToString)
                End Try

                'Videos
                Orig = gHS.Systems(OrigSystem).Roms(OrigRom).HLVideosPath
                If Orig.EndsWith("\") = True Then
                    Orig = Directory.GetParent(Orig).FullName
                End If
                Dest = gHS.Systems(DestSystem).Roms(DestRom).HLVideosPath
                Try
                    If Directory.Exists(Orig) Then
                        LogEntry(LogType._Info, "{0}", "     Source : " & Orig & "  ---> Dest : " & Dest)
                        My.Computer.FileSystem.CopyDirectory(Orig, Dest, Overwrite)
                        Application.DoEvents()
                        If MoveMedia = True Then
                            Directory.Delete(Orig, True)
                        End If
                    End If
                Catch ex As Exception
                    LogEntry(LogType._Error, "{0}", "     Cannot perform copy : " & ex.Message.ToString)
                End Try

                'Wheels
                Orig = gHS.Systems(OrigSystem).Roms(OrigRom).HLWheelsPath
                If Orig.EndsWith("\") = True Then
                    Orig = Directory.GetParent(Orig).FullName
                End If
                Dest = gHS.Systems(DestSystem).Roms(DestRom).HLWheelsPath
                Try
                    If Directory.Exists(Orig) Then
                        LogEntry(LogType._Info, "{0}", "     Source : " & Orig & "  ---> Dest : " & Dest)
                        My.Computer.FileSystem.CopyDirectory(Orig, Dest, Overwrite)
                        Application.DoEvents()
                        If MoveMedia = True Then
                            Directory.Delete(Orig, True)
                        End If
                    End If
                Catch ex As Exception
                    LogEntry(LogType._Error, "{0}", "     Cannot perform copy : " & ex.Message.ToString)
                End Try

            End If
        End If

        Form1.Show()
        Form1.Focus()
        Me.Close()
        Me.Dispose()

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Form1.Show()
        Form1.Focus()
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub ComboBox1_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedValueChanged
        DestSystem = ComboBox1.SelectedItem.ToString

        Cursor = Cursors.WaitCursor
        Try
            If gHS.Systems(DestSystem).Roms.Count < 1 Then
                gCheckRoms = False 'on force à ne pas vérifier les roms
                gHS.Systems(DestSystem).RefreshSystem(True)
            End If
        Catch ex As Exception
            MsgBox("There has been an issue loading roms for the selected system." & vbCrLf & "Please ensure System is valid", MsgBoxStyle.Critical, "Error")
        Finally
            gCheckRoms = Form1.CheckBox3.Checked 'on remet le check des roms à la valeur d'origine (définie dans le Form1.CheckBox3)
        End Try

        'load roms from selected system
        ComboBox2.Items.Clear()
        Try
            'For Each vItem In gHS.Systems(DestSystem).ReturnRomNames()
            '    ComboBox2.Items.Add(vItem)
            'Next
            For Each Entry In gHS.Systems(DestSystem).Roms
                ComboBox2.Items.Add(Entry.Name)
            Next
        Catch ex As Exception
            LogEntry(LogType._Warning, "{0}", "Problem while listing roms for system " & DestSystem & " : " & ex.Message.ToString)
        End Try
        Cursor = Cursors.Default
    End Sub

    Private Sub ComboBox2_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedValueChanged
        DestRom = ComboBox2.SelectedItem.ToString
    End Sub
End Class