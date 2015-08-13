Imports System.IO

Public Class FrmRenameRom
    Private pSystem As String
    Private pRom As String
    Private pRowID As Integer

    Public Sub New(ByVal System As String, ByVal RowID As Integer)
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.

        pSystem = System
        pRom = Form1.DataGridView4.Rows(RowID).Cells(0).Value.ToString
        pRowID = RowID

        'Default config
        Label1.Text = "Renaming [" & pSystem & "]/[" & pRom & "]"
        RadioButton1.Checked = True
        RadioButton2.Checked = False
        RadioButton3.Checked = False

        RadioButton4.Checked = False
        RadioButton5.Checked = True

        CheckBox1.Checked = True
        CheckBox2.Checked = True

        CheckBox3.Checked = True

        Button3.Enabled = False
        Button3.Visible = False
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If pRom = TextBox1.Text Then
                MsgBox("Nothing to do -> same rom !", MsgBoxStyle.Information, "St00pid?")
            Else
                RenameRom()
            End If
        Catch ex As Exception
            LogEntry(LogType._Error, "{0}", ex.Message.ToString)
        End Try
    End Sub

    Private Sub RenameRom()
        Dim RenameRomFile As Boolean = CheckBox3.Checked
        Dim NewName As String = TextBox1.Text
        Dim CopyMedia As Boolean = RadioButton1.Checked
        Dim MoveMedia As Boolean = RadioButton2.Checked
        'Dim CopyNothing As Boolean = RadioButton3.Checked
        Dim IncludeHyperSpinMedia As Boolean = CheckBox1.Checked
        Dim IncludeHyperPauseMedia As Boolean = CheckBox2.Checked
        Dim Overwrite As Boolean = CheckBox4.Checked

        'check if name is not nothing
        If NewName = "" Then
            MsgBox("Please enter a new name !", MsgBoxStyle.Information, "Oops")
            Exit Sub
        End If

        'if RenameRomFile=true & rom already existing : ask question to overwrite or not
        If RenameRomFile = True And File.Exists(gHS.Systems(pSystem).mRomPath & "\" & NewName & Path.GetExtension(gHS.Systems(pSystem).Roms(pRom).RomPath)) = True Then
            If MsgBox("Ooops ! Looks like the file you attempt to rename is already existing ..." & vbCrLf & "OVERWRITE IT ?", MsgBoxStyle.YesNo, "Warning") = MsgBoxResult.No Then
                Exit Sub
            End If
        End If

        If MsgBox("Are you sure ? XML will be saved after this operation.", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
            Try
                'change XML
                Form1.DataGridView4.Rows(pRowID).Cells(0).Value = NewName
            Catch ex As Exception
                LogEntry(LogType._Error, "{0}", "Cannot change the rom name in the list ... (pls contact the Dev)")
                Exit Sub
            End Try

            'renameRomFile
            If RenameRomFile = True Then
                Try
                    FileSystem.Rename(gHS.Systems(pSystem).Roms(pRom).RomPath, gHS.Systems(pSystem).mRomPath & "\" & NewName & Path.GetExtension(gHS.Systems(pSystem).Roms(pRom).RomPath))
                    LogEntry(LogType._Info, "{0}", "Rom renamed from " & gHS.Systems(pSystem).Roms(pRom).RomPath & " -to-> " & gHS.Systems(pSystem).mRomPath & "\" & NewName & Path.GetExtension(gHS.Systems(pSystem).Roms(pRom).RomPath))
                Catch ex As Exception
                    MsgBox("Error ! : cannot rename romfile :( --> abording", MsgBoxStyle.Critical, "Error")
                    LogEntry(LogType._Error, "{0}", "Error ! : cannot rename romfile :( --> abording")
                    Exit Sub
                End Try
            End If

            'copy data or not
            If CopyMedia = True Or MoveMedia = True Then
                Dim Orig As String = ""
                Dim Dest As String = ""

                If IncludeHyperSpinMedia = True Then
                    'then copy theme, wheel, video, art1, art2, art3, art4
                    LogEntry(LogType._Info, "{0}", "Copy started (overwrite=" & Overwrite.ToString & ") --> ")

                    'Theme
                    Orig = gHS.Systems(pSystem).Roms(pRom).ThemePath
                    Dest = Path.GetDirectoryName(gHS.Systems(pSystem).Roms(pRom).ThemePath) & "\" & NewName & ".zip"
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
                    Orig = gHS.Systems(pSystem).Roms(pRom).VideoPath
                    Dest = Path.GetDirectoryName(gHS.Systems(pSystem).Roms(pRom).VideoPath) & "\" & NewName & Path.GetExtension(gHS.Systems(pSystem).Roms(pRom).VideoPath)
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
                    Orig = gHS.Systems(pSystem).Roms(pRom).WheelPath
                    Dest = Path.GetDirectoryName(gHS.Systems(pSystem).Roms(pRom).WheelPath) & "\" & NewName & Path.GetExtension(gHS.Systems(pSystem).Roms(pRom).WheelPath)
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
                    Orig = gHS.Systems(pSystem).Roms(pRom).Art1Path
                    Dest = Path.GetDirectoryName(gHS.Systems(pSystem).Roms(pRom).Art1Path) & "\" & NewName & Path.GetExtension(gHS.Systems(pSystem).Roms(pRom).Art1Path)
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
                    Orig = gHS.Systems(pSystem).Roms(pRom).Art2Path
                    Dest = Path.GetDirectoryName(gHS.Systems(pSystem).Roms(pRom).Art2Path) & "\" & NewName & Path.GetExtension(gHS.Systems(pSystem).Roms(pRom).Art2Path)
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
                    Orig = gHS.Systems(pSystem).Roms(pRom).Art3Path
                    Dest = Path.GetDirectoryName(gHS.Systems(pSystem).Roms(pRom).Art3Path) & "\" & NewName & Path.GetExtension(gHS.Systems(pSystem).Roms(pRom).Art3Path)
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
                    Orig = gHS.Systems(pSystem).Roms(pRom).Art4Path
                    Dest = Path.GetDirectoryName(gHS.Systems(pSystem).Roms(pRom).Art4Path) & "\" & NewName & Path.GetExtension(gHS.Systems(pSystem).Roms(pRom).Art4Path)
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
                    Orig = gHS.Systems(pSystem).Roms(pRom).ArtworkPath
                    If Orig.EndsWith("\") = True Then
                        Orig = Directory.GetParent(Orig).FullName
                    End If
                    Dest = Directory.GetParent(Orig).FullName & "\" & NewName
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
                    Orig = gHS.Systems(pSystem).Roms(pRom).BackgroundsPath
                    If Orig.EndsWith("\") = True Then
                        Orig = Directory.GetParent(Orig).FullName
                    End If
                    Dest = Directory.GetParent(Orig).FullName & "\" & NewName
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
                    Orig = gHS.Systems(pSystem).Roms(pRom).BezelsPath
                    If Orig.EndsWith("\") = True Then
                        Orig = Directory.GetParent(Orig).FullName
                    End If
                    Dest = Directory.GetParent(Orig).FullName & "\" & NewName
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
                    Orig = gHS.Systems(pSystem).Roms(pRom).ControllerPath
                    If Orig.EndsWith("\") = True Then
                        Orig = Directory.GetParent(Orig).FullName
                    End If
                    Dest = Directory.GetParent(Orig).FullName & "\" & NewName
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
                    Orig = gHS.Systems(pSystem).Roms(pRom).FadePath
                    If Orig.EndsWith("\") = True Then
                        Orig = Directory.GetParent(Orig).FullName
                    End If
                    Dest = Directory.GetParent(Orig).FullName & "\" & NewName
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
                    Orig = gHS.Systems(pSystem).Roms(pRom).FontsPath
                    If Orig.EndsWith("\") = True Then
                        Orig = Directory.GetParent(Orig).FullName
                    End If
                    Dest = Directory.GetParent(Orig).FullName & "\" & NewName
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
                    Orig = gHS.Systems(pSystem).Roms(pRom).GuidesPath
                    If Orig.EndsWith("\") = True Then
                        Orig = Directory.GetParent(Orig).FullName
                    End If
                    Dest = Directory.GetParent(Orig).FullName & "\" & NewName
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
                    Orig = gHS.Systems(pSystem).Roms(pRom).ManualsPath
                    If Orig.EndsWith("\") = True Then
                        Orig = Directory.GetParent(Orig).FullName
                    End If
                    Dest = Directory.GetParent(Orig).FullName & "\" & NewName
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
                    Orig = gHS.Systems(pSystem).Roms(pRom).MenuImagesPath
                    If Orig.EndsWith("\") = True Then
                        Orig = Directory.GetParent(Orig).FullName
                    End If
                    Dest = Directory.GetParent(Orig).FullName & "\" & NewName
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
                    Orig = gHS.Systems(pSystem).Roms(pRom).MultiGamePath
                    If Orig.EndsWith("\") = True Then
                        Orig = Directory.GetParent(Orig).FullName
                    End If
                    Dest = Directory.GetParent(Orig).FullName & "\" & NewName
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
                    Orig = gHS.Systems(pSystem).Roms(pRom).MusicPath
                    If Orig.EndsWith("\") = True Then
                        Orig = Directory.GetParent(Orig).FullName
                    End If
                    Dest = Directory.GetParent(Orig).FullName & "\" & NewName
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
                    Orig = gHS.Systems(pSystem).Roms(pRom).SoundsPath
                    If Orig.EndsWith("\") = True Then
                        Orig = Directory.GetParent(Orig).FullName
                    End If
                    Dest = Directory.GetParent(Orig).FullName & "\" & NewName
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
                    Orig = gHS.Systems(pSystem).Roms(pRom).HLVideosPath
                    If Orig.EndsWith("\") = True Then
                        Orig = Directory.GetParent(Orig).FullName
                    End If
                    Dest = Directory.GetParent(Orig).FullName & "\" & NewName
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
                    Orig = gHS.Systems(pSystem).Roms(pRom).HLWheelsPath
                    If Orig.EndsWith("\") = True Then
                        Orig = Directory.GetParent(Orig).FullName
                    End If
                    Dest = Directory.GetParent(Orig).FullName & "\" & NewName
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

            'save XML
            Form1.ToolStripButton6.PerformClick()
            Form1.Show()
            Form1.Focus()
            Me.Close()
            Me.Dispose()
        Else
            Exit Sub
        End If

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Form1.Show()
        Form1.Focus()
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub RadioButton4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton4.CheckedChanged
        If RadioButton4.Checked = True Then
            Button3.Visible = True
            Button3.Enabled = True
            TextBox1.Text = ""
            CheckBox3.Checked = False
            'CheckBox3.Enabled = False
        Else
            Button3.Visible = False
            Button3.Enabled = False
            TextBox1.Text = "[enter new name]"
            CheckBox3.Checked = True
            'CheckBox3.Enabled = True
        End If
    End Sub

    Private Sub RadioButton3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged
        If RadioButton3.Checked = True Then
            CheckBox1.Checked = False
            CheckBox1.Enabled = False
            CheckBox2.Checked = False
            CheckBox2.Enabled = False
        Else
            CheckBox1.Checked = True
            CheckBox1.Enabled = True
            CheckBox2.Checked = True
            CheckBox2.Enabled = True
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim openFileDialog1 As New OpenFileDialog()
        Dim DlgFilter As String = "Roms (*.*)|*.*"
        openFileDialog1.RestoreDirectory = True
        Try
            openFileDialog1.InitialDirectory = gHS.Systems(pSystem).mRomPath
        Catch ex As Exception
            '
        End Try
        openFileDialog1.Filter = DlgFilter
        openFileDialog1.FilterIndex = 1
        If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            TextBox1.Text = Path.GetFileNameWithoutExtension(openFileDialog1.FileName)
        Else
            TextBox1.Text = ""
        End If
    End Sub
End Class