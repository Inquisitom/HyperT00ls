Imports System.IO

Public Class FrmReport
    Private vCurrSystem As String

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Public Sub New(ByVal System As String)
        vCurrSystem = System
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Label1.Text = "System : " & vCurrSystem
        TreeView1.ExpandAll()
        TextBox1.Text = Environment.CurrentDirectory & "\Reports\"
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim GetReportsPath As New FolderBrowserDialog
        GetReportsPath.Description = "Reports Folder"
        GetReportsPath.RootFolder = System.Environment.SpecialFolder.MyComputer
        If GetReportsPath.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Try
                TextBox1.Text = GetReportsPath.SelectedPath
            Catch Ex As Exception

            End Try
        Else
            TextBox1.Text = Environment.CurrentDirectory & "\Reports\"
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim vOneFile As Boolean = RadioButton1.Checked
        Dim vGenHAVE As Boolean = CheckBox1.Checked
        Dim vGenMISSING As Boolean = CheckBox2.Checked
        'vCurrSystem
        Dim vOutPath As String = TextBox1.Text

        Try
            'First, check if folder exists, or create it
            If Directory.Exists(vOutPath) = False Then
                Directory.CreateDirectory(vOutPath)
            End If
        Catch ex As Exception
            MsgBox("Cannot create directory ! : " & vOutPath)
            LogEntry(LogType._Error, "{0}", "Cannot create directory : " & vOutPath & vbCrLf & ex.Message.ToString)
            Exit Sub
        End Try

        Cursor.Current = Cursors.WaitCursor
        If vOneFile = True Then
            'we create only one file
            'if already existing : file.delete it ! gniark :)
            Dim vOutFile As String = vOutPath & "\Report_" & Date.Now.ToString("yyyy-MM-dd_HH-mm-ss") & ".txt"
            If File.Exists(vOutFile) = True Then
                File.Delete(vOutFile)
            End If
            Application.DoEvents()
            ' Create a file to write to. 
            Dim sw As StreamWriter = File.CreateText(vOutFile)

            'Handling missing
            If vGenMISSING = True Then
                sw.WriteLine("---------------- REPORT MISSING for System " & vCurrSystem & " ----------------")
                sw.WriteLine("---------------- " & Date.Now.ToString("dd-MM-yyyy HH\hmm\mss\s"))
                'Handling missing type (wheel/theme/rom/etc...)
                For Each vNode As TreeNode In TreeView1.Nodes
                    If vNode.Text = "Rom" And vNode.Checked = True Then
                        sw.WriteLine(vbCrLf)
                        sw.WriteLine("-> MISSING : " & vNode.Text)
                        For Each vRow As DataGridViewRow In Form1.DataGridView1.Rows
                            Dim vCurrRom As String = vRow.Cells(0).Value.ToString
                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).RomFound = False Then
                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                            End If
                        Next
                    End If
                    If vNode.Nodes.Count > 0 Then
                        For Each childNode As TreeNode In vNode.Nodes
                            If childNode.Checked = True Then
                                sw.WriteLine(vbCrLf)
                                sw.WriteLine("-> MISSING : " & childNode.Text)
                                For Each vRow As DataGridViewRow In Form1.DataGridView1.Rows
                                    Dim vCurrRom As String = vRow.Cells(0).Value.ToString
                                    Select Case childNode.Text
                                        Case "Wheel"
                                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).HasWheel = False Then
                                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                                            End If
                                        Case "Theme"
                                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).HasTheme = False Then
                                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                                            End If
                                        Case "Video"
                                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).HasVideo = False Then
                                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                                            End If
                                        Case "Art1"
                                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).HasArt1 = False Then
                                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                                            End If
                                        Case "Art2"
                                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).HasArt2 = False Then
                                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                                            End If
                                        Case "Art3"
                                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).HasArt3 = False Then
                                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                                            End If
                                        Case "Art4"
                                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).HasArt4 = False Then
                                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                                            End If

                                        Case "Artwork"
                                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).HasArtwork = False Then
                                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                                            End If
                                        Case "Background"
                                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).HasBackgrounds = False Then
                                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                                            End If
                                        Case "Bezel"
                                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).HasBezels = False Then
                                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                                            End If
                                        Case "Controller"
                                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).HasController = False Then
                                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                                            End If
                                        Case "Fade"
                                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).HasFade = False Then
                                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                                            End If
                                        Case "Fonts"
                                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).HasFonts = False Then
                                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                                            End If
                                        Case "Guides"
                                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).HasGuides = False Then
                                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                                            End If
                                        Case "Manuals"
                                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).HasManuals = False Then
                                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                                            End If
                                        Case "Menu Images"
                                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).HasMenuImages = False Then
                                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                                            End If
                                        Case "Multigame"
                                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).HasMultiGame = False Then
                                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                                            End If
                                        Case "Music"
                                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).HasMusic = False Then
                                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                                            End If
                                        Case "Sounds"
                                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).HasSounds = False Then
                                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                                            End If
                                        Case "Videos"
                                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).HasHLVideos = False Then
                                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                                            End If
                                        Case "Wheels"
                                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).HasHLWheels = False Then
                                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                                            End If
                                    End Select
                                Next
                            End If
                        Next
                    End If
                Next
                sw.WriteLine(vbCrLf)
                sw.WriteLine(vbCrLf)
            End If
            If vGenHAVE = True Then
                sw.WriteLine("---------------- REPORT HAVE for System " & vCurrSystem & " ----------------")
                sw.WriteLine("---------------- " & Date.Now.ToString("dd-MM-yyyy HH\hmm\mss\s"))
                'Handling have type (wheel/theme/rom/etc...)
                For Each vNode As TreeNode In TreeView1.Nodes
                    If vNode.Text = "Rom" And vNode.Checked = True Then
                        sw.WriteLine(vbCrLf)
                        sw.WriteLine("-> HAVE : " & vNode.Text)
                        For Each vRow As DataGridViewRow In Form1.DataGridView1.Rows
                            Dim vCurrRom As String = vRow.Cells(0).Value.ToString
                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).RomFound = True Then
                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                            End If
                        Next
                    End If
                    If vNode.Nodes.Count > 0 Then
                        For Each childNode As TreeNode In vNode.Nodes
                            If childNode.Checked = True Then
                                sw.WriteLine(vbCrLf)
                                sw.WriteLine("-> HAVE : " & childNode.Text)
                                For Each vRow As DataGridViewRow In Form1.DataGridView1.Rows
                                    Dim vCurrRom As String = vRow.Cells(0).Value.ToString
                                    Select Case childNode.Text
                                        Case "Wheel"
                                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).HasWheel = True Then
                                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                                            End If
                                        Case "Theme"
                                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).HasTheme = True Then
                                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                                            End If
                                        Case "Video"
                                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).HasVideo = True Then
                                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                                            End If
                                        Case "Art1"
                                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).HasArt1 = True Then
                                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                                            End If
                                        Case "Art2"
                                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).HasArt2 = True Then
                                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                                            End If
                                        Case "Art3"
                                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).HasArt3 = True Then
                                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                                            End If
                                        Case "Art4"
                                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).HasArt4 = True Then
                                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                                            End If

                                        Case "Artwork"
                                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).HasArtwork = True Then
                                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                                            End If
                                        Case "Background"
                                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).HasBackgrounds = True Then
                                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                                            End If
                                        Case "Bezel"
                                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).HasBezels = True Then
                                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                                            End If
                                        Case "Controller"
                                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).HasController = True Then
                                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                                            End If
                                        Case "Fade"
                                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).HasFade = True Then
                                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                                            End If
                                        Case "Fonts"
                                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).HasFonts = True Then
                                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                                            End If
                                        Case "Guides"
                                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).HasGuides = True Then
                                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                                            End If
                                        Case "Manuals"
                                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).HasManuals = True Then
                                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                                            End If
                                        Case "Menu Images"
                                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).HasMenuImages = True Then
                                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                                            End If
                                        Case "Multigame"
                                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).HasMultiGame = True Then
                                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                                            End If
                                        Case "Music"
                                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).HasMusic = True Then
                                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                                            End If
                                        Case "Sounds"
                                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).HasSounds = True Then
                                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                                            End If
                                        Case "Videos"
                                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).HasHLVideos = True Then
                                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                                            End If
                                        Case "Wheels"
                                            If gHS.Systems(vCurrSystem).Roms(vCurrRom).HasHLWheels = True Then
                                                sw.WriteLine(vRow.Cells(0).Value.ToString)
                                            End If
                                    End Select
                                Next
                            End If
                        Next
                    End If
                Next
                sw.WriteLine(vbCrLf)
            End If
            sw.Flush()
            sw.Close()
            Process.Start(vOutFile)
        Else
            'Multiple files
            'will be handled later

        End If

        Cursor.Current = Cursors.Default
        Me.Close()
    End Sub

    Private Sub TreeView1_NodeMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeNodeMouseClickEventArgs) Handles TreeView1.NodeMouseClick
        Select Case e.Node.Text
            Case "Hyperspin", "HyperPause"
                For Each vNode As TreeNode In e.Node.Nodes
                    vNode.Checked = e.Node.Checked
                Next
        End Select
    End Sub
End Class