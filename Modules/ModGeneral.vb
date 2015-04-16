Imports System.IO

Module ModGeneral
    Public IsLoading As Boolean = False
    Public gCONFIGFILE As String = Application.StartupPath & "\config.ini"
    Public gINI As New IniFile
    Public gHSPath As String
    Public gHLPath As String
    Public gDBPath As String
    Public gMainMenuXML As String
    Public gDisplayAllSystems As Boolean = False
    Public gCheckRoms As Boolean = False
    Public gDebugLog As Boolean = False

    'Columns
    'HS
    Public gColDesc As Boolean = True
    Public gColHSWheel As Boolean = True
    Public gColHSVideo As Boolean = True
    Public gColHSTheme As Boolean = True
    Public gColHSArt1 As Boolean = True
    Public gColHSArt2 As Boolean = True
    Public gColHSArt3 As Boolean = True
    Public gColHSArt4 As Boolean = True
    'HP
    Public gColHPArtwork As Boolean = True
    Public gColHPBackground As Boolean = True
    Public gColHPBezel As Boolean = True
    Public gColHPController As Boolean = True
    Public gColHPFade As Boolean = True
    Public gColHPFonts As Boolean = True
    Public gColHPGuides As Boolean = True
    Public gColHPManuals As Boolean = True
    Public gColHPMenuImages As Boolean = True
    Public gColHPMultigame As Boolean = True
    Public gColHPMusic As Boolean = True
    Public gColHPSounds As Boolean = True
    Public gColHPVideos As Boolean = True
    Public gColHPWheels As Boolean = True


    Public gHS As clsHyperSpin

    Public Enum LogType
        _Error = 3
        _Warning = 2
        _Debug = 1
        _Info = 0
    End Enum

#Region "Configuration File"
    Public Sub LoadConfigFile()
#If DEBUG Then
        If File.Exists(gCONFIGFILE) Then
            'file.delete(CONFIGFILE)
        End If
#End If
        If File.Exists(gCONFIGFILE) Then
            gINI.Load(gCONFIGFILE)
            gHSPath = gINI.GetSection("MAIN").GetKey("HSPath").GetValue()
            If gHSPath.StartsWith("..") Then
                gHSPath = Application.StartupPath & "\..\" & gHSPath
            End If
            If gHSPath.StartsWith("\") Then
                gHSPath = Application.StartupPath & "\"
            End If
            gDBPath = gHSPath & "\Databases\"
            gMainMenuXML = gDBPath & "\Main Menu\Main Menu.xml"
            Try
                gDisplayAllSystems = CBool(gINI.GetSection("MAIN").GetKey("DisplayAllSystems").GetValue())
            Catch ex As Exception
                gDisplayAllSystems = False
                gINI.GetSection("MAIN").AddKey("DisplayAllSystems").SetValue(gDisplayAllSystems)
            End Try
            Try
                gDebugLog = CBool(gINI.GetSection("MAIN").GetKey("Debug").GetValue())
            Catch ex As Exception
                gDebugLog = False
                gINI.GetSection("MAIN").AddKey("Debug").SetValue(gDebugLog)
            End Try
            Try
                gCheckRoms = CBool(gINI.GetSection("MAIN").GetKey("CheckRoms").GetValue())
            Catch ex As Exception
                gDebugLog = False
                gINI.GetSection("MAIN").AddKey("CheckRoms").SetValue(gCheckRoms)
            End Try


            'gColDesc
            Try
                gColDesc = CBool(gINI.GetSection("MAIN").GetKey("gColDesc").GetValue())
            Catch ex As Exception
                gColDesc = True
                gINI.GetSection("MAIN").AddKey("gColDesc").SetValue(gColDesc)
            End Try
            'gColHSWheel
            Try
                gColHSWheel = CBool(gINI.GetSection("MAIN").GetKey("gColHSWheel").GetValue())
            Catch ex As Exception
                gColHSWheel = True
                gINI.GetSection("MAIN").AddKey("gColHSWheel").SetValue(gColHSWheel)
            End Try
            'gColHSVideo
            Try
                gColHSVideo = CBool(gINI.GetSection("MAIN").GetKey("gColHSVideo").GetValue())
            Catch ex As Exception
                gColHSVideo = True
                gINI.GetSection("MAIN").AddKey("gColHSVideo").SetValue(gColHSVideo)
            End Try
            'gColHSTheme
            Try
                gColHSTheme = CBool(gINI.GetSection("MAIN").GetKey("gColHSTheme").GetValue())
            Catch ex As Exception
                gColHSTheme = True
                gINI.GetSection("MAIN").AddKey("gColHSTheme").SetValue(gColHSTheme)
            End Try
            'gColHSArt1
            Try
                gColHSArt1 = CBool(gINI.GetSection("MAIN").GetKey("gColHSArt1").GetValue())
            Catch ex As Exception
                gColHSArt1 = True
                gINI.GetSection("MAIN").AddKey("gColHSArt1").SetValue(gColHSArt1)
            End Try
            'gColHSArt2
            Try
                gColHSArt2 = CBool(gINI.GetSection("MAIN").GetKey("gColHSArt2").GetValue())
            Catch ex As Exception
                gColHSArt2 = True
                gINI.GetSection("MAIN").AddKey("gColHSArt2").SetValue(gColHSArt2)
            End Try
            'gColHSArt3
            Try
                gColHSArt3 = CBool(gINI.GetSection("MAIN").GetKey("gColHSArt3").GetValue())
            Catch ex As Exception
                gColHSArt3 = True
                gINI.GetSection("MAIN").AddKey("gColHSArt3").SetValue(gColHSArt3)
            End Try
            'gColHSArt4
            Try
                gColHSArt4 = CBool(gINI.GetSection("MAIN").GetKey("gColHSArt4").GetValue())
            Catch ex As Exception
                gColHSArt4 = True
                gINI.GetSection("MAIN").AddKey("gColHSArt4").SetValue(gColHSArt4)
            End Try

            'gColHPArtwork
            Try
                gColHPArtwork = CBool(gINI.GetSection("MAIN").GetKey("gColHPArtwork").GetValue())
            Catch ex As Exception
                gColHPArtwork = True
                gINI.GetSection("MAIN").AddKey("gColHPArtwork").SetValue(gColHPArtwork)
            End Try

            'gColHPBackground
            Try
                gColHPBackground = CBool(gINI.GetSection("MAIN").GetKey("gColHPBackground").GetValue())
            Catch ex As Exception
                gColHPBackground = True
                gINI.GetSection("MAIN").AddKey("gColHPBackground").SetValue(gColHPBackground)
            End Try

            'gColHPBezel
            Try
                gColHPBezel = CBool(gINI.GetSection("MAIN").GetKey("gColHPBezel").GetValue())
            Catch ex As Exception
                gColHPBezel = True
                gINI.GetSection("MAIN").AddKey("gColHPBezel").SetValue(gColHPBezel)
            End Try

            'gColHPController
            Try
                gColHPController = CBool(gINI.GetSection("MAIN").GetKey("gColHPController").GetValue())
            Catch ex As Exception
                gColHPController = True
                gINI.GetSection("MAIN").AddKey("gColHPController").SetValue(gColHPController)
            End Try

            'gColHPFade
            Try
                gColHPFade = CBool(gINI.GetSection("MAIN").GetKey("gColHPFade").GetValue())
            Catch ex As Exception
                gColHPFade = True
                gINI.GetSection("MAIN").AddKey("gColHPFade").SetValue(gColHPFade)
            End Try

            'gColHPFonts
            Try
                gColHPFonts = CBool(gINI.GetSection("MAIN").GetKey("gColHPFonts").GetValue())
            Catch ex As Exception
                gColHPFonts = True
                gINI.GetSection("MAIN").AddKey("gColHPFonts").SetValue(gColHPFonts)
            End Try

            'gColHPGuides
            Try
                gColHPGuides = CBool(gINI.GetSection("MAIN").GetKey("gColHPGuides").GetValue())
            Catch ex As Exception
                gColHPGuides = True
                gINI.GetSection("MAIN").AddKey("gColHPGuides").SetValue(gColHPGuides)
            End Try

            'gColHPManuals
            Try
                gColHPManuals = CBool(gINI.GetSection("MAIN").GetKey("gColHPManuals").GetValue())
            Catch ex As Exception
                gColHPManuals = True
                gINI.GetSection("MAIN").AddKey("gColHPManuals").SetValue(gColHPManuals)
            End Try

            'gColHPMenuImages
            Try
                gColHPMenuImages = CBool(gINI.GetSection("MAIN").GetKey("gColHPMenuImages").GetValue())
            Catch ex As Exception
                gColHPMenuImages = True
                gINI.GetSection("MAIN").AddKey("gColHPMenuImages").SetValue(gColHPMenuImages)
            End Try

            'gColHPMultigame
            Try
                gColHPMultigame = CBool(gINI.GetSection("MAIN").GetKey("gColHPMultigame").GetValue())
            Catch ex As Exception
                gColHPMultigame = True
                gINI.GetSection("MAIN").AddKey("gColHPMultigame").SetValue(gColHPMultigame)
            End Try

            'gColHPMusic
            Try
                gColHPMusic = CBool(gINI.GetSection("MAIN").GetKey("gColHPMusic").GetValue())
            Catch ex As Exception
                gColHPMusic = True
                gINI.GetSection("MAIN").AddKey("gColHPMusic").SetValue(gColHPMusic)
            End Try

            'gColHPSounds
            Try
                gColHPSounds = CBool(gINI.GetSection("MAIN").GetKey("gColHPSounds").GetValue())
            Catch ex As Exception
                gColHPSounds = True
                gINI.GetSection("MAIN").AddKey("gColHPSounds").SetValue(gColHPSounds)
            End Try

            'gColHPVideos
            Try
                gColHPVideos = CBool(gINI.GetSection("MAIN").GetKey("gColHPVideos").GetValue())
            Catch ex As Exception
                gColHPVideos = True
                gINI.GetSection("MAIN").AddKey("gColHPVideos").SetValue(gColHPVideos)
            End Try

            'gColHPWheels
            Try
                gColHPWheels = CBool(gINI.GetSection("MAIN").GetKey("gColHPWheels").GetValue())
            Catch ex As Exception
                gColHPWheels = True
                gINI.GetSection("MAIN").AddKey("gColHPWheels").SetValue(gColHPWheels)
            End Try

            Form1.TabControl1.SelectedIndex = 0
        Else
            Form1.TabControl1.SelectedIndex = 7
            gINI.AddSection("MAIN")
            Dim GetHSPath As New FolderBrowserDialog
            GetHSPath.Description = "Hyperspin Folder"
            GetHSPath.RootFolder = Environment.SpecialFolder.MyComputer
            If GetHSPath.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                Try
                    gHSPath = GetHSPath.SelectedPath
                    If gHSPath.EndsWith("\") = False Then
                        gHSPath = gHSPath & "\"
                    End If
                    gDBPath = gHSPath & "Databases\"
                    gMainMenuXML = gDBPath & "\Main Menu\Main Menu.xml"
                Catch Ex As Exception
                Finally
                    SaveHyperT00lsParameters()
                End Try
            Else
                gHSPath = ""
                SaveHyperT00lsParameters()
            End If
        End If

        Form1.CheckBox1.Checked = gDisplayAllSystems
        Form1.CheckBox2.Checked = gDebugLog
        Form1.CheckBox3.Checked = gCheckRoms

        'Desc
        Form1.TreeView2.Nodes("Node0").Checked = gColDesc

        'HS
        For Each vNode As TreeNode In Form1.TreeView2.Nodes("Node1").Nodes
            Select Case vNode.Text
                Case "Description"
                    vNode.Checked = gColDesc
                Case "Wheel"
                    vNode.Checked = gColHSWheel
                Case "Video"
                    vNode.Checked = gColHSVideo
                Case "Theme"
                    vNode.Checked = gColHSTheme
                Case "Art1"
                    vNode.Checked = gColHSArt1
                Case "Art2"
                    vNode.Checked = gColHSArt2
                Case "Art3"
                    vNode.Checked = gColHSArt3
                Case "Art4"
                    vNode.Checked = gColHSArt4
            End Select
        Next

        'HP
        For Each vNode As TreeNode In Form1.TreeView2.Nodes("Node10").Nodes
            Select Case vNode.Text
                Case "Artwork"
                    vNode.Checked = gColHPArtwork
                Case "Background"
                    vNode.Checked = gColHPBackground
                Case "Bezel"
                    vNode.Checked = gColHPBezel
                Case "Controller"
                    vNode.Checked = gColHPController
                Case "Fade"
                    vNode.Checked = gColHPFade
                Case "Fonts"
                    vNode.Checked = gColHPFonts
                Case "Guides"
                    vNode.Checked = gColHPGuides
                Case "Manuals"
                    vNode.Checked = gColHPManuals
                Case "Menu Images"
                    vNode.Checked = gColHPMenuImages
                Case "Multigame"
                    vNode.Checked = gColHPMultigame
                Case "Music"
                    vNode.Checked = gColHPMusic
                Case "Sounds"
                    vNode.Checked = gColHPSounds
                Case "Videos"
                    vNode.Checked = gColHPVideos
                Case "Wheels"
                    vNode.Checked = gColHPWheels
            End Select
        Next

        'Check entry in HL\Settings\settings.ini for HL path
        '[Main]
        'Hyperlaunch_Path=D:\Hyperspin\HyperLaunch\
        Dim INIHL As New IniFile
        Try
            LogEntry(LogType._Info, "Loading " & gHSPath & "\Settings\Settings.ini")
            INIHL.Load(gHSPath & "\Settings\Settings.ini")
            gHLPath = INIHL.GetSection("Main").AddKey("Hyperlaunch_Path").GetValue()
            If gHLPath.StartsWith("\") Then
                If gHSPath.EndsWith("\") Then
                    gHLPath = gHSPath & "..\" & gHLPath
                Else
                    gHLPath = gHSPath & "\..\" & gHLPath
                End If
            End If
            If gHLPath.EndsWith("\") Then
                gHLPath = Path.GetDirectoryName(gHLPath)
            End If
            If gHLPath.StartsWith("..") Then
                gHLPath = gHLPath.Replace("..", gHSPath)
            End If
            LogEntry(LogType._Info, "Finished loading configuration file for HyperLaunch.")
            Console.WriteLine("Load OK !")
        Catch ex As Exception
            LogEntry(LogType._Error, "Cannot load configuration file for HyperLaunch !")
            Console.WriteLine("Load config file ERROR !")
        Finally
            INIHL = Nothing
        End Try
    End Sub

    Public Sub SaveHyperT00lsParameters()
        If File.Exists(gCONFIGFILE) Then
            File.Delete(gCONFIGFILE)
        End If
        Application.DoEvents()

        'Desc
        gColDesc = Form1.TreeView2.Nodes("Node0").Checked
        'HS
        For Each vNode As TreeNode In Form1.TreeView2.Nodes("Node1").Nodes
            Select Case vNode.Text
                Case "Description"
                    gColDesc = vNode.Checked
                Case "Wheel"
                    gColHSWheel = vNode.Checked
                Case "Video"
                    gColHSVideo = vNode.Checked
                Case "Theme"
                    gColHSTheme = vNode.Checked
                Case "Art1"
                    gColHSArt1 = vNode.Checked
                Case "Art2"
                    gColHSArt2 = vNode.Checked
                Case "Art3"
                    gColHSArt3 = vNode.Checked
                Case "Art4"
                    gColHSArt4 = vNode.Checked
            End Select
        Next

        'HP 
        For Each vNode As TreeNode In Form1.TreeView2.Nodes("Node10").Nodes
            Select Case vNode.Text
                Case "Artwork"
                    gColHPArtwork = vNode.Checked
                Case "Background"
                    gColHPBackground = vNode.Checked
                Case "Bezel"
                    gColHPBezel = vNode.Checked
                Case "Controller"
                    gColHPController = vNode.Checked
                Case "Fade"
                    gColHPFade = vNode.Checked
                Case "Fonts"
                    gColHPFonts = vNode.Checked
                Case "Guides"
                    gColHPGuides = vNode.Checked
                Case "Manuals"
                    gColHPManuals = vNode.Checked
                Case "Menu Images"
                    gColHPMenuImages = vNode.Checked
                Case "Multigame"
                    gColHPMultigame = vNode.Checked
                Case "Music"
                    gColHPMusic = vNode.Checked
                Case "Sounds"
                    gColHPSounds = vNode.Checked
                Case "Videos"
                    gColHPVideos = vNode.Checked
                Case "Wheels"
                    gColHPWheels = vNode.Checked
            End Select
        Next

        gINI.AddSection("MAIN")
        gINI.GetSection("MAIN").AddKey("HSPath").SetValue(gHSPath)
        gINI.GetSection("MAIN").AddKey("DisplayAllSystems").SetValue(gDisplayAllSystems)
        gINI.GetSection("MAIN").AddKey("Debug").SetValue(gDebugLog)
        gINI.GetSection("MAIN").AddKey("CheckRoms").SetValue(gCheckRoms)
        gINI.GetSection("MAIN").AddKey("gColDesc").SetValue(gColDesc)
        gINI.GetSection("MAIN").AddKey("gColHSWheel").SetValue(gColHSWheel)
        gINI.GetSection("MAIN").AddKey("gColHSVideo").SetValue(gColHSVideo)
        gINI.GetSection("MAIN").AddKey("gColHSTheme").SetValue(gColHSTheme)
        gINI.GetSection("MAIN").AddKey("gColHSArt1").SetValue(gColHSArt1)
        gINI.GetSection("MAIN").AddKey("gColHSArt2").SetValue(gColHSArt2)
        gINI.GetSection("MAIN").AddKey("gColHSArt3").SetValue(gColHSArt3)
        gINI.GetSection("MAIN").AddKey("gColHSArt4").SetValue(gColHSArt4)
        gINI.GetSection("MAIN").AddKey("gColHPArtwork").SetValue(gColHPArtwork)
        gINI.GetSection("MAIN").AddKey("gColHPBackground").SetValue(gColHPBackground)
        gINI.GetSection("MAIN").AddKey("gColHPBezel").SetValue(gColHPBezel)
        gINI.GetSection("MAIN").AddKey("gColHPController").SetValue(gColHPController)
        gINI.GetSection("MAIN").AddKey("gColHPFade").SetValue(gColHPFade)
        gINI.GetSection("MAIN").AddKey("gColHPFonts").SetValue(gColHPFonts)
        gINI.GetSection("MAIN").AddKey("gColHPGuides").SetValue(gColHPGuides)
        gINI.GetSection("MAIN").AddKey("gColHPManuals").SetValue(gColHPManuals)
        gINI.GetSection("MAIN").AddKey("gColHPMenuImages").SetValue(gColHPMenuImages)
        gINI.GetSection("MAIN").AddKey("gColHPMultigame").SetValue(gColHPMultigame)
        gINI.GetSection("MAIN").AddKey("gColHPMusic").SetValue(gColHPMusic)
        gINI.GetSection("MAIN").AddKey("gColHPSounds").SetValue(gColHPSounds)
        gINI.GetSection("MAIN").AddKey("gColHPVideos").SetValue(gColHPVideos)
        gINI.GetSection("MAIN").AddKey("gColHPWheels").SetValue(gColHPWheels)
        gINI.Save(gCONFIGFILE)
        LogEntry(LogType._Info, "Config saved ...")
    End Sub

    Public Sub LogEntry(ByVal Type As LogType, ByVal Message As String)
        If gDebugLog = True Then
            Dim vDate As String
            Dim vHour As String
            Dim vMessage As String = Message
            Dim vType As String
            Dim time As DateTime = DateTime.Now
            Dim n As Integer = Form1.DataGridView3.Rows.Add()
            vDate = time.ToString("yyyy-MM-dd")
            vHour = time.ToString("HH:mm")

            Select Case Type
                Case LogType._Debug
                    vType = "Debug"
                    Form1.DataGridView3.Rows(n).DefaultCellStyle.BackColor = Color.White
                Case (LogType._Error)
                    vType = "Error"
                    Form1.DataGridView3.Rows(n).DefaultCellStyle.BackColor = Color.IndianRed
                Case LogType._Info
                    vType = "Info"
                    Form1.DataGridView3.Rows(n).DefaultCellStyle.BackColor = Color.LightGreen
                Case LogType._Warning
                    vType = "Warning"
                    Form1.DataGridView3.Rows(n).DefaultCellStyle.BackColor = Color.MediumOrchid
                Case Else
                    vType = ""
                    Form1.DataGridView3.Rows(n).DefaultCellStyle.BackColor = Color.White
            End Select

            Form1.DataGridView3.Rows(n).Cells(0).Value = vDate
            Form1.DataGridView3.Rows(n).Cells(1).Value = vHour
            Form1.DataGridView3.Rows(n).Cells(2).Value = vType
            Form1.DataGridView3.Rows(n).Cells(3).Value = vMessage
        End If
    End Sub
#End Region
End Module
