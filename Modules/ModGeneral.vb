Imports System.IO

Module ModGeneral
    Public IsLoading As Boolean = False
    Public gCONFIGFILE As String = Application.StartupPath & "\config.ini"
    Public gDebugLogFile As String = Application.StartupPath & "\HyperT00ls.log"
    Public gDebugDisplayLevel As LogType = LogType._Warning
    Public gINI As New IniFile
    Public gHSPath As String
    Public gHSOldVersion As Boolean = True
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
    Public gColEnabled As Boolean = True

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
            'gColEnabled
            Try
                gColEnabled = CBool(gINI.GetSection("MAIN").GetKey("gColEnabled").GetValue())
            Catch ex As Exception
                gColEnabled = True
                gINI.GetSection("MAIN").AddKey("gColEnabled").SetValue(gColEnabled)
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

        'Enabled
        Form1.TreeView2.Nodes("NodeEnabled").Checked = gColEnabled

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

        'HL/RocketLauncher
        Try
            gHLPath = gINI.GetSection("MAIN").GetKey("RocketLauncherPath").GetValue()
            If gHLPath = "" Then
                gHLPath = GetHLPath()
            Else
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
            End If
            If Path.HasExtension(gHLPath) = True Then
                gHLPath = Path.GetDirectoryName(gHLPath)
            End If
        Catch ex As Exception
            gHLPath = GetHLPath()
        Finally
            SaveHyperT00lsParameters()
        End Try
        Form1.TextBox3.Text = gHLPath
    End Sub

    Public Function GetHLPath() As String
        GetHLPath = ""
        Dim INIHL As New IniFile
        Try
            LogEntry(LogType._Info, "{0}", "Loading " & gHSPath & "\Settings\Settings.ini")
            INIHL.Load(gHSPath & "\Settings\Settings.ini")
            GetHLPath = INIHL.GetSection("Main").AddKey("Hyperlaunch_Path").GetValue()
            If GetHLPath.StartsWith("\") Then
                If gHSPath.EndsWith("\") Then
                    GetHLPath = gHSPath & "..\" & GetHLPath
                Else
                    GetHLPath = gHSPath & "\..\" & GetHLPath
                End If
            End If
            If GetHLPath.EndsWith("\") Then
                GetHLPath = Path.GetDirectoryName(GetHLPath)
            End If
            If GetHLPath.StartsWith("..") Then
                GetHLPath = GetHLPath.Replace("..", gHSPath)
            End If
            If Path.HasExtension(GetHLPath) = True Then
                GetHLPath = Path.GetDirectoryName(GetHLPath)
            End If
            LogEntry(LogType._Info, "{0}", "Finished loading configuration file for HyperLaunch.")
            Console.WriteLine("Load OK !")
        Catch ex1 As Exception
            LogEntry(LogType._Error, "{0}", "Cannot load configuration file for HyperLaunch !")
            Console.WriteLine("Load config file ERROR !")
        Finally
            INIHL = Nothing
        End Try
        Return GetHLPath
    End Function

    Public Sub SaveHyperT00lsParameters()
        If File.Exists(gCONFIGFILE) Then
            File.Delete(gCONFIGFILE)
        End If
        Application.DoEvents()

        'Desc
        gColDesc = Form1.TreeView2.Nodes("Node0").Checked

        'Enabled
        gColEnabled = Form1.TreeView2.Nodes("NodeEnabled").Checked

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
        gINI.GetSection("MAIN").AddKey("RocketLauncherPath").SetValue(gHLPath)
        gINI.GetSection("MAIN").AddKey("DisplayAllSystems").SetValue(gDisplayAllSystems)
        gINI.GetSection("MAIN").AddKey("Debug").SetValue(gDebugLog)
        gINI.GetSection("MAIN").AddKey("CheckRoms").SetValue(gCheckRoms)
        gINI.GetSection("MAIN").AddKey("gColDesc").SetValue(gColDesc)
        gINI.GetSection("MAIN").AddKey("gColEnabled").SetValue(gColEnabled)
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
        LogEntry(LogType._Info, "{0}", "Config saved ...")
    End Sub

    Public Function GetFileVersionInfo(ByVal filename As String) As Version
        Return Version.Parse(FileVersionInfo.GetVersionInfo(filename).FileVersion)
    End Function

    Public Sub LogEntry(ByVal Type As LogType, ByVal Func As String, ByVal Message As String)
        If gDebugLog = True Then
            Dim filePath As String = gDebugLogFile
            Using writer As New StreamWriter(filePath, True)
                Dim vDate As String
                Dim vHour As String
                Dim vType As String
                Dim vMessage As String = Message.Replace(";", " - ")
                Dim regex As New System.Text.RegularExpressions.Regex("[ \t\r\n]")
                vMessage = regex.Replace(vMessage, " ")
                Dim time As DateTime = DateTime.Now
                vDate = time.ToString("yyyy-MM-dd")
                vHour = time.ToString("HH:mm")

                Select Case Type
                    Case LogType._Debug
                        vType = "Debug"
                    Case (LogType._Error)
                        vType = "Error"
                    Case LogType._Info
                        vType = "Info"
                    Case LogType._Warning
                        vType = "Warning"
                    Case Else
                        vType = ""
                End Select

                writer.WriteLine(vDate & ";" & vHour & ";" & vType & ";" & String.Format(Func, New System.Diagnostics.StackFrame(1).GetMethod().Name) & ";" & vMessage)
            End Using
        End If
    End Sub

    Public Function GetDirnameOnly(ByVal vFullPath As String) As String
        Dim vDir As New System.IO.DirectoryInfo(vFullPath)
        Return vDir.Name
    End Function
#End Region
End Module
