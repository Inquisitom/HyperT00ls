Option Explicit On
Option Strict On

Imports System.IO
Imports System.Xml
Imports System.Collections
Imports System.Collections.ObjectModel

Public Class ClsRom
    Inherits System.Collections.CollectionBase

#Region "Declarations"
    '<game image="'" index="true" name="88games">
    '  <description>'88 Games</description>
    '  <cloneof />
    '  <crc />
    '  <manufacturer>Konami</manufacturer>
    '  <year>1988</year>
    '  <genre>Sports/Track &amp; Field</genre>
    '  <rating>AAMA - Green (Suitable For All Ages)</rating>
    '  <enabled>Yes</enabled>
    '</game>

    Public Name As String
    Public Description As String
    Public Index As Boolean
    Public Image As String
    Public Year As String
    Public Manufacturer As String
    Public CloneOf As String
    Public CRC As String
    Public Rating As String
    Public Genre As String
    Public Enabled As Boolean

    Public System As String
    Public RomPath As String
    Public RomFound As Boolean

    Public VideoPath As String
    Public HasVideo As Boolean
    Public VideoExt As String

    Public ThemePath As String
    Public HasTheme As Boolean

    Public WheelPath As String
    Public HasWheel As Boolean

    Public Art1Path As String
    Public HasArt1 As Boolean

    Public Art2Path As String
    Public HasArt2 As Boolean

    Public Art3Path As String
    Public HasArt3 As Boolean

    Public Art4Path As String
    Public HasArt4 As Boolean

    Public ArtworkPath As String
    Public HasArtwork As Boolean
    Public NumArtwork As Integer

    Public BackgroundsPath As String
    Public HasBackgrounds As Boolean
    Public NumBackgrounds As Integer

    Public BezelsPath As String
    Public HasBezels As Boolean
    Public NumBezels As Integer

    Public ControllerPath As String
    Public HasController As Boolean
    Public NumController As Integer

    Public FadePath As String
    Public HasFade As Boolean
    Public NumFade As Integer

    Public FontsPath As String
    Public HasFonts As Boolean
    Public NumFonts As Integer

    Public GuidesPath As String
    Public HasGuides As Boolean
    Public NumGuides As Integer

    Public ManualsPath As String
    Public HasManuals As Boolean
    Public NumManuals As Integer

    Public MenuImagesPath As String
    Public HasMenuImages As Boolean
    Public NumMenuImages As Integer

    Public MultiGamePath As String
    Public HasMultiGame As Boolean
    Public NumMultiGame As Integer

    Public MusicPath As String
    Public HasMusic As Boolean
    Public NumMusic As Integer

    Public SoundsPath As String
    Public HasSounds As Boolean
    Public NumSounds As Integer

    Public HLVideosPath As String
    Public HasHLVideos As Boolean
    Public NumHLVideos As Integer

    Public HLWheelsPath As String
    Public HasHLWheels As Boolean
    Public NumHLWheels As Integer

    Public IsFavorite As Boolean
#End Region

    Public Sub New(ByVal SystemName As String, _
                   ByVal RomName As String, _
                   ByVal Description As String, _
                   ByVal RomPaths As String, _
                   ByVal Extensions As String, _
                   Optional ByVal CheckMedia As Boolean = True, _
                   Optional ByVal SearchSubfolders As Boolean = False, _
                   Optional ByVal Index As Boolean = False, _
                   Optional ByVal Image As String = "", _
                   Optional ByVal Year As String = "", _
                   Optional ByVal Manufacturer As String = "", _
                   Optional ByVal CloneOf As String = "", _
                   Optional ByVal CRC As String = "", _
                   Optional ByVal Rating As String = "", _
                   Optional ByVal Genre As String = "", _
                   Optional ByVal Enabled As Boolean = True)

        With Me
            .System = SystemName
            .Name = RomName
            .Description = Description
            .Index = Index
            .Image = Image
            .Year = Year
            .Manufacturer = Manufacturer
            .CloneOf = CloneOf
            .CRC = CRC
            .Rating = Rating
            .Genre = Genre
            .Enabled = Enabled
            .RomPath = RomPaths
        End With

        If CheckMedia = True Then
            CheckRomMedia(SearchSubfolders, RomPaths, Extensions)
        End If
    End Sub

    Private Sub SearchRom(ByVal pRomName As String, ByVal pFolder As String, Optional ByVal pLookInSubfolders As Boolean = False)
        LogEntry(LogType._Debug, "{0}", "        Func SearchRom: pRomNameWithExt=" & pRomName & " pFolder=" & pFolder)
        If Right(Trim(pFolder), 1) <> "\" Then
            pFolder = Trim(pFolder) & "\"
        End If

        If File.Exists(pFolder & pRomName) Then
            Me.RomFound = True
            Me.RomPath = pFolder & pRomName
        End If

        If File.Exists(pFolder & Path.GetFileNameWithoutExtension(pRomName) & "\" & pRomName) Then
            Me.RomFound = True
            Me.RomPath = pFolder & Path.GetFileNameWithoutExtension(pRomName) & "\" & pRomName
        End If

        If Me.RomFound = False Then
            Try
                If pLookInSubfolders = True Then
                    For Each d As String In Directory.GetDirectories(pFolder)
                        SearchRom(pRomName, d, pLookInSubfolders)
                    Next
                End If
            Catch ex As System.Exception
                LogEntry(LogType._Warning, "{0}", "Searching for roms : " & ex.Message.ToString & "(" & pFolder & ")")
            End Try
        End If

        If Me.RomFound = True Then
            LogEntry(LogType._Info, "{0}", "        " & pRomName & " FOUND! in : " & Me.RomPath)
        Else
            LogEntry(LogType._Info, "{0}", "        " & pRomName & " NOT FOUND! in : " & pFolder)
        End If
    End Sub

    Public Sub CheckRomMedia(ByVal pSearchSub As Boolean, ByVal pRomFolder As String, ByVal Extensions As String)
        Dim vExtensions = Extensions
        'on reset tout à zéro
        Form1.DataGridView3.Enabled = False
        Form1.DataGridView3.Visible = False

        ResetValues()

        If gCheckRoms = True Then
            LogEntry(LogType._Debug, "{0}", Me.Name & " --> romFolder=" & pRomFolder & ", ext=" & Extensions & ")")
            Me.RomFound = False
            'First, we get ROMPATH + EXTENSIONS
            'The info is located within the INI file of the system under HyperSpin (and NOT HyperLaunch)  - but this might change!
            Dim HL3_INI As New IniFile
            Try
                'LogEntry(LogType._Info, "{0}", "ROM search : " & Me.Name)
                Dim EmulatorINI As String = gHLPath & "\Settings\" & Me.System & "\Emulators.ini"
                EmulatorINI = EmulatorINI.Replace("\\", "\")
                LogEntry(LogType._Debug, "{0}", "    EmulatorINI = " & EmulatorINI)
                If File.Exists(EmulatorINI) Then
                    LogEntry(LogType._Debug, "{0}", "    EmulatorINI = found")
                    HL3_INI.Load(EmulatorINI)
                    Dim DefaultEmul As String = HL3_INI.GetKeyValue("ROMS", "Default_Emulator")
                    LogEntry(LogType._Debug, "{0}", "    DefaultEmul = " & DefaultEmul)
                    Dim GlobEmulINI As New IniFile
                    Dim GlobEmulINIPath As String = Path.GetFullPath(gHLPath & "\Settings\Global Emulators.ini")
                    If File.Exists(GlobEmulINIPath) Then
                        Try
                            GlobEmulINI.Load(GlobEmulINIPath)
                            Extensions = GlobEmulINI.GetKeyValue(DefaultEmul, "Rom_Extension").Replace("|", ",")
                        Catch ex As Exception
                            LogEntry(LogType._Debug, "{0}", "        Error: Cannot load and get extensions of " & GlobEmulINIPath)
                        End Try
                    Else
                        LogEntry(LogType._Debug, "{0}", "        Warning: cannot find " & GlobEmulINIPath)
                    End If
                    If Extensions = "" Then
                        LogEntry(LogType._Debug, "{0}", "        Warning: no extension defined in Global Emulators.ini for " & DefaultEmul & ". Now using HyperHQ.")
                        Extensions = vExtensions
                        If Extensions = "" Then
                            LogEntry(LogType._Debug, "{0}", "        PLEASE CONFIGURE AT LEAST ROMS EXTENSIONS IN HyperHQ or HyperLaunch")
                        End If
                    End If
                    For Each HL3Folder In HL3_INI.GetKeyValue("ROMS", "Rom_Path").Split(CChar("|"))
                        If Me.RomFound = False Then
                            If HL3Folder.Contains("..") = True Then
                                HL3Folder = gHLPath & "\" & HL3Folder
                            End If
                            HL3Folder = Path.GetFullPath(HL3Folder)
                            For Each RomExt In Extensions.Split(CChar(","))
                                If Me.RomFound = False Then
                                    LogEntry(LogType._Debug, "{0}", "            Scanning in " & HL3Folder & " with extension : " & RomExt)
                                    SearchRom(Me.Name & "." & Trim(RomExt), HL3Folder, False) 'HL3 does NOT support subfolders
                                End If
                            Next
                        End If
                    Next
                Else
                    LogEntry(LogType._Info, "{0}", "EmulatorINI = not found - doing classic search (HS, not HL/RL)")
                    If pRomFolder = "" Or Extensions = "" Then
                        LogEntry(LogType._Warning, "{0}", Me.Name & ": No RomFolder or Extension set in HyperHQ and no HL/RL valid config found : aborted")
                    Else
                        For Each OneExt As String In Extensions.Split(CChar(","))
                            If Me.RomFound = False Then
                                LogEntry(LogType._Debug, "{0}", "Searching " & Me.Name & "." & Trim(OneExt) & " in : " & pRomFolder)
                                SearchRom(Me.Name & "." & Trim(OneExt), pRomFolder, pSearchSub)
                            End If
                        Next
                    End If
                End If
            Catch ex As Exception
                LogEntry(LogType._Error, "{0}", "Function : SearchRom -> " & ex.Message)
            Finally
                HL3_INI = Nothing
            End Try
        End If

        For Each tVExt As String In "png,flv,mp4".Split(CChar(","))
            Dim tVideoPath As String
            If gHSPath.EndsWith("\") Then
                tVideoPath = gHSPath & "Media\" & Me.System & "\Video\" & Me.Name & "." & tVExt
            Else
                tVideoPath = gHSPath & "\Media\" & Me.System & "\Video\" & Me.Name & "." & tVExt
            End If
            If File.Exists(tVideoPath) Then
                Me.VideoPath = tVideoPath
                If Me.VideoExt = "" Then
                    Me.VideoExt = tVExt
                Else
                    Me.VideoExt = Me.VideoExt & "/" & tVExt
                End If
                Me.HasVideo = True
            Else
                If gHS.Systems(Me.System).AltVideoPath <> "" Then
                    tVideoPath = gHS.Systems(Me.System).AltVideoPath & Me.Name & "." & tVExt
                    If File.Exists(tVideoPath) Then
                        Me.VideoPath = tVideoPath
                        If Me.VideoExt = "" Then
                            Me.VideoExt = tVExt
                        Else
                            Me.VideoExt = Me.VideoExt & "/" & tVExt
                        End If
                        Me.HasVideo = True
                    End If
                End If
            End If
        Next

        Me.WheelPath = gHSPath & "\Media\" & Me.System & "\Images\Wheel\" & Me.Name & ".png"
        If File.Exists(Me.WheelPath) Then
            Me.HasWheel = True
        End If

        Me.ThemePath = gHSPath & "\Media\" & Me.System & "\Themes\" & Me.Name & ".zip"
        If File.Exists(Me.ThemePath) Then
            Me.HasTheme = True
        End If

        Me.Art1Path = gHSPath & "\Media\" & Me.System & "\Images\Artwork1\" & Me.Name & ".png"
        If File.Exists(Me.Art1Path) Then
            Me.HasArt1 = True
        End If

        Me.Art2Path = gHSPath & "\Media\" & Me.System & "\Images\Artwork2\" & Me.Name & ".png"
        If File.Exists(Me.Art2Path) Then
            Me.HasArt2 = True
        End If

        Me.Art3Path = gHSPath & "\Media\" & Me.System & "\Images\Artwork3\" & Me.Name & ".png"
        If File.Exists(Me.Art3Path) Then
            Me.HasArt3 = True
        End If

        Me.Art4Path = gHSPath & "\Media\" & Me.System & "\Images\Artwork4\" & Me.Name & ".png"
        If File.Exists(Me.Art4Path) Then
            Me.HasArt4 = True
        End If

        Me.ArtworkPath = gHLPath & "\Media\Artwork\" & Me.System & "\" & Me.Name & "\"
        If Directory.Exists(Me.ArtworkPath) Then
            For Each pDir In Directory.GetDirectories(Me.ArtworkPath)
                Me.NumArtwork = Directory.GetFiles(pDir).Length
                If Me.NumArtwork > 0 Then
                    Me.HasArtwork = True
                End If
            Next
            Me.NumArtwork = Me.NumArtwork + Directory.GetFiles(Me.ArtworkPath).Length
            If Me.NumArtwork > 0 Then
                Me.HasArtwork = True
            End If
        End If

        Me.BackgroundsPath = gHLPath & "\Media\Backgrounds\" & Me.System & "\" & Me.Name & "\"
        If Directory.Exists(Me.BackgroundsPath) Then
            Me.NumBackgrounds = Directory.GetFiles(Me.BackgroundsPath).Length
            If Me.NumBackgrounds > 0 Then
                Me.HasBackgrounds = True
            End If
        End If

        Me.BezelsPath = gHLPath & "\Media\Bezels\" & Me.System & "\" & Me.Name & "\"
        If Directory.Exists(Me.BezelsPath) Then
            For Each pDir In Directory.GetDirectories(Me.BezelsPath)
                Me.NumBezels = Directory.GetFiles(pDir).Length
            Next
            Me.NumBezels = Me.NumBezels + Directory.GetFiles(Me.BezelsPath).Length
            If Me.NumBezels > 0 Then
                Me.HasBezels = True
            End If
        End If

        Me.ControllerPath = gHLPath & "\Media\Controller\" & Me.System & "\" & Me.Name & "\"
        If Directory.Exists(Me.ControllerPath) Then
            Me.NumController = Directory.GetFiles(Me.ControllerPath).Length
            If Me.NumController > 0 Then
                Me.HasController = True
            End If
        End If

        Me.FadePath = gHLPath & "\Media\Fade\" & Me.System & "\" & Me.Name & "\"
        If Directory.Exists(Me.FadePath) Then
            Me.NumFade = Directory.GetFiles(Me.FadePath).Length
            If Me.NumFade > 0 Then
                Me.HasFade = True
            End If
        End If

        Me.FontsPath = gHLPath & "\Media\Fonts\" & Me.System & "\" & Me.Name & "\"
        If Directory.Exists(Me.FontsPath) Then
            Me.NumFonts = Directory.GetFiles(Me.FontsPath).Length
            If Me.NumFonts > 0 Then
                Me.HasFonts = True
            End If
        End If

        Me.GuidesPath = gHLPath & "\Media\Guides\" & Me.System & "\" & Me.Name & "\"
        If Directory.Exists(Me.GuidesPath) Then
            For Each pDir In Directory.GetDirectories(Me.GuidesPath)
                Me.NumGuides = Directory.GetFiles(pDir).Length
                If Me.NumGuides > 0 Then
                    Me.HasGuides = True
                End If
            Next
            Me.NumGuides = Me.NumGuides + Directory.GetFiles(Me.GuidesPath).Length
            If Me.NumGuides > 0 Then
                Me.HasGuides = True
            End If
        End If

        Me.ManualsPath = gHLPath & "\Media\Manuals\" & Me.System & "\" & Me.Name & "\"
        If Directory.Exists(Me.ManualsPath) Then
            For Each pDir In Directory.GetDirectories(Me.ManualsPath)
                Me.NumManuals = Directory.GetFiles(pDir).Length
                If Me.NumManuals > 0 Then
                    Me.HasManuals = True
                End If
            Next
            Me.NumManuals = Me.NumManuals + Directory.GetFiles(Me.ManualsPath).Length
            If Me.NumManuals > 0 Then
                Me.HasManuals = True
            End If
        End If

        Me.MenuImagesPath = gHLPath & "\Media\Menu Images\" & Me.System & "\" & Me.Name & "\"
        If Directory.Exists(Me.MenuImagesPath) Then
            Me.NumMenuImages = Directory.GetFiles(Me.MenuImagesPath).Length
            If Me.NumMenuImages > 0 Then
                Me.HasMenuImages = True
            End If
        End If

        Me.MultiGamePath = gHLPath & "\Media\MultiGame\" & Me.System & "\" & Me.Name & "\"
        If Directory.Exists(Me.MultiGamePath) Then
            Me.NumMultiGame = Directory.GetFiles(Me.MultiGamePath).Length
            If Me.NumMultiGame > 0 Then
                Me.HasMultiGame = True
            End If
        End If

        Me.MusicPath = gHLPath & "\Media\Music\" & Me.System & "\" & Me.Name & "\"
        If Directory.Exists(Me.MusicPath) Then
            Me.NumMusic = Directory.GetFiles(Me.MusicPath).Length
            If Me.NumMusic > 0 Then
                Me.HasMusic = True
            End If
        End If

        Me.SoundsPath = gHLPath & "\Media\Sounds\" & Me.System & "\" & Me.Name & "\"
        If Directory.Exists(Me.SoundsPath) Then
            Me.NumSounds = Directory.GetFiles(Me.SoundsPath).Length
            If Me.NumSounds > 0 Then
                Me.HasSounds = True
            End If
        End If

        Me.HLVideosPath = gHLPath & "\Media\Videos\" & Me.System & "\" & Me.Name & "\"
        If Directory.Exists(Me.HLVideosPath) Then
            Me.NumHLVideos = Directory.GetFiles(Me.HLVideosPath).Length
            If Me.NumHLVideos > 0 Then
                Me.HasHLVideos = True
            End If
        End If

        Me.HLWheelsPath = gHLPath & "\Media\Wheels\" & Me.System & "\" & Me.Name & "\"
        If Directory.Exists(Me.HLWheelsPath) Then
            Me.NumHLWheels = Directory.GetFiles(Me.HLWheelsPath).Length
            If Me.NumHLWheels > 0 Then
                Me.HasHLWheels = True
            End If
        End If

        Form1.DataGridView3.Enabled = True
        Form1.DataGridView3.Visible = True
    End Sub

    Private Sub ResetValues()
        Me.HasArt1 = False
        Me.Art1Path = Nothing
        Me.HasArt2 = False
        Me.Art2Path = Nothing
        Me.HasArt3 = False
        Me.Art3Path = Nothing
        Me.HasArt4 = False
        Me.Art4Path = Nothing
        Me.HasArtwork = False
        Me.ArtworkPath = Nothing
        Me.HasBackgrounds = False
        Me.BackgroundsPath = Nothing
        Me.HasBezels = False
        Me.BezelsPath = Nothing
        Me.HasController = False
        Me.ControllerPath = Nothing
        Me.HasFade = False
        Me.FadePath = Nothing
        Me.HasFonts = False
        Me.FontsPath = Nothing
        Me.HasGuides = False
        Me.GuidesPath = Nothing
        Me.HasHLVideos = False
        Me.HLVideosPath = Nothing
        Me.HasHLWheels = False
        Me.HLWheelsPath = Nothing
        Me.HasManuals = False
        Me.ManualsPath = Nothing
        Me.HasMenuImages = False
        Me.MenuImagesPath = Nothing
        Me.HasMultiGame = False
        Me.MultiGamePath = Nothing
        Me.HasMusic = False
        Me.MusicPath = Nothing
        Me.HasSounds = False
        Me.SoundsPath = Nothing
        Me.HasTheme = False
        Me.ThemePath = Nothing
        Me.HasVideo = False
        Me.VideoPath = Nothing
        Me.HasWheel = False
        Me.WheelPath = Nothing
        Me.VideoExt = Nothing
    End Sub

#Region "Properties"


#End Region

End Class
