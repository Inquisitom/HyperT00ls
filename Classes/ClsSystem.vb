Option Explicit On
Option Strict On

Imports System.IO
Imports System.Xml
Imports System.Collections
Imports System.Collections.ObjectModel

Public Class RomsCollection
    Inherits KeyedCollection(Of String, ClsRom)

    Protected Overrides Function GetKeyForItem(ByVal Rom As ClsRom) As String
        Return Rom.Name
    End Function
End Class

Public Class ClsSystem
    Inherits System.Collections.CollectionBase

#Region "Declarations"
    Public Roms As New RomsCollection
    Private mName As String

    Public IsEnabled As Boolean = False

    Public mSearchSubfolders As Boolean = False
    Public mRomExtensions As String
    Public mRomPath As String

    Private mXMLPath As String
    Private mINIPath As String

    Private mVideoPath As String
    Public AltVideoPath As String
    Public HasVideo As Boolean
    Public VideoExt As String

    Private mThemePath As String
    Public HasTheme As Boolean

    Private mWheelPath As String
    Public HasWheel As Boolean

    Private mBetaBriteClassicPath As String
    Public HasBetaBriteClassic As Boolean

    Private mBetaBritePrismPath As String
    Public HasBetaBritePrism As Boolean

    Private mEXEPath As String
    Private mIsEXE As Boolean

    Public HasHPBackground As Boolean = False
    Public HPBackgroundFolder As String

    Public HasHPFade As Boolean = False
    Public HPFadeFolder As String

    Public HasHPBezel As Boolean = False
    Public HPBezelFolder As String

#End Region

    Public Sub New(ByVal SystemName As String, ByVal HyperSpinRootPath As String, Optional ByVal LoadRoms As Boolean = False, Optional ByVal IsSystemEXE As Boolean = False, Optional ByVal Enabled As Boolean = True)
        Me.Name = SystemName
        mXMLPath = HyperSpinRootPath & "\Databases\" & SystemName & "\" & SystemName & ".xml"

        For Each tVExt As String In "png,flv,mp4".Split(CChar(","))
            Dim tVideoPath As String = HyperSpinRootPath & "\Media\Main Menu\Video\" & SystemName & "." & tVExt
            If File.Exists(tVideoPath) Then
                VideoPath = tVideoPath
            End If
        Next

        INIPath = HyperSpinRootPath & "\Settings\" & SystemName & ".ini"
        WheelPath = HyperSpinRootPath & "\Media\Main Menu\Images\Wheel\" & SystemName & ".png"
        ThemePath = HyperSpinRootPath & "\Media\Main Menu\Themes\" & SystemName & ".zip"
        BetaBriteClassicPath = HyperSpinRootPath & "\Media\Main Menu\Images\BetaBrite Classic\" & SystemName & ".png"
        BetaBritePrismPath = HyperSpinRootPath & "\Media\Main Menu\Images\BetaBrite Prism\" & SystemName & ".png"

        HPBackgroundFolder = gHLPath & "\Media\Backgrounds\" & SystemName & "\_Default\"
        If Directory.Exists(HPBackgroundFolder) Then
            If Directory.GetFiles(HPBackgroundFolder).Count > 0 Then
                HasHPBackground = True
            End If
        End If

        HPFadeFolder = gHLPath & "\Media\Fade\" & SystemName & "\_Default\"
        If Directory.Exists(HPFadeFolder) Then
            If Directory.GetFiles(HPFadeFolder).Count > 0 Then
                HasHPFade = True
            End If
        End If

        Dim numBezels As Integer = 0
        HPBezelFolder = gHLPath & "\Media\Bezels\" & SystemName & "\_Default\"
        If Directory.Exists(HPBezelFolder) Then
            For Each pDir In Directory.GetDirectories(HPBezelFolder)
                numBezels = Directory.GetFiles(pDir).Length
            Next
            numBezels = numBezels + Directory.GetFiles(HPBezelFolder).Length
            If numBezels > 0 Then
                HasHPBezel = True
            End If
        End If

        IsEXE = IsSystemEXE
        IsEnabled = Enabled

        Dim rINI As New IniFile
        If File.Exists(INIPath) Then
            rINI.Load(INIPath)

            Try
                AltVideoPath = Trim(rINI.GetSection("video defaults").GetKey("path").GetValue())
                If AltVideoPath.EndsWith("\") = False And AltVideoPath.Length > 1 Then
                    AltVideoPath = AltVideoPath & "\"
                End If
            Catch ex As Exception
                AltVideoPath = ""
            End Try

            Try
                mRomPath = rINI.GetSection("exe info").GetKey("rompath").GetValue()
            Catch ex As Exception
                If SystemName.ToLower <> "main menu" Then
                    LogEntry(LogType._Error, "{0}", "Cannot get default ROM path in " & INIPath & " : Error occured --> " & ex.Message.ToString)
                End If
            End Try

            Try
                mRomExtensions = rINI.GetSection("exe info").GetKey("romextension").GetValue()
            Catch ex As Exception
                If SystemName.ToLower <> "main menu" Then
                    LogEntry(LogType._Error, "{0}", "Cannot get default ROM Extension in " & INIPath & " : Error occured --> " & ex.Message.ToString)
                End If
                mRomExtensions = ""
                'LogEntry(LogType._Warning, "Setting default extension to ZIP")
            End Try

            Try
                mSearchSubfolders = CBool(rINI.GetSection("exe info").GetKey("searchsubfolders").GetValue())
            Catch ex As Exception
                mSearchSubfolders = False
            End Try
            rINI = Nothing
        End If

        If LoadRoms = True Then
            LoadSystemRoms(LoadRoms)
        End If
    End Sub

    Function ReturnRomNames() As String()
        Dim vXML As XmlDocument
        Dim vNODES As XmlNodeList
        vXML = New XmlDocument()
        If File.Exists(Me.XMLPath) = False Then
            LogEntry(LogType._Error, "{0}", "Ooops ! " & Me.XMLPath & " is not found, exiting.")
            Return Nothing
        End If

        vXML.Load(Me.XMLPath)
        vNODES = vXML.SelectNodes("menu/game")
        Dim i As Integer = 0
        Dim xRoms(vNODES.Count) As String
        Try
            For Each vNode As XmlNode In vNODES
                xRoms(i) = vNode.Attributes.GetNamedItem("name").InnerText
                i = i + 1
            Next
        Catch ex As Exception
            LogEntry(LogType._Error, "{0}", "Loading Systems : Error occured --> " & ex.Message.ToString)
        End Try

        Return xRoms
    End Function

    Sub RefreshSystem(ByVal ReloadRoms As Boolean)
        mXMLPath = gHSPath & "\Databases\" & mName & "\" & mName & ".xml"

        For Each tVExt As String In "png,flv,mp4".Split(CChar(","))
            Dim tVideoPath As String = gHSPath & "\Media\Main Menu\Video\" & mName & "." & tVExt
            If File.Exists(tVideoPath) Then
                VideoPath = tVideoPath
            End If
        Next

        INIPath = gHSPath & "\Settings\" & mName & ".ini"
        WheelPath = gHSPath & "\Media\Main Menu\Images\Wheel\" & mName & ".png"
        ThemePath = gHSPath & "\Media\Main Menu\Themes\" & mName & ".zip"
        BetaBriteClassicPath = gHSPath & "\Media\Main Menu\Images\BetaBrite Classic\" & mName & ".png"
        BetaBritePrismPath = gHSPath & "\Media\Main Menu\Images\BetaBrite Prism\" & mName & ".png"

        HPBackgroundFolder = gHLPath & "\Media\Backgrounds\" & mName & "\_Default\"

        If Directory.Exists(HPBackgroundFolder) Then
            If Directory.GetFiles(HPBackgroundFolder).Count > 0 Then
                HasHPBackground = True
            End If
        End If

        HPFadeFolder = gHLPath & "\Media\Fade\" & mName & "\_Default\"
        If Directory.Exists(HPFadeFolder) Then
            If Directory.GetFiles(HPFadeFolder).Count > 0 Then
                HasHPFade = True
            End If
        End If

        Dim rINI As New IniFile
        If File.Exists(INIPath) Then
            rINI.Load(INIPath)

            Try
                AltVideoPath = Trim(rINI.GetSection("video defaults").GetKey("path").GetValue())
                If AltVideoPath.EndsWith("\") = False And AltVideoPath.Length > 1 Then
                    AltVideoPath = AltVideoPath & "\"
                End If
            Catch ex As Exception
                AltVideoPath = ""
            End Try

            Try
                mRomPath = rINI.GetSection("exe info").GetKey("rompath").GetValue()
            Catch ex As Exception
                If mName.ToLower <> "main menu" Then
                    LogEntry(LogType._Error, "{0}", "Cannot get default ROM path in " & INIPath & " : Error occured --> " & ex.Message.ToString)
                End If
            End Try

            Try
                mRomExtensions = rINI.GetSection("exe info").GetKey("romextension").GetValue()
            Catch ex As Exception
                If mName.ToLower <> "main menu" Then
                    LogEntry(LogType._Error, "{0}", "Cannot get default ROM Extension in " & INIPath & " : Error occured --> " & ex.Message.ToString)
                End If
                mRomExtensions = ""
                'LogEntry(LogType._Warning, "Setting default extension to ZIP")
            End Try

            Try
                mSearchSubfolders = CBool(rINI.GetSection("exe info").GetKey("searchsubfolders").GetValue())
            Catch ex As Exception
                mSearchSubfolders = False
            End Try
            rINI = Nothing
        End If

        If ReloadRoms = True Then
            LoadSystemRoms(ReloadRoms)
        End If
    End Sub

#Region "Properties"

    Public Sub LoadSystemRoms(ByVal SearchSubfolders As Boolean)
        Me.Roms.Clear()
        Dim vXML As XmlDocument
        Dim vNODES As XmlNodeList

        vXML = New XmlDocument()
        If File.Exists(Me.XMLPath) = False Then
            MsgBox("Ooops ! " & Me.XMLPath & " is not found, exiting.", MsgBoxStyle.Critical, "Error")
            LogEntry(LogType._Error, "{0}", "Ooops ! " & Me.XMLPath & " is not found, exiting.")
            'Form1.Close()
            Exit Sub
        End If

        Try
            vXML.Load(Me.XMLPath)
        Catch ex As Exception
            LogEntry(LogType._Error, "{0}", "Cannot load : " & Me.XMLPath & " --> ABORDING")
            Exit Sub
        End Try

        vNODES = vXML.SelectNodes("menu/game")
        Dim i As Integer = 0
        Dim vSysName As String = mName
        Dim pName As String = ""
        Dim vError As Integer = 0

        Try
            If Form1.ProgressBar1.Visible = True And Form1.TabControl1.SelectedTab.Text.ToLower = "cleanup" Then
                Form1.ProgressBar1.Maximum = vNODES.Count + 1
                Form1.ProgressBar1.Value = 0
            End If
        Catch ex As Exception
            LogEntry(LogType._Warning, "{0}", ex.Message.ToString)
            Form1.ProgressBar1.Minimum = 1
            Form1.ProgressBar1.Maximum = 2
            Form1.ProgressBar1.Value = 1
        End Try

        For Each vNode As XmlNode In vNODES
            Try
                i = i + 1
                Try
                    If Form1.ProgressBar1.Visible = True And Form1.TabControl1.SelectedTab.Text.ToLower = "cleanup" Then
                        Form1.ProgressBar1.Value = Form1.ProgressBar1.Value + 1
                    End If
                Catch ex As Exception

                End Try

                Form1.ToolStripStatusLabel1.Text = "Loading rom " & i.ToString & "/" & vNODES.Count
                pName = ""
                Dim vSysEXE As Boolean = False
                Dim pIndex As Boolean = False
                Dim pImage As String = ""
                Dim pDescription As String = ""
                Dim pYear As String = ""
                Dim pManufacturer As String = ""
                Dim pCloneOf As String = ""
                Dim pCRC As String = ""
                Dim pRating As String = ""
                Dim pGenre As String = ""
                Dim pEnabled As Boolean = True

                Try
                    Dim xtmp As String = Trim(vNode.Attributes.GetNamedItem("index").InnerText.ToLower)
                    If xtmp = "" Or xtmp = "false" Then
                        pIndex = False
                    Else
                        pIndex = True
                    End If
                Catch ex As Exception
                    pIndex = False
                End Try

                Try
                    pImage = vNode.Attributes.GetNamedItem("image").InnerText
                Catch ex As Exception
                    pImage = ""
                End Try

                Try
                    If vNode.Attributes.GetNamedItem("enabled").InnerText = "0" Then
                        pEnabled = False
                    Else
                        pEnabled = True
                    End If
                Catch ex As Exception
                    pEnabled = True
                End Try

                pName = vNode.Attributes.GetNamedItem("name").InnerText

                For Each vAttribute As Xml.XmlNode In vNode.ChildNodes
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
                    Select Case vAttribute.Name
                        Case "description"
                            pDescription = vAttribute.InnerText
                        Case "cloneof"
                            pCloneOf = vAttribute.InnerText
                        Case "manufacturer"
                            pManufacturer = vAttribute.InnerText
                        Case "crc"
                            pCRC = vAttribute.InnerText
                        Case "year"
                            pYear = vAttribute.InnerText
                        Case "genre"
                            pGenre = vAttribute.InnerText
                        Case "rating"
                            pRating = vAttribute.InnerText
                            'Case "enabled"
                            '    If pEnabled = True Then
                            '        If Trim(vAttribute.InnerText).ToLower = "yes" Or Trim(vAttribute.InnerText).ToLower = "" Then
                            '            pEnabled = True
                            '        Else
                            '            pEnabled = False
                            '        End If
                            '    End If
                        Case Else
                            '
                    End Select
                Next

                'Here create a new rom
                If SearchSubfolders = False Then
                    Me.Roms.Add(New ClsRom(vSysName, pName, pDescription, mRomPath, mRomExtensions, True, False, pIndex, pImage, pYear, pManufacturer, pCloneOf, pCRC, pRating, pGenre, pEnabled))
                Else
                    Me.Roms.Add(New ClsRom(vSysName, pName, pDescription, mRomPath, mRomExtensions, True, mSearchSubfolders, pIndex, pImage, pYear, pManufacturer, pCloneOf, pCRC, pRating, pGenre, pEnabled))
                End If
                Application.DoEvents()
            Catch ex As Exception
                Form1.ToolStripStatusLabel1.Text = "Error occured : " & ex.Message.ToString & " --> Seems rom : " & pName & " is found twice or contains errors ...SKIPPING IT"
                LogEntry(LogType._Error, "{0}", "Loading Systems : Error occured --> " & ex.Message.ToString)
                LogEntry(LogType._Error, "{0}", "Seems rom : " & pName & " is found or contains errors ...SKIPPING IT")
                vError = vError + 1
            End Try
        Next
        Form1.ToolStripStatusLabel1.Text = "Loading rom " & i.ToString & "/" & vNODES.Count & " ---> DONE !"
        If vError > 0 Then
            MsgBox("Looks like your XML contains errors, " & vError.ToString & " entries have been skipped ... check the DebugLog", MsgBoxStyle.Critical, "Errors found")
        End If
    End Sub

    Public Property Name As String
        Get
            Return mName
        End Get
        Set(ByVal Value As String)
            mName = Value
        End Set
    End Property

    Public Property XMLPath As String
        Get
            Return mXMLPath
        End Get
        Set(ByVal Value As String)
            mXMLPath = Value
        End Set
    End Property

    Public Property IsEXE As Boolean
        Get
            Return mIsEXE
        End Get
        Set(ByVal Value As Boolean)
            mIsEXE = Value
            If Value = True Then
                Try
                    gINI.Load(Me.INIPath)
                    mEXEPath = gINI.GetSection("exe info").GetKey("path").GetValue & gINI.GetSection("exe info").GetKey("exe").GetValue
                Catch ex As Exception
                    LogEntry(LogType._Error, "{0}", "Problem while trying to get the EXE path for system " & Me.mName & " (ini=" & Me.INIPath & ") : " & ex.Message.ToString)
                    mEXEPath = ""
                End Try
            Else
                mEXEPath = ""
            End If
        End Set
    End Property

    Public Property INIPath As String
        Get
            Return mINIPath
        End Get
        Set(ByVal Value As String)
            mINIPath = Value
        End Set
    End Property

    Public Property VideoPath As String
        Get
            Return mVideoPath
        End Get
        Set(ByVal Value As String)
            mVideoPath = Value
            VideoExt = Right(Value, 3)
            If File.Exists(Value) Then
                HasVideo = True
            Else
                HasVideo = False
            End If
        End Set
    End Property

    Public Property ThemePath As String
        Get
            Return mThemePath
        End Get
        Set(ByVal Value As String)
            mThemePath = Value
            If File.Exists(Value) Then
                HasTheme = True
            Else
                HasTheme = False
            End If
        End Set
    End Property

    Public Property WheelPath As String
        Get
            Return mWheelPath
        End Get
        Set(ByVal Value As String)
            mWheelPath = Value
            If File.Exists(Value) Then
                HasWheel = True
            Else
                HasWheel = False
            End If
        End Set
    End Property

    Public Property BetaBriteClassicPath As String
        Get
            Return mBetaBriteClassicPath
        End Get
        Set(ByVal Value As String)
            mBetaBriteClassicPath = Value
            If File.Exists(Value) Then
                HasBetaBriteClassic = True
            Else
                HasBetaBriteClassic = False
            End If
        End Set
    End Property

    Public Property BetaBritePrismPath As String
        Get
            Return mBetaBritePrismPath
        End Get
        Set(ByVal Value As String)
            mBetaBritePrismPath = Value
            If File.Exists(Value) Then
                HasBetaBritePrism = True
            Else
                HasBetaBritePrism = False
            End If
        End Set
    End Property

#End Region
End Class
