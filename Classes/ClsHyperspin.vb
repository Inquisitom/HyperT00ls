Option Explicit On
Option Strict On

Imports System.IO
Imports System.Xml
Imports System.Collections.ObjectModel

Public Class SystemsCollection
    Inherits KeyedCollection(Of String, ClsSystem)

    Protected Overrides Function GetKeyForItem(ByVal System As ClsSystem) As String
        Return System.Name
    End Function
End Class

Public Class clsHyperSpin
    Public HyperSpinEXEversion As String
    Public HSRootPath As String
    Public HSDatabasesPath As String
    Public HSMediaPath As String
    Public HLRootPath As String
    Public MainMenuXML As String

    Public Systems As New SystemsCollection

    Public Sub New(ByVal HSPath As String)
        If HSPath.EndsWith("\") Then
            HSPath = HSPath.Substring(0, Len(HSPath) - 1)
        End If
        Me.HSRootPath = HSPath
        Me.HSMediaPath = HSPath & "\Media"
        Me.HSDatabasesPath = Me.HSRootPath & "\Databases"
        Me.MainMenuXML = Me.HSDatabasesPath & "\Main Menu\Main Menu.xml"
        LoadSystems()
    End Sub

    Private Sub LoadSystems()
        Dim vXML As XmlDocument
        Dim vNODES As XmlNodeList

        vXML = New XmlDocument()
        If File.Exists(Me.MainMenuXML) = False Then
            'MsgBox("Ooops ! " & Me.MainMenuXML & " is not found, exiting.", MsgBoxStyle.Critical, "Error")
            LogEntry(LogType._Error, Me.MainMenuXML & " is not found !")
            LogEntry(LogType._Error, "Choose a valid HyperSpin folder path in PARAMETERS !!")

            Form1.TabPage7.Select()
            Form1.TabControl1.SelectedIndex = 3
            'Form1.Close()
            Exit Sub
        End If

        'Systems in XML
        vXML.Load(Me.MainMenuXML)
        vNODES = vXML.SelectNodes("menu/game")
        For Each vNode As XmlNode In vNODES
            Dim vSysName As String = ""
            Dim vSysEXE As Boolean = False

            For Each vAttribute As XmlAttribute In vNode.Attributes
                'Here create a new system
                Select Case vAttribute.Name
                    Case "name"
                        vSysName = vAttribute.InnerText
                    Case "exe"
                        'IsEXE
                        vSysEXE = CBool(vAttribute.InnerText)
                End Select
            Next
            'At first, we do not load the roms
            Dim vSys As New ClsSystem(vSysName, HSRootPath, False, vSysEXE, True)
            Me.Systems.Add(vSys)
        Next

        'Systems not in XML, but present as folders in Databases
        For Each vFullDir In Directory.EnumerateDirectories(gDBPath)
            Dim vDir As String = Path.GetFileNameWithoutExtension(vFullDir)
            If Me.Systems.Contains(vDir) = False Then
                Dim vSys As New ClsSystem(vDir, HSRootPath, False, False, False)
                Me.Systems.Add(vSys)
            End If
        Next

    End Sub

    Public Sub ReloadSystem(ByVal SystemName As String)
        Dim vXML As XmlDocument
        Dim vNODES As XmlNodeList

        vXML = New XmlDocument()
        If File.Exists(Me.MainMenuXML) = False Then
            Form1.Button1.PerformClick()
            'MsgBox("Ooops ! " & Me.MainMenuXML & " is not found, exiting.", MsgBoxStyle.Critical, "Error")
            LogEntry(LogType._Error, Me.MainMenuXML & " is not found, exiting.")
            'Form1.Close()
            'Exit Sub
        End If
        vXML.Load(Me.MainMenuXML)
        vNODES = vXML.SelectNodes("menu/game")
        For Each vNode As XmlNode In vNODES
            Dim vSysName As String = ""
            Dim vSysEXE As Boolean = False
            For Each vAttribute As XmlAttribute In vNode.Attributes
                'Here create a new system
                Select Case vAttribute.Name
                    Case "name"
                        vSysName = vAttribute.InnerText
                    Case "exe"
                        'IsEXE
                        vSysEXE = CBool(vAttribute.InnerText)
                End Select
            Next

            'We found the system, so we reload it, and delete the old one

            If vSysName = SystemName Then
                Dim vSys As New ClsSystem(vSysName, HSRootPath, vSysEXE)
                Me.Systems.Add(vSys)
                Exit Sub
            End If
        Next
    End Sub

    Public Sub RemoveSystem(ByVal SystemName As String)
        'First, we clear Roms
        Me.Systems(SystemName).Roms.Clear()
        'Then, we delete the System
        Me.Systems.Remove(SystemName)
    End Sub
End Class
