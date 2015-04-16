Option Explicit On
'Option Strict On

'TODO : facile - ajouter "cleanup" des videos dans le cas ou png/jpg/flv présent si MP4 est dispo
'TODO : possibilité de créer des "packages" pour Systèmes / Roms ... (facilité de partage)
'TODO : hard/prend du temps - Youtube video downloader
'TODO : integrate HyperLaunch\Data\Game Info\ for automatic rom XML adding ... (via form?)
'TODO : full fonctionnalité de HAVE/MISSING module
'TODO : un systeme drag and drop (plus de précisions)
'TODO : handle INI files within MediaManagement
'TODO : keep interface position(s) in config.ini - and load them at start

'VERSION CHANGES:
'- v1.3.2:
'   - Corrected some entries if relative paths are used
'- v1.3.1:
'   - Corrected some functions for rom search
'- v1.3.0.8:
'   - Corrected RomSearch function - HyperT00ls was not looking in %path%/romname/romname.romextension
'   - Loading XML or Check media for a System now does both at the same time - WARNING : do not forget to save your XML before Checking for Medias 
'   - Added XML Header + XML version for roms XML 
'   - Now double-clicking on ANY cell within RomsMediaCheck launches the Rom via HL3 - previously, this was working only if text on any cell was double-clicked
'   - added two new buttons in RomsMediaCheck & RomsXML to Launch associated rom entry with HL3
'   - Added check for Bezels in Systems ... and now handles Horizontal/Vertical bezels for both Roms & Systems
'   - Can now "fill missing entries" in a XML by looking in another XML you select
'- v1.3.0.6:
'   - Maintenant, le module de recherche de roms se base sur HL3 si trouvé, sinon, se base sur la config HyperSpin (donc oui, les deux sont gérés).
'   - ajout de "Crop XML" : permet de réduire un XML uniquement aux roms trouvées
'   - améliorations diverses de gestion des images
'   - HaveRoms de la fonction generate Have/Missing Report corrigé
'   - Double click sur une ligne de rom dans "Roms & Media Check" = lancement via Hyperlaunch !!
'   - diverses améliorations ...
'- v1.3.0.5:
'   - correction de nom pour la copie des pointers.png sur un système
'   - Possibilité de sélectionner les colonnes à afficher dans RomsMediaCheck (voir paramètres)
'- v1.3.0.4:
'   - charset UTF8 pour éviter les accents bizarres lors des copier/coller de guides en .txt 
'   - ajout du type de media que l'on peut copier/coller dans le menu contextuel pour chaque entrée (img et/ou txt)
'   - il est maintenant possible de redimensionner les colonnes dans RomsMediaCheck (en attendant l'option de les désactiver complètement)
'   - ajout de diverses améliorations esthétiques sur RomsMediaCheck
'- v1.3.0.3:
'   - plein de nouveaux remerciements dans le splashscreen !
'   - corrections de dernière minute (rien de grave)
'- v1.3.0.2:
'   - double-click sur item dans Systems/MediaMgmt lance l'app définie par défaut de cet item
'   - possibilité de renommer une rom dans le XML (et de copier/déplacer les medias associés)
'   - possibilité de copier/déplacer les medias associés d'une rom vers une autre (pas forcément du même système)
'- v1.3.0.1:
'   - correction mineure sur le SplashScreen
'   - Media management : affiche maintenant les répertoires dans "Guides" si trouvés
'- v1.3.0.0:
'   - possibilité de faire drag'n drop des entrées dans le XML des roms et des Systèmes 
'       !! attention, lorsqu'on sauvegarde le(s) XML, il garde l'ordre dans lequel sont affichées les entrées. Si vous faites un drag'n drop, ne triez pas par colonne après :)
'   - ajout d'un SplashScreen
'   - ajout du refresh des systèmes
'   - sécurité renforcée : plus possible de supprimer des folders depuis le right-click menu des systèmes et roms
'- v1.2.0.9:
'   - Correction de bugs divers
'   - Possibilité d'ajouter des guides en format texte depuis le presse-papiers si on a copié du texte en mémoire (hypert00ls détecte automatiquement le type : image ou texte)
'   - Possibilité de faire du MediaManagement au niveau systèmes (ajout/suppression/visualisation de medias etc ...)
'   - Ajout d'un module de génération de rapports : Have & Missing list (demande de Metheore)
'- v1.2.0.8:
'   - added check for alternate video path if set in systems config
'   - removed double-click feature in the MediaRoms view to add files. You must now use right-click in any entry in MediaManagement tab.
'   - Media Management:
'       - Added 3 ways to add media : from file, from clipboard (for images only!!), from URL
'       - all medias for selected rom are scanned and displayed in the treeview
'       - even the "Screenshots" folder in HyperPause/Artworks
'       - added way of opening the folder containing the media selected
'       - double-click on a file  in the treeview opens the file with its default associated program defined in windows
'       because the integrated web browser does not open FLV files, by double-clicking on video, it opens the FLV player
'

Imports System.IO
Imports System.Data
Imports System.Xml
Imports System.Collections
Imports System.Collections.ObjectModel
Imports System.Net
Imports System.Text.RegularExpressions

Public Class Form1
    Private fromIndex As Integer
    Private dragIndex As Integer
    Private dragRect As Rectangle

#Region "Loading"
    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        System.Windows.Forms.Application.EnableVisualStyles()
        Me.Visible = False
        'SplashScreen1.Show()
        Application.DoEvents()

        IsLoading = True
        LoadConfigFile()

        TextBox2.Text = gHSPath
        LoadSystems()

        RichTextBox1.Text = "Tips :" & vbCrLf & _
            "Goto MediaManagement to add/delete/open-folders medias related to selected rom" & vbCrLf & vbCrLf & _
            "On MediaManagement : Right click anywhere in the treeview to open the menu" & vbCrLf & vbCrLf & _
            "Double click on a Media under MediaManagement to open the file by the default associated program in windows."

        IsLoading = False
        TreeView2.ExpandAll()
        'SplashScreen1.Close()
        'Application.DoEvents()

        Me.Cursor = Cursors.Default
        Me.Visible = True
        LogEntry(LogType._Info, "All systems ready :)")
        Me.Focus()
    End Sub

    Private Sub LoadSystems()
        LogEntry(LogType._Info, "Func LoadSystems : reloading Systems in Treeview")
        gHS = Nothing
        gHS = New clsHyperSpin(gHSPath)
        Me.Text = "HyperT00ls v" & Application.ProductVersion

        'DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        TreeView1.Nodes(0).Nodes.Clear()

        'Creating node for each system
        For Each Node As ClsSystem In gHS.Systems
            If Node.Name.ToLower <> "main menu" Then
                If gDisplayAllSystems = True Then
                    Dim vTreeNode As New TreeNode
                    vTreeNode.Name = Node.Name
                    vTreeNode.Text = Node.Name
                    If Node.IsEnabled = False Then
                        vTreeNode.NodeFont = New Font(TreeView1.Font, FontStyle.Italic)
                    End If
                    TreeView1.Nodes(0).Nodes.Add(vTreeNode)
                Else
                    If Node.IsEnabled = True Then
                        Dim vTreeNode As New TreeNode
                        vTreeNode.Name = Node.Name
                        vTreeNode.Text = Node.Name
                        TreeView1.Nodes(0).Nodes.Add(vTreeNode)
                    End If
                End If
            End If
        Next

        'Expanding Systems
        TreeView1.Nodes(0).Expand()
        'Expanding Support
        'TreeView1.Nodes(1).Expand()

        FillSystemsDatagrid(DataGridView2, gDisplayAllSystems)
        DataGridView2.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)
    End Sub

#End Region

#Region "Top Menu"
    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

#End Region

#Region "Systems"
    Private Sub FillSystemsDatagrid(ByRef Grid As DataGridView, Optional ByVal LoadFolders As Boolean = False)
        LogEntry(LogType._Info, "Filling Systems Grid")
        'First, clear the grid
        Grid.Rows.Clear()

        'Then, we fill with Systems entries
        For Each vSys In gHS.Systems
            If vSys.Name.ToLower <> "main menu" Then
                Dim vRow As New DataGridViewRow
                vRow.CreateCells(Grid)

                vRow.Cells(0).Value = vSys.Name
                vRow.Cells(1).Value = vSys.IsEnabled

                If vSys.IsEnabled = True Then
                    vRow.Cells(0).Style.BackColor = Color.LightGreen
                    vRow.Cells(1).Style.BackColor = Color.LightGreen
                Else
                    vRow.Cells(0).Style.BackColor = Color.LightGray
                    vRow.Cells(1).Style.BackColor = Color.LightGray
                End If

                With vRow.Cells(2)
                    If vSys.HasWheel = True Then
                        .Value = "YES"
                        .Tag = vSys.WheelPath
                        .Style.BackColor = Color.LightGreen
                    Else
                        .Value = "NO"
                        .Tag = ""
                        .Style.BackColor = Color.Red
                    End If
                End With

                With vRow.Cells(3)
                    If vSys.HasVideo = True Then
                        .Value = vSys.VideoExt
                        .Tag = vSys.VideoPath
                        .Style.BackColor = Color.LightGreen
                    Else
                        .Value = "NO"
                        .Tag = ""
                        .Style.BackColor = Color.Red
                    End If
                End With

                With vRow.Cells(4)
                    If vSys.HasTheme = True Then
                        .Value = "YES"
                        .Tag = vSys.ThemePath
                        .Style.BackColor = Color.LightGreen
                    Else
                        .Value = "NO"
                        .Tag = ""
                        .Style.BackColor = Color.Red
                    End If
                End With

                With vRow.Cells(5)
                    If vSys.HasBetaBriteClassic = True Then
                        .Value = "YES"
                        .Tag = vSys.BetaBriteClassicPath
                        .Style.BackColor = Color.LightGreen
                    Else
                        .Value = "NO"
                        .Tag = ""
                        .Style.BackColor = Color.Red
                    End If
                End With

                With vRow.Cells(6)
                    If vSys.HasBetaBritePrism = True Then
                        .Value = "YES"
                        .Tag = vSys.BetaBritePrismPath
                        .Style.BackColor = Color.LightGreen
                    Else
                        .Value = "NO"
                        .Tag = ""
                        .Style.BackColor = Color.Red
                    End If
                End With

                With vRow.Cells(7)
                    If vSys.IsEXE = True Then
                        .Value = True
                    Else
                        .Value = False
                    End If
                    .Tag = ""
                    .Style.BackColor = Color.LightGreen
                End With

                Grid.Rows.Add(vRow)
            End If
        Next

        'And we fill with Folders not in MainMenu.xml if LoadFolders = true
        If LoadFolders = True Then
            '
        End If
    End Sub

    Private Sub FillRomsMediaCheckDatagrid(ByVal System As String, ByVal Grid As DataGridView)
        SplitContainer2.Enabled = False

        'First, clear the grid
        Grid.Enabled = False
        Grid.Rows.Clear()

        'Done before
        'gHS.Systems(System).Roms.Clear()
        'gHS.Systems(System).LoadSystemRoms(gHS.Systems(System).mSearchSubfolders)

        ToolStripProgressBar1.Maximum = gHS.Systems(System).Roms.Count
        ToolStripProgressBar1.Value = 0

        Grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Grid.AllowUserToResizeColumns = True

        'DISPLAYING COLUMNS OR NOT
        For Each col As DataGridViewColumn In Grid.Columns
            Select Case col.HeaderText
                Case "Rom Name"
                    col.Visible = True
                Case "Description"
                    col.Visible = gColDesc
                Case "Rom"
                    col.Visible = gCheckRoms
                Case "Wheel"
                    col.Visible = gColHSWheel
                Case "Video"
                    col.Visible = gColHSVideo
                Case "Art1"
                    col.Visible = gColHSArt1
                Case "Art2"
                    col.Visible = gColHSArt2
                Case "Art3"
                    col.Visible = gColHSArt3
                Case "Art4"
                    col.Visible = gColHSArt4
                Case "Theme"
                    col.Visible = gColHSTheme
                Case "Artwork"
                    col.Visible = gColHPArtwork
                Case "Backgrounds"
                    col.Visible = gColHPBackground
                Case "Bezels"
                    col.Visible = gColHPBezel
                Case "Controller"
                    col.Visible = gColHPController
                Case "Fade"
                    col.Visible = gColHPFade
                Case "Fonts"
                    col.Visible = gColHPFonts
                Case "Guides"
                    col.Visible = gColHPGuides
                Case "Manuals"
                    col.Visible = gColHPManuals
                Case "MenuImages"
                    col.Visible = gColHPMenuImages
                Case "Multigame"
                    col.Visible = gColHPMultigame
                Case "Music"
                    col.Visible = gColHPMusic
                Case "Sounds"
                    col.Visible = gColHPSounds
                Case "Videos"
                    col.Visible = gColHPVideos
                Case "HLWheels"
                    col.Visible = gColHPWheels
                Case Else
                    col.Visible = False
            End Select
            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        Next

        'Then, we fill with Systems entries
        For Each vRom In gHS.Systems(System).Roms
            'Dim vRom As ClsRom = xRom
            ToolStripProgressBar1.Value = ToolStripProgressBar1.Value + 1

            Dim vRow As New DataGridViewRow
            vRow.CreateCells(Grid)

            FillRomsMediaCheckROW(Grid, vRow, vRom, False)

            Grid.Rows.Add(vRow)
        Next

        ColorizeMediaRomsCells(Grid)
        Grid.Enabled = True

        SplitContainer2.Enabled = True
        'SplitContainer2.Refresh()
        Grid.ScrollBars = ScrollBars.None
        Grid.ScrollBars = ScrollBars.Both
        Grid.ClearSelection()
    End Sub

    Function getColumnIDbyHeader(ByVal columnHeader As String, ByVal Grid As DataGridView) As Integer
        Dim columnID As Integer = -1
        For Each column As DataGridViewColumn In Grid.Columns
            If column.HeaderText = columnHeader Then
                columnID = Grid.Columns.IndexOf(column)
                Exit For
            End If
        Next

        Return columnID
    End Function

    Function getColumnIDbyName(ByVal columnName As String, ByVal Grid As DataGridView) As Integer
        Dim columnID As Integer = -1
        For Each column As DataGridViewColumn In Grid.Columns
            If column.Name = columnName Then
                columnID = Grid.Columns.IndexOf(column)
                Exit For
            End If
        Next
        Return columnID
    End Function

    Private Sub FillRomsMediaCheckROW(ByRef Grid As DataGridView, ByRef vRow As DataGridViewRow, ByVal vRom As ClsRom, Optional ByVal Reload As Boolean = True)
        'vRow.CreateCells(Grid)

        'Name
        vRow.Cells(0).Value = vRom.Name
        'vRow.Cells(0).Style.BackColor = Color.LightGreen

        'Desc
        If gColDesc = True Then
            With vRow.Cells(getColumnIDbyName("ColDescription", Grid))
                If Len(vRom.Description) > 0 Then
                    '.Style.BackColor = Color.LightGreen
                    .Value = vRom.Description
                Else
                    .Value = ""
                    '.Style.BackColor = Color.Red
                End If
            End With
        End If

        'Enabled
        'With vRow.Cells(2)
        '    If vRom.Enabled = True Then
        '        .Value = True
        '        '.Style.BackColor = Color.LightGreen
        '    Else
        '        .Value = False
        '        '.Style.BackColor = Color.Red
        '    End If
        'End With

        'RomFound
        If gCheckRoms = True Then
            With vRow.Cells(getColumnIDbyName("Column10", Grid))
                If vRom.RomFound = True Then
                    .Value = "YES"
                    .Tag = vRom.RomPath
                    '.Style.BackColor = Color.LightGreen
                Else
                    .Value = "NO"
                    .Tag = ""
                    '.Style.BackColor = Color.Red
                End If
            End With
        End If

        'Wheel
        If gColHSWheel = True Then
            With vRow.Cells(getColumnIDbyName("Column4", Grid))
                If vRom.HasWheel = True Then
                    .Value = "YES"
                    .Tag = vRom.WheelPath
                    '.Style.BackColor = Color.LightGreen
                Else
                    .Value = "NO"
                    .Tag = ""
                    '.Style.BackColor = Color.Red
                End If
            End With
        End If

        'Video
        If gColHSVideo = True Then
            With vRow.Cells(getColumnIDbyName("Column5", Grid))
                If vRom.HasVideo = True Then
                    .Value = vRom.VideoExt
                    .Tag = vRom.VideoPath
                    '.Style.BackColor = Color.LightGreen
                Else
                    .Value = "NO"
                    .Tag = ""
                    '.Style.BackColor = Color.Red
                End If
            End With
        End If

        'Theme
        If gColHSTheme = True Then
            With vRow.Cells(getColumnIDbyName("Column2", Grid))
                If vRom.HasTheme = True Then
                    .Value = "YES"
                    .Tag = vRom.ThemePath
                    '.Style.BackColor = Color.LightGreen
                Else
                    .Value = "NO"
                    .Tag = ""
                    '.Style.BackColor = Color.Red
                End If
            End With
        End If

        'Art1
        If gColHSArt1 = True Then
            With vRow.Cells(getColumnIDbyName("Column6", Grid))
                If vRom.HasArt1 = True Then
                    .Value = "YES"
                    .Tag = vRom.Art1Path
                    '.Style.BackColor = Color.LightGreen
                Else
                    .Value = "NO"
                    '.Tag = ""
                    '.Style.BackColor = Color.Red
                End If
            End With
        End If

        'Art2
        If gColHSArt2 = True Then
            With vRow.Cells(getColumnIDbyName("Column7", Grid))
                If vRom.HasArt2 = True Then
                    .Value = "YES"
                    .Tag = vRom.Art2Path
                    '.Style.BackColor = Color.LightGreen
                Else
                    .Value = "NO"
                    '.Tag = ""
                    '.Style.BackColor = Color.Red
                End If
            End With
        End If

        'Art3
        If gColHSArt3 = True Then
            With vRow.Cells(getColumnIDbyName("Column8", Grid))
                If vRom.HasArt3 = True Then
                    .Value = "YES"
                    .Tag = vRom.Art3Path
                    '.Style.BackColor = Color.LightGreen
                Else
                    .Value = "NO"
                    '.Tag = ""
                    '.Style.BackColor = Color.Red
                End If
            End With
        End If

        'Art4
        If gColHSArt4 = True Then
            With vRow.Cells(getColumnIDbyName("Column1", Grid))
                If vRom.HasArt4 = True Then
                    .Value = "YES"
                    .Tag = vRom.Art4Path
                    '.Style.BackColor = Color.LightGreen
                Else
                    .Value = "NO"
                    '.Tag = ""
                    '.Style.BackColor = Color.Red
                End If
            End With
        End If

        'CRC (DataGridViewTextBoxColumn11)
        'vRow.Cells(11).Value = vRom.CRC.ToString

        'HL Artwork
        If gColHPArtwork = True Then
            vRow.Cells(getColumnIDbyName("ColArtwork", Grid)).Value = vRom.HasArtwork.ToString
        End If

        'Backgrounds
        If gColHPBackground = True Then
            vRow.Cells(getColumnIDbyName("ColBackgrounds", Grid)).Value = vRom.HasBackgrounds.ToString
        End If

        'Bezels
        If gColHPBezel = True Then
            vRow.Cells(getColumnIDbyName("ColBezels", Grid)).Value = vRom.HasBezels.ToString
        End If

        'Controller
        If gColHPController = True Then
            vRow.Cells(getColumnIDbyName("ColController", Grid)).Value = vRom.HasController.ToString
        End If

        'Fade
        If gColHPFade = True Then
            vRow.Cells(getColumnIDbyName("ColFade", Grid)).Value = vRom.HasFade.ToString
        End If

        'Fonts
        If gColHPFonts = True Then
            vRow.Cells(getColumnIDbyName("ColFonts", Grid)).Value = vRom.HasFonts.ToString
        End If

        'Guides
        If gColHPGuides = True Then
            vRow.Cells(getColumnIDbyName("ColGuides", Grid)).Value = vRom.HasGuides.ToString
        End If

        'Manuals
        If gColHPManuals = True Then
            vRow.Cells(getColumnIDbyName("ColManuals", Grid)).Value = vRom.HasManuals.ToString
        End If

        'Menu Images
        If gColHPMenuImages = True Then
            vRow.Cells(getColumnIDbyName("ColMenuImages", Grid)).Value = vRom.HasMenuImages.ToString
        End If

        'MultiGame
        If gColHPMultigame = True Then
            vRow.Cells(getColumnIDbyName("ColMultigame", Grid)).Value = vRom.HasMultiGame.ToString
        End If

        'Music
        If gColHPMusic = True Then
            vRow.Cells(getColumnIDbyName("ColMusic", Grid)).Value = vRom.HasMusic.ToString
        End If

        'Sounds
        If gColHPSounds = True Then
            vRow.Cells(getColumnIDbyName("ColSounds", Grid)).Value = vRom.HasSounds.ToString
        End If

        'Videos
        If gColHPVideos = True Then
            vRow.Cells(getColumnIDbyName("ColVideos", Grid)).Value = vRom.HasHLVideos.ToString
        End If

        'HLWheels
        If gColHPWheels = True Then
            vRow.Cells(getColumnIDbyName("ColHLWheels", Grid)).Value = vRom.HasHLWheels.ToString
        End If

        If Reload = True Then
            ColorizeMediaRomsCells(Grid)
            Grid.Enabled = True

            'SplitContainer2.Enabled = True
            'SplitContainer2.Refresh()
            Grid.ScrollBars = ScrollBars.None
            Grid.ScrollBars = ScrollBars.Both
            Grid.ClearSelection()
            vRow.Selected = True
            'Grid.SelectedRows(vRow.Index).Selected = True
            Grid.FirstDisplayedScrollingRowIndex = vRow.Index
        End If
    End Sub

    Private Sub FillSystemsMediaCheckROW(ByRef Grid As DataGridView, ByRef vRow As DataGridViewRow, ByVal vSys As ClsSystem)
        vRow.Cells(0).Value = vSys.Name
        vRow.Cells(1).Value = vSys.IsEnabled

        If vSys.IsEnabled = True Then
            vRow.Cells(0).Style.BackColor = Color.LightGreen
            vRow.Cells(1).Style.BackColor = Color.LightGreen
        Else
            vRow.Cells(0).Style.BackColor = Color.LightGray
            vRow.Cells(1).Style.BackColor = Color.LightGray
        End If

        With vRow.Cells(2)
            If vSys.HasWheel = True Then
                .Value = "YES"
                .Tag = vSys.WheelPath
                .Style.BackColor = Color.LightGreen
            Else
                .Value = "NO"
                .Tag = ""
                .Style.BackColor = Color.Red
            End If
        End With

        With vRow.Cells(3)
            If vSys.HasVideo = True Then
                .Value = vSys.VideoExt
                .Tag = vSys.VideoPath
                .Style.BackColor = Color.LightGreen
            Else
                .Value = "NO"
                .Tag = ""
                .Style.BackColor = Color.Red
            End If
        End With

        With vRow.Cells(4)
            If vSys.HasTheme = True Then
                .Value = "YES"
                .Tag = vSys.ThemePath
                .Style.BackColor = Color.LightGreen
            Else
                .Value = "NO"
                .Tag = ""
                .Style.BackColor = Color.Red
            End If
        End With

        With vRow.Cells(5)
            If vSys.HasBetaBriteClassic = True Then
                .Value = "YES"
                .Tag = vSys.BetaBriteClassicPath
                .Style.BackColor = Color.LightGreen
            Else
                .Value = "NO"
                .Tag = ""
                .Style.BackColor = Color.Red
            End If
        End With

        With vRow.Cells(6)
            If vSys.HasBetaBritePrism = True Then
                .Value = "YES"
                .Tag = vSys.BetaBritePrismPath
                .Style.BackColor = Color.LightGreen
            Else
                .Value = "NO"
                .Tag = ""
                .Style.BackColor = Color.Red
            End If
        End With

        With vRow.Cells(7)
            If vSys.IsEXE = True Then
                .Value = True
            Else
                .Value = False
            End If
            .Tag = ""
            .Style.BackColor = Color.LightGreen
        End With

    End Sub

    Private Sub FillRomsXMLDatagrid(ByVal System As String, ByRef Grid As DataGridView, ByVal CheckRoms As Boolean)
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

        SplitContainer4.Enabled = False

        'First, clear the grid
        Grid.Rows.Clear()

        'Done before the load
        'gHS.Systems(System).Roms.Clear()
        'gHS.Systems(System).LoadSystemRoms(CheckRoms)

        ToolStripProgressBar2.Maximum = gHS.Systems(System).Roms.Count
        ToolStripProgressBar2.Value = 0

        For Each col As DataGridViewColumn In Grid.Columns
            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        Next

        'Then, we fill with Systems entries
        For Each xRom In gHS.Systems(System).Roms
            Dim vRom As ClsRom = xRom
            ToolStripProgressBar2.Value = ToolStripProgressBar2.Value + 1

            Dim vRow As New DataGridViewRow
            vRow.CreateCells(Grid)

            'Name
            vRow.Cells(0).Value = vRom.Name.ToString
            If CheckRoms = True Then
                If vRom.RomFound = True Then
                    vRow.Cells(0).Style.BackColor = Color.LightGreen
                Else
                    vRow.Cells(0).Style.BackColor = Color.Red
                    vRow.Cells(0).Style.ForeColor = Color.White
                End If
            End If

            'Desc
            vRow.Cells(1).Value = vRom.Description.ToString
            'Manufacturer
            vRow.Cells(2).Value = vRom.Manufacturer.ToString
            vRow.Cells(3).Value = vRom.Year.ToString
            vRow.Cells(4).Value = vRom.Genre.ToString
            vRow.Cells(5).Value = vRom.Rating.ToString
            vRow.Cells(6).Value = vRom.CRC.ToString
            vRow.Cells(7).Value = vRom.CloneOf.ToString
            vRow.Cells(8).Value = vRom.Enabled
            vRow.Cells(9).Value = vRom.Image.ToString
            vRow.Cells(10).Value = vRom.Index

            Grid.Rows.Add(vRow)
        Next

        Grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        'ColorizeCells(Grid)

        SplitContainer4.Enabled = True
        'SplitContainer4.Refresh()
        'Grid.ScrollBars = ScrollBars.None
        Grid.ScrollBars = ScrollBars.Both
        Grid.ClearSelection()
    End Sub

    Private Sub ColorizeMediaRomsCells(ByVal Grid As DataGridView)
        If Grid.Name = "DataGridView1" Then
            'Colorize cells
            Dim x As Integer
            For x = 0 To Grid.RowCount - 1
                For Each vCell As DataGridViewCell In Grid.Rows(x).Cells
                    If vCell.Value Is Nothing Then
                        vCell.Style.BackColor = Color.DarkOrange
                    Else
                        Select Case vCell.Value.ToString.ToLower
                            Case "flv", "mp4", "png"
                                vCell.Style.BackColor = Color.LightGreen
                            Case "jpg/png/mp4", "png/mp4", "jpg/mp4", "png/flv", "jpg/png/flv", "jpg/flv", "jpg/png"
                                vCell.Style.BackColor = Color.LightGreen
                            Case "mp4/flv", "png/flv/mp4"
                                vCell.Style.BackColor = Color.Orange
                            Case "jpg"
                                vCell.Style.BackColor = Color.Red
                                vCell.Style.ForeColor = Color.White
                                vCell.Style.Font = New Font(TreeView1.Font, FontStyle.Bold)
                            Case "ok", "yes", "true"
                                vCell.Style.BackColor = Color.LightGreen
                            Case "no", "nok", "false"
                                vCell.Style.BackColor = Color.Red
                            Case Else
                                If vCell.ColumnIndex = 0 Then
                                    'Rom Name
                                    vCell.Style.BackColor = Color.LightGreen
                                ElseIf Grid.Columns(getColumnIDbyName("ColDescription", Grid)).Visible = True And vCell.ColumnIndex = 1 Then
                                    vCell.Style.BackColor = Color.LightGreen
                                Else
                                    vCell.Style.BackColor = Color.DarkOrange
                                End If
                        End Select
                    End If
                Next

                'If Grid.Rows(x).Cells(2).Value = True Then
                '    Grid.Rows(x).Cells(2).Style.BackColor = Color.LightGreen
                'Else
                '    Grid.Rows(x).Cells(2).Style.BackColor = Color.Gray
                'End If
                ''Rom
                'If Grid.Rows(x).Cells(3).Value = "YES" Then
                '    Grid.Rows(x).Cells(3).Style.BackColor = Color.LightGreen
                'Else
                '    Grid.Rows(x).Cells(3).Style.BackColor = Color.Red
                'End If
                ''Wheel
                'If Grid.Rows(x).Cells(4).Value = "YES" Then
                '    Grid.Rows(x).Cells(4).Style.BackColor = Color.LightGreen
                'Else
                '    Grid.Rows(x).Cells(4).Style.BackColor = Color.Red
                'End If
                ''Video
                'If Grid.Rows(x).Cells(5).Value = "NO" Then
                '    Grid.Rows(x).Cells(5).Style.BackColor = Color.Red
                'Else
                '    '"jpg,png,flv,mp4"
                '    Select Case Grid.Rows(x).Cells(5).Value
                '        Case "flv", "mp4", "png"
                '            Grid.Rows(x).Cells(5).Style.BackColor = Color.LightGreen
                '        Case "jpg/png/mp4", "png/mp4", "jpg/mp4", "png/flv", "jpg/png/flv", "jpg/flv", "jpg/png"
                '            Grid.Rows(x).Cells(5).Style.BackColor = Color.LightGreen
                '        Case "mp4/flv", "png/flv/mp4"
                '            Grid.Rows(x).Cells(5).Style.BackColor = Color.Orange
                '        Case "jpg"
                '            Grid.Rows(x).Cells(5).Style.BackColor = Color.Red
                '            Grid.Rows(x).Cells(5).Style.ForeColor = Color.White
                '            Grid.Rows(x).Cells(5).Style.Font = New Font(TreeView1.Font, FontStyle.Bold)
                '        Case Else
                '            Grid.Rows(x).Cells(5).Style.BackColor = Color.DarkOrange
                '    End Select
                'End If
                ''Theme
                'If Grid.Rows(x).Cells(6).Value = "YES" Then
                '    Grid.Rows(x).Cells(6).Style.BackColor = Color.LightGreen
                'Else
                '    Grid.Rows(x).Cells(6).Style.BackColor = Color.Red
                'End If
                ''Art1
                'If Grid.Rows(x).Cells(7).Value = "YES" Then
                '    Grid.Rows(x).Cells(7).Style.BackColor = Color.LightGreen
                'Else
                '    Grid.Rows(x).Cells(7).Style.BackColor = Color.Red
                'End If
                ''Art2
                'If Grid.Rows(x).Cells(8).Value = "YES" Then
                '    Grid.Rows(x).Cells(8).Style.BackColor = Color.LightGreen
                'Else
                '    Grid.Rows(x).Cells(8).Style.BackColor = Color.Red
                'End If
                ''Art3
                'If Grid.Rows(x).Cells(9).Value = "YES" Then
                '    Grid.Rows(x).Cells(9).Style.BackColor = Color.LightGreen
                'Else
                '    Grid.Rows(x).Cells(9).Style.BackColor = Color.Red
                'End If
                ''Art4
                'If Grid.Rows(x).Cells(10).Value = "YES" Then
                '    Grid.Rows(x).Cells(10).Style.BackColor = Color.LightGreen
                'Else
                '    Grid.Rows(x).Cells(10).Style.BackColor = Color.Red
                'End If
                ''CRC
                'Grid.Rows(x).Cells(11).Style.BackColor = Color.LightGray

                ''HL Artwork
                'If Grid.Rows(x).Cells(12).Value.ToString.ToLower = "false" Then
                '    Grid.Rows(x).Cells(12).Style.BackColor = Color.Red
                'Else
                '    Grid.Rows(x).Cells(12).Style.BackColor = Color.LightGreen
                'End If

                ''HL Backgrounds
                'If Grid.Rows(x).Cells(13).Value.ToString.ToLower = "false" Then
                '    Grid.Rows(x).Cells(13).Style.BackColor = Color.Red
                'Else
                '    Grid.Rows(x).Cells(13).Style.BackColor = Color.LightGreen
                'End If

                ''HL Bezels
                'If Grid.Rows(x).Cells(14).Value.ToString.ToLower = "false" Then
                '    Grid.Rows(x).Cells(14).Style.BackColor = Color.Red
                'Else
                '    Grid.Rows(x).Cells(14).Style.BackColor = Color.LightGreen
                'End If

                ''HL Controller
                'If Grid.Rows(x).Cells(15).Value.ToString.ToLower = "false" Then
                '    Grid.Rows(x).Cells(15).Style.BackColor = Color.Red
                'Else
                '    Grid.Rows(x).Cells(15).Style.BackColor = Color.LightGreen
                'End If

                ''HL Fade
                'If Grid.Rows(x).Cells(16).Value.ToString.ToLower = "false" Then
                '    Grid.Rows(x).Cells(16).Style.BackColor = Color.Red
                'Else
                '    Grid.Rows(x).Cells(16).Style.BackColor = Color.LightGreen
                'End If

                ''HL Fonts
                'If Grid.Rows(x).Cells(17).Value.ToString.ToLower = "false" Then
                '    Grid.Rows(x).Cells(17).Style.BackColor = Color.Red
                'Else
                '    Grid.Rows(x).Cells(17).Style.BackColor = Color.LightGreen
                'End If

                ''HL Guides
                'If Grid.Rows(x).Cells(18).Value.ToString.ToLower = "false" Then
                '    Grid.Rows(x).Cells(18).Style.BackColor = Color.Red
                'Else
                '    Grid.Rows(x).Cells(18).Style.BackColor = Color.LightGreen
                'End If

                ''HL Manuals
                'If Grid.Rows(x).Cells(19).Value.ToString.ToLower = "false" Then
                '    Grid.Rows(x).Cells(19).Style.BackColor = Color.Red
                'Else
                '    Grid.Rows(x).Cells(19).Style.BackColor = Color.LightGreen
                'End If

                ''HL Menu Images
                'If Grid.Rows(x).Cells(20).Value.ToString.ToLower = "false" Then
                '    Grid.Rows(x).Cells(20).Style.BackColor = Color.Red
                'Else
                '    Grid.Rows(x).Cells(20).Style.BackColor = Color.LightGreen
                'End If

                ''HL MultiGame
                'If Grid.Rows(x).Cells(21).Value.ToString.ToLower = "false" Then
                '    Grid.Rows(x).Cells(21).Style.BackColor = Color.Red
                'Else
                '    Grid.Rows(x).Cells(21).Style.BackColor = Color.LightGreen
                'End If

                ''HL Music
                'If Grid.Rows(x).Cells(22).Value.ToString.ToLower = "false" Then
                '    Grid.Rows(x).Cells(22).Style.BackColor = Color.Red
                'Else
                '    Grid.Rows(x).Cells(22).Style.BackColor = Color.LightGreen
                'End If

                ''HL Sounds
                'If Grid.Rows(x).Cells(23).Value.ToString.ToLower = "false" Then
                '    Grid.Rows(x).Cells(23).Style.BackColor = Color.Red
                'Else
                '    Grid.Rows(x).Cells(23).Style.BackColor = Color.LightGreen
                'End If

                ''HL Videos
                'If Grid.Rows(x).Cells(24).Value.ToString.ToLower = "false" Then
                '    Grid.Rows(x).Cells(24).Style.BackColor = Color.Red
                'Else
                '    Grid.Rows(x).Cells(24).Style.BackColor = Color.LightGreen
                'End If

                ''HL HLWheels
                'If Grid.Rows(x).Cells(25).Value.ToString.ToLower = "false" Then
                '    Grid.Rows(x).Cells(25).Style.BackColor = Color.Red
                'Else
                '    Grid.Rows(x).Cells(25).Style.BackColor = Color.LightGreen
                'End If
            Next
        End If
    End Sub

    Private Sub AddNewSystemMedia(ByVal vMediaType As SystemMediaType, ByVal vSystemName As String, ByVal vAddingType As AddingTypes)
        LogEntry(LogType._Info, "Add System Media : type->" & vMediaType.ToString & ", from->" & vAddingType.ToString & " in " & vSystemName)
        Dim openFileDialog1 As New OpenFileDialog()
        Dim DlgFilter As String = ""
        Dim destFolder As String = ""
        Dim FixedName As Boolean = True
        openFileDialog1.RestoreDirectory = True
        openFileDialog1.InitialDirectory = gHSPath
        Dim ForcePng As Boolean = False
        Dim vEndedOK As Boolean = True
        Dim vRow As DataGridViewRow = DataGridView2.SelectedRows(0)

        ToolStripStatusLabel1.Text = ""

        Select Case vMediaType
            Case SystemMediaType.HSWheel
                DlgFilter = "Wheels (*.png;*.bmp;*.gif;*.jpg)|*.png;*.bmp;*.gif;*.jpg"
                'destFolder = Path.GetDirectoryName(gHS.Systems(vSystemName).WheelPath)
                'If destFolder = Nothing Then
                destFolder = gHSPath & "\Media\Main Menu\Images\Wheel\"
                'End If
                ForcePng = True
            Case SystemMediaType.HSVideo
                DlgFilter = "Videos (*.png;*.bmp;*.gif;*.jpg;*.flv;*.mp4)|*.png;*.bmp;*.gif;*.jpg;*.flv;*.mp4"
                'destFolder = Path.GetDirectoryName(gHS.Systems(vSystemName).VideoPath)
                'If destFolder = Nothing Then
                destFolder = gHSPath & "\Media\Main Menu\Video\"
                'End If
            Case SystemMediaType.HSTheme
                DlgFilter = "Themes (*.zip)|*.zip"
                'destFolder = Path.GetDirectoryName(gHS.Systems(vSystemName).ThemePath)
                'If destFolder = Nothing Then
                destFolder = gHSPath & "\Media\Main Menu\Themes\"
                'End If
            Case SystemMediaType.HSPointer
                DlgFilter = "Pointers (*.png;*.bmp;*.gif;*.jpg)|*.png;*.bmp;*.gif;*.jpg"
                'destFolder = Path.GetDirectoryName(gHS.Systems(vSystemName).)
                'If destFolder = Nothing Then
                destFolder = gHSPath & "\Media\" & vSystemName & "\Images\Other\"
                'End If
                ForcePng = True
            Case SystemMediaType.HPBackground
                DlgFilter = "Backgrounds  (*.png;*.bmp;*.gif;*.jpg;*.tip)|*.png;*.bmp;*.gif;*.jpg;*.tip|All files (*.*)|*.*"
                destFolder = Path.GetDirectoryName(gHS.Systems(vSystemName).HPBackgroundFolder)
                FixedName = False
            Case SystemMediaType.HPFade
                DlgFilter = "Fades (*.png;*.bmp;*.gif;*.jpg)|*.png;*.bmp;*.gif;*.jpg"
                destFolder = Path.GetDirectoryName(gHS.Systems(vSystemName).HPFadeFolder)
                FixedName = False
                ForcePng = True
            Case SystemMediaType.HPBezel
                DlgFilter = "Bezels (*.png;*.bmp;*.gif;*.jpg)|*.png;*.bmp;*.gif;*.jpg"
                destFolder = Path.GetDirectoryName(gHS.Systems(vSystemName).HPBezelFolder)
                FixedName = False
                ForcePng = True
        End Select

        'just in case of
        If destFolder.EndsWith("\") = False Then
            destFolder = destFolder & "\"
        End If
        If destFolder.EndsWith("\\") Then
            destFolder = destFolder.Replace("\\", "\")
        End If

        LogEntry(LogType._Info, "         Filter applied = " & DlgFilter)
        LogEntry(LogType._Info, "         Destination folder = " & destFolder)
        LogEntry(LogType._Info, "         Fixed FileName = " & FixedName.ToString)
        LogEntry(LogType._Info, "         Force png format = " & ForcePng.ToString)
        Try
            If Directory.Exists(destFolder) = False Then
                Directory.CreateDirectory(destFolder)
                Application.DoEvents()
            End If
        Catch ex As Exception
            LogEntry(LogType._Error, "Cannot create destination folder (" & destFolder & "): " & ex.Message)
            LogEntry(LogType._Error, "Abording copy of file ...")
            vEndedOK = False
            Exit Sub
        End Try

        Select Case vAddingType
            Case AddingTypes.FromFile
                openFileDialog1.Filter = DlgFilter
                openFileDialog1.FilterIndex = 1
                If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                    Dim vExt As String = Path.GetExtension(openFileDialog1.FileName).ToLower
                    Dim DestFilename As String = ""

                    If FixedName = True Then
                        'If this is an Hyperspin file
                        If vExt = ".bmp" Or vExt = ".gif" Or vExt = ".jpg" Or vExt = ".png" Then
                            If ForcePng = True Then
                                Dim xx As Image = Image.FromFile(openFileDialog1.FileName)
                                If vMediaType = SystemMediaType.HSPointer Then
                                    DestFilename = destFolder & "Pointer.png"
                                Else
                                    DestFilename = destFolder & vSystemName & ".png"
                                End If

                                LogEntry(LogType._Info, "          Complete File destination = " & DestFilename)
                                Try
                                    Dim todel As Boolean = False
                                    UnloadSystemMediaFromPictureBoxes(DestFilename)
                                    Application.DoEvents()
                                    LogEntry(LogType._Info, "          Unloading file from Pictures Boxes")
                                    If File.Exists(DestFilename) = True Then
                                        todel = True
                                    End If
                                    Application.DoEvents()
                                    If todel = True Then
                                        Try
                                            LogEntry(LogType._Info, "          Deleting " & DestFilename)
                                            File.Delete(DestFilename)
                                            Application.DoEvents()
                                        Catch ex As Exception
                                            LogEntry(LogType._Error, "Cannot delete original file. Error: " & ex.Message)
                                            ToolStripStatusLabel1.Text = "Cannot delete original file. Error: " & ex.Message
                                        End Try
                                    End If
                                    xx.Save(DestFilename, System.Drawing.Imaging.ImageFormat.Png)
                                    LogEntry(LogType._Info, "File copied successfully")
                                Catch ex As Exception
                                    vEndedOK = False
                                    LogEntry(LogType._Error, "Cannot copy file. Original error: " & ex.Message)
                                    ToolStripStatusLabel1.Text = "Cannot copy file. Original error: " & ex.Message
                                    Exit Sub
                                End Try
                            Else
                                If vMediaType = SystemMediaType.HSPointer Then
                                    DestFilename = destFolder & "Pointer.png"
                                Else
                                    DestFilename = destFolder & vSystemName & vExt
                                End If
                                LogEntry(LogType._Info, "Complete File destination = " & DestFilename)
                                Try
                                    UnloadSystemMediaFromPictureBoxes(DestFilename)
                                    If File.Exists(DestFilename) = True Then
                                        File.Delete(DestFilename)
                                        Application.DoEvents()
                                    End If
                                    File.Copy(openFileDialog1.FileName, DestFilename, True)
                                    LogEntry(LogType._Info, "File copied successfully")
                                Catch ex As Exception
                                    vEndedOK = False
                                    LogEntry(LogType._Error, "Cannot copy file. Original error: " & ex.Message)
                                    ToolStripStatusLabel1.Text = "Cannot copy file. Original error: " & ex.Message
                                    Exit Sub
                                End Try
                            End If
                        Else
                            If vMediaType = SystemMediaType.HSPointer Then
                                DestFilename = destFolder & "Pointer.png"
                            Else
                                DestFilename = destFolder & vSystemName & ".png"
                            End If
                            LogEntry(LogType._Info, "Complete File destination = " & DestFilename)
                            Try
                                UnloadSystemMediaFromPictureBoxes(DestFilename)
                                If File.Exists(DestFilename) = True Then
                                    File.Delete(DestFilename)
                                    Application.DoEvents()
                                End If
                                File.Copy(openFileDialog1.FileName, DestFilename, True)
                                LogEntry(LogType._Info, "File copied successfully")
                            Catch ex As Exception
                                vEndedOK = False
                                LogEntry(LogType._Error, "Cannot copy file. Original error: " & ex.Message)
                                ToolStripStatusLabel1.Text = "Cannot copy file. Original error: " & ex.Message
                                Exit Sub
                            End Try
                        End If
                    Else
                        'If this is an HyperPause file
                        If ForcePng = True Then
                            Dim xx As Image = Image.FromFile(openFileDialog1.FileName)
                            Dim tempFile As String = Path.GetFileNameWithoutExtension(openFileDialog1.FileName)
                            DestFilename = destFolder & tempFile & ".png"
                            LogEntry(LogType._Info, "Complete File destination = " & DestFilename)
                            Try
                                UnloadSystemMediaFromPictureBoxes(DestFilename)
                                If File.Exists(DestFilename) = True Then
                                    File.Delete(DestFilename)
                                    Application.DoEvents()
                                End If
                                xx.Save(DestFilename, System.Drawing.Imaging.ImageFormat.Png)
                                LogEntry(LogType._Info, "File copied successfully")
                            Catch ex As Exception
                                vEndedOK = False
                                LogEntry(LogType._Error, "Cannot copy file. Original error: " & ex.Message)
                                ToolStripStatusLabel1.Text = "Cannot copy file. Original error: " & ex.Message
                                Exit Sub
                            End Try
                        Else
                            DestFilename = destFolder & Path.GetFileName(openFileDialog1.FileName).ToLower
                            LogEntry(LogType._Info, "Complete File destination = " & DestFilename)
                            Try
                                UnloadSystemMediaFromPictureBoxes(DestFilename)
                                If File.Exists(DestFilename) = True Then
                                    File.Delete(DestFilename)
                                    Application.DoEvents()
                                End If
                                File.Copy(openFileDialog1.FileName, DestFilename, True)
                                LogEntry(LogType._Info, "File copied successfully")
                            Catch ex As Exception
                                vEndedOK = False
                                LogEntry(LogType._Error, "Cannot copy file. Original error: " & ex.Message)
                                ToolStripStatusLabel1.Text = "Cannot copy file. Original error: " & ex.Message
                                Exit Sub
                            End Try
                        End If
                    End If
                End If
            Case AddingTypes.FromClipboard
                If Clipboard.ContainsImage = True Then
                    Dim DestFilename As String
                    If FixedName = True Then
                        DestFilename = destFolder & vSystemName & ".png"
                    Else
                        Dim vTmpName As String = InputBox("Enter new name for image (without extension) : ", "Filename", "")
                        If vTmpName = "" Then
                            MsgBox("You must specify a valid name ... aborting", MsgBoxStyle.Critical, "Error")
                            Exit Sub
                        Else
                            DestFilename = destFolder & vTmpName & ".png"
                        End If
                    End If
                    Try
                        LogEntry(LogType._Info, "Complete File destination = " & DestFilename)
                        UnloadSystemMediaFromPictureBoxes(DestFilename)
                        If File.Exists(DestFilename) = True Then
                            File.Delete(DestFilename)
                            Application.DoEvents()
                        End If
                        Clipboard.GetImage.Save(DestFilename, System.Drawing.Imaging.ImageFormat.Png)
                        LogEntry(LogType._Info, "File copied successfully")
                    Catch ex As Exception
                        vEndedOK = False
                        LogEntry(LogType._Error, "Cannot copy file. Original error: " & ex.Message)
                        ToolStripStatusLabel1.Text = "Cannot copy file. Original error: " & ex.Message
                    End Try
                Else
                    vEndedOK = False
                    LogEntry(LogType._Error, "Looks like clipboard does NOT contain an image, please retry.")
                    ToolStripStatusLabel1.Text = "Looks like clipboard does NOT contain an image, please retry."
                End If
            Case AddingTypes.FromURL
                Dim vURL = InputBox("Enter URL:", "Source file", "")
                If vURL = "" Then
                    MsgBox("You must specify a valid URL ... aborting", MsgBoxStyle.Critical, "Error")
                    Exit Sub
                End If
                Me.Cursor = Cursors.WaitCursor
                Application.DoEvents()
                Dim DestFilename As String
                Dim vFileName = vURL.Substring(vURL.LastIndexOf("/") + 1)
                Try
                    ' Make a WebClient.
                    Dim web_client As WebClient = New WebClient
                    ' Download the file.
                    If FixedName = True Then
                        If ForcePng = True Then
                            Dim vTmp As String = Path.GetTempPath & vFileName
                            DestFilename = destFolder & vSystemName & ".png"
                            LogEntry(LogType._Info, "Complete File destination = " & DestFilename)
                            Try
                                If Path.GetExtension(vFileName).ToLower = ".png" Then
                                    UnloadSystemMediaFromPictureBoxes(DestFilename)
                                    If File.Exists(DestFilename) = True Then
                                        File.Delete(DestFilename)
                                        Application.DoEvents()
                                    End If
                                    web_client.DownloadFile(vURL, DestFilename)
                                    LogEntry(LogType._Info, "File copied successfully")
                                Else
                                    UnloadSystemMediaFromPictureBoxes(DestFilename)
                                    web_client.DownloadFile(vURL, vTmp)
                                    Dim vImg As Image = Image.FromFile(vTmp)
                                    If File.Exists(DestFilename) = True Then
                                        File.Delete(DestFilename)
                                        Application.DoEvents()
                                    End If
                                    vImg.Save(DestFilename, System.Drawing.Imaging.ImageFormat.Png)
                                    LogEntry(LogType._Info, "File copied successfully")
                                End If
                            Catch ex As Exception
                                web_client.CancelAsync()
                                web_client = Nothing
                                File.Delete(vTmp)
                                vEndedOK = False
                                LogEntry(LogType._Error, "Cannot download/copy file. Original error: " & ex.Message)
                                ToolStripStatusLabel1.Text = "Cannot download/copy file. Original error: " & ex.Message
                                Me.Cursor = Cursors.Default
                                Exit Sub
                            End Try
                        Else
                            Try
                                DestFilename = destFolder & vSystemName & Path.GetExtension(vFileName)
                                LogEntry(LogType._Info, "Complete File destination = " & DestFilename)
                                UnloadSystemMediaFromPictureBoxes(DestFilename)
                                If File.Exists(DestFilename) = True Then
                                    File.Delete(DestFilename)
                                    Application.DoEvents()
                                End If
                                web_client.DownloadFile(vURL, DestFilename)
                                LogEntry(LogType._Info, "File copied successfully")
                            Catch ex As Exception
                                web_client.CancelAsync()
                                web_client = Nothing
                                vEndedOK = False
                                LogEntry(LogType._Error, "Cannot download/copy file. Original error: " & ex.Message)
                                ToolStripStatusLabel1.Text = "Cannot download/copy file. Original error: " & ex.Message
                                Me.Cursor = Cursors.Default
                                Exit Sub
                            End Try
                        End If
                    Else
                        DestFilename = destFolder & vFileName
                        LogEntry(LogType._Info, "Complete File destination = " & DestFilename)
                        UnloadSystemMediaFromPictureBoxes(DestFilename)
                        Try
                            If File.Exists(DestFilename) = True Then
                                File.Delete(DestFilename)
                                Application.DoEvents()
                            End If
                            web_client.DownloadFile(vURL, DestFilename)
                            LogEntry(LogType._Info, "File copied successfully")
                        Catch ex As Exception
                            vEndedOK = False
                            LogEntry(LogType._Error, "Cannot download/copy file. Original error: " & ex.Message)
                            ToolStripStatusLabel1.Text = "Cannot download/copy file. Original error: " & ex.Message
                            Me.Cursor = Cursors.Default
                            Exit Sub
                        End Try
                    End If
                Catch ex As Exception
                    vEndedOK = False
                    LogEntry(LogType._Error, "Cannot download/copy file. Original error: " & ex.Message)
                    ToolStripStatusLabel1.Text = "Cannot download/copy file. Original error: " & ex.Message
                    Me.Cursor = Cursors.Default
                End Try
                Me.Cursor = Cursors.Default
            Case AddingTypes.FromYouTube
                Dim vURL As String = InputBox("Enter full youtube URL here:", "URL", "https://www.youtube.com/watch?v=JfAc6Y9RyUA")
                If vURL = "" Then
                    vEndedOK = False
                    Exit Select
                End If
                vEndedOK = True
                Me.Visible = True
                Me.Focus()

        End Select

        'If vEndedOK = True Then
        DataGridView2.ClearSelection()
        gHS.Systems(vSystemName).RefreshSystem(False)
        FillSystemsMediaCheckROW(DataGridView2, vRow, gHS.Systems(vSystemName))
        vRow.Selected = True
        'End If
    End Sub
#End Region

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        LoadFULLSystem(TreeView1.SelectedNode.Text)
    End Sub

    'Hyperlaunch rom
    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        Dim vRow As Integer = e.RowIndex
        Dim vRom As String = DataGridView1.Rows(vRow).Cells(0).Value.ToString
        Dim vSystem As String = ToolStripLabel1.Text

        If vSystem = "" Or vSystem = "System" Or vRom = "" Then
            Exit Sub
        End If

        Try
            Dim psi As New ProcessStartInfo
            Dim p As New Process
            psi.FileName = gHLPath & "\HyperLaunch.exe"
            psi.Arguments = """" & vSystem & """ """ & vRom & """"
            LogEntry(LogType._Info, "Launching : " & psi.FileName & " with arguments: " & psi.Arguments)
            p.StartInfo = psi
            p.Start()
        Catch ex As Exception
            LogEntry(LogType._Info, "Error: cannot launch " & vRom & " in System:" & vSystem & " ! -> " & ex.Message)
        End Try
    End Sub

    Private Sub DataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged
        TreeViewMedia.Nodes.Clear()
        PictureBox1.Image = Nothing
        PictureBox2.Image = Nothing
        PictureBox3.Image = Nothing
        PictureBox4.Image = Nothing
        PictureBox5.Image = Nothing
        PictureBox11.Image = Nothing
        WBMedia.Navigate("")
        Application.DoEvents()

        If DataGridView1.Enabled = True Then
            Dim CurrentSystem As String = ToolStripLabel1.Text

            Try
                If DataGridView1.SelectedRows.Count > 0 Then
                    ToolStripButton10.Enabled = True
                    If DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString <> "" Then
                        Dim pRom As String = DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString
                        Dim pSystem As String = ToolStripLabel1.Text
                        Dim pRow As DataGridViewRow = DataGridView1.SelectedRows.Item(0)
                        Dim rootNode As New TreeNode(gHS.Systems(pSystem).Roms(pRom).Description)

                        rootNode.Nodes.Add("HyperSpin", "HyperSpin")

                        If gHS.Systems(pSystem).Roms(pRom).HasTheme Then
                            rootNode.Nodes("HyperSpin").Nodes.Add("Theme", "Theme")
                            rootNode.Nodes("HyperSpin").Nodes("Theme").Tag = Path.GetDirectoryName(gHS.Systems(pSystem).Roms(pRom).ThemePath)
                        End If

                        If gHS.Systems(pSystem).Roms(pRom).HasArt1 Or gHS.Systems(pSystem).Roms(pRom).HasArt2 Or gHS.Systems(pSystem).Roms(pRom).HasArt3 Or gHS.Systems(pSystem).Roms(pRom).HasArt4 Then
                            rootNode.Nodes("HyperSpin").Nodes.Add("Art", "Art")
                            If gHS.Systems(pSystem).Roms(pRom).HasArt1 Then
                                Dim fs As System.IO.FileStream
                                fs = New System.IO.FileStream(gHS.Systems(pSystem).Roms(pRom).Art1Path, IO.FileMode.Open, IO.FileAccess.Read)
                                PictureBox2.Image = System.Drawing.Image.FromStream(fs)
                                fs.Close()
                                'PictureBox2.ImageLocation = gHS.Systems(pSystem).Roms(pRom).Art1Path
                                rootNode.Nodes("HyperSpin").Nodes("Art").Nodes.Add("Art1", "Art1")
                                rootNode.Nodes("HyperSpin").Nodes("Art").Nodes("Art1").Tag = gHS.Systems(pSystem).Roms(pRom).Art1Path
                            End If
                            If gHS.Systems(pSystem).Roms(pRom).HasArt2 Then
                                Dim fs As System.IO.FileStream
                                fs = New System.IO.FileStream(gHS.Systems(pSystem).Roms(pRom).Art2Path, IO.FileMode.Open, IO.FileAccess.Read)
                                PictureBox3.Image = System.Drawing.Image.FromStream(fs)
                                fs.Close()
                                'PictureBox3.ImageLocation = gHS.Systems(pSystem).Roms(pRom).Art2Path
                                rootNode.Nodes("HyperSpin").Nodes("Art").Nodes.Add("Art2", "Art2")
                                rootNode.Nodes("HyperSpin").Nodes("Art").Nodes("Art2").Tag = gHS.Systems(pSystem).Roms(pRom).Art2Path
                            End If
                            If gHS.Systems(pSystem).Roms(pRom).HasArt3 Then
                                Dim fs As System.IO.FileStream
                                fs = New System.IO.FileStream(gHS.Systems(pSystem).Roms(pRom).Art3Path, IO.FileMode.Open, IO.FileAccess.Read)
                                PictureBox4.Image = System.Drawing.Image.FromStream(fs)
                                fs.Close()
                                'PictureBox4.ImageLocation = gHS.Systems(pSystem).Roms(pRom).Art3Path
                                rootNode.Nodes("HyperSpin").Nodes("Art").Nodes.Add("Art3", "Art3")
                                rootNode.Nodes("HyperSpin").Nodes("Art").Nodes("Art3").Tag = gHS.Systems(pSystem).Roms(pRom).Art3Path
                            End If
                            If gHS.Systems(pSystem).Roms(pRom).HasArt4 Then
                                Dim fs As System.IO.FileStream
                                fs = New System.IO.FileStream(gHS.Systems(pSystem).Roms(pRom).Art4Path, IO.FileMode.Open, IO.FileAccess.Read)
                                PictureBox5.Image = System.Drawing.Image.FromStream(fs)
                                fs.Close()
                                'PictureBox5.ImageLocation = gHS.Systems(pSystem).Roms(pRom).Art4Path
                                rootNode.Nodes("HyperSpin").Nodes("Art").Nodes.Add("Art4", "Art4")
                                rootNode.Nodes("HyperSpin").Nodes("Art").Nodes("Art4").Tag = gHS.Systems(pSystem).Roms(pRom).Art4Path
                            End If
                        End If

                        If gHS.Systems(pSystem).Roms(pRom).HasWheel Then
                            Dim fs As System.IO.FileStream
                            fs = New System.IO.FileStream(gHS.Systems(pSystem).Roms(pRom).WheelPath, IO.FileMode.Open, IO.FileAccess.Read)
                            PictureBox1.Image = System.Drawing.Image.FromStream(fs)
                            fs.Close()
                            'PictureBox1.ImageLocation = gHS.Systems(pSystem).Roms(pRom).WheelPath
                            rootNode.Nodes("HyperSpin").Nodes.Add("Wheel", "Wheel")
                            rootNode.Nodes("HyperSpin").Nodes("Wheel").Tag = gHS.Systems(pSystem).Roms(pRom).WheelPath
                        End If

                        If gHS.Systems(pSystem).Roms(pRom).HasVideo Then
                            rootNode.Nodes("HyperSpin").Nodes.Add("Video", "Video")
                            rootNode.Nodes("HyperSpin").Nodes("Video").Tag = gHS.Systems(pSystem).Roms(pRom).VideoPath
                        End If

                        rootNode.Nodes.Add("HyperPause", "HyperPause")

                        If gHS.Systems(pSystem).Roms(pRom).HasBezels Then
                            rootNode.Nodes("HyperPause").Nodes.Add("Bezels", "Bezels")
                            rootNode.Nodes("HyperPause").Nodes("Bezels").Tag = gHS.Systems(pSystem).Roms(pRom).BezelsPath
                            For Each pDir In Directory.GetDirectories(gHS.Systems(pSystem).Roms(pRom).BezelsPath)
                                Dim vLastDir As String = pDir.Split("\").Last
                                rootNode.Nodes("HyperPause").Nodes("Bezels").Nodes.Add(pDir, vLastDir)
                                rootNode.Nodes("HyperPause").Nodes("Bezels").Nodes(pDir).Tag = pDir.Replace("\\", "\") & "\"
                                For Each pFile In Directory.GetFiles(pDir)
                                    If Path.GetFileName(pFile) <> "Thumbs.db" Then
                                        rootNode.Nodes("HyperPause").Nodes("Bezels").Nodes(pDir).Nodes.Add(pFile, Path.GetFileName(pFile))
                                        rootNode.Nodes("HyperPause").Nodes("Bezels").Nodes(pDir).Nodes(pFile).Tag = Path.GetFullPath(pFile)
                                    End If
                                Next
                            Next
                            For Each pFile In Directory.GetFiles(gHS.Systems(pSystem).Roms(pRom).BezelsPath)
                                If Path.GetFileName(pFile) <> "Thumbs.db" Then
                                    rootNode.Nodes("HyperPause").Nodes("Bezels").Nodes.Add(pFile, Path.GetFileName(pFile))
                                    rootNode.Nodes("HyperPause").Nodes("Bezels").Nodes(pFile).Tag = Path.GetFullPath(pFile)
                                End If
                            Next
                        End If

                        If gHS.Systems(pSystem).Roms(pRom).HasArtwork Then
                            rootNode.Nodes("HyperPause").Nodes.Add("Artwork", "Artwork")
                            rootNode.Nodes("HyperPause").Nodes("Artwork").Tag = gHS.Systems(pSystem).Roms(pRom).ArtworkPath
                            For Each pDir In Directory.GetDirectories(gHS.Systems(pSystem).Roms(pRom).ArtworkPath)
                                Dim vLastDir As String = pDir.Split("\").Last
                                rootNode.Nodes("HyperPause").Nodes("Artwork").Nodes.Add(pDir, vLastDir)
                                rootNode.Nodes("HyperPause").Nodes("Artwork").Nodes(pDir).Tag = pDir.Replace("\\", "\") & "\"
                                For Each pFile In Directory.GetFiles(pDir)
                                    If Path.GetFileName(pFile) <> "Thumbs.db" Then
                                        rootNode.Nodes("HyperPause").Nodes("Artwork").Nodes(pDir).Nodes.Add(pFile, Path.GetFileName(pFile))
                                        rootNode.Nodes("HyperPause").Nodes("Artwork").Nodes(pDir).Nodes(pFile).Tag = Path.GetFullPath(pFile)
                                    End If
                                Next
                            Next

                            For Each pFile In Directory.GetFiles(gHS.Systems(pSystem).Roms(pRom).ArtworkPath)
                                If Path.GetFileName(pFile) <> "Thumbs.db" Then
                                    rootNode.Nodes("HyperPause").Nodes("Artwork").Nodes.Add(pFile, Path.GetFileName(pFile))
                                    rootNode.Nodes("HyperPause").Nodes("Artwork").Nodes(pFile).Tag = Path.GetFullPath(pFile)
                                End If
                            Next
                        End If

                        If gHS.Systems(pSystem).Roms(pRom).HasController Then
                            rootNode.Nodes("HyperPause").Nodes.Add("Controller", "Controller")
                            rootNode.Nodes("HyperPause").Nodes("Controller").Tag = gHS.Systems(pSystem).Roms(pRom).ControllerPath
                            For Each pFile In Directory.GetFiles(gHS.Systems(pSystem).Roms(pRom).ControllerPath)
                                If Path.GetFileName(pFile) <> "Thumbs.db" Then
                                    rootNode.Nodes("HyperPause").Nodes("Controller").Nodes.Add(pFile, Path.GetFileName(pFile))
                                    rootNode.Nodes("HyperPause").Nodes("Controller").Nodes(pFile).Tag = Path.GetFullPath(pFile)
                                End If
                            Next
                        End If

                        If gHS.Systems(pSystem).Roms(pRom).HasBackgrounds Then
                            rootNode.Nodes("HyperPause").Nodes.Add("Background", "Background")
                            rootNode.Nodes("HyperPause").Nodes("Background").Tag = gHS.Systems(pSystem).Roms(pRom).BackgroundsPath
                            For Each pFile In Directory.GetFiles(gHS.Systems(pSystem).Roms(pRom).BackgroundsPath)
                                If Path.GetFileName(pFile) <> "Thumbs.db" Then
                                    rootNode.Nodes("HyperPause").Nodes("Background").Nodes.Add(pFile, Path.GetFileName(pFile))
                                    rootNode.Nodes("HyperPause").Nodes("Background").Nodes(pFile).Tag = Path.GetFullPath(pFile)
                                End If
                            Next
                        End If

                        If gHS.Systems(pSystem).Roms(pRom).HasFade Then
                            rootNode.Nodes("HyperPause").Nodes.Add("Fade", "Fade")
                            rootNode.Nodes("HyperPause").Nodes("Fade").Tag = gHS.Systems(pSystem).Roms(pRom).FadePath
                            For Each pFile In Directory.GetFiles(gHS.Systems(pSystem).Roms(pRom).FadePath)
                                If Path.GetFileName(pFile) <> "Thumbs.db" Then
                                    rootNode.Nodes("HyperPause").Nodes("Fade").Nodes.Add(pFile, Path.GetFileName(pFile))
                                    rootNode.Nodes("HyperPause").Nodes("Fade").Nodes(pFile).Tag = Path.GetFullPath(pFile)
                                End If
                            Next
                        End If

                        If gHS.Systems(pSystem).Roms(pRom).HasFonts Then
                            rootNode.Nodes("HyperPause").Nodes.Add("Fonts", "Fonts")
                            rootNode.Nodes("HyperPause").Nodes("Fonts").Tag = gHS.Systems(pSystem).Roms(pRom).FontsPath
                            For Each pFile In Directory.GetFiles(gHS.Systems(pSystem).Roms(pRom).FontsPath)
                                If Path.GetFileName(pFile) <> "Thumbs.db" Then
                                    rootNode.Nodes("HyperPause").Nodes("Fonts").Nodes.Add(pFile, Path.GetFileName(pFile))
                                    rootNode.Nodes("HyperPause").Nodes("Fonts").Nodes(pFile).Tag = Path.GetFullPath(pFile)
                                End If
                            Next
                        End If

                        If gHS.Systems(pSystem).Roms(pRom).HasGuides Then
                            rootNode.Nodes("HyperPause").Nodes.Add("Guides", "Guides")
                            rootNode.Nodes("HyperPause").Nodes("Guides").Tag = gHS.Systems(pSystem).Roms(pRom).GuidesPath
                            For Each pDir In Directory.GetDirectories(gHS.Systems(pSystem).Roms(pRom).GuidesPath)
                                Dim vLastDir As String = pDir.Split("\").Last
                                rootNode.Nodes("HyperPause").Nodes("Guides").Nodes.Add(pDir, vLastDir)
                                rootNode.Nodes("HyperPause").Nodes("Guides").Nodes(pDir).Tag = pDir.Replace("\\", "\") & "\"
                                For Each pFile In Directory.GetFiles(pDir)
                                    If Path.GetFileName(pFile) <> "Thumbs.db" Then
                                        rootNode.Nodes("HyperPause").Nodes("Guides").Nodes(pDir).Nodes.Add(pFile, Path.GetFileName(pFile))
                                        rootNode.Nodes("HyperPause").Nodes("Guides").Nodes(pDir).Nodes(pFile).Tag = Path.GetFullPath(pFile)
                                    End If
                                Next
                            Next

                            For Each pFile In Directory.GetFiles(gHS.Systems(pSystem).Roms(pRom).GuidesPath)
                                If Path.GetFileName(pFile) <> "Thumbs.db" Then
                                    rootNode.Nodes("HyperPause").Nodes("Guides").Nodes.Add(pFile, Path.GetFileName(pFile))
                                    rootNode.Nodes("HyperPause").Nodes("Guides").Nodes(pFile).Tag = Path.GetFullPath(pFile)
                                End If
                            Next
                        End If

                        If gHS.Systems(pSystem).Roms(pRom).HasManuals Then
                            rootNode.Nodes("HyperPause").Nodes.Add("Manuals", "Manuals")
                            rootNode.Nodes("HyperPause").Nodes("Manuals").Tag = gHS.Systems(pSystem).Roms(pRom).ManualsPath
                            For Each pDir In Directory.GetDirectories(gHS.Systems(pSystem).Roms(pRom).ManualsPath)
                                Dim vLastDir As String = pDir.Split("\").Last
                                rootNode.Nodes("HyperPause").Nodes("Manuals").Nodes.Add(pDir, vLastDir)
                                rootNode.Nodes("HyperPause").Nodes("Manuals").Nodes(pDir).Tag = pDir.Replace("\\", "\") & "\"
                                For Each pFile In Directory.GetFiles(pDir)
                                    If Path.GetFileName(pFile) <> "Thumbs.db" Then
                                        rootNode.Nodes("HyperPause").Nodes("Manuals").Nodes(pDir).Nodes.Add(pFile, Path.GetFileName(pFile))
                                        rootNode.Nodes("HyperPause").Nodes("Manuals").Nodes(pDir).Nodes(pFile).Tag = Path.GetFullPath(pFile)
                                    End If
                                Next
                            Next

                            For Each pFile In Directory.GetFiles(gHS.Systems(pSystem).Roms(pRom).ManualsPath)
                                If Path.GetFileName(pFile) <> "Thumbs.db" Then
                                    rootNode.Nodes("HyperPause").Nodes("Manuals").Nodes.Add(pFile, Path.GetFileName(pFile))
                                    rootNode.Nodes("HyperPause").Nodes("Manuals").Nodes(pFile).Tag = Path.GetFullPath(pFile)
                                End If
                            Next
                        End If

                        If gHS.Systems(pSystem).Roms(pRom).HasMultiGame Then
                            rootNode.Nodes("HyperPause").Nodes.Add("Multigame", "Multigame")
                            rootNode.Nodes("HyperPause").Nodes("Multigame").Tag = gHS.Systems(pSystem).Roms(pRom).MultiGamePath
                            For Each pFile In Directory.GetFiles(gHS.Systems(pSystem).Roms(pRom).MultiGamePath)
                                If Path.GetFileName(pFile) <> "Thumbs.db" Then
                                    rootNode.Nodes("HyperPause").Nodes("Multigame").Nodes.Add(pFile, Path.GetFileName(pFile))
                                    rootNode.Nodes("HyperPause").Nodes("Multigame").Nodes(pFile).Tag = Path.GetFullPath(pFile)
                                End If
                            Next
                        End If

                        If gHS.Systems(pSystem).Roms(pRom).HasMusic Then
                            rootNode.Nodes("HyperPause").Nodes.Add("Music", "Music")
                            rootNode.Nodes("HyperPause").Nodes("Music").Tag = gHS.Systems(pSystem).Roms(pRom).MusicPath
                            For Each pFile In Directory.GetFiles(gHS.Systems(pSystem).Roms(pRom).MusicPath)
                                If Path.GetFileName(pFile) <> "Thumbs.db" Then
                                    rootNode.Nodes("HyperPause").Nodes("Music").Nodes.Add(pFile, Path.GetFileName(pFile))
                                    rootNode.Nodes("HyperPause").Nodes("Music").Nodes(pFile).Tag = Path.GetFullPath(pFile)
                                End If
                            Next
                        End If

                        If gHS.Systems(pSystem).Roms(pRom).HasSounds Then
                            rootNode.Nodes("HyperPause").Nodes.Add("Sounds", "Sounds")
                            rootNode.Nodes("HyperPause").Nodes("Sounds").Tag = gHS.Systems(pSystem).Roms(pRom).SoundsPath
                            For Each pFile In Directory.GetFiles(gHS.Systems(pSystem).Roms(pRom).SoundsPath)
                                If Path.GetFileName(pFile) <> "Thumbs.db" Then
                                    rootNode.Nodes("HyperPause").Nodes("Sounds").Nodes.Add(pFile, Path.GetFileName(pFile))
                                    rootNode.Nodes("HyperPause").Nodes("Sounds").Nodes(pFile).Tag = Path.GetFullPath(pFile)
                                End If
                            Next
                        End If

                        If gHS.Systems(pSystem).Roms(pRom).HasHLVideos Then
                            rootNode.Nodes("HyperPause").Nodes.Add("Videos", "Videos")
                            rootNode.Nodes("HyperPause").Nodes("Videos").Tag = gHS.Systems(pSystem).Roms(pRom).HLVideosPath
                            For Each pFile In Directory.GetFiles(gHS.Systems(pSystem).Roms(pRom).HLVideosPath)
                                If Path.GetFileName(pFile) <> "Thumbs.db" Then
                                    rootNode.Nodes("HyperPause").Nodes("Videos").Nodes.Add(pFile, Path.GetFileName(pFile))
                                    rootNode.Nodes("HyperPause").Nodes("Videos").Nodes(pFile).Tag = Path.GetFullPath(pFile)
                                End If
                            Next
                        End If

                        If gHS.Systems(pSystem).Roms(pRom).HasHLWheels Then
                            rootNode.Nodes("HyperPause").Nodes.Add("Wheel", "Wheel")
                            rootNode.Nodes("HyperPause").Nodes("Wheel").Tag = gHS.Systems(pSystem).Roms(pRom).HLWheelsPath
                            For Each pFile In Directory.GetFiles(gHS.Systems(pSystem).Roms(pRom).HLWheelsPath)
                                If Path.GetFileName(pFile) <> "Thumbs.db" Then
                                    rootNode.Nodes("HyperPause").Nodes("Wheel").Nodes.Add(pFile, Path.GetFileName(pFile))
                                    rootNode.Nodes("HyperPause").Nodes("Wheel").Nodes(pFile).Tag = Path.GetFullPath(pFile)
                                End If
                            Next
                        End If
                        TreeViewMedia.Nodes.Add(rootNode)
                        TreeViewMedia.ExpandAll()
                    End If
                Else
                    ToolStripButton10.Enabled = False
                End If
            Catch ex As Exception
                LogEntry(LogType._Error, "Error: " & ex.Message)
                TreeViewMedia.Nodes.Clear()
                WBMedia.Navigate("")
                PictureBox11.Image = Nothing
                ToolStripButton10.Enabled = False
            End Try
        End If
    End Sub

    Private Sub TreeViewMedia_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeViewMedia.AfterSelect
        Dim CurrNode As TreeNode
        CurrNode = TreeViewMedia.SelectedNode

        If CurrNode IsNot Nothing And CurrNode.Tag <> "" Then
            Select Case Path.GetExtension(CurrNode.Tag.ToString).ToLower
                Case ".png", ".bmp", ".gif", ".jpg", ".jpeg"
                    'PictureBox11.ImageLocation = CurrNode.Tag
                    Dim fs As System.IO.FileStream
                    ' Specify a valid picture file path on your computer.
                    fs = New System.IO.FileStream(CurrNode.Tag, IO.FileMode.Open, IO.FileAccess.Read)
                    PictureBox11.Image = System.Drawing.Image.FromStream(fs)
                    fs.Close()

                    PictureBox11.Visible = True
                    WBMedia.Visible = False
                    WBMedia.Navigate("")
                Case Else
                    WBMedia.Navigate(CurrNode.Tag)
                    WBMedia.Visible = True
                    PictureBox11.Visible = False
                    PictureBox11.Image = Nothing
            End Select
        Else
            PictureBox11.ImageLocation = Nothing
            WBMedia.Navigate("")
        End If
    End Sub

    Private Sub TreeViewMedia_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TreeViewMedia.MouseDown
        If TreeViewMedia.GetNodeCount(True) > 0 Then
            TreeViewMedia.ContextMenuStrip = ContextMenuStrip1
        Else
            TreeViewMedia.ContextMenuStrip = Nothing
        End If
    End Sub

    Private Sub TreeViewMedia_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TreeViewMedia.MouseUp
        If TreeViewMedia.GetNodeCount(True) > 0 Then
            If e.Button = MouseButtons.Right Then
                Try
                    If TreeViewMedia.GetNodeAt(New Point(e.X, e.Y)).Tag = "" Or Path.GetExtension(TreeViewMedia.GetNodeAt(New Point(e.X, e.Y)).Tag) = "" Then
                        DeleteSelectedFileToolStripMenuItem.Enabled = False
                        OpenToolStripMenuItem.Enabled = False
                        OpenContainingFolderToolStripMenuItem.Enabled = False
                    Else
                        DeleteSelectedFileToolStripMenuItem.Enabled = True
                        OpenToolStripMenuItem.Enabled = True
                        OpenContainingFolderToolStripMenuItem.Enabled = True
                    End If
                Catch ex As Exception
                    DeleteSelectedFileToolStripMenuItem.Enabled = False
                    OpenToolStripMenuItem.Enabled = False
                    OpenContainingFolderToolStripMenuItem.Enabled = False
                End Try
                TreeViewMedia.ContextMenuStrip.Show(TreeViewMedia, New Point(e.X, e.Y))
            End If
        End If
    End Sub

    'Double click sur Media node
    Private Sub TreeViewMedia_NodeMouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeNodeMouseClickEventArgs) Handles TreeViewMedia.NodeMouseDoubleClick
        Try
            Dim psi As New ProcessStartInfo
            Dim p As New Process
            psi.FileName = e.Node.Tag
            p.StartInfo = psi
            p.Start()
        Catch ex As Exception
            LogEntry(LogType._Error, "Error: " & ex.Message)
        End Try
    End Sub

    Private Sub DataGridView2_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView2.SelectionChanged
        Dim CurrentSystem As String
        'TREEVIEW MEDIA_SYSTEMS (treeviewSystems)
        PictureBox16.Image = Nothing
        PictureBox15.Image = Nothing
        PictureBox14.Image = Nothing
        PictureBox13.Image = Nothing
        PictureBox12.Image = Nothing
        WBSystems.Navigate("")
        'Application.DoEvents()

        TreeviewSystems.Nodes.Clear()
        '        If IsLoading = False Then
        'If DataGridView2.Enabled = True Then
        Try
            CurrentSystem = DataGridView2.SelectedRows.Item(0).Cells(0).Value
        Catch ex As Exception
            Exit Sub
        End Try

        'WHEEL
        Try
            If DataGridView2.SelectedRows.Item(0).Cells(2).Tag <> "" Then

                Dim fs As System.IO.FileStream
                ' Specify a valid picture file path on your computer.
                fs = New System.IO.FileStream(DataGridView2.SelectedRows.Item(0).Cells(2).Tag, IO.FileMode.Open, IO.FileAccess.Read)
                PictureBox15.Image = System.Drawing.Image.FromStream(fs)
                fs.Close()

            Else
                PictureBox15.Image = Nothing
            End If
        Catch ex As Exception
            PictureBox15.Image = Nothing
        End Try

        'BataBriteC
        Try
            If DataGridView2.SelectedRows.Item(0).Cells(5).Tag <> "" Then
                Dim fs As System.IO.FileStream
                ' Specify a valid picture file path on your computer.
                fs = New System.IO.FileStream(DataGridView2.SelectedRows.Item(0).Cells(5).Tag, IO.FileMode.Open, IO.FileAccess.Read)
                PictureBox14.Image = System.Drawing.Image.FromStream(fs)
                fs.Close()
            Else
                PictureBox14.Image = Nothing
            End If
        Catch ex As Exception
            PictureBox14.Image = Nothing
        End Try

        'BataBriteP
        Try
            If DataGridView2.SelectedRows.Item(0).Cells(6).Tag <> "" Then
                Dim fs As System.IO.FileStream
                ' Specify a valid picture file path on your computer.
                fs = New System.IO.FileStream(DataGridView2.SelectedRows.Item(0).Cells(6).Tag, IO.FileMode.Open, IO.FileAccess.Read)
                PictureBox13.Image = System.Drawing.Image.FromStream(fs)
                fs.Close()
            Else
                PictureBox13.Image = Nothing
            End If
        Catch ex As Exception
            PictureBox13.Image = Nothing
        End Try

        'Pointer
        Try
            Dim PointerPath As String = gHSPath & "\Media\" & CurrentSystem & "\Images\Other\Pointer.png"
            If File.Exists(PointerPath) Then
                Dim fs As System.IO.FileStream
                ' Specify a valid picture file path on your computer.
                fs = New System.IO.FileStream(PointerPath, IO.FileMode.Open, IO.FileAccess.Read)
                PictureBox12.Image = System.Drawing.Image.FromStream(fs)
                fs.Close()
            Else
                PictureBox12.Image = Nothing
            End If
        Catch ex As Exception
            PictureBox12.Image = Nothing
        End Try

        'TreeviewSystems
        If DataGridView2.SelectedRows.Count > 0 Then
            Try
                Dim pRow As DataGridViewRow = DataGridView2.SelectedRows.Item(0)
                Dim rootNode As New TreeNode(CurrentSystem)

                rootNode.Nodes.Add("HyperSpin", "HyperSpin")
                rootNode.Nodes.Add("HyperPause", "HyperPause")

                If gHS.Systems(CurrentSystem).HasTheme Then
                    rootNode.Nodes("HyperSpin").Nodes.Add("Theme", "Theme")
                    rootNode.Nodes("HyperSpin").Nodes("Theme").Tag = Path.GetDirectoryName(gHS.Systems(CurrentSystem).ThemePath)
                End If

                If gHS.Systems(CurrentSystem).HasTheme Then
                    rootNode.Nodes("HyperSpin").Nodes.Add("Wheel", "Wheel")
                    rootNode.Nodes("HyperSpin").Nodes("Wheel").Tag = gHS.Systems(CurrentSystem).WheelPath
                End If

                If gHS.Systems(CurrentSystem).HasVideo Then
                    rootNode.Nodes("HyperSpin").Nodes.Add("Video", "Video")
                    rootNode.Nodes("HyperSpin").Nodes("Video").Tag = gHS.Systems(CurrentSystem).VideoPath
                End If

                If gHS.Systems(CurrentSystem).HasBetaBriteClassic Then
                    rootNode.Nodes("HyperSpin").Nodes.Add("Betabrite Classic", "Betabrite Classic")
                    rootNode.Nodes("HyperSpin").Nodes("Betabrite Classic").Tag = gHS.Systems(CurrentSystem).BetaBriteClassicPath
                End If

                If gHS.Systems(CurrentSystem).HasBetaBritePrism Then
                    rootNode.Nodes("HyperSpin").Nodes.Add("Betabrite Prism", "Betabrite Prism")
                    rootNode.Nodes("HyperSpin").Nodes("Betabrite Prism").Tag = gHS.Systems(CurrentSystem).BetaBritePrismPath
                End If

                If gHS.Systems(CurrentSystem).HasHPBackground Then
                    rootNode.Nodes("HyperPause").Nodes.Add("Background", "Background")
                    rootNode.Nodes("HyperPause").Nodes("Background").Tag = gHS.Systems(CurrentSystem).HPBackgroundFolder
                    For Each pFile In Directory.GetFiles(gHS.Systems(CurrentSystem).HPBackgroundFolder)
                        If Path.GetFileName(pFile) <> "Thumbs.db" Then
                            rootNode.Nodes("HyperPause").Nodes("Background").Nodes.Add(pFile, Path.GetFileName(pFile))
                            rootNode.Nodes("HyperPause").Nodes("Background").Nodes(pFile).Tag = Path.GetFullPath(pFile)
                        End If
                    Next
                End If

                If gHS.Systems(CurrentSystem).HasHPFade Then
                    rootNode.Nodes("HyperPause").Nodes.Add("Fade", "Fade")
                    rootNode.Nodes("HyperPause").Nodes("Fade").Tag = gHS.Systems(CurrentSystem).HPFadeFolder
                    For Each pFile In Directory.GetFiles(gHS.Systems(CurrentSystem).HPFadeFolder)
                        If Path.GetFileName(pFile) <> "Thumbs.db" Then
                            rootNode.Nodes("HyperPause").Nodes("Fade").Nodes.Add(pFile, Path.GetFileName(pFile))
                            rootNode.Nodes("HyperPause").Nodes("Fade").Nodes(pFile).Tag = Path.GetFullPath(pFile)
                        End If
                    Next
                End If

                If gHS.Systems(CurrentSystem).HasHPBezel Then
                    rootNode.Nodes("HyperPause").Nodes.Add("Bezel", "Bezel")
                    rootNode.Nodes("HyperPause").Nodes("Bezel").Tag = gHS.Systems(CurrentSystem).HPBezelFolder
                    For Each pDir In Directory.GetDirectories(gHS.Systems(CurrentSystem).HPBezelFolder)
                        Dim vLastDir As String = pDir.Split("\").Last
                        rootNode.Nodes("HyperPause").Nodes("Bezel").Nodes.Add(pDir, vLastDir)
                        rootNode.Nodes("HyperPause").Nodes("Bezel").Nodes(pDir).Tag = pDir.Replace("\\", "\") & "\"
                        For Each pFile In Directory.GetFiles(pDir)
                            If Path.GetFileName(pFile) <> "Thumbs.db" Then
                                rootNode.Nodes("HyperPause").Nodes("Bezel").Nodes(pDir).Nodes.Add(pFile, Path.GetFileName(pFile))
                                rootNode.Nodes("HyperPause").Nodes("Bezel").Nodes(pDir).Nodes(pFile).Tag = Path.GetFullPath(pFile)
                            End If
                        Next
                    Next
                    'If gHS.Systems(CurrentSystem).HasHPBezel Then
                    '    rootNode.Nodes("HyperPause").Nodes.Add("Bezel", "Bezel")
                    '    rootNode.Nodes("HyperPause").Nodes("Bezel").Tag = gHS.Systems(CurrentSystem).HPBezelFolder
                    For Each pFile In Directory.GetFiles(gHS.Systems(CurrentSystem).HPBezelFolder)
                        If Path.GetFileName(pFile) <> "Thumbs.db" Then
                            rootNode.Nodes("HyperPause").Nodes("Bezel").Nodes.Add(pFile, Path.GetFileName(pFile))
                            rootNode.Nodes("HyperPause").Nodes("Bezel").Nodes(pFile).Tag = Path.GetFullPath(pFile)
                        End If
                    Next
                    'End If
                End If

                TreeviewSystems.Nodes.Add(rootNode)
                TreeviewSystems.ExpandAll()
            Catch ex As Exception
                LogEntry(LogType._Error, "Error: " & ex.Message)
                TreeviewSystems.Nodes.Clear()
                WBSystems.Navigate("")
                PictureBox16.Image = Nothing
            End Try
        End If

        Application.DoEvents()
        'End If
        'End If
    End Sub

    'SAVE CHANGES (MainMenu.xml)
    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        If MsgBox("Are you sure ?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            DataGridView2.CommitEdit(DataGridViewDataErrorContexts.Commit)
            'Make backup of old XML
            Dim FileSave As String = gDBPath & "\Main Menu\Main Menu_" & Date.Now.ToString("yyyy-MM-dd_HH-mm-ss") & ".xml"
            If File.Exists(FileSave) Then
                File.Delete(FileSave)
                Application.DoEvents()
            End If
            File.Move(gMainMenuXML, FileSave)
            Application.DoEvents()
            'file.delete(gMainMenuXML)
            'Application.DoEvents()
            'Recreate new XML
            Dim objWriter As New System.IO.StreamWriter(gMainMenuXML)
            Try
                objWriter.WriteLine("<menu>")
                For Each vRow As DataGridViewRow In DataGridView2.Rows
                    If vRow.Cells(7).Value = False Then
                        If vRow.Cells(1).Value = True Then
                            objWriter.WriteLine("  <game name=""" & EncodeForXml(vRow.Cells(0).Value) & """/>")
                        End If
                    Else
                        If vRow.Cells(1).Value = True Then
                            objWriter.WriteLine("  <game name=""" & EncodeForXml(vRow.Cells(0).Value) & """ exe=""true""/>")
                        End If
                    End If
                Next
                objWriter.WriteLine("</menu>")
                '<menu>
                '  <game name="ScummVM"/>
                '</menu>
            Catch ex As Exception
                ToolStripStatusLabel1.Text = ex.Message.ToString
                LogEntry(LogType._Error, "Warning, error in writing Mail Menu.xml !")
            Finally
                objWriter.Close()
                ToolStripStatusLabel1.Text = """Main Menu.xml"" backuped, and new changes applied."
            End Try
            'Set system entry in treeview in normal font or reload systems (easier)
            'LoadSystems()
            'reload stats if any (system stats)

            Application.DoEvents()
            LoadSystems()
        End If
    End Sub

    Private Sub TreeView1_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeView1.AfterSelect
        Try
            If TabControl1.SelectedIndex = 0 Then
                If e.Node.Parent.Text = "Systems" Then
                    TabControl1.SelectedIndex = 1
                End If
            End If
        Catch ex As Exception

        End Try

        If ToolStripLabel1.Text <> e.Node.Text Then
            'desactivate all roms & media check
        End If

        If ToolStripLabel2.Text <> e.Node.Text Then
            'desactivate roms XML
            ToolStripButton6.Enabled = False
        End If

    End Sub

    'Load ROMs XML
    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        LoadFULLSystem(TreeView1.SelectedNode.Text)
    End Sub

    'Check load all DB folders
    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        gDisplayAllSystems = CheckBox1.Checked

        If IsLoading = False Then
            SaveHyperT00lsParameters()
            Try
                If gHS.Systems.Count > 0 Then
                    LoadSystems()

                End If
            Catch ex As Exception

            End Try
        End If



    End Sub

    'Save <databases\system.xml>
    Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click
        'If MsgBox("Are you sure ?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
        If ToolStripLabel2.Text <> "" And ToolStripLabel2.Text <> "Systems" And ToolStripLabel2.Text <> "System" Then
            DataGridView4.CommitEdit(DataGridViewDataErrorContexts.Commit)
            Dim vTime As String = Date.Now.ToString("yyyy-MM-dd_HH-mm-ss")
            Dim vDate As String = Date.Now.ToString("MM/dd/yyyy")

            Dim vSystem As String = ToolStripLabel2.Text
            'Make backup of old XML
            Dim FileSave As String = gDBPath & "\" & vSystem & "\" & vSystem & "_" & vTime & ".xml"
            If File.Exists(FileSave) Then
                File.Delete(FileSave)
                Application.DoEvents()
            End If
            File.Move(gHS.Systems(vSystem).XMLPath, FileSave)
            Application.DoEvents()
            'Recreate new XML
            Dim objWriter As New System.IO.StreamWriter(gHS.Systems(vSystem).XMLPath)
            Try
                objWriter.WriteLine("<?xml version=""1.0""?>")
                objWriter.WriteLine("<menu>")
                objWriter.WriteLine("  <header>")
                objWriter.WriteLine("    <listname>" & vSystem & "</listname>")
                objWriter.WriteLine("    <lastlistupdate>" & vDate & "</lastlistupdate>")
                objWriter.WriteLine("    <listversion>HyperT00ls " & vTime & "</listversion>")
                objWriter.WriteLine("    <exporterversion>Generated by HyperT00ls v" & Application.ProductVersion & "</exporterversion>")
                objWriter.WriteLine("  </header>")
                For Each vRow As DataGridViewRow In DataGridView4.Rows
                    If vRow.Cells(0).Value <> "" Then
                        Dim vAdd As String = ""
                        'index="true" image="'"
                        If vRow.Cells(10).Value = True Then
                            vAdd = " index=""true"""
                        Else
                            vAdd = " index="""""
                        End If
                        If vRow.Cells(9).Value <> "" Then
                            vAdd = vAdd & " image=""" & vRow.Cells(9).Value.ToString & """"
                        Else
                            vAdd = vAdd & " image="""""
                        End If

                        'Game name
                        Dim vVal As String = ""
                        Try
                            vVal = EncodeForXml(vRow.Cells(0).Value.ToString)
                        Catch ex As Exception
                            vVal = ""
                        Finally
                            objWriter.WriteLine("  <game name=""" & vVal & """" & vAdd & ">")
                        End Try

                        'Description
                        Try
                            vVal = EncodeForXml(vRow.Cells(1).Value.ToString)
                        Catch ex As Exception
                            vVal = ""
                        Finally
                            objWriter.WriteLine("    <description>" & vVal & "</description>")
                        End Try

                        Try
                            vVal = vRow.Cells(7).Value.ToString
                        Catch ex As Exception
                            vVal = ""
                        Finally
                            objWriter.WriteLine("    <cloneof>" & vVal & "</cloneof>")
                        End Try


                        Try
                            vVal = vRow.Cells(6).Value.ToString
                        Catch ex As Exception
                            vVal = ""
                        Finally
                            objWriter.WriteLine("    <crc>" & vVal & "</crc>")
                        End Try


                        Try
                            vVal = EncodeForXml(vRow.Cells(2).Value.ToString)
                        Catch ex As Exception
                            vVal = ""
                        Finally
                            objWriter.WriteLine("    <manufacturer>" & vVal & "</manufacturer>")
                        End Try


                        Try
                            vVal = EncodeForXml(vRow.Cells(5).Value.ToString)
                        Catch ex As Exception
                            vVal = ""
                        Finally
                            objWriter.WriteLine("    <rating>" & vVal & "</rating>")
                        End Try


                        Try
                            vVal = EncodeForXml(vRow.Cells(3).Value.ToString)
                        Catch ex As Exception
                            vVal = ""
                        Finally
                            objWriter.WriteLine("    <year>" & vVal & "</year>")
                        End Try


                        Try
                            vVal = EncodeForXml(vRow.Cells(4).Value.ToString)
                        Catch ex As Exception
                            vVal = ""
                        Finally
                            objWriter.WriteLine("    <genre>" & vVal & "</genre>")
                        End Try


                        If vRow.Cells(8).Value = False Then
                            objWriter.WriteLine("    <enabled>No</enabled>")
                        Else
                            objWriter.WriteLine("    <enabled>Yes</enabled>")
                        End If
                        objWriter.WriteLine("  </game>")
                    End If
                Next
                objWriter.WriteLine("</menu>")
                'gHS.ReloadSystem(vSystem)

                'NEED TO BE CHECKED

                '<menu>
                '<game name="Alien vs Predator (USA, Europe)">
                '	<description>Alien Vs Predator</description>
                '	<cloneof></cloneof>
                '	<crc></crc>
                '	<manufacturer>Atari</manufacturer>
                '	<year>1994</year>
                '	<rating>xxxxx</rating>
                '	<genre></genre>
                '</game>
                '(...)
                '</menu>
            Catch ex As Exception
                ToolStripStatusLabel1.Text = ex.Message.ToString
            Finally
                objWriter.Close()
                Application.DoEvents()
                'gHS.Systems(vSystem).LoadSystemRoms(False)
                ToolStripStatusLabel1.Text = """" & vSystem & ".xml"" backuped, and new changes applied."
            End Try
            'Set system entry in treeview in normal font or reload systems (easier)
            'LoadSystems()
            'reload stats if any (system stats)

            Application.DoEvents()
        Else
            MsgBox("Please select a valid system ...", MsgBoxStyle.Information, "Warning")
        End If
    End Sub

    'Reload systems
    Private Sub ToolStripButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton7.Click
        LoadSystems()
    End Sub

    'Enable debug logging
    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        'Console.WriteLine("DebugMode = " & gDebugLog.ToString & " --> CheckBoxDebug=" & gCheckRoms.ToString)
        gDebugLog = CheckBox2.Checked

        'If IsLoading = False Then
        '    SaveHyperT00lsParameters()

        '    Try
        '        If gHS.Systems.Count > 0 Then
        '            LoadSystems()
        '        End If
        '    Catch ex As Exception

        '    End Try
        'End If
    End Sub

    'Launch HyperSpin
    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        Try
            Process.Start(gHS.HSRootPath & "\Hyperspin.exe")
        Catch ex As Exception
            LogEntry(LogType._Error, "Cannot launch " & gHS.HSRootPath & "\Hyperspin.exe : " & ex.Message.ToString)
        End Try
    End Sub

    'Launch HyperLaunchHQ
    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        Try
            Process.Start(gHLPath & "\HyperLaunchHQ\HyperLaunchHQ.exe")
        Catch ex As Exception
            LogEntry(LogType._Error, "Cannot launch " & gHLPath & "\HyperLaunchHQ\HyperLaunchHQ.exe : " & ex.Message.ToString)
        End Try
    End Sub

    'Launch HyperHQ
    Private Sub ToolStripButton8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton8.Click
        Try
            Process.Start(gHS.HSRootPath & "\HyperHQ.exe")
        Catch ex As Exception
            LogEntry(LogType._Error, "Cannot launch " & gHS.HSRootPath & "\HyperHQ.exe here : " & ex.Message.ToString)
        End Try
    End Sub

    'delete xml entry
    Private Sub Button7_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Try
            DataGridView4.Rows.RemoveAt(DataGridView4.SelectedCells(0).RowIndex)
            DataGridView4.CommitEdit(DataGridViewDataErrorContexts.RowDeletion)
            Label1.Text = "DO NOT FORGET TO SAVE !"
            'Call ToolStripButton6.PerformClick()
        Catch ex As Exception
            LogEntry(LogType._Error, ex.Message.ToString)
        End Try

    End Sub

    'Choose HyperSpin Path
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim GetHSPath As New FolderBrowserDialog
        GetHSPath.Description = "Hyperspin Folder"
        GetHSPath.RootFolder = Environment.SpecialFolder.MyComputer
        If GetHSPath.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Try
                gHSPath = GetHSPath.SelectedPath
                TextBox2.Text = gHSPath
            Catch Ex As Exception
            Finally
                SaveHyperT00lsParameters()
            End Try
        End If

        LoadConfigFile()
        If File.Exists(gMainMenuXML) Then
            LoadSystems()
        Else
            MsgBox("Sorry but there is an error ! Cannot find : " & gMainMenuXML & " !")
        End If

    End Sub

    'Enable/disable check of roms
    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged
        Console.WriteLine("IsLoading = " & IsLoading.ToString & " --> CheckOfRoms=" & gCheckRoms.ToString)
        gCheckRoms = CheckBox3.Checked

        If gCheckRoms = False Then
            Button13.Enabled = False
        End If

        If IsLoading = False Then

            SaveHyperT00lsParameters()

            'Try
            '    If gHS.Systems.Count > 0 Then
            '        LoadSystems()
            '    End If
            'Catch ex As Exception

            'End Try
        End If
    End Sub

    'Look for missing roms in XML
    Private Sub Button12_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Cursor = Cursors.WaitCursor
        Button2.Enabled = False
        Button3.Enabled = False
        'List all roms in folders
        'need:
        ' - system.rompath
        ' - system.romextension(s)
        ' - system.lookinsubfolders
        ' - all xml rom names
        ListBox1.Items.Clear()

        Dim vSystem As String = ToolStripLabel2.Text
        'ToolStripLabel2.Text = vSystem

        If DataGridView4.Rows.Count < 2 Then
            ToolStripButton5.PerformClick()
        End If

        Select Case vSystem
            Case "System", "Systems", "Support", "HyperT00ls", "Hyperspin-FE", "Emumovies"
                'MsgBox("Please select a valid System on the left." & vbCrLf & "Currently selected (" & vSystem & ") is not valid.", MsgBoxStyle.Information, "Information")
                'ToolStripButton5.PerformClick()
                'Button12.PerformClick()
            Case Else
                'First, list ALL roms
                'Dim RomList() As String
                Dim vErrors As String = ""
                If gHS.Systems(vSystem).mRomExtensions = "" Then
                    vErrors = "No rom extensions have been configured for this system (" & gHS.Systems(vSystem).Name & ") !"
                End If
                If gHS.Systems(vSystem).mRomPath = "" Then
                    If vErrors = "" Then
                        vErrors = "No rom path has been configured for this system (" & gHS.Systems(vSystem).Name & ") !"
                    Else
                        vErrors = vErrors & vbCrLf & "No rom path has been configured for this system (" & gHS.Systems(vSystem).Name & ") !"
                    End If
                    vErrors = vErrors & vbCrLf & "Please CORRECT this using HyperHQ"
                End If
                If vErrors <> "" Then
                    LogEntry(LogType._Error, "Error occured while searching --> " & vErrors)
                    'MsgBox(vErrors, MsgBoxStyle.Critical, "Errors")
                    Label1.Text = vErrors
                    Cursor = Cursors.Default
                    Exit Sub
                End If

                For Each OneExt As String In gHS.Systems(vSystem).mRomExtensions.Split(CChar(","))
                    'SearchRoms(Me.Name & "." & OneExt, gHS.Systems(vSystem).mRomPath, gHS.Systems(vSystem).mSearchSubfolders)
                    If gHS.Systems(vSystem).mSearchSubfolders = True Then
                        GetFiles("*." & Trim(OneExt), gHS.Systems(vSystem).mRomPath, 10000, 0)
                    Else
                        GetFiles("*." & Trim(OneExt), gHS.Systems(vSystem).mRomPath, 10000, 0)
                    End If
                Next

                Dim ListToDelete As New List(Of Object)
                'Then, get romname without path and extension and check VS the XML
                For Each item In ListBox1.Items()
                    'get item "romname" only
                    Dim ppRomName As String = Path.GetFileNameWithoutExtension(item)
                    For Each ppRow As DataGridViewRow In DataGridView4.Rows()
                        Try
                            If ppRow.Cells(0).Value.ToString.ToLower = ppRomName.ToLower Then
                                ListToDelete.Add(item)
                                'ListBox1.Items.Remove(item)
                            End If
                        Catch ex As Exception

                        End Try
                    Next
                Next

                ' Loop through and display each path.
                For Each entry In ListToDelete
                    ListBox1.Items.Remove(entry)
                Next
                'gHS.Systems(vSystem)

        End Select

        If ListBox1.Items.Count > 0 Then
            Button2.Enabled = False
            Button3.Enabled = True
        End If
        Cursor = Cursors.Default
    End Sub

    ' Recursive function which keeps moving down the directory tree until a given depth.
    Public Sub GetFiles(ByVal strFileFilter As String, ByVal strDirectory As String, ByVal intDepthLimit As Integer, ByVal intCurrentDepth As Integer)
        Label1.Text = "Info :"
        Dim folderInfo As New DirectoryInfo(strDirectory)

        ' Is the current depth on this recursion less than our limit?
        ' If so, find any directories and get into them by calling GetFiles recursively (incrementing depth count)
        If intCurrentDepth < intDepthLimit Then
            Dim directories() As DirectoryInfo
            Dim vCurrDir As String = strDirectory
            If Directory.Exists(strDirectory) = False Then
                Label1.Text = "WARNING !!! The directory " & strDirectory & " does NOT exists !!!"
                Exit Sub
            End If
            Try
                directories = folderInfo.GetDirectories()
                For Each fDirectory In directories
                    vCurrDir = fDirectory.FullName
                    ' Recursively call ourselves incrementing the depth using the given folder path.
                    GetFiles(strFileFilter, fDirectory.FullName, intDepthLimit, intCurrentDepth + 1)
                Next
            Catch ex As Exception
                LogEntry(LogType._Error, "Error occured while searching (Dir=" & vCurrDir & ", Filter=" & strFileFilter & ")--> " & ex.Message.ToString)
                StatusStrip1.Text = Trim(ex.Message.ToString)
                'MsgBox("Error during scan ... check LogViewer", MsgBoxStyle.Critical, "Error")
                Exit Sub
            End Try
        End If

        ' After we can't go further down, add any files which match our filter to listbox (in this case lstFiles)
        Dim files() As FileInfo
        files = folderInfo.GetFiles(strFileFilter)

        For Each fFile In files
            ListBox1.Items.Add(fFile.FullName)
        Next
    End Sub

    'Highlight duplicate ROM entries in XML
    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        Label1.Text = "Info :"
        Cursor = Cursors.WaitCursor
        Button2.Enabled = False
        Button3.Enabled = False

        ListBox1.Items.Clear()
        Dim vResult As Integer = 0
        Dim vDoublons As String = ""
        Try
            For intI As Integer = 0 To DataGridView4.Rows.Count - 2
                For intJ As Integer = intI + 1 To DataGridView4.Rows.Count - 2
                    If DataGridView4.Rows(intI).Cells(0).Value = DataGridView4.Rows(intJ).Cells(0).Value Then
                        DataGridView4.Rows(intJ).DefaultCellStyle.BackColor = Color.Orange
                        DataGridView4.Rows(intI).DefaultCellStyle.BackColor = Color.Orange
                        vResult = vResult + 1
                        If vDoublons = "" Then
                            vDoublons = DataGridView4.Rows(intJ).Cells(0).Value.ToString
                        Else
                            vDoublons = vDoublons & "," & DataGridView4.Rows(intJ).Cells(0).Value.ToString
                        End If
                        Cursor = Cursors.Default
                        Exit Sub
                    End If
                Next
            Next
        Catch ex As Exception
            LogEntry(LogType._Error, "Error occured while processing --> " & ex.Message.ToString)
        Finally
            ListBox1.Items.Add("Found " & vResult.ToString & " duplicate roms (highlighted in orange in the list)")
            If vDoublons <> "" Then
                ListBox1.Items.Add("ROMS = " & vDoublons)
            End If
        End Try
        Cursor = Cursors.Default
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        If ListBox1.SelectedItems.Count > 0 Then
            If ListBox1.Items(0).ToString.StartsWith("Found") = False Then
                Button2.Enabled = True
            End If
        End If
    End Sub

    'Add selected ROM to XML
    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Label1.Text = "Info :"
        Try
            Dim rowID As Integer = DataGridView4.Rows.Add()
            Dim xx As String = Path.GetFileNameWithoutExtension(ListBox1.SelectedItems(0).ToString)
            DataGridView4.Rows(rowID).Cells(0).Value = xx
            DataGridView4.Rows(rowID).Cells(1).Value = xx
            DataGridView4.Rows(rowID).Cells(8).Value = True
            DataGridView4.CommitEdit(DataGridViewDataErrorContexts.Commit)
            Label1.Text = "Warning ! DO NOT FORGET TO SAVE !"
            DataGridView4.ClearSelection()
            DataGridView4.Rows(rowID).Cells(0).Selected = True
            'Call ToolStripButton6.PerformClick()
        Catch ex As Exception
            LogEntry(LogType._Error, ex.Message.ToString)
        End Try
    End Sub

    'Add ALL found ROMs to XML
    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Label1.Text = "Info :"
        Dim vFirstRow As Integer = -1

        Try
            For Each vEntry As String In ListBox1.Items
                Dim rowID As Integer = DataGridView4.Rows.Add()
                If vFirstRow = -1 Then
                    vFirstRow = rowID
                End If
                Dim xx As String = Path.GetFileNameWithoutExtension(vEntry)
                DataGridView4.Rows(rowID).Cells(0).Value = xx
                DataGridView4.Rows(rowID).Cells(1).Value = xx
                DataGridView4.Rows(rowID).Cells(8).Value = True
                Label1.Text = "Warning ! DO NOT FORGET TO SAVE !"
                'Call ToolStripButton6.PerformClick()
            Next
            DataGridView4.CommitEdit(DataGridViewDataErrorContexts.Commit)
            DataGridView4.ClearSelection()
            DataGridView4.Rows(vFirstRow).Cells(0).Selected = True
        Catch ex As Exception
            LogEntry(LogType._Error, ex.Message.ToString)
        End Try
    End Sub

#Region "Toolstripmenu"
    'TOOLSTRIP MENU

    'Open file
    Private Sub OpenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenToolStripMenuItem.Click
        If File.Exists(TreeViewMedia.SelectedNode.Tag) Then
            Try
                Dim psi As New ProcessStartInfo
                Dim p As New Process
                psi.FileName = TreeViewMedia.SelectedNode.Tag
                p.StartInfo = psi
                p.Start()
            Catch ex As Exception
                LogEntry(LogType._Error, "Error: " & ex.Message)
            End Try
        Else
            MsgBox("File : " & TreeViewMedia.SelectedNode.Tag & vbCrLf & "Not found !", MsgBoxStyle.Critical, "Error")
            LogEntry(LogType._Warning, "File : " & TreeViewMedia.SelectedNode.Tag & vbCrLf & "Not found !")
        End If
    End Sub

    'Delete file
    Private Sub DeleteSelectedFileToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteSelectedFileToolStripMenuItem.Click
        If File.Exists(TreeViewMedia.SelectedNode.Tag) Then
            Try
                If MsgBox("Are you sure ?", MsgBoxStyle.YesNo, "Delete") = MsgBoxResult.Yes Then
                    Dim vSystemName As String = ToolStripLabel1.Text
                    Dim vRomName As String = DataGridView1.SelectedRows(0).Cells(0).Value
                    Dim vRow As DataGridViewRow = DataGridView1.SelectedRows(0)
                    UnloadMediaFromPictureBoxes(TreeViewMedia.SelectedNode.Tag)
                    File.Delete(TreeViewMedia.SelectedNode.Tag)
                    gHS.Systems(vSystemName).Roms(vRomName).CheckRomMedia(gHS.Systems(vSystemName).mSearchSubfolders, gHS.Systems(vSystemName).mRomPath, gHS.Systems(vSystemName).mRomExtensions)
                    FillRomsMediaCheckROW(DataGridView1, DataGridView1.SelectedRows(0), gHS.Systems(vSystemName).Roms(vRomName))
                    DataGridView1.ClearSelection()
                    vRow.Selected = True
                End If
            Catch ex As Exception
                LogEntry(LogType._Error, "Error: " & ex.Message)
            End Try
        Else
            MsgBox("File : " & TreeViewMedia.SelectedNode.Tag & vbCrLf & "Not found !", MsgBoxStyle.Critical, "Error")
            LogEntry(LogType._Warning, "File : " & TreeViewMedia.SelectedNode.Tag & vbCrLf & "Not found !")
        End If
    End Sub

    'Copy all Media to ...
    Private Sub CopyAllMediaToToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyAllMediaToToolStripMenuItem.Click
        Dim frmMove As New FrmCopyMedia(ToolStripLabel1.Text, DataGridView1.SelectedRows(0).Cells(0).Value)
        frmMove.Show()
    End Sub

    Private Enum AddingTypes
        FromFile
        FromURL
        FromYouTube
        FromClipboard
    End Enum

    Private Enum MediaType
        HSWheel
        HSVideo
        HSTheme
        HSArt1
        HSArt2
        HSArt3
        HSArt4
        HPArtwork
        HPBackground
        HPBezel
        HPController
        HPFade
        HPFont
        HPGuide
        HPManual
        HPMenuImage
        HPMultigame
        HPMusic
        HPSound
        HPVideo
        HPWheel
    End Enum

    Private Enum SystemMediaType
        HSWheel
        HSVideo
        HSTheme
        HSPointer
        HSBetaBriteClassic
        HSBetaBritePrism
        'HSArt4
        'HPArtwork
        HPBackground
        HPBezel
        'HPController
        HPFade
        'HPFont
        'HPGuide
        'HPManual
        'HPMenuImage
        'HPMultigame
        'HPMusic
        'HPSound
        'HPVideo
        'HPWheel
    End Enum

    Private Sub AddNewMedia(ByVal vMediaType As MediaType, ByVal vSystemName As String, ByVal vRomName As String, ByVal vAddingType As AddingTypes)
        LogEntry(LogType._Info, "AddMedia : type->" & vMediaType.ToString & ", from->" & vAddingType.ToString & " in " & vSystemName & "/" & vRomName)
        Dim openFileDialog1 As New OpenFileDialog()
        Dim DlgFilter As String = ""
        Dim destFolder As String = ""
        Dim FixedName As Boolean = True
        openFileDialog1.RestoreDirectory = True
        openFileDialog1.InitialDirectory = gHSPath
        Dim ForcePng As Boolean = False
        Dim vEndedOK As Boolean = True
        Dim vRow As DataGridViewRow = DataGridView1.SelectedRows(0)

        Select Case vMediaType
            Case MediaType.HSWheel
                DlgFilter = "Wheels (*.png;*.bmp;*.gif;*.jpg)|*.png;*.bmp;*.gif;*.jpg"
                destFolder = Path.GetDirectoryName(gHS.Systems(vSystemName).Roms(vRomName).WheelPath)
                If destFolder = Nothing Then
                    destFolder = gHSPath & "\Media\" & vSystemName & "\Images\Wheel\"
                End If
                ForcePng = True
            Case MediaType.HSVideo
                DlgFilter = "Videos (*.png;*.bmp;*.gif;*.jpg;*.flv;*.mp4)|*.png;*.bmp;*.gif;*.jpg;*.flv;*.mp4"
                destFolder = Path.GetDirectoryName(gHS.Systems(vSystemName).Roms(vRomName).VideoPath)
                If destFolder = Nothing Then
                    destFolder = gHSPath & "\Media\" & vSystemName & "\Video\"
                End If
            Case MediaType.HSTheme
                DlgFilter = "Themes (*.zip)|*.zip"
                destFolder = Path.GetDirectoryName(gHS.Systems(vSystemName).Roms(vRomName).ThemePath)
                If destFolder = Nothing Then
                    destFolder = gHSPath & "\Media\" & vSystemName & "\Themes\"
                End If
            Case MediaType.HSArt1
                DlgFilter = "Art1 (*.png;*.bmp;*.gif;*.jpg)|*.png;*.bmp;*.gif;*.jpg"
                destFolder = Path.GetDirectoryName(gHS.Systems(vSystemName).Roms(vRomName).Art1Path)
                If destFolder = Nothing Then
                    destFolder = gHSPath & "\Media\" & vSystemName & "\Images\Artwork1\"
                End If
                ForcePng = True
            Case MediaType.HSArt2
                DlgFilter = "Art2 (*.png;*.bmp;*.gif;*.jpg)|*.png;*.bmp;*.gif;*.jpg"
                destFolder = Path.GetDirectoryName(gHS.Systems(vSystemName).Roms(vRomName).Art2Path)
                If destFolder = Nothing Then
                    destFolder = gHSPath & "\Media\" & vSystemName & "\Images\Artwork2\"
                End If
                ForcePng = True
            Case MediaType.HSArt3
                DlgFilter = "Art3 (*.png;*.bmp;*.gif;*.jpg)|*.png;*.bmp;*.gif;*.jpg"
                destFolder = Path.GetDirectoryName(gHS.Systems(vSystemName).Roms(vRomName).Art3Path)
                If destFolder = Nothing Then
                    destFolder = gHSPath & "\Media\" & vSystemName & "\Images\Artwork3\"
                End If
                ForcePng = True
            Case MediaType.HSArt4
                DlgFilter = "Art4 (*.png;*.bmp;*.gif;*.jpg)|*.png;*.bmp;*.gif;*.jpg"
                destFolder = Path.GetDirectoryName(gHS.Systems(vSystemName).Roms(vRomName).Art4Path)
                If destFolder = Nothing Then
                    destFolder = gHSPath & "\Media\" & vSystemName & "\Images\Artwork4\"
                End If
                ForcePng = True
            Case MediaType.HPArtwork
                DlgFilter = "Artworks (*.png;*.bmp;*.gif;*.jpg;*.tip;*.txt;*.pdf;*.zip;*.rar;*.7z;*.lzh;*.gzip;*.tar)|*.png;*.bmp;*.gif;*.jpg;*.tip;*.txt;*.pdf;*.zip;*.rar;*.7z;*.lzh;*.gzip;*.tar|All files (*.*)|*.*"
                destFolder = Path.GetDirectoryName(gHS.Systems(vSystemName).Roms(vRomName).ArtworkPath)
                FixedName = False
            Case MediaType.HPBackground
                DlgFilter = "Backgrounds  (*.png;*.bmp;*.gif;*.jpg;*.tip)|*.png;*.bmp;*.gif;*.jpg;*.tip|All files (*.*)|*.*"
                destFolder = Path.GetDirectoryName(gHS.Systems(vSystemName).Roms(vRomName).BackgroundsPath)
                FixedName = False
            Case MediaType.HPBezel
                DlgFilter = "Wheels (*.png;*.bmp;*.gif;*.jpg)|*.png;*.bmp;*.gif;*.jpg"
                destFolder = Path.GetDirectoryName(gHS.Systems(vSystemName).Roms(vRomName).BezelsPath)
                FixedName = False
            Case MediaType.HPController
                DlgFilter = "Controllers (*.png;*.bmp;*.gif;*.jpg;*.tip;*.txt;*.pdf;*.zip;*.rar;*.7z;*.lzh;*.gzip;*.tar)|*.png;*.bmp;*.gif;*.jpg;*.tip;*.txt;*.pdf;*.zip;*.rar;*.7z;*.lzh;*.gzip;*.tar|All files (*.*)|*.*"
                destFolder = Path.GetDirectoryName(gHS.Systems(vSystemName).Roms(vRomName).ControllerPath)
                FixedName = False
            Case MediaType.HPFade
                DlgFilter = "Fades (*.png;*.bmp;*.gif;*.jpg)|*.png;*.bmp;*.gif;*.jpg"
                destFolder = Path.GetDirectoryName(gHS.Systems(vSystemName).Roms(vRomName).FadePath)
                FixedName = False
                ForcePng = True
            Case MediaType.HPFont
                DlgFilter = "Fonts (*.ttf;*.fnt;*.font)|*.ttf;*.fnt;*.font"
                destFolder = Path.GetDirectoryName(gHS.Systems(vSystemName).Roms(vRomName).FontsPath)
                FixedName = False
            Case MediaType.HPGuide
                DlgFilter = "Guides (*.png;*.bmp;*.gif;*.jpg;*.tip;*.txt;*.pdf;*.zip;*.rar;*.7z;*.lzh;*.gzip;*.tar)|*.png;*.bmp;*.gif;*.jpg;*.tip;*.txt;*.pdf;*.zip;*.rar;*.7z;*.lzh;*.gzip;*.tar|All files (*.*)|*.*"
                destFolder = Path.GetDirectoryName(gHS.Systems(vSystemName).Roms(vRomName).GuidesPath)
                FixedName = False
            Case MediaType.HPManual
                DlgFilter = "Manuals (*.png;*.bmp;*.gif;*.jpg;*.tip;*.txt;*.pdf;*.zip;*.rar;*.7z;*.lzh;*.gzip;*.tar)|*.png;*.bmp;*.gif;*.jpg;*.tip;*.txt;*.pdf;*.zip;*.rar;*.7z;*.lzh;*.gzip;*.tar|All files (*.*)|*.*"
                destFolder = Path.GetDirectoryName(gHS.Systems(vSystemName).Roms(vRomName).ManualsPath)
                FixedName = False
            Case MediaType.HPMenuImage
                DlgFilter = "Menus (*.png;*.bmp;*.gif;*.jpg)|*.png;*.bmp;*.gif;*.jpg"
                destFolder = Path.GetDirectoryName(gHS.Systems(vSystemName).Roms(vRomName).MenuImagesPath)
                FixedName = False
            Case MediaType.HPMultigame
                DlgFilter = "Multigame (*.png;*.bmp;*.gif;*.jpg)|*.png;*.bmp;*.gif;*.jpg"
                destFolder = Path.GetDirectoryName(gHS.Systems(vSystemName).Roms(vRomName).MultiGamePath)
                FixedName = False
            Case MediaType.HPMusic
                DlgFilter = "Music (*.mp3;*.m4a;*.wav;*.mid;*.wma)|*.mp3;*.m4a;*.wav;*.mid;*.wma"
                destFolder = Path.GetDirectoryName(gHS.Systems(vSystemName).Roms(vRomName).MusicPath)
                FixedName = False
            Case MediaType.HPSound
                DlgFilter = "Sounds (*.mp3;*.m4a;*.wav;*.mid;*.wma)|*.mp3;*.m4a;*.wav;*.mid;*.wma"
                destFolder = Path.GetDirectoryName(gHS.Systems(vSystemName).Roms(vRomName).SoundsPath)
                FixedName = False
            Case MediaType.HPVideo
                DlgFilter = "Videos (*.flv;*.mp4)|*.flv;*.mp4"
                destFolder = Path.GetDirectoryName(gHS.Systems(vSystemName).Roms(vRomName).HLVideosPath)
                FixedName = False
            Case MediaType.HPWheel
                DlgFilter = "Wheels (*.png;*.bmp;*.gif;*.jpg)|*.png;*.bmp;*.gif;*.jpg"
                destFolder = Path.GetDirectoryName(gHS.Systems(vSystemName).Roms(vRomName).HLWheelsPath)
                FixedName = False
                ForcePng = True
        End Select

        'just in case of
        If destFolder.EndsWith("\") = False Then
            destFolder = destFolder & "\"
        End If
        If destFolder.EndsWith("\\") Then
            destFolder = destFolder.Replace("\\", "\")
        End If

        LogEntry(LogType._Info, "         Filter applied = " & DlgFilter)
        LogEntry(LogType._Info, "         Destination folder = " & destFolder)
        LogEntry(LogType._Info, "         Fixed FileName = " & FixedName.ToString)
        LogEntry(LogType._Info, "         Force png format = " & ForcePng.ToString)
        Try
            If Directory.Exists(destFolder) = False Then
                Directory.CreateDirectory(destFolder)
                Application.DoEvents()
            End If
        Catch ex As Exception
            LogEntry(LogType._Error, "Cannot create destination folder (" & destFolder & "): " & ex.Message)
            LogEntry(LogType._Error, "Abording copy of file ...")
            vEndedOK = False
            Exit Sub
        End Try


        Select Case vAddingType
            Case AddingTypes.FromFile
                openFileDialog1.Filter = DlgFilter
                openFileDialog1.FilterIndex = 1
                If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                    Dim vExt As String = Path.GetExtension(openFileDialog1.FileName).ToLower
                    Dim DestFilename As String = ""

                    If FixedName = True Then
                        'If this is an Hyperspin file
                        If vExt = ".bmp" Or vExt = ".gif" Or vExt = ".jpg" Then
                            If ForcePng = True Then
                                Dim xx As Image = Image.FromFile(openFileDialog1.FileName)
                                DestFilename = destFolder & vRomName & ".png"
                                LogEntry(LogType._Info, "Complete File destination = " & DestFilename)
                                Try
                                    UnloadMediaFromPictureBoxes(DestFilename)
                                    If File.Exists(DestFilename) Then
                                        File.Delete(DestFilename)
                                        Application.DoEvents()
                                    End If
                                    xx.Save(DestFilename, System.Drawing.Imaging.ImageFormat.Png)
                                    LogEntry(LogType._Info, "File copied successfully")
                                Catch ex As Exception
                                    vEndedOK = False
                                    LogEntry(LogType._Error, "Cannot copy file. Original error: " & ex.Message)
                                    ToolStripStatusLabel1.Text = "Cannot copy file. Original error: " & ex.Message
                                    Exit Sub
                                End Try
                            Else
                                DestFilename = destFolder & vRomName & vExt
                                LogEntry(LogType._Info, "Complete File destination = " & DestFilename)
                                Try
                                    UnloadMediaFromPictureBoxes(DestFilename)
                                    File.Copy(openFileDialog1.FileName, DestFilename, True)
                                    LogEntry(LogType._Info, "File copied successfully")
                                Catch ex As Exception
                                    vEndedOK = False
                                    LogEntry(LogType._Error, "Cannot copy file. Original error: " & ex.Message)
                                    ToolStripStatusLabel1.Text = "Cannot copy file. Original error: " & ex.Message
                                    Exit Sub
                                End Try
                            End If
                        Else
                            DestFilename = destFolder & vRomName & vExt
                            LogEntry(LogType._Info, "Complete File destination = " & DestFilename)
                            Try
                                UnloadMediaFromPictureBoxes(DestFilename)
                                File.Copy(openFileDialog1.FileName, DestFilename, True)
                                LogEntry(LogType._Info, "File copied successfully")
                            Catch ex As Exception
                                vEndedOK = False
                                LogEntry(LogType._Error, "Cannot copy file. Original error: " & ex.Message)
                                ToolStripStatusLabel1.Text = "Cannot copy file. Original error: " & ex.Message
                                Exit Sub
                            End Try
                        End If
                    Else
                        'If this is an HyperPause file
                        If ForcePng = True Then
                            Dim xx As Image = Image.FromFile(openFileDialog1.FileName)
                            Dim tempFile As String = Path.GetFileNameWithoutExtension(openFileDialog1.FileName)
                            DestFilename = destFolder & tempFile & ".png"
                            LogEntry(LogType._Info, "Complete File destination = " & DestFilename)
                            Try
                                UnloadMediaFromPictureBoxes(DestFilename)
                                xx.Save(DestFilename, System.Drawing.Imaging.ImageFormat.Png)
                                LogEntry(LogType._Info, "File copied successfully")
                            Catch ex As Exception
                                vEndedOK = False
                                LogEntry(LogType._Error, "Cannot copy file. Original error: " & ex.Message)
                                ToolStripStatusLabel1.Text = "Cannot copy file. Original error: " & ex.Message
                                Exit Sub
                            End Try
                        Else
                            DestFilename = destFolder & Path.GetFileName(openFileDialog1.FileName).ToLower
                            LogEntry(LogType._Info, "Complete File destination = " & DestFilename)
                            Try
                                UnloadMediaFromPictureBoxes(DestFilename)
                                File.Copy(openFileDialog1.FileName, DestFilename, True)
                                LogEntry(LogType._Info, "File copied successfully")
                            Catch ex As Exception
                                vEndedOK = False
                                LogEntry(LogType._Error, "Cannot copy file. Original error: " & ex.Message)
                                ToolStripStatusLabel1.Text = "Cannot copy file. Original error: " & ex.Message
                                Exit Sub
                            End Try
                        End If
                    End If
                End If
            Case AddingTypes.FromClipboard
                If Clipboard.ContainsText = True And vMediaType = MediaType.HPGuide Then
                    File.WriteAllText(destFolder & "Guide_" & Date.Now.ToString("yyyy-MM-dd_HH-mm-ss") & ".txt", Clipboard.GetText, System.Text.Encoding.UTF8)
                ElseIf Clipboard.ContainsImage = True Then
                    Dim DestFilename As String
                    If FixedName = True Then
                        DestFilename = destFolder & vRomName & ".png"
                    Else
                        Dim vTmpName As String = InputBox("Enter new name for image (without extension) : ", "Filename", "")
                        If vTmpName = "" Then
                            MsgBox("You must specify a valid name ... aborting", MsgBoxStyle.Critical, "Error")
                            Exit Sub
                        Else
                            DestFilename = destFolder & vTmpName & ".png"
                        End If
                    End If
                    Try
                        LogEntry(LogType._Info, "Complete File destination = " & DestFilename)
                        UnloadMediaFromPictureBoxes(DestFilename)
                        Clipboard.GetImage.Save(DestFilename, System.Drawing.Imaging.ImageFormat.Png)
                        LogEntry(LogType._Info, "File copied successfully")
                    Catch ex As Exception
                        vEndedOK = False
                        LogEntry(LogType._Error, "Cannot copy file. Original error: " & ex.Message)
                        ToolStripStatusLabel1.Text = "Cannot copy file. Original error: " & ex.Message
                    End Try
                Else
                    vEndedOK = False
                    LogEntry(LogType._Error, "Looks like clipboard does NOT contain an image, please retry.")
                    ToolStripStatusLabel1.Text = "Looks like clipboard does NOT contain an image, please retry."
                End If
            Case AddingTypes.FromURL
                Dim vURL = InputBox("Enter URL:", "Source file", "")
                If vURL = "" Then
                    MsgBox("You must specify a valid URL ... aborting", MsgBoxStyle.Critical, "Error")
                    Exit Sub
                End If
                Me.Cursor = Cursors.WaitCursor
                Application.DoEvents()
                Dim DestFilename As String
                Dim vFileName = vURL.Substring(vURL.LastIndexOf("/") + 1)
                Try
                    ' Make a WebClient.
                    Dim web_client As WebClient = New WebClient
                    ' Download the file.
                    If FixedName = True Then
                        If ForcePng = True Then
                            Dim vTmp As String = Path.GetTempPath & vFileName
                            DestFilename = destFolder & vRomName & ".png"
                            LogEntry(LogType._Info, "Complete File destination = " & DestFilename)
                            Try
                                If Path.GetExtension(vFileName).ToLower = ".png" Then
                                    UnloadMediaFromPictureBoxes(DestFilename)
                                    web_client.DownloadFile(vURL, DestFilename)
                                    LogEntry(LogType._Info, "File copied successfully")
                                Else
                                    UnloadMediaFromPictureBoxes(DestFilename)
                                    web_client.DownloadFile(vURL, vTmp)
                                    Dim vImg As Image = Image.FromFile(vTmp)
                                    vImg.Save(DestFilename, System.Drawing.Imaging.ImageFormat.Png)
                                    LogEntry(LogType._Info, "File copied successfully")
                                End If
                            Catch ex As Exception
                                web_client.CancelAsync()
                                web_client = Nothing
                                File.Delete(vTmp)
                                vEndedOK = False
                                LogEntry(LogType._Error, "Cannot download/copy file. Original error: " & ex.Message)
                                ToolStripStatusLabel1.Text = "Cannot download/copy file. Original error: " & ex.Message
                                Me.Cursor = Cursors.Default
                                Exit Sub
                            End Try
                        Else
                            Try
                                DestFilename = destFolder & vRomName & Path.GetExtension(vFileName)
                                LogEntry(LogType._Info, "Complete File destination = " & DestFilename)
                                UnloadMediaFromPictureBoxes(DestFilename)
                                web_client.DownloadFile(vURL, DestFilename)
                                LogEntry(LogType._Info, "File copied successfully")
                            Catch ex As Exception
                                web_client.CancelAsync()
                                web_client = Nothing
                                vEndedOK = False
                                LogEntry(LogType._Error, "Cannot download/copy file. Original error: " & ex.Message)
                                ToolStripStatusLabel1.Text = "Cannot download/copy file. Original error: " & ex.Message
                                Me.Cursor = Cursors.Default
                                Exit Sub
                            End Try
                        End If
                    Else
                        DestFilename = destFolder & vFileName
                        LogEntry(LogType._Info, "Complete File destination = " & DestFilename)
                        UnloadMediaFromPictureBoxes(DestFilename)
                        Try
                            web_client.DownloadFile(vURL, DestFilename)
                            LogEntry(LogType._Info, "File copied successfully")
                        Catch ex As Exception
                            vEndedOK = False
                            LogEntry(LogType._Error, "Cannot download/copy file. Original error: " & ex.Message)
                            ToolStripStatusLabel1.Text = "Cannot download/copy file. Original error: " & ex.Message
                            Me.Cursor = Cursors.Default
                            Exit Sub
                        End Try
                    End If
                Catch ex As Exception
                    vEndedOK = False
                    LogEntry(LogType._Error, "Cannot download/copy file. Original error: " & ex.Message)
                    ToolStripStatusLabel1.Text = "Cannot download/copy file. Original error: " & ex.Message
                    Me.Cursor = Cursors.Default
                End Try
                Me.Cursor = Cursors.Default
            Case AddingTypes.FromYouTube
                Dim vURL As String = InputBox("Enter full youtube URL here:", "URL", "https://www.youtube.com/watch?v=JfAc6Y9RyUA")
                If vURL = "" Then
                    vEndedOK = False
                    Exit Select
                End If
                vEndedOK = True
                Me.Visible = True
                Me.Focus()

        End Select

        If vEndedOK = True Then
            gHS.Systems(vSystemName).Roms(vRomName).CheckRomMedia(gHS.Systems(vSystemName).mSearchSubfolders, gHS.Systems(vSystemName).mRomPath, gHS.Systems(vSystemName).mRomExtensions)
            FillRomsMediaCheckROW(DataGridView1, vRow, gHS.Systems(vSystemName).Roms(vRomName))
            DataGridView1.ClearSelection()
            vRow.Selected = True
        End If
    End Sub

#End Region

    Private Sub UnloadMediaFromPictureBoxes(Optional ByVal vPath As String = "")
        If vPath = "" Then
            PictureBox11.Image = Nothing
            PictureBox1.Image = Nothing
            PictureBox2.Image = Nothing
            PictureBox3.Image = Nothing
            PictureBox4.Image = Nothing
            PictureBox5.Image = Nothing
            WBMedia.Navigate("")
        Else
            If PictureBox11.ImageLocation = vPath Then
                PictureBox11.Image = Nothing
                PictureBox11.Refresh()
            End If
            If WBMedia.Url.AbsolutePath = vPath Then
                WBMedia.Navigate("")
            End If
            If PictureBox2.ImageLocation = vPath Then
                PictureBox2.Image = Nothing
                PictureBox2.Refresh()
            End If
            If PictureBox1.ImageLocation = vPath Then
                PictureBox1.Image = Nothing
                PictureBox1.Refresh()
            End If
            If PictureBox3.ImageLocation = vPath Then
                PictureBox3.Image = Nothing
                PictureBox3.Refresh()
            End If
            If PictureBox4.ImageLocation = vPath Then
                PictureBox4.Image = Nothing
                PictureBox4.Refresh()
            End If
            If PictureBox5.ImageLocation = vPath Then
                PictureBox5.Image = Nothing
                PictureBox5.Refresh()
            End If
        End If
        Application.DoEvents()
    End Sub

    Private Sub UnloadSystemMediaFromPictureBoxes(Optional ByVal vPath As String = "")
        If vPath = "" Then
            PictureBox16.Image = Nothing
            WBSystems.Navigate("")
            PictureBox12.Image = Nothing
            PictureBox13.Image = Nothing
            PictureBox14.Image = Nothing
            PictureBox15.Image = Nothing
        Else
            If PictureBox16.ImageLocation = vPath Then
                PictureBox16.Image = Nothing
                PictureBox16.Refresh()
            End If
            If WBSystems.Url.AbsolutePath = vPath Then
                WBSystems.Navigate("")
            End If
            If PictureBox12.ImageLocation = vPath Then
                PictureBox12.Image = Nothing
                PictureBox12.Refresh()
            End If
            If PictureBox13.ImageLocation = vPath Then
                PictureBox13.Image = Nothing
                PictureBox13.Refresh()
            End If
            If PictureBox14.ImageLocation = vPath Then
                PictureBox14.Image = Nothing
                PictureBox14.Refresh()
            End If
            If PictureBox15.ImageLocation = vPath Then
                PictureBox15.Image = Nothing
                PictureBox15.Refresh()
            End If
        End If
        Application.DoEvents()
    End Sub

#Region "ContextMenu Medias Roms"
    'Add new HSwheel from file
    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click
        AddNewMedia(MediaType.HSWheel, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromFile)
    End Sub
    'Add new HSwheel from clipb
    Private Sub FromClipboardToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FromClipboardToolStripMenuItem.Click
        AddNewMedia(MediaType.HSWheel, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromClipboard)
    End Sub
    'Add new HSwheel from URL
    Private Sub FromURLToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FromURLToolStripMenuItem.Click
        AddNewMedia(MediaType.HSWheel, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromURL)
    End Sub
    'Add new HSvideo from file
    Private Sub FromFileToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FromFileToolStripMenuItem.Click
        AddNewMedia(MediaType.HSVideo, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromFile)
    End Sub
    'Add new HSvideo from Youtube
    Private Sub FromYouTubeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FromYouTubeToolStripMenuItem.Click
        AddNewMedia(MediaType.HSVideo, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromYouTube)
    End Sub
    'Add new HPArtwork from URL
    Private Sub ToolStripMenuItem16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem16.Click
        AddNewMedia(MediaType.HPArtwork, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromURL)
    End Sub
    'Add new HLTheme from file
    Private Sub ThemeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ThemeToolStripMenuItem.Click
        AddNewMedia(MediaType.HSTheme, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromFile)
    End Sub
    'HPGuide from File
    Private Sub ToolStripMenuItem30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem30.Click
        AddNewMedia(MediaType.HPGuide, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromFile)
    End Sub
    'HPManual from File
    Private Sub ToolStripMenuItem34_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem34.Click
        AddNewMedia(MediaType.HPManual, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromFile)
    End Sub
    'HPController from URL
    Private Sub ToolStripMenuItem25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem25.Click
        AddNewMedia(MediaType.HPController, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromURL)
    End Sub
    'HPManual from URL
    Private Sub ToolStripMenuItem36_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem36.Click
        AddNewMedia(MediaType.HPManual, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromURL)
    End Sub
    'art1 ffile
    Private Sub ToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem2.Click
        AddNewMedia(MediaType.HSArt1, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromFile)
    End Sub

    Private Sub ToolStripMenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem5.Click
        AddNewMedia(MediaType.HSArt2, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromFile)
    End Sub

    Private Sub ToolStripMenuItem8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem8.Click
        AddNewMedia(MediaType.HSArt3, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromFile)
    End Sub

    Private Sub ToolStripMenuItem11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem11.Click
        AddNewMedia(MediaType.HSArt4, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromFile)
    End Sub

    Private Sub ToolStripMenuItem14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem14.Click
        AddNewMedia(MediaType.HPArtwork, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromFile)
    End Sub

    Private Sub ToolStripMenuItem17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem17.Click
        AddNewMedia(MediaType.HPBackground, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromFile)
    End Sub

    Private Sub ToolStripMenuItem20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem20.Click
        AddNewMedia(MediaType.HPBezel, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromFile)
    End Sub

    Private Sub ToolStripMenuItem23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem23.Click
        AddNewMedia(MediaType.HPController, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromFile)
    End Sub

    Private Sub ToolStripMenuItem26_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem26.Click
        AddNewMedia(MediaType.HPFade, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromFile)
    End Sub

    Private Sub ToolStripMenuItem29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem29.Click
        AddNewMedia(MediaType.HPFont, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromFile)
    End Sub

    Private Sub ToolStripMenuItem37_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem37.Click
        AddNewMedia(MediaType.HPMenuImage, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromFile)
    End Sub

    Private Sub ToolStripMenuItem40_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem40.Click
        AddNewMedia(MediaType.HPMultigame, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromFile)
    End Sub

    Private Sub ToolStripMenuItem43_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem43.Click
        AddNewMedia(MediaType.HPMusic, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromFile)
    End Sub

    Private Sub ToolStripMenuItem46_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem46.Click
        AddNewMedia(MediaType.HPSound, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromFile)
    End Sub

    Private Sub ToolStripMenuItem44_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem44.Click
        AddNewMedia(MediaType.HPVideo, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromFile)
    End Sub

    Private Sub ToolStripMenuItem47_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem47.Click
        AddNewMedia(MediaType.HPWheel, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromFile)
    End Sub

    Private Sub ToolStripMenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem3.Click
        AddNewMedia(MediaType.HSArt1, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromClipboard)
    End Sub

    Private Sub ToolStripMenuItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem6.Click
        AddNewMedia(MediaType.HSArt2, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromClipboard)
    End Sub

    Private Sub ToolStripMenuItem9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem9.Click
        AddNewMedia(MediaType.HSArt3, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromClipboard)
    End Sub

    Private Sub ToolStripMenuItem12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem12.Click
        AddNewMedia(MediaType.HSArt4, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromClipboard)
    End Sub

    Private Sub ToolStripMenuItem15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem15.Click
        AddNewMedia(MediaType.HPArtwork, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromClipboard)
    End Sub

    Private Sub ToolStripMenuItem18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem18.Click
        AddNewMedia(MediaType.HPBackground, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromClipboard)
    End Sub

    Private Sub ToolStripMenuItem21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem21.Click
        AddNewMedia(MediaType.HPBezel, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromClipboard)
    End Sub

    Private Sub ToolStripMenuItem24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem24.Click
        AddNewMedia(MediaType.HPController, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromClipboard)
    End Sub

    Private Sub ToolStripMenuItem27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem27.Click
        AddNewMedia(MediaType.HPFade, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromClipboard)
    End Sub

    Private Sub ToolStripMenuItem32_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem32.Click
        AddNewMedia(MediaType.HPGuide, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromClipboard)
    End Sub

    Private Sub ToolStripMenuItem35_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem35.Click
        AddNewMedia(MediaType.HPManual, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromClipboard)
    End Sub

    Private Sub ToolStripMenuItem38_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem38.Click
        AddNewMedia(MediaType.HPMenuImage, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromClipboard)
    End Sub

    Private Sub ToolStripMenuItem41_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem41.Click
        AddNewMedia(MediaType.HPMultigame, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromClipboard)
    End Sub

    Private Sub ToolStripMenuItem50_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem50.Click
        AddNewMedia(MediaType.HPWheel, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromClipboard)
    End Sub

    Private Sub ToolStripMenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem4.Click
        AddNewMedia(MediaType.HSArt1, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromURL)
    End Sub

    Private Sub ToolStripMenuItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem7.Click
        AddNewMedia(MediaType.HSArt2, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromURL)
    End Sub

    Private Sub ToolStripMenuItem10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem10.Click
        AddNewMedia(MediaType.HSArt3, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromURL)
    End Sub

    Private Sub ToolStripMenuItem13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem13.Click
        AddNewMedia(MediaType.HSArt4, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromURL)
    End Sub

    Private Sub ToolStripMenuItem19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem19.Click
        AddNewMedia(MediaType.HPBackground, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromURL)
    End Sub

    Private Sub ToolStripMenuItem22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem22.Click
        AddNewMedia(MediaType.HPBezel, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromURL)
    End Sub

    Private Sub ToolStripMenuItem28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem28.Click
        AddNewMedia(MediaType.HPFade, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromURL)
    End Sub

    Private Sub ToolStripMenuItem31_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem31.Click
        AddNewMedia(MediaType.HPFont, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromURL)
    End Sub

    Private Sub ToolStripMenuItem33_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem33.Click
        AddNewMedia(MediaType.HPGuide, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromURL)
    End Sub

    Private Sub ToolStripMenuItem39_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem39.Click
        AddNewMedia(MediaType.HPMenuImage, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromURL)
    End Sub

    Private Sub ToolStripMenuItem42_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem42.Click
        AddNewMedia(MediaType.HPMultigame, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromURL)
    End Sub

    Private Sub ToolStripMenuItem45_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem45.Click
        AddNewMedia(MediaType.HPMusic, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromURL)
    End Sub

    Private Sub ToolStripMenuItem48_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem48.Click
        AddNewMedia(MediaType.HPSound, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromURL)
    End Sub

    Private Sub ToolStripMenuItem51_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem51.Click
        AddNewMedia(MediaType.HPWheel, ToolStripLabel1.Text, DataGridView1.SelectedCells(0).Value.ToString, AddingTypes.FromURL)
    End Sub

    Private Sub OpenContainingFolderToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenContainingFolderToolStripMenuItem.Click
        If Directory.Exists(Path.GetDirectoryName(TreeViewMedia.SelectedNode.Tag)) Then
            Try
                Dim psi As New ProcessStartInfo
                Dim p As New Process
                psi.FileName = Path.GetDirectoryName(TreeViewMedia.SelectedNode.Tag)
                p.StartInfo = psi
                p.Start()
            Catch ex As Exception
                LogEntry(LogType._Error, "Error: " & ex.Message)
            End Try
        Else
            MsgBox("Folder : " & Path.GetDirectoryName(TreeViewMedia.SelectedNode.Tag) & vbCrLf & "Not found !", MsgBoxStyle.Critical, "Error")
            LogEntry(LogType._Warning, "Folder : " & Path.GetDirectoryName(TreeViewMedia.SelectedNode.Tag) & vbCrLf & "Not found !")
        End If
    End Sub

#End Region

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Dim i As Integer = 0
        If gCheckRoms = True Then
            DataGridView4.CommitEdit(DataGridViewDataErrorContexts.Commit)
            Dim vSystem As String = ToolStripLabel1.Text
            Try
                While RowsToDel() = True
                    For Each vRow As DataGridViewRow In DataGridView4.Rows
                        If vRow.Cells(0).Style.BackColor = Color.Red Then
                            DataGridView4.Rows.Remove(vRow)
                            i = i + 1
                        End If
                    Next
                End While
            Catch ex As Exception
                ToolStripStatusLabel1.Text = ex.Message.ToString
            Finally
                DataGridView4.CommitEdit(DataGridViewDataErrorContexts.RowDeletion)
                Application.DoEvents()
            End Try
            MsgBox("Done ! " & i & " entries removed, do not forget to save.", MsgBoxStyle.Information, "Job done")
            Label1.Text = "DO NOT FORGET TO SAVE !!"
        Else
            Label1.Text = "Info : YOU MUST ENABLE ROM CHECK IN PARAMETERS, and reload this XML"
        End If
    End Sub

    Private Function RowsToDel() As Boolean
        Dim out As Boolean = False
        For Each vRow As DataGridViewRow In DataGridView4.Rows
            If vRow.Cells(0).Style.BackColor = Color.Red Then
                out = True
            End If
        Next
        Return out
    End Function

    Public Shared Sub DownloadYT(ByVal videoID As String, ByVal newFilePath As String)
        Dim wc As WebClient = New WebClient
        Dim file As String = wc.DownloadString(String.Format("http://www.youtube.com/watch?v={0}", videoID))
        Console.WriteLine(file)
        Dim t As String = New Regex("(?<=&t=)[^&]*").Match(file).Value
        wc.DownloadFile(String.Format("http://www.youtube.com/get_video?t={0}=&video_id={1}", t, videoID), newFilePath)
    End Sub

    Private Sub PictureBox15_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox15.DoubleClick
        Try
            If DataGridView2.SelectedRows.Count > 0 Then
                If DataGridView2.SelectedRows.Item(0).Cells(2).Tag <> "" Then
                    Process.Start("explorer.exe", Path.GetDirectoryName(DataGridView2.SelectedRows.Item(0).Cells(2).Tag))
                Else
                    Process.Start("explorer.exe", gHS.HSMediaPath & "\Main Menu\Images\Wheel\")
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub PictureBox14_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox14.DoubleClick
        Try
            If DataGridView2.SelectedRows.Count > 0 Then
                If DataGridView2.SelectedRows.Item(0).Cells(5).Tag <> "" Then
                    Process.Start("explorer.exe", Path.GetDirectoryName(DataGridView2.SelectedRows.Item(0).Cells(5).Tag))
                Else
                    Process.Start("explorer.exe", gHS.HSMediaPath & "\Main Menu\Images\BetaBrite Classic\")
                End If
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub PictureBox13_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox13.DoubleClick
        Try
            If DataGridView2.SelectedRows.Count > 0 Then
                If DataGridView2.SelectedRows.Item(0).Cells(6).Tag <> "" Then
                    Process.Start("explorer.exe", Path.GetDirectoryName(DataGridView2.SelectedRows.Item(0).Cells(6).Tag))
                Else
                    Process.Start("explorer.exe", gHS.HSMediaPath & "\Main Menu\Images\BetaBrite Prism\")
                End If
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub PictureBox12_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox12.DoubleClick
        Try
            If DataGridView2.SelectedRows.Count > 0 Then
                Process.Start("explorer.exe", gHS.HSMediaPath & "\" & DataGridView2.SelectedRows.Item(0).Cells(0).Value & "\Images\Other\")
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub TreeviewSystems_AfterSelect(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeviewSystems.AfterSelect
        Dim CurrNode As TreeNode
        CurrNode = TreeviewSystems.SelectedNode

        If CurrNode IsNot Nothing And CurrNode.Tag <> "" Then
            Select Case Path.GetExtension(CurrNode.Tag.ToString).ToLower
                Case ".png", ".bmp", ".gif", ".jpg", ".jpeg"
                    'PictureBox16.ImageLocation = CurrNode.Tag
                    Dim fs As System.IO.FileStream
                    fs = New System.IO.FileStream(CurrNode.Tag, IO.FileMode.Open, IO.FileAccess.Read)
                    PictureBox16.Image = System.Drawing.Image.FromStream(fs)
                    fs.Close()
                    PictureBox16.Visible = True
                    WBSystems.Visible = False
                    WBSystems.Navigate("")
                Case Else
                    WBSystems.Navigate(CurrNode.Tag)
                    WBSystems.Visible = True
                    PictureBox16.Visible = False
                    PictureBox16.Image = Nothing
            End Select
        Else
            PictureBox16.ImageLocation = Nothing
            WBSystems.Navigate("")
        End If
    End Sub

    Private Sub ToolStripMenuItem55_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem55.Click
        AddNewSystemMedia(SystemMediaType.HSWheel, DataGridView2.SelectedCells(0).Value.ToString, AddingTypes.FromFile)
    End Sub

    Private Sub ToolStripMenuItem56_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem56.Click
        AddNewSystemMedia(SystemMediaType.HSWheel, DataGridView2.SelectedCells(0).Value.ToString, AddingTypes.FromClipboard)
    End Sub

    Private Sub ToolStripMenuItem57_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem57.Click
        AddNewSystemMedia(SystemMediaType.HSWheel, DataGridView2.SelectedCells(0).Value.ToString, AddingTypes.FromURL)
    End Sub

    Private Sub ToolStripMenuItem59_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem59.Click
        AddNewSystemMedia(SystemMediaType.HSVideo, DataGridView2.SelectedCells(0).Value.ToString, AddingTypes.FromFile)
    End Sub

    Private Sub ToolStripMenuItem60_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem60.Click
        AddNewSystemMedia(SystemMediaType.HSVideo, DataGridView2.SelectedCells(0).Value.ToString, AddingTypes.FromYouTube)
    End Sub

    Private Sub ToolStripMenuItem61_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem61.Click
        AddNewSystemMedia(SystemMediaType.HSTheme, DataGridView2.SelectedCells(0).Value.ToString, AddingTypes.FromFile)
    End Sub

    Private Sub ToolStripMenuItem63_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem63.Click
        AddNewSystemMedia(SystemMediaType.HSPointer, DataGridView2.SelectedCells(0).Value.ToString, AddingTypes.FromFile)
    End Sub

    Private Sub ToolStripMenuItem64_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem64.Click
        AddNewSystemMedia(SystemMediaType.HSPointer, DataGridView2.SelectedCells(0).Value.ToString, AddingTypes.FromClipboard)
    End Sub

    Private Sub ToolStripMenuItem65_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem65.Click
        AddNewSystemMedia(SystemMediaType.HSPointer, DataGridView2.SelectedCells(0).Value.ToString, AddingTypes.FromURL)
    End Sub

    Private Sub ToolStripMenuItem84_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem84.Click
        AddNewSystemMedia(SystemMediaType.HPBackground, DataGridView2.SelectedCells(0).Value.ToString, AddingTypes.FromFile)
    End Sub

    Private Sub ToolStripMenuItem85_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem85.Click
        AddNewSystemMedia(SystemMediaType.HPBackground, DataGridView2.SelectedCells(0).Value.ToString, AddingTypes.FromClipboard)
    End Sub

    Private Sub ToolStripMenuItem86_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem86.Click
        AddNewSystemMedia(SystemMediaType.HPBackground, DataGridView2.SelectedCells(0).Value.ToString, AddingTypes.FromURL)
    End Sub

    Private Sub ToolStripMenuItem96_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem96.Click
        AddNewSystemMedia(SystemMediaType.HPFade, DataGridView2.SelectedCells(0).Value.ToString, AddingTypes.FromFile)
    End Sub

    Private Sub ToolStripMenuItem97_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem97.Click
        AddNewSystemMedia(SystemMediaType.HPFade, DataGridView2.SelectedCells(0).Value.ToString, AddingTypes.FromClipboard)
    End Sub

    Private Sub ToolStripMenuItem98_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem98.Click
        AddNewSystemMedia(SystemMediaType.HPFade, DataGridView2.SelectedCells(0).Value.ToString, AddingTypes.FromURL)
    End Sub

    Private Sub ToolStripMenuItem134_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem134.Click
        If File.Exists(TreeviewSystems.SelectedNode.Tag) Then
            Try
                If MsgBox("Are you sure ?", MsgBoxStyle.YesNo, "Delete") = MsgBoxResult.Yes Then
                    Dim vSystemName As String = ToolStripLabel1.Text
                    'Dim vRomName As String = DataGridView2.SelectedRows(0).Cells(0).Value
                    Dim vRow As DataGridViewRow = DataGridView2.SelectedRows(0)
                    UnloadMediaFromPictureBoxes(TreeviewSystems.SelectedNode.Tag)
                    File.Delete(TreeviewSystems.SelectedNode.Tag)
                    gHS.Systems(vSystemName).RefreshSystem(False)
                    FillSystemsMediaCheckROW(DataGridView2, DataGridView2.SelectedRows(0), gHS.Systems(vSystemName))
                    DataGridView2.ClearSelection()
                    vRow.Selected = True
                End If
            Catch ex As Exception
                LogEntry(LogType._Error, "Error: " & ex.Message)
            End Try
        Else
            MsgBox("File : " & TreeviewSystems.SelectedNode.Tag & vbCrLf & "Not found !", MsgBoxStyle.Critical, "Error")
            LogEntry(LogType._Warning, "File : " & TreeviewSystems.SelectedNode.Tag & vbCrLf & "Not found !")
        End If
    End Sub

    Private Sub ToolStripMenuItem133_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem133.Click
        If File.Exists(TreeviewSystems.SelectedNode.Tag) Then
            Try
                Dim psi As New ProcessStartInfo
                Dim p As New Process
                psi.FileName = TreeviewSystems.SelectedNode.Tag
                p.StartInfo = psi
                p.Start()
            Catch ex As Exception
                LogEntry(LogType._Error, "Error: " & ex.Message)
            End Try
        Else
            MsgBox("File : " & TreeviewSystems.SelectedNode.Tag & vbCrLf & "Not found !", MsgBoxStyle.Critical, "Error")
            LogEntry(LogType._Warning, "File : " & TreeviewSystems.SelectedNode.Tag & vbCrLf & "Not found !")
        End If
    End Sub

    Private Sub ToolStripMenuItem132_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem132.Click
        If Directory.Exists(Path.GetDirectoryName(TreeviewSystems.SelectedNode.Tag)) Then
            Try
                Dim psi As New ProcessStartInfo
                Dim p As New Process
                psi.FileName = Path.GetDirectoryName(TreeviewSystems.SelectedNode.Tag)
                p.StartInfo = psi
                p.Start()
            Catch ex As Exception
                LogEntry(LogType._Error, "Error: " & ex.Message)
            End Try
        Else
            MsgBox("Folder : " & Path.GetDirectoryName(TreeviewSystems.SelectedNode.Tag) & vbCrLf & "Not found !", MsgBoxStyle.Critical, "Error")
            LogEntry(LogType._Warning, "Folder : " & Path.GetDirectoryName(TreeviewSystems.SelectedNode.Tag) & vbCrLf & "Not found !")
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim vSystem As String
        Try
            vSystem = TreeView1.SelectedNode.Text
        Catch ex As Exception
            MsgBox("Please select a valid System on the left.", MsgBoxStyle.Information, "Information")
            Exit Sub
        End Try
        Select Case vSystem
            Case "Systems", "Support", "HyperT00ls", "Hyperspin-FE", "Emumovies"
                MsgBox("Please select a valid System on the left." & vbCrLf & "Currently selected is not valid.", MsgBoxStyle.Information, "Information")
            Case Else
                Dim FrmReports As New FrmReport(vSystem)
                FrmReports.Show()
                FrmReports.Focus()
        End Select
    End Sub

    Private Sub TreeviewSystems_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TreeviewSystems.MouseUp
        If TreeviewSystems.GetNodeCount(True) > 0 Then
            If e.Button = MouseButtons.Right Then
                Try
                    If TreeviewSystems.GetNodeAt(New Point(e.X, e.Y)).Tag = "" Or Path.GetExtension(TreeviewSystems.GetNodeAt(New Point(e.X, e.Y)).Tag) = "" Then
                        ToolStripMenuItem134.Enabled = False
                        ToolStripMenuItem133.Enabled = False
                        ToolStripMenuItem132.Enabled = False
                    Else
                        ToolStripMenuItem134.Enabled = True
                        ToolStripMenuItem133.Enabled = True
                        ToolStripMenuItem132.Enabled = True
                    End If
                Catch ex As Exception
                    ToolStripMenuItem134.Enabled = False
                    ToolStripMenuItem133.Enabled = False
                    ToolStripMenuItem132.Enabled = False
                End Try
                TreeviewSystems.ContextMenuStrip.Show(TreeviewSystems, New Point(e.X, e.Y))
            End If
        End If
    End Sub

#Region "Drag'n Drop"
    'DATAGRIDVIEW2
    Private Sub DataGridView2_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles DataGridView2.DragDrop
        Dim p As Point = DataGridView2.PointToClient(New Point(e.X, e.Y))
        dragIndex = DataGridView2.HitTest(p.X, p.Y).RowIndex
        If (e.Effect = DragDropEffects.Move) Then
            Dim dragRow As DataGridViewRow = CType(e.Data.GetData(GetType(DataGridViewRow)), DataGridViewRow)
            If dragIndex < 0 Then dragIndex = DataGridView2.RowCount - 1
            DataGridView2.Rows.RemoveAt(fromIndex)
            DataGridView2.Rows.Insert(dragIndex, dragRow)
        End If
    End Sub

    Private Sub DataGridView2_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGridView2.MouseDown
        fromIndex = DataGridView2.HitTest(e.X, e.Y).RowIndex
        If fromIndex > -1 Then
            Dim dragSize As Size = SystemInformation.DragSize
            dragRect = New Rectangle(New Point(e.X - (dragSize.Width / 2), e.Y - (dragSize.Height / 2)), dragSize)
        Else
            dragRect = Rectangle.Empty
        End If
    End Sub

    Private Sub DataGridView2_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGridView2.MouseMove
        If (e.Button And MouseButtons.Left) = MouseButtons.Left Then
            If (dragRect <> Rectangle.Empty _
            AndAlso Not dragRect.Contains(e.X, e.Y)) Then
                DataGridView2.DoDragDrop(DataGridView2.Rows(fromIndex), DragDropEffects.Move)
            End If
        End If
    End Sub

    Private Sub DataGridView2_DragOver(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles DataGridView2.DragOver
        e.Effect = DragDropEffects.Move
    End Sub

    'DATAGRIDVIEW4
    Private Sub DataGridView4_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles DataGridView4.DragDrop
        Dim p As Point = DataGridView4.PointToClient(New Point(e.X, e.Y))
        dragIndex = DataGridView4.HitTest(p.X, p.Y).RowIndex
        If (e.Effect = DragDropEffects.Move) Then
            Dim dragRow As DataGridViewRow = CType(e.Data.GetData(GetType(DataGridViewRow)), DataGridViewRow)
            If dragIndex < 0 Then dragIndex = DataGridView4.RowCount - 1
            DataGridView4.Rows.RemoveAt(fromIndex)
            DataGridView4.Rows.Insert(dragIndex, dragRow)
        End If
    End Sub

    Private Sub DataGridView4_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGridView4.MouseDown
        fromIndex = DataGridView4.HitTest(e.X, e.Y).RowIndex
        If fromIndex > -1 Then
            Dim dragSize As Size = SystemInformation.DragSize
            dragRect = New Rectangle(New Point(e.X - (dragSize.Width / 2), e.Y - (dragSize.Height / 2)), dragSize)
        Else
            dragRect = Rectangle.Empty
        End If
    End Sub

    Private Sub DataGridView4_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGridView4.MouseMove
        If (e.Button And MouseButtons.Left) = MouseButtons.Left Then
            If (dragRect <> Rectangle.Empty _
            AndAlso Not dragRect.Contains(e.X, e.Y)) Then
                DataGridView4.DoDragDrop(DataGridView4.Rows(fromIndex), DragDropEffects.Move)
            End If
        End If
    End Sub

    Private Sub DataGridView4_DragOver(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles DataGridView4.DragOver
        e.Effect = DragDropEffects.Move
    End Sub
#End Region

    'Double-click sur treeviewSystem node lauches default associated app
    Private Sub TreeviewSystems_NodeMouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeNodeMouseClickEventArgs) Handles TreeviewSystems.NodeMouseDoubleClick
        Try
            Dim psi As New ProcessStartInfo
            Dim p As New Process
            psi.FileName = e.Node.Tag
            p.StartInfo = psi
            p.Start()
        Catch ex As Exception
            LogEntry(LogType._Error, "Error: " & ex.Message)
        End Try
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim frmRen As New FrmRenameRom(ToolStripLabel2.Text, DataGridView4.SelectedCells(0).RowIndex)
        frmRen.Show()
    End Sub

    Private Sub DataGridView4_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView4.SelectionChanged
        If DataGridView4.SelectedCells.Count > 0 Then
            Button7.Enabled = True
            Button5.Enabled = True
            ToolStripButton9.Enabled = True
        Else
            Button7.Enabled = False
            Button5.Enabled = False
            ToolStripButton9.Enabled = False
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        SaveHyperT00lsParameters()
    End Sub

    Private Sub TreeView2_NodeMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeNodeMouseClickEventArgs) Handles TreeView2.NodeMouseClick
        Select Case e.Node.Text
            Case "Hyperspin", "HyperPause"
                For Each vNode As TreeNode In e.Node.Nodes
                    vNode.Checked = e.Node.Checked
                Next
        End Select
    End Sub

    'compare & fill with other XML
    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Dim openFileDialog1 As New OpenFileDialog()
        Dim DlgFilter As String = ""
        'openFileDialog1.RestoreDirectory = True
        openFileDialog1.InitialDirectory = gHSPath
        DlgFilter = "XML (*.xml)|*.xml"
        openFileDialog1.Filter = DlgFilter
        openFileDialog1.FilterIndex = 1
        If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Dim XMLPath As String = openFileDialog1.FileName
            'Desc
            Dim vXML As New XmlDocument()
            If File.Exists(XMLPath) = False Then
                MsgBox("Ooops ! " & XMLPath & " is not found, exiting.", MsgBoxStyle.Critical, "Error")
                LogEntry(LogType._Error, "Ooops ! " & XMLPath & " is not found, exiting.")
                'Form1.Close()
                Exit Sub
            End If

            Dim vNODES As XmlNodeList
            vXML.Load(XMLPath)
            vNODES = vXML.SelectNodes("menu/game")

            For Each vROw As DataGridViewRow In DataGridView4.Rows
                Dim CurrRom As String = vROw.Cells(0).Value
                Dim pName As String
                For Each vNode As XmlNode In vNODES
                    pName = vNode.Attributes.GetNamedItem("name").InnerText
                    If pName = CurrRom Then
                        'we have our info ;)

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

                        For Each vAttribute As Xml.XmlNode In vNode.ChildNodes
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
                                Case Else
                                    '
                            End Select
                        Next

                        If vROw.Cells(1).Value = "" Then
                            vROw.Cells(1).Value = pDescription
                        End If
                        If vROw.Cells(2).Value = "" Then
                            vROw.Cells(2).Value = pManufacturer
                        End If
                        If vROw.Cells(3).Value = "" Then
                            vROw.Cells(3).Value = pYear
                        End If
                        If vROw.Cells(4).Value = "" Then
                            vROw.Cells(4).Value = pGenre
                        End If
                        If vROw.Cells(5).Value = "" Then
                            vROw.Cells(5).Value = pRating
                        End If
                        If vROw.Cells(6).Value = "" Then
                            vROw.Cells(6).Value = pCRC
                        End If
                        If vROw.Cells(7).Value = "" Then
                            vROw.Cells(7).Value = pCloneOf
                        End If
                    End If
                Next
            Next
        Else
            Exit Sub
        End If

    End Sub

    Private Sub LoadFULLSystem(ByVal systemName As String)
        Dim vSystem As String
        Try
            vSystem = systemName
        Catch ex As Exception
            MsgBox("Please select a valid System on the left.", MsgBoxStyle.Information, "Information")
            Exit Sub
        End Try

        DataGridView1.Rows.Clear()
        DataGridView4.Rows.Clear()
        DataGridView3.SuspendLayout()

        ToolStripLabel2.Text = vSystem
        ListBox1.Items.Clear()
        UnloadMediaFromPictureBoxes()
        UnloadSystemMediaFromPictureBoxes()

        ToolStripLabel1.Text = vSystem
        ToolStripLabel2.Text = vSystem

        Select Case vSystem
            Case "Systems", "Support", "HyperT00ls", "Hyperspin-FE", "Emumovies"
                MsgBox("Please select a valid System on the left." & vbCrLf & "Currently selected (" & vSystem & ") is not valid.", MsgBoxStyle.Information, "Information")
                Button12.Enabled = False
                Button13.Enabled = False
                Button14.Enabled = False
                Button2.Enabled = False
                Button3.Enabled = False
                Button8.Enabled = False
                ToolStripButton6.Enabled = False
                ToolStripButton11.Enabled = False
            Case Else
                If gHS.Systems(vSystem).IsEXE = False Then
                    If File.Exists(gHS.Systems(vSystem).XMLPath) Then
                        gHS.Systems(vSystem).Roms.Clear()
                        gHS.Systems(vSystem).LoadSystemRoms(gCheckRoms)

                        'part1
                        FillRomsMediaCheckDatagrid(vSystem, DataGridView1)
                        Button4.Enabled = True

                        'part2
                        FillRomsXMLDatagrid(vSystem, DataGridView4, gCheckRoms)
                        Button12.Enabled = True
                        If gCheckRoms = True Then
                            Button13.Enabled = True
                        Else
                            Button13.Enabled = False
                        End If
                        Button14.Enabled = True
                        Button8.Enabled = True
                        ToolStripButton6.Enabled = True
                        ToolStripButton11.Enabled = True
                    Else
                        MsgBox("Oops, """ & vSystem & ".xml"" does not exist in your ""HyperSpin\Databases\" & vSystem & """ folder.", MsgBoxStyle.Information, "Information")
                        Button12.Enabled = False
                        Button13.Enabled = False
                        Button14.Enabled = False
                        Button8.Enabled = False
                        ToolStripButton6.Enabled = False
                        ToolStripButton11.Enabled = False
                    End If
                Else
                    ToolStripStatusLabel1.Text = "System is EXE, and thus, does not have any roms inside"
                    Button12.Enabled = False
                    Button13.Enabled = False
                    Button14.Enabled = False
                    Button8.Enabled = False
                    ToolStripButton6.Enabled = False
                    ToolStripButton11.Enabled = False
                End If
        End Select
        DataGridView3.ResumeLayout()
    End Sub

    Private Sub ToolStripMenuItem67_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem67.Click
        AddNewSystemMedia(SystemMediaType.HPBezel, DataGridView2.SelectedCells(0).Value.ToString, AddingTypes.FromFile)
    End Sub

    Private Sub ToolStripMenuItem68_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem68.Click
        AddNewSystemMedia(SystemMediaType.HPBezel, DataGridView2.SelectedCells(0).Value.ToString, AddingTypes.FromClipboard)
    End Sub

    Private Sub ToolStripMenuItem69_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem69.Click
        AddNewSystemMedia(SystemMediaType.HPBezel, DataGridView2.SelectedCells(0).Value.ToString, AddingTypes.FromURL)
    End Sub

    Private Sub ToolStripButton9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton9.Click
        Dim vRow As Integer = DataGridView4.SelectedCells(0).RowIndex
        Dim vRom As String = DataGridView4.Rows(vRow).Cells(0).Value.ToString
        Dim vSystem As String = ToolStripLabel2.Text

        If vSystem = "" Or vSystem = "System" Or vRom = "" Then
            Exit Sub
        End If

        Try
            Dim psi As New ProcessStartInfo
            Dim p As New Process
            psi.FileName = gHLPath & "\HyperLaunch.exe"
            psi.Arguments = """" & vSystem & """ """ & vRom & """"
            LogEntry(LogType._Info, "Launching : " & psi.FileName & " with arguments: " & psi.Arguments)
            p.StartInfo = psi
            p.Start()
        Catch ex As Exception
            LogEntry(LogType._Info, "Error: cannot launch " & vRom & " in System:" & vSystem & " ! -> " & ex.Message)
        End Try
    End Sub

    Private Sub ToolStripButton10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton10.Click
        Dim vRow As Integer = DataGridView1.SelectedCells(0).RowIndex
        Dim vRom As String = DataGridView1.Rows(vRow).Cells(0).Value.ToString
        Dim vSystem As String = ToolStripLabel1.Text

        If vSystem = "" Or vSystem = "System" Or vRom = "" Then
            Exit Sub
        End If

        Try
            Dim psi As New ProcessStartInfo
            Dim p As New Process
            psi.FileName = gHLPath & "\HyperLaunch.exe"
            psi.Arguments = """" & vSystem & """ """ & vRom & """"
            LogEntry(LogType._Info, "Launching : " & psi.FileName & " with arguments: " & psi.Arguments)
            p.StartInfo = psi
            p.Start()
        Catch ex As Exception
            LogEntry(LogType._Info, "Error: cannot launch " & vRom & " in System:" & vSystem & " ! -> " & ex.Message)
        End Try
    End Sub

    'YOUTUBE BATCH DOWNLOADER
    Private Sub ToolStripButton11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton11.Click
        'youtube-dl "ytsearch:allintitle:3D Twist & Match PSP" -o "c:\location you want your videos\3D Twist & Match (USA) (minis) (PSN).mp4" --max-quality 18

    End Sub
End Class

'HYPERLAUNCH MEDIA
'HyperPause Media Quick Start Guide:
'Artwork
'Backgrounds
'Bezels
'Controller
'Fade
'Fonts
'Guides
'Manuals
'Menu Images
'MultiGame
'Music
'Sounds
'Videos
'Wheels

'    Manuals
'        Preferred format: PDFs and PNGs
'        Accepted formats: PDF, TXT, PNG, BMP, GIF, JPG and TIP. Compressed Image files in ZIP, RAR, 7Z, LZH, GZIP, and TAR formats are also supported.
'        Main Contents: Game, system, controller and emulator manuals
'        Folder Location:
'            general system manuals: .\HyperLaunch\Media\Manuals\%systemname%\_Default\ Console Manual.pdf
'            game manuals: .\HyperLaunch\Media\Manuals\%systemname%\%romname% \Manual.pdf (if the game is a multidisc game and has manual contents that serves to all discs you should place the manual on .\HyperLaunch\Media\Manuals\%systemname%\%romnamew ithoutdiscinfo%\Manual.pdf)
'            Scanned Manual images can be grouped together by simply copying all image files on the same folder and placing the folder on the HyperPause manual folder (for example, .\HyperLaunch\Media\Manuals\Sony Playstation\Final Fantasy VII (USA)\Scanned Manual\image *.png).
'            You can also group images by compressing them at the same file together. Just remember that the decompress time will cause a small delay at HyperPause first run. (for example, .\HyperLaunch\Media\Manuals\Sony Playstation\Final Fantasy VII (USA)\Scanned Manual.7z).

'    Guides
'        Preferred format: TXTs and PNGs
'        Accepted formats: PDF, TXT, PNG, BMP, GIF, JPG and TIP. Compressed Image files in ZIP, RAR, 7Z, LZH, GZIP, and TAR formats are also supported.
'        Main Contents: txt walkthroughs and Map images
'        Folder Location:
'            game guides: .\HyperLaunch\Media\Guides\%systemname%\%romname%\ Walkthroughs.txt (if the game is a multidisc game and has guide contents that serves to all discs you should place the guide file on .\HyperLaunch\Media\Guides\%systemname%\%romnamewi thoutdiscinfo%\Walkthroughs.txt)
'            Scanned Map Guide images can be grouped by simply copying all image files on the same folder and placing the folder on the HyperPause manual folder (for example, .\HyperLaunch\Media\Guides\Sony PlayStation\Final Fantasy VII (USA)\Maps\image *.png)
'            You can also group images by compressing them at the same file together. Just remember that the decompress time will cause a small delay at HyperPause first run. (for example, .\HyperLaunch\Media\Manuals\Sony Playstation\Final Fantasy VII (USA)\Maps.7z).
'        Warning: If you can, avoid using too many and too large txt files as this will slow down HP loading.

'    Videos
'        Preferred format: avi, wmv and mp4
'        Accepted formats: any format that windows media player can reproduce. Try to stay with the natively reproduced ones to avoid compatibility issues between users.
'        Main Contents: System and game commercials and game video guides.
'        Folder Location:
'            general system videos: .\HyperLaunch\Media\Videos\%systemname%\commercial 1.avi
'            game videos: .\HyperLaunch\Media\Videos\%systemname%\%romname%\ Video Guide.avi (if the game is a multidisc game and has video guide contents that serves to all discs you should place the video files on .\HyperLaunch\Media\Videos\%systemname%\%romnamewi thoutdiscinfo%\Video Guide.avi)

'    Music
'        Preferred format: mp3 (others: m4a, wav, mid, wma)
'        Accepted formats: any format that windows media player can reproduce. Try to stay with the natively reproduced ones to avoid compatibility issues between users.
'        Main Contents: Game soundtracks, system advert and commercial music
'        Folder Location:
'            general system sounds: .\HyperLaunch\Media\Music\%systemname%\musicfile.m p3
'            game sounds: .\HyperLaunch\Media\Music\%systemname%\%romname%\m usicfile.mp3 (if the game is a multidisc game and has music guide contents that serves to all discs you should place the muisc files on .\HyperLaunch\Media\Music\%systemname%\%romnamewit houtdiscinfo%\musicfile.mp3, these music files will only be considered if no music file is found at the romname folder)
'        Warning, HP creates a playlist file to reproduce these songs. If you add more music files you should delete the previous created playlist file to allow HyperPause to create a new one with the new music additions. 

'    Backgrounds
'        Preffered format: PNGs
'        Accepted formats: PNG, BMP, GIF, JPG and TIP.
'        Main Contents: Background Images
'        Folder Location:
'            general HP background: .\HyperLaunch\Media\Backgrounds\_Default\backgroun d 1.png
'            general system backgrounds: .\HyperLaunch\Media\Backgrounds\%systemname%\_Defa ult\background 1.png
'            game sounds: .\HyperLaunch\Media\Backgrounds\%systemname%\%romn ame%\background 1.png (if the game is a multidisc game and you want to use the background image to all discs you should named like this: .\HyperLaunch\Media\Backgrounds\%systemname%\%romn amewithoutdiscinfo%\background 1.png)
'        You can add as many background images as you like to the above described folders. HyperPause and Multigame features will choose the on to show aleatory at each time that you enter the menu.

'    Controls
'        Preferred format: PNGs
'        Accepted formats: PDF, TXT, PNG, BMP, GIF, JPG and TIP. Compressed Image files in ZIP, RAR, 7Z, LZH, GZIP, and TAR formats are also supported.
'        Main Contents: Controller Images Created on CPWizard batch export.
'        Folder Location:
'            general system controller image: .\HyperLaunch\Media\Controller\%systemname%\_Defau lt\theNameThatYouWantForTheControllerShowingInTheM enu.png
'            game controller image: .\HyperLaunch\Media\Controller\%systemname%\%romna me%\theNameThatYouWantForTheControllerShowingInThe Menu.png (if the game is a multidisc game and you want to use the controller image to all discs you should named like this: .\HyperLaunch\Media\Controller\%systemname%\%romna mewithoutdiscinfo%\theNameThatYouWantForTheControl lerShowingInTheMenu.png)
'        Warning: This is a much more user dependable feature.

'    Artwork
'        Preferred format: PNGs
'        Accepted formats: PDF, TXT, PNG, BMP, GIF, JPG and TIP. Compressed Image files in ZIP, RAR, 7Z, LZH, GZIP, and TAR formats are also supported.
'        Main Contents: Artwork Images related with the games. Marquees, cabinet images, console images, ...
'        Folder Location:*
'            general system controller image: .\HyperLaunch\Media\Artwork\%systemname%\_Default\ theNameThatYouWantForTheControllerShowingInTheMenu .png
'            game controller image: .\HyperLaunch\Media\Artwork\%systemname%\%romname% \theNameThatYouWantForTheControllerShowingInTheMen u.png (if the game is a multidisc game and you want to use the controller image to all discs you should named like this: .\HyperLaunch\Media\Artwork\%systemname%\%romnamew ithoutdiscinfo%\theNameThatYouWantForTheController ShowingInTheMenu.png)

'    Game Info
'        The ini files present at the .\HyperLaunch\Data\Game Info\ folder describe every info that it will be show at HyperPause screen when a game is show.
'        HyperPause will try to find the ini section corresponding to the game in these files, however it will not always find it because the name differs from the HyperSpin default game names.
'        If someone can try to be responsible for that, it would be nice if the ini files would be reorganized according the HyperSpin xml databases following the HyperSpin naming convention and only describing games that are on the HyperSpin xmls.
'        Rain created a greate utility to produce the base game info files for any system that does not have it. If you want to request a game info ini file for any system not present at the HyperLaunch release, just let us know.
'        If a game is a multidisc game you can just remove the (Disc 1) part of the rom name to show the same description on all discs.
'        Anyone is up to the task???? 

'    Moves Lists
'        The folder .\HyperLaunch\Media\Moves List\Icons contains the moves list icons. You can change then for whatever you like it more. Just remember to keep the same file name.
'        The files present at the .\HyperLaunch\Data\Moves List\ folder have the info contained on the command.dat project for the moves list descriptions (http://home.comcast.net/~plotor/).
'        For more info please read: https://sites.google.com/site/hyperl...oves-list-menu