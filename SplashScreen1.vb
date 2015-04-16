Public NotInheritable Class SplashScreen1

    Private Sub SplashScreen1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Set up the dialog text at runtime according to the application's assembly information.  
        'Application title
        If My.Application.Info.Title <> "" Then
            ApplicationTitle.Text = My.Application.Info.Title & "'s loading ..."
        Else
            'If the application title is missing, use the application name, without the extension
            ApplicationTitle.Text = System.IO.Path.GetFileNameWithoutExtension(My.Application.Info.AssemblyName)
        End If

        'Format the version information using the text set into the Version control at design time as the
        '  formatting string.  This allows for effective localization if desired.
        '  Build and revision information could be included by using the following code and changing the 
        '  Version control's designtime text to "Version {0}.{1:00}.{2}.{3}" or something similar.  See
        '  String.Format() in Help for more information.
        '
        '    Version.Text = System.String.Format(Version.Text, My.Application.Info.Version.Major, My.Application.Info.Version.Minor, My.Application.Info.Version.Build, My.Application.Info.Version.Revision)

        Version.Text = Application.ProductVersion

        'Copyright info
        Copyright.Text = "Thanks to HyperFreeSpin forums !"

        'xx.Show()

    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Process.Start(LinkLabel1.Text)
    End Sub

    Private Sub LinkLabel1_MouseMove(ByVal sender As Object, ByVal e As System.EventArgs) Handles LinkLabel1.MouseMove
        Me.Cursor = Cursors.Hand
        Application.DoEvents()
    End Sub

    Private Sub LinkLabel1_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles LinkLabel1.MouseLeave
        Me.Cursor = Cursors.Default
        Application.DoEvents()
    End Sub

    Private Sub LinkLabel1_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles LinkLabel1.MouseEnter
        Me.Cursor = Cursors.Hand
        Application.DoEvents()
    End Sub

    Private Sub LinkLabel1_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles LinkLabel1.MouseHover
        Me.Cursor = Cursors.Hand
        Application.DoEvents()
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Process.Start(LinkLabel2.Text)
    End Sub

    Private Sub LinkLabel2_MouseMove(ByVal sender As Object, ByVal e As System.EventArgs) Handles LinkLabel2.MouseMove
        Me.Cursor = Cursors.Hand
        Application.DoEvents()
    End Sub

    Private Sub LinkLabel2_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles LinkLabel2.MouseLeave
        Me.Cursor = Cursors.Default
        Application.DoEvents()
    End Sub

    Private Sub LinkLabel2_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles LinkLabel2.MouseEnter
        Me.Cursor = Cursors.Hand
        Application.DoEvents()
    End Sub

    Private Sub LinkLabel2_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles LinkLabel2.MouseHover
        Me.Cursor = Cursors.Hand
        Application.DoEvents()
    End Sub

    Private Sub LinkLabel3_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked
        Process.Start(LinkLabel3.Text)
    End Sub

    Private Sub LinkLabel3_MouseMove(ByVal sender As Object, ByVal e As System.EventArgs) Handles LinkLabel3.MouseMove
        Me.Cursor = Cursors.Hand
        Application.DoEvents()
    End Sub

    Private Sub LinkLabel3_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles LinkLabel3.MouseLeave
        Me.Cursor = Cursors.Default
        Application.DoEvents()
    End Sub

    Private Sub LinkLabel3_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles LinkLabel3.MouseEnter
        Me.Cursor = Cursors.Hand
        Application.DoEvents()
    End Sub

    Private Sub LinkLabel3_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles LinkLabel3.MouseHover
        Me.Cursor = Cursors.Hand
        Application.DoEvents()
    End Sub

    Private Sub LinkLabel4_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel4.LinkClicked
        Process.Start(LinkLabel4.Text)
    End Sub

    Private Sub LinkLabel4_MouseMove(ByVal sender As Object, ByVal e As System.EventArgs) Handles LinkLabel4.MouseMove
        Me.Cursor = Cursors.Hand
        Application.DoEvents()
    End Sub

    Private Sub LinkLabel4_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles LinkLabel4.MouseLeave
        Me.Cursor = Cursors.Default
        Application.DoEvents()
    End Sub

    Private Sub LinkLabel4_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles LinkLabel4.MouseEnter
        Me.Cursor = Cursors.Hand
        Application.DoEvents()
    End Sub

    Private Sub LinkLabel4_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles LinkLabel4.MouseHover
        Me.Cursor = Cursors.Hand
        Application.DoEvents()
    End Sub
End Class
