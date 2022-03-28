Imports System.Windows.Forms
Imports System.Data.SqlClient
Imports System.Data
Imports System.IO

Public Class Biocare
    Dim cmd, cmd1 As New SqlCommand
    Dim drd, drd1 As SqlDataReader
    Private Sub ShowNewForm(ByVal sender As Object, ByVal e As EventArgs)
        ' Create a new instance of the child form.
        Dim ChildForm As New System.Windows.Forms.Form
        ' Make it a child of this MDI form before showing it.
        ChildForm.MdiParent = Me

        m_ChildFormNumber += 1
        ChildForm.Text = "Window " & m_ChildFormNumber

        ChildForm.Show()
    End Sub

    Private Sub OpenFile(ByVal sender As Object, ByVal e As EventArgs)
        Dim OpenFileDialog As New OpenFileDialog
        OpenFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        OpenFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        If (OpenFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK) Then
            Dim FileName As String = OpenFileDialog.FileName
            ' TODO: Add code here to open the file.
        End If
    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim SaveFileDialog As New SaveFileDialog
        SaveFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        SaveFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"

        If (SaveFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK) Then
            Dim FileName As String = SaveFileDialog.FileName
            ' TODO: Add code here to save the current contents of the form to a file.
        End If
    End Sub


    Private Sub ExitToolsStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.Close()
    End Sub

    Private Sub CutToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        ' Use My.Computer.Clipboard to insert the selected text or images into the clipboard
    End Sub

    Private Sub CopyToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        ' Use My.Computer.Clipboard to insert the selected text or images into the clipboard
    End Sub

    Private Sub PasteToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        'Use My.Computer.Clipboard.GetText() or My.Computer.Clipboard.GetData to retrieve information from the clipboard.
    End Sub

  


    Private Sub CascadeToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.Cascade)
    End Sub

    Private Sub TileVerticalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.TileVertical)
    End Sub

    Private Sub TileHorizontalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.TileHorizontal)
    End Sub

    Private Sub ArrangeIconsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.ArrangeIcons)
    End Sub

    Private Sub CloseAllToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        ' Close all child forms of the parent.
        For Each ChildForm As Form In Me.MdiChildren
            ChildForm.Close()
        Next
    End Sub

    Private m_ChildFormNumber As Integer

    Private Sub Biocare_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        Me.WindowState = 1
        Me.WindowState = FormWindowState.Maximized
        For Each ctl As Control In Me.Controls
            If TypeOf ctl Is MdiClient Then
                ctl.BackColor = Me.BackColor
            End If
        Next ctl
    End Sub
    Private Sub login()
        Dim password As String
        openconn()
        cmd = New SqlCommand("Select username,pass From  usermaster Where  username='" & Trim(TextBox1.Text) & "'", cnn)
        drd = cmd.ExecuteReader
        If drd.Read Then
            Dim s As String
            s = drd("username")
            password = drd("pass")


            Dim jQ As String = s.ToString.ToUpper()
            Dim kQ As String = Trim(Me.TextBox1.Text.ToString.ToUpper())
            If jQ = kQ Then

                If password = Trim(TextBox2.Text) Then
                    Label9.Visible = True
                    Label10.Visible = True
                    Label7.Visible = True
                    Label8.Visible = True
                    Label8.Text = Environment.MachineName
                    Label9.Text = Now.ToLongDateString
                    Label10.Text = TimeOfDay.ToString("h:mm:ss tt")
                    Label7.Text = kQ
                    FileMenu.Enabled = True
                    EditMenu.Enabled = True




                    drd.Close()
                    Panel1.Visible = False
                    'GroupBox1.Visible = False

                Else

                    MsgBox("Inavalid Password")

                    Exit Sub
                End If

            Else

                Exit Sub
            End If
        Else
            MsgBox("Invalid UserName")
            Exit Sub
        End If
        ' Me.TxtName.Text = ""
        Me.TextBox1.Focus()

        closeconn()
     

        Exit Sub

    End Sub
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

        login()
    End Sub

    Private Sub EditMenu_Click(sender As System.Object, e As System.EventArgs) Handles EditMenu.Click

    End Sub

    Private Sub RedoToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub LateReportNameWiseToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles LateReportNameWiseToolStripMenuItem.Click
        namewisereport.MdiParent = Me
        namewisereport.Show()
        namewisereport.BringToFront()
        Biocare_Empolyee.Close()

    End Sub

    Private Sub ImsertReportDatabaseToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ImsertReportDatabaseToolStripMenuItem.Click
        Biocare_Empolyee.MdiParent = Me
        Biocare_Empolyee.Show()
        Biocare_Empolyee.BringToFront()

    End Sub


    Private Sub EToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles EToolStripMenuItem.Click
        EarlyNameWiseReport.MdiParent = Me
        EarlyNameWiseReport.Show()
        EarlyNameWiseReport.BringToFront()
        Biocare_Empolyee.Close()

    End Sub

    Private Sub EveningReportToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles EveningReportToolStripMenuItem.Click
        Evening_Report.MdiParent = Me
        Evening_Report.Show()
        Evening_Report.BringToFront()
        Biocare_Empolyee.Close()

    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            SendKeys.Send("{tab}")
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        If Asc(e.KeyChar) = 13 Then
            SendKeys.Send("{tab}")
        End If
    End Sub

    Private Sub TextBox2_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub Button1_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles Button1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            SendKeys.Send("{tab}")
        End If
    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripMenuItem1.Click
        AbsentWise_Report.MdiParent = Me
        AbsentWise_Report.Show()
        AbsentWise_Report.BringToFront()
        Biocare_Empolyee.Close()

    End Sub

    Private Sub UserMasterToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles UserMasterToolStripMenuItem.Click
        UserMaster.MdiParent = Me
        UserMaster.Show()
        UserMaster.BringToFront()
        Biocare_Empolyee.Close()

    End Sub

    Private Sub ChangePasswordToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs)
        Changepassword.MdiParent = Me
        Changepassword.Show()
        Changepassword.BringToFront()
    End Sub

    Private Sub ChangePasswordToolStripMenuItem_Click_1(sender As System.Object, e As System.EventArgs) Handles ChangePasswordToolStripMenuItem.Click
        Changepassword.MdiParent = Me
        Changepassword.Show()
        Changepassword.BringToFront()
        Biocare_Empolyee.Close()

    End Sub

    Private Sub LateReportToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles LateReportToolStripMenuItem.Click
        Late_Coming.MdiParent = Me
        Late_Coming.Show()
        Late_Coming.BringToFront()
        Biocare_Empolyee.Close()
    End Sub


    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub MenuStrip_ItemClicked(sender As System.Object, e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles MenuStrip.ItemClicked

    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Private Sub FromCheckToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FromCheckToolStripMenuItem.Click
        Form1.MdiParent = Me
        Form1.Show()
    End Sub

    Private Sub OverTimeReportToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles OverTimeReportToolStripMenuItem.Click
        OvertimeReport.MdiParent = Me
        OvertimeReport.Show()
        OvertimeReport.BringToFront()
        Biocare_Empolyee.Close()
    End Sub
End Class
