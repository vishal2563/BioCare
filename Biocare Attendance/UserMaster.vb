Imports System.Data.SqlClient
Imports System.Data
Imports System.IO
Public Class UserMaster
    Dim cmd, cmd1 As New SqlCommand
    Dim drd, drd1 As SqlDataReader
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs)
        AddRecords()
    End Sub
    Private Sub AddRecords()


        openconn()
        If TextBox1.Text = "" Then
            MsgBox("Enter the User Name")
        ElseIf TextBox2.Text = "" Then
            MsgBox("Enter the password")
        ElseIf TextBox2.Text <> TextBox3.Text Then
            MsgBox("Incorrect Password")

        Else

            cmd = New SqlCommand("insert into usermaster(username,pass) values (@username,@pass)", cnn)
            cmd.Parameters.AddWithValue("@username", (TextBox1.Text))
            cmd.Parameters.AddWithValue("@pass", (TextBox2.Text))
            cmd.ExecuteNonQuery()
            MsgBox("Record Enter")
        End If
       
        empty()
        closeconn()


    End Sub
    Sub empty()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""

    End Sub
    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs)
        Me.Close()

    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs)
        If Asc(e.KeyChar) = 13 Then
            SendKeys.Send("{tab}")
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs)
        If Asc(e.KeyChar) = 13 Then
            SendKeys.Send("{tab}")
        End If
    End Sub

    Private Sub TextBox2_TextChanged(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub TextBox3_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs)
        If Asc(e.KeyChar) = 13 Then
            SendKeys.Send("{tab}")
        End If
    End Sub

    Private Sub TextBox3_TextChanged(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub UserMaster_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'Me.WindowState = 1
        'Me.WindowState = FormWindowState.Maximized
    End Sub

    Private Sub Button2_Click_1(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click_1(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        AddRecords()
    End Sub
End Class