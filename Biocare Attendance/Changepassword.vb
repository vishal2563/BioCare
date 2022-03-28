Imports System.Data.SqlClient
Imports System.Data
Imports System.IO
Public Class Changepassword
    Dim cmd, cmd1 As New SqlCommand
    Dim drd, drd1 As SqlDataReader
    Private Sub Label1_Click(sender As System.Object, e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub updt()

        openconn()
        Dim cb As String = ("Update usermaster set pass=@d1 where username='" & (TextBox1.Text) & "' ")
        cmd = New SqlCommand(cb)
        cmd.Connection = cnn
        cmd.Parameters.AddWithValue("@d1", TextBox2.Text)
       
        cmd.ExecuteNonQuery()
        MessageBox.Show("Successfully updated", "Restaurant Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
       

        empty()
    End Sub
    Private Sub empty()
        TextBox1.Text = ""
        TextBox2.Text = ""

    End Sub
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        updt()

    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        Me.Close()
    End Sub
End Class