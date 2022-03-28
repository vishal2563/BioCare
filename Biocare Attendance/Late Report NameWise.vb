Imports System
Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Data.SqlClient
Imports System.Data
Imports System.Configuration
Imports Excel = Microsoft.Office.Interop.Excel
Public Class Late_Report_NameWise
    Dim cmd, cmd1, cmd2 As New SqlCommand
    Dim drd, drd1 As SqlDataReader
    Dim drd2 As SqlDataAdapter
    Dim r As Integer = 0
    Dim sqlstmt As String = ""
    Public dr
    Dim ENo As String = ""
    Private Property DataSet As DataTable
    Private Sub Late_Report_NameWise_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub refress()
        openconn()
        r = 0
        DataGridView1.Rows.Clear()

        cmd = New SqlCommand("select empname ,shiftstart,intime,latearr,tdate from Attendace where latearr <30 ", cnn)
        drd = cmd.ExecuteReader
        While drd.Read
            DataGridView1.Rows.Add()
            DataGridView1.Rows(r).Cells(0).Value = drd("empname")
            DataGridView1.Rows(r).Cells(1).Value = drd("shiftstart")
            DataGridView1.Rows(r).Cells(2).Value = drd("intime")
            DataGridView1.Rows(r).Cells(3).Value = drd("latearr")

            DataGridView1.Rows(r).Cells(5).Value = drd("tdate")
            r = r + 1

        End While
        drd.Close()


        cmd = New SqlCommand("select latearr from Attendace where latearr >30 ", cnn)
        drd = cmd.ExecuteReader
        While drd.Read
            DataGridView1.Rows(r).Cells(4).Value = drd("latearr")
            r = r + 1
        End While
        drd.Close()
        closeconn()
    End Sub
End Class