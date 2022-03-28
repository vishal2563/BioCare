Imports System
Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Data.SqlClient
Imports System.Data
Imports System.Configuration
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data.OleDb
Public Class Form1

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Addrecord()
        Try
            ' readfile()

            Dim dlg As New OpenFileDialog
            dlg.ShowDialog()
            If dlg.ShowDialog = Windows.Forms.DialogResult.OK Then
                Dim fileName As String
                fileName = dlg.FileName
                Label3.Text = fileName
                '   MsgBox(fileName)
            End If
            '  MsgBox(path)
           Dim MyConnection As System.Data.OleDb.OleDbConnection
            Dim DtSet As System.Data.DataSet
            Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
            MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Label3.Text & ";Extended Properties=Excel 8.0;")
            MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$]", MyConnection)
            MyCommand.TableMappings.Add("Table", "Net-informations.com")
            DtSet = New System.Data.DataSet
            MyCommand.Fill(DtSet)
            DataGridView1.DataSource = DtSet.Tables(0)
            MyConnection.Close()
        Catch ex As Exception
        End Try
        '' EmptyRow()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Addrecord()
    End Sub
End Class
