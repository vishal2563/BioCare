Imports System
Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Data.SqlClient
Imports System.Data
Imports System.Configuration
Imports Excel = Microsoft.Office.Interop.Excel
Public Class Earlyname
    Dim cmd, cmd1, cmd2 As New SqlCommand
    Dim drd, drd1 As SqlDataReader
    Dim drd2 As SqlDataAdapter
    Dim r As Integer = 0
    Dim sqlstmt As String = ""
    Public dr
    Dim ENo As String = ""
    Dim str As String
    Dim shec As String

    Private Sub Earlyname_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    End Sub
    Public Function SaveTextToFile(ByVal strData As String, ByVal FullPath As String, Optional ByVal ErrInfo As String = "") As Boolean
        'Dim Contents As String
        Dim Saved As Boolean = False
        Dim objReader As IO.StreamWriter
        Try
            objReader = New IO.StreamWriter(FullPath)
            objReader.Write(strData)
            objReader.Close()
            Saved = True
        Catch Ex As Exception
            ErrInfo = Ex.Message
        End Try
        Return Saved
    End Function

    Private Sub visitMicrosoft2()
        System.Diagnostics.Process.Start("E:\Earlyname.xls")
    End Sub
    Private Sub created()
        Dim HTMLTitle As String = ""
        Dim HTMLText As String = ""
        Dim HTMLFileName As String = ""
        Dim strFile As String
        'Dim str As Date
        'Dim rowcounter As Integer
        'Dim rowtablecounter As Integer
        'rowcounter = 0
        'rowtablecounter = 0
        ' ----------------------
        ' -- Prepare String --
        ' ----------------------
        strFile = ""

        strFile = strFile & "<html>" & vbNewLine
        strFile = strFile & "<head>" & vbNewLine
        strFile = strFile & "<title>" & HTMLTitle & "</title>" & vbNewLine
        strFile = strFile & "</head><body>"
        strFile = strFile & " <table>"
        strFile = strFile & "<tr> "
        strFile = strFile & "<td>Date</td> "
        strFile = strFile & "<td>Name</td> "

        strFile = strFile & " <td>scheduled in time </td>"
        strFile = strFile & " <td>	<td>actual time </td>"
        strFile = strFile & " <td>late by &lt;30 min </td>"
        strFile = strFile & " <td>late by &lt;30 min </td>"
        strFile = strFile & " </tr>"
        strFile = strFile & " </table>"
        strFile = strFile & "  </form>"
        strFile = strFile & "  </body></html>"
        SaveTextToFile(strFile, "E:\" & HTMLFileName & "Earlyname.xls")
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        created()
        visitMicrosoft2()
    End Sub
End Class