Option Explicit On
Imports System
Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Data.SqlClient
Imports System.Data
Imports System.Configuration
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices.Marshal
Public Class AbsentWise_Report
    Dim cmd, cmd1, cmd2 As New SqlCommand
    Dim drd, drd1 As SqlDataReader
    Dim drd2 As SqlDataAdapter
    Dim r As Integer = 0
    Dim sqlstmt As String = ""
    Public dr
    Dim ENo As String = ""
    Private Property DataSet As DataTable
    Private Sub AbsentWise_Report_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.StartPosition = FormStartPosition.Manual
    End Sub
    Private Sub fetchasbsent()
        Dim percentage As Decimal = 0

        Try
            openconn()
            r = 0
            DataGridView1.Rows.Clear()

            'cmd = New SqlCommand("SELECT DISTINCT empcode, empname,tdate, '26'AS Total_Working_Days, COUNT(status) AS Absent,COUNT(status) * 100 / 26 AS percentage FROM  Attendace WHERE tdate >='" & (DateTimePicker1.Value.ToString("MM/dd/yyyy")) & "' and tdate <='" & (DateTimePicker2.Value.ToString("MM/dd/yyyy")) & "' GROUP BY empname, empcode,tdate ", cnn)
            ' cmd1 = New SqlCommand("SELECT DISTINCT empcode, empname,'26'AS Total_Working_Days, COUNT(status) AS Absent,COUNT(status) * 100 / 26 AS percentage FROM  Attendace WHERE status ='A' and tdate >='" & DateTimePicker1.Value.Date.ToString("MM/dd/yyyy") & "' and tdate <='" & DateTimePicker2.Value.Date.ToString("MM/dd/yyyy") & "' GROUP BY empname, empcode ", cnn)
            cmd1 = New SqlCommand("SELECT DISTINCT empcode, empname,srno, COUNT(status) AS Absent  FROM  Attendace WHERE status ='A' and tdate >='" & DateTimePicker1.Value.Date.ToString("MM/dd/yyyy") & "' and tdate <='" & DateTimePicker2.Value.Date.ToString("MM/dd/yyyy") & "' and companyname ='" & ComboBox2.Text & "' GROUP BY empname, empcode,srno ORDER BY empname ", cnn)
            ' MsgBox("hii")
            drd = cmd1.ExecuteReader
            'MsgBox("hello")
            While drd.Read
                'MsgBox("v")
                DataGridView1.Rows.Add()
                DataGridView1.Rows(r).Cells(0).Value = drd("empcode")
                DataGridView1.Rows(r).Cells(1).Value = drd("empname")
                ' DataGridView1.Rows(r).Cells(2).Value = drd("Total_Working_Days")
                DataGridView1.Rows(r).Cells(2).Value = TextBox2.Text
                DataGridView1.Rows(r).Cells(3).Value = drd("Absent")
                percentage = Val(DataGridView1.Rows(r).Cells(3).Value) / Val(DataGridView1.Rows(r).Cells(2).Value) * 100
                DataGridView1.Rows(r).Cells(4).Value = Math.Round(percentage, 2)
                'DataGridView1.Rows(r).Cells(5).Value = drd("tdate")
                r = r + 1
            End While
            drd.Close()
            closeconn()
        Catch ex As Exception
        End Try
        If DataGridView1.Rows.Count = 1 Then
            MsgBox("Record Empty")
        Else
        End If

    End Sub
   
    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)

            'System.rumtime.interopservice.marshal.releasecomobject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
    Private Sub totaldys()
        If Val(TextBox1.Text) < 12 Then
            Label7.Text = 1
        ElseIf Val(TextBox1.Text) > 13 And Val(TextBox1.Text) < 21 Then
            Label7.Text = 2
        ElseIf Val(TextBox1.Text) >= 21 And Val(TextBox1.Text) < 28 Then
            Label7.Text = 3
        ElseIf Val(TextBox1.Text) >= 28 And Val(TextBox1.Text) <= 31 Then
            Label7.Text = 4
        End If
        TextBox2.Text = Val(TextBox1.Text) - Val(Label7.Text)
    End Sub
    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        ' Try
        '    ' MsgBox("hello")
        '    Button3.Text = "Please Wait"
        '    Button3.Enabled = False
        '    SaveFileDialog1.Filter = "Excel document (*.xlsx)|*.xlsx"
        '    If SaveFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then


        '        Dim xlpp As Microsoft.Office.Interop.Excel.Application
        '        Dim xlworkbook As Microsoft.Office.Interop.Excel.Workbook
        '        Dim xlworksheet As Microsoft.Office.Interop.Excel.Worksheet
        '        Dim misvalue As Object = System.Reflection.Missing.Value
        '        Dim i As Integer
        '        Dim j As Integer



        '        xlpp = New Microsoft.Office.Interop.Excel.Application
        '        xlworkbook = xlpp.Workbooks.Add(misvalue)
        '        xlworksheet = xlworkbook.Sheets("sheet1")

        '        For i = 0 To DataGridView1.RowCount - 2
        '            For j = 0 To DataGridView1.ColumnCount - 1
        '                For k As Integer = 1 To DataGridView1.Columns.Count

        '                    xlworksheet.Cells(1, k) = DataGridView1.Columns(k - 1).HeaderText
        '                    xlworksheet.Cells(i + 1, j + 1) = DataGridView1(j, i).Value.ToString()

        '                Next
        '            Next
        '        Next
        '        xlworksheet.SaveAs(SaveFileDialog1.FileName)
        '        xlworkbook.Close()
        '        xlpp.Quit()
        '        releaseObject(xlpp)
        '        releaseObject(xlworkbook)
        '        releaseObject(xlworksheet)
        '        MsgBox("save")
        '        Button3.Text = "Export to Ms Excel"
        '        'Button2.Enabled = False
        '    End If
        'Catch ex As Exception
        '    Return
        'End Try
        CreatePage2()
        visitMicrosoft2()
    End Sub
    Private Sub visitMicrosoft2()
        System.Diagnostics.Process.Start("D:\AdsentReport.xls")
    End Sub
    Private Sub CreatePage2()


        Dim HTMLTitle As String = ""
        Dim HTMLText As String = ""
        Dim HTMLFileName As String = ""
        Dim strFile As String
        Dim dt As DateTime = DateTime.Now
        Dim ts As TimeSpan
        Dim time1 As TimeSpan
        Dim rowcounter As Integer
        Dim rowtablecounter As Integer
        rowcounter = 0
        rowtablecounter = 0
        ' ----------------------
        ' -- Prepare String --
        ' ----------------------
        strFile = ""

        strFile = strFile & "<html>" & vbNewLine
        strFile = strFile & "<head>" & vbNewLine
        strFile = strFile & "<title>" & HTMLTitle & "</title>" & vbNewLine
        strFile = strFile & "</head><body>"

        strFile = strFile & " <table style='width: 450px' border=0>"
        strFile = strFile & "  <tr> "
        strFile = strFile & " <td align='right'> "
        strFile = strFile & " </td>"
        strFile = strFile & "</tr>"
        strFile = strFile & "</table>"
        strFile = strFile & " <center><b> <Font Size=4> " & Label1.Text & "</Font></b></center>"
        strFile = strFile & "<center><b></b></center>"

        strFile = strFile & " <center><b> <Font Size=4> For the period:" & DateTimePicker1.Text & " to " & DateTimePicker2.Text & "</Font></b></center>"
        strFile = strFile & "<center><b></b></center>"

         strFile = strFile & " <table  border='1'>"
        strFile = strFile & "  <tr> "
        strFile = strFile & "  <td> "


        strFile = strFile & " <table style='width: 0' border='0'>"
        strFile = strFile & "  <tr> "
        For Each col In DataGridView1.Columns

            If col.Visible = True Then

                strFile = strFile & "      <td style=height: 20px; 'width: 0'>"
                strFile = strFile & "     <center> <b> " & col.HeaderText & " </b> </center></td>"

            End If

        Next

        strFile = strFile & "  </tr>"

        Dim col2 As DataGridViewColumn

        Dim R As DataGridViewRow

        For Each R In Me.DataGridView1.Rows

            strFile = strFile & "  <tr>"
            For Each col2 In DataGridView1.Columns
                If col2.Visible = True Then

                    strFile = strFile & "       <td style='Height: 21px'> <center>"

                    strFile = strFile & " " & R.Cells(col2.Index).Value & " </center></td>"

                End If
            Next

            strFile = strFile & "  <tr>"

        Next
        strFile = strFile & "</table>"
        strFile = strFile & "  </form>"
        strFile = strFile & "  </body></html>"
        SaveTextToFile(strFile, "D:\" & HTMLFileName & "AdsentReport.xls")
    End Sub

    Public Function SaveTextToFile(ByVal strData As String, ByVal FullPath As String, Optional ByVal ErrInfo As String = "") As Boolean
        Dim Contents As String
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
    Private Sub DATEDIFF()
        Dim diff As System.TimeSpan = DateTimePicker2.Value.Date.Subtract(DateTimePicker1.Value.Date)
        Dim diff1 As System.TimeSpan = DateTimePicker2.Value.Date - DateTimePicker1.Value.Date

        Dim diff2 As String = (DateTimePicker2.Value.Date - DateTimePicker1.Value.Date).TotalDays.ToString()
        TextBox1.Text = Val(diff2) + 1
        ' MsgBox(diff2.ToString())
    End Sub
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        fetchasbsent()
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Me.Close()

    End Sub

    Private Sub Label1_Click(sender As System.Object, e As System.EventArgs) Handles Label1.Click

    End Sub
    Private Sub datedchage()
        Dim i, y, m As Integer

        '  i = DateTimePicker1.Value.Date.Month
        y = DateTimePicker1.Value.Date.Year

        If ComboBox1.Text = "January" Then

            i = 1
        ElseIf ComboBox1.Text = "February" Then
            i = 2

        ElseIf ComboBox1.Text = "March" Then

            i = 3
        ElseIf ComboBox1.Text = "April" Then
            i = 4

        ElseIf ComboBox1.Text = "May" Then
            i = 5

        ElseIf ComboBox1.Text = "June" Then
            i = 6

        ElseIf ComboBox1.Text = "July" Then
            i = 7
        ElseIf ComboBox1.Text = "August" Then
            i = 8
        ElseIf ComboBox1.Text = "September" Then

            i = 9
        ElseIf ComboBox1.Text = "October" Then
            i = 10
        ElseIf ComboBox1.Text = "November" Then
            i = 11
        ElseIf ComboBox1.Text = "December" Then
            i = 12

        End If

        If i = 1 Then

            DateTimePicker1.Value = "1/" & i & "/" & y

            DateTimePicker2.Value = "31/" & i & "/" & y
        End If


        If i = 2 Then
            DateTimePicker1.Value = "1/" & i & "/" & y

            DateTimePicker2.Value = "28/" & i & "/" & y
        End If
        If i = 3 Then

            DateTimePicker1.Value = "1/" & i & "/" & y

            DateTimePicker2.Value = "31/" & i & "/" & y
        End If
        If i = 4 Then
            DateTimePicker1.Value = "1/" & i & "/" & y

            DateTimePicker2.Value = "30/" & i & "/" & y
        End If

        If i = 5 Then
            DateTimePicker1.Value = "1/" & i & "/" & y

            DateTimePicker2.Value = "31/" & i & "/" & y
        End If
        If i = 6 Then
            DateTimePicker1.Value = "1/" & i & "/" & y

            DateTimePicker2.Value = "30/" & i & "/" & y
        End If


        If i = 7 Then

            DateTimePicker1.Value = "1/" & i & "/" & y

            DateTimePicker2.Value = "31/" & i & "/" & y
        End If

        If i = 8 Then
            DateTimePicker1.Value = "1/" & i & "/" & y

            DateTimePicker2.Value = "31/" & i & "/" & y
        End If

        If i = 9 Then
            DateTimePicker1.Value = "1/" & i & "/" & y

            DateTimePicker2.Value = "30/" & i & "/" & y
        End If
        If i = 10 Then
            DateTimePicker1.Value = "1/" & i & "/" & y

            DateTimePicker2.Value = "31/" & i & "/" & y
        End If
        If i = 11 Then
            DateTimePicker1.Value = "1/" & i & "/" & y


            DateTimePicker2.Value = "30/" & i & "/" & y
        End If
        If i = 12 Then
            DateTimePicker1.Value = "1/" & i & "/" & y

            DateTimePicker2.Value = "31/" & i & "/" & y
        End If

    End Sub


    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub GroupBox1_Enter(sender As System.Object, e As System.EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        DATEDIFF()
        totaldys()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        datedchage()
        DATEDIFF()
        totaldys()
        'If ComboBox1.Text = "January" Then
        '    Label5.Text = 31
        '    Label6.Text = Val(Label5.Text) - 4
        'ElseIf ComboBox1.Text = "February" Then
        '    Label5.Text = 28
        '    Label6.Text = Val(Label5.Text) - 4
        'ElseIf ComboBox1.Text = "March" Then
        '    Label5.Text = 31
        '    Label6.Text = Val(Label5.Text) - 4
        'ElseIf ComboBox1.Text = "April" Then
        '    Label5.Text = 30
        '    Label6.Text = Val(Label5.Text) - 4
        'ElseIf ComboBox1.Text = "May" Then
        '    Label5.Text = 31
        '    Label6.Text = Val(Label5.Text) - 4
        'ElseIf ComboBox1.Text = "June" Then
        '    Label5.Text = 30
        '    Label6.Text = Val(Label5.Text) - 4
        'ElseIf ComboBox1.Text = "July" Then
        '    Label5.Text = 31
        '    Label6.Text = Val(Label5.Text) - 4
        'ElseIf ComboBox1.Text = "August" Then
        '    Label5.Text = 31
        '    Label6.Text = Val(Label5.Text) - 4
        'ElseIf ComboBox1.Text = "September" Then
        '    Label5.Text = 30
        '    Label6.Text = Val(Label5.Text) - 4
        'ElseIf ComboBox1.Text = "October" Then
        '    Label5.Text = 31
        '    Label6.Text = Val(Label5.Text) - 4
        'ElseIf ComboBox1.Text = "November" Then
        '    Label5.Text = 30
        '    Label6.Text = Val(Label5.Text) - 4
        '      ElseIf ComboBox1.Text = "December" Then
        '    Label5.Text = 31
        '    Label6.Text = Val(Label5.Text) - 4

        'End If

       

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
             CreatePage2()

        Dim objProcess As New System.Diagnostics.ProcessStartInfo
        With objProcess
            .FileName = "d:\AdsentReport.xls"
            .WindowStyle = ProcessWindowStyle.Hidden
            .Verb = "print"
            .CreateNoWindow = True
            .UseShellExecute = True
        End With
        '   Try
        System.Diagnostics.Process.Start(objProcess)

        'Catch ex As Exception
        ' System.IO.File.Delete("d:\Report.xls")
        '  MessageBox.Show(ex.Message)
        'End Try
        '       Try


        System.IO.File.Delete("d:\AdsentReport.xls")
    End Sub

    Private Sub DateTimePicker2_Validated(sender As Object, e As System.EventArgs) Handles DateTimePicker2.Validated
        DATEDIFF()
        totaldys()
    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As System.Object, e As System.EventArgs) Handles DateTimePicker2.ValueChanged
        DATEDIFF()
        totaldys()
    End Sub

    Private Sub Label6_Click(sender As System.Object, e As System.EventArgs) Handles Label6.Click

    End Sub
End Class