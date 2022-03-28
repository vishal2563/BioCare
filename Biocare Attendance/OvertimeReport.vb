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
Public Class OvertimeReport
    Dim cmd, cmd1, cmd2 As New SqlCommand
    Dim drd, drd1 As SqlDataReader
    Dim drd2 As SqlDataAdapter
    Dim r As Integer = 0
    Dim sqlstmt As String = ""
    Public dr
    Dim str As String
    Dim shec As String
    Dim ENo As String = ""
    Private Property DataSet As DataTable
    Private Sub OvertimeReport_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.Top = 0
        'Me.WindowState = 1
        'Me.WindowState = FormWindowState.Maximized
        Microsoft.Win32.Registry.SetValue("HKEY_CURRENT_USER\Control Panel\International", "sShortDate", "dd/MM/yyyy")
        empnameref()

    End Sub

    Private Sub empnameref()
        openconn()
        r = 0
        DataGridView2.Rows.Clear()
        cmd = New SqlCommand("select distinct empname ,empcode  from Attendace where empname like'" & TextBox1.Text & "%'", cnn)
        drd = cmd.ExecuteReader
        While drd.Read
            DataGridView2.Rows.Add()
            DataGridView2.Rows(r).Cells("empname").Value = drd("empname")
            DataGridView2.Rows(r).Cells("empcode").Value = drd("empcode")
            r = r + 1
        End While
        drd.Close()
        closeconn()
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


    Private Sub refress()
        Dim Starttime As New DateTime
        Dim endtime As New DateTime
        Dim intime As New DateTime
        Dim outtime As New DateTime
        Dim newdate As Date
        Dim count As Integer = 0
        Dim countmin As Integer = 0
        Dim countzero As Integer = 0
        Dim count4 As New TimeSpan(0, 0, 0, 0)
        Dim count5 As New TimeSpan(0, 0, 0, 0)
        Dim count6 As New TimeSpan(0, 0, 0, 0)
        Dim totalovertime As New TimeSpan(0, 0, 0, 0)

        'Try

        openconn()
        r = 0
        DataGridView1.Rows.Clear()
        cmd = New SqlCommand("select tdate,empname ,shiftstart ,shiftend,intime ,outtime  from Attendace where status ='P' and tdate >='" & (DateTimePicker1.Value.ToString("MM/dd/yyyy")) & "' and tdate <='" & (DateTimePicker2.Value.ToString("MM/dd/yyyy")) & "' and empname ='" & TextBox1.Text & "' and companyname='" & ComboBox3.Text & "'", cnn)
        'cmd = New SqlCommand("select tdate,empname ,shiftstart ,shiftend,intime ,outtime  from Attendace where status ='P'", cnn)
        drd = cmd.ExecuteReader
        While drd.Read
            newdate = drd("tdate")
            If IsDBNull(drd("shiftstart")) Then
                Starttime = "00:00:00"
            Else
                Starttime = drd("shiftstart")
            End If

            If IsDBNull(drd("shiftend")) Then
                endtime = "00:00:00"
            Else
                endtime = drd("shiftend")
            End If

            If IsDBNull(drd("intime")) Then
                intime = "00:00:00"
            Else
                intime = drd("intime")
            End If


            If IsDBNull(drd("outtime")) Then
                outtime = "00:00:00"
            Else
                outtime = drd("outtime")
            End If

      
            DataGridView1.Rows.Add()
            DataGridView1.Rows(r).Cells("date1").Value = newdate.ToString("dd/MM/yyyy")
            DataGridView1.Rows(r).Cells("empname1").Value = drd("empname")
            DataGridView1.Rows(r).Cells("shiftstart1").Value = Starttime.ToString("HH:mm:ss")
            DataGridView1.Rows(r).Cells("shiftend1").Value = endtime.ToString("HH:mm:ss")
            DataGridView1.Rows(r).Cells("intime1").Value = intime.ToString("HH:mm:ss")
            DataGridView1.Rows(r).Cells("outtime1").Value = outtime.ToString("HH:mm:ss")
            Try

                count4 = TimeSpan.Parse(DataGridView1.Rows(r).Cells("outtime1").Value) - TimeSpan.Parse(DataGridView1.Rows(r).Cells("intime1").Value)
                DataGridView1.Rows(r).Cells("totalhrs1").Value = count4

            Catch ex As Exception

            End Try




            count5 = TimeSpan.Parse(Label5.Text)

            count6 = count4 - count5
            DataGridView1.Rows(r).Cells("overtime2").Value = count6
            totalovertime = totalovertime + count6

            '  Try

            ' count5 = TimeSpan.Parse(Label5.Text)

            'If count4 > count5 Then
            '    DataGridView1.Rows(r).Cells("overtime2").Value = "Y"
            'Else
            '    DataGridView1.Rows(r).Cells("overtime2").Value = "N"
            'End If

            ' Catch ex As Exception

            ' End Try
            r = r + 1
        End While
        Label7.Text = totalovertime.ToString()
        count4 = TimeSpan.Zero
        count6 = TimeSpan.Zero
        drd.Close()
        closeconn()


        If DataGridView1.Rows.Count = 1 Then
            MsgBox("Record Empty")
        Else

        End If

        'Catch ex As Exception

        'End Try



    End Sub

    Private Sub DataGridView2_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub DataGridView2_Click(sender As Object, e As System.EventArgs) Handles DataGridView2.Click
        TextBox1.Text = DataGridView2.CurrentRow.Cells("empname").Value
        DataGridView2.Visible = False
    End Sub


    Private Sub visitMicrosoft2()
        System.Diagnostics.Process.Start("D:\OvertimeReport.xls")
    End Sub

    Public Function SaveTextToFile(ByVal strData As String, ByVal FullPath As String, Optional ByVal ErrInfo As String = "") As Boolean
        ' Dim Contents As String
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



    Private Sub CreatePage2()


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

        strFile = strFile & " <table style='width: 200px' border=0>"
        strFile = strFile & "  <tr> "
        strFile = strFile & " <td align='right'> "
        strFile = strFile & " </td>"
        strFile = strFile & "</tr>"
        strFile = strFile & "</table>"
        strFile = strFile & " <center><b> <Font Size=4> " & Label1.Text & "</Font></b></center>"
        strFile = strFile & "<center><b></b></center>"
        strFile = strFile & " <center><b> <Font Size=4>Name of Company:" & ComboBox3.Text & "</Font></b></center>"
        strFile = strFile & "<center><b></b></center>"
        strFile = strFile & " <center><b> <Font Size=4> Month:" & ComboBox1.Text & "</Font></b></center>"
        strFile = strFile & "<center><b></b></center>"
        strFile = strFile & " <center><b> <Font Size=4> Name of Employee:" & TextBox1.Text & "</Font></b></center>"
        strFile = strFile & "<center><b></b></center>"
        strFile = strFile & " <center><b> <Font Size=4> From Date:" & DateTimePicker1.Text & " To Date " & DateTimePicker2.Text & "</Font></b></center>"
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
                    Try

                     strFile = strFile & " " & R.Cells(col2.Index).Value.ToString() & " "
                    Catch ex As Exception

                    End Try

                    strFile = strFile & "  </center> </td>"
                End If
            Next

            strFile = strFile & "  <tr>"

        Next

        strFile = strFile & "</table>"
        strFile = strFile & "<table>"
        strFile = strFile & "  <tr> "
        strFile = strFile & " <td align='right'> Total OverTime:" & Label7.Text & ""
        strFile = strFile & " </td>"
        strFile = strFile & "</tr>"
        strFile = strFile & "</table>"
        strFile = strFile & "  </form>"
        strFile = strFile & "  </body></html>"
        SaveTextToFile(strFile, "D:\" & HTMLFileName & "OvertimeReport.xls")
    End Sub

    Private Sub DataGridView2_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles DataGridView2.KeyDown
        If TextBox1.Text = "" Then
            DataGridView2.Visible = False
        End If
        If DataGridView2.Rows.Count > 0 Then
            If e.KeyCode = 38 And DataGridView2.CurrentRow.Index = 0 Then
                If DataGridView2.CurrentCell.ColumnIndex = 1 Then
                    DateTimePicker1.Focus()
                Else
                    'TextBox4.Focus()
                End If
            End If
            If (e.KeyCode = Keys.Enter) Then
                DateTimePicker1.Focus()
                e.Handled = True
            End If
            If e.KeyCode = 13 Then
                TextBox1.Text = DataGridView2.CurrentRow.Cells("empname").Value
                'TextBox4.Text = DataGridView2.CurrentRow.Cells(1).Value
                DataGridView2.Visible = False
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown
        If TextBox1.Text = "" Then
            DataGridView2.Visible = False
        End If

        If e.KeyCode = Keys.Enter Then
            DataGridView2.Focus()

        End If


        If DataGridView2.Rows.Count <> 0 Then
            If e.KeyCode = 40 Then
                DataGridView2.Visible = True
                DataGridView2.Height = 150
                DataGridView2.Width = 250
                DataGridView2.Focus()
            End If
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text = "" Then
            DataGridView2.Visible = False
        End If
        If TextBox1.Text <> "" Then
            DataGridView2.Visible = True
            DataGridView2.Height = 150
            DataGridView2.Width = 250
            Dim start, length As Integer
            Dim str1 As String
            For start = 0 To DataGridView2.Rows.Count - 1
                str1 = StrConv(TextBox1.Text, VbStrConv.Uppercase)
                length = Len(str1)
                If str1 = Mid(DataGridView2.Rows(start).Cells("empname").Value, 1, length) Then
                    DataGridView2.CurrentCell = DataGridView2.Item(0, start)
                    DataGridView2.FirstDisplayedScrollingRowIndex = start
                    Exit Sub
                End If
            Next
        End If
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        refress()
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        datedchage()
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        CreatePage2()
        visitMicrosoft2()
    End Sub

    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        CreatePage2()

        Dim objProcess As New System.Diagnostics.ProcessStartInfo
        With objProcess
            .FileName = "d:\lateCommingNameWise.xls"
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


        System.IO.File.Delete("d:\OvertimeReport.xls")
    End Sub
End Class