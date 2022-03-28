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
Public Class EarlyNameWiseReport
    Dim cmd, cmd1, cmd2 As New SqlCommand
    Dim drd, drd1 As SqlDataReader
    Dim drd2 As SqlDataAdapter
    Dim r As Integer = 0
    Dim sqlstmt As String = ""
    Public dr
    Dim ENo As String = ""
    Dim str As String
    Dim shec As String

    Private Sub EarlyNameWiseReport_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.Top = 0
        'Me.WindowState = 1
        'Me.WindowState = FormWindowState.Maximized
        Label13.Visible = False
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
    Private Sub refress()

        Dim Starttime As New DateTime
        Dim endtime As New DateTime
        Dim newdate As Date
        Dim count As Decimal = 0
        Dim countmin As Decimal = 0
        Dim countzero As Decimal = 0
        Try
            openconn()
            r = 0
            DataGridView1.Rows.Clear()

            cmd = New SqlCommand("select  empname ,shiftend ,outtime ,tdate  from Attendace where  status ='P' and tdate >='" & (DateTimePicker1.Value.ToString("MM/dd/yyyy")) & "' and tdate <='" & (DateTimePicker2.Value.ToString("MM/dd/yyyy")) & "' And empname='" & (TextBox1.Text) & "' and companyname='" & ComboBox3.Text & "'", cnn)
            drd = cmd.ExecuteReader
            While drd.Read
                newdate = drd("tdate")
                Starttime = drd("shiftend")
                endtime = drd("outtime")
                DataGridView1.Rows.Add()
                DataGridView1.Rows(r).Cells(0).Value = newdate.ToString("dd/MM/yyyy")
                DataGridView1.Rows(r).Cells(1).Value = drd("empname")
                Label5.Text = DataGridView1.Rows.Count - 1
                DataGridView1.Rows(r).Cells(2).Value = Starttime.ToString("HH:mm:ss")
                DataGridView1.Rows(r).Cells(3).Value = endtime.ToString("HH:mm:ss")
                DataGridView1.Rows(r).Cells(6).Value = DateDiff(DateInterval.Minute, DataGridView1.Rows(r).Cells(3).Value, DataGridView1.Rows(r).Cells(2).Value)
                If DataGridView1.Rows(r).Cells(6).Value > 0 And DataGridView1.Rows(r).Cells(6).Value <= 30 Then
                    DataGridView1.Rows(r).Cells(4).Value = DataGridView1.Rows(r).Cells(6).Value
                    count = count + 1
                Else
                    DataGridView1.Rows(r).Cells(4).Value = 0
                    ' countzero = countzero + 1
                End If
                If DataGridView1.Rows(r).Cells(6).Value > 30 Then
                    DataGridView1.Rows(r).Cells(5).Value = DataGridView1.Rows(r).Cells(6).Value
                    countmin = countmin + 1
                ElseIf DataGridView1.Rows(r).Cells(4).Value = 0 Then
                    DataGridView1.Rows(r).Cells(5).Value = 0
                    countzero = countzero + 1
                Else
                    'DataGridView1.Rows(r).Cells(5).Value = 0

                End If
                r = r + 1

            End While
            Label6.Text = count
            Label9.Text = countmin
            Label10.Text = countzero
            drd.Close()
            closeconn()

        Catch ex As Exception

        End Try
        If DataGridView1.Rows.Count = 1 Then
            MsgBox("Record Empty")
        Else

        End If


    End Sub

    Private Sub refressfiveteenmin()

        Dim Starttime As New DateTime
        Dim endtime As New DateTime
        Dim newdate As Date
        Dim count As Decimal = 0
        Dim countmin As Decimal = 0
        Dim countzero As Decimal = 0
        Try
            openconn()
            r = 0
            DataGridView3.Rows.Clear()

            cmd = New SqlCommand("select  empname ,shiftend ,outtime ,tdate  from Attendace where  status ='P' and tdate >='" & (DateTimePicker1.Value.ToString("MM/dd/yyyy")) & "' and tdate <='" & (DateTimePicker2.Value.ToString("MM/dd/yyyy")) & "' And empname='" & (TextBox1.Text) & "' and companyname='" & ComboBox3.Text & "'", cnn)
            drd = cmd.ExecuteReader
            While drd.Read
                newdate = drd("tdate")
                Starttime = drd("shiftend")
                endtime = drd("outtime")
                DataGridView3.Rows.Add()
                DataGridView3.Rows(r).Cells(0).Value = newdate.ToString("dd/MM/yyyy")
                DataGridView3.Rows(r).Cells(1).Value = drd("empname")
                Label5.Text = DataGridView3.Rows.Count - 1
                DataGridView3.Rows(r).Cells(2).Value = Starttime.ToString("HH:mm:ss")
                DataGridView3.Rows(r).Cells(3).Value = endtime.ToString("HH:mm:ss")
                DataGridView3.Rows(r).Cells(6).Value = DateDiff(DateInterval.Minute, DataGridView3.Rows(r).Cells(3).Value, DataGridView3.Rows(r).Cells(2).Value)
                If DataGridView3.Rows(r).Cells(6).Value > 0 And DataGridView3.Rows(r).Cells(6).Value <= 15 Then
                    DataGridView3.Rows(r).Cells(4).Value = DataGridView3.Rows(r).Cells(6).Value
                    count = count + 1
                Else
                    DataGridView3.Rows(r).Cells(4).Value = 0
                    ' countzero = countzero + 1
                End If
                If DataGridView3.Rows(r).Cells(6).Value > 15 Then
                    DataGridView3.Rows(r).Cells(5).Value = DataGridView3.Rows(r).Cells(6).Value
                    countmin = countmin + 1
                ElseIf DataGridView3.Rows(r).Cells(4).Value = 0 Then
                    DataGridView3.Rows(r).Cells(5).Value = 0
                    countzero = countzero + 1
                Else
                    'DataGridView1.Rows(r).Cells(5).Value = 0

                End If
                r = r + 1

            End While
            Label6.Text = count
            Label9.Text = countmin
            Label10.Text = countzero
            drd.Close()
            closeconn()

        Catch ex As Exception

        End Try
        If DataGridView3.Rows.Count = 1 Then
            MsgBox("Record Empty")
        Else

        End If


    End Sub

    Private Sub cal()
        Try

            Dim a, b, c, d, e, f, g, h, i As Decimal
            a = Label5.Text
            b = Label6.Text
            c = b / a * 100
            Label7.Text = Val(c)
            Label7.Text = Math.Round(Val(Label7.Text), 2)
            d = Label9.Text
            e = d / a * 100
            Label8.Text = Val(e)
            Label8.Text = Math.Round(Val(Label8.Text), 2)
            g = Label10.Text
            ' i = f + g
            h = Val(g) / Val(a) * 100
            Label11.Text = Val(h)
            Label11.Text = Math.Round(Val(Label11.Text), 2)
            Label13.Text = Val(Label7.Text) + Val(Label8.Text) + Val(Label11.Text)
            Label13.Text = Math.Round(Val(Label13.Text), 0)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub visitMicrosoft2()
        System.Diagnostics.Process.Start("D:\Earlyname.xls")
    End Sub
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

        strFile = strFile & " <table style='width: 450px' border=0>"
        strFile = strFile & "  <tr> "
        strFile = strFile & " <td align='right'> "
        strFile = strFile & " </td>"
        strFile = strFile & "</tr>"
        strFile = strFile & "</table>"
        strFile = strFile & " <center><b> <Font Size=4> " & Label1.Text & "</Font></b></center>"
        strFile = strFile & "<center><b></b></center>"

        strFile = strFile & " <center><b> <Font Size=4>Company Name:" & ComboBox3.Text & "</Font></b></center>"
        strFile = strFile & "<center><b></b></center>"
        strFile = strFile & " <center><b> <Font Size=4>Report:" & ComboBox2.Text & "</Font></b></center>"
        strFile = strFile & "<center><b></b></center>"
        strFile = strFile & " <center><b> <Font Size=4>Month:" & ComboBox1.Text & "</Font></b></center>"
        strFile = strFile & "<center><b></b></center>"

        strFile = strFile & " <center><b> <Font Size=4> Name of Employee:" & TextBox1.Text & "</Font></b></center>"
        strFile = strFile & "<center><b></b></center>"
        strFile = strFile & " <center><b> <Font Size=4> From Date:" & DateTimePicker1.Text & " To Date:" & DateTimePicker2.Text & "</Font></b></center>"
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
        strFile = strFile & "<table>"
        strFile = strFile & "<tr >"
        strFile = strFile & "<td>&nbsp;</td>"
        strFile = strFile & "<td>&nbsp;</td>"
        strFile = strFile & "</tr>"
        strFile = strFile & "  <tr>"
        strFile = strFile & "<td>" & Label14.Text & ":</td>"
        strFile = strFile & "<td>" & Label7.Text & "</td>"
        strFile = strFile & "</tr>"
        strFile = strFile & "  <tr>"
        strFile = strFile & " <td>" & Label15.Text & ":</td>"
        strFile = strFile & " <td>" & Label8.Text & "</td>"
        strFile = strFile & "   </tr>"
        strFile = strFile & "  <tr>"
        strFile = strFile & " <td>" & Label16.Text & "</td>"
        strFile = strFile & " <td>" & Label11.Text & "</td>"
        strFile = strFile & "  </tr>"
        strFile = strFile & "  <tr>"
        strFile = strFile & " <td>&nbsp;</td>"
        strFile = strFile & " <td>" & Label12.Text & "</td>"
        strFile = strFile & "   </tr>"
        strFile = strFile & " </table>"

        strFile = strFile & "  </form>"
        strFile = strFile & "  </body></html>"
        SaveTextToFile(strFile, "D:\" & HTMLFileName & "Earlyname.xls")
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


    Private Sub visitMicrosoft3()
        System.Diagnostics.Process.Start("D:\EarlynameReport.xls")
    End Sub
    Private Sub CreatePage3()


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

        strFile = strFile & " <table style='width: 450px' border=0>"
        strFile = strFile & "  <tr> "
        strFile = strFile & " <td align='right'> "
        strFile = strFile & " </td>"
        strFile = strFile & "</tr>"
        strFile = strFile & "</table>"
        strFile = strFile & " <center><b> <Font Size=4> " & Label1.Text & "</Font></b></center>"
        strFile = strFile & "<center><b></b></center>"

        strFile = strFile & " <center><b> <Font Size=4>Company Name:" & ComboBox3.Text & "</Font></b></center>"
        strFile = strFile & "<center><b></b></center>"
        strFile = strFile & " <center><b> <Font Size=4>Report:" & ComboBox2.Text & "</Font></b></center>"
        strFile = strFile & "<center><b></b></center>"
        strFile = strFile & " <center><b> <Font Size=4>Month:" & ComboBox1.Text & "</Font></b></center>"
        strFile = strFile & "<center><b></b></center>"

        strFile = strFile & " <center><b> <Font Size=4> Name of Employee:" & TextBox1.Text & "</Font></b></center>"
        strFile = strFile & "<center><b></b></center>"
        strFile = strFile & " <center><b> <Font Size=4> From Date:" & DateTimePicker1.Text & " To Date:" & DateTimePicker2.Text & "</Font></b></center>"
        strFile = strFile & "<center><b></b></center>"
        strFile = strFile & " <table  border='1'>"
        strFile = strFile & "  <tr> "
        strFile = strFile & "  <td> "


        strFile = strFile & " <table style='width: 0' border='0'>"
        strFile = strFile & "  <tr> "
        For Each col In DataGridView3.Columns

            If col.Visible = True Then

                strFile = strFile & "      <td style=height: 20px; 'width: 0'>"
                strFile = strFile & "     <center> <b> " & col.HeaderText & " </b> </center></td>"

            End If

        Next

        strFile = strFile & "  </tr>"

        Dim col2 As DataGridViewColumn

        Dim R As DataGridViewRow

        For Each R In Me.DataGridView3.Rows

            strFile = strFile & "  <tr>"
            For Each col2 In DataGridView3.Columns
                If col2.Visible = True Then

                    strFile = strFile & "       <td style='Height: 21px'> <center>"

                    strFile = strFile & " " & R.Cells(col2.Index).Value & " </center></td>"

                End If
            Next

            strFile = strFile & "  <tr>"

        Next
        strFile = strFile & "</table>"
        strFile = strFile & "<table>"
        strFile = strFile & "<tr >"
        strFile = strFile & "<td>&nbsp;</td>"
        strFile = strFile & "<td>&nbsp;</td>"
        strFile = strFile & "</tr>"
        strFile = strFile & "  <tr>"
        strFile = strFile & "<td>" & Label14.Text & ":</td>"
        strFile = strFile & "<td>" & Label7.Text & "</td>"
        strFile = strFile & "</tr>"
        strFile = strFile & "  <tr>"
        strFile = strFile & " <td>" & Label15.Text & ":</td>"
        strFile = strFile & " <td>" & Label8.Text & "</td>"
        strFile = strFile & "   </tr>"
        strFile = strFile & "  <tr>"
        strFile = strFile & " <td>" & Label16.Text & "</td>"
        strFile = strFile & " <td>" & Label11.Text & "</td>"
        strFile = strFile & "  </tr>"
        strFile = strFile & "  <tr>"
        strFile = strFile & " <td>&nbsp;</td>"
        strFile = strFile & " <td>" & Label12.Text & "</td>"
        strFile = strFile & "   </tr>"
        strFile = strFile & " </table>"

        strFile = strFile & "  </form>"
        strFile = strFile & "  </body></html>"
        SaveTextToFile(strFile, "D:\" & HTMLFileName & "EarlynameReport.xls")
    End Sub

  
    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click

        If ComboBox2.Text = "Late by 15 Min" Then
            CreatePage3()

            visitMicrosoft3()
        ElseIf ComboBox2.Text = "Late by 30 Min" Then
            CreatePage2()

            visitMicrosoft2()
        End If

       
        

    End Sub
   
    Private Sub showlab()
        Label12.Visible = True
        Label7.Visible = True
        Label8.Visible = True
        Label11.Visible = True
        ' Label5.Visible = True
        'Label13.Visible = True

    End Sub
    Private Sub mis()
        openconn()
        cmd = New SqlCommand("select status  from Attendace where empname  ='" & TextBox1.Text & "'", cnn)
        drd = cmd.ExecuteReader
        If drd.Read Then
            Label17.Text = drd("status")
        End If
        drd.Close()
        closeconn()

    End Sub
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        If ComboBox2.Text = "Late by 15 Min" And ComboBox3.Text <> "Select" Then
            'mis()
            'If Label17.Text = "A" Or Label17.Text = "MIS" Then
            '    MsgBox("This is Record A Or MIS")
            ' ElseIf Label17.Text = "P" Then
            refressfiveteenmin()
            showlab()
            cal()
            If Val(Label7.Text) > 0 Or Val(Label8.Text) > 0 Or Val(Label11.Text) > 0 Then
                Label13.Visible = True
            Else
                Label13.Visible = False
            End If
            ' End If
        ElseIf ComboBox2.Text = "Late by 30 Min" And ComboBox3.Text <> "Select" Then
            'mis()
            'If Label17.Text = "A" Or Label17.Text = "MIS" Then
            '    MsgBox("This is Record A Or MIS")
            ' ElseIf Label17.Text = "P" Then
            refress()
            showlab()
            cal()
            If Val(Label7.Text) > 0 Or Val(Label8.Text) > 0 Or Val(Label11.Text) > 0 Then
                Label13.Visible = True
            Else
                Label13.Visible = False
            End If
        Else
            MsgBox("Select Report")
            ' End If
        End If

        
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

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

    Private Sub DataGridView2_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub DataGridView2_Click(sender As Object, e As System.EventArgs) Handles DataGridView2.Click
        TextBox1.Text = DataGridView2.CurrentRow.Cells("empname").Value
        DataGridView2.Visible = False
    End Sub

    Private Sub DataGridView2_Enter(sender As Object, e As System.EventArgs) Handles DataGridView2.Enter
        
    End Sub

    Private Sub DataGridView2_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles DataGridView2.KeyDown
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
    Private Sub ComboBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        datedchage()
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

        If ComboBox2.Text = "Late by 15 Min" Then
            CreatePage3()

            Dim objProcess As New System.Diagnostics.ProcessStartInfo
            With objProcess
                .FileName = "d:\EarlynameReport.xls"
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


            System.IO.File.Delete("d:\EarlynameReport.xls")
            ' Catch ex As Exception
            ' Close Microsoft Word
        ElseIf ComboBox2.Text = "Late by 30 Min" Then
            CreatePage2()

            Dim objProcess As New System.Diagnostics.ProcessStartInfo
            With objProcess
                .FileName = "d:\Earlyname.xls"
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


            System.IO.File.Delete("d:\Earlyname.xls")
            ' Catch ex As Exception
            ' Close Microsoft Word
        End If

       
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.Text = "Late by 15 Min" Then
            DataGridView3.Visible = True
            DataGridView1.Visible = False
            Label14.Text = " %Age Early < 15 minutes"
            Label15.Text = " %Age Early > 15 minutes"

        ElseIf ComboBox2.Text = "Late by 30 Min" Then
            DataGridView1.Visible = True
            DataGridView3.Visible = False
            Label14.Text = " %Age Early < 30 minutes"
            Label15.Text = " %Age Early > 30 minutes"
        End If
    End Sub

    Private Sub DataGridView3_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView3.CellContentClick

    End Sub

    Private Sub Panel1_Paint(sender As System.Object, e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint

    End Sub
End Class