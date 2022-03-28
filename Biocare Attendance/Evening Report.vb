Option Explicit On
Imports System
Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Web
Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices.Marshal
Public Class Evening_Report
    Dim cmd, cmd1, cmd2 As New SqlCommand
    Dim drd, drd1 As SqlDataReader
    Dim drd2 As SqlDataAdapter
    Dim r As Integer = 0
    Dim sqlstmt As String = ""
    Public dr
    Dim ENo As String = ""
    Dim str As String
    Dim shec As String
    Dim fso As New Scripting.FileSystemObject
    Dim fil1 As Scripting.File

    Private Sub Evening_Report_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.Top = 0
        'Me.WindowState = 1
        'Me.WindowState = FormWindowState.Maximized
        Label5.Visible = False
    End Sub
    Private Sub refress()
        Dim Starttime As New DateTime
        Dim endtime As New DateTime

        Dim newdate As Date
        Dim count As Integer = 0
        Dim countmin As Integer = 0
        Dim countzero As Integer = 0

        Try
            openconn()
            r = 0
            DataGridView1.Rows.Clear()
            cmd = New SqlCommand("select empname ,shiftend ,outtime ,earlydep,tdate,status  from Attendace where status ='P' and  tdate >='" & (DateTimePicker1.Value.ToString("MM/dd/yyyy")) & "' and tdate <='" & (DateTimePicker2.Value.ToString("MM/dd/yyyy")) & "' and companyname='" & ComboBox3.Text & "'", cnn)
            drd = cmd.ExecuteReader
            While drd.Read
                newdate = drd("tdate")
                Starttime = drd("shiftend")
                DataGridView1.Rows.Add()
                DataGridView1.Rows(r).Cells(0).Value = drd("empname")
                Label1.Text = DataGridView1.Rows.Count - 1
                DataGridView1.Rows(r).Cells(1).Value = Starttime.ToString("HH:mm:ss")
                If IsDBNull(drd("outtime")) Then
                    endtime = "00:00:00"
                Else
                    endtime = drd("outtime")
                End If
                DataGridView1.Rows(r).Cells(2).Value = endtime.ToString("HH:mm:ss")
                DataGridView1.Rows(r).Cells(6).Value = DateDiff(DateInterval.Minute, DataGridView1.Rows(r).Cells(2).Value, DataGridView1.Rows(r).Cells(1).Value)
                If DataGridView1.Rows(r).Cells(6).Value > 0 And DataGridView1.Rows(r).Cells(6).Value <= 30 And DataGridView1.Rows(r).Cells(2).Value <> "00:00:00" Then
                    DataGridView1.Rows(r).Cells(3).Value = DataGridView1.Rows(r).Cells(6).Value
                    count = count + 1
                ElseIf DataGridView1.Rows(r).Cells(4).Value = 0 Then
                    DataGridView1.Rows(r).Cells(3).Value = 0
                End If
                DataGridView1.Rows(r).Cells(5).Value = newdate.ToString("dd/MM/yyyy")
                ' MsgBox(drd("status"))

                DataGridView1.Rows(r).Cells(7).Value = drd("status")
                'If DataGridView1.Rows(r).Cells(7).Value = "MIS" Or DataGridView1.Rows(r).Cells(7).Value = "A" Then
                '    DataGridView1.Rows(r).Cells(6).Value = ""
                'End If
                If DataGridView1.Rows(r).Cells(6).Value > 30 Then
                    DataGridView1.Rows(r).Cells(4).Value = DataGridView1.Rows(r).Cells(6).Value
                    countmin = countmin + 1
                ElseIf DataGridView1.Rows(r).Cells(3).Value = 0 Then
                    DataGridView1.Rows(r).Cells(4).Value = 0
                    countzero = countzero + 1
                End If
                r = r + 1
            End While
            Label2.Text = count
            Label3.Text = countmin
            Label4.Text = countzero
            drd.Close()

            closeconn()

            ' ElseIf DataGridView1.Rows(r).Cells(3).Value = 0 Then
            'DataGridView1.Rows(r).Cells(4).Value = 0
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
        Dim count As Integer = 0
        Dim countmin As Integer = 0
        Dim countzero As Integer = 0

        Try
            openconn()
            r = 0
            DataGridView2.Rows.Clear()
            cmd = New SqlCommand("select empname ,shiftend ,outtime ,earlydep,tdate,status  from Attendace where status ='P' and  tdate >='" & (DateTimePicker1.Value.ToString("MM/dd/yyyy")) & "' and tdate <='" & (DateTimePicker2.Value.ToString("MM/dd/yyyy")) & "' and companyname='" & ComboBox3.Text & "' ", cnn)
            drd = cmd.ExecuteReader
            While drd.Read
                newdate = drd("tdate")
                Starttime = drd("shiftend")
                DataGridView2.Rows.Add()
                DataGridView2.Rows(r).Cells(0).Value = drd("empname")
                Label1.Text = DataGridView2.Rows.Count - 1
                DataGridView2.Rows(r).Cells(1).Value = Starttime.ToString("HH:mm:ss")
                If IsDBNull(drd("outtime")) Then
                    endtime = "00:00:00"
                Else
                    endtime = drd("outtime")
                End If
                DataGridView2.Rows(r).Cells(2).Value = endtime.ToString("HH:mm:ss")
                DataGridView2.Rows(r).Cells(6).Value = DateDiff(DateInterval.Minute, DataGridView2.Rows(r).Cells(2).Value, DataGridView2.Rows(r).Cells(1).Value)
                If DataGridView2.Rows(r).Cells(6).Value > 0 And DataGridView2.Rows(r).Cells(6).Value <= 15 And DataGridView2.Rows(r).Cells(2).Value <> "00:00:00" Then
                    DataGridView2.Rows(r).Cells(3).Value = DataGridView2.Rows(r).Cells(6).Value
                    count = count + 1
                ElseIf DataGridView2.Rows(r).Cells(4).Value = 0 Then
                    DataGridView2.Rows(r).Cells(3).Value = 0
                End If
                DataGridView2.Rows(r).Cells(5).Value = newdate.ToString("dd/MM/yyyy")
                ' MsgBox(drd("status"))

                DataGridView2.Rows(r).Cells(7).Value = drd("status")
                'If DataGridView1.Rows(r).Cells(7).Value = "MIS" Or DataGridView1.Rows(r).Cells(7).Value = "A" Then
                '    DataGridView1.Rows(r).Cells(6).Value = ""
                'End If
                If DataGridView2.Rows(r).Cells(6).Value > 15 Then
                    DataGridView2.Rows(r).Cells(4).Value = DataGridView2.Rows(r).Cells(6).Value
                    countmin = countmin + 1
                ElseIf DataGridView2.Rows(r).Cells(3).Value = 0 Then
                    DataGridView2.Rows(r).Cells(4).Value = 0
                    countzero = countzero + 1
                End If
                r = r + 1
            End While
            Label2.Text = count
            Label3.Text = countmin
            Label4.Text = countzero
            drd.Close()

            closeconn()

            ' ElseIf DataGridView1.Rows(r).Cells(3).Value = 0 Then
            'DataGridView1.Rows(r).Cells(4).Value = 0
        Catch ex As Exception

        End Try
        If DataGridView2.Rows.Count = 1 Then
            MsgBox("Record Empty")
        Else

        End If


    End Sub

    Private Sub cal()
        Try

            Dim a, b, c, d, e, f, g, j As Decimal
            a = Label1.Text
            b = Label2.Text
            c = b / a * 100
            Label6.Text = Val(c)
            Label6.Text = Math.Round(Val(Label6.Text), 2)
            d = Label3.Text
            e = d / a * 100
            Label7.Text = Val(e)
            Label7.Text = Math.Round(Val(Label7.Text), 2)
            f = Label4.Text
            j = f / a * 100

            Label8.Text = Val(j)
            Label8.Text = Math.Round(Val(Label8.Text), 1)
            '  Label5.Text = Val(Label6.Text) + Val(Label7.Text) + Val(Label8.Text)

        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Me.Close()

    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click

        If ComboBox2.Text = "Late by 15 Min" Then
            CreatePage3()
            visitMicrosoft3()
        ElseIf ComboBox2.Text = "Late by 30 Min" Then
            CreatePage2()
            visitMicrosoft2()
        End If

      
    End Sub
    Private Sub visitMicrosoft2()
        System.Diagnostics.Process.Start("D:\Evening Report.xls")
    End Sub
    Private Sub CreatePage2()


        Dim HTMLTitle As String = ""
        Dim HTMLText As String = ""
        Dim HTMLFileName As String = ""
        Dim strFile As String
        Dim str As Date
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
        strFile = strFile & " <center><b> <Font Size=4>  Early leaving Report </Font></b></center>"
        strFile = strFile & "<center><b></b></center>"
        strFile = strFile & " <center><b> <Font Size=4> Company Name:" & ComboBox3.Text & "</Font></b></center>"
        strFile = strFile & "<center><b></b></center>"
        strFile = strFile & " <center><b> <Font Size=4> Report:" & ComboBox2.Text & "</Font></b></center>"
        strFile = strFile & "<center><b></b></center>"
        strFile = strFile & " <center><b> <Font Size=4> Month:" & ComboBox1.Text & "</Font></b></center>"
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
        strFile = strFile & "<table border='1' width='100%'>"
        strFile = strFile & "<tr >"
        strFile = strFile & "<td>&nbsp;</td>"
        strFile = strFile & "<td>&nbsp;</td>"
        strFile = strFile & "</tr>"
        strFile = strFile & "  <tr>"
        strFile = strFile & "<td>" & Label14.Text & ":</td>"
        strFile = strFile & "<td>" & Label6.Text & "</td>"
        strFile = strFile & "</tr>"
        strFile = strFile & "  <tr>"
        strFile = strFile & " <td>" & Label15.Text & ":</td>"
        strFile = strFile & " <td>" & Label7.Text & "</td>"
        strFile = strFile & "   </tr>"
        strFile = strFile & "  <tr>"
        strFile = strFile & " <td>" & Label16.Text & "</td>"
        strFile = strFile & " <td>" & Label8.Text & "</td>"
        strFile = strFile & "  </tr>"
        strFile = strFile & "  <tr>"
        strFile = strFile & " <td>&nbsp;</td>"
        strFile = strFile & " <td>" & Label9.Text & "</td>"
        strFile = strFile & "   </tr>"
        strFile = strFile & " </table>"

        strFile = strFile & "  </form>"
        strFile = strFile & "  </body></html>"
        SaveTextToFile(strFile, "D:\" & HTMLFileName & "Evening Report.xls")
    End Sub

    Private Sub visitMicrosoft3()
        System.Diagnostics.Process.Start("D:\EveningReports.xls")
    End Sub
    Private Sub CreatePage3()


        Dim HTMLTitle As String = ""
        Dim HTMLText As String = ""
        Dim HTMLFileName As String = ""
        Dim strFile As String
        Dim str As Date
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
        strFile = strFile & " <center><b> <Font Size=4>  Early leaving Report </Font></b></center>"
        strFile = strFile & "<center><b></b></center>"
        strFile = strFile & " <center><b> <Font Size=4> Company Name:" & ComboBox3.Text & "</Font></b></center>"
        strFile = strFile & "<center><b></b></center>"
        strFile = strFile & " <center><b> <Font Size=4> Report:" & ComboBox2.Text & "</Font></b></center>"
        strFile = strFile & "<center><b></b></center>"
        strFile = strFile & " <center><b> <Font Size=4> Month:" & ComboBox1.Text & "</Font></b></center>"
        strFile = strFile & "<center><b></b></center>"
        strFile = strFile & " <center><b> <Font Size=4> From Date:" & DateTimePicker1.Text & " To Date:" & DateTimePicker2.Text & "</Font></b></center>"
        strFile = strFile & "<center><b></b></center>"

        strFile = strFile & " <table  border='1'>"
        strFile = strFile & "  <tr> "
        strFile = strFile & "  <td> "


        strFile = strFile & " <table style='width: 0' border='0'>"
        strFile = strFile & "  <tr> "
        For Each col In DataGridView2.Columns

            If col.Visible = True Then

                strFile = strFile & "      <td style=height: 20px; 'width: 0'>"
                strFile = strFile & "     <center> <b> " & col.HeaderText & " </b> </center></td>"

            End If

        Next

        strFile = strFile & "  </tr>"

        Dim col2 As DataGridViewColumn

        Dim R As DataGridViewRow

        For Each R In Me.DataGridView2.Rows

            strFile = strFile & "  <tr>"
            For Each col2 In DataGridView2.Columns
                If col2.Visible = True Then

                    strFile = strFile & "       <td style='Height: 21px'> <center>"

                    strFile = strFile & " " & R.Cells(col2.Index).Value & " </center></td>"

                End If
            Next

            strFile = strFile & "  <tr>"

        Next
        strFile = strFile & "</table>"
        strFile = strFile & "<table border='1' width='100%'>"
        strFile = strFile & "<tr >"
        strFile = strFile & "<td>&nbsp;</td>"
        strFile = strFile & "<td>&nbsp;</td>"
        strFile = strFile & "</tr>"
        strFile = strFile & "  <tr>"
        strFile = strFile & "<td>" & Label14.Text & ":</td>"
        strFile = strFile & "<td>" & Label6.Text & "</td>"
        strFile = strFile & "</tr>"
        strFile = strFile & "  <tr>"
        strFile = strFile & " <td>" & Label15.Text & ":</td>"
        strFile = strFile & " <td>" & Label7.Text & "</td>"
        strFile = strFile & "   </tr>"
        strFile = strFile & "  <tr>"
        strFile = strFile & " <td>" & Label16.Text & "</td>"
        strFile = strFile & " <td>" & Label8.Text & "</td>"
        strFile = strFile & "  </tr>"
        strFile = strFile & "  <tr>"
        strFile = strFile & " <td>&nbsp;</td>"
        strFile = strFile & " <td>" & Label9.Text & "</td>"
        strFile = strFile & "   </tr>"
        strFile = strFile & " </table>"

        strFile = strFile & "  </form>"
        strFile = strFile & "  </body></html>"
        SaveTextToFile(strFile, "D:\" & HTMLFileName & "EveningReports.xls")
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
    Private Sub showlab()

        Label7.Visible = True
        Label9.Visible = True
        Label6.Visible = True
        Label8.Visible = True




    End Sub
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

        If ComboBox2.Text = "Late by 15 Min" And ComboBox3.Text <> "Select" Then
            refressfiveteenmin()
            showlab()
            cal()
            If Val(Label6.Text) > 0 Or Val(Label7.Text) > 0 Or Val(Label8.Text) > 0 Then
                Label5.Visible = True
            Else
                Label5.Visible = False
            End If

        ElseIf ComboBox2.Text = "Late by 30 Min" And ComboBox3.Text <> "Select" Then
            refress()
            showlab()
            cal()
            If Val(Label6.Text) > 0 Or Val(Label7.Text) > 0 Or Val(Label8.Text) > 0 Then
                Label5.Visible = True
            Else
                Label5.Visible = False
            End If
        Else
            MsgBox("Select Report")
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

    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        Me.Close()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        If ComboBox2.Text = "Late by 15 Min" Then
            CreatePage3()

            Dim objProcess As New System.Diagnostics.ProcessStartInfo
            With objProcess
                .FileName = "d:\EveningReports.xls"
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


            System.IO.File.Delete("d:\EveningReports.xls")
            ' Close Microsoft Word
        ElseIf ComboBox2.Text = "Late by 30 Min" Then
            CreatePage2()

            Dim objProcess As New System.Diagnostics.ProcessStartInfo
            With objProcess
                .FileName = "d:\EveningReport.xls"
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


            System.IO.File.Delete("d:\EveningReport.xls")
            ' Close Microsoft Word
        End If


      

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.Text = "Late by 15 Min" Then
            DataGridView2.Visible = True
            DataGridView1.Visible = False
            Label14.Text = "%Age Early < 15 minutes"
            Label15.Text = " %Age Early > 15 minutes"

        ElseIf ComboBox2.Text = "Late by 30 Min" Then
            DataGridView1.Visible = True
            DataGridView2.Visible = False
            Label14.Text = "%Age Early < 30 minutes"
            Label15.Text = "%Age Early > 30 minutes"
        End If
    End Sub

    Private Sub ComboBox2_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.Validated
        'If ComboBox2.Text = "Late by 15 Min" Then
        '    DataGridView2.Visible = True
        '    DataGridView1.Visible = False
        'ElseIf ComboBox2.Text = "Late by 30 Min" Then
        '    DataGridView1.Visible = True
        '    DataGridView2.Visible = False
        'End If
    End Sub

    Private Sub GroupBox1_Enter(sender As System.Object, e As System.EventArgs) Handles GroupBox1.Enter

    End Sub
End Class