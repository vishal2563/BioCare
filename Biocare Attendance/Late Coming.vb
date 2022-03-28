Option Explicit On
Imports System
Imports System.IO
Imports System.Text
Imports System.Data.SqlClient
Imports System.Data
Imports System.Configuration
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices.Marshal

Public Class Late_Coming
    Dim drd As SqlDataReader
    Dim cmd, cmd1, cmd2 As New SqlCommand
    Dim drd2 As SqlDataAdapter
    Dim r As Integer = 0
    Dim sqlstmt As String = ""
    Public dr
    Dim ENo As String = ""
    Private Property DataSet As DataTable
    Dim abc As Date
    Dim str As String
    Dim shec As String
    Dim i As Integer
    Dim j As Integer
    'Option Explicit On
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        If ComboBox2.Text = "Late by 15 Min" And ComboBox3.Text <> "Select" Then
            refressfiveteenmin()
            showlab()
            cal()
            If Val(Label7.Text) > 0 Or Val(Label8.Text) > 0 Or Val(Label11.Text) > 0 Then
                Label12.Visible = True
            Else
                Label12.Visible = False
            End If
        ElseIf ComboBox2.Text = "Late by 30 Min" And ComboBox3.Text <> "Select" Then
            refress()
            showlab()
            cal()
            If Val(Label7.Text) > 0 Or Val(Label8.Text) > 0 Or Val(Label11.Text) > 0 Then
                Label12.Visible = True
            Else
                Label12.Visible = False
            End If
        Else
            MsgBox("Select Report")
        End If 
    End Sub
    Private Sub showlab()
        'Label12.Visible = True
        Label11.Visible = True
        Label8.Visible = True
        Label7.Visible = True
    End Sub
    Private Sub refress()
        Try
            Dim bal As Integer
            Dim Count As Integer = 0
            Dim countmin As Integer = 0
            Dim countmaax As Integer = 0
            Dim Starttime As New DateTime
            Dim endtime As New DateTime
            Dim countzero As Integer = 0
            ' Dim count As Integer
            Dim empcount As Integer = 0
            openconn()
            r = 0
            DataGridView1.Rows.Clear()
            'cmd = New SqlCommand("select  empname ,shiftstart,intime,latearr ,tdate from Attendace where latearr <30 And  tdate >='" & (DateTimePicker1.Value.ToString("MM/dd/yyyy")) & "' and tdate <='" & (DateTimePicker2.Value.ToString("MM/dd/yyyy")) & "'", cnn)
            cmd = New SqlCommand("select latearr,shiftstart ,intime,empname,tdate  from Attendace where  status ='P' and tdate >='" & (DateTimePicker1.Value.ToString("MM/dd/yyyy")) & "' and tdate <='" & (DateTimePicker2.Value.ToString("MM/dd/yyyy")) & "' and companyname='" & ComboBox3.Text & "'", cnn)
            drd = cmd.ExecuteReader
            While drd.Read
                'Try
                Starttime = drd("shiftstart")
                If IsDBNull(drd("intime")) Then
                    endtime = "00:00:00"
                Else
                    endtime = drd("intime")
                End If
                DataGridView1.Rows.Add()
                DataGridView1.Rows(r).Cells(0).Value = drd("empname")
                Label10.Text = DataGridView1.Rows.Count - 1
                Label4.Text = DataGridView1.Rows(r).Cells(3).Value
                DataGridView1.Rows(r).Cells(1).Value = Starttime.ToString("HH:mm:ss")

                DataGridView1.Rows(r).Cells(2).Value = endtime.ToString("HH:mm:ss")
                '  End If
                Try
                    DataGridView1.Rows(r).Cells(6).Value = DateDiff(DateInterval.Minute, DataGridView1.Rows(r).Cells(1).Value, DataGridView1.Rows(r).Cells(2).Value)
                Catch ex As Exception
                End Try
                If DataGridView1.Rows(r).Cells(6).Value > 0 And DataGridView1.Rows(r).Cells(6).Value < 30 Then
                    DataGridView1.Rows(r).Cells(3).Value = DataGridView1.Rows(r).Cells(6).Value
                    Count = Count + 1
                Else
                    DataGridView1.Rows(r).Cells(3).Value = 0
                End If
                If DataGridView1.Rows(r).Cells(6).Value > 30 Then
                    DataGridView1.Rows(r).Cells(4).Value = DataGridView1.Rows(r).Cells(6).Value
                    countmin = countmin + 1

                ElseIf DataGridView1.Rows(r).Cells(3).Value = 0 Then
                    DataGridView1.Rows(r).Cells(4).Value = 0
                    countzero = countzero + 1
                End If
                ' DataGridView1.Rows(r).Cells(3).Value = drd("latearr")
                'DataGridView1.Rows(r).Cells(4).Value = "0"
                DataGridView1.Rows(r).Cells(5).Value = drd("tdate")

                r = r + 1
                'Catch ex As Exception

                'End Try
            End While
            Label4.Text = Count
            Label5.Text = countzero
            Label6.Text = countmin
            ' Label9.Text = countmaax
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
        '  Try
        Dim bal As Integer
        Dim Count As Integer = 0
        Dim countmin As Integer = 0
        Dim countmaax As Integer = 0
        Dim Starttime As New DateTime
        Dim endtime As New DateTime
        Dim countzero As Integer = 0
        ' Dim count As Integer
        Dim empcount As Integer = 0
        openconn()
        r = 0
        DataGridView2.Rows.Clear()
        'cmd = New SqlCommand("select  empname ,shiftstart,intime,latearr ,tdate from Attendace where latearr <30 And  tdate >='" & (DateTimePicker1.Value.ToString("MM/dd/yyyy")) & "' and tdate <='" & (DateTimePicker2.Value.ToString("MM/dd/yyyy")) & "'", cnn)
        cmd = New SqlCommand("select latearr,shiftstart ,intime,empname,tdate  from Attendace where  status ='P' and tdate >='" & (DateTimePicker1.Value.ToString("MM/dd/yyyy")) & "' and tdate <='" & (DateTimePicker2.Value.ToString("MM/dd/yyyy")) & "' and companyname='" & ComboBox3.Text & "'", cnn)
        drd = cmd.ExecuteReader
        While drd.Read
            'Try
            Starttime = drd("shiftstart")
            If IsDBNull(drd("intime")) Then
                endtime = "00:00:00"
            Else
                endtime = drd("intime")
            End If

            DataGridView2.Rows.Add()

            DataGridView2.Rows(r).Cells(0).Value = drd("empname")
            Label10.Text = DataGridView2.Rows.Count - 1
            'Label4.Text = DataGridView1.Rows(r).Cells(3).Value
            DataGridView2.Rows(r).Cells(1).Value = Starttime.ToString("HH:mm:ss")

            DataGridView2.Rows(r).Cells(2).Value = endtime.ToString("HH:mm:ss")
            '  End If

            '   Try
            DataGridView2.Rows(r).Cells(6).Value = DateDiff(DateInterval.Minute, DataGridView2.Rows(r).Cells(1).Value, DataGridView2.Rows(r).Cells(2).Value)
            'Catch ex As Exception

            'End Try

            If DataGridView2.Rows(r).Cells(6).Value > 0 And DataGridView2.Rows(r).Cells(6).Value < 15 Then
                DataGridView2.Rows(r).Cells(3).Value = DataGridView2.Rows(r).Cells(6).Value

                Count = Count + 1


            Else
                DataGridView2.Rows(r).Cells(3).Value = 0

            End If
            If DataGridView2.Rows(r).Cells(6).Value > 15 Then
                DataGridView2.Rows(r).Cells(4).Value = DataGridView2.Rows(r).Cells(6).Value
                countmin = countmin + 1

            ElseIf DataGridView2.Rows(r).Cells(3).Value = 0 Then
                DataGridView2.Rows(r).Cells(4).Value = 0
                countzero = countzero + 1
            End If
            ' DataGridView1.Rows(r).Cells(3).Value = drd("latearr")
            'DataGridView1.Rows(r).Cells(4).Value = "0"
            DataGridView2.Rows(r).Cells(5).Value = drd("tdate")

            r = r + 1
            'Catch ex As Exception

            'End Try
        End While
        Label4.Text = Count
        Label5.Text = countzero
        Label6.Text = countmin
        ' Label9.Text = countmaax
        drd.Close()
        closeconn()

        'Catch ex As Exception

        'End Try
        If DataGridView1.Rows.Count = 1 Then
            MsgBox("Record Empty")
        Else

        End If

    End Sub

    Private Sub cal()
        Try
            Dim a, b, c, d, e, f, g, h, i As Decimal
            a = Label10.Text
            b = Label4.Text
            c = b / a * 100
            Label7.Text = Val(c)
            Label7.Text = Math.Round(Val(Label7.Text), 2)
            d = Label6.Text
            e = d / a * 100
            Label8.Text = Val(e)
            Label8.Text = Math.Round(Val(Label8.Text), 2)
            f = Label5.Text
            'g = Label10.Text
            ' i = f + g
            h = f / a * 100
            Label11.Text = Val(h)
            Label11.Text = Math.Round(Val(Label11.Text), 2)
            'Label12.Text = Val(Label7.Text) + Val(Label8.Text) + Val(Label11.Text)
        Catch ex As Exception

        End Try
    End Sub



    Private Sub Late_Coming_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.Top = 0
        'Me.WindowState = 1
        'Me.WindowState = FormWindowState.Maximized
        Label12.Visible = False
        'refress()
        ' refrs()
        ' Me.ReportViewer1.RefreshReport()
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
    Private Sub visitMicrosoft2()
        System.Diagnostics.Process.Start("D:\LateCommingReport.xls")
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
        strFile = strFile & " <center><b> <Font Size=4>  Late Comming Report </Font></b></center>"
        strFile = strFile & "<center><b></b></center>"
        strFile = strFile & " <center><b> <Font Size=4> Company Name:" & ComboBox3.Text & "</Font></b></center>"
        strFile = strFile & "<center><b></b></center>"
        strFile = strFile & " <center><b> <Font Size=4> Report:" & ComboBox2.Text & "</Font></b></center>"
        strFile = strFile & "<center><b></b></center>"
        strFile = strFile & " <center><b> <Font Size=4> Month:" & ComboBox1.Text & "</Font></b></center>"
        strFile = strFile & "<center><b></b></center>"
        strFile = strFile & " <center><b> <Font Size=4> From Date:" & DateTimePicker1.Text & " To date:" & DateTimePicker2.Text & "</Font></b></center>"
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
        SaveTextToFile(strFile, "D:\" & HTMLFileName & "LateCommingReport.xls")
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
    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    


    Private Sub GroupBox1_Enter(sender As System.Object, e As System.EventArgs) Handles GroupBox1.Enter

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
    Private Sub DateTimePicker1_ValueChanged(sender As System.Object, e As System.EventArgs) Handles DateTimePicker1.ValueChanged
      

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        datedchage()

    End Sub

    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        Me.Close()
    End Sub
    Private bitmap As Bitmap
    Private Sub print()
        'Dim height As Integer = DataGridView1.Height
        'DataGridView1.Height = DataGridView1.RowCount * DataGridView1.RowTemplate.Height
        'bitmap = New Bitmap(Me.DataGridView1.Width, Me.DataGridView1.Height)
        'DataGridView1.DrawToBitmap(bitmap, New Rectangle(0, 0, Me.DataGridView1.Width, Me.DataGridView1.Height))
        'PrintPreviewDialog1.PrintPreviewControl.Zoom = 1
        'PrintPreviewDialog1.ShowDialog()
        'DataGridView1.Height = height

       

    End Sub
    Private Sub Button4_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '  Try

        If ComboBox2.Text = "Late by 15 Min" Then
            CreatePage3()

            Dim objProcess As New System.Diagnostics.ProcessStartInfo
            With objProcess
                .FileName = "d:\LateCommingReports.xls"
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


            System.IO.File.Delete("d:\LateCommingReports.xls")
            ' Catch ex As Exception

            ' End Try
        ElseIf ComboBox2.Text = "Late by 30 Min" Then
            CreatePage2()

            Dim objProcess As New System.Diagnostics.ProcessStartInfo
            With objProcess
                .FileName = "d:\LateCommingReport.xls"
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


            System.IO.File.Delete("d:\LateCommingReport.xls")
            ' Catch ex As Exception

            ' End Try
        End If


       
    End Sub

    Private Sub Label10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label10.Click

    End Sub

 
    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs)


    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub




    Private Sub ComboBox2_SelectedIndexChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.Text = "Late by 15 Min" Then
            DataGridView2.Visible = True
            DataGridView1.Visible = False
            Label14.Text = "%Age Late < 15 minutes"
            Label15.Text = "%Age Late > 15 minutes"

        ElseIf ComboBox2.Text = "Late by 30 Min" Then
            Label14.Text = "%Age Late < 30 minutes"
            Label15.Text = "%Age Late > 30 minutes"
            DataGridView1.Visible = True
            DataGridView2.Visible = False
        End If
    End Sub

    Private Sub visitMicrosoft3()
        System.Diagnostics.Process.Start("D:\LateCommingReports.xls")
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
        strFile = strFile & " <center><b> <Font Size=4>  Late Comming Report </Font></b></center>"
        strFile = strFile & "<center><b></b></center>"
        strFile = strFile & "<center><b></b></center>"
        strFile = strFile & " <center><b> <Font Size=4> Company Name:" & ComboBox3.Text & "</Font></b></center>"
        strFile = strFile & "<center><b></b></center>"
        strFile = strFile & " <center><b> <Font Size=4> Report:" & ComboBox2.Text & "</Font></b></center>"
        strFile = strFile & "<center><b></b></center>"
        strFile = strFile & " <center><b> <Font Size=4> Month:" & ComboBox1.Text & "</Font></b></center>"
        strFile = strFile & "<center><b></b></center>"
        strFile = strFile & " <center><b> <Font Size=4> From Date:" & DateTimePicker1.Text & " To date:" & DateTimePicker2.Text & "</Font></b></center>"
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
        SaveTextToFile(strFile, "D:\" & HTMLFileName & "LateCommingReports.xls")
    End Sub


End Class