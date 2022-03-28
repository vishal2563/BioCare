Imports System
Imports System.IO
Imports System.Text
Imports System.Data.SqlClient
Imports System.Data
Imports System.Configuration
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Net.Mail
Public Class Late_Report
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
    Private Sub Late_Report_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.WindowState = 1
        Me.WindowState = FormWindowState.Maximized
    End Sub

    Private Sub visitMicrosoft2()
        System.Diagnostics.Process.Start("E:\LateComming Report.xls")
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

        strFile = strFile & " <center><b> <Font Size=4> For the Day:" & DateTimePicker1.Text & "</Font></b></center>"
        strFile = strFile & "<center><b></b></center>"

        strFile = strFile & " <table style='width: 450px' border='1'>"
        strFile = strFile & "  <tr> "
        strFile = strFile & "     <td style='height: 20px; text-align: right'>"
        strFile = strFile & "   <center>  <b> Emp Name  </b> </center> </td>"

        strFile = strFile & "     <td style='height: 20px; text-align: right'>"
        strFile = strFile & "      <center>  <b>Schedule In Time </b> </center></td>"

        strFile = strFile & "     <td style='height: 20px; text-align: right'>"
        strFile = strFile & "       <center> <b>Actual Time </b> </center> </td>"

        strFile = strFile & "     <td style='height: 20px; text-align: right'>"
        strFile = strFile & "      <center>  <b> Late by <30 Minute  </b> </center></td>"
        strFile = strFile & "     <td style='height: 20px; text-align: right'>"
        strFile = strFile & "      <center>  <b> Late by >30 Minute  </b> </center></td>"
        strFile = strFile & "    <td style='height: 20px'>"
        strFile = strFile & "      <center> <b> Date </b> </center> </td>"


        strFile = strFile & "  </tr>"



        Dim b As DataGridViewRow
        For Each b In Me.DataGridView1.Rows
            '  If r.Visible = True Then

            strFile = strFile & " <tr>"
            strFile = strFile & "<td style='height: 21px' align='right'> <center>"
            strFile = strFile & " &nbsp; " & b.Cells(0).Value & " </center></td>"
            strFile = strFile & "<td style='height: 21px'> <center>"
            strFile = strFile & "           " & b.Cells(1).Value & " </center></td>"
            strFile = strFile & "       <td style='height: 21px' align='right'><center>"
            strFile = strFile & "          &nbsp; " & b.Cells(2).Value & " </center> </td>"

            strFile = strFile & "       <td style='height: 21px' align='right'> "
            strFile = strFile & "           " & (b.Cells(3).Value) & " </td>"
            strFile = strFile & "       <td style='height: 21px' align='right'> "
            strFile = strFile & "           " & (b.Cells(4).Value) & " </td>"
            strFile = strFile & "       <td style='height: 21px' align='right'> "
            strFile = strFile & "           " & (b.Cells(5).Value) & " </td>"

            strFile = strFile & "   </tr>"
            '  End If

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
        SaveTextToFile(strFile, "E:\" & HTMLFileName & "LateComming Report.xls")
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

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        refress()
        showlab()
        cal()
    End Sub
    Private Sub showlab()
        Label12.Visible = True
        Label11.Visible = True
        Label8.Visible = True
        Label7.Visible = True


    End Sub
    Private Sub refress()
        openconn()
        r = 0
        DataGridView1.Rows.Clear()

        'cmd = New SqlCommand("select  empname ,shiftstart,intime,latearr ,tdate from Attendace where latearr <30 And  tdate >='" & (DateTimePicker1.Value.ToString("MM/dd/yyyy")) & "' and tdate <='" & (DateTimePicker2.Value.ToString("MM/dd/yyyy")) & "'", cnn)
        cmd = New SqlCommand("select latearr ,SUBSTRING(shiftstart,12,17 )as starttime,SUBSTRING(intime,12,15) as intimes,empname,tdate  from Attendace where latearr <30 And  tdate >='" & (DateTimePicker1.Value.ToString("MM/dd/yyyy")) & "' and tdate <='" & (DateTimePicker2.Value.ToString("MM/dd/yyyy")) & "'", cnn)
        drd = cmd.ExecuteReader
        While drd.Read
            Try


                DataGridView1.Rows.Add()
                ' DataGridView1.Rows(r).Cells(0).Value = drd("empname")
                DataGridView1.Rows(r).Cells("actual").Value = drd("starttime")
                DataGridView1.Rows(r).Cells("actual").Value = drd("intimes")
                'DataGridView1.Rows(r).Cells(6).Value = drd("latearr")
                'DataGridView1.Rows(r).Cells(4).Value = "0"
                ' DataGridView1.Rows(r).Cells(5).Value = drd("tdate")
                r = r + 1
            Catch ex As Exception

            End Try
        End While
        drd.Close()


        'cmd = New SqlCommand("select empname ,shiftstart,intime,latearr,tdate from Attendace where latearr >30 And  tdate >='" & (DateTimePicker1.Value.ToString("MM/dd/yyyy")) & "' and tdate <='" & (DateTimePicker2.Value.ToString("MM/dd/yyyy")) & "'", cnn)
        'drd = cmd.ExecuteReader
        'While drd.Read
        '    str = (drd("shiftstart").ToString.Substring(0, 8))
        '    shec = (drd("intime").ToString.Substring(0, 8))
        '    DataGridView1.Rows.Add()

        '    DataGridView1.Rows(r).Cells(0).Value = drd("empname")
        '    DataGridView1.Rows(r).Cells(1).Value = str
        '    DataGridView1.Rows(r).Cells(2).Value = shec
        '    DataGridView1.Rows(r).Cells(3).Value = "0"
        '    DataGridView1.Rows(r).Cells(4).Value = drd("latearr")
        '    DataGridView1.Rows(r).Cells(5).Value = drd("tdate")
        '    r = r + 1

        'End While
        'drd.Close()
        cmd = New SqlCommand("select  COUNT (empname) as emp5  from Attendace ", cnn)
        drd = cmd.ExecuteReader
        While drd.Read
            Label4.Text = drd("emp5")
            Label4.Text = Math.Round(Val(Label4.Text), 2)
            r = r + 1
        End While
        drd.Close()
        cmd = New SqlCommand("select COUNT (latearr)as latearr7  from Attendace where latearr <30  and latearr > 0", cnn)
        drd = cmd.ExecuteReader
        While drd.Read
            Label5.Text = drd("latearr7")
            Label5.Text = Math.Round(Val(Label5.Text), 2)
            r = r + 1
        End While
        drd.Close()
        cmd = New SqlCommand("select  count(latearr) as late from Attendace where latearr >30 and latearr > 0", cnn)
        drd = cmd.ExecuteReader
        While drd.Read
            Label6.Text = drd("late")
            Label6.Text = Math.Round(Val(Label6.Text), 2)
            r = r + 1
        End While
        drd.Close()

        cmd = New SqlCommand("select  count(latearr) as latearr5 from Attendace where latearr=0 And latearr >30", cnn)
        drd = cmd.ExecuteReader
        While drd.Read
            Label9.Text = drd("latearr5")
            Label9.Text = Math.Round(Val(Label9.Text), 2)
            r = r + 1
        End While
        drd.Close()

        cmd = New SqlCommand("select  count(latearr) as larrr from Attendace where latearr=0 And latearr <30", cnn)
        drd = cmd.ExecuteReader
        While drd.Read
            Label10.Text = drd("larrr")
            Label10.Text = Math.Round(Val(Label10.Text), 2)
            r = r + 1
        End While
        drd.Close()

        closeconn()
    End Sub

    Private Sub cal()
        Try

            Dim a, b, c, d, e, f, g, h, i As Decimal
            a = Label4.Text
            b = Label5.Text
            c = b / a * 100
            Label7.Text = Val(c)
            Label7.Text = Math.Round(Val(Label7.Text), 2)
            d = Label6.Text
            e = d / a * 100
            Label8.Text = Val(e)
            Label8.Text = Math.Round(Val(Label8.Text), 2)
            f = Label9.Text
            g = Label10.Text
            i = f + g
            h = i / a * 100
            Label11.Text = Val(h)
            Label11.Text = Math.Round(Val(Label11.Text), 2)

            'Label12.Text = Val(Label7.Text) + Val(Label8.Text) + Val(Label11.Text)

        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        CreatePage2()
        visitMicrosoft2()
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        Me.Close()
    End Sub
End Class