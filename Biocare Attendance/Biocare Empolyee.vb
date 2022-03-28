Imports System
Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Data.SqlClient
Imports System.Data
Imports System.Configuration
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data.OleDb
Public Class Biocare_Empolyee
    Dim cmd, cmd1, cmd2 As New SqlCommand
    Dim drd, drd1 As SqlDataReader
    Dim drd2 As SqlDataAdapter
    Dim r As Integer = 0
    Dim sqlstmt As String = ""
    Public dr
    Dim ENo As String = ""
    Private Property DataSet As DataTable
    Private Sub Biocare_Empolyee_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        ' Biocare.PictureBox2.Visible = True
    End Sub
    Private Sub Biocare_Empolyee_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        ' Biocare.PictureBox2.Visible = False
        '  Me.WindowState = 1
        ' Me.WindowState = FormWindowState.Maximized
        fetch()
        'Addrecord()
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
            MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & Label3.Text & ";Extended Properties=Excel 8.0;")
            MyCommand = New System.Data.OleDb.OleDbDataAdapter("select F1 as [SrNo],F4 as [dated] ,F5 as [Emp code],F9 as [Card No] ,F16 as [Emp Name],F30 as [Shift Start],F38 as [Shift End],F45 as [In Time],F50 as [Out Time],F56 as [Shif Hrs],F64 as [Work Hrs],F72 as [Ot Hhrs],F78 as [Earl yarr],F85 as [Lat Arr],F91 as [Late Deep],F98 as [Early Deep],F105 as [Status],F45 as [Tdate] from [Sheet1$]", MyConnection)
          

            MyCommand.TableMappings.Add("Table", "dstables")
            DtSet = New System.Data.DataSet
            MyCommand.Fill(DtSet)
            DataGridView1.DataSource = DtSet.Tables(0)
            ' MsgBox("hello
            '    Dim SS As String = DtSet.Tables(0).Rows.Item("F1").ToString()
            'MsgBox(SS)
            MyConnection.Close()
            ' End If
        Catch ex As Exception
        End Try
        '' EmptyRow()
    End Sub
    Private Sub readfile()
        Dim dlg As New OpenFileDialog
        dlg.ShowDialog()
        If dlg.ShowDialog = Windows.Forms.DialogResult.OK Then
            Dim fileName As String
            fileName = dlg.FileName
            Label2.Text = fileName
            'MsgBox(fileName)
        End If
        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
        Dim i As Integer
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        'Try
        xlWorkBook = xlApp.Workbooks.Open(Label2.Text)
        xlWorkSheet = xlWorkBook.Worksheets("Sheet1")
        'For i = 1 To 11
        '    xlWorkSheet.Cells(i).EntireRow.Delete()
        'Next
        xlWorkSheet.Range("AT2:AV2").Value = "10:07"
        xlWorkBook.Save()
         
        'Catch ex As Exception
        '    MessageBox2(ex.Message, Me.Page)
        'Finally
        '    If Not IsNothing(xlWorkBook) Then
        ' MsgBox("Read File")
        xlWorkBook.Close()
        '    End If
        '    xlApp.Quit()
        ' End Try
    End Sub




    Private Sub fetch()
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
            cmd = New SqlCommand("select MAX(tdate) as dated2 from Attendace ", cnn)
            drd = cmd.ExecuteReader
            newdate = drd("dated2")

            While drd.Read
                TextBox3.Text = newdate.ToString("dd/MM/yyyy")
            End While
            drd.Close()
            closeconn()
            ' ElseIf DataGridView1.Rows(r).Cells(3).Value = 0 Then
            'DataGridView1.Rows(r).Cells(4).Value = 0
        Catch ex As Exception
        End Try
    End Sub
    Private Sub EmptyRow()
        Dim dt As Date
        Dim tdt2 As String = ""
        Dim r1 As DataGridViewRow
        For Each r1 In DataGridView2.Rows
            Try
                If (r1.Cells(0).Value.ToString) = "" And (r1.Cells(1).Value.ToString) = "" Then
                    '    MsgBox(r1.Cells(1).Value)
                    r1.Visible = False
                    '   tdt2 = r1.Cells(1).Value
                End If

            'If (r1.Cells(6).Value.ToString) <> "" Then
            '    dt = r1.Cells(6).Value.ToString
            '    TextBox1.Text = dt.ToString("dd/MM/yyyy")

            'Else

            'End If
            Catch ex As Exception

            End Try
        Next
    End Sub
    Private Sub fetchdatagridrow()
        Dim r2 As DataGridViewRow
        Dim dt As Date
        For Each r2 In DataGridView1.Rows

            Dim r1 As Integer = Me.DataGridView2.Rows.Add()
            With DataGridView2

                .Item(0, r1).Value = r2.Cells(0).Value
                .Item(1, r1).Value = r2.Cells(1).Value
                .Item(2, r1).Value = r2.Cells(2).Value

                .Item(3, r1).Value = r2.Cells(3).Value
                .Item(4, r1).Value = r2.Cells(4).Value
                .Item(5, r1).Value = r2.Cells(5).Value
                .Item(6, r1).Value = r2.Cells(6).Value


                .Item(7, r1).Value = r2.Cells(7).Value
                If TextBox2.Text = "" Then
                    Try

                        If IsDBNull(.Item(7, r1).Value) Then

                        Else
                            dt = CDate(.Item(7, r1).Value)

                            TextBox2.Text = dt.ToString("MM/dd/yyyy")
                        End If

                    Catch ex As Exception

                    End Try
                End If
                .Item(8, r1).Value = r2.Cells(8).Value
                .Item(9, r1).Value = r2.Cells(9).Value
                .Item(10, r1).Value = r2.Cells(10).Value
                .Item(11, r1).Value = r2.Cells(11).Value
                .Item(12, r1).Value = r2.Cells(12).Value
                .Item(13, r1).Value = r2.Cells(13).Value
                .Item(14, r1).Value = r2.Cells(14).Value
                .Item(15, r1).Value = r2.Cells(15).Value
                .Item(16, r1).Value = r2.Cells(16).Value
                .Item(17, r1).Value = r2.Cells(17).Value
                'If .Item(0, r1).Value = "" And .Item(1, r1).Value = "" Then
                '    .Rows(r1).Visible = False
                'End If

            End With

        Next

        'EmptyRow()

    End Sub
    Private Sub FetchDate()
        Dim tdt2 As String
        Dim tdate4 As Date
        Dim r1 As DataGridViewRow
        For Each r1 In DataGridView2.Rows
            If r1.Visible = True Then
                Try

                    If r1.Cells("dated1").Value.ToString <> "" Then
                        tdate4 = CDate(r1.Cells("dated1").Value)

                        TextBox1.Text = tdate4.ToString("MM/dd/yyyy")
                        r1.Cells("tdate1").Value = TextBox1.Text
                    Else
                        r1.Cells("dated1").Value = TextBox1.Text
                        r1.Cells("tdate1").Value = TextBox1.Text


                    End If


                    ' If r1.Cells("status1").Value.ToString() = "A" Or r1.Cells("status1").Value.ToString() = "WO-I" Or r1.Cells("status1").Value.ToString() = "NA" Or r1.Cells("status1").Value.ToString() = "MIS" Then

                    If r1.Cells("status1").Value = "A" Or r1.Cells("status1").Value = "WO-I" Or r1.Cells("status1").Value = "NA" Or r1.Cells("status1").Value = "MIS" Then


                        r1.Cells("tdate1").Value = TextBox1.Text

                        If r1.Cells("tdate1").Value.ToString = "" Then
                            'tdate4 = TextBox2.Text.ToString("yyyy/MM/dd")

                            r1.Cells("tdate1").Value = TextBox2.Text

                        End If


                    End If


                    '  If r1.Cells("status1").Value.ToString() = "A" Or r1.Cells("status1").Value.ToString() = "WO-I" Or r1.Cells("status1").Value.ToString() = "NA" Or r1.Cells("status1").Value.ToString() = "MIS" Then
                    If r1.Cells("status1").Value = "A" Or r1.Cells("status1").Value = "WO-I" Or r1.Cells("status1").Value = "NA" Or r1.Cells("status1").Value = "MIS" Then


                        r1.Cells("dated1").Value = TextBox1.Text
                        r1.Cells("InTime1").Value = r1.Cells("dated1").Value


                    End If

                Catch ex As Exception

                End Try

            End If
        Next
    End Sub
     
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Addrecord()
        fetchdatagridrow()
        'FetchDate()
        FetchFirstRecords()
        EMPTRYROWON0LYDATE()
        FetchINTimeRecord()
        FetchOUTTimeRecord()
     
    End Sub
    Private Sub record2()
        'Try
        openconn()
        For r As Integer = 0 To DataGridView2.Rows.Count - 1
            Try


                If DataGridView2.Rows(r).Visible = True Then
                    If DataGridView2.Rows(r).Cells("SrNo1").Value.ToString <> "" Then
                        cmd = New SqlCommand("insert into Attendace(srno,empcode,cardno, empname,shifhrs,workhrs,othrs,earlyarr,latearr,latedep,earlydep,status,shiftstart,shiftend,intime,outtime,tdate) values (@srno,@empcode,@cardno,@empname,@shifhrs,@workhrs,@othrs,@earlyarr,@latearr,@latedep,@earlydep,@status,@shiftstart,@shiftend,@intime,@outtime,@tdate)", cnn)
                        cmd.Parameters.AddWithValue("@srno", DataGridView2.Rows(r).Cells("SrNo1").Value)
                        cmd.Parameters.AddWithValue("@empcode", DataGridView2.Rows(r).Cells("Empcode1").Value)
                        cmd.Parameters.AddWithValue("@cardno", DataGridView2.Rows(r).Cells("CardNo1").Value)
                        cmd.Parameters.AddWithValue("@empname", DataGridView2.Rows(r).Cells("EmpName1").Value)
                        cmd.Parameters.AddWithValue("@shiftstart", DataGridView2.Rows(r).Cells("ShiftStart1").Value)
                        cmd.Parameters.AddWithValue("@shiftend", DataGridView2.Rows(r).Cells("ShiftEnd1").Value)
                        cmd.Parameters.AddWithValue("@intime", DataGridView2.Rows(r).Cells("InTime1").Value)
                        cmd.Parameters.AddWithValue("@outtime", DataGridView2.Rows(r).Cells("OutTime1").Value)
                        cmd.Parameters.AddWithValue("@shifhrs", DataGridView2.Rows(r).Cells("ShifHrs1").Value)
                        cmd.Parameters.AddWithValue("@workhrs", DataGridView2.Rows(r).Cells("WorkHrs1").Value)
                        cmd.Parameters.AddWithValue("@othrs", DataGridView2.Rows(r).Cells("OtHhrs1").Value)
                        cmd.Parameters.AddWithValue("@earlyarr", DataGridView2.Rows(r).Cells("Earlyarr1").Value)
                        cmd.Parameters.AddWithValue("@latearr", DataGridView2.Rows(r).Cells("LatArr1").Value)
                        cmd.Parameters.AddWithValue("@latedep", DataGridView2.Rows(r).Cells("LateDeep1").Value)
                        cmd.Parameters.AddWithValue("@earlydep", DataGridView2.Rows(r).Cells("EarlyDeep1").Value)
                        cmd.Parameters.AddWithValue("@status", DataGridView2.Rows(r).Cells("Status1").Value)
                        ' If IsDBNull(DataGridView1.Rows(r).Cells(6).Value) Then
                        cmd.Parameters.AddWithValue("@tdate", DataGridView2.Rows(r).Cells("Tdate1").Value)
                        'Else
                        '     cmd.Parameters.AddWithValue("@tdate", DataGridView1.Rows(r).Cells(6).Value)
                        'End If

                        cmd.ExecuteNonQuery()
                    End If

                End If
            Catch ex As Exception

            End Try

        Next


        'Catch ex As Exception
        'End Try
        MsgBox("Record Updated")
        DataGridView2.Rows.Clear()
        closeconn()
    End Sub
   

    Private Sub record()
        Try

            openconn()
            For r As Integer = 0 To DataGridView1.Rows.Count - 1
                'Try


                If DataGridView1.Rows(r).Visible = True Then
                    If DataGridView1.Rows(r).Cells(0).Value.ToString <> "" Then
                        cmd = New SqlCommand("insert into Attendace(srno,tdate,empcode,cardno, empname,shifhrs,workhrs,othrs,earlyarr,latearr,latedep,earlydep,status,shiftstart,shiftend,intime,outtime,companyname) values (@srno,@tdate,@empcode,@cardno,@empname,@shifhrs,@workhrs,@othrs,@earlyarr,@latearr,@latedep,@earlydep,@status,@shiftstart,@shiftend,@intime,@outtime,@companyname)", cnn)
                        cmd.Parameters.AddWithValue("@srno", DataGridView1.Rows(r).Cells(0).Value)

                        cmd.Parameters.AddWithValue("@tdate", CDate(DataGridView1.Rows(r).Cells(1).Value))
                        cmd.Parameters.AddWithValue("@empcode", DataGridView1.Rows(r).Cells(2).Value)

                        cmd.Parameters.AddWithValue("@cardno", DataGridView1.Rows(r).Cells(3).Value)
                        cmd.Parameters.AddWithValue("@empname", DataGridView1.Rows(r).Cells(4).Value)
                        cmd.Parameters.AddWithValue("@shiftstart", DataGridView1.Rows(r).Cells(5).Value)
                        cmd.Parameters.AddWithValue("@shiftend", DataGridView1.Rows(r).Cells(6).Value)
                        cmd.Parameters.AddWithValue("@intime", DataGridView1.Rows(r).Cells(7).Value)
                        cmd.Parameters.AddWithValue("@outtime", DataGridView1.Rows(r).Cells(8).Value)
                        cmd.Parameters.AddWithValue("@shifhrs", DataGridView1.Rows(r).Cells(9).Value)
                        cmd.Parameters.AddWithValue("@workhrs", DataGridView1.Rows(r).Cells(10).Value)
                        cmd.Parameters.AddWithValue("@othrs", DataGridView1.Rows(r).Cells(11).Value)
                        cmd.Parameters.AddWithValue("@earlyarr", DataGridView1.Rows(r).Cells(12).Value)
                        cmd.Parameters.AddWithValue("@latearr", DataGridView1.Rows(r).Cells(13).Value)
                        cmd.Parameters.AddWithValue("@latedep", DataGridView1.Rows(r).Cells(14).Value)
                        cmd.Parameters.AddWithValue("@earlydep", DataGridView1.Rows(r).Cells(15).Value)
                        cmd.Parameters.AddWithValue("@status", DataGridView1.Rows(r).Cells(16).Value)
                        cmd.Parameters.AddWithValue("@companyname", Trim(ComboBox1.Text))
                        ' If IsDBNull(DataGridView1.Rows(r).Cells(6).Value) Then

                        'Else
                        '     cmd.Parameters.AddWithValue("@tdate", DataGridView1.Rows(r).Cells(6).Value)
                        'End If

                        cmd.ExecuteNonQuery()
                    End If

                End If
                'Catch ex As Exception

                'End Try

            Next



            MsgBox("Record Updated")
            DataGridView1.Rows.Clear()
            closeconn()
        Catch ex As Exception
        End Try

    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs)

    End Sub
    Private Sub Button2_Click_1(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        'If Label2.Text = Label3.Text Then
        If ComboBox1.Text = "Select" Then
            MsgBox("Select Company Name")
        Else
            record()
            MsgBox("Record Updated")
        End If
       
        '  DataGridView1.Rows.Clear()

        ' Else
        '  MsgBox("File Not Same")
        'End If

    End Sub
   
    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        Try

            Dim dlg As New OpenFileDialog

            dlg.ShowDialog()

            If dlg.ShowDialog = Windows.Forms.DialogResult.OK Then

                Dim fileName As String

                fileName = dlg.FileName

                Label2.Text = fileName

                'MsgBox(fileName)

            End If

            Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()

            Dim i As Integer

            Dim xlWorkBook As Excel.Workbook

            Dim xlWorkSheet As Excel.Worksheet
            'Try
            Dim A As String

            Dim COMP As String
            xlWorkBook = xlApp.Workbooks.Open(Label2.Text)

            xlWorkSheet = xlWorkBook.Worksheets("Sheet1")

            A = xlWorkSheet.Range("DB13").Value.ToString
            COMP = xlWorkSheet.Range("B5").Value.ToString
            Dim WORDAAR As String() = COMP.Split(":")
            TextBox6.Text = WORDAAR(1)
            TextBox6.Text = Trim(TextBox6.Text)
            If A <> "P" Then


                xlWorkSheet.Range("AT13").Value = "10:00"

                xlWorkSheet.Range("AY13").Value = "10:00"

            End If

            For i = 1 To 11

                xlWorkSheet.Cells(i).EntireRow.Delete()

            Next


            xlWorkBook.Save()
            'Catch ex As Exception
            '    MessageBox2(ex.Message, Me.Page)
            'Finally
            '    If Not IsNothing(xlWorkBook) Then
            MsgBox("Read File")

            xlWorkBook.Close()
            Button3.Visible = False
            compare()
            '    End If
            '    xlApp.Quit()
            ' End Try

        Catch ex As Exception

        End Try
    End Sub
    Private Sub FetchFirstRecords()

        Dim RR As DataGridViewRow

        For Each RR In DataGridView1.Rows

            If Not IsDBNull(RR.Cells(1).Value) Then

                TextBox5.Text = RR.Cells(1).Value
                 
                If TextBox4.Text = "" Then

                    TextBox4.Text = TextBox5.Text

                    Dim d As Date = CDate(TextBox4.Text)

                    d = d.AddDays(-1)

                    TextBox5.Text = d

                End If

                Exit Sub

            Else

            End If
        Next
    End Sub
    Private Sub compare()
        If TextBox6.Text = ComboBox1.Text Then
            Button1.Enabled = True
        End If
    End Sub
    Private Sub FetchINTimeRecord()

        Dim RR As DataGridViewRow
        For Each RR In DataGridView1.Rows
            If IsDBNull(RR.Cells(7).Value) Then
                RR.Cells(7).Value = "00:00:00"
                Exit Sub
            Else
                'RR.Cells(7).Value = RR.Cells(7).Value
            End If
        Next
    End Sub
    Private Sub FetchOUTTimeRecord()
        'Dim RR As DataGridViewRow
        'For Each RR In DataGridView1.Rows
        '    If IsDBNull(RR.Cells(8).Value) Then
        '        RR.Cells(8).Value = "00:00:00"
        '        '   Exit Sub
        '    Else
        '        'RR.Cells(8).Value = RR.Cells(8).Value
        '    End If
        'Next
    End Sub


    Private Sub EMPTRYROWON0LYDATE()
        ' Try
        Dim RR As DataGridViewRow
        For Each RR In DataGridView1.Rows
            If Not IsDBNull(RR.Cells(1).Value) Then
                TextBox5.Text = RR.Cells(1).Value
                If TextBox4.Text = "" Then
                    TextBox4.Text = TextBox5.Text
                End If
            Else

            End If

            If IsDBNull(RR.Cells(1).Value) Then

                RR.Cells(1).Value = TextBox5.Text

            End If


            If IsDBNull(RR.Cells(0).Value) Then
                RR.Visible = False
            End If


        Next

       
        ' Catch ex As Exception
        
         ' End Try


    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        'FetchDate()
        Me.Close()
        '  Biocare.PictureBox2.Visible = True
    End Sub

    Private Sub Panel1_Paint(sender As System.Object, e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint

    End Sub

    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        Me.Close()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Button3.Enabled = True
        'DataGridView1.Rows.Clear()

    End Sub
End Class