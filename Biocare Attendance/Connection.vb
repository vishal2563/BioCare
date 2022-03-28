Imports System.Data.SqlClient
Imports System.Data
Imports System.IO

Module Connection
    Public db, cnn, cnn0, cnn1, cnn2, cnn3, cnn4, con As New SqlConnection
    Public loginflag As String = "N"
    Public firm As Integer = "0"
    Public LeaveCalc As String = " "
    Public FirmFormOpen As String = " "
    Public SqlServerName() As String
    Public SqlDataBase() As String
    Public SqlUserId() As String
    Public SqlPassword() As String
    Public FinYear As String = "2018-19"

    Public firstCopyTitle As String
    Public SecondCopyTitle As String
    Public ThirdCopyTitle As String
    Public FourthCopyTitle As String
    Public fifthCopyTitle As String

    Public editAccount As Boolean
    Public editItem As Boolean
    Public editBatch As Boolean
    Public editBOM As Boolean

    Public editSaleOrder As Boolean
    Public editSale As Boolean
    Public editSaleReturn As Boolean

    Public editPurchOrder As Boolean
    Public editPurch As Boolean
    Public editPurchReturn As Boolean

    Public editBankBook As Boolean
    Public editCashBook As Boolean
    Public editJournal As Boolean
    Public editOpStock As Boolean

    Public editSeries As Boolean

    Public print_Po_number As Integer = 0
    Public print_bill_number As Integer = 0
    Public print_CreditNote_number As Integer = 0
    Public print_DebitNote_number As Integer = 0

    Public Format_SaleBill As Integer = 0
    Public StateCode_ForBillPrint As String = ""

    Public publicreportpath As String
    Public publicperiodstart As Date
    Public publicperiodend As Date
    Public rpttype As String
    Public mac As String
    Public showpic As Boolean
    Public type As String = ""

    'Public str As String = readpath() & "User Id=sa;Password=zipzap"
    Public str As String = readpath() & "User Id=sa;Password=sa@123"

    Public Sub myglobal()
        Dim aa() As String = str.Split(";")
        SqlServerName = aa(0).Split("=")
        SqlDataBase = aa(1).Split("=")
        SqlUserId = aa(2).Split("=")
        SqlPassword = aa(3).Split("=")
    End Sub

    Public Function App_Path() As String
        Return System.AppDomain.CurrentDomain.BaseDirectory()
    End Function

    Public Function readpath() As String
        Dim path As String = App_Path() & "\datapath.txt"
        Dim d As String
        Try
            Dim sr As StreamReader = New StreamReader(path)
            d = sr.ReadLine()
            sr.Close()
            Return d
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Public Sub openconn0()
        cnn0 = New SqlConnection(str)
        cnn0.Open()
    End Sub

    Public Sub closeconn0()
        cnn0.Close()
    End Sub

    Public Sub openconn()
        'cnn = New SqlConnection("Data Source=" & SqlServerName & "; user id=" & SqlUserId & "; password=" & SqlPassword & ";Initial Catalog=" & SqlDataBase & ";")
        cnn = New SqlConnection(str)
        cnn.Open()
    End Sub

    Public Sub closeconn()
        cnn.Close()
    End Sub

    Public Sub openconn1()
        cnn1 = New SqlConnection(str)
        cnn1.Open()
    End Sub

    Public Sub closeconn1()
        cnn1.Close()
    End Sub

    Public Sub openconn2()
        cnn2 = New SqlConnection(str)
        cnn2.Open()
    End Sub

    Public Sub closeconn2()
        cnn2.Close()
    End Sub

    Public Sub openconn3()
        cnn3 = New SqlConnection(str)
        cnn3.Open()
    End Sub

    Public Sub closeconn3()
        cnn3.Close()
    End Sub

    Public Sub openconn4()
        cnn4 = New SqlConnection(str)
        cnn4.Open()
    End Sub

    Public Sub closeconn4()
        cnn4.Close()
    End Sub
   
    Public Function autonum(ByVal col As String, ByVal table As String) As Integer
        openconn()
        Dim i As Integer = 0
        Dim cmd As SqlCommand
        Dim drd As SqlDataReader
        cmd = New SqlCommand("select max(" & col & ") as temp from " & table & "", cnn)
        drd = cmd.ExecuteReader
        If drd.Read = True Then
            i = IIf(IsDBNull(drd("temp")), 0, drd("temp"))
            i = i + 1
        End If
        drd.Close()
        closeconn()
        Return i
    End Function
  


    Public Sub ValidarTeclaNumeros(ByVal sender As Object, ByVal e As KeyPressEventArgs)
        ''Dim Ct As TextBox
        'Ct = sender
        'If [Char].IsDigit(e.KeyChar) OrElse [Char].IsControl(e.KeyChar) Then
        '    'ok 
        '    'e.Handled = True
        'ElseIf [Char].IsPunctuation(e.KeyChar) Then
        '    If (Ct.Text.Contains(",") OrElse Ct.Text.Contains(".")) Then
        '        e.Handled = True
        '    End If


        'Else
        '    e.Handled = True
        'End If
    End Sub

    Public Sub exportme(ByVal datagridview1 As DataGridView, Optional title As String = "")
        Dim strfn2 As String = System.IO.Path.GetTempFileName
        Dim tempfileattb() As String = Split(strfn2, ".")
        Dim tempfilename As String = tempfileattb(0) & ".xls"
        Dim fs As New IO.StreamWriter(tempfilename, False)
        fs.WriteLine("<?xml version=""1.0""?>")
        fs.WriteLine("<?mso-application progid=""Excel.Sheet""?>")
        fs.WriteLine("<ss:Workbook xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet"">")
        fs.WriteLine("    <ss:Styles>")
        fs.WriteLine("        <ss:Style ss:ID=""1"">")
        fs.WriteLine("           <ss:Font ss:Bold=""1""/>")
        fs.WriteLine("        </ss:Style>")
        fs.WriteLine("    </ss:Styles>")
        fs.WriteLine("    <ss:Worksheet ss:Name=""Sheet1"">")
        fs.WriteLine("        <ss:Table>")
        For x As Integer = 0 To datagridview1.Columns.Count - 1
            fs.WriteLine("            <ss:Column ss:Width=""{0}""/>",
                         datagridview1.Columns.Item(x).Width)
        Next
        fs.WriteLine("            <ss:Row ss:StyleID=""1"">")
        For i As Integer = 0 To datagridview1.Columns.Count - 1
            fs.WriteLine("                <ss:Cell>")
            fs.WriteLine(String.Format(
                         "                   <ss:Data ss:Type=""String"">{0}</ss:Data>",
                                       datagridview1.Columns.Item(i).HeaderText))
            fs.WriteLine("                </ss:Cell>")
        Next
        fs.WriteLine("            </ss:Row>")
        For intRow As Integer = 0 To datagridview1.RowCount - 1
            fs.WriteLine(String.Format("            <ss:Row ss:Height =""{0}"">",
                                       datagridview1.Rows(intRow).Height))
            For intCol As Integer = 0 To datagridview1.Columns.Count - 1
                fs.WriteLine("                <ss:Cell>")

                Dim cellvalue As String
                cellvalue = ""
                If datagridview1.Rows(intRow).Cells(intCol).Value Is Nothing Then
                    cellvalue = ""
                ElseIf IsDBNull(datagridview1.Rows(intRow).Cells(intCol).Value) Then
                    cellvalue = ""
                ElseIf IsNumeric(datagridview1.Rows(intRow).Cells(intCol).Value) Then
                    cellvalue = Val(datagridview1.Rows(intRow).Cells(intCol).Value)
                Else
                    cellvalue = datagridview1.Rows(intRow).Cells(intCol).Value.ToString
                End If

                If cellvalue Is Nothing Then
                    cellvalue = ""
                End If

                If cellvalue.Contains("<") Then
                    cellvalue = cellvalue.Replace("<", "Less than ")
                End If

                If cellvalue.Contains(">") Then
                    cellvalue = cellvalue.Replace(">", "Greater than ")
                End If
                fs.WriteLine(String.Format(
                             "                   <ss:Data ss:Type=""String"">{0}</ss:Data>",
                                           cellvalue))
                fs.WriteLine("                </ss:Cell>")
            Next
            fs.WriteLine("            </ss:Row>")
        Next
        fs.WriteLine("        </ss:Table>")
        fs.WriteLine("    </ss:Worksheet>")
        fs.WriteLine("</ss:Workbook>")
        fs.Close()
        'MsgBox("File Exported to d:\exported.xls")
        ' MsgBox(tempfilename)
        Process.Start(tempfilename)
    End Sub



End Module
