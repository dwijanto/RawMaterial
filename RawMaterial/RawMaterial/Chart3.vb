Imports Microsoft.Office.Interop
Imports System.IO
Imports DJLib
Imports DJLib.ExcelStuff
Imports DJLib.Dbtools


Public Class Chart3
    Private UseAllCurrency As Boolean = True
    Private _CrcyCode As String
    Dim dbtools1 As New Dbtools(myUserid, myPassword)

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        ' Add any initialization after the InitializeComponent() call.
        Dim DbTools1 As New DJLib.Dbtools(myUserid, myPassword)
        Dim sqlstr As String = "select '' as master union all (select sheetname as master from rmmaterialraw group by sheetname order by sheetname) union all Select 'CURRENCY';" & _
                               "(Select keyid,updateby,source,links,area,currency,latestupdate,unit,sheetname,rawmaterialname,rawmaterialname || ' (' || source || ' - ' || unit || ')' as namesource from rmmaterialraw order by sheetname,rawmaterialname )union all" & _
                               " select 'EUR',null,null,null,null,null,null,null,'CURRENCY','EUR','EUR' union all select 'USD',null,null,null,null,null,null,null,'CURRENCY','USD','USD' union all select 'CNY',null,null,null,null,null,null,null,'CURRENCY','CNY','CNY' union all ((select crcycode,null,null,null,null,null,null,null,'CURRENCY',crcycode,crcycode from rmcrcytx  group by crcycode  order by crcycode));" & _
                               " select 'EUR' as crcycode union all select 'USD' as crcycode union all select 'CNY' as crcycode union all (select crcycode from rmcrcytx group by crcycode  union all select 'EUR' as crcycode order by crcycode);"
        Dim message As String = vbEmpty
        Dim Dataset1 As New DataSet
        Dim Dataset2 As New DataSet
        Dim Dataset3 As New DataSet
        Dim Dataset4 As New DataSet
        Dim Dataset5 As New DataSet
        Dim Dataset6 As New DataSet
        'FormMenu.ToolStripStatusLabel1.Text = "Please Wait...Initialize Chart1"

        FormMenu.Refresh()
        DbTools1.getDataSet(sqlstr, Dataset1, message)
        Dataset2 = Dataset1.Copy
        Dataset3 = Dataset1.Copy
        Dataset4 = Dataset1.Copy
        Dataset5 = Dataset1.Copy
        Dataset6 = Dataset1.Copy

        RmCrcy1.SeriesSelection1.GroupBox1.Text = "Series 1"
        RmCrcy1.SeriesSelection1.DataSet = Dataset1
        RmCrcy1.SeriesSelection2.GroupBox1.Text = "Series 2"
        RmCrcy1.SeriesSelection2.DataSet = Dataset2
        RmCrcy1.SeriesSelection3.GroupBox1.Text = "Series 3"
        RmCrcy1.SeriesSelection3.DataSet = Dataset3
        RmCrcy1.SeriesSelection4.GroupBox1.Text = "Series 4"
        RmCrcy1.SeriesSelection4.DataSet = Dataset4
        RmCrcy1.SeriesSelection5.GroupBox1.Text = "Series 5"
        RmCrcy1.SeriesSelection5.DataSet = Dataset5
        RmCrcy1.SeriesSelection6.GroupBox1.Text = "Series 6"
        RmCrcy1.SeriesSelection6.DataSet = Dataset6
        EnabledCurrency(True)

    End Sub

    Private Sub MultiCurrenciesCharts_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim sqlstr As String = "select '' as master union all (select sheetname as master from rmmaterialraw group by sheetname order by sheetname) union all Select 'CURRENCY';(Select * from rmmaterialraw order by sheetname,rawmaterialname )union all" & _
                               "select null,null,null,null,null,null,null,null,'CURRENCY','EUR' union all select null,null,null,null,null,null,null,null,'CURRENCY','USD' union all select null,null,null,null,null,null,null,null,'CURRENCY','CNY' union all ((select null,null,null,null,null,null,null,null,'CURRENCY',crcycode from rmcrcytx  group by crcycode  order by crcycode));" & _
                               "select 'EUR' as crcycode union all select 'USD' as crcycode union all select 'CNY' as crcycode union all (select crcycode from rmcrcytx group by crcycode  union all select 'EUR' as crcycode order by crcycode);"
    End Sub

    Private Sub RmCrcy1_CurrencyChanged(ByRef Crcycode As String) Handles RmCrcy1.CurrencyChanged
        _CrcyCode = Crcycode
        If UseAllCurrency Then

            setRocomboboxtxt(Crcycode)

        End If
    End Sub


    Private Sub RmCrcy1_UseAllCurrency(ByRef checked As Boolean) Handles RmCrcy1.UseAllCurrency
        UseAllCurrency = checked
        EnabledCurrency(checked)
        If checked Then RmCrcy1_CurrencyChanged(_CrcyCode)
    End Sub

    Private Sub EnabledCurrency(ByRef checked As Boolean)
        RmCrcy1.SeriesSelection2.RoComboBox3.Enabled = Not checked
        RmCrcy1.SeriesSelection3.RoComboBox3.Enabled = Not checked
        RmCrcy1.SeriesSelection4.RoComboBox3.Enabled = Not checked
        RmCrcy1.SeriesSelection5.RoComboBox3.Enabled = Not checked
        RmCrcy1.SeriesSelection6.RoComboBox3.Enabled = Not checked
    End Sub
    Private Sub setRocomboboxtxt(ByRef crcycode As String)
        RmCrcy1.SeriesSelection2.RoComboBox3.Text = crcycode
        RmCrcy1.SeriesSelection3.RoComboBox3.Text = crcycode
        RmCrcy1.SeriesSelection4.RoComboBox3.Text = crcycode
        RmCrcy1.SeriesSelection5.RoComboBox3.Text = crcycode
        RmCrcy1.SeriesSelection6.RoComboBox3.Text = crcycode
    End Sub

    Private Sub RmCrcy1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RmCrcy1.Load

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Export to Excel
        Dim Result As Boolean = False
        Cursor.Current = Cursors.WaitCursor
        Dim Filename As String = "MyMultiSeries-" & Format(DateTime.Today, "yyyyMMdd") & ".xlsx"
        Dim source As String = ""
        'ask export location
        Dim DirectoryBrowser As FolderBrowserDialog = New FolderBrowserDialog
        DirectoryBrowser.Description = "Which directory do you want to use?"
        If (DirectoryBrowser.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Application.DoEvents()
            source = DirectoryBrowser.SelectedPath & "\" & Filename

            'Excel Variable
            Dim oXl As Excel.Application = Nothing
            Dim oWb As Excel.Workbook = Nothing
            Dim oSheet As Excel.Worksheet = Nothing
            Dim SheetName As String = vbEmpty
            Dim oRange As Excel.Range = Nothing
            Dim J As Integer = 0
            Dim K As Integer = 0
            Dim Col As Long = 0
            Dim Row As Long = 0

            'Need these variable to kill excel
            Dim aprocesses() As Process = Nothing '= Process.GetProcesses
            Dim aprocess As Process = Nothing


            Try
                'Create Object Excel 
                oXl = CType(CreateObject("Excel.Application"), Excel.Application)
                Application.DoEvents()
                oXl.Visible = True
                'get process pid
                aprocesses = Process.GetProcesses
                For i = 0 To aprocesses.GetUpperBound(0)
                    If aprocesses(i).MainWindowHandle.ToString = oXl.Hwnd.ToString Then
                        aprocess = aprocesses(i)
                        Exit For
                    End If
                    Application.DoEvents()
                Next
                oXl.Visible = False
                Application.DoEvents()
                oXl.DisplayAlerts = False
                oWb = oXl.Workbooks.Open(Application.StartupPath & "\templates\MultiSeriesTemplate.xltx")
                'Loop for chart
                oSheet = oWb.Worksheets(1)
                oSheet.Name = "Graph"
                oWb.Worksheets(1).select()
                Dim MemoryStream1 As MemoryStream
                MemoryStream1 = New MemoryStream
                Dim charting As Object = RmCrcy1.Chart1
                charting.SaveImage(MemoryStream1, System.Drawing.Imaging.ImageFormat.Bmp)
                Dim bmp As New Bitmap(MemoryStream1)
                Clipboard.SetDataObject(bmp)
                '   paste to excel
                oSheet.Paste()
                '   adjust location
                Dim myobj As Object = oSheet.Shapes.Item(oSheet.Shapes.Count)
                myobj.top = 10
                myobj.left = 10
                'end loop
                Application.DoEvents()

                'save excel
                'Label10.Text = "Raw Material Figures available  From " & Format(CDate(Chart1.Series(0).Points(0).AxisLabel), "dd-MMM-yyyy") & " to " & Format(CDate(Chart1.Series(0).Points(Chart1.Series(0).Points.Count - 1).AxisLabel), "dd-MMM-yyyy")
                'put other info

                oSheet.Cells(23, 1) = "Date selection : " & Format(RmCrcy1.DateTimePicker1.Value, "dd-MMM-yyyy") & " To " & Format(RmCrcy1.DateTimePicker2.Value, "dd-MMM-yyyy")

                Dim iterate As Integer = 1
                If RmCrcy1.SeriesSelection1.SeriesLegend <> "" Then
                    oSheet.Cells(24 + iterate, 1) = RmCrcy1.SeriesSelection1.SeriesLegend
                    oSheet.Cells(24 + iterate, 2) = RmCrcy1.SeriesSelection1.TextBox2.Text
                    oSheet.Cells(24 + iterate, 3) = RmCrcy1.SeriesSelection1.TextBox1.Text
                    oSheet.Cells(24 + iterate, 4) = RmCrcy1.SeriesSelection1.TextBox3.Text
                    oSheet.Cells(24 + iterate, 5) = RmCrcy1.SeriesSelection1.TextBox4.Text
                    iterate = iterate + 1
                End If
                If RmCrcy1.SeriesSelection2.SeriesLegend <> "" Then
                    oSheet.Cells(24 + iterate, 1) = RmCrcy1.SeriesSelection2.SeriesLegend
                    oSheet.Cells(24 + iterate, 2) = RmCrcy1.SeriesSelection2.TextBox2.Text
                    oSheet.Cells(24 + iterate, 3) = RmCrcy1.SeriesSelection2.TextBox1.Text
                    oSheet.Cells(24 + iterate, 4) = RmCrcy1.SeriesSelection2.TextBox3.Text
                    oSheet.Cells(24 + iterate, 5) = RmCrcy1.SeriesSelection2.TextBox4.Text
                    iterate = iterate + 1
                End If
                If RmCrcy1.SeriesSelection3.SeriesLegend <> "" Then
                    oSheet.Cells(24 + iterate, 1) = RmCrcy1.SeriesSelection3.SeriesLegend
                    oSheet.Cells(24 + iterate, 2) = RmCrcy1.SeriesSelection3.TextBox2.Text
                    oSheet.Cells(24 + iterate, 3) = RmCrcy1.SeriesSelection3.TextBox1.Text
                    oSheet.Cells(24 + iterate, 4) = RmCrcy1.SeriesSelection3.TextBox3.Text
                    oSheet.Cells(24 + iterate, 5) = RmCrcy1.SeriesSelection3.TextBox4.Text
                    iterate = iterate + 1
                End If
                If RmCrcy1.SeriesSelection4.SeriesLegend <> "" Then
                    oSheet.Cells(24 + iterate, 1) = RmCrcy1.SeriesSelection4.SeriesLegend
                    oSheet.Cells(24 + iterate, 2) = RmCrcy1.SeriesSelection4.TextBox2.Text
                    oSheet.Cells(24 + iterate, 3) = RmCrcy1.SeriesSelection4.TextBox1.Text
                    oSheet.Cells(24 + iterate, 4) = RmCrcy1.SeriesSelection4.TextBox3.Text
                    oSheet.Cells(24 + iterate, 5) = RmCrcy1.SeriesSelection4.TextBox4.Text
                    iterate = iterate + 1
                End If
                If RmCrcy1.SeriesSelection5.SeriesLegend <> "" Then
                    oSheet.Cells(24 + iterate, 1) = RmCrcy1.SeriesSelection5.SeriesLegend
                    oSheet.Cells(24 + iterate, 2) = RmCrcy1.SeriesSelection5.TextBox2.Text
                    oSheet.Cells(24 + iterate, 3) = RmCrcy1.SeriesSelection5.TextBox1.Text
                    oSheet.Cells(24 + iterate, 4) = RmCrcy1.SeriesSelection5.TextBox3.Text
                    oSheet.Cells(24 + iterate, 5) = RmCrcy1.SeriesSelection5.TextBox4.Text
                    iterate = iterate + 1
                End If
                If RmCrcy1.SeriesSelection6.SeriesLegend <> "" Then
                    oSheet.Cells(24 + iterate, 1) = RmCrcy1.SeriesSelection6.SeriesLegend
                    oSheet.Cells(24 + iterate, 2) = RmCrcy1.SeriesSelection6.TextBox2.Text
                    oSheet.Cells(24 + iterate, 3) = RmCrcy1.SeriesSelection6.TextBox1.Text
                    oSheet.Cells(24 + iterate, 4) = RmCrcy1.SeriesSelection6.TextBox3.Text
                    oSheet.Cells(24 + iterate, 5) = RmCrcy1.SeriesSelection6.TextBox4.Text
                    iterate = iterate + 1
                End If

                oSheet.Cells.EntireColumn.AutoFit()

                Dim StringBuilder1 As New System.Text.StringBuilder
                Dim mysqlstr As String
                Dim iSheet As Integer = 0
                For i = 0 To 5                   
                    mysqlstr = ""
                    Dim mytitle As String = ""
                    Select Case i
                        Case 0
                            mysqlstr = RmCrcy1.SeriesSelection1.Sqlstr
                            mytitle = RmCrcy1.SeriesSelection1.SeriesLegend
                        Case 1
                            mysqlstr = RmCrcy1.SeriesSelection2.Sqlstr
                            mytitle = RmCrcy1.SeriesSelection2.SeriesLegend
                        Case 2
                            mysqlstr = RmCrcy1.SeriesSelection3.Sqlstr
                            mytitle = RmCrcy1.SeriesSelection3.SeriesLegend
                        Case 3
                            mysqlstr = RmCrcy1.SeriesSelection4.Sqlstr
                            mytitle = RmCrcy1.SeriesSelection4.SeriesLegend
                        Case 4
                            mysqlstr = RmCrcy1.SeriesSelection5.Sqlstr
                            mytitle = RmCrcy1.SeriesSelection5.SeriesLegend
                        Case 5
                            mysqlstr = RmCrcy1.SeriesSelection6.Sqlstr
                            mytitle = RmCrcy1.SeriesSelection6.SeriesLegend
                    End Select
                    oSheet = oWb.Worksheets(2 + iSheet)
                    If mysqlstr <> "" Then
                        oWb.Worksheets(2 + iSheet).select()
                        If mytitle.Length > 30 Then
                            mytitle = mytitle.Substring(0, 28)
                        End If
                        oSheet.Name = iSheet + 1 & " " & ValidateSheetName(mytitle)
                        ExcelStuff.FillDataSource(oWb, 2 + iSheet, mysqlstr, dbtools1)
                        'oSheet.Columns("F:F").NumberFormat = "0.000"
                        oSheet.Cells.EntireColumn.AutoFit()
                        iSheet = iSheet + 1
                    End If
                Next

                oWb.Worksheets(1).select()
                Filename = ValidateFileName(DirectoryBrowser.SelectedPath, source)
                oWb.SaveAs(Filename)
                Application.DoEvents()
                Result = True
                'FormMenu.setBubbleMessage("Export To Excel", "Done")
            Catch ex As Exception
                MsgBox(ex.Message)

            Finally
                'clear excel from memory
                oXl.Quit()
                oXl.Visible = True
                releaseComObject(oSheet)
                releaseComObject(oWb)
                releaseComObject(oXl)
                GC.Collect()
                GC.WaitForPendingFinalizers()
                Try
                    If Not aprocess Is Nothing Then
                        aprocess.Kill()
                    End If
                Catch ex As Exception
                End Try
                Application.DoEvents()
                If Result Then
                    If MsgBox("File name: " & Filename & vbCr & vbCr & "Open the file?", vbYesNo, "Export To Excel") = DialogResult.Yes Then
                        Process.Start(Filename)
                    End If
                End If

                Cursor.Current = Cursors.Default

            End Try

        End If
    End Sub
End Class