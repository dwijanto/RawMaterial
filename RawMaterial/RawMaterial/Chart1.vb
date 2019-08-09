Imports System.Data.Common
Imports DJLib
Imports DJLib.Dbtools
Imports DJLib.ExcelStuff
Imports System.IO

Imports Microsoft.Office.Interop

Public Class Chart1

    Public Sub New()
        Cursor.Current = Cursors.WaitCursor
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        Dim DbTools1 As New DJLib.Dbtools(myUserid, myPassword)
        Dim sqlstr As String = "select '' as sheetname union all (select sheetname from rmmaterialraw group by sheetname order by sheetname);" & _
            "Select keyid,updateby,source,links,area,currency,latestupdate,unit,sheetname,rawmaterialname,rawmaterialname || ' (' || source || ' - ' || unit || ')' as namesource from rmmaterialraw order by sheetname,rawmaterialname;select 'EUR' as crcycode union all select 'USD' as crcycode union all select 'CNY' as crcycode union all (select crcycode from rmcrcytx group by crcycode  union all select 'EUR' as crcycode order by crcycode);"
        Dim message As String = vbEmpty
        Dim Dataset1 As New DataSet
        Dim Dataset2 As New DataSet
        Dim Dataset3 As New DataSet
        Dim Dataset4 As New DataSet
        Dim Dataset5 As New DataSet
        Dim Dataset6 As New DataSet
        FormMenu.ToolStripStatusLabel1.Text = "Please Wait...Initialize Chart1"
        'FormMenu.Refresh()
        Application.DoEvents()
        DbTools1.getDataSet(sqlstr, Dataset1, message)
        MyChart1.DataSet1 = Dataset1

        FormMenu.ToolStripStatusLabel1.Text = "Please Wait...Initialize Chart2"
        'FormMenu.Refresh()
        Application.DoEvents()
        'DbTools1.getDataSet(sqlstr, Dataset2, message)
        Dataset2 = Dataset1.Copy
        MyChart2.DataSet1 = Dataset2

        FormMenu.ToolStripStatusLabel1.Text = "Please Wait...Initialize Chart3"
        'FormMenu.Refresh()
        Application.DoEvents()
        'DbTools1.getDataSet(sqlstr, Dataset3, message)
        Dataset3 = Dataset1.Copy
        MyChart3.DataSet1 = Dataset3

        FormMenu.ToolStripStatusLabel1.Text = "Please Wait...Initialize Chart4"
        'FormMenu.Refresh()
        Application.DoEvents()
        'DbTools1.getDataSet(sqlstr, Dataset4, message)
        Dataset4 = Dataset1.Copy
        MyChart4.DataSet1 = Dataset4

        FormMenu.ToolStripStatusLabel1.Text = "Please Wait...Initialize Chart5"
        'FormMenu.Refresh()
        Application.DoEvents()
        'DbTools1.getDataSet(sqlstr, Dataset5, message)
        Dataset5 = Dataset1.Copy
        MyChart5.DataSet1 = Dataset5

        FormMenu.ToolStripStatusLabel1.Text = "Please Wait...Initialize Chart6"
        'FormMenu.Refresh()
        Application.DoEvents()
        'DbTools1.getDataSet(sqlstr, Dataset6, message)
        Dataset6 = Dataset1.Copy
        MyChart6.DataSet1 = Dataset6
        FormMenu.ToolStripStatusLabel1.Text = ""
        Cursor.Current = Cursors.Default
    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Export to Excel
        Dim Result As Boolean = False
        Cursor.Current = Cursors.WaitCursor
        Dim Filename As String = "MyChart-" & Format(DateTime.Today, "yyyyMMdd") & ".xlsx"
        Dim source As String = ""
        'ask export location
        Dim DirectoryBrowser As FolderBrowserDialog = New FolderBrowserDialog
        DirectoryBrowser.Description = "Which directory do you want to use?"
        If (DirectoryBrowser.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Application.DoEvents()
            source = DirectoryBrowser.SelectedPath & "\" & filename

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
                oWb = oXl.Workbooks.Open(Application.StartupPath & "\templates\ChartTemplate.xltx")
                'Loop for chart
                oSheet = oWb.Worksheets(1)
                oSheet.Name = "Graph"
                Dim MemoryStream1 As MemoryStream
                Dim position(,) As Double = {{74, 0}, {74, 462}, {297, 0}, {297, 462}, {525, 0}, {525, 462}}
                Dim mycoll As New Collection
                mycoll.Add(MyChart1.Chart1, 0)
                mycoll.Add(MyChart2.Chart1, 1)
                mycoll.Add(MyChart3.Chart1, 2)
                mycoll.Add(MyChart4.Chart1, 3)
                mycoll.Add(MyChart5.Chart1, 4)
                mycoll.Add(MyChart6.Chart1, 5)

                For i = 0 To 5
                    '   copy chart
                    MemoryStream1 = New MemoryStream
                    Dim charting As Object = mycoll(i + 1)
                    charting.SaveImage(MemoryStream1, System.Drawing.Imaging.ImageFormat.Bmp)
                    Dim bmp As New Bitmap(MemoryStream1)
                    Clipboard.SetDataObject(bmp)
                    '   paste to excel
                    oSheet.Paste()
                    '   adjust location
                    Dim myobj As Object = oSheet.Shapes.Item(oSheet.Shapes.Count)
                    myobj.top = position(i, 0)
                    myobj.left = position(i, 1)
                    'end loop
                    Application.DoEvents()
                Next
                'save excel
                'Label10.Text = "Raw Material Figures available  From " & Format(CDate(Chart1.Series(0).Points(0).AxisLabel), "dd-MMM-yyyy") & " to " & Format(CDate(Chart1.Series(0).Points(Chart1.Series(0).Points.Count - 1).AxisLabel), "dd-MMM-yyyy")
                'put other info
                If MyChart1.Chart1.Series(0).Points.Count <> 0 Then
                    oSheet.Cells(2, 1) = MyChart1.ComboBox1.Text.ToString
                    oSheet.Cells(2, 2) = MyChart1.ComboBox2.Text.ToString
                    oSheet.Cells(2, 3) = MyChart1.ComboBox3.Text.ToString
                    oSheet.Cells(2, 4) = Format(CDate(MyChart1.Chart1.Series(0).Points(0).AxisLabel), "dd-MMM-yyyy") ' MyChart1.DateTimePicker1.Value.ToString
                    oSheet.Cells(2, 5) = Format(CDate(MyChart1.Chart1.Series(0).Points(MyChart1.Chart1.Series(0).Points.Count - 1).AxisLabel), "dd-MMM-yyyy")   'MyChart1.DateTimePicker2.Value.ToString
                    oSheet.Cells(2, 6) = MyChart1.TextBox4.Text.ToString
                    oSheet.Cells(4, 4) = MyChart1.TextBox1.Text.ToString
                    oSheet.Cells(4, 5) = MyChart1.TextBox2.Text.ToString
                    oSheet.Cells(4, 6) = MyChart1.TextBox3.Text.ToString
                End If
                If MyChart2.Chart1.Series(0).Points.Count <> 0 Then
                    oSheet.Cells(2, 8) = MyChart2.ComboBox1.Text.ToString
                    oSheet.Cells(2, 9) = MyChart2.ComboBox2.Text.ToString
                    oSheet.Cells(2, 10) = MyChart2.ComboBox3.Text.ToString
                    oSheet.Cells(2, 11) = Format(CDate(MyChart2.Chart1.Series(0).Points(0).AxisLabel), "dd-MMM-yyyy") 'MyChart2.DateTimePicker1.Value.ToString
                    oSheet.Cells(2, 12) = Format(CDate(MyChart2.Chart1.Series(0).Points(MyChart2.Chart1.Series(0).Points.Count - 1).AxisLabel), "dd-MMM-yyyy") 'MyChart2.DateTimePicker2.Value.ToString
                    oSheet.Cells(2, 13) = MyChart2.TextBox4.Text.ToString
                    oSheet.Cells(4, 11) = MyChart2.TextBox1.Text.ToString
                    oSheet.Cells(4, 12) = MyChart2.TextBox2.Text.ToString
                    oSheet.Cells(4, 13) = MyChart2.TextBox3.Text.ToString
                End If
                If MyChart3.Chart1.Series(0).Points.Count <> 0 Then
                    oSheet.Cells(17, 1) = MyChart3.ComboBox1.Text.ToString
                    oSheet.Cells(17, 2) = MyChart3.ComboBox2.Text.ToString
                    oSheet.Cells(17, 3) = MyChart3.ComboBox3.Text.ToString
                    oSheet.Cells(17, 4) = Format(CDate(MyChart3.Chart1.Series(0).Points(0).AxisLabel), "dd-MMM-yyyy") 'MyChart3.DateTimePicker1.Value.ToString
                    oSheet.Cells(17, 5) = Format(CDate(MyChart3.Chart1.Series(0).Points(MyChart3.Chart1.Series(0).Points.Count - 1).AxisLabel), "dd-MMM-yyyy") 'MyChart3.DateTimePicker2.Value.ToString
                    oSheet.Cells(17, 6) = MyChart3.TextBox4.Text.ToString
                    oSheet.Cells(19, 4) = MyChart3.TextBox1.Text.ToString
                    oSheet.Cells(19, 5) = MyChart3.TextBox2.Text.ToString
                    oSheet.Cells(19, 6) = MyChart3.TextBox3.Text.ToString
                End If
                If MyChart4.Chart1.Series(0).Points.Count <> 0 Then
                    oSheet.Cells(17, 8) = MyChart4.ComboBox1.Text.ToString
                    oSheet.Cells(17, 9) = MyChart4.ComboBox2.Text.ToString
                    oSheet.Cells(17, 10) = MyChart4.ComboBox3.Text.ToString
                    oSheet.Cells(17, 11) = Format(CDate(MyChart4.Chart1.Series(0).Points(0).AxisLabel), "dd-MMM-yyyy") 'MyChart4.DateTimePicker1.Value.ToString
                    oSheet.Cells(17, 12) = Format(CDate(MyChart4.Chart1.Series(0).Points(MyChart4.Chart1.Series(0).Points.Count - 1).AxisLabel), "dd-MMM-yyyy") ' MyChart4.DateTimePicker2.Value.ToString
                    oSheet.Cells(17, 13) = MyChart4.TextBox4.Text.ToString
                    oSheet.Cells(19, 11) = MyChart4.TextBox1.Text.ToString
                    oSheet.Cells(19, 12) = MyChart4.TextBox2.Text.ToString
                    oSheet.Cells(19, 13) = MyChart4.TextBox3.Text.ToString
                End If
                If MyChart5.Chart1.Series(0).Points.Count <> 0 Then
                    oSheet.Cells(31, 1) = MyChart5.ComboBox1.Text.ToString
                    oSheet.Cells(31, 2) = MyChart5.ComboBox2.Text.ToString
                    oSheet.Cells(31, 3) = MyChart5.ComboBox3.Text.ToString
                    oSheet.Cells(31, 4) = Format(CDate(MyChart5.Chart1.Series(0).Points(0).AxisLabel), "dd-MMM-yyyy") 'MyChart5.DateTimePicker1.Value.ToString
                    oSheet.Cells(31, 5) = Format(CDate(MyChart5.Chart1.Series(0).Points(MyChart5.Chart1.Series(0).Points.Count - 1).AxisLabel), "dd-MMM-yyyy") 'MyChart5.DateTimePicker2.Value.ToString
                    oSheet.Cells(31, 6) = MyChart5.TextBox4.Text.ToString
                    oSheet.Cells(33, 4) = MyChart5.TextBox1.Text.ToString
                    oSheet.Cells(33, 5) = MyChart5.TextBox2.Text.ToString
                    oSheet.Cells(33, 6) = MyChart5.TextBox3.Text.ToString
                End If
                If MyChart6.Chart1.Series(0).Points.Count <> 0 Then
                    oSheet.Cells(31, 8) = MyChart6.ComboBox1.Text.ToString
                    oSheet.Cells(31, 9) = MyChart6.ComboBox2.Text.ToString
                    oSheet.Cells(31, 10) = MyChart6.ComboBox3.Text.ToString
                    oSheet.Cells(31, 11) = Format(CDate(MyChart6.Chart1.Series(0).Points(0).AxisLabel), "dd-MMM-yyyy") 'MyChart6.DateTimePicker1.Value.ToString
                    oSheet.Cells(31, 12) = Format(CDate(MyChart6.Chart1.Series(0).Points(MyChart6.Chart1.Series(0).Points.Count - 1).AxisLabel), "dd-MMM-yyyy") 'MyChart6.DateTimePicker2.Value.ToString
                    oSheet.Cells(31, 13) = MyChart6.TextBox4.Text.ToString
                    oSheet.Cells(33, 11) = MyChart6.TextBox1.Text.ToString
                    oSheet.Cells(33, 12) = MyChart6.TextBox2.Text.ToString
                    oSheet.Cells(33, 13) = MyChart6.TextBox3.Text.ToString
                End If
                Dim StringBuilder1 As New System.Text.StringBuilder
                Dim dataset1 As New DataSet
                Dim iSheet As Integer = 0
                For i = 0 To 5
                    StringBuilder1 = New System.Text.StringBuilder
                    dataset1 = New DataSet
                    Select Case i
                        Case 0
                            dataset1 = MyChart1.getDataset
                        Case 1
                            dataset1 = MyChart2.getDataset
                        Case 2
                            dataset1 = MyChart3.getDataset
                        Case 3
                            dataset1 = MyChart4.getDataset
                        Case 4
                            dataset1 = MyChart5.getDataset
                        Case 5
                            dataset1 = MyChart6.getDataset
                    End Select
                    oSheet = oWb.Worksheets(2 + iSheet)
                    If Not (dataset1 Is Nothing) Then
                        If dataset1.Tables.Count <> 0 Then
                            oWb.Worksheets(2 + iSheet).select()
                            Dim myname As String = dataset1.Tables(0).TableName
                            If dataset1.Tables(0).TableName.Length > 30 Then
                                myname = dataset1.Tables(0).TableName.Substring(0, 29)
                            End If
                            oSheet.Name = iSheet + 1 & " " & ValidateSheetName(myname)
                            If Not ConvertDataForExcel(StringBuilder1, dataset1.Tables(0).DefaultView) Then
                                Return
                            End If
                            'put other info
                            Dim abc As String = StringBuilder1.ToString
                            Clipboard.SetDataObject(StringBuilder1.ToString, False)
                            oRange = oSheet.Range("A1")
                            oRange.Select()
                            oSheet.Paste()
                            oSheet.Cells.EntireColumn.AutoFit()
                            iSheet = iSheet + 1
                        End If
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


    Private Sub Chart1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class