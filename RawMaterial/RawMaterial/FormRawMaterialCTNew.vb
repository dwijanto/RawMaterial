Imports DJLib
Imports DJLib.Dbtools
Imports Microsoft.Office.Interop

Public Class FormRawMaterialCT
    Dim sqlstr As String
    Private Dbtools1 As New Dbtools(myUserid, myPassword)
    Dim DataSet1 As DataSet
    Dim DataSet2 As DataSet
    Dim DataSet3 As DataSet
    Dim DataSet4 As DataSet
    Dim DataSet5 As DataSet
    Dim DataSet6 As DataSet
    Dim myCol As New Collection

    Private Sub Set_RecordCount(ByVal RecordCount As Long) Handles DataGridRM1.RecordCount, DataGridRM2.RecordCount, DataGridRM3.RecordCount, DataGridRM4.RecordCount, DataGridRM5.RecordCount, DataGridRM6.RecordCount
        ToolStripStatusLabel1.Text = "Record Count : " & RecordCount
    End Sub

    Private Sub FormRawMaterialCT_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'sqlstr = "select setval('rawmaterial',1,false);select * from crosstab('select txdate,myid,txamount from rmmaterialrawtx rm " & _
        '         " join (select keyid , nextval(''rawmaterial'') as myid from (select keyid from rmmaterialraw where sheetname = ''CARBON'' order by keyid)as foo) tb on tb.keyid = rm.keyid order by 1','select * from generate_series(1," & 38 & ")m') as ("
        Dim errorMessage As String = String.Empty
        If Dataset1 Is Nothing Then
            Dataset1 = New DataSet
            getData(DataSet1, 0)
            DataGridRM1.DataSet1 = Dataset1
            DataGridRM1.LoadData()
            DataSet1.Tables(0).TableName = "CARBON"
            Dbtools1.GetDataCollection("select date_part('day',dvalue), cvalue from rmparamdt where paramname = 'sheet' order by cvalue;", myCol, errorMessage)
        End If
    End Sub


    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        Cursor = Cursors.WaitCursor
        Select Case TabControl1.SelectedIndex
            Case 0
                If Dataset1 Is Nothing Then
                    Dataset1 = New DataSet
                    getData(DataSet1, 0)
                    DataGridRM1.DataSet1 = Dataset1
                    DataGridRM1.LoadData()
                    DataSet1.Tables(0).TableName = "CARBON"
                Else
                    Set_RecordCount(DataGridRM1.GetRecordCount)
                End If
            Case 1
                If DataSet2 Is Nothing Then
                    DataSet2 = New DataSet
                    getData(DataSet2, 1)
                    DataGridRM2.DataSet1 = DataSet2
                    DataGridRM2.LoadData()
                    DataSet2.Tables(0).TableName = "NON FERROUS"
                Else
                    Set_RecordCount(DataGridRM2.GetRecordCount)
                End If
            Case 2
                If DataSet3 Is Nothing Then
                    DataSet3 = New DataSet
                    getData(DataSet3, 2)
                    DataGridRM3.DataSet1 = DataSet3
                    DataGridRM3.LoadData()
                    DataSet3.Tables(0).TableName = "OTHERS"
                Else
                    Set_RecordCount(DataGridRM3.GetRecordCount)
                End If
            Case 3
                If DataSet4 Is Nothing Then
                    DataSet4 = New DataSet
                    getData(DataSet4, 3)
                    DataGridRM4.DataSet1 = DataSet4
                    DataGridRM4.LoadData()
                    DataSet4.Tables(0).TableName = "PLASTICS"
                Else
                    Set_RecordCount(DataGridRM4.GetRecordCount)
                End If
            Case 4
                If DataSet5 Is Nothing Then
                    DataSet5 = New DataSet
                    getData(DataSet5, 4)
                    DataGridRM5.DataSet1 = DataSet5
                    DataGridRM5.LoadData()
                    DataSet5.Tables(0).TableName = "SCRAP"
                Else
                    Set_RecordCount(DataGridRM5.GetRecordCount)
                End If
            Case 5
                If DataSet6 Is Nothing Then
                    DataSet6 = New DataSet
                    getData(DataSet6, 5)
                    DataGridRM6.DataSet1 = DataSet6
                    DataGridRM6.LoadData()
                    DataSet6.Tables(0).TableName = "STAINLESS STEEL"
                Else
                    Set_RecordCount(DataGridRM6.GetRecordCount)
                End If
        End Select
        Cursor = Cursors.Default
    End Sub

    Private Sub getData(ByVal Dataset As DataSet, ByRef TabIndex As Integer)
        Dim errorMessage As String = String.Empty
        Dim DataTable As New DataTable
        Dim myDataset As New DataSet
        sqlstr = "select setval('rawmaterial',1,false);select nextval('rawmaterial') ||' '|| materialname as materialname from (select (rawmaterialname || ' (' || source || ' - ' || unit || ')' || ' (' || case when length(currency) > 0  then currency else 'No Crcy' end  || ')'  ) as materialname from rmmaterialraw where upper(sheetname) = '" & UCase(TabControl1.TabPages(TabIndex).Text) & "' order by rawmaterialname) as foo"
        Dbtools1.getDataSet(sqlstr, myDataset, errorMessage)

        Dim myList(myDataset.Tables(1).Rows.Count - 1) As String

        Dim i As Integer = 0
        For Each row As DataRow In myDataset.Tables(1).Rows
            myList(i) = Chr(34) & Trim(row.Item(0).ToString) & Chr(34) & " numeric"
            i = i + 1
        Next
        sqlstr = "select setval('rawmaterial',1,false);select * from crosstab('select txdate,myid,txamount from rmmaterialrawtx rm " & _
                 " join (select keyid , nextval(''rawmaterial'') as myid from (select keyid from rmmaterialraw where upper(sheetname) = ''" & UCase(TabControl1.TabPages(TabIndex).Text) & "'' order by keyid)as foo) tb on tb.keyid = rm.keyid order by 1','select * from generate_series(1," & myDataset.Tables(1).Rows.Count & ")m') as (""Date"" date," & String.Join(",", myList) & ");" & _
                 " select '' as crcycode union all select 'EUR' as crcycode union all select 'USD' as crcycode union all select 'CNY' as crcycode union all (select crcycode from rmcrcytx group by crcycode  union all select 'EUR' as crcycode order by crcycode);" & _
                 " select 'All' as rawmaterialname,'0' as keyid,'' as sheetname union all select rawmaterialname,keyid,sheetname  from rmmaterialraw where upper(sheetname) = '" & UCase(TabControl1.TabPages(TabIndex).Text) & "' order by keyid;"

        sqlstr = "select setval('rawmaterial',1,false);" & _
                 "select * from crosstab('select txdate,myid,txamount from rmmaterialrawtx rm " & _
                         " join (select keyid , nextval(''rawmaterial'') as myid from (select keyid from rmmaterialraw where upper(sheetname) = ''" & UCase(TabControl1.TabPages(TabIndex).Text) & "'' order by rawmaterialname)as foo) tb on tb.keyid = rm.keyid order by 1','select * from generate_series(1," & myDataset.Tables(1).Rows.Count & ")m') as (""Date"" date," & String.Join(",", myList) & ");" & _
                         " select '' as crcycode union all select 'EUR' as crcycode union all select 'USD' as crcycode union all select 'CNY' as crcycode union all (select crcycode from rmcrcytx group by crcycode  union all select 'EUR' as crcycode order by crcycode);" & _
                         " select '  All' as rawmaterialname,'0' as keyid,'' as sheetname,'  All' as namesource union all select rawmaterialname,keyid,sheetname,rawmaterialname || ' (' || source || ' - ' || unit || ')' as namesource   from rmmaterialraw where upper(sheetname) = '" & UCase(TabControl1.TabPages(TabIndex).Text) & "' order by rawmaterialname;"

        Dbtools1.getDataSet(sqlstr, Dataset, errorMessage)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If ExportSelectedItems() Then 'getcountlist
            ExportSelectedItemsToExcel()
        Else
            If TabSelection.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Dim myDataSetCollection As New Collection

                If TabSelection.RadioButton1.Checked Then
                    Select Case TabControl1.SelectedIndex
                        Case 0
                            myDataSetCollection.Add(DataSet1, 0)
                        Case 1
                            myDataSetCollection.Add(DataSet2, 0)
                        Case 2
                            myDataSetCollection.Add(DataSet3, 0)
                        Case 3
                            myDataSetCollection.Add(DataSet4, 0)
                        Case 4
                            myDataSetCollection.Add(DataSet5, 0)
                        Case 5
                            myDataSetCollection.Add(DataSet6, 0)
                    End Select
                Else
                    If DataSet1 Is Nothing Then
                        DataSet1 = New DataSet
                        getData(DataSet1, 0)
                        DataGridRM1.DataSet1 = DataSet1
                        DataGridRM1.LoadData()
                    End If
                    If DataSet2 Is Nothing Then
                        DataSet2 = New DataSet
                        getData(DataSet2, 1)
                        DataGridRM2.DataSet1 = DataSet2
                        DataGridRM2.LoadData()
                    End If
                    If DataSet3 Is Nothing Then
                        DataSet3 = New DataSet
                        getData(DataSet3, 2)
                        DataGridRM3.DataSet1 = DataSet3
                        DataGridRM3.LoadData()
                    End If
                    If DataSet4 Is Nothing Then
                        DataSet4 = New DataSet
                        getData(DataSet4, 3)
                        DataGridRM4.DataSet1 = DataSet4
                        DataGridRM4.LoadData()
                    End If
                    If DataSet5 Is Nothing Then
                        DataSet5 = New DataSet
                        getData(DataSet5, 4)
                        DataGridRM5.DataSet1 = DataSet5
                        DataGridRM5.LoadData()
                    End If
                    If DataSet6 Is Nothing Then
                        DataSet6 = New DataSet
                        getData(DataSet6, 5)
                        DataGridRM6.DataSet1 = DataSet6
                        DataGridRM6.LoadData()
                    End If
                    myDataSetCollection.Add(DataSet1, 0)
                    myDataSetCollection.Add(DataSet2, 1)
                    myDataSetCollection.Add(DataSet3, 3)
                    myDataSetCollection.Add(DataSet4, 4)
                    myDataSetCollection.Add(DataSet5, 5)
                    myDataSetCollection.Add(DataSet6, 6)
                End If
                Dim Filename As String = "MyRawData-" & Format(DateTime.Today, "yyyyMMdd") & ".xlsx"
                Cursor = Cursors.WaitCursor
                ExcelStuff.ExportToExcel(Filename, myDataSetCollection, 1)
                Cursor = Cursors.Default
            End If
        End If
    End Sub

    Private Function ExportSelectedItems() As Boolean
        Dim myResult As Boolean = True
        Dim mycount As Integer = 0
        Select Case TabControl1.SelectedIndex
            Case 0

                mycount = DataGridRM1.Countlist(DataGridRM1.CheckedListBox1)
                If mycount = 0 Or mycount = DataGridRM1.CheckedListBox1.Items.Count Then myResult = False
                If DataGridRM1.ComboBox1.Text <> "" Then myResult = True
            Case 1

                mycount = DataGridRM2.Countlist(DataGridRM2.CheckedListBox1)
                If mycount = 0 Or mycount = DataGridRM2.CheckedListBox1.Items.Count Then myResult = False
                If DataGridRM2.ComboBox1.Text <> "" Then myResult = True
            Case 2

                mycount = DataGridRM3.Countlist(DataGridRM3.CheckedListBox1)
                If mycount = 0 Or mycount = DataGridRM3.CheckedListBox1.Items.Count Then myResult = False
                If DataGridRM3.ComboBox1.Text <> "" Then myResult = True
            Case 3

                mycount = DataGridRM4.Countlist(DataGridRM4.CheckedListBox1)
                If mycount = 0 Or mycount = DataGridRM4.CheckedListBox1.Items.Count Then myResult = False
                If DataGridRM4.ComboBox1.Text <> "" Then myResult = True
            Case 4

                mycount = DataGridRM5.Countlist(DataGridRM5.CheckedListBox1)
                If mycount = 0 Or mycount = DataGridRM5.CheckedListBox1.Items.Count Then myResult = False
                If DataGridRM5.ComboBox1.Text <> "" Then myResult = True
            Case 5

                mycount = DataGridRM6.Countlist(DataGridRM6.CheckedListBox1)
                If mycount = 0 Or mycount = DataGridRM6.CheckedListBox1.Items.Count Then myResult = False
                If DataGridRM6.ComboBox1.Text <> "" Then myResult = True
        End Select
        Return myResult
    End Function

    Private Sub ExportSelectedItemsToExcel()
        'get period myCol.Item(ComboBox1.Text)
        Dim myvalue As Integer = myCol.Item(TabControl1.SelectedIndex + 1)
        'populate items
        Dim mysb As New System.Text.StringBuilder
        Dim myCriteria As String = String.Empty
        Dim myTabname As String = UCase(TabControl1.TabPages(TabIndex).Text)
        Select Case TabControl1.SelectedIndex
            Case 0
                myCriteria = DataGridRM1.getStringBuilderSingle(myTabname)
                mysb = SBClass.GenSBRMSingle(myvalue, DataGridRM1.ComboBox1.Text)
            Case 1
                myCriteria = DataGridRM2.getStringBuilderSingle(myTabname)
                mysb = SBClass.GenSBRMSingle(myvalue, DataGridRM2.ComboBox1.Text)
            Case 2
                myCriteria = DataGridRM3.getStringBuilderSingle(myTabname)
                mysb = SBClass.GenSBRMSingle(myvalue, DataGridRM3.ComboBox1.Text)
            Case 3
                myCriteria = DataGridRM4.getStringBuilderSingle(myTabname)
                mysb = SBClass.GenSBRMSingle(myvalue, DataGridRM4.ComboBox1.Text)
            Case 4
                myCriteria = DataGridRM5.getStringBuilderSingle(myTabname)
                mysb = SBClass.GenSBRMSingle(myvalue, DataGridRM5.ComboBox1.Text)
            Case 5
                myCriteria = DataGridRM6.getStringBuilderSingle(myTabname)
                mysb = SBClass.GenSBRMSingle(myvalue, DataGridRM6.ComboBox1.Text)
        End Select
        'MsgBox(mysb.ToString)

        Dim sqlstr As String = mysb.ToString & myCriteria

        'ExcelStuff.ExportToExcelAskDirectory(Filename, mysb.ToString & myCriteria, Dbtools1, , , , 2)
 


        'Export to Excel
        Dim Result As Boolean = False
        Cursor.Current = Cursors.WaitCursor
        Dim Filename As String = "MyRawData-" & Format(DateTime.Today, "yyyyMMdd") & ".xlsx"
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
                oWb = oXl.Workbooks.Open(Application.StartupPath & "\templates\RawDataTemplate.xltx")

                'Loop for chart
                oWb.Worksheets(2).select()


                ExcelStuff.FillDataSource(oWb, 2, sqlstr, Dbtools1)

                oSheet = oWb.Worksheets(2)
                oSheet.Name = "DATA"
                'set DbRange
                oWb.Names.Add(Name:="DBRange", RefersToR1C1:="=OFFSET('" & oSheet.Name & "'!R1C1,0,0,COUNTA('" & oSheet.Name & "'!C1),COUNTA('" & oSheet.Name & "'!R1))")

                'Go To Worksheet(1)
                oSheet = oWb.Worksheets(1)
                oSheet.Name = "PivotTable"
                oWb.Worksheets(1).select()

                oWb.PivotCaches.Add(Excel.XlPivotTableSourceType.xlDatabase, "DBRange").CreatePivotTable(oSheet.Name & "!R6C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
                oSheet.PivotTables("PivotTable1").columngrand = False
                oSheet.PivotTables("PivotTable1").rowgrand = False
                oSheet.PivotTables("PivotTable1").ingriddropzones = True
                oSheet.PivotTables("PivotTable1").rowaxislayout(Excel.XlLayoutRowType.xlTabularRow)

                'add Rowfields
                oSheet.PivotTables("PivotTable1").PivotFields("Tx Date").orientation = Excel.XlPivotFieldOrientation.xlRowField
               
                'add columnfield
                oSheet.PivotTables("PivotTable1").PivotFields("rawmaterialname").orientation = Excel.XlPivotFieldOrientation.xlColumnField


                'add datafield
                oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("Original Amount"), " Original Amount", Excel.XlConsolidationFunction.xlSum)
                oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("conversion"), " Conversion", Excel.XlConsolidationFunction.xlSum)
                Dim myName As String = " " & oSheet.PivotTables("PivotTable1").PivotFields(5).name
                oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields(5), myName, Excel.XlConsolidationFunction.xlSum)
                oSheet.PivotTables("PivotTable1").DataPivotField.Orientation = Excel.XlPivotFieldOrientation.xlColumnField
                oSheet.PivotTables("PivotTable1").DataPivotField.position = 2
                'sort column period                
                oSheet.Cells.EntireColumn.ColumnWidth = 20
                oSheet.Rows("7:7").wraptext = True                
                oSheet.Name = "PivotTable " & myTabname
                oWb.Worksheets(1).select()

                'Do Pivot Table

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