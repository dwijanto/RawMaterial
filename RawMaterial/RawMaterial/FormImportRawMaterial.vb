Imports Microsoft.Office.Interop
Imports System.Data
Imports System.Data.Common
Imports DJLib
Imports DJLib.Dbtools
Imports System.IO
Imports System.ComponentModel
Imports Npgsql

Public Class FormImportRawMaterial
    Private WithEvents backgroundworker1 As BackgroundWorker
    'Private DataAdapterHeader As DbDataAdapter
    'Private DataAdapterDetail As DbDataAdapter
    Private DataAdapterHeader As NpgsqlDataAdapter
    Private DataAdapterDetail As NpgsqlDataAdapter
    Private Dataset1 As DataSet
    Private DataTableHeader1 As DataTable
    Private DataTableDetail1 As DataTable
    Private DataHeaderCollection As New Collection()
    Private SheetCollection As New Collection
    Private Dbtools1 As New Dbtools(myUserid, myPassword)
    Private Param As DbParameter
    Private FileName As String
    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        backgroundworker1 = New BackgroundWorker
        ' Add any initialization after the InitializeComponent() call.
    End Sub

    Private Sub FormImportRawMaterial_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If backgroundworker1.IsBusy Then
            MsgBox("Please wait until the current process is finished")
            e.Cancel = True
        End If
    End Sub

    Private Sub FormImportExcel_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TextBox1.Text = My.Settings.ImportExcelFolder
        Dim errorMessage As String = vbEmpty
        If Not Dbtools1.GetDataCollection("select ivalue,cvalue from rmparamdt where paramname = 'sheet'", SheetCollection, errorMessage) Then
            MessageBox.Show(errorMessage)
        End If

    End Sub
    Private Sub ConnectData()
        'Dim factory As DbProviderFactory = DbProviderFactories.GetFactory(My.Settings.dbProviderFactory)
        'Dim conn As DbConnection = factory.CreateConnection
        Dim conn = New NpgsqlConnection(Dbtools1.getConnectionString)
        conn.ConnectionString = Dbtools1.getConnectionString
        Try
            'Dim SelectCmd As DbCommand = factory.CreateCommand
            Dim selectcmd = New NpgsqlCommand()
            SelectCmd.Connection = conn
            SelectCmd.CommandText = "select keyid,updateby,source,links,area,currency,unit,sheetname,rawmaterialname from rmmaterialraw"
            selectcmd.CommandType = CommandType.Text

            'DataAdapterHeader = factory.CreateDataAdapter
            DataAdapterHeader = New NpgsqlDataAdapter
            DataAdapterHeader.SelectCommand = selectcmd

            'Dim CommandBuilder As DbCommandBuilder = factory.CreateCommandBuilder
            Dim CommandBuilder = New NpgsqlCommandBuilder
            CommandBuilder.DataAdapter = DataAdapterHeader
            DataAdapterHeader.InsertCommand = CommandBuilder.GetInsertCommand
            DataAdapterHeader.UpdateCommand = CommandBuilder.GetUpdateCommand

            'Dim SelectCmd2 As DbCommand = factory.CreateCommand
            Dim selectcmd2 = New NpgsqlCommand()
            SelectCmd2.Connection = conn
            SelectCmd2.CommandText = "select materialrawtxid,txdate,crcycode,txamount,unit,keyid from rmmaterialrawtx"
            selectcmd2.CommandType = CommandType.Text

            'DataAdapterDetail = factory.CreateDataAdapter
            DataAdapterDetail = New NpgsqlDataAdapter
            DataAdapterDetail.SelectCommand = selectcmd2

            'Dim CommandBuilder2 As DbCommandBuilder = factory.CreateCommandBuilder
            Dim CommandBuilder2 = New NpgsqlCommandBuilder

            CommandBuilder2.DataAdapter = DataAdapterDetail
            DataAdapterDetail.InsertCommand = CommandBuilder2.GetInsertCommand

            'Dim updateCmd As DbCommand = factory.CreateCommand
            Dim updateCmd = New NpgsqlCommand()
            updateCmd.Connection = conn
            updateCmd.CommandText = "Update rmmaterialrawtx set txdate=@txdate,keyid=@keyid,crcycode=@crcycode,txamount=@txamount,unit=@unit where materialrawtxid = @materialrawtxid"
            DataAdapterDetail.UpdateCommand = updateCmd
            Application.DoEvents()
            Param = updateCmd.CreateParameter
            Param.SourceColumn = "txdate"
            Param.ParameterName = "@txdate"
            'Param.DbType = DbType.Date
            Param.Value = New DateTime(DataRowVersion.Current).ToString("yyyy-MM-dd HH:mm:ss.fffff")
            updateCmd.Parameters.Add(Param)

            Param = updateCmd.CreateParameter
            Param.SourceColumn = "keyid"
            Param.ParameterName = "@keyid"
            Param.DbType = DbType.String
            Param.Value = DataRowVersion.Current
            updateCmd.Parameters.Add(Param)

            Param = updateCmd.CreateParameter
            Param.SourceColumn = "crcycode"
            Param.ParameterName = "@crcycode"
            Param.DbType = DbType.String
            Param.Value = DataRowVersion.Current
            updateCmd.Parameters.Add(Param)

            Param = updateCmd.CreateParameter
            Param.SourceColumn = "txamount"
            Param.ParameterName = "@txamount"
            Param.DbType = DbType.Decimal
            Param.Value = DataRowVersion.Current
            updateCmd.Parameters.Add(Param)

            Param = updateCmd.CreateParameter
            Param.SourceColumn = "unit"
            Param.ParameterName = "@unit"
            Param.DbType = DbType.String
            Param.Value = DataRowVersion.Current
            updateCmd.Parameters.Add(Param)

            Param = updateCmd.CreateParameter
            Param.SourceColumn = "materialrawtxid"
            Param.ParameterName = "@materialrawtxid"
            Param.DbType = DbType.Int64
            Param.Value = DataRowVersion.Original
            updateCmd.Parameters.Add(Param)

            DataAdapterDetail.DeleteCommand = CommandBuilder2.GetDeleteCommand
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub ConnectDataFactory()
        Dim factory As DbProviderFactory = DbProviderFactories.GetFactory(My.Settings.dbProviderFactory)
        Dim conn As DbConnection = factory.CreateConnection
        conn.ConnectionString = Dbtools1.getConnectionString
        Try
            Dim SelectCmd As DbCommand = factory.CreateCommand
            SelectCmd.Connection = conn
            SelectCmd.CommandText = "select keyid,updateby,source,links,area,currency,unit,sheetname,rawmaterialname from rmmaterialraw"
            SelectCmd.CommandType = CommandType.Text
            DataAdapterHeader = factory.CreateDataAdapter
            DataAdapterHeader.SelectCommand = SelectCmd
            Dim CommandBuilder As DbCommandBuilder = factory.CreateCommandBuilder
            CommandBuilder.DataAdapter = DataAdapterHeader
            DataAdapterHeader.InsertCommand = CommandBuilder.GetInsertCommand
            DataAdapterHeader.UpdateCommand = CommandBuilder.GetUpdateCommand

            Dim SelectCmd2 As DbCommand = factory.CreateCommand
            SelectCmd2.Connection = conn
            SelectCmd2.CommandText = "select materialrawtxid,txdate,crcycode,txamount,unit,keyid from rmmaterialrawtx"
            SelectCmd2.CommandType = CommandType.Text
            DataAdapterDetail = factory.CreateDataAdapter
            DataAdapterDetail.SelectCommand = SelectCmd2

            Dim CommandBuilder2 As DbCommandBuilder = factory.CreateCommandBuilder
            CommandBuilder2.DataAdapter = DataAdapterDetail
            DataAdapterDetail.InsertCommand = CommandBuilder2.GetInsertCommand

            Dim updateCmd As DbCommand = factory.CreateCommand
            updateCmd.Connection = conn
            updateCmd.CommandText = "Update rmmaterialrawtx set txdate=@txdate,keyid=@keyid,crcycode=@crcycode,txamount=@txamount,unit=@unit where materialrawtxid = @materialrawtxid"
            DataAdapterDetail.UpdateCommand = updateCmd
            Application.DoEvents()
            Param = updateCmd.CreateParameter
            Param.SourceColumn = "txdate"
            Param.ParameterName = "@txdate"
            'Param.DbType = DbType.Date
            Param.Value = New DateTime(DataRowVersion.Current).ToString("yyyy-MM-dd HH:mm:ss.fffff")
            updateCmd.Parameters.Add(Param)

            Param = updateCmd.CreateParameter
            Param.SourceColumn = "keyid"
            Param.ParameterName = "@keyid"
            Param.DbType = DbType.String
            Param.Value = DataRowVersion.Current
            updateCmd.Parameters.Add(Param)

            Param = updateCmd.CreateParameter
            Param.SourceColumn = "crcycode"
            Param.ParameterName = "@crcycode"
            Param.DbType = DbType.String
            Param.Value = DataRowVersion.Current
            updateCmd.Parameters.Add(Param)

            Param = updateCmd.CreateParameter
            Param.SourceColumn = "txamount"
            Param.ParameterName = "@txamount"
            Param.DbType = DbType.Decimal
            Param.Value = DataRowVersion.Current
            updateCmd.Parameters.Add(Param)

            Param = updateCmd.CreateParameter
            Param.SourceColumn = "unit"
            Param.ParameterName = "@unit"
            Param.DbType = DbType.String
            Param.Value = DataRowVersion.Current
            updateCmd.Parameters.Add(Param)

            Param = updateCmd.CreateParameter
            Param.SourceColumn = "materialrawtxid"
            Param.ParameterName = "@materialrawtxid"
            Param.DbType = DbType.Int64
            Param.Value = DataRowVersion.Original
            updateCmd.Parameters.Add(Param)

            DataAdapterDetail.DeleteCommand = CommandBuilder2.GetDeleteCommand
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub FillData()

        Try
            Dataset1 = New DataSet("MyDb")
            DataAdapterHeader.FillSchema(Dataset1, SchemaType.Source)
            DataAdapterHeader.MissingSchemaAction = MissingSchemaAction.AddWithKey 'Will add Primary & Unique key info
            DataTableHeader1 = New DataTable("materialraw")
            DataAdapterHeader.Fill(DataTableHeader1)

            DataAdapterDetail.FillSchema(Dataset1, SchemaType.Source)
            'DataAdapterDetail.MissingSchemaAction = MissingSchemaAction.AddWithKey 'Will add Primary & Unique key info
            Application.DoEvents()
            DataTableDetail1 = New DataTable("materialrawtx")
            DataAdapterDetail.Fill(DataTableDetail1)
            Application.DoEvents()
            Dim Keys(2) As DataColumn
            Keys(0) = DataTableDetail1.Columns(1) 'txdate
            Keys(1) = DataTableDetail1.Columns(5) 'keyid
            DataTableDetail1.PrimaryKey = Keys
            'DataView1 = New DataView(Dataset1.Tables("product")) 
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnImport.Click
        If Not (backgroundworker1.IsBusy) Then

            'get filename
            OpenFileDialog1.FileName = ""
            If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                TextBox3.Text = "File Checking..Please wait."
                Application.DoEvents()
                If OpenFileDialog1.CheckFileExists And OpenFileDialog1.CheckPathExists Then
                    FileName = OpenFileDialog1.FileName
                    TextBox1.Text = FileName
                    TextBox3.Text = "Open connection..."
                    Application.DoEvents()
                    Call ConnectData()
                    Application.DoEvents()
                    Call FillData()
                    Application.DoEvents()
                    TextBox3.Text = "Open Worker..."
                    Try
                        backgroundworker1.WorkerReportsProgress = True
                        backgroundworker1.WorkerSupportsCancellation = True
                        backgroundworker1.RunWorkerAsync()
                    Catch ex As Exception
                        MessageBox.Show(ex.Message)
                    End Try
                    
                End If


            End If

        Else
            MsgBox("Please wait until the current process is finished")
        End If
    End Sub

    Private Sub backgroundworker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles backgroundworker1.DoWork
        backgroundworker1.ReportProgress(3, TextBox3.Text & "Start")
        'Dim DirectoryInfo As New DirectoryInfo(My.Settings.ImportExcelFolder)
        'Dim i As Integer = 0
        'Dim TotalFiles As Integer = DirectoryInfo.GetFiles.Length
        'Dim Message As String = vbEmpty
        Dim Status As Boolean
        'backgroundworker1.ReportProgress(2, "")
        'backgroundworker1.ReportProgress(3, "")

        Status = ImportExcel(FileName)

        'For Each File In DirectoryInfo.GetFiles
        '    i += 1
        '    backgroundworker1.ReportProgress(3, "Processing file " & i & " of " & TotalFiles)
        '    'only interest in excel file
        '    If Path.GetExtension(File.FullName) = ".xlsx" Or Path.GetExtension(File.FullName) = ".xls" Then
        '        backgroundworker1.ReportProgress(2, File.FullName)
        '        Status = ImportExcel(File.FullName)
        '    Else
        '        Status = False
        '        Call MoveToFolder(My.Settings.ErrorFolder, File.FullName, Message)
        '    End If
        'Next
        backgroundworker1.ReportProgress(3, TextBox3.Text & ". Done.")
    End Sub

    Private Sub backgroundworker1_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles backgroundworker1.ProgressChanged
        Select Case e.ProgressPercentage
            Case 2
                TextBox2.Text = e.UserState
            Case 3
                TextBox3.Text = e.UserState
        End Select
    End Sub

    Private Sub backgroundworker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles backgroundworker1.RunWorkerCompleted
        btnImport.Enabled = True
        FormMenu.setBubbleMessage("Import Excel to DB", "Done")
        If CheckBoxClose.Checked Then
            Me.Close()
        End If
    End Sub

    Private Function ImportExcel1(ByVal FileName As String) As Boolean
        backgroundworker1.ReportProgress(3, "Excel..")
        Dim myret As Boolean = False
        Dim message As String = vbEmpty
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
        Dim aprocesses() As Process = Process.GetProcesses
        Dim aprocess As Process = Nothing
        Try
            backgroundworker1.ReportProgress(3, "Create Object..")
            oXl = CType(CreateObject("Excel.Application"), Excel.Application)
            oXl.Visible = True

            aprocesses = Process.GetProcesses
            For i = 0 To aprocesses.GetUpperBound(0)
                If aprocesses(i).MainWindowHandle.ToString = oXl.Hwnd.ToString Then
                    aprocess = aprocesses(i)
                    Exit For
                End If
            Next
            oXl.Visible = False

            oXl.DisplayAlerts = False
            backgroundworker1.ReportProgress(3, "Opening Excel..")
            oWb = oXl.Workbooks.Open(FileName)
            For i = 1 To oWb.Worksheets.Count
                backgroundworker1.ReportProgress(3, "Processing Sheet " & i & " of " & oWb.Worksheets.Count)
                SheetName = oWb.Worksheets(i).name
                oSheet = oWb.Worksheets(i)
                If ValidSheet(SheetName) Then
                    oSheet = oWb.Worksheets(i)

                    getUsedRange(oSheet, Row, Col)
                    backgroundworker1.ReportProgress(3, SheetName)
                    'get header
                    J = 2
                    For J = 2 To Col
                        If oSheet.Cells(1, J).value <> "" Then
                            'find key in table
                            Dim DataRow1 As DataRow = DataTableHeader1.Rows.Find(oSheet.Cells(1, J).value)
                            If IsNothing(DataRow1) Then
                                DataRow1 = DataTableHeader1.NewRow
                                DataRow1("keyid") = oSheet.Cells(1, J).value
                                DataRow1("updateby") = CleanChar(oSheet.Cells(2, J).value)
                                DataRow1("source") = CleanChar(oSheet.Cells(3, J).value)
                                DataRow1("links") = CleanChar(oSheet.Cells(4, J).value)
                                DataRow1("area") = CleanChar(oSheet.Cells(5, J).value)
                                DataRow1("currency") = CleanChar(oSheet.Cells(6, J).value)
                                DataRow1("unit") = CleanChar(oSheet.Cells(7, J).value)
                                DataRow1("rawmaterialname") = CleanChar(oSheet.Cells(8, J).value)
                                DataRow1("sheetname") = SheetName
                                DataTableHeader1.Rows.Add(DataRow1)
                            Else
                                DataRow1.BeginEdit()
                                Dim keyid1 As String = CType(oSheet.Cells(1, J).value, String)
                                DataRow1("keyid") = keyid1
                                DataRow1("updateby") = CleanChar(oSheet.Cells(2, J).value)
                                DataRow1("source") = CleanChar(oSheet.Cells(3, J).value)
                                DataRow1("links") = CleanChar(oSheet.Cells(4, J).value)
                                DataRow1("area") = CleanChar(oSheet.Cells(5, J).value)
                                DataRow1("currency") = CleanChar(oSheet.Cells(6, J).value)
                                DataRow1("unit") = CleanChar(oSheet.Cells(7, J).value)
                                DataRow1("rawmaterialname") = CleanChar(oSheet.Cells(8, J).value)
                                DataRow1("sheetname") = SheetName
                                DataRow1.EndEdit()
                            End If

                            'Update Detail
                            K = 9
                            For K = 9 To Row
                                'Detail
                                If CType(oSheet.Cells(K, 1).value, String) <> "" Then
                                    Dim pkey(1) As Object
                                    pkey(0) = oSheet.Cells(K, 1).value
                                    pkey(1) = oSheet.Cells(1, J).value
                                    Dim Datarow2 As DataRow = DataTableDetail1.Rows.Find(pkey)
                                    If IsNothing(Datarow2) Then
                                        If IsNumeric(oSheet.Cells(K, J).value) Then
                                            Datarow2 = DataTableDetail1.NewRow
                                            Datarow2("txdate") = oSheet.Cells(K, 1).value
                                            Datarow2("txamount") = ValidNumber(oSheet.Cells(K, J).value)
                                            Datarow2("crcycode") = oSheet.Cells(6, J).value
                                            Datarow2("keyid") = oSheet.Cells(1, J).value
                                            Datarow2("unit") = oSheet.Cells(7, J).value
                                            DataTableDetail1.Rows.Add(Datarow2)
                                        End If
                                    Else
                                        'prevent blank value
                                        If CType(oSheet.Cells(K, J).value, String) <> "" Then
                                            Datarow2.BeginEdit()
                                            Datarow2("txdate") = DateFormatDotNet(oSheet.Cells(K, 1).value)
                                            Datarow2("txamount") = ValidNumber(oSheet.Cells(K, J).value)
                                            Datarow2("crcycode") = oSheet.Cells(6, J).value
                                            Datarow2("keyid") = oSheet.Cells(1, J).value
                                            Datarow2("unit") = oSheet.Cells(7, J).value
                                            Datarow2.EndEdit()
                                        End If

                                    End If
                                End If
                                backgroundworker1.ReportProgress(3, "Processing Sheet " & SheetName & " " & i & " of " & oWb.Worksheets.Count & ". Col " & J & " of " & Col & " Row " & K & " of " & Row)
                            Next K
                        End If
                    Next J
                    backgroundworker1.ReportProgress(3, "Send to DB: Detail")
                    DataAdapterHeader.Update(DataTableHeader1)
                    backgroundworker1.ReportProgress(3, "Send to DB: Header")
                    DataAdapterDetail.Update(DataTableDetail1)
                End If
            Next
            myret = True
        Catch ex As Exception
            backgroundworker1.ReportProgress(3, ex.Message & " Sheet Name: " & SheetName & "Col:" & J & " Row:" & K)
        Finally
            oXl.Quit()
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

            'Try
            'If myret Then
            '    Call MoveToFolder(My.Settings.SuccessFolder, FileName, message)
            'Else
            '    Call MoveToFolder(My.Settings.ErrorFolder, FileName, message)
            'End If
            '    Catch ex As Exception
            '    MsgBox(ex.Message)
            'End Try
        End Try
        Return myret
    End Function

    Private Function ImportExcel(ByVal FileName As String) As Boolean
        backgroundworker1.ReportProgress(3, "Excel..")
        Dim myret As Boolean = False
        Dim message As String = vbEmpty
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
        Dim aprocesses() As Process = Process.GetProcesses
        Dim aprocess As Process = Nothing
        Try
            backgroundworker1.ReportProgress(3, "Create Object..")
            oXl = CType(CreateObject("Excel.Application"), Excel.Application)
            oXl.Visible = True

            aprocesses = Process.GetProcesses
            For i = 0 To aprocesses.GetUpperBound(0)
                If aprocesses(i).MainWindowHandle.ToString = oXl.Hwnd.ToString Then
                    aprocess = aprocesses(i)
                    Exit For
                End If
            Next
            oXl.Visible = False

            oXl.DisplayAlerts = False
            backgroundworker1.ReportProgress(3, "Opening Excel..")
            oWb = oXl.Workbooks.Open(FileName)
            For i = 1 To oWb.Worksheets.Count
                backgroundworker1.ReportProgress(3, "Processing Sheet " & i & " of " & oWb.Worksheets.Count)
                SheetName = oWb.Worksheets(i).name
                oSheet = oWb.Worksheets(i)
                If ValidSheet(SheetName) Then
                    oSheet = oWb.Worksheets(i)

                    getUsedRange(oSheet, Row, Col)
                    oRange = oSheet.Range("A1")
                    Row = GetLastRow(oXl, oSheet, oRange)
                    backgroundworker1.ReportProgress(3, SheetName)
                    'get header
                    J = 2
                    For J = 2 To Col
                        If oSheet.Cells(1, J).value <> "" Then
                            'find key in table
                            Dim DataRow1 As DataRow = DataTableHeader1.Rows.Find(oSheet.Cells(1, J).value)
                            If IsNothing(DataRow1) Then
                                DataRow1 = DataTableHeader1.NewRow
                                DataRow1("keyid") = oSheet.Cells(1, J).value
                                DataRow1("updateby") = CleanChar(oSheet.Cells(2, J).value)
                                DataRow1("source") = CleanChar(oSheet.Cells(3, J).value)
                                DataRow1("links") = CleanChar(oSheet.Cells(4, J).value)
                                DataRow1("area") = CleanChar(oSheet.Cells(5, J).value)
                                DataRow1("currency") = CleanChar(oSheet.Cells(6, J).value)
                                DataRow1("unit") = CleanChar(oSheet.Cells(7, J).value)
                                DataRow1("rawmaterialname") = CleanChar(oSheet.Cells(8, J).value)
                                DataRow1("sheetname") = SheetName
                                DataTableHeader1.Rows.Add(DataRow1)
                            Else
                                DataRow1.BeginEdit()
                                Dim keyid1 As String = CType(oSheet.Cells(1, J).value, String)
                                DataRow1("keyid") = keyid1
                                DataRow1("updateby") = CleanChar(oSheet.Cells(2, J).value)
                                DataRow1("source") = CleanChar(oSheet.Cells(3, J).value)
                                DataRow1("links") = CleanChar(oSheet.Cells(4, J).value)
                                DataRow1("area") = CleanChar(oSheet.Cells(5, J).value)
                                DataRow1("currency") = CleanChar(oSheet.Cells(6, J).value)
                                DataRow1("unit") = CleanChar(oSheet.Cells(7, J).value)
                                DataRow1("rawmaterialname") = CleanChar(oSheet.Cells(8, J).value)
                                DataRow1("sheetname") = SheetName
                                DataRow1.EndEdit()
                            End If

                            'Update Detail
                            K = 9
                            For K = 9 To Row
                                'Detail
                                If CType(oSheet.Cells(K, 1).value, String) <> "" Then
                                    Dim pkey(1) As Object
                                    pkey(0) = oSheet.Cells(K, 1).value
                                    pkey(1) = oSheet.Cells(1, J).value
                                    Dim Datarow2 As DataRow = DataTableDetail1.Rows.Find(pkey)
                                    If IsNothing(Datarow2) Then
                                        If IsNumeric(oSheet.Cells(K, J).value) Then
                                            If Not CType(oSheet.Cells(K, J).value, String) = "0" Then
                                                Datarow2 = DataTableDetail1.NewRow
                                                Datarow2("txdate") = oSheet.Cells(K, 1).value
                                                Datarow2("txamount") = ValidNumber(oSheet.Cells(K, J).value)
                                                Datarow2("crcycode") = oSheet.Cells(6, J).value
                                                Datarow2("keyid") = oSheet.Cells(1, J).value
                                                Datarow2("unit") = oSheet.Cells(7, J).value
                                                DataTableDetail1.Rows.Add(Datarow2)
                                            End If
                                        End If
                                    Else
                                        'prevent blank value
                                        If CType(oSheet.Cells(K, J).value, String) = "0" Then
                                            Datarow2.BeginEdit()
                                            Datarow2.Delete()
                                            Datarow2.EndEdit()
                                        ElseIf CType(oSheet.Cells(K, J).value, String) <> "" Then
                                            Datarow2.BeginEdit()
                                            Datarow2("txdate") = DateFormatDotNet(oSheet.Cells(K, 1).value)
                                            Datarow2("txamount") = ValidNumber(oSheet.Cells(K, J).value)
                                            Datarow2("crcycode") = oSheet.Cells(6, J).value
                                            Datarow2("keyid") = oSheet.Cells(1, J).value
                                            Datarow2("unit") = oSheet.Cells(7, J).value
                                            Datarow2.EndEdit()
                                        End If

                                    End If
                                End If
                                backgroundworker1.ReportProgress(3, "Processing Sheet " & SheetName & " " & i & " of " & oWb.Worksheets.Count & ". Col " & J & " of " & Col & " Row " & K & " of " & Row)
                            Next K
                        End If
                    Next J
                    backgroundworker1.ReportProgress(3, "Send to DB: Detail")
                    DataAdapterHeader.Update(DataTableHeader1)
                    backgroundworker1.ReportProgress(3, "Send to DB: Header")
                    DataAdapterDetail.Update(DataTableDetail1)
                End If
            Next
            myret = True
        Catch ex As Exception
            backgroundworker1.ReportProgress(3, ex.Message & " Sheet Name: " & SheetName & "Col:" & J & " Row:" & K)
        Finally
            oXl.Quit()
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

            'Try
            'If myret Then
            '    Call MoveToFolder(My.Settings.SuccessFolder, FileName, message)
            'Else
            '    Call MoveToFolder(My.Settings.ErrorFolder, FileName, message)
            'End If
            '    Catch ex As Exception
            '    MsgBox(ex.Message)
            'End Try
        End Try
        Return myret
    End Function

    Private Function ValidSheet(ByRef SheetName As String) As Boolean
        Dim myret As Boolean = False
        Try
            Dim myvalue As Integer = SheetCollection.Item(SheetName)
            myret = True
        Catch ex As Exception
        End Try
        Return myret
    End Function
    Private Function CleanChar(ByVal myString As String) As String
        Return Replace(myString, Chr(10), " ")
    End Function

    Public Function GetLastRow(ByVal oxl As Excel.Application, ByVal osheet As Excel.Worksheet, ByVal range As Excel.Range) As Long
        Dim lastrow As Long = 1
        oxl.ScreenUpdating = False
        Try
            lastrow = osheet.Cells.Find("*", range, , , Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious).Row
        Catch ex As Exception
        End Try
        Return lastrow
        oxl.ScreenUpdating = True
    End Function
End Class