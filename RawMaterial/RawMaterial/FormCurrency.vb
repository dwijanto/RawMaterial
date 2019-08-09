
Imports DJLib
Imports DJLib.Dbtools
Imports Microsoft.Office.Interop
Imports System.IO
Public Class FormCurrency

    Dim DataSet1 As DataSet
    Dim Dataset2 As DataSet
    Dim CurrencyManager1 As CurrencyManager
    Dim Bindingsource1 As New BindingSource
    Dim CurrTable As New DataTable
    Dim myFullPath As String = System.AppDomain.CurrentDomain.BaseDirectory & "myTmp.txt"

    Private Sub FormCurrency_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Application.DoEvents() 'we need this to refresh form main menu
        Cursor.Current = Cursors.WaitCursor

        Dim result As Boolean = False
        Dim message As String = String.Empty

        Dim DbTools1 As New DJLib.Dbtools(myUserid, myPassword)
        'Dim sqlstr As String = String.Empty
        DataSet1 = New DataSet
        Dim sqlstr As String = "select crcycode from rmcrcytx group by crcycode order by crcycode;"
        DbTools1.getDataSet(sqlstr, DataSet1)

        Dim myList(DataSet1.Tables(0).Rows.Count - 1) As String

        Dim i As Integer = 0
        For Each row As DataRow In DataSet1.Tables(0).Rows
            myList(i) = Chr(34) & row.Item(0).ToString & Chr(34) & " numeric"
            i = i + 1
        Next
        sqlstr = "select * from crosstab('select crcydate,rmtbcrcyid,crcyamount from rmcrcytx rm join rmtbcrcy tb on tb.crcycode = rm.crcycode order by 1','select * from generate_series(1," & DataSet1.Tables(0).Rows.Count & ")m') as (""Date"" date," & String.Join(",", myList) & ");"
        Dataset2 = New DataSet
        DbTools1.getDataSet(sqlstr, DataSet2)

        BindingData()
        Cursor.Current = Cursors.Default
    End Sub
    Private Sub FillData()

    End Sub
    Private Sub BindingData()
        Try
           
            Bindingsource1.DataSource = Dataset2.Tables(0).DefaultView 'DataSet1.Tables(0).DefaultView 'DataTable1 'dataview1
            DataGridView1.DataSource = Bindingsource1
            With DataGridView1
                .RowsDefaultCellStyle.BackColor = Color.Bisque
                .AlternatingRowsDefaultCellStyle.BackColor = Color.White
            End With
            Me.DataGridView1.Columns("Date").DefaultCellStyle.Format = "dd-MM-yyyy"
            CurrencyManager1 = CType(Me.BindingContext(Bindingsource1), CurrencyManager)
            ShowRecordCount(CurrencyManager1.Count)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub ShowRecordCount(ByRef RecordCount As Long)
        Me.ToolStripStatusLabel1.Text = "Record Count: " & RecordCount '& DataGridView1.Rows.Count.ToString
    End Sub

    Private Sub ButtonExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonExport.Click
        
        Dim Filename As String = "MyRawDataCurrency-" & Format(DateTime.Today, "yyyyMMdd") & ".xlsx"
        Call ExcelStuff.ExportToExcel(Filename, Dataset2)
    End Sub

    Private Sub DateTimePicker2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker2.ValueChanged, DateTimePicker1.ValueChanged
        applyFilter()
    End Sub

    Private Sub ApplyFilter()
        Dim myArray(0) As String
        Dim i As Integer = 0
        If DateTimePicker1.Checked Then
            ReDim Preserve myArray(0)
            DateTimePicker1.Format = DateTimePickerFormat.Short
            myArray(0) = "Date > =#" & DateTimePicker1.Value.Date & "#"
            i = i + 1

        End If

        If DateTimePicker2.Checked Then
            ReDim Preserve myArray(i)
            DateTimePicker1.Format = DateTimePickerFormat.Short
            myArray(i) = "Date <=#" & DateTimePicker2.Value.Date & "#"
            i = i + 1
        End If
        Dim myfilter As String = Join(myArray, " and ")
        Dataset2.Tables(0).DefaultView.RowFilter = myfilter
        DateTimePicker1.Format = DateTimePickerFormat.Custom
        ShowRecordCount(CurrencyManager1.Count)
    End Sub
End Class