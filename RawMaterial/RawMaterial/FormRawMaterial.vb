Imports DJLib
Imports DJLib.Dbtools
Imports Microsoft.Office.Interop
Public Class FormRawMaterial
    Dim DataSet1 As DataSet
    Dim CurrencyManager1 As CurrencyManager
    Dim Bindingsource1 As New BindingSource


    Private Sub FormRawMaterial_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load        
        Application.DoEvents() 'we need this to refresh form main menu
        Cursor.Current = Cursors.WaitCursor
        DataSet1 = New DataSet
        Dim DbTools1 As New DJLib.Dbtools(myUserid, myPassword)
        Dim sqlstr As String = " SELECT m.sheetname as ""Sheet Name"", mtx.txdate as ""Tx Date"", mtx.crcycode as ""Crcy"", m.rawmaterialname as ""Material Name"",mtx.txamount as ""Amount"", mtx.unit as ""Unit"",m.source as ""Source"", m.links as ""Links"",m.area as ""Area"", m.keyid as ""Key Id"", m.updateby as ""Update By"",  m.latestupdate as ""Last Update"",  mtx.materialrawtxid as id " & _
                               " from  rmmaterialraw m " & _
                               " LEFT JOIN rmmaterialrawtx mtx ON mtx.keyid::text = m.keyid::text order by ""Sheet Name"",""Tx Date""; select '' as sheetname union all (select sheetname from rmmaterialraw group by sheetname order by sheetname);"
        DbTools1.getDataSet(sqlstr, DataSet1)
        'filldata
        BindingData()
        Cursor.Current = Cursors.Default
    End Sub
    Private Sub FillData()

    End Sub
    Private Sub BindingData()
        Try
            ComboBox1.DataSource = DataSet1.Tables(1).DefaultView
            ComboBox1.DisplayMember = "sheetname"
            ComboBox1.ValueMember = "sheetname"



            Bindingsource1.DataSource = DataSet1.Tables(0).DefaultView 'DataTable1 'dataview1
            DataGridView1.DataSource = Bindingsource1
            With DataGridView1
                .RowsDefaultCellStyle.BackColor = Color.Bisque
                .AlternatingRowsDefaultCellStyle.BackColor = Color.White
            End With
            Me.DataGridView1.Columns("Tx Date").DefaultCellStyle.Format = "dd-MM-yyyy"
            ComboBox2.Items.Clear()
            For i As Integer = 0 To DataGridView1.Columns.Count - 1
                Me.DataGridView1.Columns("Material Name").Width = 80
                ComboBox2.Items.Add(Me.DataGridView1.Columns(i).Name.ToString)
            Next

            Me.DataGridView1.Columns("Material Name").Width = 400
            Me.DataGridView1.Columns("Links").Width = 350
            Me.DataGridView1.Columns("Source").Width = 350
            CurrencyManager1 = CType(Me.BindingContext(Bindingsource1), CurrencyManager)
            ShowRecordCount()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Cursor.Current = Cursors.WaitCursor

        FormRawMaterial_Load(Me, e)
        Cursor.Current = Cursors.Default
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.Text <> "" Then
            DataSet1.Tables(0).DefaultView.RowFilter = "[Sheet Name] =" & escapeString(ComboBox1.SelectedValue.ToString)
        Else
            DataSet1.Tables(0).DefaultView.RowFilter = ""
        End If
        ShowRecordCount()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        Try
            If ComboBox2.Text <> "" Then
                DataSet1.Tables(0).DefaultView.Sort = ComboBox2.Text
            End If
        Catch ex As Exception

        End Try
        
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Dim Additionalfilter As String = ""

        Dim CurrentFilter As String
        If ComboBox1.Text <> "" Then
            CurrentFilter = "[Sheet Name] =" & escapeString(ComboBox1.SelectedValue.ToString)
        Else
            CurrentFilter = ""
        End If
        Try
            If TextBox1.Text = "" Then
                DataSet1.Tables(0).DefaultView.RowFilter = CurrentFilter
            Else
                If CurrentFilter <> "" Then
                    CurrentFilter = CurrentFilter & " and "
                End If
                DataSet1.Tables(0).DefaultView.RowFilter = CurrentFilter & "[" & ComboBox2.Text & "] like " & escapeString("*" & TextBox1.Text & "*")
            End If

        Catch ex As Exception

        End Try
        ShowRecordCount()
    End Sub
    Private Sub ShowRecordCount()
        Me.ToolStripStatusLabel1.Text = "Records Count: " & DataGridView1.Rows.Count.ToString
    End Sub

    Private Sub ButtonExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonExport.Click
        
        Dim Filename As String = "MyRawData-" & Format(DateTime.Today, "yyyyMMdd") & ".xlsx"
        ExcelStuff.ExportToExcel(Filename, DataSet1)
    End Sub

End Class