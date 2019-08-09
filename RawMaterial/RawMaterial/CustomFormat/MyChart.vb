Imports DJLib
Imports DJLib.Dbtools
Imports System.Data.Common
Imports System.Text
Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Collections.Generic
Imports System.Linq
Imports Npgsql

Public Class MyChart
    Dim unit As String = vbEmpty
    'Dim DataAdapter1 As DbDataAdapter
    Dim DataAdapter1 As NpgsqlDataAdapter
    Dim mDataset1 As DataSet
    Dim mDataset2 As DataSet
    Dim DataTable1 As DataTable
    Dim DataTable2 As DataTable
    Dim DataTable3 As DataTable
    Dim myCheck As Boolean = False
    Dim myCheck2 As Boolean = False
    Dim mydate() As Date
    Dim myCol As New Collection
    Dim errorMessage As String
    Dim Dbtools1 As New Dbtools(myUserid, myPassword)

    Public Property DataSet1 As DataSet
        Get
            Return mDataset1
        End Get
        Set(ByVal value As DataSet)
            mDataset1 = value
        End Set
    End Property

    Public ReadOnly Property getDataset As DataSet
        Get
            Return mDataset2
        End Get
    End Property


    Private Sub MyChart_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Not DesignMode Then
            Try
                'Call ConnectData()

                Call FillData()
                Call BindingData()
                Dbtools1.GetDataCollection("select date_part('day',dvalue), cvalue from rmparamdt where paramname = 'sheet';", myCol, errorMessage)
                myCheck = True
                Call ComboBox1_SelectedIndexChanged(Me, e)
                myCheck2 = True
            Catch ex As Exception

            End Try

        End If
    End Sub

    Private Sub FillData()
        If Not mDataset1 Is Nothing Then
            Try
                DataTable1 = New DataTable("SheetName")
                DataTable2 = New DataTable("RawMaterial")
                DataTable3 = New DataTable("Currency")
                DataTable1 = mDataset1.Tables(0)
                DataTable2 = mDataset1.Tables(1)
                DataTable3 = mDataset1.Tables(2)
            Catch ex As Exception
            End Try
        End If
    End Sub
    Private Sub BindingData()
        If Not mDataset1 Is Nothing Then
            Try
                ComboBox1.DataSource = DataTable1.DefaultView
                ComboBox1.DisplayMember = "sheetname"
                ComboBox1.ValueMember = "sheetname"
                ComboBox2.DataSource = Nothing
                ComboBox3.DataSource = DataTable3.DefaultView
                ComboBox3.DisplayMember = "crcycode"
                ComboBox3.ValueMember = "crcycode"
            Catch ex As Exception
            End Try
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        If myCheck Then
            Try
                Dim myFilter As String = "sheetname='" & ComboBox1.Text.ToString & "'"
                DataTable2.DefaultView.RowFilter = myFilter
                ComboBox2.DataSource = DataTable2.DefaultView
                ComboBox2.DisplayMember = "namesource" '"rawmaterialname"
                ComboBox2.ValueMember = "keyid"
                myCheck2 = True
                ComboBox2_SelectedIndexChanged(Me, e)
            Catch ex As Exception
            End Try
        End If
    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        If myCheck2 Then
            If Not mDataset1 Is Nothing Then
                Try
                    unit = ""
                    Dim myrow() As DataRow = Nothing
                    myrow = mDataset1.Tables(1).Select("keyid='" + ComboBox2.SelectedValue.ToString & "'")
                    If myrow(0).ItemArray(7).ToString <> "" Then
                        unit = " (" & myrow(0).ItemArray(7).ToString & " )"
                    End If
                Catch ex As Exception
                End Try
                Dim toolTip1 As ToolTip = New ToolTip()
                toolTip1.AutoPopDelay = 0
                toolTip1.InitialDelay = 0
                toolTip1.ReshowDelay = 0
                toolTip1.ShowAlways = True
                toolTip1.SetToolTip(Me.ComboBox2, ComboBox2.Text.ToString())
            End If
        End If
    End Sub
    Private Sub Button1_ClickFactory(ByVal sender As System.Object, ByVal e As System.EventArgs)

        If ComboBox1.Text = "" Then
            MessageBox.Show("Please select ""Material Type"" from list", "Charts")
            ComboBox1.Focus()
            Exit Sub
        End If

        If ComboBox3.Text = "" Then
            MessageBox.Show("Please select ""Currency"" from list", "Charts")
            ComboBox3.Focus()
            Exit Sub
        End If

        Cursor.Current = Cursors.WaitCursor
        Dim DbTools1 As New DJLib.Dbtools(myUserid, myPassword)
        Dim stringbuilder1 As New StringBuilder
        Dim StringBuilder2 As New StringBuilder
        Dim factory As DbProviderFactory = DbProviderFactories.GetFactory(My.Settings.dbProviderFactory)
        Dim conn As DbConnection = factory.CreateConnection
        Dim DataReader As DbDataReader = Nothing
        Dim DataReader2 As DbDataReader = Nothing
        conn.ConnectionString = DbTools1.getConnectionString

        Try
            conn.Open()

            Dim myvalue As Integer = myCol.Item(ComboBox1.Text)

            If Not ((ComboBox1.Text = "EUR") And (ComboBox2.Text = "EUR")) Then
                stringbuilder1 = SBClass.GenSBRM(myvalue, ComboBox3.SelectedValue.ToString, ComboBox2.SelectedValue.ToString, DateTimePicker1.Value, DateTimePicker2.Value)
            Else
                stringbuilder1 = SBClass.GenSBCurrenciesLine(DateTimePicker1.Value, DateTimePicker2.Value)
            End If



            Dim selectcmd1 As DbCommand = factory.CreateCommand
            selectcmd1.Connection = conn
            selectcmd1.CommandText = stringbuilder1.ToString
            selectcmd1.CommandType = CommandType.Text

            mDataset2 = New DataSet
            DbTools1.getDataSet(stringbuilder1.ToString, mDataset2)
            DataReader = selectcmd1.ExecuteReader

            If Not DataReader.HasRows Then
                MessageBox.Show("Data not available", "Charts")
            End If

            Chart1.DataSource = DataReader
            Chart1.Series(0).XValueMember = "Tx Date"
            Chart1.Series(0).YValueMembers = "Tx Amount"
            Chart1.Series(0).LegendText = ComboBox2.Text & unit
            Chart1.DataBind()
            Chart1.ChartAreas.Item(0).AxisX.Interval = 0
            Chart1.ChartAreas.Item(0).AxisY.Title = ComboBox3.Text
            mDataset2.Tables(0).TableName = ComboBox2.Text
            ' Zoom into the X axis
            Chart1.ChartAreas.Item(0).AxisX.ScaleView.Zoom(2, 3)

            ' Enable range selection and zooming end user interface
            Chart1.ChartAreas.Item(0).CursorX.IsUserEnabled = True
            Chart1.ChartAreas.Item(0).CursorX.IsUserSelectionEnabled = True
            Chart1.ChartAreas.Item(0).AxisX.ScaleView.Zoomable = True
            Chart1.ChartAreas.Item(0).AxisX.ScrollBar.IsPositionedInside = True
            Chart1.ChartAreas.Item(0).AxisX.ScaleView.ZoomReset()

            DataReader.Close()

            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
            Label10.Text = ""
            Try
                Dim selector As Func(Of DataPoint, Decimal) = Function(str) str.YValues(0)

                TextBox1.Text = Format(Chart1.Series(0).Points.Max(selector), "#,##0.00")
                TextBox2.Text = Format(Chart1.Series(0).Points.Min(selector), "#,##0.00")
                TextBox3.Text = Format(Chart1.Series(0).Points.Average(selector), "#,##0.00")
                TextBox4.Text = Format((Chart1.Series(0).Points(Chart1.Series(0).Points.Count - 1).YValues(0) / Chart1.Series(0).Points(0).YValues(0)) - 1, "##0.00%")
                Label10.Text = "Figures available  From " & Format(CDate(Chart1.Series(0).Points(0).AxisLabel), "dd-MMM-yyyy") & " to " & Format(CDate(Chart1.Series(0).Points(Chart1.Series(0).Points.Count - 1).AxisLabel), "dd-MMM-yyyy")
                Dim maxvalue = Chart1.Series(0).Points.Max(selector)
                Dim minValue = Chart1.Series(0).Points.Min(selector)
                Chart1.ChartAreas(0).AxisY.Maximum = misc.getRound(maxvalue, 1)
                Chart1.ChartAreas(0).AxisY.Minimum = misc.getRound(minValue, -1)
                If Chart1.ChartAreas(0).AxisY.Maximum - Chart1.ChartAreas(0).AxisY.Minimum = 0 Then
                    Chart1.ChartAreas(0).AxisY.Maximum += 0.01
                    Chart1.ChartAreas(0).AxisY.Minimum -= 0.01
                End If
            Catch ex As Exception

            End Try

            conn.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Charts")
        End Try
        Cursor.Current = Cursors.Default

    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If ComboBox1.Text = "" Then
            MessageBox.Show("Please select ""Material Type"" from list", "Charts")
            ComboBox1.Focus()
            Exit Sub
        End If

        If ComboBox3.Text = "" Then
            MessageBox.Show("Please select ""Currency"" from list", "Charts")
            ComboBox3.Focus()
            Exit Sub
        End If

        Cursor.Current = Cursors.WaitCursor
        Dim DbTools1 As New DJLib.Dbtools(myUserid, myPassword)
        Dim stringbuilder1 As New StringBuilder
        Dim StringBuilder2 As New StringBuilder
        'Dim factory As DbProviderFactory = DbProviderFactories.GetFactory(My.Settings.dbProviderFactory)
        'Dim conn As DbConnection = factory.CreateConnection
        Dim conn = New NpgsqlConnection(DbTools1.getConnectionString)
        Dim DataReader As DbDataReader = Nothing
        Dim DataReader2 As DbDataReader = Nothing
        conn.ConnectionString = DbTools1.getConnectionString

        Try
            conn.Open()

            Dim myvalue As Integer = myCol.Item(ComboBox1.Text)

            If Not ((ComboBox1.Text = "EUR") And (ComboBox2.Text = "EUR")) Then
                stringbuilder1 = SBClass.GenSBRM(myvalue, ComboBox3.SelectedValue.ToString, ComboBox2.SelectedValue.ToString, DateTimePicker1.Value, DateTimePicker2.Value)
            Else
                stringbuilder1 = SBClass.GenSBCurrenciesLine(DateTimePicker1.Value, DateTimePicker2.Value)
            End If


            
            'Dim selectcmd1 As DbCommand = factory.CreateCommand
            Dim selectcmd1 = New NpgsqlCommand()
            selectcmd1.Connection = conn
            selectcmd1.CommandText = stringbuilder1.ToString
            selectcmd1.CommandType = CommandType.Text

            mDataset2 = New DataSet
            DbTools1.getDataSet(stringbuilder1.ToString, mDataset2)
            DataReader = selectcmd1.ExecuteReader

            If Not DataReader.HasRows Then
                MessageBox.Show("Data not available", "Charts")
            End If

            Chart1.DataSource = DataReader
            Chart1.Series(0).XValueMember = "Tx Date"
            Chart1.Series(0).YValueMembers = "Tx Amount"
            Chart1.Series(0).LegendText = ComboBox2.Text & unit
            Chart1.DataBind()
            Chart1.ChartAreas.Item(0).AxisX.Interval = 0
            Chart1.ChartAreas.Item(0).AxisY.Title = ComboBox3.Text
            mDataset2.Tables(0).TableName = ComboBox2.Text
            ' Zoom into the X axis
            Chart1.ChartAreas.Item(0).AxisX.ScaleView.Zoom(2, 3)

            ' Enable range selection and zooming end user interface
            Chart1.ChartAreas.Item(0).CursorX.IsUserEnabled = True
            Chart1.ChartAreas.Item(0).CursorX.IsUserSelectionEnabled = True
            Chart1.ChartAreas.Item(0).AxisX.ScaleView.Zoomable = True
            Chart1.ChartAreas.Item(0).AxisX.ScrollBar.IsPositionedInside = True
            Chart1.ChartAreas.Item(0).AxisX.ScaleView.ZoomReset()

            DataReader.Close()

            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
            Label10.Text = ""
            Try
                Dim selector As Func(Of DataPoint, Decimal) = Function(str) str.YValues(0)

                TextBox1.Text = Format(Chart1.Series(0).Points.Max(selector), "#,##0.00")
                TextBox2.Text = Format(Chart1.Series(0).Points.Min(selector), "#,##0.00")
                TextBox3.Text = Format(Chart1.Series(0).Points.Average(selector), "#,##0.00")
                TextBox4.Text = Format((Chart1.Series(0).Points(Chart1.Series(0).Points.Count - 1).YValues(0) / Chart1.Series(0).Points(0).YValues(0)) - 1, "##0.00%")
                Label10.Text = "Figures available  From " & Format(CDate(Chart1.Series(0).Points(0).AxisLabel), "dd-MMM-yyyy") & " to " & Format(CDate(Chart1.Series(0).Points(Chart1.Series(0).Points.Count - 1).AxisLabel), "dd-MMM-yyyy")
                Dim maxvalue = Chart1.Series(0).Points.Max(selector)
                Dim minValue = Chart1.Series(0).Points.Min(selector)
                Chart1.ChartAreas(0).AxisY.Maximum = misc.getRound(maxvalue, 1)
                Chart1.ChartAreas(0).AxisY.Minimum = misc.getRound(minValue, -1)
                If Chart1.ChartAreas(0).AxisY.Maximum - Chart1.ChartAreas(0).AxisY.Minimum = 0 Then
                    Chart1.ChartAreas(0).AxisY.Maximum += 0.01
                    Chart1.ChartAreas(0).AxisY.Minimum -= 0.01
                End If
            Catch ex As Exception

            End Try

            conn.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Charts")
        End Try
        Cursor.Current = Cursors.Default

    End Sub
    Private Sub Button1_Click1(ByVal sender As System.Object, ByVal e As System.EventArgs)

        If ComboBox1.Text = "" Then
            MessageBox.Show("Please select ""Material Type"" from list", "Charts")
            ComboBox1.Focus()
            Exit Sub
        End If

        If ComboBox3.Text = "" Then
            MessageBox.Show("Please select ""Currency"" from list", "Charts")
            ComboBox3.Focus()
            Exit Sub
        End If

        Cursor.Current = Cursors.WaitCursor
        Dim DbTools1 As New DJLib.Dbtools(myUserid, myPassword)
        Dim stringbuilder1 As New StringBuilder
        Dim StringBuilder2 As New StringBuilder
        Dim factory As DbProviderFactory = DbProviderFactories.GetFactory(My.Settings.dbProviderFactory)
        Dim conn As DbConnection = factory.CreateConnection
        Dim DataReader As DbDataReader = Nothing
        Dim DataReader2 As DbDataReader = Nothing
        conn.ConnectionString = DbTools1.getConnectionString

        Try
            conn.Open()

            Dim myvalue As Integer = myCol.Item(ComboBox1.Text)

            stringbuilder1.Append("select txdate::text as ""Tx Date"",txamount as ""Original Amount"",(case when tx.crcycode = 'EUR' then 1 else c1.crcyamount end ) * " & IIf(ComboBox3.SelectedValue.ToString = "EUR", 1, "c2.crcyamount") & " as conversion ,tx.crcycode  as ""Original Crcy"", txamount / ( case when tx.crcycode = 'EUR' then 1 else c1.crcyamount end )  * " & IIf(ComboBox3.SelectedValue.ToString = "EUR", 1, "c2.crcyamount") & "  as ""Tx Amount"" , '" & ComboBox3.SelectedValue.ToString & "' as ""Crcy""  from rmmaterialrawtx tx ") '   where keyid =  '")

            Select Case myvalue
                Case 1
                    stringbuilder1.Append(" left join rmcrcytx c1 on c1.crcydate = tx.txdate and c1.crcycode= tx.crcycode")
                    stringbuilder1.Append(" left join rmcrcytx c2 on c2.crcydate = tx.txdate and c2.crcycode= '")
                Case 7
                    stringbuilder1.Append(" left join (select avg(crcyamount) as crcyamount, date_part('year',crcydate) as myyear,date_part('week',crcydate) as myweek,crcycode from rmcrcytx group by date_part('week',crcydate),date_part('year',crcydate),crcycode order by date_part('year',crcydate),date_part('week',crcydate)) as c1 on c1.myyear = date_part('year',txdate) and c1.myweek = date_part('week',txdate) and c1.crcycode = tx.crcycode ")
                    stringbuilder1.Append(" left join (select avg(crcyamount) as crcyamount, date_part('year',crcydate) as myyear,date_part('week',crcydate) as myweek,crcycode from rmcrcytx group by date_part('week',crcydate),date_part('year',crcydate),crcycode order by date_part('year',crcydate),date_part('week',crcydate)) as c2 on c2.myyear = date_part('year',txdate) and c2.myweek = date_part('week',txdate) and c2.crcycode = '")
                Case 31
                    stringbuilder1.Append(" left join (select avg(crcyamount) as crcyamount, date_part('year',crcydate) as myyear,date_part('month',crcydate) as mymonth,crcycode from rmcrcytx group by date_part('month',crcydate),date_part('year',crcydate),crcycode order by date_part('year',crcydate),date_part('month',crcydate)) as c1 on c1.myyear = date_part('year',txdate) and c1.mymonth = date_part('month',txdate) and c1.crcycode = tx.crcycode ")
                    stringbuilder1.Append(" left join (select avg(crcyamount) as crcyamount, date_part('year',crcydate) as myyear,date_part('month',crcydate) as mymonth,crcycode from rmcrcytx group by date_part('month',crcydate),date_part('year',crcydate),crcycode order by date_part('year',crcydate),date_part('month',crcydate)) as c2 on c2.myyear = date_part('year',txdate) and c2.mymonth = date_part('month',txdate) and c2.crcycode = '")
            End Select

            stringbuilder1.Append(ComboBox3.SelectedValue.ToString)
            stringbuilder1.Append("'")

            stringbuilder1.Append(" where keyid = '")
            stringbuilder1.Append(ComboBox2.SelectedValue.ToString)
            stringbuilder1.Append("' and txdate >= ")
            stringbuilder1.Append(DateFormatyyyyMMdd(DateTimePicker1.Value))
            stringbuilder1.Append(" and txdate <= ")
            stringbuilder1.Append(DateFormatyyyyMMdd(DateTimePicker2.Value))
            stringbuilder1.Append(" order by txdate")


            Dim selectcmd1 As DbCommand = factory.CreateCommand
            selectcmd1.Connection = conn
            selectcmd1.CommandText = stringbuilder1.ToString
            selectcmd1.CommandType = CommandType.Text

            mDataset2 = New DataSet
            DbTools1.getDataSet(stringbuilder1.ToString, mDataset2)
            DataReader = selectcmd1.ExecuteReader

            If Not DataReader.HasRows Then
                MessageBox.Show("Data not available", "Charts")
            End If

            Chart1.DataSource = DataReader
            Chart1.Series(0).XValueMember = "Tx Date"
            Chart1.Series(0).YValueMembers = "Tx Amount"
            Chart1.Series(0).LegendText = ComboBox2.Text & unit
            Chart1.DataBind()
            Chart1.ChartAreas.Item(0).AxisX.Interval = 0
            Chart1.ChartAreas.Item(0).AxisY.Title = ComboBox3.Text
            ' Zoom into the X axis
            Chart1.ChartAreas.Item(0).AxisX.ScaleView.Zoom(2, 3)

            ' Enable range selection and zooming end user interface
            Chart1.ChartAreas.Item(0).CursorX.IsUserEnabled = True
            Chart1.ChartAreas.Item(0).CursorX.IsUserSelectionEnabled = True
            Chart1.ChartAreas.Item(0).AxisX.ScaleView.Zoomable = True
            Chart1.ChartAreas.Item(0).AxisX.ScrollBar.IsPositionedInside = True
            Chart1.ChartAreas.Item(0).AxisX.ScaleView.ZoomReset()

            DataReader.Close()

            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
            Label10.Text = ""
            Try
                Dim selector As Func(Of DataPoint, Decimal) = Function(str) str.YValues(0)

                TextBox1.Text = Format(Chart1.Series(0).Points.Max(selector), "#,##0.00")
                TextBox2.Text = Format(Chart1.Series(0).Points.Min(selector), "#,##0.00")
                TextBox3.Text = Format(Chart1.Series(0).Points.Average(selector), "#,##0.00")
                TextBox4.Text = Format((Chart1.Series(0).Points(Chart1.Series(0).Points.Count - 1).YValues(0) / Chart1.Series(0).Points(0).YValues(0)), "##0.00%")
                Label10.Text = "Raw Material Figures available  From " & Format(CDate(Chart1.Series(0).Points(0).AxisLabel), "dd-MMM-yyyy") & " to " & Format(CDate(Chart1.Series(0).Points(Chart1.Series(0).Points.Count - 1).AxisLabel), "dd-MMM-yyyy")
            Catch ex As Exception

            End Try

            conn.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Charts")
        End Try
        Cursor.Current = Cursors.Default

    End Sub

End Class
