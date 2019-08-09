Imports DJLib
Imports DJLib.Dbtools
Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Data.Common
Imports Npgsql

Public Class crcyChart
    Private _Dataset As DataSet
    Private Dataset2 As DataSet
    Private DataTable1 As DataTable
    Private DataTable2 As DataTable
    Public Property Dataset() As DataSet
        Get
            Return _Dataset
        End Get
        Set(ByVal value As DataSet)
            _Dataset = value
        End Set
    End Property

    Public ReadOnly Property getDataset As DataSet
        Get
            Return Dataset2
        End Get
    End Property
    Private Sub crcyChart_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not DesignMode Then
            Try
                Call FillData()
                Call BindingData()
            Catch ex As Exception

            End Try
        End If
    End Sub
    Private Sub FillData()
        If Not _Dataset Is Nothing Then
            Try
                DataTable1 = New DataTable("Crcy1")
                DataTable2 = New DataTable("Crcy2")

                DataTable1 = _Dataset.Tables(0)
                DataTable2 = _Dataset.Tables(1)
            Catch ex As Exception
            End Try
        End If
    End Sub
    Private Sub BindingData()
        If Not _Dataset Is Nothing Then
            Try
                RoComboBox1.DataSource = DataTable1.DefaultView
                RoComboBox1.DisplayMember = "crcycode"
                RoComboBox1.ValueMember = "crcycode"
                RoComboBox2.DataSource = DataTable2.DefaultView
                RoComboBox2.DisplayMember = "crcycode"
                RoComboBox2.ValueMember = "crcycode"
            Catch ex As Exception

            End Try
        End If

    End Sub

    Private Sub Button1_ClickFactory(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Cursor.Current = Cursors.WaitCursor
        Dim DbTools1 As New DJLib.Dbtools(myUserid, myPassword)
        Dim stringbuilder1 As New System.Text.StringBuilder
        Dim StringBuilder2 As New System.Text.StringBuilder
        Dim factory As DbProviderFactory = DbProviderFactories.GetFactory(My.Settings.dbProviderFactory)
        Dim conn As DbConnection = factory.CreateConnection
        Dim DataReader As DbDataReader = Nothing
        Dim DataReader2 As DbDataReader = Nothing
        conn.ConnectionString = DbTools1.getConnectionString
        Try
            conn.Open()
            If Not ((RoComboBox1.Text = "EUR") And (RoComboBox2.Text = "EUR")) Then
                stringbuilder1 = SBClass.GenSBCurrencies(RoComboBox1.Text, RoComboBox2.Text, DateTimePicker1.Value, DateTimePicker2.Value)
            Else
                stringbuilder1 = SBClass.GenSBCurrenciesLine(DateTimePicker1.Value, DateTimePicker2.Value)
            End If


            Dim selectcmd1 As DbCommand = factory.CreateCommand
            selectcmd1.Connection = conn
            selectcmd1.CommandText = stringbuilder1.ToString
            selectcmd1.CommandType = CommandType.Text

            Dataset2 = New DataSet
            DbTools1.getDataSet(stringbuilder1.ToString, Dataset2)
            DataReader = selectcmd1.ExecuteReader

            If Not DataReader.HasRows Then
                MessageBox.Show("Data not available", "Charts")

            End If

            Chart1.DataSource = DataReader
            Chart1.Series(0).XValueMember = "TX Date"
            Chart1.Series(0).YValueMembers = "Amount"
            Chart1.Series(Chart1.Series.Count - 1).ChartType = SeriesChartType.Spline

            Chart1.Series(0).LegendText = RoComboBox2.Text
            Dataset2.Tables(0).TableName = RoComboBox1.Text & " VS " & RoComboBox2.Text
            Chart1.DataBind()
            Chart1.ChartAreas.Item(0).AxisX.Interval = 0
            Chart1.ChartAreas.Item(0).AxisY.Title = RoComboBox1.Text
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

                TextBox1.Text = Format(Chart1.Series(0).Points.Max(selector), "#,##0.0000")
                TextBox2.Text = Format(Chart1.Series(0).Points.Min(selector), "#,##0.0000")
                TextBox3.Text = Format(Chart1.Series(0).Points.Average(selector), "#,##0.0000")
                TextBox4.Text = Format((Chart1.Series(0).Points(Chart1.Series(0).Points.Count - 1).YValues(0) / Chart1.Series(0).Points(0).YValues(0)) - 1, "##0.00%")
                Label10.Text = "Currencies Figures available From " & Format(CDate(Chart1.Series(0).Points(0).AxisLabel), "dd-MMM-yyyy") & " to " & Format(CDate(Chart1.Series(0).Points(Chart1.Series(0).Points.Count - 1).AxisLabel), "dd-MMM-yyyy")

                'get round
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

        Cursor.Current = Cursors.WaitCursor
        Dim DbTools1 As New DJLib.Dbtools(myUserid, myPassword)
        Dim stringbuilder1 As New System.Text.StringBuilder
        Dim StringBuilder2 As New System.Text.StringBuilder
        'Dim factory As DbProviderFactory = DbProviderFactories.GetFactory(My.Settings.dbProviderFactory)
        Dim conn = New NpgsqlConnection(DbTools1.getConnectionString)
        'Dim conn As DbConnection = factory.CreateConnection
        Dim DataReader As DbDataReader = Nothing
        Dim DataReader2 As DbDataReader = Nothing
        conn.ConnectionString = DbTools1.getConnectionString
        Try
            conn.Open()
            If Not ((RoComboBox1.Text = "EUR") And (RoComboBox2.Text = "EUR")) Then
                stringbuilder1 = SBClass.GenSBCurrencies(RoComboBox1.Text, RoComboBox2.Text, DateTimePicker1.Value, DateTimePicker2.Value)
            Else
                stringbuilder1 = SBClass.GenSBCurrenciesLine(DateTimePicker1.Value, DateTimePicker2.Value)
            End If


            'Dim selectcmd1 As DbCommand = factory.CreateCommand
            Dim selectcmd1 = New NpgsqlCommand()
            selectcmd1.Connection = conn
            selectcmd1.CommandText = stringbuilder1.ToString
            selectcmd1.CommandType = CommandType.Text

            Dataset2 = New DataSet
            DbTools1.getDataSet(stringbuilder1.ToString, Dataset2)
            DataReader = selectcmd1.ExecuteReader

            If Not DataReader.HasRows Then
                MessageBox.Show("Data not available", "Charts")

            End If

            Chart1.DataSource = DataReader
            Chart1.Series(0).XValueMember = "TX Date"
            Chart1.Series(0).YValueMembers = "Amount"
            Chart1.Series(Chart1.Series.Count - 1).ChartType = SeriesChartType.Spline

            Chart1.Series(0).LegendText = RoComboBox2.Text
            Dataset2.Tables(0).TableName = RoComboBox1.Text & " VS " & RoComboBox2.Text
            Chart1.DataBind()
            Chart1.ChartAreas.Item(0).AxisX.Interval = 0
            Chart1.ChartAreas.Item(0).AxisY.Title = RoComboBox1.Text
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

                TextBox1.Text = Format(Chart1.Series(0).Points.Max(selector), "#,##0.0000")
                TextBox2.Text = Format(Chart1.Series(0).Points.Min(selector), "#,##0.0000")
                TextBox3.Text = Format(Chart1.Series(0).Points.Average(selector), "#,##0.0000")
                TextBox4.Text = Format((Chart1.Series(0).Points(Chart1.Series(0).Points.Count - 1).YValues(0) / Chart1.Series(0).Points(0).YValues(0)) - 1, "##0.00%")
                Label10.Text = "Currencies Figures available From " & Format(CDate(Chart1.Series(0).Points(0).AxisLabel), "dd-MMM-yyyy") & " to " & Format(CDate(Chart1.Series(0).Points(Chart1.Series(0).Points.Count - 1).AxisLabel), "dd-MMM-yyyy")

                'get round
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

        Cursor.Current = Cursors.WaitCursor
        Dim DbTools1 As New DJLib.Dbtools(myUserid, myPassword)
        Dim stringbuilder1 As New System.Text.StringBuilder
        Dim StringBuilder2 As New System.Text.StringBuilder
        Dim factory As DbProviderFactory = DbProviderFactories.GetFactory(My.Settings.dbProviderFactory)
        Dim conn As DbConnection = factory.CreateConnection
        Dim DataReader As DbDataReader = Nothing
        Dim DataReader2 As DbDataReader = Nothing
        conn.ConnectionString = DbTools1.getConnectionString
        Try
            conn.Open()

            stringbuilder1.Append("select rm.crcydate::text as ""TX Date""," & IIf(RoComboBox1.Text = "EUR", 1, "rm.crcyamount") & " / " & IIf(RoComboBox2.Text = "EUR", 1, "rm2.crcyamount") & " as ""Amount"",rm.crcycode from rmcrcytx rm")
            stringbuilder1.Append(" left join rmcrcytx rm2 on rm2.crcydate = rm.crcydate and rm2.crcycode = '" & RoComboBox2.Text & "'")
            stringbuilder1.Append(" where rm.crcydate>=") '2010-01-01'  
            stringbuilder1.Append(DateFormatyyyyMMdd(DateTimePicker1.Value))
            stringbuilder1.Append(" and rm.crcydate <=")  '2010-10-01' 
            stringbuilder1.Append(DateFormatyyyyMMdd(DateTimePicker2.Value))
            stringbuilder1.Append(" and rm.crcycode = '" & IIf(RoComboBox1.Text = "EUR", RoComboBox2.Text, RoComboBox1.Text) & "'")
            stringbuilder1.Append(" and not " & IIf(RoComboBox1.Text = "EUR", 1, "rm.crcyamount") & " / " & IIf(RoComboBox2.Text = "EUR", 1, "rm2.crcyamount") & " isnull")
            stringbuilder1.Append(" order by rm.crcydate")

            Dim selectcmd1 As DbCommand = factory.CreateCommand
            selectcmd1.Connection = conn
            selectcmd1.CommandText = stringbuilder1.ToString
            selectcmd1.CommandType = CommandType.Text

            Dataset2 = New DataSet
            DbTools1.getDataSet(stringbuilder1.ToString, Dataset2)
            DataReader = selectcmd1.ExecuteReader

            If Not DataReader.HasRows Then
                MessageBox.Show("Data not available", "Charts")

            End If

            Chart1.DataSource = DataReader
            Chart1.Series(0).XValueMember = "TX Date"
            Chart1.Series(0).YValueMembers = "Amount"
            Chart1.Series(0).LegendText = RoComboBox1.Text
            Chart1.DataBind()
            Chart1.ChartAreas.Item(0).AxisX.Interval = 0
            Chart1.ChartAreas.Item(0).AxisY.Title = RoComboBox2.Text
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
                Label10.Text = "Currencies Figures available From " & Format(CDate(Chart1.Series(0).Points(0).AxisLabel), "dd-MMM-yyyy") & " to " & Format(CDate(Chart1.Series(0).Points(Chart1.Series(0).Points.Count - 1).AxisLabel), "dd-MMM-yyyy")
            Catch ex As Exception

            End Try

            conn.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Charts")
        End Try
        Cursor.Current = Cursors.Default

    End Sub

    
End Class
