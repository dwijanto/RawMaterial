Imports System.Data.Common
Imports DJLib
Imports DJLib.Dbtools
Imports System.Windows.Forms.DataVisualization.Charting

Public Class RMCrcy

    Public Event UseAllCurrency(ByRef checked As Boolean)
    Public Event CurrencyChanged(ByRef Crcycode As String)
    Private BaseFirstTime As Boolean = True
    Private Chart1FirstTime As Boolean = True
    Private BaseMax As Decimal = 0
    Private BaseMin As Decimal = 0
    'Private _Dataset2 As DataSet
    'Private _Dataset3 As DataSet

    'Public ReadOnly Property getDataset2() As DataSet
    '    Get
    '        Return _Dataset2
    '    End Get
    'End Property

    'Public ReadOnly Property getDataset3() As DataSet
    '    Get
    '        Return _Dataset3
    '    End Get
    'End Property
    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        RaiseEvent UseAllCurrency(CheckBox2.Checked)
    End Sub

    Private Sub SeriesSelection1_CurrencyChanged(ByVal CrCyCode As String) Handles SeriesSelection1.CurrencyChanged
        RaiseEvent CurrencyChanged(CrCyCode)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        BaseFirstTime = True
        Chart1FirstTime = True
        If Chart1.ChartAreas.Count > 0 Then
            Chart1.ChartAreas.Remove(Chart1.ChartAreas.Item(0))            
        End If
        Chart1.ChartAreas.Add("ChartArea1")
        Chart1.ChartAreas.Item(0).AxisY.IsStartedFromZero = False
        Chart1.ChartAreas.Item(0).AxisY.IsMarginVisible = False
        Chart1.ChartAreas.Item(0).AxisY.LogarithmBase = 10
        Chart1.ChartAreas.Item(0).AxisX.IsLabelAutoFit = True
        Chart1.Legends(0).LegendStyle = LegendStyle.Column
        Chart1.Titles(0).Text = "Material Trends: " & getTitle(DateTimePicker1.Value, DateTimePicker2.Value)
        '
        For i = 0 To Chart1.Series.Count - 1
            Chart1.Series.Remove(Chart1.Series(0))
        Next

        If SeriesSelection1.RoComboBox1.Text <> "" Then
            Call GenerateSeries(SeriesSelection1)
        End If
        If SeriesSelection2.RoComboBox1.Text <> "" Then
            Call GenerateSeries(SeriesSelection2)
        End If
        If SeriesSelection3.RoComboBox1.Text <> "" Then
            Call GenerateSeries(SeriesSelection3)
        End If
        If SeriesSelection4.RoComboBox1.Text <> "" Then
            Call GenerateSeries(SeriesSelection4)
        End If
        If SeriesSelection5.RoComboBox1.Text <> "" Then
            Call GenerateSeries(SeriesSelection5)
        End If
        If SeriesSelection6.RoComboBox1.Text <> "" Then
            Call GenerateSeries(SeriesSelection6)
        End If
    End Sub
    Private Function getTitle(ByRef date1 As Date, ByRef date2 As Date) As String
        Dim mytitle As String = String.Empty
        If date1.Year = date2.Year Then
            mytitle = date1.Year.ToString
        Else
            mytitle = date1.Year & " - " & date2.Year
        End If

        Return mytitle
    End Function
    Private Sub GenerateSeries(ByVal series As seriesSelection)
        If series.RoComboBox1.Text = "CURRENCY" Then
            Call GenerateSeriesCurrency(series)
        Else
            Call GenerateSeriesRM(series)
        End If
    End Sub
    Private Sub GenerateSeriesCurrency1(ByVal series As seriesSelection)

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
            If Not (series.RoComboBox2.Text = "EUR") And (series.RoComboBox3.Text = "EUR") Then
                stringbuilder1.Append("select rm.crcydate::text as ""TX Date""," & IIf(series.RoComboBox2.Text = "EUR", 1, "rm.crcyamount") & " / " & IIf(series.RoComboBox3.Text = "EUR", 1, "rm2.crcyamount") & " as ""Amount"" from rmcrcytx rm")
                stringbuilder1.Append(" left join rmcrcytx rm2 on rm2.crcydate = rm.crcydate and rm2.crcycode = '" & series.RoComboBox3.Text & "'")
                stringbuilder1.Append(" where rm.crcydate>=") '2010-01-01'  
                stringbuilder1.Append(DateFormatyyyyMMdd(DateTimePicker1.Value))
                stringbuilder1.Append(" and rm.crcydate <=")  '2010-10-01' 
                stringbuilder1.Append(DateFormatyyyyMMdd(DateTimePicker2.Value))
                stringbuilder1.Append(" and rm.crcycode = '" & IIf(series.RoComboBox2.Text = "EUR", series.RoComboBox3.Text, series.RoComboBox2.Text) & "'")
                stringbuilder1.Append(" and not " & IIf(series.RoComboBox2.Text = "EUR", 1, "rm.crcyamount") & " / " & IIf(series.RoComboBox3.Text = "EUR", 1, "rm2.crcyamount") & " isnull")
                stringbuilder1.Append(" order by rm.crcydate")
            Else
                stringbuilder1.Append("select crcydate as ""TX Date"",  1 as ""Amount"" from rmcrcytx ")
                stringbuilder1.Append("where crcydate >= ")
                stringbuilder1.Append(DateFormatyyyyMMdd(DateTimePicker1.Value))
                stringbuilder1.Append("and crcydate <= ")
                stringbuilder1.Append(DateFormatyyyyMMdd(DateTimePicker2.Value))
                stringbuilder1.Append("group by crcydate order by ""TX Date""")
            End If

            Dim selectcmd1 As DbCommand = factory.CreateCommand
            selectcmd1.Connection = conn
            selectcmd1.CommandText = stringbuilder1.ToString
            selectcmd1.CommandType = CommandType.Text

            '_Dataset2 = New DataSet
            'DbTools1.getDataSet(stringbuilder1.ToString, _Dataset2)
            DataReader = selectcmd1.ExecuteReader



            If Not DataReader.HasRows Then
                MessageBox.Show("Data not available", "Charts")
            Else

                Dim chartSeries1 As New Series()
                Chart1.Series.Add(chartSeries1)
                Dim mylabel As String = IIf(series.RoComboBox2.Text = series.RoComboBox3.Text, "Line", series.RoComboBox2.Text & " VS " & series.RoComboBox3.Text)



                'create series
                While DataReader.Read
                    chartSeries1.Points.AddXY(CDate(DataReader.Item(0)), DataReader.Item(1))
                End While

                Chart1.Series(Chart1.Series.Count - 1).LegendText = mylabel

                Chart1.Series(Chart1.Series.Count - 1).ChartType = SeriesChartType.Spline  ' SeriesChartType.Line
                Chart1.Series(Chart1.Series.Count - 1).XValueMember = "TX Date"
                Chart1.Series(Chart1.Series.Count - 1).YValueMembers = "Amount"
                Chart1.DataBind()
                Chart1.ChartAreas.Item(0).AxisX.Interval = 0

                ' Zoom into the X axis
                Chart1.ChartAreas.Item(0).AxisX.ScaleView.Zoom(2, 3)

                ' Enable range selection and zooming end user interface
                Chart1.ChartAreas.Item(0).CursorX.IsUserEnabled = True
                Chart1.ChartAreas.Item(0).CursorX.IsUserSelectionEnabled = True
                Chart1.ChartAreas.Item(0).AxisX.ScaleView.Zoomable = True
                Chart1.ChartAreas.Item(0).AxisX.ScrollBar.IsPositionedInside = True
                Chart1.ChartAreas.Item(0).AxisX.ScaleView.ZoomReset()


                DataReader.Close()

                series.TextBox1.Text = ""
                series.TextBox2.Text = ""
                series.TextBox3.Text = ""
                series.TextBox4.Text = ""

                Try
                    Dim selector As Func(Of System.Windows.Forms.DataVisualization.Charting.DataPoint, Decimal) = Function(str) str.YValues(0)

                    series.TextBox1.Text = Format(Chart1.Series(Chart1.Series.Count - 1).Points.Min(selector), "#,##0.00##")
                    series.TextBox2.Text = Format(Chart1.Series(Chart1.Series.Count - 1).Points.Max(selector), "#,##0.00##")
                    series.TextBox3.Text = Format(Chart1.Series(Chart1.Series.Count - 1).Points.Average(selector), "#,##0.00##")
                    series.TextBox4.Text = Format((Chart1.Series(Chart1.Series.Count - 1).Points(Chart1.Series(Chart1.Series.Count - 1).Points.Count - 1).YValues(0) / Chart1.Series(Chart1.Series.Count - 1).Points(0).YValues(0)), "##0.00%")

                    Dim myMax As Double = Chart1.ChartAreas(0).AxisY.Maximum
                    Chart1.ChartAreas.Item(0).AxisY.IsStartedFromZero = True
                    If Not (Chart1FirstTime And myMax = -1) Then
                        Chart1.ChartAreas(0).AxisY.Maximum = Math.Max(myMax, Chart1.Series(Chart1.Series.Count - 1).Points.Max(selector))
                        Dim myMin As Double = 0
                        myMin = IIf(Chart1.ChartAreas(0).AxisY.Minimum = 0, Chart1.Series(Chart1.Series.Count - 1).Points.Min(selector), Chart1.ChartAreas(0).AxisY.Minimum)
                        Dim Validmymin As Double = IIf(myMin = Chart1.ChartAreas(0).AxisY.Maximum, myMin - 1, myMin)
                        Chart1.ChartAreas(0).AxisY.Minimum = Math.Min(Validmymin, Chart1.Series(Chart1.Series.Count - 1).Points.Min(selector))

                    End If
                    Chart1FirstTime = False
                Catch ex As Exception
                End Try

                'using base 100 then
                If CheckBox1.Checked Then
                    Dim FirstPoint As Decimal = Chart1.Series(Chart1.Series.Count - 1).Points(0).YValues(0)
                    Dim basepoint As Decimal = 100 / FirstPoint
                    If BaseFirstTime Then
                        For i = 0 To chartSeries1.Points.Count - 1
                            chartSeries1.Points.Remove(chartSeries1.Points(0))
                        Next
                    End If
                    stringbuilder1.Remove(0, stringbuilder1.Length)
                    If Not (series.RoComboBox2.Text = "EUR") And (series.RoComboBox3.Text = "EUR") Then
                        stringbuilder1.Append("select rm.crcydate::text as ""TX Date"",(" & IIf(series.RoComboBox2.Text = "EUR", 1, "rm.crcyamount") & " / " & IIf(series.RoComboBox3.Text = "EUR", 1, "rm2.crcyamount") & " * " & basepoint & ") as ""Amount"" from rmcrcytx rm")
                        stringbuilder1.Append(" left join rmcrcytx rm2 on rm2.crcydate = rm.crcydate and rm2.crcycode = '" & series.RoComboBox3.Text & "'")
                        stringbuilder1.Append(" where rm.crcydate>=") '2010-01-01'  
                        stringbuilder1.Append(DateFormatyyyyMMdd(DateTimePicker1.Value))
                        stringbuilder1.Append(" and rm.crcydate <=")  '2010-10-01' 
                        stringbuilder1.Append(DateFormatyyyyMMdd(DateTimePicker2.Value))
                        stringbuilder1.Append(" and rm.crcycode = '" & IIf(series.RoComboBox2.Text = "EUR", series.RoComboBox3.Text, series.RoComboBox2.Text) & "'")
                        stringbuilder1.Append(" and not " & IIf(series.RoComboBox2.Text = "EUR", 1, "rm.crcyamount") & " / " & IIf(series.RoComboBox3.Text = "EUR", 1, "rm2.crcyamount") & " isnull")
                        stringbuilder1.Append(" order by rm.crcydate")
                    Else

                        stringbuilder1.Append("select crcydate as ""TX Date"",  100 as ""Amount"" from rmcrcytx ")
                        stringbuilder1.Append("where crcydate >= ")
                        stringbuilder1.Append(DateFormatyyyyMMdd(DateTimePicker1.Value))
                        stringbuilder1.Append("and crcydate <= ")
                        stringbuilder1.Append(DateFormatyyyyMMdd(DateTimePicker2.Value))
                        stringbuilder1.Append("group by crcydate order by ""TX Date""")
                    End If

                    selectcmd1.Connection = conn
                    selectcmd1.CommandText = stringbuilder1.ToString
                    selectcmd1.CommandType = CommandType.Text

                    '_Dataset3 = New DataSet
                    'DbTools1.getDataSet(stringbuilder1.ToString, _Dataset3)
                    DataReader = selectcmd1.ExecuteReader


                    If Not DataReader.HasRows Then
                        MessageBox.Show("Data not available", "Charts")
                    Else


                        For i = 0 To chartSeries1.Points.Count - 1
                            chartSeries1.Points.Remove(chartSeries1.Points(0))
                        Next


                        'create series
                        While DataReader.Read
                            chartSeries1.Points.AddXY(CDate(DataReader.Item(0)), Math.Round(DataReader.Item(1), 1))
                        End While

                        Chart1.Series(Chart1.Series.Count - 1).ChartType = SeriesChartType.Spline 'SeriesChartType.Line
                        Chart1.Series(Chart1.Series.Count - 1).XValueMember = "TX Date"
                        Chart1.Series(Chart1.Series.Count - 1).YValueMembers = "Amount"

                        Chart1.DataBind()
                        Chart1.ChartAreas.Item(0).AxisX.Interval = 0

                        ' Zoom into the X axis
                        Chart1.ChartAreas.Item(0).AxisX.ScaleView.Zoom(2, 3)

                        ' Enable range selection and zooming end user interface
                        Chart1.ChartAreas.Item(0).CursorX.IsUserEnabled = True
                        Chart1.ChartAreas.Item(0).CursorX.IsUserSelectionEnabled = True
                        Chart1.ChartAreas.Item(0).AxisX.ScaleView.Zoomable = True
                        Chart1.ChartAreas.Item(0).AxisX.ScrollBar.IsPositionedInside = True
                        Chart1.ChartAreas.Item(0).AxisX.ScaleView.ZoomReset()

                        Dim selector As Func(Of System.Windows.Forms.DataVisualization.Charting.DataPoint, Decimal) = Function(str) str.YValues(0)


                        If BaseFirstTime Then
                            BaseMax = Chart1.Series(Chart1.Series.Count - 1).Points.Max(selector)
                            BaseMin = Chart1.Series(Chart1.Series.Count - 1).Points.Min(selector)
                            Dim validBaseMin As Decimal = IIf(BaseMin = BaseMax, BaseMin - 1, BaseMin)
                            BaseMin = validBaseMin
                            BaseFirstTime = False
                        End If

                        Chart1.ChartAreas(0).AxisY.Maximum = Math.Max(BaseMax, Chart1.Series(Chart1.Series.Count - 1).Points.Max(selector))
                        Chart1.ChartAreas(0).AxisY.Minimum = Math.Min(BaseMin, Chart1.Series(Chart1.Series.Count - 1).Points.Min(selector))
                        BaseMax = Chart1.ChartAreas(0).AxisY.Maximum
                        BaseMin = Chart1.ChartAreas(0).AxisY.Minimum
                        DataReader.Close()
                        mylabel = IIf(series.RoComboBox2.Text = series.RoComboBox3.Text, "Base100", series.RoComboBox2.Text & " VS " & series.RoComboBox3.Text)
                        Chart1.Series(Chart1.Series.Count - 1).LegendText = mylabel

                    End If
                End If

            End If
            conn.Close()


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            MessageBox.Show(ex.Message, "Charts")
        End Try
        Cursor.Current = Cursors.Default
    End Sub

    Private Sub GenerateSeriesCurrency(ByVal series As seriesSelection)

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
            If Not ((series.RoComboBox2.Text = "EUR") And (series.RoComboBox3.Text = "EUR")) Then
                stringbuilder1 = SBClass.GenSBCurrencies(series.RoComboBox2.Text, series.RoComboBox3.Text, DateTimePicker1.Value, DateTimePicker2.Value)
            Else
                stringbuilder1 = SBClass.GenSBCurrenciesLine(DateTimePicker1.Value, DateTimePicker2.Value)
            End If

            Dim selectcmd1 As DbCommand = factory.CreateCommand
            selectcmd1.Connection = conn
            selectcmd1.CommandText = stringbuilder1.ToString
            selectcmd1.CommandType = CommandType.Text

            '_Dataset2 = New DataSet
            'DbTools1.getDataSet(stringbuilder1.ToString, _Dataset2)
            series.Sqlstr = stringbuilder1.ToString

            DataReader = selectcmd1.ExecuteReader

            If Not DataReader.HasRows Then
                MessageBox.Show("Data not available", "Charts")
            Else

                Dim chartSeries1 As New Series()
                Chart1.Series.Add(chartSeries1)
                Dim mylabel As String = IIf(series.RoComboBox2.Text = series.RoComboBox3.Text, "Line", series.RoComboBox2.Text & " VS " & series.RoComboBox3.Text)
                Chart1.Series(Chart1.Series.Count - 1).LegendText = mylabel
                series.SeriesLegend = mylabel

                'create series
                While DataReader.Read
                    chartSeries1.Points.AddXY(CDate(DataReader.Item(0)), DataReader.Item(1))
                End While


                Chart1.Series(Chart1.Series.Count - 1).ChartType = SeriesChartType.Spline  ' SeriesChartType.Line
                Chart1.Series(Chart1.Series.Count - 1).XValueMember = "TX Date"
                Chart1.Series(Chart1.Series.Count - 1).YValueMembers = "Amount"
                Chart1.DataBind()
                Chart1.ChartAreas.Item(0).AxisX.Interval = 0
                ' Zoom into the X axis
                Chart1.ChartAreas.Item(0).AxisX.ScaleView.Zoom(2, 3)
                ' Enable range selection and zooming end user interface
                Chart1.ChartAreas.Item(0).CursorX.IsUserEnabled = True
                Chart1.ChartAreas.Item(0).CursorX.IsUserSelectionEnabled = True
                Chart1.ChartAreas.Item(0).AxisX.ScaleView.Zoomable = True
                Chart1.ChartAreas.Item(0).AxisX.ScrollBar.IsPositionedInside = True
                Chart1.ChartAreas.Item(0).AxisX.ScaleView.ZoomReset()
                Chart1.ChartAreas.Item(0).AxisX.Interval = 1


                DataReader.Close()

                series.TextBox1.Text = ""
                series.TextBox2.Text = ""
                series.TextBox3.Text = ""
                series.TextBox4.Text = ""

                Try
                    Dim selector As Func(Of System.Windows.Forms.DataVisualization.Charting.DataPoint, Decimal) = Function(str) str.YValues(0)

                    series.TextBox1.Text = Format(Chart1.Series(Chart1.Series.Count - 1).Points.Min(selector), "#,##0.00##")
                    series.TextBox2.Text = Format(Chart1.Series(Chart1.Series.Count - 1).Points.Max(selector), "#,##0.00##")
                    series.TextBox3.Text = Format(Chart1.Series(Chart1.Series.Count - 1).Points.Average(selector), "#,##0.00##")
                    series.TextBox4.Text = Format((Chart1.Series(Chart1.Series.Count - 1).Points(Chart1.Series(Chart1.Series.Count - 1).Points.Count - 1).YValues(0) / Chart1.Series(Chart1.Series.Count - 1).Points(0).YValues(0)), "##0.00%")

                    Dim myMax As Double = Chart1.ChartAreas(0).AxisY.Maximum
                    Chart1.ChartAreas.Item(0).AxisY.IsStartedFromZero = True
                    If Not (Chart1FirstTime And myMax = -1) Then
                        Chart1.ChartAreas(0).AxisY.Maximum = Math.Max(myMax, Chart1.Series(Chart1.Series.Count - 1).Points.Max(selector))
                        Dim myMin As Double = 0
                        myMin = IIf(Chart1.ChartAreas(0).AxisY.Minimum = 0, Chart1.Series(Chart1.Series.Count - 1).Points.Min(selector), Chart1.ChartAreas(0).AxisY.Minimum)
                        Dim Validmymin As Double = IIf(myMin = Chart1.ChartAreas(0).AxisY.Maximum, myMin - 1, myMin)
                        Chart1.ChartAreas(0).AxisY.Minimum = Math.Min(Validmymin, Chart1.Series(Chart1.Series.Count - 1).Points.Min(selector))

                    End If
                    Chart1FirstTime = False
                Catch ex As Exception
                End Try

                'using base 100 then
                If CheckBox1.Checked Then
                    Dim FirstPoint As Decimal = Chart1.Series(Chart1.Series.Count - 1).Points(0).YValues(0)
                    Dim basepoint As Decimal = 100 / FirstPoint
                    If BaseFirstTime Then
                        For i = 0 To chartSeries1.Points.Count - 1
                            chartSeries1.Points.Remove(chartSeries1.Points(0))
                        Next
                    End If
                    stringbuilder1.Remove(0, stringbuilder1.Length)
                    If Not ((series.RoComboBox2.Text = "EUR") And (series.RoComboBox3.Text = "EUR")) Then
                        stringbuilder1 = SBClass.GenSBCurrencies(series.RoComboBox2.Text, series.RoComboBox3.Text, DateTimePicker1.Value, DateTimePicker2.Value, basepoint)
                    Else
                        stringbuilder1 = SBClass.GenSBCurrenciesLine(DateTimePicker1.Value, DateTimePicker2.Value, 100)
                    End If

                    selectcmd1.Connection = conn
                    selectcmd1.CommandText = stringbuilder1.ToString
                    selectcmd1.CommandType = CommandType.Text

                    '_Dataset3 = New DataSet
                    'DbTools1.getDataSet(stringbuilder1.ToString, _Dataset3)

                    DataReader = selectcmd1.ExecuteReader


                    If Not DataReader.HasRows Then
                        MessageBox.Show("Data not available", "Charts")
                    Else


                        For i = 0 To chartSeries1.Points.Count - 1
                            chartSeries1.Points.Remove(chartSeries1.Points(0))
                        Next


                        'create series
                        While DataReader.Read
                            chartSeries1.Points.AddXY(CDate(DataReader.Item(0)), Math.Round(DataReader.Item(1), 1))
                        End While

                        Chart1.Series(Chart1.Series.Count - 1).ChartType = SeriesChartType.Spline 'SeriesChartType.Line
                        Chart1.Series(Chart1.Series.Count - 1).XValueMember = "TX Date"
                        Chart1.Series(Chart1.Series.Count - 1).YValueMembers = "Amount"

                        Chart1.DataBind()
                        Chart1.ChartAreas.Item(0).AxisX.Interval = 1

                        ' Zoom into the X axis
                        Chart1.ChartAreas.Item(0).AxisX.ScaleView.Zoom(2, 3)

                        ' Enable range selection and zooming end user interface
                        Chart1.ChartAreas.Item(0).CursorX.IsUserEnabled = True
                        Chart1.ChartAreas.Item(0).CursorX.IsUserSelectionEnabled = True
                        Chart1.ChartAreas.Item(0).AxisX.ScaleView.Zoomable = True
                        Chart1.ChartAreas.Item(0).AxisX.ScrollBar.IsPositionedInside = True
                        Chart1.ChartAreas.Item(0).AxisX.ScaleView.ZoomReset()

                        Dim selector As Func(Of System.Windows.Forms.DataVisualization.Charting.DataPoint, Decimal) = Function(str) str.YValues(0)


                        If BaseFirstTime Then
                            BaseMax = Chart1.Series(Chart1.Series.Count - 1).Points.Max(selector)
                            BaseMin = Chart1.Series(Chart1.Series.Count - 1).Points.Min(selector)
                            Dim validBaseMin As Decimal = IIf(BaseMin = BaseMax, BaseMin - 1, BaseMin)
                            BaseMin = validBaseMin
                            BaseFirstTime = False
                        End If

                        Chart1.ChartAreas(0).AxisY.Maximum = Math.Max(BaseMax, Chart1.Series(Chart1.Series.Count - 1).Points.Max(selector))
                        Chart1.ChartAreas(0).AxisY.Minimum = Math.Min(BaseMin, Chart1.Series(Chart1.Series.Count - 1).Points.Min(selector))
                        BaseMax = Chart1.ChartAreas(0).AxisY.Maximum
                        BaseMin = Chart1.ChartAreas(0).AxisY.Minimum
                        DataReader.Close()
                        mylabel = IIf(series.RoComboBox2.Text = series.RoComboBox3.Text, "Base100", series.RoComboBox2.Text & " VS " & series.RoComboBox3.Text)
                        Chart1.Series(Chart1.Series.Count - 1).LegendText = mylabel
                        series.SeriesLegend = mylabel
                    End If
                End If

            End If
            conn.Close()


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            MessageBox.Show(ex.Message, "Charts")
        End Try
        Cursor.Current = Cursors.Default
    End Sub

    Private Sub GenerateSeriesRM1(ByVal series As seriesSelection)
        Dim mycol As New Collection
        Dim DbTools1 As New DJLib.Dbtools(myUserid, myPassword)
        Dim stringbuilder1 As New System.Text.StringBuilder
        Dim StringBuilder2 As New System.Text.StringBuilder
        Dim factory As DbProviderFactory = DbProviderFactories.GetFactory(My.Settings.dbProviderFactory)
        Dim conn As DbConnection = factory.CreateConnection
        Dim DataReader As DbDataReader = Nothing
        Dim DataReader2 As DbDataReader = Nothing
        conn.ConnectionString = DbTools1.getConnectionString
        Dim errormessage As String = String.Empty
        Try
            conn.Open()

            DbTools1.GetDataCollection("select date_part('day',dvalue), cvalue from rmparamdt where paramname = 'sheet';", mycol, errormessage)
            Dim myvalue As Integer = mycol.Item(series.RoComboBox1.Text)

            stringbuilder1.Append("select txdate::text as ""Tx Date"",txamount as ""Original Amount"",(case when tx.crcycode = 'EUR' then 1 else c1.crcyamount end ) * " & IIf(series.RoComboBox3.SelectedValue.ToString = "EUR", 1, "c2.crcyamount") & " as conversion ,tx.crcycode  as ""Original Crcy"", txamount / ( case when tx.crcycode = 'EUR' then 1 else c1.crcyamount end )  * " & IIf(series.RoComboBox3.SelectedValue.ToString = "EUR", 1, "c2.crcyamount") & "  as ""Tx Amount"" , '" & series.RoComboBox3.SelectedValue.ToString & "' as ""Crcy""  from rmmaterialrawtx tx ") '   where keyid =  '")

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

            stringbuilder1.Append(series.RoComboBox3.SelectedValue.ToString)
            stringbuilder1.Append("'")

            stringbuilder1.Append(" where keyid = '")
            stringbuilder1.Append(series.RoComboBox2.SelectedValue.ToString)
            stringbuilder1.Append("' and txdate >= ")
            stringbuilder1.Append(DateFormatyyyyMMdd(DateTimePicker1.Value))
            stringbuilder1.Append(" and txdate <= ")
            stringbuilder1.Append(DateFormatyyyyMMdd(DateTimePicker2.Value))
            stringbuilder1.Append(" order by txdate")

            Dim selectcmd1 As DbCommand = factory.CreateCommand
            selectcmd1.Connection = conn
            selectcmd1.CommandText = stringbuilder1.ToString
            selectcmd1.CommandType = CommandType.Text

            '_Dataset2 = New DataSet
            'DbTools1.getDataSet(stringbuilder1.ToString, _Dataset2)
            DataReader = selectcmd1.ExecuteReader



            If Not DataReader.HasRows Then
                MessageBox.Show("Data not available", "Charts")
            Else

                Dim chartSeries1 As New Series()
                Chart1.Series.Add(chartSeries1)
                Chart1.Series(Chart1.Series.Count - 1).LegendText = series.RoComboBox2.Text & " VS " & series.RoComboBox3.Text

                'create series
                While DataReader.Read
                    chartSeries1.Points.AddXY(CDate(DataReader.Item(0)), DataReader.Item(4))
                End While


                Chart1.Series(Chart1.Series.Count - 1).ChartType = SeriesChartType.Spline ' SeriesChartType.Line
                Chart1.Series(Chart1.Series.Count - 1).XValueMember = "Tx Date"
                Chart1.Series(Chart1.Series.Count - 1).YValueMembers = "Tx Amount"
                Chart1.DataBind()
                Chart1.ChartAreas.Item(0).AxisX.Interval = 0

                ' Zoom into the X axis
                Chart1.ChartAreas.Item(0).AxisX.ScaleView.Zoom(2, 3)

                ' Enable range selection and zooming end user interface
                Chart1.ChartAreas.Item(0).CursorX.IsUserEnabled = True
                Chart1.ChartAreas.Item(0).CursorX.IsUserSelectionEnabled = True
                Chart1.ChartAreas.Item(0).AxisX.ScaleView.Zoomable = True
                Chart1.ChartAreas.Item(0).AxisX.ScrollBar.IsPositionedInside = True
                Chart1.ChartAreas.Item(0).AxisX.ScaleView.ZoomReset()

                DataReader.Close()

                series.TextBox1.Text = ""
                series.TextBox2.Text = ""
                series.TextBox3.Text = ""
                series.TextBox4.Text = ""

                Try
                    Dim selector As Func(Of System.Windows.Forms.DataVisualization.Charting.DataPoint, Decimal) = Function(str) str.YValues(0)

                    series.TextBox1.Text = Format(Chart1.Series(Chart1.Series.Count - 1).Points.Min(selector), "#,##0.00##")
                    series.TextBox2.Text = Format(Chart1.Series(Chart1.Series.Count - 1).Points.Max(selector), "#,##0.00##")
                    series.TextBox3.Text = Format(Chart1.Series(Chart1.Series.Count - 1).Points.Average(selector), "#,##0.00##")
                    series.TextBox4.Text = Format((Chart1.Series(Chart1.Series.Count - 1).Points(Chart1.Series(Chart1.Series.Count - 1).Points.Count - 1).YValues(0) / Chart1.Series(Chart1.Series.Count - 1).Points(0).YValues(0)), "##0.00%")

                    Dim myMax As Double = Chart1.ChartAreas(0).AxisY.Maximum

                    Chart1.ChartAreas(0).AxisY.Maximum = Math.Max(myMax, Chart1.Series(Chart1.Series.Count - 1).Points.Max(selector))
                    Dim myMin As Double = 0
                    myMin = IIf(Chart1.ChartAreas(0).AxisY.Minimum = 0, Chart1.Series(Chart1.Series.Count - 1).Points.Min(selector), Chart1.ChartAreas(0).AxisY.Minimum)
                    Chart1.ChartAreas(0).AxisY.Minimum = Math.Min(myMin, Chart1.Series(Chart1.Series.Count - 1).Points.Min(selector))
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try

                'using base 100 then
                If CheckBox1.Checked Then
                    Dim FirstPoint As Decimal = Chart1.Series(Chart1.Series.Count - 1).Points(0).YValues(0)
                    If FirstPoint <> 0 Then
                        Dim basepoint As Decimal = 100 / FirstPoint
                        If BaseFirstTime Then
                            For i = 0 To chartSeries1.Points.Count - 1
                                chartSeries1.Points.Remove(chartSeries1.Points(0))
                            Next
                        End If


                        stringbuilder1.Remove(0, stringbuilder1.Length)


                        stringbuilder1.Append("select txdate::text as ""Tx Date"",txamount as ""Original Amount"",(case when tx.crcycode = 'EUR' then 1 else c1.crcyamount end ) * " & IIf(series.RoComboBox3.SelectedValue.ToString = "EUR", 1, "c2.crcyamount") & " as conversion ,tx.crcycode  as ""Original Crcy"", txamount / ( case when tx.crcycode = 'EUR' then 1 else c1.crcyamount end )  * " & IIf(series.RoComboBox3.SelectedValue.ToString = "EUR", 1, "c2.crcyamount") & " * " & basepoint & "  as ""Tx Amount"" , '" & series.RoComboBox3.SelectedValue.ToString & "' as ""Crcy""  from rmmaterialrawtx tx ") '   where keyid =  '")

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

                        stringbuilder1.Append(series.RoComboBox3.SelectedValue.ToString)
                        stringbuilder1.Append("'")

                        stringbuilder1.Append(" where keyid = '")
                        stringbuilder1.Append(series.RoComboBox2.SelectedValue.ToString)
                        stringbuilder1.Append("' and txdate >= ")
                        stringbuilder1.Append(DateFormatyyyyMMdd(DateTimePicker1.Value))
                        stringbuilder1.Append(" and txdate <= ")
                        stringbuilder1.Append(DateFormatyyyyMMdd(DateTimePicker2.Value))
                        stringbuilder1.Append(" order by txdate")

                        selectcmd1.Connection = conn
                        selectcmd1.CommandText = stringbuilder1.ToString
                        selectcmd1.CommandType = CommandType.Text

                        '_Dataset3 = New DataSet
                        'DbTools1.getDataSet(stringbuilder1.ToString, _Dataset3)
                        DataReader = selectcmd1.ExecuteReader


                        If Not DataReader.HasRows Then
                            MessageBox.Show("Data not available", "Charts")
                        Else


                            For i = 0 To chartSeries1.Points.Count - 1
                                chartSeries1.Points.Remove(chartSeries1.Points(0))
                            Next


                            'create series
                            While DataReader.Read
                                chartSeries1.Points.AddXY(CDate(DataReader.Item(0)), Math.Round(DataReader.Item(4), 1))
                            End While

                            Chart1.Series(Chart1.Series.Count - 1).ChartType = SeriesChartType.Spline  'SeriesChartType.Line
                            Chart1.Series(Chart1.Series.Count - 1).XValueMember = "Tx Date"
                            Chart1.Series(Chart1.Series.Count - 1).YValueMembers = "Tx Amount"

                            Chart1.DataBind()
                            Chart1.ChartAreas.Item(0).AxisX.Interval = 1

                            ' Zoom into the X axis
                            Chart1.ChartAreas.Item(0).AxisX.ScaleView.Zoom(2, 3)

                            ' Enable range selection and zooming end user interface
                            Chart1.ChartAreas.Item(0).CursorX.IsUserEnabled = True
                            Chart1.ChartAreas.Item(0).CursorX.IsUserSelectionEnabled = True
                            Chart1.ChartAreas.Item(0).AxisX.ScaleView.Zoomable = True
                            Chart1.ChartAreas.Item(0).AxisX.ScrollBar.IsPositionedInside = True
                            Chart1.ChartAreas.Item(0).AxisX.ScaleView.ZoomReset()

                            Dim selector As Func(Of System.Windows.Forms.DataVisualization.Charting.DataPoint, Decimal) = Function(str) str.YValues(0)
                            If BaseFirstTime Then
                                BaseMax = Chart1.Series(Chart1.Series.Count - 1).Points.Max(selector)
                                BaseMin = Chart1.Series(Chart1.Series.Count - 1).Points.Min(selector)
                                BaseFirstTime = False
                            End If

                            Chart1.ChartAreas(0).AxisY.Maximum = Math.Max(BaseMax, Chart1.Series(Chart1.Series.Count - 1).Points.Max(selector))
                            Chart1.ChartAreas(0).AxisY.Minimum = Math.Min(BaseMin, Chart1.Series(Chart1.Series.Count - 1).Points.Min(selector))
                            BaseMax = Chart1.ChartAreas(0).AxisY.Maximum
                            BaseMin = Chart1.ChartAreas(0).AxisY.Minimum
                            DataReader.Close()

                        End If
                    End If

                End If
                conn.Close()

            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            MessageBox.Show(ex.Message, "Charts")
        End Try
        Cursor.Current = Cursors.Default
    End Sub
    Private Sub GenerateSeriesRM(ByVal series As seriesSelection)
        Dim mycol As New Collection
        Dim DbTools1 As New DJLib.Dbtools(myUserid, myPassword)
        Dim stringbuilder1 As New System.Text.StringBuilder
        Dim StringBuilder2 As New System.Text.StringBuilder
        Dim factory As DbProviderFactory = DbProviderFactories.GetFactory(My.Settings.dbProviderFactory)
        Dim conn As DbConnection = factory.CreateConnection
        Dim DataReader As DbDataReader = Nothing
        Dim DataReader2 As DbDataReader = Nothing
        conn.ConnectionString = DbTools1.getConnectionString
        Dim errormessage As String = String.Empty
        Try
            conn.Open()

            DbTools1.GetDataCollection("select date_part('day',dvalue), cvalue from rmparamdt where paramname = 'sheet';", mycol, errormessage)
            Dim myvalue As Integer = mycol.Item(series.RoComboBox1.Text)


            stringbuilder1 = SBClass.GenSBRM(myvalue, series.RoComboBox3.SelectedValue.ToString, series.RoComboBox2.SelectedValue.ToString, DateTimePicker1.Value, DateTimePicker2.Value)
            Dim selectcmd1 As DbCommand = factory.CreateCommand
            selectcmd1.Connection = conn
            selectcmd1.CommandText = stringbuilder1.ToString
            selectcmd1.CommandType = CommandType.Text

            '_Dataset2 = New DataSet
            'DbTools1.getDataSet(stringbuilder1.ToString, _Dataset2)
            DataReader = selectcmd1.ExecuteReader
            series.Sqlstr = stringbuilder1.ToString


            If Not DataReader.HasRows Then
                MessageBox.Show("Data not available", "Charts")
            Else

                Dim chartSeries1 As New Series()
                Chart1.Series.Add(chartSeries1)
                'Chart1.Series(Chart1.Series.Count - 1).LegendText = series.RoComboBox2.Text & " VS " & series.RoComboBox3.Text
                'series.SeriesLegend = Chart1.Series(Chart1.Series.Count - 1).LegendText

                'create series
                Dim firstdate As String = Nothing
                Dim lastDate As String = Nothing
                While DataReader.Read
                    chartSeries1.Points.AddXY(CDate(DataReader.Item(0)), DataReader.Item(4))
                    If firstdate Is Nothing Then
                        firstdate = DataReader.Item(0).ToString
                    Else
                        lastDate = DataReader.Item(0).ToString
                    End If
                End While
                Chart1.Series(Chart1.Series.Count - 1).LegendText = series.RoComboBox2.Text & " VS " & series.RoComboBox3.Text & " (Data available from " & lastDate & " To " & firstdate & ")"
                series.SeriesLegend = Chart1.Series(Chart1.Series.Count - 1).LegendText
                Chart1.Series(Chart1.Series.Count - 1).ChartType = SeriesChartType.Spline ' SeriesChartType.Line
                Chart1.Series(Chart1.Series.Count - 1).XValueMember = "Tx Date"
                Chart1.Series(Chart1.Series.Count - 1).YValueMembers = "Tx Amount"
                Chart1.DataBind()

                Chart1.ChartAreas.Item(0).AxisX.Interval = 0

                ' Zoom into the X axis
                Chart1.ChartAreas.Item(0).AxisX.ScaleView.Zoom(2, 3)

                ' Enable range selection and zooming end user interface
                Chart1.ChartAreas.Item(0).CursorX.IsUserEnabled = True
                Chart1.ChartAreas.Item(0).CursorX.IsUserSelectionEnabled = True
                Chart1.ChartAreas.Item(0).AxisX.ScaleView.Zoomable = True
                Chart1.ChartAreas.Item(0).AxisX.ScrollBar.IsPositionedInside = True
                Chart1.ChartAreas.Item(0).AxisX.ScaleView.ZoomReset()
                Chart1.ChartAreas.Item(0).AxisX.Interval = 1
                DataReader.Close()

                series.TextBox1.Text = ""
                series.TextBox2.Text = ""
                series.TextBox3.Text = ""
                series.TextBox4.Text = ""

                Try
                    Dim selector As Func(Of System.Windows.Forms.DataVisualization.Charting.DataPoint, Decimal) = Function(str) str.YValues(0)

                    series.TextBox1.Text = Format(Chart1.Series(Chart1.Series.Count - 1).Points.Min(selector), "#,##0.00##")
                    series.TextBox2.Text = Format(Chart1.Series(Chart1.Series.Count - 1).Points.Max(selector), "#,##0.00##")
                    series.TextBox3.Text = Format(Chart1.Series(Chart1.Series.Count - 1).Points.Average(selector), "#,##0.00##")
                    series.TextBox4.Text = Format((Chart1.Series(Chart1.Series.Count - 1).Points(Chart1.Series(Chart1.Series.Count - 1).Points.Count - 1).YValues(0) / Chart1.Series(Chart1.Series.Count - 1).Points(0).YValues(0)), "##0.00%")

                    Dim myMax As Double = Chart1.ChartAreas(0).AxisY.Maximum

                    Chart1.ChartAreas(0).AxisY.Maximum = Math.Max(myMax, Chart1.Series(Chart1.Series.Count - 1).Points.Max(selector))
                    Dim myMin As Double = 0
                    myMin = IIf(Chart1.ChartAreas(0).AxisY.Minimum = 0, Chart1.Series(Chart1.Series.Count - 1).Points.Min(selector), Chart1.ChartAreas(0).AxisY.Minimum)
                    Chart1.ChartAreas(0).AxisY.Minimum = Math.Min(myMin, Chart1.Series(Chart1.Series.Count - 1).Points.Min(selector))
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try

                'using base 100 then
                If CheckBox1.Checked Then
                    Dim FirstPoint As Decimal = Chart1.Series(Chart1.Series.Count - 1).Points(0).YValues(0)
                    If FirstPoint <> 0 Then
                        Dim basepoint As Decimal = 100 / FirstPoint
                        If BaseFirstTime Then
                            For i = 0 To chartSeries1.Points.Count - 1
                                chartSeries1.Points.Remove(chartSeries1.Points(0))
                            Next
                        End If


                        stringbuilder1.Remove(0, stringbuilder1.Length)
                        'stringbuilder1 = SBClass.GenSBRM(myvalue, series.RoComboBox3.SelectedValue.ToString, series.RoComboBox2.SelectedValue.ToString, DateTimePicker1.Value, DateTimePicker2.Value)
                        stringbuilder1 = SBClass.GenSBRM(myvalue, series.RoComboBox3.SelectedValue.ToString, series.RoComboBox2.SelectedValue.ToString, DateTimePicker1.Value, DateTimePicker2.Value, basepoint)
                       

                        selectcmd1.Connection = conn
                        selectcmd1.CommandText = stringbuilder1.ToString
                        selectcmd1.CommandType = CommandType.Text
                        series.Sqlstr = stringbuilder1.ToString
                        '_Dataset3 = New DataSet
                        'DbTools1.getDataSet(stringbuilder1.ToString, _Dataset3)
                        DataReader = selectcmd1.ExecuteReader


                        If Not DataReader.HasRows Then
                            MessageBox.Show("Data not available", "Charts")
                        Else


                            For i = 0 To chartSeries1.Points.Count - 1
                                chartSeries1.Points.Remove(chartSeries1.Points(0))
                            Next


                            'create series
                            While DataReader.Read
                                'chartSeries1.Points.AddXY(CDate(DataReader.Item(0)), Math.Round(DataReader.Item(4), 1))
                                chartSeries1.Points.AddXY(CDate(DataReader.Item(0)), DataReader.Item(5))
                            End While

                            Chart1.Series(Chart1.Series.Count - 1).ChartType = SeriesChartType.Spline  'SeriesChartType.Line
                            Chart1.Series(Chart1.Series.Count - 1).XValueMember = "Tx Date"
                            Chart1.Series(Chart1.Series.Count - 1).YValueMembers = "Tx Amount"

                            Chart1.DataBind()
                            Chart1.ChartAreas.Item(0).AxisX.Interval = 1

                            ' Zoom into the X axis
                            Chart1.ChartAreas.Item(0).AxisX.ScaleView.Zoom(2, 3)

                            ' Enable range selection and zooming end user interface
                            Chart1.ChartAreas.Item(0).CursorX.IsUserEnabled = True
                            Chart1.ChartAreas.Item(0).CursorX.IsUserSelectionEnabled = True
                            Chart1.ChartAreas.Item(0).AxisX.ScaleView.Zoomable = True
                            Chart1.ChartAreas.Item(0).AxisX.ScrollBar.IsPositionedInside = True
                            Chart1.ChartAreas.Item(0).AxisX.ScaleView.ZoomReset()

                            Dim selector As Func(Of System.Windows.Forms.DataVisualization.Charting.DataPoint, Decimal) = Function(str) str.YValues(0)
                            If BaseFirstTime Then
                                BaseMax = misc.getRound(Chart1.Series(Chart1.Series.Count - 1).Points.Max(selector), 1)
                                BaseMin = misc.getRound(Chart1.Series(Chart1.Series.Count - 1).Points.Min(selector), -1)
                                BaseFirstTime = False
                            End If

                            Chart1.ChartAreas(0).AxisY.Maximum = Math.Max(BaseMax, misc.getRound(Chart1.Series(Chart1.Series.Count - 1).Points.Max(selector), 1))
                            Chart1.ChartAreas(0).AxisY.Minimum = Math.Min(BaseMin, misc.getRound(Chart1.Series(Chart1.Series.Count - 1).Points.Min(selector), -1))
                            BaseMax = Chart1.ChartAreas(0).AxisY.Maximum
                            BaseMin = Chart1.ChartAreas(0).AxisY.Minimum
                            DataReader.Close()

                        End If
                    End If

                End If
                conn.Close()

            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            MessageBox.Show(ex.Message, "Charts")
        End Try
        Cursor.Current = Cursors.Default
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        SeriesSelection1.RoComboBox1.Text = ""
        SeriesSelection2.RoComboBox1.Text = ""
        SeriesSelection3.RoComboBox1.Text = ""
        SeriesSelection4.RoComboBox1.Text = ""
        SeriesSelection5.RoComboBox1.Text = ""
        SeriesSelection6.RoComboBox1.Text = ""
        Button1_Click(Me, e)
        Chart1.Titles(0).Text = ""
    End Sub
End Class
