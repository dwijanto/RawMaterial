<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class RMCrcy
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim ChartArea1 As System.Windows.Forms.DataVisualization.Charting.ChartArea = New System.Windows.Forms.DataVisualization.Charting.ChartArea()
        Dim Legend1 As System.Windows.Forms.DataVisualization.Charting.Legend = New System.Windows.Forms.DataVisualization.Charting.Legend()
        Dim Series1 As System.Windows.Forms.DataVisualization.Charting.Series = New System.Windows.Forms.DataVisualization.Charting.Series()
        Dim Title1 As System.Windows.Forms.DataVisualization.Charting.Title = New System.Windows.Forms.DataVisualization.Charting.Title()
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker()
        Me.DateTimePicker2 = New System.Windows.Forms.DateTimePicker()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.CheckBox2 = New System.Windows.Forms.CheckBox()
        Me.Chart1 = New System.Windows.Forms.DataVisualization.Charting.Chart()
        Me.SeriesSelection6 = New RawMaterial.seriesSelection()
        Me.SeriesSelection5 = New RawMaterial.seriesSelection()
        Me.SeriesSelection4 = New RawMaterial.seriesSelection()
        Me.SeriesSelection3 = New RawMaterial.seriesSelection()
        Me.SeriesSelection2 = New RawMaterial.seriesSelection()
        Me.SeriesSelection1 = New RawMaterial.seriesSelection()
        Me.Panel1.SuspendLayout()
        CType(Me.Chart1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.CustomFormat = "dd-MMM-yyyy"
        Me.DateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePicker1.Location = New System.Drawing.Point(71, 12)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(118, 20)
        Me.DateTimePicker1.TabIndex = 8
        '
        'DateTimePicker2
        '
        Me.DateTimePicker2.CustomFormat = "dd-MMM-yyyy"
        Me.DateTimePicker2.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePicker2.Location = New System.Drawing.Point(260, 12)
        Me.DateTimePicker2.Name = "DateTimePicker2"
        Me.DateTimePicker2.Size = New System.Drawing.Size(118, 20)
        Me.DateTimePicker2.TabIndex = 9
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Checked = True
        Me.CheckBox1.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox1.Location = New System.Drawing.Point(462, 15)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(71, 17)
        Me.CheckBox1.TabIndex = 13
        Me.CheckBox1.Text = "Base 100"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Lavender
        Me.Panel1.Controls.Add(Me.Button2)
        Me.Panel1.Controls.Add(Me.Button1)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.CheckBox2)
        Me.Panel1.Controls.Add(Me.DateTimePicker2)
        Me.Panel1.Controls.Add(Me.DateTimePicker1)
        Me.Panel1.Controls.Add(Me.CheckBox1)
        Me.Panel1.Location = New System.Drawing.Point(3, 3)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1035, 47)
        Me.Panel1.TabIndex = 102
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(726, 6)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(150, 33)
        Me.Button2.TabIndex = 111
        Me.Button2.Text = "Clear Selections && Chart"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(882, 6)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(138, 33)
        Me.Button1.TabIndex = 110
        Me.Button1.Text = "Update Chart"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(565, 15)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(52, 13)
        Me.Label3.TabIndex = 17
        Me.Label3.Text = "Currency:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(199, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(52, 13)
        Me.Label2.TabIndex = 16
        Me.Label2.Text = "End Date"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(10, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(55, 13)
        Me.Label1.TabIndex = 15
        Me.Label1.Text = "Start Date"
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Checked = True
        Me.CheckBox2.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox2.Location = New System.Drawing.Point(623, 14)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(81, 17)
        Me.CheckBox2.TabIndex = 14
        Me.CheckBox2.Text = "Same for all"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'Chart1
        '
        Me.Chart1.BackColor = System.Drawing.Color.Beige
        Me.Chart1.BorderlineColor = System.Drawing.Color.DarkOrange
        Me.Chart1.BorderlineDashStyle = System.Windows.Forms.DataVisualization.Charting.ChartDashStyle.Solid
        Me.Chart1.BorderlineWidth = 2
        Me.Chart1.BorderSkin.BorderDashStyle = System.Windows.Forms.DataVisualization.Charting.ChartDashStyle.DashDot
        Me.Chart1.BorderSkin.PageColor = System.Drawing.SystemColors.ControlLightLight
        Me.Chart1.BorderSkin.SkinStyle = System.Windows.Forms.DataVisualization.Charting.BorderSkinStyle.Emboss
        ChartArea1.AxisX.LabelAutoFitStyle = CType(((((((System.Windows.Forms.DataVisualization.Charting.LabelAutoFitStyles.IncreaseFont Or System.Windows.Forms.DataVisualization.Charting.LabelAutoFitStyles.DecreaseFont) _
                    Or System.Windows.Forms.DataVisualization.Charting.LabelAutoFitStyles.StaggeredLabels) _
                    Or System.Windows.Forms.DataVisualization.Charting.LabelAutoFitStyles.LabelsAngleStep30) _
                    Or System.Windows.Forms.DataVisualization.Charting.LabelAutoFitStyles.LabelsAngleStep45) _
                    Or System.Windows.Forms.DataVisualization.Charting.LabelAutoFitStyles.LabelsAngleStep90) _
                    Or System.Windows.Forms.DataVisualization.Charting.LabelAutoFitStyles.WordWrap), System.Windows.Forms.DataVisualization.Charting.LabelAutoFitStyles)
        ChartArea1.AxisY.IsMarginVisible = False
        ChartArea1.AxisY.IsStartedFromZero = False
        ChartArea1.Name = "ChartArea1"
        Me.Chart1.ChartAreas.Add(ChartArea1)
        Legend1.Alignment = System.Drawing.StringAlignment.Center
        Legend1.BackColor = System.Drawing.Color.Beige
        Legend1.Docking = System.Windows.Forms.DataVisualization.Charting.Docking.Bottom
        Legend1.LegendStyle = System.Windows.Forms.DataVisualization.Charting.LegendStyle.Column
        Legend1.Name = "Legend1"
        Legend1.TextWrapThreshold = 100
        Me.Chart1.Legends.Add(Legend1)
        Me.Chart1.Location = New System.Drawing.Point(3, 305)
        Me.Chart1.Name = "Chart1"
        Me.Chart1.Palette = System.Windows.Forms.DataVisualization.Charting.ChartColorPalette.None
        Series1.BorderWidth = 2
        Series1.ChartArea = "ChartArea1"
        Series1.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line
        Series1.Color = System.Drawing.Color.DarkOliveGreen
        Series1.Legend = "Legend1"
        Series1.LegendText = " "
        Series1.Name = "Series1"
        Series1.ToolTip = "Value: #VAL{#,##0.00} #LABEL \nDate : #AXISLABEL"
        Me.Chart1.Series.Add(Series1)
        Me.Chart1.Size = New System.Drawing.Size(1032, 377)
        Me.Chart1.TabIndex = 109
        Me.Chart1.Text = "Chart1"
        Title1.Name = "MaterialTrends"
        Me.Chart1.Titles.Add(Title1)
        '
        'SeriesSelection6
        '
        Me.SeriesSelection6.DataSet = Nothing
        Me.SeriesSelection6.DataSet1 = Nothing
        Me.SeriesSelection6.DataSet2 = Nothing
        Me.SeriesSelection6.Location = New System.Drawing.Point(521, 222)
        Me.SeriesSelection6.Name = "SeriesSelection6"
        Me.SeriesSelection6.SeriesLegend = Nothing
        Me.SeriesSelection6.Size = New System.Drawing.Size(517, 77)
        Me.SeriesSelection6.Sqlstr = Nothing
        Me.SeriesSelection6.TabIndex = 108
        '
        'SeriesSelection5
        '
        Me.SeriesSelection5.DataSet = Nothing
        Me.SeriesSelection5.DataSet1 = Nothing
        Me.SeriesSelection5.DataSet2 = Nothing
        Me.SeriesSelection5.Location = New System.Drawing.Point(3, 222)
        Me.SeriesSelection5.Name = "SeriesSelection5"
        Me.SeriesSelection5.SeriesLegend = Nothing
        Me.SeriesSelection5.Size = New System.Drawing.Size(517, 77)
        Me.SeriesSelection5.Sqlstr = Nothing
        Me.SeriesSelection5.TabIndex = 107
        '
        'SeriesSelection4
        '
        Me.SeriesSelection4.DataSet = Nothing
        Me.SeriesSelection4.DataSet1 = Nothing
        Me.SeriesSelection4.DataSet2 = Nothing
        Me.SeriesSelection4.Location = New System.Drawing.Point(521, 139)
        Me.SeriesSelection4.Name = "SeriesSelection4"
        Me.SeriesSelection4.SeriesLegend = Nothing
        Me.SeriesSelection4.Size = New System.Drawing.Size(517, 77)
        Me.SeriesSelection4.Sqlstr = Nothing
        Me.SeriesSelection4.TabIndex = 106
        '
        'SeriesSelection3
        '
        Me.SeriesSelection3.DataSet = Nothing
        Me.SeriesSelection3.DataSet1 = Nothing
        Me.SeriesSelection3.DataSet2 = Nothing
        Me.SeriesSelection3.Location = New System.Drawing.Point(3, 139)
        Me.SeriesSelection3.Name = "SeriesSelection3"
        Me.SeriesSelection3.SeriesLegend = Nothing
        Me.SeriesSelection3.Size = New System.Drawing.Size(517, 77)
        Me.SeriesSelection3.Sqlstr = Nothing
        Me.SeriesSelection3.TabIndex = 105
        '
        'SeriesSelection2
        '
        Me.SeriesSelection2.DataSet = Nothing
        Me.SeriesSelection2.DataSet1 = Nothing
        Me.SeriesSelection2.DataSet2 = Nothing
        Me.SeriesSelection2.Location = New System.Drawing.Point(521, 56)
        Me.SeriesSelection2.Name = "SeriesSelection2"
        Me.SeriesSelection2.SeriesLegend = Nothing
        Me.SeriesSelection2.Size = New System.Drawing.Size(517, 77)
        Me.SeriesSelection2.Sqlstr = Nothing
        Me.SeriesSelection2.TabIndex = 104
        '
        'SeriesSelection1
        '
        Me.SeriesSelection1.DataSet = Nothing
        Me.SeriesSelection1.DataSet1 = Nothing
        Me.SeriesSelection1.DataSet2 = Nothing
        Me.SeriesSelection1.Location = New System.Drawing.Point(3, 56)
        Me.SeriesSelection1.Name = "SeriesSelection1"
        Me.SeriesSelection1.SeriesLegend = Nothing
        Me.SeriesSelection1.Size = New System.Drawing.Size(517, 77)
        Me.SeriesSelection1.Sqlstr = Nothing
        Me.SeriesSelection1.TabIndex = 103
        '
        'RMCrcy
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.Chart1)
        Me.Controls.Add(Me.SeriesSelection6)
        Me.Controls.Add(Me.SeriesSelection5)
        Me.Controls.Add(Me.SeriesSelection4)
        Me.Controls.Add(Me.SeriesSelection3)
        Me.Controls.Add(Me.SeriesSelection2)
        Me.Controls.Add(Me.SeriesSelection1)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "RMCrcy"
        Me.Size = New System.Drawing.Size(1041, 682)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.Chart1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Protected Friend WithEvents SeriesSelection1 As RawMaterial.seriesSelection
    Protected Friend WithEvents SeriesSelection2 As RawMaterial.seriesSelection
    Protected Friend WithEvents SeriesSelection3 As RawMaterial.seriesSelection
    Protected Friend WithEvents SeriesSelection4 As RawMaterial.seriesSelection
    Protected Friend WithEvents SeriesSelection5 As RawMaterial.seriesSelection
    Protected Friend WithEvents SeriesSelection6 As RawMaterial.seriesSelection
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Protected Friend WithEvents DateTimePicker1 As System.Windows.Forms.DateTimePicker
    Protected Friend WithEvents DateTimePicker2 As System.Windows.Forms.DateTimePicker
    Protected Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Protected Friend WithEvents Chart1 As System.Windows.Forms.DataVisualization.Charting.Chart
    Friend WithEvents Button2 As System.Windows.Forms.Button

End Class
