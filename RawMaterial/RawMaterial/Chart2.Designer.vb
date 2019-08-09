<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Chart2
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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
        Me.Button1 = New System.Windows.Forms.Button()
        Me.CrcyChart6 = New RawMaterial.crcyChart()
        Me.CrcyChart5 = New RawMaterial.crcyChart()
        Me.CrcyChart4 = New RawMaterial.crcyChart()
        Me.CrcyChart3 = New RawMaterial.crcyChart()
        Me.CrcyChart2 = New RawMaterial.crcyChart()
        Me.CrcyChart1 = New RawMaterial.crcyChart()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(14, 12)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(93, 28)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "Export"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'CrcyChart6
        '
        Me.CrcyChart6.BackColor = System.Drawing.SystemColors.Window
        Me.CrcyChart6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CrcyChart6.Dataset = Nothing
        Me.CrcyChart6.Location = New System.Drawing.Point(604, 660)
        Me.CrcyChart6.Name = "CrcyChart6"
        Me.CrcyChart6.Size = New System.Drawing.Size(584, 301)
        Me.CrcyChart6.TabIndex = 6
        '
        'CrcyChart5
        '
        Me.CrcyChart5.BackColor = System.Drawing.SystemColors.Window
        Me.CrcyChart5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CrcyChart5.Dataset = Nothing
        Me.CrcyChart5.Location = New System.Drawing.Point(14, 660)
        Me.CrcyChart5.Name = "CrcyChart5"
        Me.CrcyChart5.Size = New System.Drawing.Size(584, 301)
        Me.CrcyChart5.TabIndex = 5
        '
        'CrcyChart4
        '
        Me.CrcyChart4.BackColor = System.Drawing.SystemColors.Window
        Me.CrcyChart4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CrcyChart4.Dataset = Nothing
        Me.CrcyChart4.Location = New System.Drawing.Point(604, 353)
        Me.CrcyChart4.Name = "CrcyChart4"
        Me.CrcyChart4.Size = New System.Drawing.Size(584, 301)
        Me.CrcyChart4.TabIndex = 4
        '
        'CrcyChart3
        '
        Me.CrcyChart3.BackColor = System.Drawing.SystemColors.Window
        Me.CrcyChart3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CrcyChart3.Dataset = Nothing
        Me.CrcyChart3.Location = New System.Drawing.Point(12, 353)
        Me.CrcyChart3.Name = "CrcyChart3"
        Me.CrcyChart3.Size = New System.Drawing.Size(584, 301)
        Me.CrcyChart3.TabIndex = 3
        '
        'CrcyChart2
        '
        Me.CrcyChart2.BackColor = System.Drawing.SystemColors.Window
        Me.CrcyChart2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CrcyChart2.Dataset = Nothing
        Me.CrcyChart2.Location = New System.Drawing.Point(604, 46)
        Me.CrcyChart2.Name = "CrcyChart2"
        Me.CrcyChart2.Size = New System.Drawing.Size(584, 301)
        Me.CrcyChart2.TabIndex = 1
        '
        'CrcyChart1
        '
        Me.CrcyChart1.BackColor = System.Drawing.SystemColors.Window
        Me.CrcyChart1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CrcyChart1.Dataset = Nothing
        Me.CrcyChart1.Location = New System.Drawing.Point(12, 46)
        Me.CrcyChart1.Name = "CrcyChart1"
        Me.CrcyChart1.Size = New System.Drawing.Size(586, 301)
        Me.CrcyChart1.TabIndex = 0
        '
        'ChartCurrency
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(996, 947)
        Me.Controls.Add(Me.CrcyChart6)
        Me.Controls.Add(Me.CrcyChart5)
        Me.Controls.Add(Me.CrcyChart4)
        Me.Controls.Add(Me.CrcyChart3)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.CrcyChart2)
        Me.Controls.Add(Me.CrcyChart1)
        Me.Name = "ChartCurrency"
        Me.Text = "Currency Charts"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents CrcyChart1 As RawMaterial.crcyChart
    Friend WithEvents CrcyChart2 As RawMaterial.crcyChart
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents CrcyChart3 As RawMaterial.crcyChart
    Friend WithEvents CrcyChart4 As RawMaterial.crcyChart
    Friend WithEvents CrcyChart5 As RawMaterial.crcyChart
    Friend WithEvents CrcyChart6 As RawMaterial.crcyChart
End Class
