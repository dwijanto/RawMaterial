<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Chart1
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Chart1))
        Me.Button1 = New System.Windows.Forms.Button()
        Me.MyChart6 = New RawMaterial.MyChart()
        Me.MyChart5 = New RawMaterial.MyChart()
        Me.MyChart4 = New RawMaterial.MyChart()
        Me.MyChart3 = New RawMaterial.MyChart()
        Me.MyChart2 = New RawMaterial.MyChart()
        Me.MyChart1 = New RawMaterial.MyChart()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(3, 12)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(97, 24)
        Me.Button1.TabIndex = 6
        Me.Button1.Text = "Export"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'MyChart6
        '
        Me.MyChart6.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.MyChart6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.MyChart6.DataSet1 = Nothing
        Me.MyChart6.Location = New System.Drawing.Point(595, 661)
        Me.MyChart6.Name = "MyChart6"
        Me.MyChart6.Size = New System.Drawing.Size(586, 303)
        Me.MyChart6.TabIndex = 5
        '
        'MyChart5
        '
        Me.MyChart5.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.MyChart5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.MyChart5.DataSet1 = Nothing
        Me.MyChart5.Location = New System.Drawing.Point(3, 661)
        Me.MyChart5.Name = "MyChart5"
        Me.MyChart5.Size = New System.Drawing.Size(586, 303)
        Me.MyChart5.TabIndex = 4
        '
        'MyChart4
        '
        Me.MyChart4.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.MyChart4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.MyChart4.DataSet1 = Nothing
        Me.MyChart4.Location = New System.Drawing.Point(595, 352)
        Me.MyChart4.Name = "MyChart4"
        Me.MyChart4.Size = New System.Drawing.Size(586, 303)
        Me.MyChart4.TabIndex = 3
        '
        'MyChart3
        '
        Me.MyChart3.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.MyChart3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.MyChart3.DataSet1 = Nothing
        Me.MyChart3.Location = New System.Drawing.Point(3, 352)
        Me.MyChart3.Name = "MyChart3"
        Me.MyChart3.Size = New System.Drawing.Size(586, 303)
        Me.MyChart3.TabIndex = 2
        '
        'MyChart2
        '
        Me.MyChart2.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.MyChart2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.MyChart2.DataSet1 = Nothing
        Me.MyChart2.Location = New System.Drawing.Point(595, 43)
        Me.MyChart2.Name = "MyChart2"
        Me.MyChart2.Size = New System.Drawing.Size(586, 303)
        Me.MyChart2.TabIndex = 1
        '
        'MyChart1
        '
        Me.MyChart1.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.MyChart1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.MyChart1.DataSet1 = Nothing
        Me.MyChart1.Location = New System.Drawing.Point(3, 43)
        Me.MyChart1.Name = "MyChart1"
        Me.MyChart1.Size = New System.Drawing.Size(586, 303)
        Me.MyChart1.TabIndex = 0
        '
        'Chart1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(1084, 802)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.MyChart6)
        Me.Controls.Add(Me.MyChart5)
        Me.Controls.Add(Me.MyChart4)
        Me.Controls.Add(Me.MyChart3)
        Me.Controls.Add(Me.MyChart2)
        Me.Controls.Add(Me.MyChart1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Chart1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Raw Material Charts"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents MyChart1 As RawMaterial.MyChart
    Friend WithEvents MyChart2 As RawMaterial.MyChart
    Friend WithEvents MyChart3 As RawMaterial.MyChart
    Friend WithEvents MyChart4 As RawMaterial.MyChart
    Friend WithEvents MyChart5 As RawMaterial.MyChart
    Friend WithEvents MyChart6 As RawMaterial.MyChart
    Friend WithEvents Button1 As System.Windows.Forms.Button
End Class
