<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Chart3
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
        Me.RmCrcy1 = New RawMaterial.RMCrcy()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(12, 12)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(149, 32)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Export To Excel"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'RmCrcy1
        '
        Me.RmCrcy1.Location = New System.Drawing.Point(12, 50)
        Me.RmCrcy1.Name = "RmCrcy1"
        Me.RmCrcy1.Size = New System.Drawing.Size(1041, 718)
        Me.RmCrcy1.TabIndex = 2
        '
        'Chart3
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(1076, 770)
        Me.Controls.Add(Me.RmCrcy1)
        Me.Controls.Add(Me.Button1)
        Me.MinimumSize = New System.Drawing.Size(1084, 798)
        Me.Name = "Chart3"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Multi Currencies & Charts"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents RmCrcy1 As RawMaterial.RMCrcy
End Class
