<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class distance
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
        Me.txtLocLAT1 = New System.Windows.Forms.TextBox()
        Me.txtLocLON1 = New System.Windows.Forms.TextBox()
        Me.txtLocLAT2 = New System.Windows.Forms.TextBox()
        Me.txtLocLON2 = New System.Windows.Forms.TextBox()
        Me.txtDistance = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(532, 38)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'txtLocLAT1
        '
        Me.txtLocLAT1.Location = New System.Drawing.Point(310, 38)
        Me.txtLocLAT1.Name = "txtLocLAT1"
        Me.txtLocLAT1.Size = New System.Drawing.Size(100, 20)
        Me.txtLocLAT1.TabIndex = 1
        '
        'txtLocLON1
        '
        Me.txtLocLON1.Location = New System.Drawing.Point(310, 64)
        Me.txtLocLON1.Name = "txtLocLON1"
        Me.txtLocLON1.Size = New System.Drawing.Size(100, 20)
        Me.txtLocLON1.TabIndex = 2
        '
        'txtLocLAT2
        '
        Me.txtLocLAT2.Location = New System.Drawing.Point(310, 104)
        Me.txtLocLAT2.Name = "txtLocLAT2"
        Me.txtLocLAT2.Size = New System.Drawing.Size(100, 20)
        Me.txtLocLAT2.TabIndex = 3
        '
        'txtLocLON2
        '
        Me.txtLocLON2.Location = New System.Drawing.Point(310, 141)
        Me.txtLocLON2.Name = "txtLocLON2"
        Me.txtLocLON2.Size = New System.Drawing.Size(100, 20)
        Me.txtLocLON2.TabIndex = 4
        '
        'txtDistance
        '
        Me.txtDistance.Location = New System.Drawing.Point(310, 181)
        Me.txtDistance.Name = "txtDistance"
        Me.txtDistance.Size = New System.Drawing.Size(100, 20)
        Me.txtDistance.TabIndex = 5
        '
        'distance
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(688, 261)
        Me.Controls.Add(Me.txtDistance)
        Me.Controls.Add(Me.txtLocLON2)
        Me.Controls.Add(Me.txtLocLAT2)
        Me.Controls.Add(Me.txtLocLON1)
        Me.Controls.Add(Me.txtLocLAT1)
        Me.Controls.Add(Me.Button1)
        Me.Name = "distance"
        Me.Text = "distance"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents txtLocLAT1 As System.Windows.Forms.TextBox
    Friend WithEvents txtLocLON1 As System.Windows.Forms.TextBox
    Friend WithEvents txtLocLAT2 As System.Windows.Forms.TextBox
    Friend WithEvents txtLocLON2 As System.Windows.Forms.TextBox
    Friend WithEvents txtDistance As System.Windows.Forms.TextBox
End Class
