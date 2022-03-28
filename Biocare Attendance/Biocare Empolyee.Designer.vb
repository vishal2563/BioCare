<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Biocare_Empolyee
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
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.TextBox4 = New System.Windows.Forms.TextBox()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.DataGridView2 = New System.Windows.Forms.DataGridView()
        Me.SrNo1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dated1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Empcode1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CardNo1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.EmpName1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ShiftStart1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ShiftEnd1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.InTime1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.OutTime1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ShifHrs1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.WorkHrs1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.OtHhrs1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Earlyarr1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.LatArr1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.LateDeep1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.EarlyDeep1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Status1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Tdate1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.TextBox5 = New System.Windows.Forms.TextBox()
        Me.TextBox6 = New System.Windows.Forms.TextBox()
        Me.Panel1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.FromArgb(CType(CType(188, Byte), Integer), CType(CType(124, Byte), Integer), CType(CType(54, Byte), Integer))
        Me.Panel1.Controls.Add(Me.Button5)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Location = New System.Drawing.Point(2, 1)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1603, 36)
        Me.Panel1.TabIndex = 0
        '
        'Button5
        '
        Me.Button5.FlatAppearance.BorderColor = System.Drawing.Color.White
        Me.Button5.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button5.ForeColor = System.Drawing.Color.White
        Me.Button5.Location = New System.Drawing.Point(1507, 8)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(37, 23)
        Me.Button5.TabIndex = 26
        Me.Button5.Text = "X"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(532, 2)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(247, 28)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "BioCare Attendance "
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.ComboBox1)
        Me.GroupBox1.Controls.Add(Me.TextBox6)
        Me.GroupBox1.Controls.Add(Me.DataGridView1)
        Me.GroupBox1.Controls.Add(Me.TextBox4)
        Me.GroupBox1.Controls.Add(Me.TextBox5)
        Me.GroupBox1.Controls.Add(Me.Button4)
        Me.GroupBox1.Controls.Add(Me.TextBox3)
        Me.GroupBox1.Controls.Add(Me.TextBox2)
        Me.GroupBox1.Controls.Add(Me.DataGridView2)
        Me.GroupBox1.Controls.Add(Me.TextBox1)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Button3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Button2)
        Me.GroupBox1.Controls.Add(Me.Button1)
        Me.GroupBox1.Location = New System.Drawing.Point(2, 46)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(1356, 733)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Attendance"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(536, 21)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(106, 16)
        Me.Label4.TabIndex = 16
        Me.Label4.Text = "Compnay Name"
        '
        'ComboBox1
        '
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Items.AddRange(New Object() {"BIOCARE MEDITECH PVT LTD", "WWES IT SOLUTIONS PVT LTD"})
        Me.ComboBox1.Location = New System.Drawing.Point(648, 18)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(121, 21)
        Me.ComboBox1.TabIndex = 15
        Me.ComboBox1.Text = "Select"
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(2, 66)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(767, 419)
        Me.DataGridView1.TabIndex = 0
        '
        'TextBox4
        '
        Me.TextBox4.Location = New System.Drawing.Point(938, 10)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(72, 20)
        Me.TextBox4.TabIndex = 13
        '
        'Button4
        '
        Me.Button4.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button4.Location = New System.Drawing.Point(436, 18)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(75, 27)
        Me.Button4.TabIndex = 10
        Me.Button4.Text = "Exit"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(1283, 10)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(73, 20)
        Me.TextBox2.TabIndex = 8
        '
        'DataGridView2
        '
        Me.DataGridView2.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        Me.DataGridView2.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.DataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView2.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.SrNo1, Me.dated1, Me.Empcode1, Me.CardNo1, Me.EmpName1, Me.ShiftStart1, Me.ShiftEnd1, Me.InTime1, Me.OutTime1, Me.ShifHrs1, Me.WorkHrs1, Me.OtHhrs1, Me.Earlyarr1, Me.LatArr1, Me.LateDeep1, Me.EarlyDeep1, Me.Status1, Me.Tdate1})
        Me.DataGridView2.GridColor = System.Drawing.SystemColors.ActiveCaption
        Me.DataGridView2.Location = New System.Drawing.Point(786, 66)
        Me.DataGridView2.Name = "DataGridView2"
        Me.DataGridView2.Size = New System.Drawing.Size(570, 419)
        Me.DataGridView2.TabIndex = 7
        '
        'SrNo1
        '
        Me.SrNo1.HeaderText = "SrNo"
        Me.SrNo1.Name = "SrNo1"
        '
        'dated1
        '
        Me.dated1.HeaderText = "dated"
        Me.dated1.Name = "dated1"
        '
        'Empcode1
        '
        Me.Empcode1.HeaderText = "Emp code"
        Me.Empcode1.Name = "Empcode1"
        '
        'CardNo1
        '
        Me.CardNo1.HeaderText = "Card No"
        Me.CardNo1.Name = "CardNo1"
        '
        'EmpName1
        '
        Me.EmpName1.HeaderText = "Emp Name"
        Me.EmpName1.Name = "EmpName1"
        '
        'ShiftStart1
        '
        Me.ShiftStart1.HeaderText = "Shift Start"
        Me.ShiftStart1.Name = "ShiftStart1"
        '
        'ShiftEnd1
        '
        Me.ShiftEnd1.HeaderText = "Shift End"
        Me.ShiftEnd1.Name = "ShiftEnd1"
        '
        'InTime1
        '
        Me.InTime1.HeaderText = "In Time"
        Me.InTime1.Name = "InTime1"
        '
        'OutTime1
        '
        Me.OutTime1.HeaderText = "Out Time"
        Me.OutTime1.Name = "OutTime1"
        '
        'ShifHrs1
        '
        Me.ShifHrs1.HeaderText = "Shif Hrs"
        Me.ShifHrs1.Name = "ShifHrs1"
        '
        'WorkHrs1
        '
        Me.WorkHrs1.HeaderText = "Work Hrs"
        Me.WorkHrs1.Name = "WorkHrs1"
        '
        'OtHhrs1
        '
        Me.OtHhrs1.HeaderText = "Ot Hhrs"
        Me.OtHhrs1.Name = "OtHhrs1"
        '
        'Earlyarr1
        '
        Me.Earlyarr1.HeaderText = "Earl yarr"
        Me.Earlyarr1.Name = "Earlyarr1"
        '
        'LatArr1
        '
        Me.LatArr1.HeaderText = "Lat Arr"
        Me.LatArr1.Name = "LatArr1"
        '
        'LateDeep1
        '
        Me.LateDeep1.HeaderText = "Late Deep"
        Me.LateDeep1.Name = "LateDeep1"
        '
        'EarlyDeep1
        '
        Me.EarlyDeep1.HeaderText = "Early Deep"
        Me.EarlyDeep1.Name = "EarlyDeep1"
        '
        'Status1
        '
        Me.Status1.HeaderText = "Status"
        Me.Status1.Name = "Status1"
        '
        'Tdate1
        '
        Me.Tdate1.HeaderText = "Tdate"
        Me.Tdate1.Name = "Tdate1"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(942, 40)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(39, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Label3"
        '
        'Button3
        '
        Me.Button3.Enabled = False
        Me.Button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button3.Location = New System.Drawing.Point(355, 18)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(75, 27)
        Me.Button3.TabIndex = 4
        Me.Button3.Text = "Read File"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(795, 23)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(39, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Label2"
        '
        'Button2
        '
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.Location = New System.Drawing.Point(274, 18)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 28)
        Me.Button2.TabIndex = 2
        Me.Button2.Text = "Submit"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(71, 19)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(197, 27)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Import Into Database"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(1016, 10)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(53, 20)
        Me.TextBox1.TabIndex = 6
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(1207, 10)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(70, 20)
        Me.TextBox3.TabIndex = 9
        '
        'TextBox5
        '
        Me.TextBox5.Location = New System.Drawing.Point(1075, 10)
        Me.TextBox5.Name = "TextBox5"
        Me.TextBox5.Size = New System.Drawing.Size(54, 20)
        Me.TextBox5.TabIndex = 12
        '
        'TextBox6
        '
        Me.TextBox6.Location = New System.Drawing.Point(1135, 10)
        Me.TextBox6.Name = "TextBox6"
        Me.TextBox6.Size = New System.Drawing.Size(66, 20)
        Me.TextBox6.TabIndex = 14
        '
        'Biocare_Empolyee
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightYellow
        Me.ClientSize = New System.Drawing.Size(1370, 749)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "Biocare_Empolyee"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Biocare_Empolyee"
        Me.Panel1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents DataGridView2 As System.Windows.Forms.DataGridView
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents SrNo1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dated1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Empcode1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CardNo1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents EmpName1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ShiftStart1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ShiftEnd1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents InTime1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents OutTime1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ShifHrs1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents WorkHrs1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents OtHhrs1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Earlyarr1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LatArr1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LateDeep1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents EarlyDeep1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Status1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Tdate1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TextBox6 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox5 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
End Class
