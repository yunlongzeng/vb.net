<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form 重写 Dispose，以清理组件列表。
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows 窗体设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改它。  
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Button5 = New System.Windows.Forms.Button()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(55, 90)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(113, 37)
        Me.Button1.TabIndex = 7
        Me.Button1.Text = "Search"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(213, 39)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(100, 22)
        Me.TextBox2.TabIndex = 8
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(240, 19)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(54, 17)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "IDCode"
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(349, 39)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(183, 22)
        Me.TextBox3.TabIndex = 8
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(406, 19)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(73, 17)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "TypeCode"
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(287, 90)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(192, 37)
        Me.Button3.TabIndex = 7
        Me.Button3.Text = "Insert or replace"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'ImageList1
        '
        Me.ImageList1.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit
        Me.ImageList1.ImageSize = New System.Drawing.Size(16, 20)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        '
        'DataGridView1
        '
        Me.DataGridView1.BackgroundColor = System.Drawing.SystemColors.Info
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(22, 149)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowTemplate.Height = 24
        Me.DataGridView1.Size = New System.Drawing.Size(669, 150)
        Me.DataGridView1.TabIndex = 12
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(578, 90)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(113, 37)
        Me.Button4.TabIndex = 7
        Me.Button4.Text = "Delete"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.TextBox1.Location = New System.Drawing.Point(87, 39)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(100, 22)
        Me.TextBox1.TabIndex = 8
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(34, 19)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(176, 17)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "SmartCode(8,14,18,19,23)"
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(578, 32)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(113, 37)
        Me.Button5.TabIndex = 7
        Me.Button5.Text = "Refresh"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AccessibleName = ""
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(714, 333)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TextBox3)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button1)
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Form1"
        Me.TopMost = True
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button1 As Windows.Forms.Button
    Friend WithEvents TextBox2 As Windows.Forms.TextBox
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents TextBox3 As Windows.Forms.TextBox
    Friend WithEvents Label3 As Windows.Forms.Label
    Friend WithEvents Button3 As Windows.Forms.Button
    Friend WithEvents ImageList1 As Windows.Forms.ImageList
    Friend WithEvents DataGridView1 As Windows.Forms.DataGridView
    Friend WithEvents Button4 As Windows.Forms.Button
    Friend WithEvents TextBox1 As Windows.Forms.TextBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents Button5 As Windows.Forms.Button
End Class
