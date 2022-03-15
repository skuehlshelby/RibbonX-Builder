<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
		Me.Panel2 = New System.Windows.Forms.Panel()
		Me.TabControl1 = New System.Windows.Forms.TabControl()
		Me.TabPage1 = New System.Windows.Forms.TabPage()
		Me.TabPage2 = New System.Windows.Forms.TabPage()
		Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
		Me.Button1 = New System.Windows.Forms.Button()
		Me.Button2 = New System.Windows.Forms.Button()
		Me.Panel2.SuspendLayout
		Me.TabControl1.SuspendLayout
		Me.TableLayoutPanel1.SuspendLayout
		Me.SuspendLayout
		'
		'Panel2
		'
		Me.Panel2.Controls.Add(Me.TabControl1)
		Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
		Me.Panel2.Location = New System.Drawing.Point(0, 0)
		Me.Panel2.Name = "Panel2"
		Me.Panel2.Size = New System.Drawing.Size(800, 450)
		Me.Panel2.TabIndex = 1
		'
		'TabControl1
		'
		Me.TabControl1.Controls.Add(Me.TabPage1)
		Me.TabControl1.Controls.Add(Me.TabPage2)
		Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
		Me.TabControl1.Location = New System.Drawing.Point(0, 0)
		Me.TabControl1.Name = "TabControl1"
		Me.TabControl1.SelectedIndex = 0
		Me.TabControl1.Size = New System.Drawing.Size(800, 450)
		Me.TabControl1.TabIndex = 0
		'
		'TabPage1
		'
		Me.TabPage1.Location = New System.Drawing.Point(4, 22)
		Me.TabPage1.Name = "TabPage1"
		Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
		Me.TabPage1.Size = New System.Drawing.Size(792, 424)
		Me.TabPage1.TabIndex = 0
		Me.TabPage1.Text = "TabPage1"
		Me.TabPage1.UseVisualStyleBackColor = true
		'
		'TabPage2
		'
		Me.TabPage2.Location = New System.Drawing.Point(4, 22)
		Me.TabPage2.Name = "TabPage2"
		Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
		Me.TabPage2.Size = New System.Drawing.Size(792, 424)
		Me.TabPage2.TabIndex = 1
		Me.TabPage2.Text = "TabPage2"
		Me.TabPage2.UseVisualStyleBackColor = true
		'
		'TableLayoutPanel1
		'
		Me.TableLayoutPanel1.ColumnCount = 5
		Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 24!))
		Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 24!))
		Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 4!))
		Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 24!))
		Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 24!))
		Me.TableLayoutPanel1.Controls.Add(Me.Button1, 1, 1)
		Me.TableLayoutPanel1.Controls.Add(Me.Button2, 3, 1)
		Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Bottom
		Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 370)
		Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
		Me.TableLayoutPanel1.RowCount = 3
		Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333!))
		Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333!))
		Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333!))
		Me.TableLayoutPanel1.Size = New System.Drawing.Size(800, 80)
		Me.TableLayoutPanel1.TabIndex = 2
		'
		'Button1
		'
		Me.Button1.Dock = System.Windows.Forms.DockStyle.Fill
		Me.Button1.Location = New System.Drawing.Point(195, 29)
		Me.Button1.Name = "Button1"
		Me.Button1.Size = New System.Drawing.Size(186, 20)
		Me.Button1.TabIndex = 0
		Me.Button1.Text = "Button1"
		Me.Button1.UseVisualStyleBackColor = true
		'
		'Button2
		'
		Me.Button2.Dock = System.Windows.Forms.DockStyle.Fill
		Me.Button2.Location = New System.Drawing.Point(419, 29)
		Me.Button2.Name = "Button2"
		Me.Button2.Size = New System.Drawing.Size(186, 20)
		Me.Button2.TabIndex = 1
		Me.Button2.Text = "Button2"
		Me.Button2.UseVisualStyleBackColor = true
		'
		'Form1
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6!, 13!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(800, 450)
		Me.Controls.Add(Me.TableLayoutPanel1)
		Me.Controls.Add(Me.Panel2)
		Me.Name = "Form1"
		Me.Text = "Form1"
		Me.Panel2.ResumeLayout(false)
		Me.TabControl1.ResumeLayout(false)
		Me.TableLayoutPanel1.ResumeLayout(false)
		Me.ResumeLayout(false)

End Sub

	Friend WithEvents Panel2 As Panel
	Friend WithEvents TabControl1 As TabControl
	Friend WithEvents TabPage1 As TabPage
	Friend WithEvents TabPage2 As TabPage
	Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
	Friend WithEvents Button1 As Button
	Friend WithEvents Button2 As Button
End Class
