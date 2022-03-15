<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Example
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
		Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
		Me.Button1 = New System.Windows.Forms.Button()
		Me.Button2 = New System.Windows.Forms.Button()
		Me.TableLayoutPanel1.SuspendLayout
		Me.SuspendLayout
		'
		'TableLayoutPanel1
		'
		Me.TableLayoutPanel1.ColumnCount = 5
		Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10!))
		Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 35!))
		Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10!))
		Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 35!))
		Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10!))
		Me.TableLayoutPanel1.Controls.Add(Me.Button1, 1, 1)
		Me.TableLayoutPanel1.Controls.Add(Me.Button2, 3, 1)
		Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Bottom
		'Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 350)
		Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
		Me.TableLayoutPanel1.RowCount = 3
		Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33334!))
		Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33334!))
		Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33334!))
		'Me.TableLayoutPanel1.Size = New System.Drawing.Size(800, 100)
		'Me.TableLayoutPanel1.TabIndex = 0
		'
		'Button1
		'
		Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right),System.Windows.Forms.AnchorStyles)
		Me.Button1.BackColor = System.Drawing.SystemColors.ControlDark
		Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
		Me.Button1.ForeColor = System.Drawing.SystemColors.ButtonFace
		'Me.Button1.Location = New System.Drawing.Point(83, 38)
		Me.Button1.Name = "Button1"
		'Me.Button1.Size = New System.Drawing.Size(274, 23)
		'Me.Button1.TabIndex = 0
		Me.Button1.Text = "Button1"
		Me.Button1.UseVisualStyleBackColor = false
		'
		'Button2
		'
		Me.Button2.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right),System.Windows.Forms.AnchorStyles)
		Me.Button2.BackColor = System.Drawing.SystemColors.ControlDark
		Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
		Me.Button2.ForeColor = System.Drawing.SystemColors.ButtonFace
		'Me.Button2.Location = New System.Drawing.Point(443, 38)
		Me.Button2.Name = "Button2"
		'Me.Button2.Size = New System.Drawing.Size(274, 23)
		'Me.Button2.TabIndex = 1
		Me.Button2.Text = "Button2"
		Me.Button2.UseVisualStyleBackColor = false
		'
		'Example
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6!, 13!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(800, 450)
		Me.Controls.Add(Me.TableLayoutPanel1)
		Me.Name = "Example"
		Me.Text = "Example"
		Me.TableLayoutPanel1.ResumeLayout(false)
		Me.ResumeLayout(false)

End Sub

	Friend WithEvents TableLayoutPanel1 As Windows.Forms.TableLayoutPanel
	Friend WithEvents Button1 As Windows.Forms.Button
	Friend WithEvents Button2 As Windows.Forms.Button
End Class
