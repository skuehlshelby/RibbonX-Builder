Imports System.Drawing
Imports System.Windows.Forms

Public NotInheritable Class MainForm
	Inherits Form
	Implements IDisposable


	Private ReadOnly addPanel As TableLayoutPanel
	Private ReadOnly removePanel As TableLayoutPanel
	Private ReadOnly topPanel As TableLayoutPanel
	Private ReadOnly bottomPanel As TableLayoutPanel
	Private ReadOnly buttons As ICollection(Of Button)
	Private ReadOnly labels As ICollection(Of Label)
	Private ReadOnly palette As ColorPalette

	Private Structure ColorPalette
		Public Sub New(background As Color, backgroundMuted As Color, foregroundHighContrast As Color, foreground As Color)
			Me.New()
			Me.Background = background
			Me.BackgroundMuted = backgroundMuted
			Me.ForegroundHighContrast = foregroundHighContrast
			Me.Foreground = foreground
		End Sub

		Public ReadOnly Property Background As Color
		Public ReadOnly Property BackgroundMuted As Color
		Public ReadOnly Property ForegroundHighContrast As Color
		Public ReadOnly Property Foreground As Color
	End Structure

	Public Sub New()
		palette = New ColorPalette(
			Color.FromArgb(23, 23, 23),
			Color.FromArgb(68, 68, 68),
			Color.FromArgb(218, 0, 55),
			Color.FromArgb(237, 237, 237))

		Font = New Font("Montserrat", 12.0!, FontStyle.Regular)

		buttons = New Button() {
			New Button() With {.Text = "OK", .Name = "OK"},
			New Button() With {.Text = "Cancel", .Name = "Cancel"}
		}

		labels = New Label() {
			New Label() With {.Text = "Register", .Name = "Register", .ForeColor = palette.Foreground},
			New Label() With {.Text = "Un-Register", .Name = "Un-Register", .ForeColor = palette.Foreground}
		}

		topPanel = New TableLayoutPanel()
		topPanel.SuspendLayout()

		With topPanel
			.RowCount = 3
			.ColumnCount = 5

			With .ColumnStyles
				.Add(New ColumnStyle(SizeType.Percent, 10.0!))
				.Add(New ColumnStyle(SizeType.Percent, 35.0!))
				.Add(New ColumnStyle(SizeType.Percent, 10.0!))
				.Add(New ColumnStyle(SizeType.Percent, 35.0!))
				.Add(New ColumnStyle(SizeType.Percent, 10.0!))
			End With

			With .RowStyles
				.Add(New RowStyle(SizeType.Percent, 33.3!))
				.Add(New RowStyle(SizeType.Percent, 33.3!))
				.Add(New RowStyle(SizeType.Percent, 33.3!))
			End With

			With .Controls
				.Add(labels.First(), 1, 1)
				.Add(labels.Last(), 3, 1)
			End With

			.Dock = DockStyle.Top
		End With

		bottomPanel = New TableLayoutPanel()
		bottomPanel.SuspendLayout()
		SuspendLayout()

		With bottomPanel
			.RowCount = 3
			.ColumnCount = 5

			With .ColumnStyles
				.Add(New ColumnStyle(SizeType.Percent, 10.0!))
				.Add(New ColumnStyle(SizeType.Percent, 35.0!))
				.Add(New ColumnStyle(SizeType.Percent, 10.0!))
				.Add(New ColumnStyle(SizeType.Percent, 35.0!))
				.Add(New ColumnStyle(SizeType.Percent, 10.0!))
			End With

			With .RowStyles
				.Add(New RowStyle(SizeType.Percent, 33.3!))
				.Add(New RowStyle(SizeType.Percent, 33.3!))
				.Add(New RowStyle(SizeType.Percent, 33.3!))
			End With

			With .Controls
				.Add(buttons.First(), 1, 1)
				.Add(buttons.Last(), 3, 1)
			End With

			.Dock = DockStyle.Bottom
			.Name = "OkayCancelPanel"
		End With

		For Each btn As Button In buttons
			With btn
				.Anchor = AnchorStyles.Left Or AnchorStyles.Right
				.FlatStyle = FlatStyle.Flat
				.FlatAppearance.BorderSize = 0
				.BackColor = palette.BackgroundMuted
				.ForeColor = palette.Foreground
				.UseVisualStyleBackColor = False
			End With
		Next

		Dim cancel As Button = buttons.Single(Function(b) b.Name.Equals("Cancel"))
		AddHandler cancel.Click, AddressOf OnCancel
		CancelButton = cancel

		AutoScaleDimensions = New SizeF(6.0!, 13.0!)
		AutoScaleMode = AutoScaleMode.Font
		BackColor = palette.Background
		ClientSize = New Size(800, 450)
		Controls.Add(bottomPanel)
		Controls.Add(topPanel)
		Name = "Main Form"
		Text = "COM Registration Utility"
		topPanel.ResumeLayout(False)
		bottomPanel.ResumeLayout(False)
		ResumeLayout(False)
	End Sub

	Private Sub OnCancel(sender As Object, e As EventArgs)
		Hide()
	End Sub

	Protected Overrides Sub Dispose(disposing As Boolean)
		Try
			If disposing AndAlso buttons IsNot Nothing Then
				For Each btn As Button In buttons
					btn.Dispose()
				Next

				For Each lbl As Label In labels
					lbl.Dispose()
				Next

				bottomPanel.Dispose()
			End If
		Finally
			MyBase.Dispose(disposing)
		End Try
	End Sub
End Class
