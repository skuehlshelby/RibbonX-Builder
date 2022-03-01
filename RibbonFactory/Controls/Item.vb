Imports RibbonFactory.ControlInterfaces
Imports stdole

Namespace Controls

	''' <summary>
	''' Represents an item in a combobox, dropdown, or gallery.
	''' </summary>
	Public NotInheritable Class Item
		Inherits RibbonElement
		Implements ILabel
		Implements IScreenTip
		Implements ISuperTip
		Implements IImage

		Private _label As String
		Private _screenTip As String
		Private _superTip As String
		Private _image As IPictureDisp
		Private ReadOnly _synchronizeLabelAndScreenTip As Boolean

		Friend Sub New(id As String, label As String, screenTip As String, superTip As String, image As IPictureDisp, Optional synchronizeLabelAndScreenTip As Boolean = True, Optional tag As Object = Nothing)
			MyBase.New(tag)
			Me.ID = id
			_label = label
			_screenTip = screenTip
			_superTip = superTip
			_image = image
			_synchronizeLabelAndScreenTip = synchronizeLabelAndScreenTip
		End Sub

		Public Overrides ReadOnly Property ID As String

		Public Overrides ReadOnly Property XML As String
			Get
				Return String.Empty
			End Get
		End Property

		Public Property Label As String Implements ILabel.Label
			Get
				Return _label
			End Get
			Set
				If Not _label.Equals(Value, StringComparison.OrdinalIgnoreCase) Then
					_label = Value

					If _synchronizeLabelAndScreenTip Then
						_screenTip = Value
					End If

					RefreshNeeded()
				End If
			End Set
		End Property

		Public Property ScreenTip As String Implements IScreenTip.ScreenTip
			Get
				Return _screenTip
			End Get
			Set
				If Not _screenTip.Equals(Value, StringComparison.OrdinalIgnoreCase) Then
					_screenTip = Value
					RefreshNeeded()
				End If
			End Set
		End Property

		Public Property SuperTip As String Implements ISuperTip.SuperTip
			Get
				Return _superTip
			End Get
			Set
				If Not _superTip.Equals(Value, StringComparison.OrdinalIgnoreCase) Then
					_superTip = Value
					RefreshNeeded()
				End If
			End Set
		End Property

		Public Property Image As IPictureDisp Implements IImage.Image
			Get
				Return _image
			End Get
			Set
				If Not ReferenceEquals(_image, Value) Then
					_image = Value
					RefreshNeeded()
				End If
			End Set
		End Property

		Public Shared Function Blank() As Item
			Return New Item(String.Empty, String.Empty, String.Empty, String.Empty, Nothing, False, Nothing)
		End Function

	End Class

End Namespace