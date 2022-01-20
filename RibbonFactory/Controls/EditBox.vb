Imports RibbonFactory.ControlInterfaces
Imports RibbonFactory.RibbonAttributes
Imports stdole

Namespace Controls

	Public Class EditBox
		Inherits RibbonElement
		Implements IEnabled
		Implements IVisible
		Implements ILabel
		Implements IScreenTip
		Implements ISuperTip
		Implements IShowLabel
		Implements IKeyTip
		Implements IImage
		Implements IShowImage
		Implements IText
		Implements IMaxLength
		Implements IDefaultProvider

		Private ReadOnly _attributes As AttributeSet

		Friend Sub New(attributes As AttributeSet, Optional tag As Object = Nothing)
			MyBase.New(tag)
			_attributes = attributes
			AddHandler _attributes.AttributeChanged, AddressOf RefreshNeeded
		End Sub

#Region "Events"

		Public Event BeforeTextChange As EventHandler(Of BeforeTextChangeEventArgs)

		Public Event TextChanged As EventHandler(Of TextChangedEventArgs)

#End Region

#Region "Base Class Overrides"

		Public Overrides ReadOnly Property ID As String
			Get
				Return _attributes.ReadOnlyLookup(Of String)(AttributeName.Id).GetValue()
			End Get
		End Property

		Public Overrides ReadOnly Property XML As String
			Get
				Return $"<editBox { _attributes }/>"
			End Get
		End Property

#End Region

#Region "Display Text"

		Public Property KeyTip As KeyTip Implements IKeyTip.KeyTip
			Get
				Return _attributes.ReadOnlyLookup(Of KeyTip)(AttributeName.Keytip).GetValue()
			End Get
			Set
				_attributes.ReadWriteLookup(Of KeyTip)(AttributeName.GetKeytip).SetValue(Value)
			End Set
		End Property

		Public Property Label As String Implements ILabel.Label
			Get
				Return _attributes.ReadOnlyLookup(Of String)(AttributeName.Label).GetValue()
			End Get
			Set
				_attributes.ReadWriteLookup(Of String)(AttributeName.Label).SetValue(Value)
			End Set
		End Property

		Public Property ShowLabel As Boolean Implements IShowLabel.ShowLabel
			Get
				Return _attributes.ReadOnlyLookup(Of Boolean)(AttributeName.ShowLabel).GetValue()
			End Get
			Set
				_attributes.ReadWriteLookup(Of Boolean)(AttributeName.GetShowLabel).SetValue(Value)
			End Set
		End Property

		Public Property ScreenTip As String Implements IScreenTip.ScreenTip
			Get
				Return _attributes.ReadOnlyLookup(Of String)(AttributeName.Screentip).GetValue()
			End Get
			Set
				_attributes.ReadWriteLookup(Of String)(AttributeName.GetScreentip).SetValue(Value)
			End Set
		End Property

		Public Property SuperTip As String Implements ISuperTip.SuperTip
			Get
				Return _attributes.ReadOnlyLookup(Of String)(AttributeName.Supertip).GetValue()
			End Get
			Set
				_attributes.ReadWriteLookup(Of String)(AttributeName.GetSupertip).SetValue(Value)
			End Set
		End Property

#End Region

#Region "Enabled and Visible"

		Public Property Enabled As Boolean Implements IEnabled.Enabled
			Get
				Return _attributes.ReadOnlyLookup(Of Boolean)(AttributeName.Enabled).GetValue()
			End Get
			Set
				_attributes.ReadWriteLookup(Of Boolean)(AttributeName.Enabled).SetValue(Value)
			End Set
		End Property

		Public Property Visible As Boolean Implements IVisible.Visible
			Get
				Return _attributes.ReadOnlyLookup(Of Boolean)(AttributeName.Visible).GetValue()
			End Get
			Set
				_attributes.ReadWriteLookup(Of Boolean)(AttributeName.Visible).SetValue(Value)
			End Set
		End Property

#End Region

#Region "Image"

		Public Property Image As IPictureDisp Implements IImage.Image
			Get
				Return _attributes.ReadOnlyLookup(Of IPictureDisp)(AttributeCategory.Image).GetValue()
			End Get
			Set
				_attributes.ReadWriteLookup(Of IPictureDisp)(AttributeCategory.Image).SetValue(Value)
			End Set
		End Property

		Public Property ShowImage As Boolean Implements IShowImage.ShowImage
			Get
				Return _attributes.ReadOnlyLookup(Of Boolean)(AttributeName.ShowImage).GetValue()
			End Get
			Set
				_attributes.ReadWriteLookup(Of Boolean)(AttributeName.GetShowImage).SetValue(Value)
			End Set
		End Property

#End Region

#Region "MaxLength"

		Public ReadOnly Property MaxLength As Integer Implements IMaxLength.MaxLength
			Get
				Return _attributes.ReadOnlyLookup(Of Integer)(AttributeName.MaxLength).GetValue()
			End Get
		End Property

#End Region

#Region "Action"

		Public Property Text As String Implements IText.Text
			Get
				Return _attributes.ReadOnlyLookup(Of String)(AttributeName.GetText).GetValue()
			End Get
			Set
				Dim initialValue As String = Text

				If Not initialValue.Equals(Value, StringComparison.OrdinalIgnoreCase) Then
					Try
						SuspendAutomaticRefreshing()

						Dim e As BeforeTextChangeEventArgs = New BeforeTextChangeEventArgs(Value)

						RaiseEvent BeforeTextChange(Me, e)

						If Not e.Cancel Then
							Dim updatedValue As String = e.NewText

							_attributes.ReadWriteLookup(Of String)(AttributeCategory.Text).SetValue(updatedValue)

							RaiseEvent TextChanged(Me, New TextChangedEventArgs(updatedValue, initialValue))
						End If
					Finally
						ResumeAutomaticRefreshing()
					End Try
				End If
			End Set
		End Property

#End Region

#Region "Helpers For Control Cloning"

		Friend Function GetBeforeTextChangeInvocationList() As EventHandler(Of BeforeTextChangeEventArgs)()
			Dim e As EventHandler(Of BeforeTextChangeEventArgs) = BeforeTextChangeEvent

			If e Is Nothing Then
				Return New EventHandler(Of BeforeTextChangeEventArgs)() {}
			Else
				Return e.
					GetInvocationList().
					OfType(Of EventHandler(Of BeforeTextChangeEventArgs)).
					ToArray()
			End If
		End Function

		Friend Function GetTextChangedInvocationList() As EventHandler(Of TextChangedEventArgs)()
			Dim e As EventHandler(Of TextChangedEventArgs) = TextChangedEvent

			If e Is Nothing Then
				Return New EventHandler(Of TextChangedEventArgs)() {}
			Else
				Return e.
					GetInvocationList().
					OfType(Of EventHandler(Of TextChangedEventArgs)).
					ToArray()
			End If
		End Function

		Private Function GetDefaults() As AttributeSet Implements IDefaultProvider.GetDefaults
			Return _attributes
		End Function

#End Region

		Public NotInheritable Class BeforeTextChangeEventArgs
			Inherits EventArgs

			Public Sub New(newText As String)
				Me.NewText = newText
				Cancel = False
			End Sub

			Public Property Cancel As Boolean

			Public Property NewText As String

		End Class

		Public Class TextChangedEventArgs
			Inherits EventArgs

			Public Sub New(newText As String, oldText As String)
				Me.NewText = newText
				Me.OldText = oldText
			End Sub

			Public ReadOnly Property NewText As String

			Public ReadOnly Property OldText As String

		End Class

	End Class

End Namespace