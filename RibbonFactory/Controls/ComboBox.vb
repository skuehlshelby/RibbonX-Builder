Imports RibbonFactory.ControlInterfaces
Imports RibbonFactory.RibbonAttributes
Imports stdole

Namespace Controls

	Public NotInheritable Class ComboBox
		Inherits Container(Of Item)
		Implements IEnabled
		Implements IVisible
		Implements ILabel
		Implements IScreenTip
		Implements ISuperTip
		Implements IShowLabel
		Implements IKeyTip
		Implements IImage
		Implements IShowImage
		Implements IMaxLength
		Implements IText
		Implements IDefaultProvider

		Private ReadOnly _attributes As AttributeSet

		Friend Sub New(attributes As AttributeSet, Optional tag As Object = Nothing)
			MyBase.New(New List(Of Item), tag)
			_attributes = attributes
			AddHandler _attributes.AttributeChanged, AddressOf RefreshNeeded
		End Sub

#Region "Events"

		Public Event BeforeTextChange As EventHandler(Of BeforeTextChangeEventArgs)

		Public Event TextChanged As EventHandler(Of TextChangedEventArgs)

#End Region

#Region "RibbonElement Overrides"

		Public Overrides ReadOnly Property ID As String
			Get
				Return _attributes.ReadOnlyLookup(Of String)(AttributeName.Id).GetValue()
			End Get
		End Property

		Public Overrides ReadOnly Property XML As String
			Get
				Return $"<comboBox { _attributes }/>"
			End Get
		End Property

#End Region

#Region "Container Overrides"

		Public Overrides Sub Add(item As Item)
			AddHandler item.ValueChanged, AddressOf OnChildItemChange

			MyBase.Add(item)
		End Sub

		Public Overrides Sub Clear()
			If Items.Any() Then

				For Each item As Item In Me
					RemoveHandler item.ValueChanged, AddressOf OnChildItemChange
				Next

				Items.Clear()
				RefreshNeeded()
			End If
		End Sub

		Public Overrides Function Remove(item As Item) As Boolean
			If Items.Remove(item) Then
				RemoveHandler item.ValueChanged, AddressOf OnChildItemChange

				RefreshNeeded()

				Return True
			Else
				Return False
			End If
		End Function

		Friend Overrides Sub Flatten(results As ICollection(Of RibbonElement))
			results.Add(Me)
		End Sub

#End Region

#Region "Container Helpers"

		Private Sub OnChildItemChange(sender As Object, e As ValueChangedEventArgs)
			RefreshNeeded()
		End Sub

#End Region

#Region "Enabled and Visible"

		Public Property Enabled As Boolean Implements IEnabled.Enabled
			Get
				Return _attributes.ReadOnlyLookup(Of Boolean)(AttributeName.Enabled).GetValue()
			End Get
			Set
				_attributes.ReadWriteLookup(Of Boolean)(AttributeName.GetEnabled).SetValue(Value)
			End Set
		End Property

		Public Property Visible As Boolean Implements IVisible.Visible
			Get
				Return _attributes.ReadOnlyLookup(Of Boolean)(AttributeName.Visible).GetValue()
			End Get
			Set
				_attributes.ReadWriteLookup(Of Boolean)(AttributeName.GetVisible).SetValue(Value)
			End Set
		End Property

#End Region

#Region "DisplayText"

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
				_attributes.ReadWriteLookup(Of String)(AttributeName.GetLabel).SetValue(Value)
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

#Region "Image and ShowImage"

		Public Property Image As IPictureDisp Implements IImage.Image
			Get
				Return _attributes.ReadOnlyLookup(Of IPictureDisp)(AttributeName.GetImage).GetValue()
			End Get
			Set
				_attributes.ReadWriteLookup(Of IPictureDisp)(AttributeName.GetImage).SetValue(Value)
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

						Dim e As BeforeTextChangeEventArgs = New BeforeTextChangeEventArgs(Value, Items.ToArray())

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

			Public Sub New(newText As String, items As ICollection(Of Item))
				Me.NewText = newText
				Me.Items = items
				Cancel = False
			End Sub

			Public Property Cancel As Boolean

			Public ReadOnly Property Items As ICollection(Of Item)

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