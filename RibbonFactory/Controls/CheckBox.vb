Imports RibbonFactory.ControlInterfaces
Imports RibbonFactory.RibbonAttributes

Namespace Controls

	Public Class CheckBox
		Inherits RibbonElement
		Implements IEnabled
		Implements IVisible
		Implements ILabel
		Implements IScreenTip
		Implements ISuperTip
		Implements IDescription
		Implements IKeyTip
		Implements IPressed
		Implements IDefaultProvider

		Private ReadOnly _attributes As AttributeSet

		Friend Sub New(attributes As AttributeSet, Optional tag As Object = Nothing)
			MyBase.New(tag)
			_attributes = attributes
			AddHandler _attributes.AttributeChanged, AddressOf RefreshNeeded
		End Sub

#Region "Events"

		Public Event BeforeStateChange As EventHandler(Of BeforeStateChangeEventArgs)

		Public Event StateChanged As EventHandler(Of StateChangedEventArgs)

#End Region

#Region "Base Class Overrides"

		Public Overrides ReadOnly Property ID As String
			Get
				Return _attributes.ReadOnlyLookup(Of String)(AttributeName.Id).GetValue()
			End Get
		End Property

		Public Overrides ReadOnly Property XML As String
			Get
				Return $"<checkBox { _attributes }/>"
			End Get
		End Property

#End Region

#Region "Display Text"

		Public Property Description As String Implements IDescription.Description
			Get
				Return _attributes.ReadOnlyLookup(Of String)(AttributeName.Description).GetValue()
			End Get
			Set
				_attributes.ReadWriteLookup(Of String)(AttributeName.GetDescription).SetValue(Value)
			End Set
		End Property

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

#Region "Action"

		Public Property Pressed As Boolean Implements IPressed.Pressed
			Get
				Return _attributes.ReadOnlyLookup(Of Boolean)(AttributeName.GetPressed).GetValue()
			End Get
			Set
				Try
					SuspendAutomaticRefreshing()

					Dim e As BeforeStateChangeEventArgs = New BeforeStateChangeEventArgs(Pressed)

					RaiseEvent BeforeStateChange(Me, e)

					If Not e.Cancel Then
						_attributes.ReadWriteLookup(Of Boolean)(AttributeCategory.Pressed).SetValue(Value)

						RaiseEvent StateChanged(Me, New StateChangedEventArgs(Value))
					End If
				Finally
					ResumeAutomaticRefreshing()
				End Try
			End Set
		End Property

#End Region

#Region "Control Cloning Helpers"

		Friend Function GetBeforeStateChangeInvocationList() As EventHandler(Of BeforeStateChangeEventArgs)()
			Dim e As EventHandler(Of BeforeStateChangeEventArgs) = BeforeStateChangeEvent

			If e Is Nothing Then
				Return New EventHandler(Of BeforeStateChangeEventArgs)() {}
			Else
				Return e.
					GetInvocationList().
					OfType(Of EventHandler(Of BeforeStateChangeEventArgs)).
					ToArray()
			End If
		End Function

		Friend Function GetStateChangedInvocationList() As EventHandler(Of StateChangedEventArgs)()
			Dim e As EventHandler(Of StateChangedEventArgs) = StateChangedEvent

			If e Is Nothing Then
				Return New EventHandler(Of StateChangedEventArgs)() {}
			Else
				Return e.
					GetInvocationList().
					OfType(Of EventHandler(Of StateChangedEventArgs)).
					ToArray()
			End If
		End Function

		Private Function GetDefaults() As AttributeSet Implements IDefaultProvider.GetDefaults
			Return _attributes
		End Function

#End Region

		Public NotInheritable Class BeforeStateChangeEventArgs
			Inherits EventArgs

			Public Sub New(isCurrentlyChecked As Boolean)
				Me.IsCurrentlyChecked = isCurrentlyChecked
			End Sub

			Public Property Cancel As Boolean

			Public ReadOnly Property IsCurrentlyChecked As Boolean

		End Class

		Public NotInheritable Class StateChangedEventArgs
			Inherits EventArgs

			Public Sub New(isCurrentlyChecked As Boolean)
				Me.IsCurrentlyChecked = isCurrentlyChecked
			End Sub

			Public ReadOnly Property IsCurrentlyChecked As Boolean

		End Class

	End Class

End Namespace