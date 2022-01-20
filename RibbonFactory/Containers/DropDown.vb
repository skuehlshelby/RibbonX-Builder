Imports System.Text.RegularExpressions
Imports RibbonFactory.ControlInterfaces
Imports RibbonFactory.Controls
Imports RibbonFactory.RibbonAttributes
Imports RibbonFactory.Utilities
Imports stdole

Namespace Containers

	Public NotInheritable Class DropDown
		Inherits Container(Of Item)
		Implements ISelect
		Implements IEnabled
		Implements IVisible
		Implements IKeyTip
		Implements ILabel
		Implements IScreenTip
		Implements ISuperTip
		Implements IShowLabel
		Implements IImage
		Implements IShowImage
		Implements IDefaultProvider

		Private ReadOnly _attributes As AttributeSet

		Friend Sub New(attributes As AttributeSet, buttons As ICollection(Of Button), Optional tag As Object = Nothing)
			MyBase.New(New List(Of Item), tag)
			_attributes = attributes
			AddHandler _attributes.AttributeChanged, AddressOf RefreshNeeded
			Me.Buttons = buttons
		End Sub

		Public ReadOnly Property Buttons As ICollection(Of Button)

#Region "Events"

		Public Event BeforeSelectionChange As EventHandler(Of BeforeSelectionChangeEventArgs)

		Public Event SelectionChanged As EventHandler(Of SelectionChangedEventArgs)

#End Region

#Region "RibbonElement Overrides"

		Public Overrides ReadOnly Property ID As String
			Get
				Return _attributes.ReadOnlyLookup(Of String)(AttributeName.Id).GetValue()
			End Get
		End Property

		Public Overrides ReadOnly Property XML As String
			Get
				Dim regex As Regex = New Regex("(?:size|getSize)=""\w+""")

				If Buttons IsNot Nothing AndAlso Buttons.Any() Then
					Return String.Join(Environment.NewLine, $"<dropDown { _attributes }>", regex.Replace(String.Join(Environment.NewLine, Buttons), String.Empty), "</dropDown>")
				Else
					Return $"<dropDown { _attributes }/>"
				End If
			End Get
		End Property

#End Region

#Region "Container Overrides"

		Public Overrides Sub Add(item As Item)
			MyBase.Add(item)

			AddHandler item.ValueChanged, AddressOf OnChildItemChange

			If Items.Count = 1 Then
				Selected = item
			End If
		End Sub

		Public Overrides Sub Clear()
			If Items.Any() Then
				Selected = Nothing

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

				If Selected.Equals(item) Then
					Selected = Items.FirstOrDefault()
				End If

				RefreshNeeded()

				Return True
			Else
				Return False
			End If
		End Function

		Friend Overrides Sub Flatten(results As ICollection(Of RibbonElement))
			results.Add(Me)

			For Each button As Button In Buttons
				results.Add(button)
			Next
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
				Return _attributes.ReadOnlyLookup(Of KeyTip)(AttributeCategory.KeyTip).GetValue()
			End Get
			Set
				_attributes.ReadWriteLookup(Of KeyTip)(AttributeCategory.KeyTip).SetValue(Value)
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

		Public Property ShowLabel As Boolean Implements IShowLabel.ShowLabel
			Get
				Return _attributes.ReadOnlyLookup(Of Boolean)(AttributeName.ShowLabel).GetValue()
			End Get
			Set
				_attributes.ReadWriteLookup(Of Boolean)(AttributeName.GetShowLabel).SetValue(Value)
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

#Region "Action"

		Public Property Selected As Item Implements ISelect.Selected
			Get
				Return _attributes.ReadOnlyLookup(Of Item)(AttributeCategory.SelectedItemPosition).GetValue()
			End Get
			Set
				Try
					SuspendAutomaticRefreshing()

					Dim e As BeforeSelectionChangeEventArgs = New BeforeSelectionChangeEventArgs(Selected, Value, Items.ToArray())

					RaiseEvent BeforeSelectionChange(Me, e)

					If Not e.Cancel Then
						_attributes.ReadWriteLookup(Of Item)(AttributeCategory.SelectedItemPosition).SetValue(e.NewSelection)

						RaiseEvent SelectionChanged(Me, New SelectionChangedEventArgs(e.NewSelection))
					End If
				Finally
					ResumeAutomaticRefreshing()
				End Try
			End Set
		End Property

		Public Property SelectedItemIndex As Integer Implements ISelect.SelectedItemIndex
			Get
				Return Items.IndexOf(Selected)
			End Get
			Set
				Selected = Items(Value)
			End Set
		End Property

#End Region

#Region "Helpers For Control Cloning"

		Friend Function GetBeforeSelectionChangeInvocationList() As EventHandler(Of BeforeSelectionChangeEventArgs)()
			Dim e As EventHandler(Of BeforeSelectionChangeEventArgs) = BeforeSelectionChangeEvent

			If e Is Nothing Then
				Return New EventHandler(Of BeforeSelectionChangeEventArgs)() {}
			Else
				Return e.
					GetInvocationList().
					OfType(Of EventHandler(Of BeforeSelectionChangeEventArgs)).
					ToArray()
			End If
		End Function

		Friend Function GetSelectionChangedInvocationList() As EventHandler(Of SelectionChangedEventArgs)()
			Dim e As EventHandler(Of SelectionChangedEventArgs) = SelectionChangedEvent

			If e Is Nothing Then
				Return New EventHandler(Of SelectionChangedEventArgs)() {}
			Else
				Return e.
					GetInvocationList().
					OfType(Of EventHandler(Of SelectionChangedEventArgs)).
					ToArray()
			End If
		End Function

		Private Function GetDefaults() As AttributeSet Implements IDefaultProvider.GetDefaults
			Return _attributes
		End Function

#End Region

		Public Class BeforeSelectionChangeEventArgs
			Inherits EventArgs

			Public Sub New(currentSelection As Item, newSelection As Item, availableSelections As ICollection(Of Item))
				Me.CurrentSelection = currentSelection
				Me.NewSelection = newSelection
				Me.AvailableSelections = availableSelections
				Cancel = False
			End Sub

			Public ReadOnly Property AvailableSelections As ICollection(Of Item)

			Public Property Cancel As Boolean

			Public ReadOnly Property CurrentSelection As Item

			Public Property NewSelection As Item

		End Class

		Public Class SelectionChangedEventArgs
			Inherits EventArgs

			Public Sub New(currentSelection As Item)
				Me.CurrentSelection = currentSelection
			End Sub

			Public ReadOnly Property CurrentSelection As Item

		End Class

	End Class

End Namespace