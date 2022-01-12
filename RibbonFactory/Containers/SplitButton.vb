Imports System.Text.RegularExpressions
Imports RibbonFactory.ControlInterfaces
Imports RibbonFactory.Enums
Imports RibbonFactory.RibbonAttributes

Namespace Containers

	Public NotInheritable Class SplitButton
		Inherits Container(Of RibbonElement)
		Implements IVisible
		Implements IEnabled
		Implements IKeyTip
		Implements ISize
		Implements IShowLabel
		Implements IDefaultProvider

		Private ReadOnly _attributes As AttributeSet

		Friend Sub New(attributes As AttributeSet, Optional tag As Object = Nothing)
			MyBase.New(New RibbonElement() {Nothing, Nothing}, tag)

			_attributes = attributes
			AddHandler _attributes.AttributeChanged, AddressOf RefreshNeeded
		End Sub

		Friend Sub New(button As RibbonElement, menu As Menu, attributes As AttributeSet, Optional tag As Object = Nothing)
			MyBase.New(New RibbonElement() {button, menu}, tag)

			_attributes = attributes
			AddHandler _attributes.AttributeChanged, AddressOf RefreshNeeded
		End Sub

		Public Overrides ReadOnly Property ID As String
			Get
				Return _attributes.ReadOnlyLookup(Of String)(AttributeName.Id).GetValue()
			End Get
		End Property

		Public Overrides ReadOnly Property XML As String
			Get
				Dim regex As Regex = New Regex("(?:size|getSize|visible|getVisible)=""\w+""")

				If Items.Any() Then
					Return _
						regex.Replace(String.Join(Environment.NewLine, $"<splitButton { _attributes }>",
									String.Join(Environment.NewLine, Items),
									"</splitButton>"), String.Empty)

				Else
					Return $"<splitButton { _attributes } />"
				End If
			End Get
		End Property

		Public ReadOnly Property Button As RibbonElement
			Get
				Return Items(0)
			End Get
		End Property

		Public ReadOnly Property Menu As Menu
			Get
				Return DirectCast(Items(1), Menu)
			End Get
		End Property

		Public Property Visible As Boolean Implements IVisible.Visible
			Get
				Return _attributes.ReadOnlyLookup(Of Boolean)(AttributeName.Visible).GetValue()
			End Get
			Set
				_attributes.ReadWriteLookup(Of Boolean)(AttributeName.GetVisible).SetValue(Value)
			End Set
		End Property

		Public Property Enabled As Boolean Implements IEnabled.Enabled
			Get
				Return _attributes.ReadOnlyLookup(Of Boolean)(AttributeName.Enabled).GetValue()
			End Get
			Set
				_attributes.ReadWriteLookup(Of Boolean)(AttributeName.GetEnabled).SetValue(Value)
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

		Public Property Size As ControlSize Implements ISize.Size
			Get
				Return _attributes.ReadOnlyLookup(Of ControlSize)(AttributeName.Size).GetValue()
			End Get
			Set
				_attributes.ReadWriteLookup(Of ControlSize)(AttributeName.GetSize).SetValue(Value)
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

		Private Function GetDefaults() As AttributeSet Implements IDefaultProvider.GetDefaults
			Return _attributes
		End Function

	End Class

End Namespace