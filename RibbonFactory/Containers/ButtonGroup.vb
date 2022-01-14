Imports System.Text.RegularExpressions
Imports RibbonFactory.ControlInterfaces
Imports RibbonFactory.RibbonAttributes

Namespace Containers

	Public NotInheritable Class ButtonGroup
		Inherits Container(Of RibbonElement)
		Implements IVisible
		Implements IDefaultProvider

		Private ReadOnly _attributes As AttributeSet

		Friend Sub New(items As ICollection(Of RibbonElement), attributes As AttributeSet, Optional tag As Object = Nothing)
			MyBase.New(items, tag)
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
				Dim regex As Regex = New Regex("(?:size|getSize)=""\w+""")

				Return String.Join(Environment.NewLine, $"<buttonGroup { _attributes }>", regex.Replace(String.Join(Environment.NewLine, Items), String.Empty), "</buttonGroup>")
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

		Private Function GetDefaults() As AttributeSet Implements IDefaultProvider.GetDefaults
			Return _attributes
		End Function

	End Class

End Namespace