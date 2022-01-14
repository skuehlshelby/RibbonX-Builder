Imports System.Text.RegularExpressions
Imports RibbonFactory.ControlInterfaces
Imports RibbonFactory.Enums
Imports RibbonFactory.RibbonAttributes
Imports stdole

Namespace Containers

	Public NotInheritable Class Menu
		Inherits Container(Of RibbonElement)
		Implements IVisible
		Implements IEnabled
		Implements ILabel
		Implements IShowLabel
		Implements IDescription
		Implements IScreenTip
		Implements ISuperTip
		Implements IKeyTip
		Implements IImage
		Implements IShowImage
		Implements ISize
		Implements IItemSize
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

				Return String.Join(Environment.NewLine, $"<menu { _attributes }>", regex.Replace(String.Join(Environment.NewLine, Items), String.Empty), "</menu>")
			End Get
		End Property

		Public Property Visible As Boolean Implements IVisible.Visible
			Get
				Return _attributes.ReadOnlyLookup(Of Boolean)(AttributeCategory.Visibility).GetValue()
			End Get
			Set
				_attributes.ReadWriteLookup(Of Boolean)(AttributeCategory.Visibility).SetValue(Value)
			End Set
		End Property

		Public Property Enabled As Boolean Implements IEnabled.Enabled
			Get
				Return _attributes.ReadOnlyLookup(Of Boolean)(AttributeCategory.Enabled).GetValue()
			End Get
			Set
				_attributes.ReadWriteLookup(Of Boolean)(AttributeCategory.Enabled).SetValue(Value)
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

		Public Property KeyTip As KeyTip Implements IKeyTip.KeyTip
			Get
				Return _attributes.ReadOnlyLookup(Of KeyTip)(AttributeName.Keytip).GetValue()
			End Get
			Set
				_attributes.ReadWriteLookup(Of KeyTip)(AttributeName.GetKeytip).SetValue(Value)

			End Set
		End Property

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

		Public Property Description As String Implements IDescription.Description
			Get
				Return _attributes.ReadOnlyLookup(Of String)(AttributeName.Description).GetValue()
			End Get
			Set
				_attributes.ReadWriteLookup(Of String)(AttributeName.GetDescription).SetValue(Value)
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

		Public ReadOnly Property ItemSize As ControlSize Implements IItemSize.ItemSize
			Get
				Return _attributes.ReadOnlyLookup(Of ControlSize)(AttributeName.ItemSize).GetValue()
			End Get
		End Property

		Public Property Size As ControlSize Implements ISize.Size
			Get
				Return _attributes.ReadOnlyLookup(Of ControlSize)(AttributeName.Size).GetValue()
			End Get
			Set
				_attributes.ReadWriteLookup(Of ControlSize)(AttributeName.GetSize).SetValue(Value)
			End Set
		End Property

		Private Function GetDefaults() As AttributeSet Implements IDefaultProvider.GetDefaults
			Return _attributes
		End Function

	End Class

End Namespace