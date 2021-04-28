Imports RibbonFactory.Component_Interfaces
Imports RibbonFactory.Enums
Imports RibbonFactory.RibbonAttributes
Imports RibbonFactory.RibbonAttributes.Categories.Description
Imports RibbonFactory.RibbonAttributes.Categories.Screentip
Imports RibbonFactory.RibbonAttributes.Categories.Enabled
Imports RibbonFactory.RibbonAttributes.Categories.ID
Imports RibbonFactory.RibbonAttributes.Categories.Image
Imports RibbonFactory.RibbonAttributes.Categories.KeyTip
Imports RibbonFactory.RibbonAttributes.Categories.Label
Imports RibbonFactory.RibbonAttributes.Categories.Size
Imports RibbonFactory.RibbonAttributes.Categories.SuperTip
Imports RibbonFactory.RibbonAttributes.Categories.Visible
Imports stdole

Namespace Containers

    Public NotInheritable Class Menu
        Inherits RibbonElement
        Implements IEnumerable(Of RibbonElement)
        Implements IReadonlyCollection(Of RibbonElement)
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

        Private ReadOnly _attributes As AttributeGroup
        Private ReadOnly _items As List(Of RibbonElement) = New List(Of RibbonElement)()

        Friend Sub New(attributes As AttributeGroup, Optional tag As Object = Nothing)
            MyBase.New(tag)
            _attributes = attributes
        End Sub

        Public Overrides ReadOnly Property ID As String
            Get
                Return _attributes.Lookup(Of Id).GetValue()
            End Get
        End Property

        Public Overrides ReadOnly Property XML As String
            Get
                Return _
                    String.Join(Environment.NewLine, $"<{NameOf(Menu).ToLower()} {String.Join(" ", _attributes) }>",
                                String.Join(Environment.NewLine, _items), $"</{NameOf(Menu).ToLower()}>")
            End Get
        End Property

        Public Property Visible As Boolean Implements IVisible.Visible
            Get
                Return _attributes.Lookup(Of Visible).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetVisible).SetValue(Value)
                Refresh()
            End Set
        End Property

        Public Property Enabled As Boolean Implements IEnabled.Enabled
            Get
                Return _attributes.Lookup(Of Enabled).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetEnabled).SetValue(Value)
                Refresh()
            End Set
        End Property

        Public Property Label As String Implements ILabel.Label
            Get
                Return _attributes.Lookup(Of Label).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetLabel).SetValue(Value)
                Refresh()
            End Set
        End Property

        Public Property ScreenTip As String Implements IScreenTip.ScreenTip

            Get
                Return _attributes.Lookup(Of Screentip).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetScreenTip).SetValue(Value)
                Refresh()
            End Set
        End Property

        Public Property SuperTip As String Implements ISuperTip.SuperTip
            Get
                Return _attributes.Lookup(Of SuperTip).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetSuperTip).SetValue(value)
                Refresh()
            End Set
        End Property

        Public Property KeyTip As KeyTip Implements IKeyTip.KeyTip
            Get
                Return _attributes.Lookup (Of Categories.KeyTip.KeyTip).GetValue()
            End Get
            Set
                _attributes.Lookup (Of GetKeyTip).SetValue(value)
                Refresh()
            End Set
        End Property

        Public Property Image As IPictureDisp Implements IImage.Image
            Get
                Return _attributes.Lookup(Of GetImage).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetImage).SetValue(value)
                Refresh()
            End Set
        End Property

        Public ReadOnly Property IsCustom As Boolean Implements IImage.IsCustom
            Get
                Return TypeOf _attributes.Lookup(Of ImageBase) Is GetImage
            End Get
        End Property

        Public Property ShowImage As Boolean Implements IShowImage.ShowImage
            Get
                Return _attributes.Lookup(Of ShowImage).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetShowImage).SetValue(value)
                Refresh()
            End Set
        End Property

        Public Property Description As String Implements IDescription.Description
            Get
                Return _attributes.Lookup(Of Description).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetDescription).SetValue(value)
                Refresh()
            End Set
        End Property

        Public Property ShowLabel As Boolean Implements IShowLabel.ShowLabel
            Get
                Return _attributes.Lookup(Of ShowLabel).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetShowLabel).SetValue(value)
                Refresh()
            End Set
        End Property

        Public Property Size As ControlSize Implements ISize.Size
            Get
                Return _attributes.Lookup(Of Size).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetSize).SetValue(value)
                Refresh()
            End Set
        End Property

        Public ReadOnly Property Count As Integer Implements IReadOnlyCollection(Of RibbonElement).Count
            Get
                Return DirectCast(_items, IReadOnlyCollection(Of RibbonElement)).Count
            End Get
        End Property

        Public Function GetEnumerator() As IEnumerator(Of RibbonElement) Implements IEnumerable(Of RibbonElement).GetEnumerator
            Return DirectCast(_items, IEnumerable(Of RibbonElement)).GetEnumerator()
        End Function

        Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return DirectCast(_items, IEnumerable).GetEnumerator()
        End Function
    End Class

End Namespace