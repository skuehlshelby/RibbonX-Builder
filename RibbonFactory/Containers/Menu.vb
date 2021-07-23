Imports RibbonFactory.Component_Interfaces
Imports RibbonFactory.Enums
Imports RibbonFactory.RibbonAttributes


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
                Return _attributes.ReadOnlyLookup(Of String)(AttributeName.Id).GetValue()
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

        Public Property SuperTip As String Implements ISupertip.Supertip
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

        Public Property Size As ControlSize Implements ISize.Size
            Get
                Return _attributes.ReadOnlyLookup(Of ControlSize)(AttributeName.Size).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of ControlSize)(AttributeName.GetSize).SetValue(Value)
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