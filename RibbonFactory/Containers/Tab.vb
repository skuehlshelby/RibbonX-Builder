Imports RibbonFactory.Component_Interfaces
Imports RibbonFactory.RibbonAttributes
Imports RibbonFactory.RibbonAttributes.Categories.ID
Imports RibbonFactory.RibbonAttributes.Categories.KeyTip
Imports RibbonFactory.RibbonAttributes.Categories.Label
Imports RibbonFactory.RibbonAttributes.Categories.Visible

Namespace Containers

    Public NotInheritable Class Tab
        Inherits RibbonElement
        Implements IEnumerable(Of RibbonElement)
        Implements IVisible
        Implements IKeyTip
        Implements ILabel

        Private ReadOnly _attributes As AttributeGroup
        Private ReadOnly _childElements As List(Of RibbonElement) = New List(Of RibbonElement)()

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
                Return String.Join(Environment.NewLine,$"<tab {String.Join(" ", _attributes) }>", String.Join(Environment.NewLine, _childElements), "<>")
            End Get
        End Property

        Public Property Visible As Boolean Implements IVisible.Visible
            Get
                Return _attributes.Lookup (Of Visible).GetValue()
            End Get
            Set
                _attributes.Lookup (Of GetVisible).SetValue(value)
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

        Public Property Label As String Implements ILabel.Label
            Get
                Return _attributes.Lookup (Of Label).GetValue()
            End Get
            Set
                _attributes.Lookup (Of GetLabel).SetValue(value)
                Refresh()
            End Set
        End Property

        Public Function AddGroup(group As Group) As Group
            _childElements.Add(group)
            Return group
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            Return obj.GetHashCode() = GetHashCode() AndAlso TypeOf obj Is Tab
        End Function

        Public Function GetEnumerator() As IEnumerator(Of RibbonElement) Implements IEnumerable(Of RibbonElement).GetEnumerator
            Return _childElements.GetEnumerator()
        End Function

        Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return _childElements.GetEnumerator()
        End Function
    End Class
End NameSpace