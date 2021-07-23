Imports RibbonFactory.Component_Interfaces
Imports RibbonFactory.RibbonAttributes


Namespace Containers

    Public NotInheritable Class Tab
        Inherits RibbonElement
        Implements IEnumerable(Of RibbonElement)
        Implements IVisible
        Implements IKeyTip
        Implements ILabel

        Private ReadOnly _attributes As AttributeGroup
        Private ReadOnly _groups As List(Of Group) = New List(Of Group)()

        Friend Sub New(attributes As AttributeGroup, Optional tag As Object = Nothing)
            MyBase.New(tag)
            _attributes = attributes
        End Sub

        Public ReadOnly Property Groups As IReadOnlyList(Of Group)
            Get
                Return _groups
            End Get
        End Property

        Public Function AddGroup(group As Group) As Group
            _groups.Add(group)
            Return group
        End Function

        Public Overrides ReadOnly Property ID As String
            Get
                Return _attributes.ReadOnlyLookup(Of String)(AttributeName.Id).GetValue()
            End Get
        End Property

        Public Overrides ReadOnly Property XML As String
            Get
                Return _
                    String.Join(Environment.NewLine, $"<tab { _attributes }>",
                                String.Join(Environment.NewLine, _groups), $"</tab>")
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

        Public Overrides Function Equals(obj As Object) As Boolean
            Return obj.GetHashCode() = GetHashCode() AndAlso TypeOf obj Is Tab
        End Function

        Public Function GetEnumerator() As IEnumerator(Of RibbonElement) Implements IEnumerable(Of RibbonElement).GetEnumerator
            Return _groups.GetEnumerator()
        End Function

        Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return _groups.GetEnumerator()
        End Function

    End Class
End NameSpace