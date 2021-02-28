Imports RibbonFactory.Component_Interfaces
Imports RibbonFactory.RibbonAttributes

Namespace Containers

    Public NotInheritable Class Tab
        Inherits RibbonElement
        Implements IEnable
        Implements IEnumerable(Of RibbonElement)

        Private ReadOnly _displayName As String
        Private ReadOnly _groups As List(Of Group) = new List(Of Group)

        Friend Sub New(displayName As String, Optional tag As Object = Nothing)
            MyBase.New(tag)
            _displayName = displayName
        End Sub

        Public Function AddGroup() As Group
            Dim group As Group = New Group(New AttributeGroup())
            _groups.Add(group)
            Return group
        End Function

        Public Overrides ReadOnly Property ID As String

        Public Overrides ReadOnly Property XML As String
            Get
                Return _
                    $"<tab>{String.Join(vbNewLine + vbTab + vbTab + vbTab + vbTab, _groups.Select(Function(g) g.XML))}</tab>"
            End Get
        End Property

        Public Property Enabled As Boolean Implements IEnable.Enabled
            Get
                Throw New NotImplementedException()
            End Get
            Set(value As Boolean)
                Throw New NotImplementedException()
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