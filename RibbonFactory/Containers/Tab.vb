Public Class Tab
    Inherits RibbonElement
    Implements IDisablable
    Implements IContainer(Of Group)
    Implements IEnumerable(Of Group)

    Private ReadOnly Groups As List(Of Group) = New List(Of Group)

    Friend Sub New(ByVal DisplayName As String, Optional ByVal Tag As Object = Nothing)
        MyBase.New(RibbonElement.FabricateID(Of Tab)(), Tag)
    End Sub
    Public Overrides ReadOnly Property XML As String
        Get
            Return String.Format("<tab id=""{0}""", "{1}", "></tab>", ID, String.Join(vbNewLine & vbTab, Groups.Select(Function(G) G.XML)))
        End Get
    End Property

    Public ReadOnly Property Items As List(Of RibbonElement) Implements IContainer.Items
        Get
            Return Groups.Cast(Of RibbonElement)
        End Get
    End Property

    Public Property Enabled As Boolean Implements IDisablable.Enabled
        Get
            Throw New NotImplementedException()
        End Get
        Set(value As Boolean)
            Throw New NotImplementedException()
        End Set
    End Property

    Public Property Visible As Boolean Implements IDisablable.Visible
        Get
            Throw New NotImplementedException()
        End Get
        Set(value As Boolean)
            Throw New NotImplementedException()
        End Set
    End Property

    Public Sub Add(ParamArray Items() As Group) Implements IContainer(Of Group).Add
        Groups.AddRange(Items)
    End Sub

    Public Sub Clear() Implements IContainer.Clear
        Groups.Clear()
    End Sub

    Public Overrides Function Equals(obj As Object) As Boolean
        If obj.GetHashCode() = GetHashCode() Then
            Return TypeOf obj Is Tab
        Else
            Return False
        End If
    End Function

    Public Function GetEnumerator() As IEnumerator(Of Group) Implements IEnumerable(Of Group).GetEnumerator
        Return Items.GetEnumerator
    End Function

    Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Items.GetEnumerator
    End Function
End Class
