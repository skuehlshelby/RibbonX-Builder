Public Class Group
    Inherits RibbonElement
    Implements IContainer(Of RibbonElement)

    Public Sub New(Optional ByVal Tag As Object = Nothing)
        MyBase.New(RibbonElement.FabricateID(Of Group), Tag)
    End Sub
    Public Overrides ReadOnly Property XML As String
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public ReadOnly Property Items As List(Of RibbonElement) Implements IContainer.Items
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public Sub Add(ParamArray Items() As RibbonElement) Implements IContainer(Of RibbonElement).Add
        Throw New NotImplementedException()
    End Sub

    Public Sub Clear() Implements IContainer.Clear
        Throw New NotImplementedException()
    End Sub

    Public Overrides Function Equals(obj As Object) As Boolean
        Throw New NotImplementedException()
    End Function
End Class
