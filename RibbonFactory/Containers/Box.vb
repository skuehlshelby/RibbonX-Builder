Public Class Box
    Inherits RibbonElement
    Implements IContainer(Of RibbonElement)

    Private Structure TThis
        Public Orientation As BoxStyle
    End Structure

    Private This As TThis

    Public Sub New(ByVal Orientation As BoxStyle, Optional ByVal Tag As Object = Nothing)
        MyBase.New(RibbonElement.FabricateID(Of Group), Tag)

        This = New TThis With {
            .Orientation = Orientation
        }
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
