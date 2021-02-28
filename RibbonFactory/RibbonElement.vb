Public MustInherit Class RibbonElement
    Protected Sub New(Optional tag As Object = Nothing)
        Me.Tag = tag
    End Sub

    Public MustOverride ReadOnly Property ID As String

    Public Property Tag As Object

    Public MustOverride ReadOnly Property XML As String

    Public Overrides Function ToString() As String
        Return XML
    End Function

    Public MustOverride Overrides Function Equals(obj As Object) As Boolean

    Public Overrides Function GetHashCode() As Integer
        Return ID.GetHashCode()
    End Function
End Class
