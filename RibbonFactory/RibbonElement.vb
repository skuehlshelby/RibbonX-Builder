Public MustInherit Class RibbonElement
    Implements IEquatable(Of RibbonElement)
    
    Public Event ValueChanged(sender As Object, e As ValueChangedEventArgs)

    Protected Sub OnValueChanged()
        RaiseEvent ValueChanged(Me, New ValueChangedEventArgs(ID))
    End Sub

    Protected Friend Sub New(Optional tag As Object = Nothing)
        Me.Tag = tag
    End Sub

    Public MustOverride ReadOnly Property ID As String

    Public Property Tag As Object

    Public MustOverride ReadOnly Property XML As String

    Public Overrides Function ToString() As String
        Return XML
    End Function

    Public  Overrides Function GetHashCode() As Integer
        Return ID.GetHashCode()
    End Function

    Public Overloads Overrides Function Equals(obj As Object) As Boolean
        Return Equals(TryCast(obj, RibbonElement))
    End Function
    
    Public Overloads Function Equals(other As RibbonElement) As Boolean Implements IEquatable(Of RibbonElement).Equals
        Return other IsNot Nothing AndAlso other.ID.Equals(ID)
    End Function

End Class
