Friend Interface IControlProperty
    ReadOnly Property Governs As Attributes.AttributeGroup
    ReadOnly Property XML As String
    Function ToString() As String
    Function Equals(obj As Object) As Boolean
    Function GetHashCode() As Integer
End Interface
Friend Interface IControlProperty(Of T)
    Inherits IControlProperty
    Function GetValue() As T
End Interface
Friend Interface IDynamicControlProperty(Of T)
    Inherits IControlProperty(Of T)
    Sub SetValue(Value As T)
End Interface