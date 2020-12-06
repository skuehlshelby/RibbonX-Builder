Public Interface IResizable
    Property Size As ControlSize
End Interface
Public Interface IResizable(Of T)
    Overloads Function Large() As T
    Overloads Function Large(ByVal UseCallback As Boolean, Optional ByVal CallbackName As String = "GetSize") As T
    Overloads Function Normal() As T
    Overloads Function Normal(ByVal UseCallback As Boolean, Optional ByVal CallbackName As String = "GetSize") As T
End Interface
