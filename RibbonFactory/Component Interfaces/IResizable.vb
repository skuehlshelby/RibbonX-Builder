Imports RibbonFactory.Enums

Namespace Component_Interfaces
    Public Interface IResizable
        Property Size As ControlSize
    End Interface
    Public Interface IResizable(Of T)
        Overloads Function Large() As T
        Overloads Function Large(UseCallback As Boolean, Optional CallbackName As String = "GetSize") As T
        Overloads Function Normal() As T
        Overloads Function Normal(UseCallback As Boolean, Optional CallbackName As String = "GetSize") As T
    End Interface
End NameSpace