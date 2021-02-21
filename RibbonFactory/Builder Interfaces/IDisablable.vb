Namespace Builder_Interfaces
    Public Interface IDisablable(Of Out T)
        Overloads Function Invisible() As T
        Overloads Function Invisible(Callback As FromControl(Of Boolean)) As T
        Overloads Function Disabled() As T
        Overloads Function Disabled(Callback As FromControl(Of Boolean)) As T
    End Interface
End NameSpace