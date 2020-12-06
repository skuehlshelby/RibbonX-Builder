Public Interface IDisablable(Of Out T)
    Overloads Function Invisible() As T
    Overloads Function Invisible(ByVal Callback As CallbackSignatures.FromControl(Of Boolean)) As T
    Overloads Function Disabled() As T
    Overloads Function Disabled(ByVal Callback As CallbackSignatures.FromControl(Of Boolean)) As T
End Interface
