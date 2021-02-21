Namespace Builder_Interfaces
    Public Interface IVeryDescribable(Of Out T)
        Overloads Function WithDescription(Description As String) As T
        Overloads Function WithDescription(Description As String, Callback As FromControl(Of String)) As T
    End Interface
End NameSpace