Namespace Builder_Interfaces
    Public Interface IDescribable(Of Out T)
        Overloads Function WithLabel(Label As String, Optional copyToScreenTip As Boolean = True) As T
        Overloads Function WithLabel(Label As String, Callback As FromControl(Of String), Optional CopyToScreentip As Boolean = True) As T
        Overloads Function WithScreenTip(ScreenTip As String) As T
        Overloads Function WithScreenTip(ScreenTip As String, Callback As FromControl(Of String)) As T
        Overloads Function WithSupertip(Supertip As String) As T
        Overloads Function WithSupertip(Supertip As String, Callback As FromControl(Of String)) As T
    End Interface
End NameSpace