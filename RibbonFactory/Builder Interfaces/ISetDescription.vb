Imports RibbonFactory.Builders

Namespace Builder_Interfaces
    Public Interface ISetDescription(Of T As Builder)
        Function WithDescription(description As String) As T

        Function WithDescription(description As String, callback As FromControl(Of String)) As T
    End Interface
End NameSpace