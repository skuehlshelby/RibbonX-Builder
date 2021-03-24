Imports RibbonFactory.Builders

Namespace Builder_Interfaces
    Public Interface ISetDescription(Of T As RibbonElement)
        Function WithDescription(description As String) As Builder(Of T)

        Function WithDescription(description As String, callback As FromControl(Of String)) As Builder(Of T)
    End Interface
End NameSpace