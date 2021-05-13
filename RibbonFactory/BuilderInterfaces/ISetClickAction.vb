Imports RibbonFactory.Builders

Namespace BuilderInterfaces
    Public Interface ISetClickAction(Of T As Builder)
        Function ThatDoes(callback As OnAction, action As Action) As T
    End Interface
End NameSpace