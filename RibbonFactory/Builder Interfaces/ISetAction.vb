Imports RibbonFactory.Builders

Namespace Builder_Interfaces
    Public Interface ISetAction(OF T As Builder)
        Function ThatDoes(callback As OnAction, action As Action) As T
    End Interface
End NameSpace