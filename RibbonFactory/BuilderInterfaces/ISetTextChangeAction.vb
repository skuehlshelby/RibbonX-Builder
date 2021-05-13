Imports RibbonFactory.Builders

Namespace BuilderInterfaces
    Public Interface ISetTextChangeAction(Of T As Builder)
        
        Function ThatDoes(callback As TextChanged, action As Action) As T
    
    End Interface
End NameSpace