Imports RibbonFactory.Builders

Namespace BuilderInterfaces
    Public Interface ISetToggleAction(Of T As Builder)
        
        Function ThatDoes(callback As ButtonPressed, action As Action) As T
        
    
    End Interface
End NameSpace