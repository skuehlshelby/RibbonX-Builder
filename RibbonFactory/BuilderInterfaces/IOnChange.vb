Namespace BuilderInterfaces

    Public Interface IOnChange(Of T)
        
        Function ThatDoes(action As Action, callback As TextChanged) As T
    
    End Interface

End NameSpace