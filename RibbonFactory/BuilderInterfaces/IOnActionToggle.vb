Namespace BuilderInterfaces

    Public Interface IOnActionToggle(Of T)
        
        Function ThatDoes(action As Action, callback As ButtonPressed) As T
    
    End Interface

End NameSpace