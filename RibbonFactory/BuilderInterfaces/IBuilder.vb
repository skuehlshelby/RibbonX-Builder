Namespace BuilderInterfaces

    Public Interface IBuilder(Of Out TRibbonElement As RibbonElement)
        
        Function Build(Optional tag As Object = Nothing) As TRibbonElement

    End Interface

End NameSpace