Namespace BuilderInterfaces

    Public Interface IOnActionClick(Of TRibbonElement As RibbonElement, Out TBuilder)

        Function ThatDoes(action As Action(Of TRibbonElement), callback As OnAction) As TBuilder

    End Interface

End NameSpace