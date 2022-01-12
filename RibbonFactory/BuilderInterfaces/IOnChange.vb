Namespace BuilderInterfaces

    Public Interface IOnChange(Of TRibbonElement As RibbonElement, Out TBuilder)

        Function ThatDoes(action As Action(Of TRibbonElement), callback As TextChanged) As TBuilder

    End Interface

End NameSpace