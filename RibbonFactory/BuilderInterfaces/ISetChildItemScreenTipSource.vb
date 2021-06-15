Imports RibbonFactory.Builders

Namespace BuilderInterfaces
    Public Interface ISetChildItemScreenTipSource(Of T As Builder)

        Function GetItemScreenTipFrom(screenTipSource As FromControlAndIndex(Of String)) As T

    End Interface
End Namespace