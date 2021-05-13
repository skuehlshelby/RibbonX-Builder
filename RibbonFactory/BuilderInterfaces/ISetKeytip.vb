Imports RibbonFactory.Builders

Namespace BuilderInterfaces

    Public Interface ISetKeyTip(Of T As Builder)
        Function WithKeyTip(keyTip As KeyTip) As T
        Function WithKeyTip(keyTip As KeyTip, callback As FromControl(Of KeyTip)) As T
    End Interface

End NameSpace