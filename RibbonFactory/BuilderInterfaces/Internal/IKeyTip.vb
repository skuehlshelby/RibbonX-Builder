Namespace BuilderInterfaces

    Public Interface IKeyTip(Of T)

        Function WithKeyTip(keyTip As KeyTip) As T

        Function WithKeyTip(keyTip As KeyTip, callback As FromControl(Of KeyTip)) As T

    End Interface


End NameSpace