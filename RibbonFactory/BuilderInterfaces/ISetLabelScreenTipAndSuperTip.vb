Imports RibbonFactory.Builders

Namespace BuilderInterfaces

    Public Interface ISetLabelScreenTipAndSuperTip(Of T As Builder)
        Function WithLabel(label As String, Optional copyToScreenTip As Boolean = True) As T

        Function WithLabel(label As String, callback As FromControl(Of String), Optional copyToScreenTip As Boolean = True) As T

        Function WithScreenTip(screenTip As String) As T

        Function WithScreenTip(screenTip As String, callback As FromControl(Of String)) As T

        Function WithSuperTip(superTip As String) As T

        Function WithSuperTip(superTip As String, callback As FromControl(Of String)) As T
    End Interface
End NameSpace