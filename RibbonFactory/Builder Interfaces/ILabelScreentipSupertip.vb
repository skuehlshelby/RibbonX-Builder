Imports RibbonFactory.Builders

Namespace Builder_Interfaces

    Public Interface ISetLabelScreenTipAndSuperTip(Of T As RibbonElement)
        Function WithLabel(label As String, Optional copyToScreenTip As Boolean = True) As Builder(Of T)

        Function WithLabel(label As String, callback As FromControl(Of String), Optional copyToScreenTip As Boolean = True) As Builder(Of T)

        Function WithScreenTip(screenTip As String) As Builder(Of T)

        Function WithScreenTip(screenTip As String, callback As FromControl(Of String)) As Builder(Of T)

        Function WithSuperTip(superTip As String) As Builder(Of T)

        Function WithSuperTip(superTip As String, callback As FromControl(Of String)) As Builder(Of T)
    End Interface
End NameSpace