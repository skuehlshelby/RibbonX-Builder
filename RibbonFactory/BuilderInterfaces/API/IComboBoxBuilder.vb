Imports RibbonFactory.Controls
Imports RibbonFactory.Controls.Events

Namespace BuilderInterfaces.API

    Public Interface IComboBoxBuilder
        Inherits IID(Of IComboBoxBuilder)
        Inherits IEnabled(Of IComboBoxBuilder)
        Inherits IVisible(Of IComboBoxBuilder)
        Inherits ILabelScreenTipSuperTip(Of IComboBoxBuilder)
        Inherits IShowLabel(Of IComboBoxBuilder)
        Inherits IKeyTip(Of IComboBoxBuilder)
        Inherits IImage(Of IComboBoxBuilder)
        Inherits IShowImage(Of IComboBoxBuilder)
        Inherits IMaxLength(Of IComboBoxBuilder)
        Inherits ISizeString(Of IComboBoxBuilder)
        Inherits IInsert(Of IComboBoxBuilder)
        Inherits IGetItemId(Of IComboBoxBuilder)
        Inherits IGetItemCount(Of IComboBoxBuilder)
        Inherits IGetItemLabel(Of IComboBoxBuilder)
        Inherits IGetItemScreentip(Of IComboBoxBuilder)
        Inherits IGetItemSupertip(Of IComboBoxBuilder)
        Inherits IGetItemImage(Of IComboBoxBuilder)
        Inherits IShowItemImage(Of IComboBoxBuilder)

        Function WithText(text As String, getText As FromControl(Of String), setText As TextChanged) As IComboBoxBuilder
        Function BeforeTextChange(ParamArray actions() As Action(Of BeforeTextChangeEventArgs)) As IComboBoxBuilder
        Function OnTextChange(ParamArray actions() As Action(Of TextChangeEventArgs)) As IComboBoxBuilder
    End Interface

End Namespace