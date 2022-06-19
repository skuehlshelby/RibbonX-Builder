Imports RibbonX.Callbacks
Imports RibbonX.Controls.EventArgs
Imports RibbonX.Builders.Internal.Standardization

Namespace Builders

    Public Interface IEditBoxBuilder
        Inherits IID(Of IEditBoxBuilder)
        Inherits IEnabled(Of IEditBoxBuilder)
        Inherits IVisible(Of IEditBoxBuilder)
        Inherits IInsert(Of IEditBoxBuilder)
        Inherits ILabelScreenTipSuperTip(Of IEditBoxBuilder)
        Inherits IKeyTip(Of IEditBoxBuilder)
        Inherits IImage(Of IEditBoxBuilder)
        Inherits IShowLabel(Of IEditBoxBuilder)
        Inherits IMaxLength(Of IEditBoxBuilder)
        Inherits ISizeString(Of IEditBoxBuilder)

        Function WithText(text As String, getText As FromControl(Of String), setText As TextChanged) As IEditBoxBuilder
        Function BeforeTextChange(ParamArray actions() As EventHandler(Of BeforeTextChangeEventArgs)) As IEditBoxBuilder
        Function AfterTextChange(ParamArray actions() As EventHandler(Of TextChangeEventArgs)) As IEditBoxBuilder

    End Interface

End Namespace


