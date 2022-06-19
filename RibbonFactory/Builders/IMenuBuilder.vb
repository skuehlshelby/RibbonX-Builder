Imports RibbonX.Builders.Internal.Standardization

Namespace Builders

    Public Interface IMenuBuilder
        Inherits IInsert(Of IMenuBuilder)
        Inherits IEnabled(Of IMenuBuilder)
        Inherits IVisible(Of IMenuBuilder)
        Inherits ILabelScreenTipSuperTip(Of IMenuBuilder)
        Inherits IShowLabel(Of IMenuBuilder)
        Inherits IImage(Of IMenuBuilder)
        Inherits IShowImage(Of IMenuBuilder)
        Inherits IItemSize(Of IMenuBuilder)
        Inherits IDescription(Of IMenuBuilder)
        Inherits ISize(Of IMenuBuilder)
        Inherits IKeyTip(Of IMenuBuilder)
    End Interface

End Namespace