Imports RibbonX.Images
Imports RibbonX.SimpleTypes

Namespace Properties

    Public Interface IRibbonElementProperty
    End Interface

    Public Interface IBoxStyle
        Inherits IRibbonElementProperty
        ReadOnly Property BoxStyle As BoxStyle
    End Interface

    Public Interface IClickable
        Inherits IRibbonElementProperty
        Sub Click()
    End Interface

    Public Interface IDescription
        Inherits IRibbonElementProperty
        Property Description As String
    End Interface

    Public Interface IEnabled
        Inherits IRibbonElementProperty
        Property Enabled As Boolean
    End Interface

    Public Interface IImage
        Inherits IRibbonElementProperty
        Property Image As RibbonImage
    End Interface

    Public Interface IItemDimensions
        Inherits IRibbonElementProperty
        Property ItemHeight As Integer
        Property ItemWidth As Integer
    End Interface

    Public Interface IItemSize
        Inherits IRibbonElementProperty
        ReadOnly Property ItemSize As ControlSize
    End Interface

    Public Interface IKeyTip
        Inherits IRibbonElementProperty
        Property KeyTip As KeyTip
    End Interface

    Public Interface ILabel
        Inherits IRibbonElementProperty
        Property Label As String
    End Interface

    Public Interface IMaxLength
        Inherits IRibbonElementProperty
        ReadOnly Property MaxLength As Byte
    End Interface

    Public Interface IChecked
        Inherits IRibbonElementProperty
        Property IsChecked As Boolean
    End Interface

    Public Interface IRowsAndColumns
        Inherits IRibbonElementProperty
        ReadOnly Property Rows As Integer
        ReadOnly Property Columns As Integer
    End Interface

    Public Interface IScreenTip
        Inherits IRibbonElementProperty
        Property ScreenTip As String
    End Interface

    Public Interface ISelect
        Inherits IRibbonElementProperty
        Property Selected As IItem
        Property SelectedItemIndex As Integer
    End Interface

    Public Interface IShowImage
        Inherits IRibbonElementProperty
        Property ShowImage As Boolean
    End Interface

    Public Interface IShowLabel
        Inherits IRibbonElementProperty
        Property ShowLabel As Boolean
    End Interface

    Public Interface ISize
        Inherits IRibbonElementProperty
        Property Size As ControlSize
    End Interface

    Public Interface ISuperTip
        Inherits IRibbonElementProperty
        Property SuperTip As String
    End Interface

    Public Interface IText
        Inherits IRibbonElementProperty
        Property Text As String
    End Interface

    Public Interface ITitle
        Inherits IRibbonElementProperty
        Property Title As String
    End Interface

    Public Interface IVisible
        Inherits IRibbonElementProperty
        Property Visible As Boolean
    End Interface

    Public Interface IItemTemplateable
        Sub AddTemplate(template As IItemTemplate)
        Sub AddTemplatedItem(item As Object)
        Sub RemoveTemplatedItem(item As Object)
    End Interface

End Namespace