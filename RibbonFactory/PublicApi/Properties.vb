Imports RibbonX.Api
Imports RibbonX.Images
Imports RibbonX.SimpleTypes

Namespace FutureApi

    Public Interface IBoxStyle
        ReadOnly Property BoxStyle As BoxStyle
    End Interface

    Public Interface IClickable
        Sub Click()
    End Interface

    Public Interface IDescription
        Property Description As String
    End Interface

    Public Interface IEnabled
        Property Enabled As Boolean
    End Interface

    Public Interface IImage
        Property Image As RibbonImage
    End Interface

    Public Interface IItemDimensions
        Property ItemHeight As Integer
        Property ItemWidth As Integer
    End Interface

    Public Interface IItemSize
        ReadOnly Property ItemSize As ControlSize
    End Interface

    Public Interface IKeyTip
        Property KeyTip As KeyTip
    End Interface

    Public Interface ILabel
        Property Label As String
    End Interface

    Public Interface IMaxLength
        ReadOnly Property MaxLength As Byte
    End Interface

    Public Interface IChecked
        Property IsChecked As Boolean
    End Interface

    Public Interface IRowsAndColumns
        ReadOnly Property Rows As Integer
        ReadOnly Property Columns As Integer
    End Interface

    Public Interface IScreenTip
        Property ScreenTip As String
    End Interface

    Public Interface ISelect
        Property Selected As IItem
        Property SelectedItemIndex As Integer
    End Interface

    Public Interface IShowImage
        Property ShowImage As Boolean
    End Interface

    Public Interface IShowLabel
        Property ShowLabel As Boolean
    End Interface

    Public Interface ISize
        Property Size As ControlSize
    End Interface

    Public Interface ISuperTip
        Property SuperTip As String
    End Interface

    Public Interface IText
        Property Text As String
    End Interface

    Public Interface ITitle
        Property Title As String
    End Interface

    Public Interface IVisible
        Property Visible As Boolean
    End Interface

End Namespace