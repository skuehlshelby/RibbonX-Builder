Imports System.Drawing

Namespace Builders

    Public Interface IItemBuilder

        Function WithId(id As String) As IItemBuilder

        Function WithLabel(label As String, Optional copyToScreentip As Boolean = True) As IItemBuilder

        Function WithScreenTip(screenTip As String) As IItemBuilder

        Function WithSuperTip(superTip As String) As IItemBuilder

        Function WithImage(image As Bitmap) As IItemBuilder

        Function WithImage(image As Icon) As IItemBuilder

    End Interface

End Namespace
