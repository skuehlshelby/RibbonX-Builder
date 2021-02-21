Imports System.Drawing
Imports RibbonFactory.Enums
Imports stdole

Namespace Component_Interfaces
    Public Interface IGraphic
        Property Image As Object
        ReadOnly Property IsCustom As Boolean
        Property ShowImage As Boolean
    End Interface
    Public Interface IGraphic(Of T)
        Overloads Function WithImage(Image As ImageMSO) As T
        Overloads Function WithImage(Image As IPictureDisp, Optional CallbackName As String = "GetImage") As T
        Overloads Function WithImage(Image As Bitmap, Optional CallbackName As String = "GetImage") As T
    End Interface
End NameSpace