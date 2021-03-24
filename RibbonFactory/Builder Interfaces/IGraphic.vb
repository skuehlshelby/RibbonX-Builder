Imports System.Drawing
Imports RibbonFactory.Enums
Imports stdole

Namespace Builder_Interfaces
    Public Interface IGraphic(Of T)
        Overloads Function WithImage(Image As ImageMSO) As T
        Overloads Function WithImage(Image As IPictureDisp, Optional CallbackName As String = "GetImage") As T
        Overloads Function WithImage(Image As Bitmap, Optional CallbackName As String = "GetImage") As T
    End Interface
End NameSpace