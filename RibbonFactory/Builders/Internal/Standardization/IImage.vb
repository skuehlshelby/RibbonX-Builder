Imports System.Drawing
Imports RibbonX.ComTypes.StdOle
Imports RibbonX.Callbacks
Imports RibbonX.Images.BuiltIn

Namespace Builders.Internal.Standardization

    Public Interface IImage(Of T)

        Function WithImage(imagePath As String) As T

        Function WithImage(image As ImageMSO) As T

        Function WithImage(image As Bitmap, callback As FromControl(Of IPictureDisp)) As T

        Function WithImage(image As Icon, callback As FromControl(Of IPictureDisp)) As T

    End Interface

End Namespace