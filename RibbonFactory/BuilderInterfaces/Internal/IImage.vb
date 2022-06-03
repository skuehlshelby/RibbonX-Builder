Imports System.Drawing
Imports RibbonX.Enums.ImageMSO
Imports RibbonX.ComTypes.StdOle

Namespace BuilderInterfaces

    Public Interface IImage(Of T)

         Function WithImage(imagePath As String) As T

         Function WithImage(image As ImageMSO) As T
         
         Function WithImage(image As Bitmap, callback As FromControl(Of IPictureDisp)) As T

         Function WithImage(image As Icon, callback As FromControl(Of IPictureDisp)) As T

    End Interface

End NameSpace