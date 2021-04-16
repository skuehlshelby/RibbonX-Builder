Imports System.Drawing
Imports RibbonFactory.Builders
Imports RibbonFactory.Enums
Imports RibbonFactory.Enums.ImageMSO
Imports stdole

Namespace Builder_Interfaces
    Public Interface ISetImage(Of T As Builder)
        Overloads Function WithImage(image As ImageMSO) As T
        Overloads Function WithImage(image As IPictureDisp) As T
        Overloads Function WithImage(image As Bitmap) As T
        Overloads Function WithImage(image As IPictureDisp, callback As FromControl(Of IPictureDisp)) As T
        Overloads Function WithImage(image As Bitmap, callback As FromControl(Of IPictureDisp)) As T
    End Interface
End NameSpace