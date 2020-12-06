Imports Bitmap = System.Drawing.Bitmap
Imports IPictureDisp = stdole.IPictureDisp
Public Interface IGraphic
    Property Image As Object
    ReadOnly Property IsCustom As Boolean
    Property ShowImage As Boolean
End Interface
Public Interface IGraphic(Of T)
    Overloads Function WithImage(ByVal Image As ImageMSO) As T
    Overloads Function WithImage(ByVal Image As IPictureDisp, Optional ByVal CallbackName As String = "GetImage") As T
    Overloads Function WithImage(ByVal Image As Bitmap, Optional ByVal CallbackName As String = "GetImage") As T
End Interface
