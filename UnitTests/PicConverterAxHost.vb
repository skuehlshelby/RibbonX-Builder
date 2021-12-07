Imports System.Drawing
Imports System.Windows.Forms
Imports stdole

Public Class PicConverterAxHost
    Inherits AxHost

    Public Sub New()
        MyBase.New("c932ba85-4374-101b-a56c-00aa003668dc")
    End Sub

    Public Function ToIPictureDisp(image As Image) As IPictureDisp
        Return DirectCast(GetIPictureDispFromPicture(image), IPictureDisp)
    End Function
End Class