
Imports System.IO
Imports System.Reflection
Imports System.Drawing
Imports RibbonFactory

<TestClass()>
Public Class ImageConversion

    <TestMethod()>
    Public Sub ToIPictureDisp()
        Dim assmbly As Assembly = Assembly.GetExecutingAssembly()
        Dim stream As Stream = assmbly.GetManifestResourceStream("UnitTests.CloudIcon.PNG")
        Dim image As Bitmap = New Bitmap(stream)

        PictureDispConverter.ImageToPictureDisp(image)
    End Sub

End Class