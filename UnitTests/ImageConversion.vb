Imports System.IO
Imports System.Reflection
Imports System.Drawing
Imports RibbonFactory
Imports stdole

<TestClass()>
Public Class ImageConversion

    <TestMethod()>
    Public Sub ToIPictureDisp()
        Dim assembly As Assembly = Assembly.GetExecutingAssembly()
        Dim stream As Stream = assembly.GetManifestResourceStream("UnitTests.CloudIcon.PNG")
        Dim image As Bitmap = New Bitmap(stream)

        Dim pic As IPictureDisp = ConvertToIPictureDisplay(image)

        Assert.IsNotNull(pic)
    End Sub

    <TestMethod()>
    Public Sub ToIPictureDispWithAxHost()
        Dim assembly As Assembly = Assembly.GetExecutingAssembly()
        Dim stream As Stream = assembly.GetManifestResourceStream("UnitTests.CloudIcon.PNG")
        Dim image As Bitmap = New Bitmap(stream)

        Dim pic As IPictureDisp = New PicConverterAxHost().ToIPictureDisp(image)

        Assert.IsNotNull(pic)
    End Sub

End Class