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

        Dim pic As IPictureDisp = Utilities.ConvertToIPictureDisplay(image)

        Assert.IsTrue(pic.Height > 0)
        Assert.IsTrue(pic.Width > 0)
    End Sub

End Class