Imports System.IO
Imports System.Reflection
Imports NUnit.Framework
Imports RibbonFactory
Imports System.Drawing

Public Class IPictureDispConversion
    <Test>
    Public Sub CanConvertIPicture()
        Dim assmbly As Assembly = Assembly.GetExecutingAssembly()
        Dim stream As Stream = assmbly.GetManifestResourceStream("Cloud Icon.PNG")

        Dim CloudIcon As Bitmap = New Bitmap(stream)
    End Sub
End Class
