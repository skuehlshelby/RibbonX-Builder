Imports NUnit.Framework
Public Class ControlPropertyTests
    Private Const Template As String = "{0}=""{1}"""
    <Test>
    Public Sub StaticXMLElement()
        System.Console.WriteLine(String.Format(Template, "label", "So Labeled"))
    End Sub

    Private Sub TestAction(ByVal control As NetOffice.OfficeApi.IRibbonControl)
    End Sub

    <Test>
    Public Sub AreSame()
        Debug.WriteLine(String.Format("{0} {1} {2}", RibbonFactory.Attributes.Image.ImageMso, If(RibbonFactory.Attributes.Image.ImageMso.Equals(RibbonFactory.Attributes.Image.GetImage), "is", "is not"), RibbonFactory.Attributes.Image.GetImage))
    End Sub
End Class
