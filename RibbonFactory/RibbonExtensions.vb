Imports System.IO
Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Xml
Imports System.Xml.Schema
Imports RibbonFactory.Containers

Public Module RibbonExtensions
    <Extension>
    Public Function GetErrors(ribbon As Ribbon) As RibbonErrorLog
        Return GetErrors(ribbon.Build())
    End Function

    Public Function GetErrors(ribbonX As String) As RibbonErrorLog
        Return New RibbonErrorLog(ribbonX, GetCustomUIVersion2009())
    End Function

    Public Function GetCustomUIVersion2009() As XmlSchema
        Using reader As StreamReader = New StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream("RibbonFactory.RibbonX.xsd"))
            Return XmlSchema.Read(XmlReader.Create(reader), AddressOf ValidationEventHandler)
        End Using
    End Function

    Private Sub ValidationEventHandler(sender As Object, e As ValidationEventArgs)
        Console.WriteLine($"{e.Severity.ToString().ToUpper}: {e.Message}")
    End Sub
End Module