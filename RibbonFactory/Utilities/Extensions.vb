Imports System.IO
Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Xml
Imports System.Xml.Schema

Public Module Extensions

    <Extension()>
    Public Sub Add(Of T)(collection As ICollection(Of T), ParamArray items As T())
        For Each item As T In items
            collection.Add(item)
        Next
    End Sub

    <Extension()>
    Public Function BoxHorizontal(_1 As RxApi, ParamArray controls() As IBoxAddable) As IBox
        Return RxApi.Box(Sub(b) b.Visible().Horizontal().WithControls(controls))
    End Function

    <Extension()>
    Public Function GetCustomUIVersion2009(_1 As RxApi) As XmlSchema
        Dim schema As XmlSchema

        Using stream As Stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("RibbonX.RibbonX.xsd")
            Using reader As XmlTextReader = New XmlTextReader(stream)
                schema = XmlSchema.Read(reader, AddressOf ValidationEventHandler)
            End Using
        End Using

        Return schema
    End Function

    Private Sub ValidationEventHandler(sender As Object, e As ValidationEventArgs)
        Debug.WriteLine($"{e.Severity.ToString().ToUpper}: {e.Message}")
    End Sub
End Module
