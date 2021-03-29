Imports System.IO
Imports System.Reflection
Imports System.Xml
Imports System.Xml.Schema
Imports NetOffice.OfficeApi

Namespace Containers

    Public NotInheritable Class Ribbon

        Private Const BoilerPlate As String =
            "<?xml version=""1.0"" encoding=""utf-8"" ?>" _
            & vbNewLine & vbNewLine & _
            "<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">" _
            & vbNewLine & vbTab & "<ribbon {0}>" _
            & vbNewLine & vbTab & vbTab & "<tabs>" _
            & vbNewLine & vbTab & vbTab & vbTab & "{1}" _
            & vbNewLine & vbTab & vbTab & "</tabs>" _                            
            & vbNewLine & vbTab & "</ribbon>" _
            & vbNewLine & "</customUI>"

        Private ReadOnly _startFromScratch As Boolean
        Private ReadOnly _nameSpace As String
        Private ReadOnly _allElements As List(Of RibbonElement) = New List(Of RibbonElement)
        Private _ribbon As IRibbonUI

        Public Sub New(Optional [nameSpace] As String = Nothing, Optional startFromScratch As Boolean = False)
            _nameSpace = [nameSpace]
            _startFromScratch = startFromScratch
        End Sub

        Public Function GetItem(Of T)(itemID As String) As T
            Throw New NotImplementedException
        End Function

        Private Function SearchByID(itemID As String) As RibbonElement
            Throw New NotImplementedException
        End Function

        Public ReadOnly Property Tabs As ICollection(Of Tab)

        Public Function Build() As String
            Dim ribbonX As String = String.Format(BoilerPlate, $"startFromScratch=""{_startFromScratch.ToString().ToLower()}""", String.Join(vbNewLine, Tabs.Select(Function(T) T.XML)))
            ValidateRibbon(ribbonX)
            Return ribbonX
        End Function

        Public Sub AssignRibbonUI(ribbonUi As IRibbonUI)
            _ribbon = If(_ribbon, ribbonUi)
        End Sub

        Public Shared Sub ValidateRibbon(ribbonX As String)
            Dim assembly As Assembly = Assembly.GetExecutingAssembly()

            Dim xsd As XmlSchemaSet = New XmlSchemaSet()

            Using reader As StreamReader = New StreamReader(assembly.GetManifestResourceStream("RibbonFactory.RibbonX.xsd"))
                xsd.Add("http://schemas.microsoft.com/office/2009/07/customui", XmlReader.Create(reader))

                Dim ribbon As XDocument = XDocument.Parse(ribbonX)

                ribbon.Validate(xsd, AddressOf LogXMLError)
            End Using
        End Sub

        Private Shared Sub LogXMLError(sender As Object, e As ValidationEventArgs)
            Debug.WriteLine($"{e.Severity.ToString().ToUpper}: {e.Message}")
        End Sub
    End Class
End NameSpace