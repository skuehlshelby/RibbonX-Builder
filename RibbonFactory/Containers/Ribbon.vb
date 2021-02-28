Imports System.IO
Imports System.Reflection
Imports System.Xml
Imports System.Xml.Schema
Imports NetOffice.OfficeApi
Imports RibbonFactory.Component_Interfaces

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
            Tabs = New List(Of Tab)()
        End Sub

        Public ReadOnly Property Tabs As List(Of Tab)

        Public ReadOnly Property AllElements As List(Of RibbonElement)
            Get
                Return _allElements
            End Get
        End Property

        Public Function AddTab(displayName As String) As Tab
            Dim newTab As Tab = New Tab(displayName)
            Tabs.Add(newTab)
            Return newTab
        End Function

        Public Function Build() As String
            Dim ribbonX As String = String.Format(BoilerPlate, $"startFromScratch=""{_startFromScratch.ToString().ToLower()}""", String.Join(vbNewLine, Tabs.Select(Function(T) T.XML)))
            ValidateRibbon(ribbonX)
            Return ribbonX
        End Function

        Private Sub AddElements(container As IContainer)
            For Each e As RibbonElement In container.Items
                If Not _allElements.Contains(e) Then
                    _allElements.Add(e)
                End If
                If TypeOf e Is IContainer Then
                    'AddElements(E)
                End If
            Next e
        End Sub

        Private Function Flatten(container As IContainer) As List(Of RibbonElement)
            Flatten = New List(Of RibbonElement)

            If container.Items.Any(Function(item) TypeOf item Is IContainer) Then
                For Each c As IContainer In container.Items.Where(Function(item) TypeOf item Is IContainer).Select(Function(item) (DirectCast(item, IContainer)))
                    Flatten.AddRange(Flatten(c)) 
                Next c
            Else
                Flatten.AddRange(container.Items)
            End If
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