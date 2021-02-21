Imports System.Xml.Schema
Imports RibbonFactory.Component_Interfaces

Namespace Containers

    Public NotInheritable Class Ribbon

        Private Const BoilerPlate As String =
            "<?xml version=""1.0"" encoding=""utf-8"" ?>

            <customUI xmlns=""http//schemas.microsoft.com/office/2009/07/customui"">
                <ribbon {0}>
                    <tabs>
                        {1}
                    </tabs>                            
                </ribbon>
            </customUI>"

        Private ReadOnly _startFromScratch As Boolean
        Private ReadOnly _nameSpace As String
        Private ReadOnly _allElements As List(Of RibbonElement) = New List(Of RibbonElement)

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
            '_Tabs.ForEach(Sub(T) AddElements(T))
            Return String.Format(BoilerPlate, $"startFromScratch=""{_startFromScratch}""", String.Join(vbNewLine, Tabs.Select(Function(T) T.XML)))
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

        Public Sub ValidateRibbon(ribbonX As String)
            Dim ribbonXValidationSettings As Xml.XmlReaderSettings = New Xml.XmlReaderSettings With {
                    .ValidationType = Xml.ValidationType.Schema
                    }

            ribbonXValidationSettings.Schemas.Add("http://schemas.microsoft.com/office/2009/07/customui", "RibbonX.xsd")

            AddHandler ribbonXValidationSettings.ValidationEventHandler, AddressOf LogXMLError

            Dim validator As Xml.XmlReader = Xml.XmlReader.Create(ribbonX, ribbonXValidationSettings)

            Do
                If Not validator.Read() Then
                    Exit Do
                End If
            Loop
        End Sub

        Private Shared Sub LogXMLError(sender As Object, e As ValidationEventArgs)
            Debug.WriteLine($"{e.Severity.ToString().ToUpper}: {e.Message}")
        End Sub
    End Class
End NameSpace