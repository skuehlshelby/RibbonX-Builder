Imports System.Xml.Schema

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

    Private Structure TThis
        Public StartFromScratch As Boolean
        Public [NameSpace] As String
        Public AllElements As List(Of RibbonElement)
        Public Tabs As List(Of Tab)
    End Structure

    Private ReadOnly This As TThis

    Public Sub New(Optional ByVal [NameSpace] As String = Nothing, Optional ByVal StartFromScratch As Boolean = False)
        This = New TThis With {
            .AllElements = New List(Of RibbonElement),
            .[NameSpace] = [NameSpace],
            .StartFromScratch = StartFromScratch,
            .Tabs = New List(Of Tab)
        }
    End Sub

    Public ReadOnly Property Tabs As List(Of Tab)
        Get
            Return This.Tabs
        End Get
    End Property

    Public ReadOnly Property AllElements As List(Of RibbonElement)
        Get
            Return This.AllElements
        End Get
    End Property

    Public Function AddTab(ByVal DisplayName As String) As Tab
        Dim NewTab As Tab = New Tab(DisplayName)
        Tabs.Add(NewTab)
        Return NewTab
    End Function

    Public Function Build() As String
        This.Tabs.ForEach(Sub(T) AddElements(T))
        Return String.Format(BoilerPlate, If(This.StartFromScratch, "startFromScratch=""true""", String.Empty), String.Join(vbNewLine, Tabs.Select(Function(T) T.XML)))
    End Function

    Private Sub AddElements(ByVal Container As IContainer)
        For Each E As RibbonElement In Container.Items
            If Not This.AllElements.Contains(E) Then
                This.AllElements.Add(E)
            End If
            If TypeOf E Is IContainer Then
                AddElements(E)
            End If
        Next E
    End Sub

    Public Sub ValidateRibbon(ByVal RibbonX As String)
        Dim RibbonXValidationSettings As Xml.XmlReaderSettings = New Xml.XmlReaderSettings With {
            .ValidationType = Xml.ValidationType.Schema
        }

        RibbonXValidationSettings.Schemas.Add("http://schemas.microsoft.com/office/2009/07/customui", "RibbonX.xsd")

        AddHandler RibbonXValidationSettings.ValidationEventHandler, AddressOf LogXMLError

        Dim Validator As Xml.XmlReader = Xml.XmlReader.Create(RibbonX, RibbonXValidationSettings)

        Do
            If Not Validator.Read() Then
                Exit Do
            End If
        Loop
    End Sub

    Private Sub LogXMLError(ByVal Sender As Object, ByVal E As ValidationEventArgs)
        Debug.WriteLine(String.Format("{0}: {1}", E.Severity.ToString().ToUpper, E.Message))
    End Sub
End Class
