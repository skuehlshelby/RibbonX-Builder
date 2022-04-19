﻿Imports System.IO
Imports System.Reflection
Imports System.Xml
Imports System.Xml.Schema
Imports Microsoft.Office.Core
Imports RibbonFactory.BuilderInterfaces.API
Imports RibbonFactory.Builders
Imports RibbonFactory.RibbonAttributes
Imports RibbonFactory.Utilities.Testing

Namespace Controls

    Public NotInheritable Class Ribbon
        Implements IRibbonUI
        Implements IEnumerable(Of RibbonElement)
        Private _ribbon As IRibbonUI
        Private ReadOnly _elements As IDictionary(Of String, RibbonElement)
        Private ReadOnly _attributes As AttributeSet

#Region "Initialization"

        Public Sub New(ParamArray tabs() As Tab)
            Me.New(Nothing, tabs)
        End Sub

        Public Sub New(configuration As Action(Of IRibbonBuilder), ParamArray tabs() As Tab)
            _elements = Flatten(tabs).ToDictionary(Function(element) element.ID)

            Dim builder As RibbonBuilder = New RibbonBuilder()

            configuration = If(configuration, Sub(b) Return)
            configuration.Invoke(builder)

            _attributes = builder.Build()

            RibbonX = String.Join(Environment.NewLine,
                                  "<?xml version=""1.0"" encoding=""utf-8"" ?>",
                                  $"<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" {_attributes.ToString(AttributeCategory.OnLoad)} {_attributes.ToString(AttributeCategory.LoadImage)}>",
                                  $"<ribbon {_attributes.ToString(AttributeCategory.StartFromScratch)}>",
                                  "<tabs>",
                                  String.Join(Environment.NewLine, tabs.Select(Function(tab) tab.XML)),
                                  "</tabs>",
                                  "</ribbon>",
                                  "</customUI>")

            RegisterEvents(_elements.Values)
        End Sub

        Private Function Flatten(tabs As ICollection(Of Tab)) As ICollection(Of RibbonElement)
            Dim results As ICollection(Of RibbonElement) = New List(Of RibbonElement)

            For Each tab As Tab In tabs
                tab.Flatten(results)
            Next

            Return results
        End Function

        Private Sub RegisterEvents(elements As IEnumerable(Of RibbonElement))
            For Each element As RibbonElement In elements
                AddHandler element.ValueChanged, AddressOf HandleValueChanged
            Next
        End Sub

        Private Sub HandleValueChanged(sender As Object, e As ValueChangedEventArgs)
            _ribbon.InvalidateControl(e.ID)
        End Sub

#End Region

#Region "API"

        Public ReadOnly Property RibbonX As String

        Public Sub AssignRibbonUI(ribbonUi As IRibbonUI)
            _ribbon = If(_ribbon, ribbonUi)
        End Sub

        Public Function GetElement(id As String) As RibbonElement
            Return _elements.Item(id)
        End Function

        Public Function GetElement(Of TInterface)(id As String) As TInterface
            Return CType(CType(GetElement(id), Object), TInterface)
        End Function

        Public Function GetContainer(Of TRibbonElement As RibbonElement)(id As String) As Container(Of TRibbonElement)
            Return CType(GetElement(id), Container(Of TRibbonElement))
        End Function

        Public Function GetContainerItem(parentId As String, index As Integer) As Item
            Return GetContainer(Of Item)(parentId)(index)
        End Function

        Public Function GetErrors() As RibbonErrorLog
            Return GetErrors(RibbonX)
        End Function

        Public Shared Function GetErrors(ribbonX As String) As RibbonErrorLog
            Return New RibbonErrorLog(ribbonX, GetCustomUIVersion2009())
        End Function

        Private Shared Function GetCustomUIVersion2009() As XmlSchema
            Using reader As StreamReader = New StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream("RibbonFactory.RibbonX.xsd"))
                Return XmlSchema.Read(XmlReader.Create(reader), AddressOf ValidationEventHandler)
            End Using
        End Function

        Private Shared Sub ValidationEventHandler(sender As Object, e As ValidationEventArgs)
            Debug.WriteLine($"{e.Severity.ToString().ToUpper}: {e.Message}")
        End Sub

#Region "IRibbonUI Members"

        Public Sub RefreshAll() Implements IRibbonUI.Invalidate
            _ribbon.Invalidate()
        End Sub

        Public Sub RefreshControl(controlID As String) Implements IRibbonUI.InvalidateControl
            _ribbon.InvalidateControl(controlID)
        End Sub

        Public Sub RefreshBuiltInControl(msoID As String) Implements IRibbonUI.InvalidateControlMso
            _ribbon.InvalidateControlMso(msoID)
        End Sub

        Public Sub ActivateTab(controlID As String) Implements IRibbonUI.ActivateTab
            _ribbon.ActivateTab(controlID)
        End Sub

        Public Sub ActivateTabMso(controlID As String) Implements IRibbonUI.ActivateTabMso
            _ribbon.ActivateTabMso(controlID)
        End Sub

        Public Sub ActivateTabQ(controlID As String, [namespace] As String) Implements IRibbonUI.ActivateTabQ
            _ribbon.ActivateTabQ(controlID, [namespace])
        End Sub

#End Region

#End Region

        Public Function GetEnumerator() As IEnumerator(Of RibbonElement) Implements IEnumerable(Of RibbonElement).GetEnumerator
            Return _elements.Values.GetEnumerator()
        End Function

        Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return GetEnumerator()
        End Function

    End Class

End Namespace