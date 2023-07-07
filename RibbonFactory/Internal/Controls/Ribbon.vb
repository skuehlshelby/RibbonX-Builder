Imports System.IO
Imports System.Reflection
Imports System.Windows.Forms
Imports System.Xml
Imports System.Xml.Schema
Imports RibbonX.ComTypes.Microsoft.Office.Core
Imports RibbonX.InternalApi
Imports RibbonX.Properties
Imports RibbonX.Testing
Imports RibbonX.Utilities

Namespace Controls

    <DebuggerDisplay("{GetDebuggerDisplay(),nq}")>
    Friend NotInheritable Class Ribbon
        Implements IRibbon
        Implements IRibbonUI
        Implements IEnumerable(Of IRibbonElement)
        Private _ribbon As IRibbonUI
        Private ReadOnly _elements As IDictionary(Of String, IRibbonElement)
        Private ReadOnly _attributes As IPropertyCollection

#Region "Initialization"

        Public Sub New(attributes As IPropertyCollection, ParamArray tabs() As ITab)
            _attributes = attributes

            _elements = ConvertElementsToDictionaryAndThrowIfDuplicatesArePresent(tabs)

            Dim customUiAttributes As ICollection(Of String) = _attributes.
                Raw(Category.OnLoad, Category.LoadImage).
                Select(Function(p) p.ToXml()).
                ToArray()

            RibbonX = String.Join(Environment.NewLine,
                                        "<?xml version=""1.0"" encoding=""utf-8"" ?>",
                                        $"<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" {String.Join(" ", customUiAttributes)}>",
                                        $"<ribbon {_attributes.Raw(Category.StartFromScratch).ToXml()}>",
                                        "<tabs>",
                                        String.Join(Environment.NewLine, tabs.Select(Function(tab) tab.ToXml())),
                                        "</tabs>",
                                        "</ribbon>",
                                        "</customUI>").NoDoubleSpace()

            RegisterEvents(_elements.Values)
        End Sub

        Private Sub RegisterEvents(elements As IEnumerable(Of IRibbonElement))
            For Each element As IRibbonElement In elements
                AddHandler element.RefreshNeeded, AddressOf HandleValueChanged
            Next
        End Sub

        Private Sub HandleValueChanged(sender As Object, e As RefreshNeededEventArgs)
            If _ribbon IsNot Nothing Then
                _ribbon.InvalidateControl(e.ID)
            Else
                Throw New MissingOnLoadException()
            End If
        End Sub

        Private Function ConvertElementsToDictionaryAndThrowIfDuplicatesArePresent(tabs() As ITab) As Dictionary(Of String, IRibbonElement)
            Dim results As ICollection(Of IRibbonElement) = Flatten(tabs)

            Dim duplicateElements As ICollection(Of String) = results.
                GroupBy(Function(e) e.Id).
                Where(Function(i) i.Count() > 1).
                Select(Function(g) g.Key).
                ToArray()

            If duplicateElements.Any() Then
                Throw New InvalidOperationException($"More than one instance of {If(1 < duplicateElements.Count, "elements", "element")} {String.Join(", ", duplicateElements)} {If(1 < duplicateElements.Count, "were", "was")} detected. Use .Clone() to add elements more than once.")
            End If

            Return results.ToDictionary(Function(e) e.Id)
        End Function

        Private Function Flatten(tabs() As ITab) As ICollection(Of IRibbonElement)
            Dim results As ICollection(Of IRibbonElement) = New List(Of IRibbonElement)()

            Dim unprocessed As Stack(Of IRibbonElement) = New Stack(Of IRibbonElement)(tabs)

            While unprocessed.Any()
                Dim element As IRibbonElement = unprocessed.Pop()

                If TypeOf element Is IEnumerable(Of IRibbonElement) Then
                    For Each child As IRibbonElement In DirectCast(element, IEnumerable(Of IRibbonElement))
                        unprocessed.Push(child)
                    Next
                End If

                results.Add(element)
            End While

            Return results
        End Function
#End Region

#Region "API"

        Public ReadOnly Property RibbonX As String Implements IRibbon.RibbonX

        Public ReadOnly Property Count As Integer Implements IReadOnlyCollection(Of IRibbonElement).Count
            Get
                Return _elements.Count
            End Get
        End Property

        Public Sub AssignRibbonUI(ribbonUi As IRibbonUI) Implements IRibbon.ReceiveIRibbonUI
            _ribbon = If(_ribbon, ribbonUi)
        End Sub

        Public Function GetElement(id As String) As IRibbonElement Implements IRibbon.GetElement
            Return _elements.Item(id)
        End Function

        Public Function GetElement(Of TInterface As IRibbonElementProperty)(id As String) As TInterface Implements IRibbon.GetElementProperty
            Return CType(CType(GetElement(id), Object), TInterface)
        End Function

        Public Function GetContainer(Of TRibbonElement As IRibbonElement)(id As String) As IReadOnlyCollection(Of TRibbonElement) Implements IRibbon.GetContainer
            Return CType(GetElement(id), IReadOnlyCollection(Of TRibbonElement))
        End Function

        Public Function GetContainerItem(parentId As String, index As Integer) As IItem Implements IRibbon.GetContainerItem
            Return GetContainer(Of IItem)(parentId)(index)
        End Function

        Public Function GetErrors() As RibbonErrorLog Implements IRibbon.GetErrors
            Return GetErrors(RibbonX)
        End Function

        Public Shared Function GetErrors(ribbonX As String) As RibbonErrorLog
            Return New RibbonErrorLog(ribbonX, GetCustomUIVersion2009())
        End Function

        Private Shared Function GetCustomUIVersion2009() As XmlSchema
            Dim schema As XmlSchema

            Using stream As Stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("RibbonX.RibbonX.xsd")
                Using reader As XmlTextReader = New XmlTextReader(stream)
                    schema = XmlSchema.Read(reader, AddressOf ValidationEventHandler)
                End Using
            End Using

            Return schema
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

        Public Function GetEnumerator() As IEnumerator(Of IRibbonElement) Implements IEnumerable(Of IRibbonElement).GetEnumerator
            Return _elements.Values.GetEnumerator()
        End Function

        Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return GetEnumerator()
        End Function

        Private Function GetDebuggerDisplay() As String
            Return $"Ribbon {_attributes}"
        End Function

    End Class

End Namespace