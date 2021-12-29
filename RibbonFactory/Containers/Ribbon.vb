Imports Microsoft.Office.Core
Imports RibbonFactory.Controls

Namespace Containers

    Public NotInheritable Class Ribbon
        Implements IRibbonUI
        Private _ribbon As IRibbonUI
        Private ReadOnly _elements As IDictionary(Of String, RibbonElement)

#Region "Initialization"

        Friend Sub New(elements As IEnumerable(Of RibbonElement), ribbonX As String)
            _elements = elements.Distinct().ToDictionary(Function(element) element.ID)
            RegisterEvents(_elements.Values)
            Me.RibbonX = ribbonX
        End Sub

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

        Public Property RibbonX as String

        Public Sub AssignRibbonUI(ribbonUi As IRibbonUI)
            _ribbon = If(_ribbon, ribbonUi)
        End Sub

        Public Function GetElement(id As String) As RibbonElement
            Return _elements.Item(id)
        End Function

        Public Function GetElement(Of TRibbonElement As RibbonElement)(predicate As Func(Of TRibbonElement, Boolean)) As TRibbonElement
            Return _elements.Values.OfType(Of TRibbonElement)().Single(predicate)
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

    End Class

End NameSpace