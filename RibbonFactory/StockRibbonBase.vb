Imports System.Runtime.InteropServices
Imports RibbonX.Callbacks
Imports RibbonX.Controls
Imports RibbonX.Controls.Properties
Imports RibbonX.ComTypes.StdOle
Imports RibbonX.ComTypes.Extensibility
Imports RibbonX.ComTypes.Microsoft.Office.Core
Imports RibbonX.SimpleTypes



<ComVisible(True)>
    Public MustInherit Class StockRibbonBase
        Implements IDTExtensibility2
        Implements IRibbonExtensibility
        Implements ICreateCallbacks

        Protected ReadOnly Extensions As ICollection(Of IDTExtensibility2)
        Protected ReadOnly Ribbon As Ribbon
        Protected ReadOnly ImageCache As IDictionary(Of String, IPictureDisp) = New Dictionary(Of String, IPictureDisp)
        Protected HostApp As Object

        Protected Sub New(ParamArray extensions() As IDTExtensibility2)
            Ribbon = BuildRibbon()
            Me.Extensions = extensions
        End Sub

        Protected MustOverride Function BuildRibbon() As Ribbon

        Public Function GetCustomUI(ribbonID As String) As String Implements IRibbonExtensibility.GetCustomUI
            Return Ribbon.RibbonX
        End Function

#Region "IDTExtensibility2 Members"

        Public Sub OnConnection(application As Object, connectMode As ext_ConnectMode, addInInst As Object, ByRef custom As Array) Implements IDTExtensibility2.OnConnection
            HostApp = application

            For Each extension As IDTExtensibility2 In Extensions
                extension.OnConnection(application, connectMode, addInInst, custom)
            Next
        End Sub

        Public Sub OnDisconnection(removeMode As ext_DisconnectMode, ByRef custom As Array) Implements IDTExtensibility2.OnDisconnection
            For Each extension As IDTExtensibility2 In Extensions
                extension.OnDisconnection(removeMode, custom)
            Next

            HostApp = Nothing

            For cycle As Integer = 1 To 2
                GC.Collect()
                GC.WaitForPendingFinalizers()
            Next
        End Sub

        Public Sub OnAddInsUpdate(ByRef custom As Array) Implements IDTExtensibility2.OnAddInsUpdate
            For Each extension As IDTExtensibility2 In Extensions
                extension.OnAddInsUpdate(custom)
            Next
        End Sub

        Public Sub OnStartupComplete(ByRef custom As Array) Implements IDTExtensibility2.OnStartupComplete
            For Each extension As IDTExtensibility2 In Extensions
                extension.OnStartupComplete(custom)
            Next
        End Sub

        Public Sub OnBeginShutdown(ByRef custom As Array) Implements IDTExtensibility2.OnBegInShutdown
            For Each extension As IDTExtensibility2 In Extensions
                extension.OnBegInShutdown(custom)
            Next
        End Sub

#End Region

#Region "Callbacks"

        Public Sub OnLoad(ribbonUI As IRibbonUI) Implements ICreateCallbacks.OnLoad
            Ribbon.AssignRibbonUI(ribbonUI)
        End Sub

        Public Function LoadImage(imageID As String) As IPictureDisp Implements ICreateCallbacks.LoadImage
            Return ImageCache(imageID)
        End Function

        Public Function GetEnabled(control As IRibbonControl) As Boolean Implements ICreateCallbacks.GetEnabled
            Return Ribbon.GetElement(Of IEnabled)(control.Id).Enabled
        End Function

        Public Function GetVisible(control As IRibbonControl) As Boolean Implements ICreateCallbacks.GetVisible
            Return Ribbon.GetElement(Of IVisible)(control.Id).Visible
        End Function

        Public Function GetLabel(control As IRibbonControl) As String Implements ICreateCallbacks.GetLabel
            Return Ribbon.GetElement(Of ILabel)(control.Id).Label
        End Function

        Public Function GetShowLabel(control As IRibbonControl) As Boolean Implements ICreateCallbacks.GetShowLabel
            Return Ribbon.GetElement(Of IShowLabel)(control.Id).ShowLabel
        End Function

        Public Function GetDescription(control As IRibbonControl) As String Implements ICreateCallbacks.GetDescription
            Return Ribbon.GetElement(Of IDescription)(control.Id).Description
        End Function

        Public Function GetScreenTip(control As IRibbonControl) As String Implements ICreateCallbacks.GetScreenTip
            Return Ribbon.GetElement(Of IScreenTip)(control.Id).ScreenTip
        End Function

        Public Function GetSuperTip(control As IRibbonControl) As String Implements ICreateCallbacks.GetSuperTip
            Return Ribbon.GetElement(Of ISuperTip)(control.Id).SuperTip
        End Function

        Public Function GetKeyTip(control As IRibbonControl) As KeyTip Implements ICreateCallbacks.GetKeyTip
            Return Ribbon.GetElement(Of IKeyTip)(control.Id).KeyTip
        End Function

        Public Function GetSize(control As IRibbonControl) As ControlSize Implements ICreateCallbacks.GetSize
            Return Ribbon.GetElement(Of ISize)(control.Id).Size
        End Function

        Public Function GetImage(control As IRibbonControl) As IPictureDisp Implements ICreateCallbacks.GetImage
            Return Ribbon.GetElement(Of IImage)(control.Id).Image.AsIPictureDisp()
        End Function

        Public Function GetShowImage(control As IRibbonControl) As Boolean Implements ICreateCallbacks.GetShowImage
            Return Ribbon.GetElement(Of IShowImage)(control.Id).ShowImage
        End Function

        Public Function GetPressed(control As IRibbonControl) As Boolean Implements ICreateCallbacks.GetPressed
            Return Ribbon.GetElement(Of IPressed)(control.Id).Checked
        End Function

        Public Function GetText(control As IRibbonControl) As String Implements ICreateCallbacks.GetText
            Return Ribbon.GetElement(Of IText)(control.Id).Text
        End Function

        Public Function GetItemCount(control As IRibbonControl) As Integer Implements ICreateCallbacks.GetItemCount
            Return Ribbon.GetContainer(Of Item)(control.Id).Count
        End Function

        Public Function GetItemID(control As IRibbonControl, index As Integer) As String Implements ICreateCallbacks.GetItemID
            Return Ribbon.GetContainerItem(control.Id, index).ID
        End Function

        Public Function GetItemImage(control As IRibbonControl, index As Integer) As IPictureDisp Implements ICreateCallbacks.GetItemImage
            Return Ribbon.GetContainerItem(control.Id, index).Image.AsIPictureDisp()
        End Function

        Public Function GetItemLabel(control As IRibbonControl, index As Integer) As String Implements ICreateCallbacks.GetItemLabel
            Return Ribbon.GetContainerItem(control.Id, index).Label
        End Function

        Public Function GetItemScreenTip(control As IRibbonControl, index As Integer) As String Implements ICreateCallbacks.GetItemScreenTip
            Return Ribbon.GetContainerItem(control.Id, index).ScreenTip
        End Function

        Public Function GetItemSuperTip(control As IRibbonControl, index As Integer) As String Implements ICreateCallbacks.GetItemSuperTip
            Return Ribbon.GetContainerItem(control.Id, index).SuperTip
        End Function

        Public Function GetItemHeight(control As IRibbonControl) As Integer Implements ICreateCallbacks.GetItemHeight
            Return Ribbon.GetElement(Of IItemDimensions)(control.Id).ItemHeight
        End Function

        Public Function GetItemWidth(control As IRibbonControl) As Integer Implements ICreateCallbacks.GetItemWidth
            Return Ribbon.GetElement(Of IItemDimensions)(control.Id).ItemWidth
        End Function

        Public Function GetSelectedItemID(control As IRibbonControl) As String Implements ICreateCallbacks.GetSelectedItemID
            Return Ribbon.GetElement(Of ISelect)(control.Id).Selected.ID
        End Function

        Public Function GetSelectedItemIndex(control As IRibbonControl) As Integer Implements ICreateCallbacks.GetSelectedItemIndex
            Return Ribbon.GetElement(Of ISelect)(control.Id).SelectedItemIndex
        End Function

        Public Sub OnAction(control As IRibbonControl) Implements ICreateCallbacks.OnAction
            Ribbon.GetElement(Of IClickable)(control.Id).Click()
        End Sub

        Public Sub OnChange(control As IRibbonControl, text As String) Implements ICreateCallbacks.OnChange
            Ribbon.GetElement(Of IText)(control.Id).Text = text
        End Sub

        Public Sub OnSelectionChange(control As IRibbonControl, selectedId As String, selectedIndex As Integer) Implements ICreateCallbacks.OnSelectionChange
            Ribbon.GetElement(Of ISelect)(control.Id).SelectedItemIndex = selectedIndex
        End Sub

        Public Sub OnButtonToggle(control As IRibbonControl, pressed As Boolean) Implements ICreateCallbacks.OnButtonToggle
            Ribbon.GetElement(Of IPressed)(control.Id).Checked = pressed
        End Sub

#End Region

End Class