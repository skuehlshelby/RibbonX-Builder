Imports System.Drawing
Imports RibbonFactory.BuilderInterfaces.API
Imports RibbonFactory.Controls
Imports RibbonFactory.Controls.Events
Imports RibbonFactory.Enums.ImageMSO
Imports RibbonFactory.Enums.MSO
Imports stdole

Namespace Builders

    Friend NotInheritable Class DropDownBuilder
        Inherits BuilderBase(Of DropDown)
        Implements IDropdownBuilder

        Public Sub New()
            MyBase.New()
        End Sub

        Public Sub New(template As RibbonElement)
            MyBase.New(template)
        End Sub

        Private Function AsWideAs(sizeString As String) As IDropdownBuilder Implements IDropdownBuilder.AsWideAs
            SizeStringBase(sizeString)
            Return Me
        End Function

        Private Function Disabled() As IDropdownBuilder Implements IDropdownBuilder.Disabled
            DisabledBase()
            Return Me
        End Function

        Private Function Disabled(callback As FromControl(Of Boolean)) As IDropdownBuilder Implements IDropdownBuilder.Disabled
            DisabledBase(callback)
            Return Me
        End Function

        Private Function Enabled() As IDropdownBuilder Implements IDropdownBuilder.Enabled
            EnabledBase()
            Return Me
        End Function

        Private Function Enabled(callback As FromControl(Of Boolean)) As IDropdownBuilder Implements IDropdownBuilder.Enabled
            EnabledBase(callback)
            Return Me
        End Function

        Private Function GetItemCountFrom(callback As FromControl(Of Integer)) As IDropdownBuilder Implements IDropdownBuilder.GetItemCountFrom
            ItemCountBase(callback)
            Return Me
        End Function

        Private Function GetItemIdFrom(callback As FromControlAndIndex(Of String)) As IDropdownBuilder Implements IDropdownBuilder.GetItemIdFrom
            ItemIDBase(callback)
            Return Me
        End Function

        Private Function GetItemImageFrom(callback As FromControlAndIndex(Of IPictureDisp)) As IDropdownBuilder Implements IDropdownBuilder.GetItemImageFrom
            ItemImageBase(callback)
            Return Me
        End Function

        Private Function GetItemLabelFrom(callback As FromControlAndIndex(Of String)) As IDropdownBuilder Implements IDropdownBuilder.GetItemLabelFrom
            ItemLabelBase(callback)
            Return Me
        End Function

        Private Function GetItemScreenTipFrom(callback As FromControlAndIndex(Of String)) As IDropdownBuilder Implements IDropdownBuilder.GetItemScreenTipFrom
            ItemScreentipBase(callback)
            Return Me
        End Function

        Private Function GetItemSuperTipFrom(callback As FromControlAndIndex(Of String)) As IDropdownBuilder Implements IDropdownBuilder.GetItemSuperTipFrom
            ItemSupertipBase(callback)
            Return Me
        End Function

        Private Function GetSelectedItemIdFrom(getSelected As FromControl(Of String), setSelected As SelectionChanged) As IDropdownBuilder Implements IDropdownBuilder.GetSelectedItemIdFrom
            SelectedBase(getSelected, setSelected)
            Return Me
        End Function

        Private Function GetSelectedItemIndexFrom(getSelected As FromControl(Of Integer), setSelected As SelectionChanged) As IDropdownBuilder Implements IDropdownBuilder.GetSelectedItemIndexFrom
            SelectedBase(getSelected, setSelected)
            Return Me
        End Function

        Private Function HideImage() As IDropdownBuilder Implements IDropdownBuilder.HideImage
            HideImageBase()
            Return Me
        End Function

        Private Function HideImage(getShowImage As FromControl(Of Boolean)) As IDropdownBuilder Implements IDropdownBuilder.HideImage
            HideImageBase(getShowImage)
            Return Me
        End Function

        Private Function HideItemImages() As IDropdownBuilder Implements IDropdownBuilder.HideItemImages
            HideItemImageBase()
            Return Me
        End Function

        Private Function HideLabel() As IDropdownBuilder Implements IDropdownBuilder.HideLabel
            HideLabelBase()
            Return Me
        End Function

        Private Function HideLabel(callback As FromControl(Of Boolean)) As IDropdownBuilder Implements IDropdownBuilder.HideLabel
            HideLabelBase(callback)
            Return Me
        End Function

        Private Function HideLabelsOnChildItems() As IDropdownBuilder Implements IDropdownBuilder.HideLabelsOnChildItems
            HideItemLabelBase()
            Return Me
        End Function

        Private Function InsertAfterMSO(builtInControl As MSO) As IDropdownBuilder Implements IDropdownBuilder.InsertAfter
            InsertAfterMsoBase(builtInControl)
            Return Me
        End Function

        Private Function InsertAfterQ(qualifiedControl As RibbonElement) As IDropdownBuilder Implements IDropdownBuilder.InsertAfter
            InsertAfterQBase(qualifiedControl)
            Return Me
        End Function

        Private Function InsertBeforeMSO(builtInControl As MSO) As IDropdownBuilder Implements IDropdownBuilder.InsertBefore
            InsertBeforeMsoBase(builtInControl)
            Return Me
        End Function

        Private Function InsertBeforeQ(qualifiedControl As RibbonElement) As IDropdownBuilder Implements IDropdownBuilder.InsertBefore
            InsertBeforeQBase(qualifiedControl)
            Return Me
        End Function

        Private Function Invisible() As IDropdownBuilder Implements IDropdownBuilder.Invisible
            InvisibleBase()
            Return Me
        End Function

        Private Function Invisible(callback As FromControl(Of Boolean)) As IDropdownBuilder Implements IDropdownBuilder.Invisible
            InvisibleBase(callback)
            Return Me
        End Function

        Private Function ShowImage() As IDropdownBuilder Implements IDropdownBuilder.ShowImage
            ShowImageBase()
            Return Me
        End Function

        Private Function ShowImage(getShowImage As FromControl(Of Boolean)) As IDropdownBuilder Implements IDropdownBuilder.ShowImage
            ShowImageBase(getShowImage)
            Return Me
        End Function

        Private Function ShowItemImages() As IDropdownBuilder Implements IDropdownBuilder.ShowItemImages
            ShowItemImageBase()
            Return Me
        End Function

        Private Function ShowLabel() As IDropdownBuilder Implements IDropdownBuilder.ShowLabel
            ShowLabelBase()
            Return Me
        End Function

        Private Function ShowLabel(callback As FromControl(Of Boolean)) As IDropdownBuilder Implements IDropdownBuilder.ShowLabel
            ShowLabelBase(callback)
            Return Me
        End Function

        Private Function ShowLabelsOnChildItems() As IDropdownBuilder Implements IDropdownBuilder.ShowLabelsOnChildItems
            ShowItemLabelBase()
            Return Me
        End Function

        Private Function Visible() As IDropdownBuilder Implements IDropdownBuilder.Visible
            VisibleBase()
            Return Me
        End Function

        Private Function Visible(callback As FromControl(Of Boolean)) As IDropdownBuilder Implements IDropdownBuilder.Visible
            VisibleBase(callback)
            Return Me
        End Function

        Private Function WithId(id As String) As IDropdownBuilder Implements IDropdownBuilder.WithId
            WithIdBase(id)
            Return Me
        End Function

        Private Function WithIdMso(mso As MSO) As IDropdownBuilder Implements IDropdownBuilder.WithIdMso
            WithIdBase(mso)
            Return Me
        End Function

        Private Function WithIdQ([namespace] As String, id As String) As IDropdownBuilder Implements IDropdownBuilder.WithIdQ
            WithIdBase([namespace], id)
            Return Me
        End Function

        Private Function WithImage(image As ImageMSO) As IDropdownBuilder Implements IDropdownBuilder.WithImage
            ImageBase(image)
            Return Me
        End Function

        Private Function WithImage(image As Bitmap, callback As FromControl(Of IPictureDisp)) As IDropdownBuilder Implements IDropdownBuilder.WithImage
            ImageBase(image, callback)
            Return Me
        End Function

        Private Function WithImage(image As Icon, callback As FromControl(Of IPictureDisp)) As IDropdownBuilder Implements IDropdownBuilder.WithImage
            ImageBase(image, callback)
            Return Me
        End Function

        Private Function WithImage(imagePath As String) As IDropdownBuilder Implements IDropdownBuilder.WithImage
            ImageBase(imagePath)
            Return Me
        End Function

        Private Function WithKeyTip(keyTip As KeyTip) As IDropdownBuilder Implements IDropdownBuilder.WithKeyTip
            KeyTipBase(keyTip)
            Return Me
        End Function

        Private Function WithKeyTip(keyTip As KeyTip, callback As FromControl(Of KeyTip)) As IDropdownBuilder Implements IDropdownBuilder.WithKeyTip
            KeyTipBase(keyTip, callback)
            Return Me
        End Function

        Private Function WithLabel(label As String, Optional copyToScreenTip As Boolean = True) As IDropdownBuilder Implements IDropdownBuilder.WithLabel
            LabelBase(label)
            ShowLabelBase()
            If copyToScreenTip Then ScreentipBase(label)
            Return Me
        End Function

        Private Function WithLabel(label As String, callback As FromControl(Of String), Optional copyToScreenTip As Boolean = True) As IDropdownBuilder Implements IDropdownBuilder.WithLabel
            LabelBase(label, callback)
            ShowLabelBase()
            If copyToScreenTip Then ScreentipBase(label, callback)
            Return Me
        End Function

        Private Function WithScreenTip(screenTip As String) As IDropdownBuilder Implements IDropdownBuilder.WithScreenTip
            ScreentipBase(screenTip)
            Return Me
        End Function

        Private Function WithScreenTip(screenTip As String, callback As FromControl(Of String)) As IDropdownBuilder Implements IDropdownBuilder.WithScreenTip
            ScreentipBase(screenTip, callback)
            Return Me
        End Function

        Private Function WithSuperTip(superTip As String) As IDropdownBuilder Implements IDropdownBuilder.WithSuperTip
            SupertipBase(superTip)
            Return Me
        End Function

        Private Function WithSuperTip(superTip As String, callback As FromControl(Of String)) As IDropdownBuilder Implements IDropdownBuilder.WithSuperTip
            SupertipBase(superTip, callback)
            Return Me
        End Function

        Private Function BeforeSelectionChange(ParamArray actions() As Action(Of BeforeSelectionChangeEventArgs)) As IDropdownBuilder Implements IDropdownBuilder.BeforeSelectionChange
            AddBeforeSelectionChangeHandlers(actions)
            Return Me
        End Function

        Private Function AfterSelectionChanged(ParamArray actions() As Action(Of SelectionChangeEventArgs)) As IDropdownBuilder Implements IDropdownBuilder.AfterSelectionChanged
            AddSelectionChangeHandlers(actions)
            Return Me
        End Function

        Private Function RouteSelectionChangeTo(callBack As SelectionChanged) As IDropdownBuilder Implements IDropdownBuilder.RouteSelectionChangeTo
            OnActionBase(callBack)
            Return Me
        End Function

    End Class

End Namespace