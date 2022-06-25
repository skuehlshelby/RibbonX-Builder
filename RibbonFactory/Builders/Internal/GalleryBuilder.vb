Imports System.Drawing
Imports RibbonX.Controls
Imports RibbonX.Controls.Base
Imports RibbonX.Controls.BuiltIn
Imports RibbonX.Controls.EventArgs
Imports RibbonX.ComTypes.StdOle
Imports RibbonX.SimpleTypes
Imports RibbonX.Callbacks
Imports RibbonX.Builders.Internal.Base
Imports RibbonX.Images.BuiltIn

Namespace Builders

    Friend NotInheritable Class GalleryBuilder
        Inherits BuilderBase(Of Gallery)
        Implements IGalleryBuilder

        Public Sub New()
            MyBase.New()
        End Sub

        Public Sub New(template As RibbonElement)
            MyBase.New(template)
        End Sub

        Private Function WithId(id As String) As IGalleryBuilder Implements IGalleryBuilder.WithId
            WithIdBase(id)
            Return Me
        End Function

        Private Function WithIdQ([namespace] As String, id As String) As IGalleryBuilder Implements IGalleryBuilder.WithIdQ
            WithIdBase([namespace], id)
            Return Me
        End Function

        Private Function WithIdMso(mso As MSO) As IGalleryBuilder Implements IGalleryBuilder.WithIdMso
            WithIdBase(mso)
            Return Me
        End Function

        Private Function Enabled() As IGalleryBuilder Implements IGalleryBuilder.Enabled
            EnabledBase()
            Return Me
        End Function

        Private Function Enabled(callback As FromControl(Of Boolean)) As IGalleryBuilder Implements IGalleryBuilder.Enabled
            EnabledBase(callback)
            Return Me
        End Function

        Private Function Disabled() As IGalleryBuilder Implements IGalleryBuilder.Disabled
            DisabledBase()
            Return Me
        End Function

        Private Function Disabled(callback As FromControl(Of Boolean)) As IGalleryBuilder Implements IGalleryBuilder.Disabled
            DisabledBase(callback)
            Return Me
        End Function

        Private Function Visible() As IGalleryBuilder Implements IGalleryBuilder.Visible
            VisibleBase()
            Return Me
        End Function

        Private Function Visible(callback As FromControl(Of Boolean)) As IGalleryBuilder Implements IGalleryBuilder.Visible
            VisibleBase(callback)
            Return Me
        End Function

        Private Function Invisible() As IGalleryBuilder Implements IGalleryBuilder.Invisible
            InvisibleBase()
            Return Me
        End Function

        Private Function Invisible(callback As FromControl(Of Boolean)) As IGalleryBuilder Implements IGalleryBuilder.Invisible
            InvisibleBase(callback)
            Return Me
        End Function

        Private Function InsertBeforeMSO(builtInControl As MSO) As IGalleryBuilder Implements IGalleryBuilder.InsertBefore
            InsertBeforeMsoBase(builtInControl)
            Return Me
        End Function

        Private Function InsertBeforeQ(qualifiedControl As RibbonElement) As IGalleryBuilder Implements IGalleryBuilder.InsertBefore
            InsertBeforeQBase(qualifiedControl)
            Return Me
        End Function

        Private Function InsertAfterMSO(builtInControl As MSO) As IGalleryBuilder Implements IGalleryBuilder.InsertAfter
            InsertAfterMsoBase(builtInControl)
            Return Me
        End Function

        Private Function InsertAfterQ(qualifiedControl As RibbonElement) As IGalleryBuilder Implements IGalleryBuilder.InsertAfter
            InsertAfterQBase(qualifiedControl)
            Return Me
        End Function

        Private Function WithLabel(label As String, Optional copyToScreenTip As Boolean = True) As IGalleryBuilder Implements IGalleryBuilder.WithLabel
            LabelBase(label)
            ShowLabelBase()
            If copyToScreenTip Then ScreentipBase(label)
            Return Me
        End Function

        Private Function WithLabel(label As String, callback As FromControl(Of String), Optional copyToScreenTip As Boolean = True) As IGalleryBuilder Implements IGalleryBuilder.WithLabel
            LabelBase(label, callback)
            ShowLabelBase()
            If copyToScreenTip Then ScreentipBase(label, callback)
            Return Me
        End Function

        Private Function WithScreenTip(screenTip As String) As IGalleryBuilder Implements IGalleryBuilder.WithScreenTip
            ScreentipBase(screenTip)
            Return Me
        End Function

        Private Function WithScreenTip(screenTip As String, callback As FromControl(Of String)) As IGalleryBuilder Implements IGalleryBuilder.WithScreenTip
            ScreentipBase(screenTip, callback)
            Return Me
        End Function

        Private Function WithSuperTip(superTip As String) As IGalleryBuilder Implements IGalleryBuilder.WithSuperTip
            SupertipBase(superTip)
            Return Me
        End Function

        Private Function WithSuperTip(superTip As String, callback As FromControl(Of String)) As IGalleryBuilder Implements IGalleryBuilder.WithSuperTip
            SupertipBase(superTip, callback)
            Return Me
        End Function

        Private Function ShowLabel() As IGalleryBuilder Implements IGalleryBuilder.ShowLabel
            ShowLabelBase()
            Return Me
        End Function

        Private Function ShowLabel(callback As FromControl(Of Boolean)) As IGalleryBuilder Implements IGalleryBuilder.ShowLabel
            ShowLabelBase(callback)
            Return Me
        End Function

        Private Function HideLabel() As IGalleryBuilder Implements IGalleryBuilder.HideLabel
            HideLabelBase()
            Return Me
        End Function

        Private Function HideLabel(callback As FromControl(Of Boolean)) As IGalleryBuilder Implements IGalleryBuilder.HideLabel
            HideLabelBase(callback)
            Return Me
        End Function

        Private Function InvalidateContentOnDrop() As IGalleryBuilder Implements IGalleryBuilder.InvalidateContentOnDrop
            InvalidateContentOnDropBase()
            Return Me
        End Function

        Private Function WithImage(image As ImageMSO) As IGalleryBuilder Implements IGalleryBuilder.WithImage
            ImageBase(image)
            Return Me
        End Function

        Private Function WithImage(image As ImageMSO, callback As FromControl(Of String)) As IGalleryBuilder Implements IGalleryBuilder.WithImage
            ImageBase(image, callback)
            Return Me
        End Function

        Private Function WithImage(image As Bitmap, callback As FromControl(Of IPictureDisp)) As IGalleryBuilder Implements IGalleryBuilder.WithImage
            ImageBase(image, callback)
            Return Me
        End Function

        Private Function WithImage(image As Icon, callback As FromControl(Of IPictureDisp)) As IGalleryBuilder Implements IGalleryBuilder.WithImage
            ImageBase(image, callback)
            Return Me
        End Function

        Private Function WithImage(imageId As String, image As Bitmap) As IGalleryBuilder Implements IGalleryBuilder.WithImage
            ImageBase(imageId, image)
            Return Me
        End Function

        Private Function WithImage(imageId As String, image As Icon) As IGalleryBuilder Implements IGalleryBuilder.WithImage
            ImageBase(imageId, image)
            Return Me
        End Function

        Private Function WithKeyTip(keyTip As KeyTip) As IGalleryBuilder Implements IGalleryBuilder.WithKeyTip
            KeyTipBase(keyTip)
            Return Me
        End Function

        Private Function WithKeyTip(keyTip As KeyTip, callback As FromControl(Of KeyTip)) As IGalleryBuilder Implements IGalleryBuilder.WithKeyTip
            KeyTipBase(keyTip, callback)
            Return Me
        End Function

        Private Function WithDescription(description As String) As IGalleryBuilder Implements IGalleryBuilder.WithDescription
            DescriptionBase(description)
            Return Me
        End Function

        Private Function WithDescription(description As String, callback As FromControl(Of String)) As IGalleryBuilder Implements IGalleryBuilder.WithDescription
            DescriptionBase(description, callback)
            Return Me
        End Function

        Private Function GetItemIdFrom(callback As FromControlAndIndex(Of String)) As IGalleryBuilder Implements IGalleryBuilder.GetItemIdFrom
            ItemIDBase(callback)
            Return Me
        End Function

        Private Function GetItemCountFrom(callback As FromControl(Of Integer)) As IGalleryBuilder Implements IGalleryBuilder.GetItemCountFrom
            ItemCountBase(callback)
            Return Me
        End Function

        Private Function GetItemLabelFrom(callback As FromControlAndIndex(Of String)) As IGalleryBuilder Implements IGalleryBuilder.GetItemLabelFrom
            ItemLabelBase(callback)
            Return Me
        End Function

        Private Function ShowLabelsOnChildItems() As IGalleryBuilder Implements IGalleryBuilder.ShowLabelsOnChildItems
            ShowItemLabelBase()
            Return Me
        End Function

        Private Function HideLabelsOnChildItems() As IGalleryBuilder Implements IGalleryBuilder.HideLabelsOnChildItems
            HideItemLabelBase()
            Return Me
        End Function

        Private Function GetItemScreenTipFrom(callback As FromControlAndIndex(Of String)) As IGalleryBuilder Implements IGalleryBuilder.GetItemScreenTipFrom
            ItemScreentipBase(callback)
            Return Me
        End Function

        Private Function GetItemSuperTipFrom(callback As FromControlAndIndex(Of String)) As IGalleryBuilder Implements IGalleryBuilder.GetItemSuperTipFrom
            ItemSupertipBase(callback)
            Return Me
        End Function

        Private Function GetItemImageFrom(callback As FromControlAndIndex(Of IPictureDisp)) As IGalleryBuilder Implements IGalleryBuilder.GetItemImageFrom
            ItemImageBase(callback)
            Return Me
        End Function

        Private Function GetItemImageFrom(callback As FromControlAndIndex(Of String)) As IGalleryBuilder Implements IGalleryBuilder.GetItemImageFrom
            ItemImageBase(callback)
            Return Me
        End Function

        Private Function ShowItemImages() As IGalleryBuilder Implements IGalleryBuilder.ShowItemImages
            ShowItemImageBase()
            Return Me
        End Function

        Private Function HideItemImages() As IGalleryBuilder Implements IGalleryBuilder.HideItemImages
            HideItemImageBase()
            Return Me
        End Function

        Private Function GetSelectedItemIdFrom(getSelected As FromControl(Of String), setSelected As SelectionChanged) As IGalleryBuilder Implements IGalleryBuilder.GetSelectedItemIdFrom
            SelectedBase(getSelected, setSelected)
            Return Me
        End Function

        Private Function GetSelectedItemIndexFrom(getSelected As FromControl(Of Integer), setSelected As SelectionChanged) As IGalleryBuilder Implements IGalleryBuilder.GetSelectedItemIndexFrom
            SelectedBase(getSelected, setSelected)
            Return Me
        End Function

        Private Function ShowImage() As IGalleryBuilder Implements IGalleryBuilder.ShowImage
            ShowImageBase()
            Return Me
        End Function

        Private Function ShowImage(getShowImage As FromControl(Of Boolean)) As IGalleryBuilder Implements IGalleryBuilder.ShowImage
            ShowImageBase(getShowImage)
            Return Me
        End Function

        Private Function HideImage() As IGalleryBuilder Implements IGalleryBuilder.HideImage
            HideImageBase()
            Return Me
        End Function

        Private Function HideImage(getShowImage As FromControl(Of Boolean)) As IGalleryBuilder Implements IGalleryBuilder.HideImage
            HideImageBase(getShowImage)
            Return Me
        End Function

        Private Function Large() As IGalleryBuilder Implements IGalleryBuilder.Large
            LargeBase()
            Return Me
        End Function

        Private Function Large(callback As FromControl(Of ControlSize)) As IGalleryBuilder Implements IGalleryBuilder.Large
            LargeBase(callback)
            Return Me
        End Function

        Private Function Normal() As IGalleryBuilder Implements IGalleryBuilder.Normal
            NormalBase()
            Return Me
        End Function

        Private Function Normal(callback As FromControl(Of ControlSize)) As IGalleryBuilder Implements IGalleryBuilder.Normal
            NormalBase(callback)
            Return Me
        End Function

        Private Function WithRowCount(rowCount As Integer) As IGalleryBuilder Implements IGalleryBuilder.WithRowCount
            RowCountBase(rowCount)
            Return Me
        End Function

        Private Function WithColumnCount(columnCount As Integer) As IGalleryBuilder Implements IGalleryBuilder.WithColumnCount
            ColumnCountBase(columnCount)
            Return Me
        End Function

        Private Function DoNotShowInRibbon() As IGalleryBuilder Implements IGalleryBuilder.DoNotShowInRibbon
            DoNotShowInRibbonBase()
            Return Me
        End Function

        Private Function WithItemDimensions(dimensions As Size) As IGalleryBuilder Implements IGalleryBuilder.WithItemDimensions
            Return WithItemDimensions(dimensions.Height, dimensions.Width)
        End Function

        Private Function WithItemDimensions(height As Integer, width As Integer) As IGalleryBuilder Implements IGalleryBuilder.WithItemDimensions
            ItemHeightBase(height)
            ItemWidthBase(width)
            Return Me
        End Function

        Private Function WithItemDimensions(dimensions As Size, heightCallback As FromControl(Of Integer), widthCallback As FromControl(Of Integer)) As IGalleryBuilder Implements IGalleryBuilder.WithItemDimensions
            Return WithItemDimensions(dimensions.Height, dimensions.Width, heightCallback, widthCallback)
        End Function

        Private Function WithItemDimensions(height As Integer, width As Integer, heightCallback As FromControl(Of Integer), widthCallback As FromControl(Of Integer)) As IGalleryBuilder Implements IGalleryBuilder.WithItemDimensions
            ItemHeightBase(height, heightCallback)
            ItemWidthBase(width, widthCallback)
            Return Me
        End Function


        Private Function RouteSelectionChangeTo(callback As SelectionChanged) As IGalleryBuilder Implements IGalleryBuilder.RouteSelectionChangeTo
            OnActionBase(callback)
            Return Me
        End Function

        Private Function BeforeSelectionChange(ParamArray actions() As Action(Of BeforeSelectionChangeEventArgs)) As IGalleryBuilder Implements IGalleryBuilder.BeforeSelectionChange
            AddBeforeSelectionChangeHandlers(actions)
            Return Me
        End Function

        Private Function AfterSelectionChanged(ParamArray actions() As Action(Of SelectionChangeEventArgs)) As IGalleryBuilder Implements IGalleryBuilder.AfterSelectionChanged
            AddSelectionChangeHandlers(actions)
            Return Me
        End Function

    End Class

End Namespace