Imports System.Drawing
Imports RibbonX.Api
Imports RibbonX.Api.Internal
Imports RibbonX.Callbacks
Imports RibbonX.ComTypes.StdOle
Imports RibbonX.Controls
Imports RibbonX.Controls.BuiltIn
Imports RibbonX.Images.BuiltIn
Imports RibbonX.SimpleTypes

Namespace InternalApi

    Friend NotInheritable Class ComboBoxBuilder
        Inherits BuilderBase(Of ComboBox)
        Implements IComboBoxBuilder

#Region "Enabled and Visible"

        Private Function Enabled() As IComboBoxBuilder Implements IComboBoxBuilder.Enabled
            EnabledBase()
            Return Me
        End Function

        Private Function Enabled(callback As FromControl(Of Boolean)) As IComboBoxBuilder Implements IComboBoxBuilder.Enabled
            EnabledBase(callback)
            Return Me
        End Function

        Private Function Disabled() As IComboBoxBuilder Implements IComboBoxBuilder.Disabled
            DisabledBase()
            Return Me
        End Function

        Private Function Disabled(callback As FromControl(Of Boolean)) As IComboBoxBuilder Implements IComboBoxBuilder.Disabled
            DisabledBase(callback)
            Return Me
        End Function

        Private Function Visible() As IComboBoxBuilder Implements IComboBoxBuilder.Visible
            VisibleBase()
            Return Me
        End Function

        Private Function Visible(callback As FromControl(Of Boolean)) As IComboBoxBuilder Implements IComboBoxBuilder.Visible
            VisibleBase(callback)
            Return Me
        End Function

        Private Function Invisible() As IComboBoxBuilder Implements IComboBoxBuilder.Invisible
            InvisibleBase()
            Return Me
        End Function

        Private Function Invisible(callback As FromControl(Of Boolean)) As IComboBoxBuilder Implements IComboBoxBuilder.Invisible
            InvisibleBase(callback)
            Return Me
        End Function

#End Region

#Region "Display Text"

        Private Function WithLabel(label As String, Optional copyToScreenTip As Boolean = True) As IComboBoxBuilder Implements IComboBoxBuilder.WithLabel
            LabelBase(label)
            ShowLabelBase()

            If copyToScreenTip Then ScreentipBase(label)
            Return Me
        End Function

        Private Function WithLabel(label As String, callback As FromControl(Of String), Optional copyToScreenTip As Boolean = True) As IComboBoxBuilder Implements IComboBoxBuilder.WithLabel
            LabelBase(label, callback)
            ShowLabelBase()

            If copyToScreenTip Then ScreentipBase(label, callback)
            Return Me
        End Function

        Private Function ShowLabel() As IComboBoxBuilder Implements IComboBoxBuilder.ShowLabel
            ShowLabelBase()
            Return Me
        End Function

        Private Function ShowLabel(callback As FromControl(Of Boolean)) As IComboBoxBuilder Implements IComboBoxBuilder.ShowLabel
            ShowLabelBase(callback)
            Return Me
        End Function

        Private Function HideLabel() As IComboBoxBuilder Implements IComboBoxBuilder.HideLabel
            HideLabelBase()
            Return Me
        End Function

        Private Function HideLabel(callback As FromControl(Of Boolean)) As IComboBoxBuilder Implements IComboBoxBuilder.HideLabel
            HideLabelBase(callback)
            Return Me
        End Function

        Private Function WithScreenTip(screenTip As String) As IComboBoxBuilder Implements IComboBoxBuilder.WithScreenTip
            ScreentipBase(screenTip)
            Return Me
        End Function

        Private Function WithScreenTip(screenTip As String, callback As FromControl(Of String)) As IComboBoxBuilder Implements IComboBoxBuilder.WithScreenTip
            ScreentipBase(screenTip, callback)
            Return Me
        End Function

        Private Function WithSuperTip(superTip As String) As IComboBoxBuilder Implements IComboBoxBuilder.WithSuperTip
            SupertipBase(superTip)
            Return Me
        End Function

        Private Function WithSuperTip(superTip As String, callback As FromControl(Of String)) As IComboBoxBuilder Implements IComboBoxBuilder.WithSuperTip
            SupertipBase(superTip, callback)
            Return Me
        End Function

        Private Function WithKeyTip(keyTip As KeyTip) As IComboBoxBuilder Implements IComboBoxBuilder.WithKeyTip
            KeyTipBase(keyTip)
            Return Me
        End Function

        Private Function WithKeyTip(keyTip As KeyTip, callback As FromControl(Of KeyTip)) As IComboBoxBuilder Implements IComboBoxBuilder.WithKeyTip
            KeyTipBase(keyTip, callback)
            Return Me
        End Function

#End Region

#Region "Image, ShowImage, and ShowItemImages"

        Private Function WithImage(image As ImageMSO) As IComboBoxBuilder Implements IComboBoxBuilder.WithImage
            ImageBase(image)
            ShowImageBase()
            Return Me
        End Function

        Public Function WithImage(image As ImageMSO, callback As FromControl(Of String)) As IComboBoxBuilder Implements IComboBoxBuilder.WithImage
            ImageBase(image, callback)
            ShowImageBase()
            Return Me
        End Function

        Private Function WithImage(image As Bitmap, callback As FromControl(Of IPictureDisp)) As IComboBoxBuilder Implements IComboBoxBuilder.WithImage
            ImageBase(image, callback)
            ShowImageBase()
            Return Me
        End Function

        Private Function WithImage(image As Icon, callback As FromControl(Of IPictureDisp)) As IComboBoxBuilder Implements IComboBoxBuilder.WithImage
            ImageBase(image, callback)
            ShowImageBase()
            Return Me
        End Function

        Public Function WithImage(imageId As String, image As Bitmap) As IComboBoxBuilder Implements IComboBoxBuilder.WithImage
            ImageBase(imageId, image)
            ShowImageBase()
            Return Me
        End Function

        Public Function WithImage(imageId As String, image As Icon) As IComboBoxBuilder Implements IComboBoxBuilder.WithImage
            ImageBase(imageId, image)
            ShowImageBase()
            Return Me
        End Function

        Private Function ShowImage() As IComboBoxBuilder Implements IComboBoxBuilder.ShowImage
            ShowImageBase()
            Return Me
        End Function

        Private Function ShowImage(getShowImage As FromControl(Of Boolean)) As IComboBoxBuilder Implements IComboBoxBuilder.ShowImage
            ShowImageBase(getShowImage)
            Return Me
        End Function

        Private Function HideImage() As IComboBoxBuilder Implements IComboBoxBuilder.HideImage
            HideImageBase()
            Return Me
        End Function

        Private Function HideImage(getShowImage As FromControl(Of Boolean)) As IComboBoxBuilder Implements IComboBoxBuilder.HideImage
            HideImageBase(getShowImage)
            Return Me
        End Function

        Private Function ShowItemImages() As IComboBoxBuilder Implements IComboBoxBuilder.ShowItemImages
            ShowItemImageBase()
            Return Me
        End Function

        Private Function HideItemImages() As IComboBoxBuilder Implements IComboBoxBuilder.HideItemImages
            HideItemImageBase()
            Return Me
        End Function

#End Region

#Region "Insert Before/After"

        Private Function InsertBeforeMSO(builtInControl As MSO) As IComboBoxBuilder Implements IComboBoxBuilder.InsertBefore
            InsertBeforeMsoBase(builtInControl)
            Return Me
        End Function

        Private Function InsertBeforeQ(qualifiedControl As IRibbonElement) As IComboBoxBuilder Implements IComboBoxBuilder.InsertBefore
            InsertBeforeQBase(qualifiedControl)
            Return Me
        End Function

        Private Function InsertAfterMSO(builtInControl As MSO) As IComboBoxBuilder Implements IComboBoxBuilder.InsertAfter
            InsertAfterMsoBase(builtInControl)
            Return Me
        End Function

        Private Function InsertAfterQ(qualifiedControl As IRibbonElement) As IComboBoxBuilder Implements IComboBoxBuilder.InsertAfter
            InsertAfterQBase(qualifiedControl)
            Return Me
        End Function

#End Region

#Region "Custom Ids"

        Private Function WithId(id As String) As IComboBoxBuilder Implements IComboBoxBuilder.WithId
            WithIdBase(id)
            Return Me
        End Function

        Private Function WithIdQ([namespace] As String, id As String) As IComboBoxBuilder Implements IComboBoxBuilder.WithIdQ
            WithIdBase([namespace], id)
            Return Me
        End Function

        Private Function WithIdMso(mso As MSO) As IComboBoxBuilder Implements IComboBoxBuilder.WithIdMso
            WithIdBase(mso)
            Return Me
        End Function

#End Region

#Region "Size"

        Private Function AsWideAs(sizeString As String) As IComboBoxBuilder Implements IComboBoxBuilder.AsWideAs
            SizeStringBase(sizeString)
            Return Me
        End Function

        Private Function WithMaximumInputLength(maxLength As Byte) As IComboBoxBuilder Implements IComboBoxBuilder.WithMaximumInputLength
            MaxLengthBase(maxLength)
            Return Me
        End Function

#End Region

#Region "Actions"

        Private Function WithText(text As String, getText As FromControl(Of String), setText As TextChanged) As IComboBoxBuilder Implements IComboBoxBuilder.WithText
            TextBase(text, getText, setText)
            Return Me
        End Function

        Public Function WithText(initialValue As String, getText As FromControl(Of String), setText As TextChanged, action As Action(Of IActionBuilder(Of IComboBox, String))) As IComboBoxBuilder Implements IComboBoxBuilder.WithText
            TextBase(initialValue, getText, setText)
            OnActionBase(action, NameOf(IComboBox.Changing), NameOf(IComboBox.Changed))
            Return Me
        End Function

#End Region

#Region "Child-Item Callbacks"

        Private Function GetItemIdFrom(callback As FromControlAndIndex(Of String)) As IComboBoxBuilder Implements IComboBoxBuilder.GetItemIdFrom
            ItemIDBase(callback)
            Return Me
        End Function

        Private Function GetItemCountFrom(callback As FromControl(Of Integer)) As IComboBoxBuilder Implements IComboBoxBuilder.GetItemCountFrom
            ItemCountBase(callback)
            Return Me
        End Function

        Private Function GetItemLabelFrom(callback As FromControlAndIndex(Of String)) As IComboBoxBuilder Implements IComboBoxBuilder.GetItemLabelFrom
            ItemLabelBase(callback)
            Return Me
        End Function

        Private Function GetItemScreenTipFrom(callback As FromControlAndIndex(Of String)) As IComboBoxBuilder Implements IComboBoxBuilder.GetItemScreenTipFrom
            ItemScreentipBase(callback)
            Return Me
        End Function

        Private Function GetItemSuperTipFrom(callback As FromControlAndIndex(Of String)) As IComboBoxBuilder Implements IComboBoxBuilder.GetItemSuperTipFrom
            ItemSupertipBase(callback)
            Return Me
        End Function

        Private Function GetItemImageFrom(callback As FromControlAndIndex(Of IPictureDisp)) As IComboBoxBuilder Implements IComboBoxBuilder.GetItemImageFrom
            ItemImageBase(callback)
            Return Me
        End Function

        Private Function GetItemImageFrom(callback As FromControlAndIndex(Of String)) As IComboBoxBuilder Implements IComboBoxBuilder.GetItemImageFrom
            ItemImageBase(callback)
            Return Me
        End Function

#End Region

        Public Function WithTag(tag As Object) As IComboBoxBuilder Implements ITag(Of IComboBoxBuilder).WithTag
            WithTagBase(tag)
            Return Me
        End Function

        Public Function FromTemplate(template As IRibbonElement) As IComboBoxBuilder Implements IComboBoxBuilder.FromTemplate
            FromTemplateBase(template)
            Return Me
        End Function

        Public Overrides Function Build() As ComboBox
            Return New ComboBox(ControlProperties, Tag)
        End Function

        Public Shared Function FromAction(Optional action As Action(Of IComboBoxBuilder) = Nothing) As IComboBox
            Dim instance As ComboBoxBuilder = New ComboBoxBuilder()

            action?.Invoke(instance)

            Return instance.Build()
        End Function

    End Class

End Namespace