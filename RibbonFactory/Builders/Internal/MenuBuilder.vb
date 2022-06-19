Imports System.Drawing
Imports RibbonX.Controls
Imports RibbonX.Controls.Base
Imports RibbonX.Controls.BuiltIn
Imports RibbonX.ComTypes.StdOle
Imports RibbonX.SimpleTypes
Imports RibbonX.Callbacks
Imports RibbonX.Builders.Internal.Base
Imports RibbonX.Images.BuiltIn
Imports RibbonX.Builders.Internal.Standardization

Namespace Builders
    Friend NotInheritable Class MenuBuilder
        Inherits BuilderBase(Of Menu)
        Implements IMenuBuilder

        Public Sub New()
            MyBase.New()
        End Sub

        Public Sub New(template As RibbonElement)
            MyBase.New(template)
        End Sub

        Public Function Enabled() As IMenuBuilder Implements IMenuBuilder.Enabled
            EnabledBase()
            Return Me
        End Function

        Public Function Enabled(callback As FromControl(Of Boolean)) As IMenuBuilder Implements IMenuBuilder.Enabled
            EnabledBase(callback)
            Return Me
        End Function

        Public Function Disabled() As IMenuBuilder Implements IMenuBuilder.Disabled
            DisabledBase()
            Return Me
        End Function

        Public Function Disabled(callback As FromControl(Of Boolean)) As IMenuBuilder Implements IMenuBuilder.Disabled
            DisabledBase(callback)
            Return Me
        End Function

        Public Function Visible() As IMenuBuilder Implements IMenuBuilder.Visible
            VisibleBase()
            Return Me
        End Function

        Public Function Visible(callback As FromControl(Of Boolean)) As IMenuBuilder Implements IMenuBuilder.Visible
            VisibleBase(callback)
            Return Me
        End Function

        Public Function Invisible() As IMenuBuilder Implements IMenuBuilder.Invisible
            InvisibleBase()
            Return Me
        End Function

        Public Function Invisible(callback As FromControl(Of Boolean)) As IMenuBuilder Implements IMenuBuilder.Invisible
            InvisibleBase(callback)
            Return Me
        End Function

        Public Function WithLabel(label As String, Optional copyToScreenTip As Boolean = True) As IMenuBuilder Implements IMenuBuilder.WithLabel
            LabelBase(label)
            ShowLabelBase()
            If copyToScreenTip Then ScreentipBase(label)
            Return Me
        End Function

        Public Function WithLabel(label As String, callback As FromControl(Of String), Optional copyToScreenTip As Boolean = True) As IMenuBuilder Implements IMenuBuilder.WithLabel
            LabelBase(label, callback)
            ShowLabelBase()
            If copyToScreenTip Then ScreentipBase(label, callback)
            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String) As IMenuBuilder Implements IMenuBuilder.WithScreenTip
            ScreentipBase(screenTip)
            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String, callback As FromControl(Of String)) As IMenuBuilder Implements IMenuBuilder.WithScreenTip
            ScreentipBase(screenTip, callback)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String) As IMenuBuilder Implements IMenuBuilder.WithSuperTip
            SupertipBase(superTip)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String, callback As FromControl(Of String)) As IMenuBuilder Implements IMenuBuilder.WithSuperTip
            SupertipBase(superTip, callback)
            Return Me
        End Function

        Public Function ShowLabel() As IMenuBuilder Implements IMenuBuilder.ShowLabel
            ShowLabelBase()
            Return Me
        End Function

        Public Function ShowLabel(callback As FromControl(Of Boolean)) As IMenuBuilder Implements IMenuBuilder.ShowLabel
            ShowLabelBase(callback)
            Return Me
        End Function

        Public Function HideLabel() As IMenuBuilder Implements IMenuBuilder.HideLabel
            HideLabelBase()
            Return Me
        End Function

        Public Function HideLabel(callback As FromControl(Of Boolean)) As IMenuBuilder Implements IMenuBuilder.HideLabel
            HideLabelBase(callback)
            Return Me
        End Function

        Public Function WithImage(image As ImageMSO) As IMenuBuilder Implements IMenuBuilder.WithImage
            ImageBase(image)
            ShowImageBase()
            Return Me
        End Function

        Public Function WithImage(image As Bitmap, callback As FromControl(Of IPictureDisp)) As IMenuBuilder Implements IMenuBuilder.WithImage
            ImageBase(image, callback)
            ShowImageBase()
            Return Me
        End Function

        Public Function WithImage(image As Icon, callback As FromControl(Of IPictureDisp)) As IMenuBuilder Implements IMenuBuilder.WithImage
            ImageBase(image, callback)
            ShowImageBase()
            Return Me
        End Function

        Public Function WithImage(imagePath As String) As IMenuBuilder Implements IMenuBuilder.WithImage
            ImageBase(imagePath)
            ShowImageBase()
            Return Me
        End Function

        Public Function ShowImage() As IMenuBuilder Implements IShowImage(Of IMenuBuilder).ShowImage
            ShowImageBase()
            Return Me
        End Function

        Public Function ShowImage(getShowImage As FromControl(Of Boolean)) As IMenuBuilder Implements IShowImage(Of IMenuBuilder).ShowImage
            ShowImageBase(getShowImage)
            Return Me
        End Function

        Public Function HideImage() As IMenuBuilder Implements IShowImage(Of IMenuBuilder).HideImage
            HideImageBase()
            Return Me
        End Function

        Public Function HideImage(getShowImage As FromControl(Of Boolean)) As IMenuBuilder Implements IShowImage(Of IMenuBuilder).HideImage
            HideImageBase(getShowImage)
            Return Me
        End Function

        Public Function WithDescription(description As String) As IMenuBuilder Implements IMenuBuilder.WithDescription
            DescriptionBase(description)
            Return Me
        End Function

        Public Function WithDescription(description As String, callback As FromControl(Of String)) As IMenuBuilder Implements IMenuBuilder.WithDescription
            DescriptionBase(description, callback)
            Return Me
        End Function

        Public Function Large() As IMenuBuilder Implements IMenuBuilder.Large
            LargeBase()
            Return Me
        End Function

        Public Function Large(callback As FromControl(Of ControlSize)) As IMenuBuilder Implements IMenuBuilder.Large
            LargeBase(callback)
            Return Me
        End Function

        Public Function Normal() As IMenuBuilder Implements IMenuBuilder.Normal
            NormalBase()
            Return Me
        End Function

        Public Function Normal(callback As FromControl(Of ControlSize)) As IMenuBuilder Implements IMenuBuilder.Normal
            NormalBase(callback)
            Return Me
        End Function

        Public Function WithLargeItems() As IMenuBuilder Implements IMenuBuilder.WithLargeItems
            LargeItemsBase()
            Return Me
        End Function

        Public Function WithNormalSizeItems() As IMenuBuilder Implements IMenuBuilder.WithNormalSizeItems
            NormalItemsBase()
            Return Me
        End Function

        Public Function WithKeyTip(keyTip As KeyTip) As IMenuBuilder Implements IMenuBuilder.WithKeyTip
            KeyTipBase(keyTip)
            Return Me
        End Function

        Public Function WithKeyTip(keyTip As KeyTip, callback As FromControl(Of KeyTip)) As IMenuBuilder Implements IMenuBuilder.WithKeyTip
            KeyTipBase(keyTip, callback)
            Return Me
        End Function

        Public Function InsertBeforeMSO(builtInControl As MSO) As IMenuBuilder Implements IMenuBuilder.InsertBefore
            InsertBeforeMsoBase(builtInControl)
            Return Me
        End Function

        Public Function InsertBeforeQ(qualifiedControl As RibbonElement) As IMenuBuilder Implements IMenuBuilder.InsertBefore
            InsertBeforeQBase(qualifiedControl)
            Return Me
        End Function

        Public Function InsertAfterMSO(builtInControl As MSO) As IMenuBuilder Implements IMenuBuilder.InsertAfter
            InsertAfterMsoBase(builtInControl)
            Return Me
        End Function

        Public Function InsertAfterQ(qualifiedControl As RibbonElement) As IMenuBuilder Implements IMenuBuilder.InsertAfter
            InsertAfterQBase(qualifiedControl)
            Return Me
        End Function

    End Class

End Namespace