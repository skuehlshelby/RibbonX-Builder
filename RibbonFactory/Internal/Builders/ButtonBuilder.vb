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
    Friend NotInheritable Class ButtonBuilder
        Inherits BuilderBase(Of Button)
        Implements IButtonBuilder

        Public Overrides Function Build() As Button
            Return New Button(ControlProperties, Tag)
        End Function

        Public Function WithId(id As String) As IButtonBuilder Implements IButtonBuilder.WithId
            WithIdBase(id)
            Return Me
        End Function

        Public Function WithIdQ([namespace] As String, id As String) As IButtonBuilder Implements IButtonBuilder.WithIdQ
            WithIdBase([namespace], id)
            Return Me
        End Function

        Public Function WithIdMso(mso As MSO) As IButtonBuilder Implements IButtonBuilder.WithIdMso
            WithIdBase(mso)
            Return Me
        End Function

        Public Function Enabled() As IButtonBuilder Implements IButtonBuilder.Enabled
            EnabledBase()
            Return Me
        End Function

        Public Function Enabled(callback As FromControl(Of Boolean)) As IButtonBuilder Implements IButtonBuilder.Enabled
            EnabledBase(callback)
            Return Me
        End Function

        Public Function Disabled() As IButtonBuilder Implements IButtonBuilder.Disabled
            DisabledBase()
            Return Me
        End Function

        Public Function Disabled(callback As FromControl(Of Boolean)) As IButtonBuilder Implements IButtonBuilder.Disabled
            DisabledBase(callback)
            Return Me
        End Function

        Public Function Visible() As IButtonBuilder Implements IButtonBuilder.Visible
            VisibleBase()
            Return Me
        End Function

        Public Function Visible(callback As FromControl(Of Boolean)) As IButtonBuilder Implements IButtonBuilder.Visible
            VisibleBase(callback)
            Return Me
        End Function

        Public Function Invisible() As IButtonBuilder Implements IButtonBuilder.Invisible
            InvisibleBase()
            Return Me
        End Function

        Public Function Invisible(callback As FromControl(Of Boolean)) As IButtonBuilder Implements IButtonBuilder.Invisible
            InvisibleBase(callback)
            Return Me
        End Function

        Public Function WithLabel(label As String, Optional copyToScreenTip As Boolean = True) As IButtonBuilder Implements IButtonBuilder.WithLabel
            LabelBase(label)
            ShowLabelBase()
            If copyToScreenTip Then ScreentipBase(label)
            Return Me
        End Function

        Public Function WithLabel(label As String, callback As FromControl(Of String), Optional copyToScreenTip As Boolean = True) As IButtonBuilder Implements IButtonBuilder.WithLabel
            LabelBase(label, callback)
            ShowLabelBase()
            If copyToScreenTip Then ScreentipBase(label, callback)
            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String) As IButtonBuilder Implements IButtonBuilder.WithScreenTip
            ScreentipBase(screenTip)
            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String, callback As FromControl(Of String)) As IButtonBuilder Implements IButtonBuilder.WithScreenTip
            ScreentipBase(screenTip, callback)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String) As IButtonBuilder Implements IButtonBuilder.WithSuperTip
            SupertipBase(superTip)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String, callback As FromControl(Of String)) As IButtonBuilder Implements IButtonBuilder.WithSuperTip
            SupertipBase(superTip, callback)
            Return Me
        End Function

        Public Function ShowLabel() As IButtonBuilder Implements IButtonBuilder.ShowLabel
            ShowLabelBase()
            Return Me
        End Function

        Public Function ShowLabel(callback As FromControl(Of Boolean)) As IButtonBuilder Implements IButtonBuilder.ShowLabel
            ShowLabelBase(callback)
            Return Me
        End Function

        Public Function HideLabel() As IButtonBuilder Implements IButtonBuilder.HideLabel
            HideLabelBase()
            Return Me
        End Function

        Public Function HideLabel(callback As FromControl(Of Boolean)) As IButtonBuilder Implements IButtonBuilder.HideLabel
            HideLabelBase(callback)
            Return Me
        End Function

        Public Function Large() As IButtonBuilder Implements IButtonBuilder.Large
            LargeBase()
            Return Me
        End Function

        Public Function Large(callback As FromControl(Of ControlSize)) As IButtonBuilder Implements IButtonBuilder.Large
            LargeBase(callback)
            Return Me
        End Function

        Public Function Normal() As IButtonBuilder Implements IButtonBuilder.Normal
            NormalBase()
            Return Me
        End Function

        Public Function Normal(callback As FromControl(Of ControlSize)) As IButtonBuilder Implements IButtonBuilder.Normal
            NormalBase(callback)
            Return Me
        End Function

        Public Function WithImage(image As ImageMSO) As IButtonBuilder Implements IButtonBuilder.WithImage
            ImageBase(image)
            ShowImageBase()
            Return Me
        End Function

        Public Function WithImage(image As ImageMSO, callback As FromControl(Of String)) As IButtonBuilder Implements IButtonBuilder.WithImage
            ImageBase(image, callback)
            ShowImageBase()
            Return Me
        End Function

        Public Function WithImage(image As Bitmap, callback As FromControl(Of IPictureDisp)) As IButtonBuilder Implements IButtonBuilder.WithImage
            ImageBase(image, callback)
            ShowImageBase()
            Return Me
        End Function

        Public Function WithImage(image As Icon, callback As FromControl(Of IPictureDisp)) As IButtonBuilder Implements IButtonBuilder.WithImage
            ImageBase(image, callback)
            ShowImageBase()
            Return Me
        End Function

        Public Function WithImage(imageId As String, image As Bitmap) As IButtonBuilder Implements IButtonBuilder.WithImage
            ImageBase(imageId, image)
            ShowImageBase()
            Return Me
        End Function

        Public Function WithImage(imageId As String, image As Icon) As IButtonBuilder Implements IButtonBuilder.WithImage
            ImageBase(imageId, image)
            ShowImageBase()
            Return Me
        End Function

        Public Function WithDescription(description As String) As IButtonBuilder Implements IButtonBuilder.WithDescription
            DescriptionBase(description)
            Return Me
        End Function

        Public Function WithDescription(description As String, callback As FromControl(Of String)) As IButtonBuilder Implements IButtonBuilder.WithDescription
            DescriptionBase(description, callback)
            Return Me
        End Function

        Public Function WithKeyTip(keyTip As KeyTip) As IButtonBuilder Implements IButtonBuilder.WithKeyTip
            KeyTipBase(keyTip)
            Return Me
        End Function

        Public Function WithKeyTip(keyTip As KeyTip, callback As FromControl(Of KeyTip)) As IButtonBuilder Implements IButtonBuilder.WithKeyTip
            KeyTipBase(keyTip, callback)
            Return Me
        End Function

        Public Function InsertBeforeMSO(builtInControl As MSO) As IButtonBuilder Implements IButtonBuilder.InsertBefore
            InsertBeforeMsoBase(builtInControl)
            Return Me
        End Function

        Public Function InsertBeforeQ(qualifiedControl As IRibbonElement) As IButtonBuilder Implements IButtonBuilder.InsertBefore
            InsertBeforeQBase(qualifiedControl)
            Return Me
        End Function

        Public Function InsertAfterMSO(builtInControl As MSO) As IButtonBuilder Implements IButtonBuilder.InsertAfter
            InsertAfterMsoBase(builtInControl)
            Return Me
        End Function

        Public Function InsertAfterQ(qualifiedControl As IRibbonElement) As IButtonBuilder Implements IButtonBuilder.InsertAfter
            InsertAfterQBase(qualifiedControl)
            Return Me
        End Function

        Public Function ShowImage() As IButtonBuilder Implements IButtonBuilder.ShowImage
            ShowImageBase()
            Return Me
        End Function

        Public Function ShowImage(getShowImage As FromControl(Of Boolean)) As IButtonBuilder Implements IButtonBuilder.ShowImage
            ShowImageBase(getShowImage)
            Return Me
        End Function

        Public Function HideImage() As IButtonBuilder Implements IButtonBuilder.HideImage
            HideImageBase()
            Return Me
        End Function

        Public Function HideImage(getShowImage As FromControl(Of Boolean)) As IButtonBuilder Implements IButtonBuilder.HideImage
            HideImageBase(getShowImage)
            Return Me
        End Function

        Public Function WithTag(tag As Object) As IButtonBuilder Implements IButtonBuilder.WithTag
            WithTagBase(tag)
            Return Me
        End Function

        Public Function OnClick(callback As OnAction) As IButtonBuilder Implements IButtonBuilder.OnClick
            OnActionBase(callback)
            Return Me
        End Function

        Public Function OnClick(callback As OnAction, action As Action(Of IActionBuilder(Of IButton))) As IButtonBuilder Implements IButtonBuilder.OnClick
            OnActionBase(callback)
            OnActionBase(action, NameOf(IButton.Clicking), NameOf(IButton.Clicked))
            Return Me
        End Function

        Public Function FromTemplate(template As IRibbonElement) As IButtonBuilder Implements IButtonBuilder.FromTemplate
            FromTemplateBase(template)
            Return Me
        End Function

        Public Shared Function FromAction(Optional action As Action(Of IButtonBuilder) = Nothing) As IButton
            Dim instance As ButtonBuilder = New ButtonBuilder()

            action?.Invoke(instance)

            Return instance.Build()
        End Function

    End Class

End Namespace