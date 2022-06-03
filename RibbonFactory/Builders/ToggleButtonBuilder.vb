Imports System.Drawing
Imports RibbonX.BuilderInterfaces.API
Imports RibbonX.Controls
Imports RibbonX.Controls.Events
Imports RibbonX.Enums
Imports RibbonX.Enums.ImageMSO
Imports RibbonX.Enums.MSO
Imports RibbonX.ComTypes.StdOle

Namespace Builders

    Friend Class ToggleButtonBuilder
        Inherits BuilderBase(Of ToggleButton)
        Implements IToggleButtonBuilder

        Public Sub New()
            MyBase.New()
        End Sub

        Public Sub New(template As RibbonElement)
            MyBase.New(template)
        End Sub

        Private Function WithId(id As String) As IToggleButtonBuilder Implements IToggleButtonBuilder.WithId
            WithIdBase(id)
            Return Me
        End Function

        Private Function WithIdQ([namespace] As String, id As String) As IToggleButtonBuilder Implements IToggleButtonBuilder.WithIdQ
            WithIdBase([namespace], id)
            Return Me
        End Function

        Private Function WithIdMso(mso As MSO) As IToggleButtonBuilder Implements IToggleButtonBuilder.WithIdMso
            WithIdBase(mso)
            Return Me
        End Function

        Private Function Enabled() As IToggleButtonBuilder Implements IToggleButtonBuilder.Enabled
            EnabledBase()
            Return Me
        End Function

        Private Function Enabled(callback As FromControl(Of Boolean)) As IToggleButtonBuilder Implements IToggleButtonBuilder.Enabled
            EnabledBase(callback)
            Return Me
        End Function

        Private Function Disabled() As IToggleButtonBuilder Implements IToggleButtonBuilder.Disabled
            DisabledBase()
            Return Me
        End Function

        Private Function Disabled(callback As FromControl(Of Boolean)) As IToggleButtonBuilder Implements IToggleButtonBuilder.Disabled
            DisabledBase(callback)
            Return Me
        End Function

        Private Function Visible() As IToggleButtonBuilder Implements IToggleButtonBuilder.Visible
            VisibleBase()
            Return Me
        End Function

        Private Function Visible(callback As FromControl(Of Boolean)) As IToggleButtonBuilder Implements IToggleButtonBuilder.Visible
            VisibleBase(callback)
            Return Me
        End Function

        Private Function Invisible() As IToggleButtonBuilder Implements IToggleButtonBuilder.Invisible
            InvisibleBase()
            Return Me
        End Function

        Private Function Invisible(callback As FromControl(Of Boolean)) As IToggleButtonBuilder Implements IToggleButtonBuilder.Invisible
            InvisibleBase(callback)
            Return Me
        End Function

        Private Function WithLabel(label As String, Optional copyToScreenTip As Boolean = True) As IToggleButtonBuilder Implements IToggleButtonBuilder.WithLabel
            LabelBase(label)
            ShowLabelBase()
            If copyToScreenTip Then ScreentipBase(label)
            Return Me
        End Function

        Private Function WithLabel(label As String, callback As FromControl(Of String), Optional copyToScreenTip As Boolean = True) As IToggleButtonBuilder Implements IToggleButtonBuilder.WithLabel
            LabelBase(label, callback)
            ShowLabelBase()
            If copyToScreenTip Then ScreentipBase(label, callback)
            Return Me
        End Function

        Private Function ShowLabel() As IToggleButtonBuilder Implements IToggleButtonBuilder.ShowLabel
            ShowLabelBase()
            Return Me
        End Function

        Private Function ShowLabel(callback As FromControl(Of Boolean)) As IToggleButtonBuilder Implements IToggleButtonBuilder.ShowLabel
            ShowLabelBase(callback)
            Return Me
        End Function

        Private Function HideLabel() As IToggleButtonBuilder Implements IToggleButtonBuilder.HideLabel
            HideLabelBase()
            Return Me
        End Function

        Private Function HideLabel(callback As FromControl(Of Boolean)) As IToggleButtonBuilder Implements IToggleButtonBuilder.HideLabel
            HideLabelBase(callback)
            Return Me
        End Function

        Private Function WithScreenTip(screenTip As String) As IToggleButtonBuilder Implements IToggleButtonBuilder.WithScreenTip
            ScreentipBase(screenTip)
            Return Me
        End Function

        Private Function WithScreenTip(screenTip As String, callback As FromControl(Of String)) As IToggleButtonBuilder Implements IToggleButtonBuilder.WithScreenTip
            ScreentipBase(screenTip, callback)
            Return Me
        End Function

        Private Function WithSuperTip(superTip As String) As IToggleButtonBuilder Implements IToggleButtonBuilder.WithSuperTip
            SupertipBase(superTip)
            Return Me
        End Function

        Private Function WithSuperTip(superTip As String, callback As FromControl(Of String)) As IToggleButtonBuilder Implements IToggleButtonBuilder.WithSuperTip
            SupertipBase(superTip, callback)
            Return Me
        End Function

        Private Function WithKeyTip(keyTip As KeyTip) As IToggleButtonBuilder Implements IToggleButtonBuilder.WithKeyTip
            KeyTipBase(keyTip)
            Return Me
        End Function

        Private Function WithKeyTip(keyTip As KeyTip, callback As FromControl(Of KeyTip)) As IToggleButtonBuilder Implements IToggleButtonBuilder.WithKeyTip
            KeyTipBase(keyTip, callback)
            Return Me
        End Function

        Private Function WithImage(image As ImageMSO) As IToggleButtonBuilder Implements IToggleButtonBuilder.WithImage
            ImageBase(image)
            Return Me
        End Function

        Private Function WithImage(image As Bitmap, callback As FromControl(Of IPictureDisp)) As IToggleButtonBuilder Implements IToggleButtonBuilder.WithImage
            ImageBase(image, callback)
            Return Me
        End Function

        Private Function WithImage(image As Icon, callback As FromControl(Of IPictureDisp)) As IToggleButtonBuilder Implements IToggleButtonBuilder.WithImage
            ImageBase(image, callback)
            Return Me
        End Function

        Private Function WithImage(imagePath As String) As IToggleButtonBuilder Implements IToggleButtonBuilder.WithImage
            ImageBase(imagePath)
            Return Me
        End Function

        Private Function WithDescription(description As String) As IToggleButtonBuilder Implements IToggleButtonBuilder.WithDescription
            DescriptionBase(description)
            Return Me
        End Function

        Private Function WithDescription(description As String, callback As FromControl(Of String)) As IToggleButtonBuilder Implements IToggleButtonBuilder.WithDescription
            DescriptionBase(description, callback)
            Return Me
        End Function

        Private Function Large() As IToggleButtonBuilder Implements IToggleButtonBuilder.Large
            LargeBase()
            Return Me
        End Function

        Private Function Large(callback As FromControl(Of ControlSize)) As IToggleButtonBuilder Implements IToggleButtonBuilder.Large
            LargeBase(callback)
            Return Me
        End Function

        Private Function Normal() As IToggleButtonBuilder Implements IToggleButtonBuilder.Normal
            NormalBase()
            Return Me
        End Function

        Private Function Normal(callback As FromControl(Of ControlSize)) As IToggleButtonBuilder Implements IToggleButtonBuilder.Normal
            NormalBase(callback)
            Return Me
        End Function

        Private Function InsertBeforeMSO(builtInControl As MSO) As IToggleButtonBuilder Implements IToggleButtonBuilder.InsertBefore
            InsertBeforeMsoBase(builtInControl)
            Return Me
        End Function

        Private Function InsertBeforeQ(qualifiedControl As RibbonElement) As IToggleButtonBuilder Implements IToggleButtonBuilder.InsertBefore
            InsertBeforeQBase(qualifiedControl)
            Return Me
        End Function

        Private Function InsertAfterMSO(builtInControl As MSO) As IToggleButtonBuilder Implements IToggleButtonBuilder.InsertAfter
            InsertAfterMsoBase(builtInControl)
            Return Me
        End Function

        Private Function InsertAfterQ(qualifiedControl As RibbonElement) As IToggleButtonBuilder Implements IToggleButtonBuilder.InsertAfter
            InsertAfterQBase(qualifiedControl)
            Return Me
        End Function

        Private Function ShowImage() As IToggleButtonBuilder Implements IToggleButtonBuilder.ShowImage
            ShowImageBase()
            Return Me
        End Function

        Private Function ShowImage(getShowImage As FromControl(Of Boolean)) As IToggleButtonBuilder Implements IToggleButtonBuilder.ShowImage
            ShowImageBase(getShowImage)
            Return Me
        End Function

        Private Function HideImage() As IToggleButtonBuilder Implements IToggleButtonBuilder.HideImage
            HideImageBase()
            Return Me
        End Function

        Private Function HideImage(getShowImage As FromControl(Of Boolean)) As IToggleButtonBuilder Implements IToggleButtonBuilder.HideImage
            HideImageBase(getShowImage)
            Return Me
        End Function

        Private Function Checked(getChecked As FromControl(Of Boolean), setChecked As ButtonPressed) As IToggleButtonBuilder Implements IToggleButtonBuilder.Checked
            PressedBase(True, getChecked, setChecked)
            Return Me
        End Function

        Private Function Unchecked(getChecked As FromControl(Of Boolean), setChecked As ButtonPressed) As IToggleButtonBuilder Implements IToggleButtonBuilder.Unchecked
            PressedBase(False, getChecked, setChecked)
            Return Me
        End Function

        Private Function BeforeStateChange(ParamArray actions() As Action(Of BeforeToggleChangeEventArgs)) As IToggleButtonBuilder Implements IToggleButtonBuilder.BeforeStateChange
            AddBeforeToggleHandlers(actions)
            Return Me
        End Function

        Private Function AfterStateChange(ParamArray actions() As Action(Of ToggleChangeEventArgs)) As IToggleButtonBuilder Implements IToggleButtonBuilder.AfterStateChange
            AddToggleHandlers(actions)
            Return Me
        End Function

    End Class

End Namespace