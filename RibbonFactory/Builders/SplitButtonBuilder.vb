Imports RibbonFactory.BuilderInterfaces.API
Imports RibbonFactory.Controls
Imports RibbonFactory.Enums
Imports RibbonFactory.Enums.MSO

Namespace Builders
    Friend NotInheritable Class SplitButtonBuilder
        Inherits BuilderBase(Of SplitButton)
        Implements ISplitButtonBuilder

        Public Sub New()
            MyBase.New()
        End Sub

        Public Sub New(template As RibbonElement)
            MyBase.New(template)
        End Sub

        Private Function WithId(id As String) As ISplitButtonBuilder Implements ISplitButtonBuilder.WithId
            WithIdBase(id)
            Return Me
        End Function

        Private Function WithIdQ([namespace] As String, id As String) As ISplitButtonBuilder Implements ISplitButtonBuilder.WithIdQ
            WithIdBase([namespace], id)
            Return Me
        End Function

        Private Function WithIdMso(mso As MSO) As ISplitButtonBuilder Implements ISplitButtonBuilder.WithIdMso
            WithIdBase(mso)
            Return Me
        End Function

        Private Function Enabled() As ISplitButtonBuilder Implements ISplitButtonBuilder.Enabled
            EnabledBase()
            Return Me
        End Function

        Private Function Enabled(callback As FromControl(Of Boolean)) As ISplitButtonBuilder Implements ISplitButtonBuilder.Enabled
            EnabledBase(callback)
            Return Me
        End Function

        Private Function Disabled() As ISplitButtonBuilder Implements ISplitButtonBuilder.Disabled
            DisabledBase()
            Return Me
        End Function

        Private Function Disabled(callback As FromControl(Of Boolean)) As ISplitButtonBuilder Implements ISplitButtonBuilder.Disabled
            DisabledBase(callback)
            Return Me
        End Function

        Private Function Visible() As ISplitButtonBuilder Implements ISplitButtonBuilder.Visible
            VisibleBase()
            Return Me
        End Function

        Private Function Visible(callback As FromControl(Of Boolean)) As ISplitButtonBuilder Implements ISplitButtonBuilder.Visible
            VisibleBase(callback)
            Return Me
        End Function

        Private Function Invisible() As ISplitButtonBuilder Implements ISplitButtonBuilder.Invisible
            InvisibleBase()
            Return Me
        End Function

        Private Function Invisible(callback As FromControl(Of Boolean)) As ISplitButtonBuilder Implements ISplitButtonBuilder.Invisible
            InvisibleBase(callback)
            Return Me
        End Function

        Private Function WithKeyTip(keyTip As KeyTip) As ISplitButtonBuilder Implements ISplitButtonBuilder.WithKeyTip
            KeyTipBase(keyTip)
            Return Me
        End Function

        Private Function WithKeyTip(keyTip As KeyTip, callback As FromControl(Of KeyTip)) As ISplitButtonBuilder Implements ISplitButtonBuilder.WithKeyTip
            KeyTipBase(keyTip, callback)
            Return Me
        End Function

        Private Function ShowLabel() As ISplitButtonBuilder Implements ISplitButtonBuilder.ShowLabel
            ShowLabelBase()
            Return Me
        End Function

        Private Function ShowLabel(callback As FromControl(Of Boolean)) As ISplitButtonBuilder Implements ISplitButtonBuilder.ShowLabel
            ShowLabelBase(callback)
            Return Me
        End Function

        Private Function HideLabel() As ISplitButtonBuilder Implements ISplitButtonBuilder.HideLabel
            HideLabelBase()
            Return Me
        End Function

        Private Function HideLabel(callback As FromControl(Of Boolean)) As ISplitButtonBuilder Implements ISplitButtonBuilder.HideLabel
            HideLabelBase(callback)
            Return Me
        End Function

        Public Function Large() As ISplitButtonBuilder Implements ISplitButtonBuilder.Large
            LargeBase()
            Return Me
        End Function

        Public Function Large(callback As FromControl(Of ControlSize)) As ISplitButtonBuilder Implements ISplitButtonBuilder.Large
            LargeBase(callback)
            Return Me
        End Function

        Public Function Normal() As ISplitButtonBuilder Implements ISplitButtonBuilder.Normal
            NormalBase()
            Return Me
        End Function

        Public Function Normal(callback As FromControl(Of ControlSize)) As ISplitButtonBuilder Implements ISplitButtonBuilder.Normal
            NormalBase(callback)
            Return Me
        End Function

        Private Function InsertBeforeMSO(builtInControl As MSO) As ISplitButtonBuilder Implements ISplitButtonBuilder.InsertBefore
            InsertBeforeMsoBase(builtInControl)
            Return Me
        End Function

        Private Function InsertBeforeQ(qualifiedControl As RibbonElement) As ISplitButtonBuilder Implements ISplitButtonBuilder.InsertBefore
            InsertBeforeQBase(qualifiedControl)
            Return Me
        End Function

        Private Function InsertAfterMSO(builtInControl As MSO) As ISplitButtonBuilder Implements ISplitButtonBuilder.InsertAfter
            InsertAfterMsoBase(builtInControl)
            Return Me
        End Function

        Private Function InsertAfterQ(qualifiedControl As RibbonElement) As ISplitButtonBuilder Implements ISplitButtonBuilder.InsertAfter
            InsertAfterQBase(qualifiedControl)
            Return Me
        End Function

    End Class

End Namespace