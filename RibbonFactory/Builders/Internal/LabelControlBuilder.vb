Imports RibbonX.Controls
Imports RibbonX.Controls.Base
Imports RibbonX.Controls.BuiltIn
Imports RibbonX.Callbacks
Imports RibbonX.Builders.Internal.Base


Namespace Builders
    Friend Class LabelControlBuilder
        Inherits BuilderBase(Of LabelControl)
        Implements ILabelControlBuilder

        Public Sub New()
        End Sub

        Public Sub New(template As RibbonElement)
            MyBase.New(template)
        End Sub

        Private Function WithId(id As String) As ILabelControlBuilder Implements ILabelControlBuilder.WithId
            WithIdBase(id)
            Return Me
        End Function

        Private Function WithIdQ([namespace] As String, id As String) As ILabelControlBuilder Implements ILabelControlBuilder.WithIdQ
            WithIdBase([namespace], id)
            Return Me
        End Function

        Private Function WithIdMso(mso As MSO) As ILabelControlBuilder Implements ILabelControlBuilder.WithIdMso
            WithIdBase(mso)
            Return Me
        End Function

        Private Function Enabled() As ILabelControlBuilder Implements ILabelControlBuilder.Enabled
            EnabledBase()
            Return Me
        End Function

        Private Function Enabled(callback As FromControl(Of Boolean)) As ILabelControlBuilder Implements ILabelControlBuilder.Enabled
            EnabledBase(callback)
            Return Me
        End Function

        Private Function Disabled() As ILabelControlBuilder Implements ILabelControlBuilder.Disabled
            DisabledBase()
            Return Me
        End Function

        Private Function Disabled(callback As FromControl(Of Boolean)) As ILabelControlBuilder Implements ILabelControlBuilder.Disabled
            DisabledBase(callback)
            Return Me
        End Function

        Private Function Visible() As ILabelControlBuilder Implements ILabelControlBuilder.Visible
            VisibleBase()
            Return Me
        End Function

        Private Function Visible(callback As FromControl(Of Boolean)) As ILabelControlBuilder Implements ILabelControlBuilder.Visible
            VisibleBase(callback)
            Return Me
        End Function

        Private Function Invisible() As ILabelControlBuilder Implements ILabelControlBuilder.Invisible
            InvisibleBase()
            Return Me
        End Function

        Private Function Invisible(callback As FromControl(Of Boolean)) As ILabelControlBuilder Implements ILabelControlBuilder.Invisible
            InvisibleBase(callback)
            Return Me
        End Function

        Private Function WithLabel(label As String, Optional copyToScreenTip As Boolean = True) As ILabelControlBuilder Implements ILabelControlBuilder.WithLabel
            LabelBase(label)
            ShowLabelBase()

            If copyToScreenTip Then ScreentipBase(label)
            Return Me
        End Function

        Private Function WithLabel(label As String, callback As FromControl(Of String), Optional copyToScreenTip As Boolean = True) As ILabelControlBuilder Implements ILabelControlBuilder.WithLabel
            LabelBase(label, callback)
            ShowLabelBase()

            If copyToScreenTip Then ScreentipBase(label, callback)
            Return Me
        End Function

        Private Function WithScreenTip(screenTip As String) As ILabelControlBuilder Implements ILabelControlBuilder.WithScreenTip
            ScreentipBase(screenTip)
            Return Me
        End Function

        Private Function WithScreenTip(screenTip As String, callback As FromControl(Of String)) As ILabelControlBuilder Implements ILabelControlBuilder.WithScreenTip
            ScreentipBase(screenTip, callback)
            Return Me
        End Function

        Private Function WithSuperTip(superTip As String) As ILabelControlBuilder Implements ILabelControlBuilder.WithSuperTip
            SupertipBase(superTip)
            Return Me
        End Function

        Private Function WithSuperTip(superTip As String, callback As FromControl(Of String)) As ILabelControlBuilder Implements ILabelControlBuilder.WithSuperTip
            SupertipBase(superTip, callback)
            Return Me
        End Function

        Private Function ShowLabel() As ILabelControlBuilder Implements ILabelControlBuilder.ShowLabel
            ShowLabelBase()
            Return Me
        End Function

        Private Function ShowLabel(callback As FromControl(Of Boolean)) As ILabelControlBuilder Implements ILabelControlBuilder.ShowLabel
            ShowLabelBase(callback)
            Return Me
        End Function

        Private Function HideLabel() As ILabelControlBuilder Implements ILabelControlBuilder.HideLabel
            HideLabelBase()
            Return Me
        End Function

        Private Function HideLabel(callback As FromControl(Of Boolean)) As ILabelControlBuilder Implements ILabelControlBuilder.HideLabel
            HideLabelBase(callback)
            Return Me
        End Function

        Private Function InsertBeforeMSO(builtInControl As MSO) As ILabelControlBuilder Implements ILabelControlBuilder.InsertBefore
            InsertBeforeMsoBase(builtInControl)
            Return Me
        End Function

        Private Function InsertBeforeQ(qualifiedControl As RibbonElement) As ILabelControlBuilder Implements ILabelControlBuilder.InsertBefore
            InsertBeforeQBase(qualifiedControl)
            Return Me
        End Function

        Private Function InsertAfterMSO(builtInControl As MSO) As ILabelControlBuilder Implements ILabelControlBuilder.InsertAfter
            InsertAfterMsoBase(builtInControl)
            Return Me
        End Function

        Private Function InsertAfterQ(qualifiedControl As RibbonElement) As ILabelControlBuilder Implements ILabelControlBuilder.InsertAfter
            InsertAfterQBase(qualifiedControl)
            Return Me
        End Function

    End Class

End NameSpace