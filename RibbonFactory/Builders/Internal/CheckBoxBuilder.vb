Imports RibbonX.Controls
Imports RibbonX.Controls.Base
Imports RibbonX.Controls.EventArgs
Imports RibbonX.Controls.BuiltIn
Imports RibbonX.Builders.Internal.Base
Imports RibbonX.Callbacks
Imports RibbonX.SimpleTypes

Namespace Builders

    Friend Class CheckBoxBuilder
        Inherits BuilderBase(Of CheckBox)
        Implements ICheckBoxBuilder

        Public Sub New()
            MyBase.New()
        End Sub

        Public Sub New(template As RibbonElement)
            MyBase.New(template)
        End Sub

        Private Function Enabled() As ICheckBoxBuilder Implements ICheckBoxBuilder.Enabled
            EnabledBase()
            Return Me
        End Function

        Private Function Enabled(callback As FromControl(Of Boolean)) As ICheckBoxBuilder Implements ICheckBoxBuilder.Enabled
            EnabledBase(callback)
            Return Me
        End Function

        Private Function Disabled() As ICheckBoxBuilder Implements ICheckBoxBuilder.Disabled
            DisabledBase()
            Return Me
        End Function

        Private Function Disabled(callback As FromControl(Of Boolean)) As ICheckBoxBuilder Implements ICheckBoxBuilder.Disabled
            DisabledBase(callback)
            Return Me
        End Function

        Private Function Visible() As ICheckBoxBuilder Implements ICheckBoxBuilder.Visible
            VisibleBase()
            Return Me
        End Function

        Private Function Visible(callback As FromControl(Of Boolean)) As ICheckBoxBuilder Implements ICheckBoxBuilder.Visible
            VisibleBase(callback)
            Return Me
        End Function

        Private Function Invisible() As ICheckBoxBuilder Implements ICheckBoxBuilder.Invisible
            InvisibleBase()
            Return Me
        End Function

        Private Function Invisible(callback As FromControl(Of Boolean)) As ICheckBoxBuilder Implements ICheckBoxBuilder.Invisible
            InvisibleBase(callback)
            Return Me
        End Function

        Private Function WithLabel(label As String, Optional copyToScreenTip As Boolean = True) As ICheckBoxBuilder Implements ICheckBoxBuilder.WithLabel
            LabelBase(label)

            If copyToScreenTip Then ScreentipBase(label)
            Return Me
        End Function

        Private Function WithLabel(label As String, callback As FromControl(Of String), Optional copyToScreenTip As Boolean = True) As ICheckBoxBuilder Implements ICheckBoxBuilder.WithLabel
            LabelBase(label, callback)

            If copyToScreenTip Then ScreentipBase(label, callback)
            Return Me
        End Function

        Private Function WithScreenTip(screenTip As String) As ICheckBoxBuilder Implements ICheckBoxBuilder.WithScreenTip
            ScreentipBase(screenTip)
            Return Me
        End Function

        Private Function WithScreenTip(screenTip As String, callback As FromControl(Of String)) As ICheckBoxBuilder Implements ICheckBoxBuilder.WithScreenTip
            ScreentipBase(screenTip, callback)
            Return Me
        End Function

        Private Function WithSuperTip(superTip As String) As ICheckBoxBuilder Implements ICheckBoxBuilder.WithSuperTip
            SupertipBase(superTip)
            Return Me
        End Function

        Private Function WithSuperTip(superTip As String, callback As FromControl(Of String)) As ICheckBoxBuilder Implements ICheckBoxBuilder.WithSuperTip
            SupertipBase(superTip, callback)
            Return Me
        End Function

        Private Function WithDescription(description As String) As ICheckBoxBuilder Implements ICheckBoxBuilder.WithDescription
            DescriptionBase(description)
            Return Me
        End Function

        Private Function WithDescription(description As String, callback As FromControl(Of String)) As ICheckBoxBuilder Implements ICheckBoxBuilder.WithDescription
            DescriptionBase(description, callback)
            Return Me
        End Function

        Private Function WithKeyTip(keyTip As KeyTip) As ICheckBoxBuilder Implements ICheckBoxBuilder.WithKeyTip
            KeyTipBase(keyTip)
            Return Me
        End Function

        Private Function WithKeyTip(keyTip As KeyTip, callback As FromControl(Of KeyTip)) As ICheckBoxBuilder Implements ICheckBoxBuilder.WithKeyTip
            KeyTipBase(keyTip, callback)
            Return Me
        End Function

        Private Function Checked(getChecked As FromControl(Of Boolean), setChecked As ButtonPressed) As ICheckBoxBuilder Implements ICheckBoxBuilder.Checked
            PressedBase(True, getChecked, setChecked)
            Return Me
        End Function

        Private Function Unchecked(getChecked As FromControl(Of Boolean), setChecked As ButtonPressed) As ICheckBoxBuilder Implements ICheckBoxBuilder.Unchecked
            PressedBase(False, getChecked, setChecked)
            Return Me
        End Function

        Private Function InsertBeforeMSO(builtInControl As MSO) As ICheckBoxBuilder Implements ICheckBoxBuilder.InsertBefore
            InsertBeforeMsoBase(builtInControl)
            Return Me
        End Function

        Private Function InsertBeforeQ(qualifiedControl As RibbonElement) As ICheckBoxBuilder Implements ICheckBoxBuilder.InsertBefore
            InsertBeforeQBase(qualifiedControl)
            Return Me
        End Function

        Private Function InsertAfterMSO(builtInControl As MSO) As ICheckBoxBuilder Implements ICheckBoxBuilder.InsertAfter
            InsertAfterMsoBase(builtInControl)
            Return Me
        End Function

        Private Function InsertAfterQ(qualifiedControl As RibbonElement) As ICheckBoxBuilder Implements ICheckBoxBuilder.InsertAfter
            InsertAfterQBase(qualifiedControl)
            Return Me
        End Function

        Private Function WithId(id As String) As ICheckBoxBuilder Implements ICheckBoxBuilder.WithId
            WithIdBase(id)
            Return Me
        End Function

        Private Function WithIdQ([namespace] As String, id As String) As ICheckBoxBuilder Implements ICheckBoxBuilder.WithIdQ
            WithIdBase([namespace], id)
            Return Me
        End Function

        Private Function WithIdMso(mso As MSO) As ICheckBoxBuilder Implements ICheckBoxBuilder.WithIdMso
            WithIdBase(mso)
            Return Me
        End Function

        Private Function BeforeStateChange(ParamArray actions() As Action(Of BeforeToggleChangeEventArgs)) As ICheckBoxBuilder Implements ICheckBoxBuilder.BeforeStateChange
            AddBeforeToggleHandlers(actions)
            Return Me
        End Function

        Private Function AfterStateChange(ParamArray actions() As Action(Of ToggleChangeEventArgs)) As ICheckBoxBuilder Implements ICheckBoxBuilder.AfterCheckStateChanged
            AddToggleHandlers(actions)
            Return Me
        End Function

    End Class

End Namespace