Imports RibbonFactory.BuilderInterfaces
Imports RibbonFactory.Controls
Imports RibbonFactory.Enums.MSO

Namespace Builders
    Public Class LabelControlBuilder
        Inherits Builder(Of LabelControl)
        Implements ISetEnabled(Of LabelControlBuilder)
        Implements ISetInsertionPoint(Of LabelControlBuilder)
        Implements ISetVisibility(Of LabelControlBuilder)
        Implements ISetLabelScreenTipAndSuperTip(Of LabelControlBuilder)
        Implements ISetLabelVisibility(Of LabelControlBuilder)
        
        Public Overrides Function Build() As LabelControl
            Return Build(Nothing)
        End Function

        Public Overrides Function Build(tag As Object) As LabelControl
            Return New LabelControl(Attributes, tag)
        End Function

        Public Function Enabled() As LabelControlBuilder Implements ISetEnabled(Of LabelControlBuilder).Enabled
            AddEnabled(enabled:= True)
            Return Me
        End Function

        Public Function Enabled(callback As FromControl(Of Boolean)) As LabelControlBuilder Implements ISetEnabled(Of LabelControlBuilder).Enabled
            AddEnabled(enabled:= True, callback:= callback)
            Return Me
        End Function

        Public Function Disabled() As LabelControlBuilder Implements ISetEnabled(Of LabelControlBuilder).Disabled
            AddEnabled(enabled:= False)
            Return Me
        End Function

        Public Function Disabled(callback As FromControl(Of Boolean)) As LabelControlBuilder Implements ISetEnabled(Of LabelControlBuilder).Disabled
            AddEnabled(enabled:= False, callback:= callback)
            Return Me
        End Function

        Public Function Visible() As LabelControlBuilder Implements ISetVisibility(Of LabelControlBuilder).Visible
            AddVisible(visible:= True)
            Return Me
        End Function

        Public Function Visible(callback As FromControl(Of Boolean)) As LabelControlBuilder Implements ISetVisibility(Of LabelControlBuilder).Visible
            AddVisible(visible:= True, callback := callback)
            Return Me
        End Function

        Public Function Invisible() As LabelControlBuilder Implements ISetVisibility(Of LabelControlBuilder).Invisible
            AddVisible(visible:= True)
            Return Me
        End Function

        Public Function Invisible(callback As FromControl(Of Boolean)) As LabelControlBuilder Implements ISetVisibility(Of LabelControlBuilder).Invisible
            AddVisible(visible:= True, callback := callback)
            Return Me
        End Function

        Public Function WithLabel(label As String, Optional copyToScreenTip As Boolean = True) As LabelControlBuilder Implements ISetLabelScreenTipAndSuperTip(Of LabelControlBuilder).WithLabel
            AddLabel(label:= label)
            AddShowLabel(showLabel:= True)
            
            If copyToScreenTip Then
                AddScreenTip(screenTip:= label)
            End If

            Return Me
        End Function

        Public Function WithLabel(label As String, callback As FromControl(Of String), Optional copyToScreenTip As Boolean = True) As LabelControlBuilder Implements ISetLabelScreenTipAndSuperTip(Of LabelControlBuilder).WithLabel
            AddLabel(label:= label, callback:= callback)
            AddShowLabel(showLabel:= True)
            
            If copyToScreenTip Then
                AddScreenTip(screenTip:= label, callback:= callback)
            End If

            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String) As LabelControlBuilder Implements ISetLabelScreenTipAndSuperTip(Of LabelControlBuilder).WithScreenTip
            AddScreenTip(screenTip:= screenTip)
            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String, callback As FromControl(Of String)) As LabelControlBuilder Implements ISetLabelScreenTipAndSuperTip(Of LabelControlBuilder).WithScreenTip
            AddScreenTip(screenTip:= screenTip, callback:= callback)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String) As LabelControlBuilder Implements ISetLabelScreenTipAndSuperTip(Of LabelControlBuilder).WithSuperTip
            AddSuperTip(superTip:= superTip)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String, callback As FromControl(Of String)) As LabelControlBuilder Implements ISetLabelScreenTipAndSuperTip(Of LabelControlBuilder).WithSuperTip
            AddSuperTip(superTip:= superTip, callback:= callback)
            Return Me
        End Function

        Public Function ShowLabel() As LabelControlBuilder Implements ISetLabelVisibility(Of LabelControlBuilder).ShowLabel
            AddShowLabel(showLabel:= True)
            Return Me
        End Function

        Public Function ShowLabel(callback As FromControl(Of Boolean)) As LabelControlBuilder Implements ISetLabelVisibility(Of LabelControlBuilder).ShowLabel
            AddShowLabel(showLabel:= True, callback:= callback)
            Return Me
        End Function

        Public Function HideLabel() As LabelControlBuilder Implements ISetLabelVisibility(Of LabelControlBuilder).HideLabel
            AddShowLabel(showLabel:= False)
            Return Me
        End Function

        Public Function HideLabel(callback As FromControl(Of Boolean)) As LabelControlBuilder Implements ISetLabelVisibility(Of LabelControlBuilder).HideLabel
            AddShowLabel(showLabel:= False, callback:= callback)
            Return Me
        End Function

        Private Function InsertBeforeMSO(builtInControl As MSO) As LabelControlBuilder Implements ISetInsertionPoint(Of LabelControlBuilder).InsertBefore
            InsertBefore(builtInControl)
            Return Me
        End Function

        Private Function InsertBeforeQ(qualifiedControl As RibbonElement) As LabelControlBuilder Implements ISetInsertionPoint(Of LabelControlBuilder).InsertBefore
            InsertBefore(qualifiedControl)
            Return Me
        End Function

        Private Function InsertAfterMSO(builtInControl As MSO) As LabelControlBuilder Implements ISetInsertionPoint(Of LabelControlBuilder).InsertAfter
            InsertAfter(builtInControl)
            Return Me
        End Function

        Private Function InsertAfterQ(qualifiedControl As RibbonElement) As LabelControlBuilder Implements ISetInsertionPoint(Of LabelControlBuilder).InsertAfter
            InsertAfter(qualifiedControl)
            Return Me
        End Function
    End Class
End NameSpace