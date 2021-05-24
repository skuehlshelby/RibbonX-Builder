Imports System.Drawing
Imports RibbonFactory.BuilderInterfaces
Imports RibbonFactory.Controls
Imports RibbonFactory.Enums
Imports RibbonFactory.Enums.ImageMSO
Imports RibbonFactory.Enums.MSO
Imports stdole

Namespace Builders
    
    Public Class ToggleButtonBuilder
        Inherits Builder(Of ToggleButton)
        Implements ISetVisibility(Of ToggleButtonBuilder)
        Implements ISetEnabled(Of ToggleButtonBuilder)
        Implements ISetInsertionPoint(Of ToggleButtonBuilder)
        Implements ISetLabelScreenTipAndSuperTip(Of ToggleButtonBuilder)
        Implements ISetLabelVisibility(Of ToggleButtonBuilder)
        Implements ISetKeyTip(Of ToggleButtonBuilder)
        Implements ISetImage(Of ToggleButtonBuilder)
        Implements ISetDescription(Of ToggleButtonBuilder)
        Implements ISetSize(Of ToggleButtonBuilder)
        Implements ISetToggleAction(Of ToggleButtonBuilder)

        Public Overrides Function Build(tag As Object) As ToggleButton
            Return New ToggleButton(attributes:= Attributes, tag:= tag)
        End Function

        Public Function Enabled() As ToggleButtonBuilder Implements ISetEnabled(Of ToggleButtonBuilder).Enabled
            AddEnabled(enabled:= True)
            Return Me
        End Function

        Public Function Enabled(callback As FromControl(Of Boolean)) As ToggleButtonBuilder Implements ISetEnabled(Of ToggleButtonBuilder).Enabled
            AddEnabled(enabled:= True, callback:= callback)
            Return Me
        End Function

        Public Function Disabled() As ToggleButtonBuilder Implements ISetEnabled(Of ToggleButtonBuilder).Disabled
            AddEnabled(enabled:= False)
            Return Me
        End Function

        Public Function Disabled(callback As FromControl(Of Boolean)) As ToggleButtonBuilder Implements ISetEnabled(Of ToggleButtonBuilder).Disabled
            AddEnabled(enabled:= False, callback:= callback)
            Return Me
        End Function

        Public Function Visible() As ToggleButtonBuilder Implements ISetVisibility(Of ToggleButtonBuilder).Visible
            AddVisible(visible:= True)
            Return Me
        End Function

        Public Function Visible(callback As FromControl(Of Boolean)) As ToggleButtonBuilder Implements ISetVisibility(Of ToggleButtonBuilder).Visible
            AddVisible(visible:= True, callback:= callback)
            Return Me
        End Function

        Public Function Invisible() As ToggleButtonBuilder Implements ISetVisibility(Of ToggleButtonBuilder).Invisible
            AddVisible(visible:= False)
            Return Me
        End Function

        Public Function Invisible(callback As FromControl(Of Boolean)) As ToggleButtonBuilder Implements ISetVisibility(Of ToggleButtonBuilder).Invisible
            AddVisible(visible:= False, callback:= callback)
            Return Me
        End Function

        Public Function WithLabel(label As String, Optional copyToScreenTip As Boolean = True) As ToggleButtonBuilder Implements ISetLabelScreenTipAndSuperTip(Of ToggleButtonBuilder).WithLabel
            AddLabel(label:= label)
            
            If copyToScreenTip Then
                AddScreenTip(screenTip:= label)
            End If
            
            Return Me
        End Function

        Public Function WithLabel(label As String, callback As FromControl(Of String), Optional copyToScreenTip As Boolean = True) As ToggleButtonBuilder Implements ISetLabelScreenTipAndSuperTip(Of ToggleButtonBuilder).WithLabel
            AddLabel(label:= label, callback:= callback)
            
            If copyToScreenTip Then
                AddScreenTip(screenTip:= label, callback:= callback)
            End If
            
            Return Me
        End Function

        Public Function ShowLabel() As ToggleButtonBuilder Implements ISetLabelVisibility(Of ToggleButtonBuilder).ShowLabel
            AddShowLabel(showLabel:= True)
            Return Me
        End Function

        Public Function ShowLabel(callback As FromControl(Of Boolean)) As ToggleButtonBuilder Implements ISetLabelVisibility(Of ToggleButtonBuilder).ShowLabel
            AddShowLabel(showLabel:= True, callback:= callback)
            Return Me
        End Function

        Public Function HideLabel() As ToggleButtonBuilder Implements ISetLabelVisibility(Of ToggleButtonBuilder).HideLabel
            AddShowLabel(showLabel:= False)
            Return Me
        End Function

        Public Function HideLabel(callback As FromControl(Of Boolean)) As ToggleButtonBuilder Implements ISetLabelVisibility(Of ToggleButtonBuilder).HideLabel
            AddShowLabel(showLabel:= False, callback:= callback)
            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String) As ToggleButtonBuilder Implements ISetLabelScreenTipAndSuperTip(Of ToggleButtonBuilder).WithScreenTip
            AddScreenTip(screenTip:= screenTip)
            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String, callback As FromControl(Of String)) As ToggleButtonBuilder Implements ISetLabelScreenTipAndSuperTip(Of ToggleButtonBuilder).WithScreenTip
            AddScreenTip(screenTip:= screenTip, callback:= callback)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String) As ToggleButtonBuilder Implements ISetLabelScreenTipAndSuperTip(Of ToggleButtonBuilder).WithSuperTip
            AddSuperTip(superTip:= superTip)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String, callback As FromControl(Of String)) As ToggleButtonBuilder Implements ISetLabelScreenTipAndSuperTip(Of ToggleButtonBuilder).WithSuperTip
            AddSuperTip(superTip:= superTip, callback:= callback)
            Return Me
        End Function

        Public Function WithKeyTip(keyTip As KeyTip) As ToggleButtonBuilder Implements ISetKeyTip(Of ToggleButtonBuilder).WithKeyTip
            AddKeyTip(keyTip:= keyTip)
            Return Me
        End Function

        Public Function WithKeyTip(keyTip As KeyTip, callback As FromControl(Of KeyTip)) As ToggleButtonBuilder Implements ISetKeyTip(Of ToggleButtonBuilder).WithKeyTip
            AddKeyTip(keyTip:= keyTip, callback:= callback)
            Return Me
        End Function

        Public Function WithImage(image As ImageMSO) As ToggleButtonBuilder Implements ISetImage(Of ToggleButtonBuilder).WithImage
            AddImage(image)
            Return Me
        End Function

        Public Function WithImage(image As IPictureDisp, callback As FromControl(Of IPictureDisp)) As ToggleButtonBuilder Implements ISetImage(Of ToggleButtonBuilder).WithImage
            AddImage(image, callback:= callback)
            Return Me
        End Function

        Public Function WithImage(image As Bitmap, callback As FromControl(Of IPictureDisp)) As ToggleButtonBuilder Implements ISetImage(Of ToggleButtonBuilder).WithImage
            AddImage(image, callback:= callback)
            Return Me
        End Function

        Public Function WithDescription(description As String) As ToggleButtonBuilder Implements ISetDescription(Of ToggleButtonBuilder).WithDescription
            AddDescription(description:= description)
            Return Me
        End Function

        Public Function WithDescription(description As String, callback As FromControl(Of String)) As ToggleButtonBuilder Implements ISetDescription(Of ToggleButtonBuilder).WithDescription
            AddDescription(description:= description)
            Return Me
        End Function

        Public Function Large() As ToggleButtonBuilder Implements ISetSize(Of ToggleButtonBuilder).Large
            AddLarge()
            Return Me
        End Function

        Public Function Large(callback As FromControl(Of ControlSize)) As ToggleButtonBuilder Implements ISetSize(Of ToggleButtonBuilder).Large
            AddLarge(callback)
            Return Me
        End Function

        Public Function Normal() As ToggleButtonBuilder Implements ISetSize(Of ToggleButtonBuilder).Normal
            AddNormal()
            Return Me
        End Function

        Public Function Normal(callback As FromControl(Of ControlSize)) As ToggleButtonBuilder Implements ISetSize(Of ToggleButtonBuilder).Normal
            AddNormal(callback)
            Return Me
        End Function

        Public Function ThatDoes(callback As ButtonPressed, action As Action) As ToggleButtonBuilder Implements ISetToggleAction(Of ToggleButtonBuilder).ThatDoes
            AddAction(callback:= callback, action:= action)
            Return Me
        End Function

        Public Function InsertBeforeMSO(builtInControl As MSO) As ToggleButtonBuilder Implements ISetInsertionPoint(Of ToggleButtonBuilder).InsertBefore
            InsertBefore(builtInControl)
            Return Me
        End Function

        Public Function InsertBeforeQ(qualifiedControl As RibbonElement) As ToggleButtonBuilder Implements ISetInsertionPoint(Of ToggleButtonBuilder).InsertBefore
            InsertBefore(qualifiedControl)
            Return Me
        End Function

        Public Function InsertAfterMSO(builtInControl As MSO) As ToggleButtonBuilder Implements ISetInsertionPoint(Of ToggleButtonBuilder).InsertAfter
            InsertAfter(builtInControl)
            Return Me
        End Function

        Public Function InsertAfterQ(qualifiedControl As RibbonElement) As ToggleButtonBuilder Implements ISetInsertionPoint(Of ToggleButtonBuilder).InsertAfter
            InsertAfter(qualifiedControl)
            Return Me
        End Function
        
    End Class
    
End NameSpace