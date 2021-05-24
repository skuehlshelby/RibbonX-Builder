Imports System.Drawing
Imports RibbonFactory.BuilderInterfaces
Imports RibbonFactory.Controls
Imports RibbonFactory.Enums
Imports RibbonFactory.Enums.ImageMSO
Imports RibbonFactory.Enums.MSO
Imports stdole

Namespace Builders
    Public NotInheritable Class ButtonBuilder
        Inherits Builder(Of Button)
        Implements ISetEnabled(Of ButtonBuilder)
        Implements ISetVisibility(Of ButtonBuilder)
        Implements ISetLabelScreenTipAndSuperTip(Of ButtonBuilder)
        Implements ISetLabelVisibility(Of ButtonBuilder)
        Implements ISetClickAction(Of ButtonBuilder)
        Implements ISetSize(Of ButtonBuilder)
        Implements ISetImage(Of ButtonBuilder)
        Implements ISetDescription(Of ButtonBuilder)
        Implements ISetKeyTip(Of ButtonBuilder)
        Implements ISetInsertionPoint(Of ButtonBuilder)

        Public Overrides Function Build(tag As Object) As Button
            Return New Button(Attributes, tag)
        End Function

        Public Function Enabled() As ButtonBuilder Implements ISetEnabled(Of ButtonBuilder).Enabled
            AddEnabled(enabled:=True)
            Return Me
        End Function

        Public Function Enabled(callback As FromControl(Of Boolean)) As ButtonBuilder Implements ISetEnabled(Of ButtonBuilder).Enabled
            AddEnabled(enabled:=True, callback:= callback)
            Return Me
        End Function

        Public Function Disabled() As ButtonBuilder Implements ISetEnabled(Of ButtonBuilder).Disabled
            AddEnabled(enabled:=False)
            Return Me
        End Function

        Public Function Disabled(callback As FromControl(Of Boolean)) As ButtonBuilder Implements ISetEnabled(Of ButtonBuilder).Disabled
            AddEnabled(enabled:=False, callback:= callback)
            Return Me
        End Function

        Public Function Visible() As ButtonBuilder Implements ISetVisibility(Of ButtonBuilder).Visible
            AddVisible(visible:= True)
            Return Me
        End Function

        Public Function Visible(callback As FromControl(Of Boolean)) As ButtonBuilder Implements ISetVisibility(Of ButtonBuilder).Visible
            AddVisible(visible:= True, callback:= callback)
            Return Me
        End Function

        Public Function Invisible() As ButtonBuilder Implements ISetVisibility(Of ButtonBuilder).Invisible
            AddVisible(visible:= False)
            Return Me
        End Function

        Public Function Invisible(callback As FromControl(Of Boolean)) As ButtonBuilder Implements ISetVisibility(Of ButtonBuilder).Invisible
            AddVisible(visible:= False, callback:= callback)
            Return Me
        End Function

        Public Function WithLabel(label As String, Optional copyToScreenTip As Boolean = True) As ButtonBuilder Implements ISetLabelScreenTipAndSuperTip(Of ButtonBuilder).WithLabel
            AddLabel(label:= label)
            AddShowLabel(showLabel:= True)
            
            If copyToScreenTip Then
                AddScreenTip(screenTip:= label)
            End If

            Return Me
        End Function

        Public Function WithLabel(label As String, callback As FromControl(Of String), Optional copyToScreenTip As Boolean = True) As ButtonBuilder Implements ISetLabelScreenTipAndSuperTip(Of ButtonBuilder).WithLabel
            AddLabel(label:= label, callback:= callback)
            AddShowLabel(showLabel:= True)
            
            If copyToScreenTip Then
                AddScreenTip(screenTip:= label, callback:= callback)
            End If

            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String) As ButtonBuilder Implements ISetLabelScreenTipAndSuperTip(Of ButtonBuilder).WithScreenTip
            AddScreenTip(screenTip:= screenTip)
            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String, callback As FromControl(Of String)) As ButtonBuilder Implements ISetLabelScreenTipAndSuperTip(Of ButtonBuilder).WithScreenTip
            AddScreenTip(screenTip:= screenTip, callback:= callback)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String) As ButtonBuilder Implements ISetLabelScreenTipAndSuperTip(Of ButtonBuilder).WithSuperTip
            AddSuperTip(superTip:= superTip)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String, callback As FromControl(Of String)) As ButtonBuilder Implements ISetLabelScreenTipAndSuperTip(Of ButtonBuilder).WithSuperTip
            AddSuperTip(superTip:= superTip, callback:= callback)
            Return Me
        End Function

        Public Function ShowLabel() As ButtonBuilder Implements ISetLabelVisibility(Of ButtonBuilder).ShowLabel
            AddShowLabel(showLabel:= True)
            Return Me
        End Function

        Public Function ShowLabel(callback As FromControl(Of Boolean)) As ButtonBuilder Implements ISetLabelVisibility(Of ButtonBuilder).ShowLabel
            AddShowLabel(showLabel:= True, callback:= callback)
            Return Me
        End Function

        Public Function HideLabel() As ButtonBuilder Implements ISetLabelVisibility(Of ButtonBuilder).HideLabel
            AddShowLabel(showLabel:= False)
            Return Me
        End Function

        Public Function HideLabel(callback As FromControl(Of Boolean)) As ButtonBuilder Implements ISetLabelVisibility(Of ButtonBuilder).HideLabel
            AddShowLabel(showLabel:= False, callback:= callback)
            Return Me
        End Function

        Public Function ThatDoes(callback As OnAction, action As Action) As ButtonBuilder Implements ISetClickAction(Of ButtonBuilder).ThatDoes
            AddAction(callback:= callback, action:= action)
            Return Me
        End Function

        Public Function Large() As ButtonBuilder Implements ISetSize(Of ButtonBuilder).Large
            AddLarge()
            Return Me
        End Function

        Public Function Large(callback As FromControl(Of ControlSize)) As ButtonBuilder Implements ISetSize(Of ButtonBuilder).Large
            AddLarge(callback)
            Return Me
        End Function

        Public Function Normal() As ButtonBuilder Implements ISetSize(Of ButtonBuilder).Normal
            AddNormal()
            Return Me
        End Function

        Public Function Normal(callback As FromControl(Of ControlSize)) As ButtonBuilder Implements ISetSize(Of ButtonBuilder).Normal
            AddNormal(callback)
            Return Me
        End Function

        Public Function WithImage(image As ImageMSO) As ButtonBuilder Implements ISetImage(Of ButtonBuilder).WithImage
            AddImage(image)
            Return Me
        End Function

        Public Function WithImage(image As IPictureDisp, callback As FromControl(Of IPictureDisp)) As ButtonBuilder Implements ISetImage(Of ButtonBuilder).WithImage
            AddImage(image, callback:= callback)
            Return Me
        End Function

        Public Function WithImage(image As Bitmap, callback As FromControl(Of IPictureDisp)) As ButtonBuilder Implements ISetImage(Of ButtonBuilder).WithImage
            AddImage(image, callback:= callback)
            Return Me
        End Function

        Public Function WithDescription(description As String) As ButtonBuilder Implements ISetDescription(Of ButtonBuilder).WithDescription
            AddDescription(description:= description)
            Return Me
        End Function

        Public Function WithDescription(description As String, callback As FromControl(Of String)) As ButtonBuilder Implements ISetDescription(Of ButtonBuilder).WithDescription
            AddDescription(description:= description, callback:= callback)
            Return Me
        End Function

        Public Function WithKeyTip(keyTip As KeyTip) As ButtonBuilder Implements ISetKeyTip(Of ButtonBuilder).WithKeyTip
            AddKeyTip(keyTip)
            Return Me
        End Function

        Public Function WithKeyTip(keyTip As KeyTip, callback As FromControl(Of KeyTip)) As ButtonBuilder Implements ISetKeyTip(Of ButtonBuilder).WithKeyTip
            AddKeyTip(keyTip, callback:= callback)
            Return Me
        End Function

        Public Function InsertBeforeMSO(builtInControl As MSO) As ButtonBuilder Implements ISetInsertionPoint(Of ButtonBuilder).InsertBefore
            InsertBefore(builtInControl)
            Return Me
        End Function

        Public Function InsertBeforeQ(qualifiedControl As RibbonElement) As ButtonBuilder Implements ISetInsertionPoint(Of ButtonBuilder).InsertBefore
            InsertBefore(qualifiedControl)
            Return Me
        End Function

        Public Function InsertAfterMSO(builtInControl As MSO) As ButtonBuilder Implements ISetInsertionPoint(Of ButtonBuilder).InsertAfter
            InsertAfter(builtInControl)
            Return Me
        End Function

        Public Function InsertAfterQ(qualifiedControl As RibbonElement) As ButtonBuilder Implements ISetInsertionPoint(Of ButtonBuilder).InsertAfter
            InsertAfter(qualifiedControl)
            Return Me
        End Function
    End Class
End NameSpace