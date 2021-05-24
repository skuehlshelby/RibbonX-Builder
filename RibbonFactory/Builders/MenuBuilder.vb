Imports System.Drawing
Imports RibbonFactory.BuilderInterfaces
Imports RibbonFactory.Containers
Imports RibbonFactory.Enums
Imports RibbonFactory.Enums.ImageMSO
Imports RibbonFactory.Enums.MSO
Imports stdole

Namespace Builders
    Public Class MenuBuilder
        Inherits Builder(Of Menu)
        Implements ISetInsertionPoint(Of MenuBuilder)
        Implements ISetEnabled(Of MenuBuilder)
        Implements ISetVisibility(Of MenuBuilder)
        Implements ISetLabelScreenTipAndSuperTip(Of MenuBuilder)
        Implements ISetLabelVisibility(Of MenuBuilder)
        Implements ISetImage(Of MenuBuilder)
        Implements ISetDescription(Of MenuBuilder)
        Implements ISetSize(Of MenuBuilder)
        Implements ISetKeyTip(Of MenuBuilder)

        'TODO Create Interface to set size of child menu items

        Public Overrides Function Build(tag As Object) As Menu
            Return New Menu(attributes:=Attributes, tag:=tag)
        End Function

        Public Function Enabled() As MenuBuilder Implements ISetEnabled(Of MenuBuilder).Enabled
            AddEnabled(enabled:=True)
            Return Me
        End Function

        Public Function Enabled(callback As FromControl(Of Boolean)) As MenuBuilder Implements ISetEnabled(Of MenuBuilder).Enabled
            AddEnabled(enabled:=True, callback:=callback)
            Return Me
        End Function

        Public Function Disabled() As MenuBuilder Implements ISetEnabled(Of MenuBuilder).Disabled
            AddEnabled(enabled:=False)
            Return Me
        End Function

        Public Function Disabled(callback As FromControl(Of Boolean)) As MenuBuilder Implements ISetEnabled(Of MenuBuilder).Disabled
            AddEnabled(enabled:=False, callback:=callback)
            Return Me
        End Function

        Public Function Visible() As MenuBuilder Implements ISetVisibility(Of MenuBuilder).Visible
            AddVisible(visible:=True)
            Return Me
        End Function

        Public Function Visible(callback As FromControl(Of Boolean)) As MenuBuilder Implements ISetVisibility(Of MenuBuilder).Visible
            AddVisible(visible:=True, callback:=callback)
            Return Me
        End Function

        Public Function Invisible() As MenuBuilder Implements ISetVisibility(Of MenuBuilder).Invisible
            AddVisible(visible:=False)
            Return Me
        End Function

        Public Function Invisible(callback As FromControl(Of Boolean)) As MenuBuilder Implements ISetVisibility(Of MenuBuilder).Invisible
            AddVisible(visible:=False, callback:=callback)
            Return Me
        End Function

        Public Function WithLabel(label As String, Optional copyToScreenTip As Boolean = True) As MenuBuilder Implements ISetLabelScreenTipAndSuperTip(Of MenuBuilder).WithLabel
            AddLabel(label:=label)
            AddShowLabel(showLabel:=True)

            If copyToScreenTip Then
                AddScreenTip(screenTip:=label)
            End If

            Return Me
        End Function

        Public Function WithLabel(label As String, callback As FromControl(Of String), Optional copyToScreenTip As Boolean = True) As MenuBuilder Implements ISetLabelScreenTipAndSuperTip(Of MenuBuilder).WithLabel
            AddLabel(label:=label, callback:=callback)
            AddShowLabel(showLabel:=True)

            If copyToScreenTip Then
                AddScreenTip(screenTip:=label, callback:=callback)
            End If

            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String) As MenuBuilder Implements ISetLabelScreenTipAndSuperTip(Of MenuBuilder).WithScreenTip
            AddScreenTip(screenTip:=screenTip)
            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String, callback As FromControl(Of String)) As MenuBuilder Implements ISetLabelScreenTipAndSuperTip(Of MenuBuilder).WithScreenTip
            AddScreenTip(screenTip:=screenTip, callback:=callback)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String) As MenuBuilder Implements ISetLabelScreenTipAndSuperTip(Of MenuBuilder).WithSuperTip
            AddSuperTip(superTip:=superTip)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String, callback As FromControl(Of String)) As MenuBuilder Implements ISetLabelScreenTipAndSuperTip(Of MenuBuilder).WithSuperTip
            AddSuperTip(superTip:=superTip, callback:=callback)
            Return Me
        End Function

        Public Function ShowLabel() As MenuBuilder Implements ISetLabelVisibility(Of MenuBuilder).ShowLabel
            AddShowLabel(showLabel:=True)
            Return Me
        End Function

        Public Function ShowLabel(callback As FromControl(Of Boolean)) As MenuBuilder Implements ISetLabelVisibility(Of MenuBuilder).ShowLabel
            AddShowLabel(showLabel:=True, callback:=callback)
            Return Me
        End Function

        Public Function HideLabel() As MenuBuilder Implements ISetLabelVisibility(Of MenuBuilder).HideLabel
            AddShowLabel(showLabel:=False)
            Return Me
        End Function

        Public Function HideLabel(callback As FromControl(Of Boolean)) As MenuBuilder Implements ISetLabelVisibility(Of MenuBuilder).HideLabel
            AddShowLabel(showLabel:=False, callback:=callback)
            Return Me
        End Function

        Public Function WithImage(image As ImageMSO) As MenuBuilder Implements ISetImage(Of MenuBuilder).WithImage
            AddImage(image)
            Return Me
        End Function

        Public Function WithImage(image As IPictureDisp, callback As FromControl(Of IPictureDisp)) As MenuBuilder Implements ISetImage(Of MenuBuilder).WithImage
            AddImage(image, callback:=callback)
            Return Me
        End Function

        Public Function WithImage(image As Bitmap, callback As FromControl(Of IPictureDisp)) As MenuBuilder Implements ISetImage(Of MenuBuilder).WithImage
            AddImage(image, callback:=callback)
            Return Me
        End Function

        Public Function WithDescription(description As String) As MenuBuilder Implements ISetDescription(Of MenuBuilder).WithDescription
            AddDescription(description:=description)
            Return Me
        End Function

        Public Function WithDescription(description As String, callback As FromControl(Of String)) As MenuBuilder Implements ISetDescription(Of MenuBuilder).WithDescription
            AddDescription(description:=description, callback:=callback)
            Return Me
        End Function

        Public Function Large() As MenuBuilder Implements ISetSize(Of MenuBuilder).Large
            AddLarge()
            Return Me
        End Function

        Public Function Large(callback As FromControl(Of ControlSize)) As MenuBuilder Implements ISetSize(Of MenuBuilder).Large
            AddLarge(callback)
            Return Me
        End Function

        Public Function Normal() As MenuBuilder Implements ISetSize(Of MenuBuilder).Normal
            AddNormal()
            Return Me
        End Function

        Public Function Normal(callback As FromControl(Of ControlSize)) As MenuBuilder Implements ISetSize(Of MenuBuilder).Normal
            AddNormal(callback)
            Return Me
        End Function

        Public Function WithKeyTip(keyTip As KeyTip) As MenuBuilder Implements ISetKeyTip(Of MenuBuilder).WithKeyTip
            AddKeyTip(keyTip)
            Return Me
        End Function

        Public Function WithKeyTip(keyTip As KeyTip, callback As FromControl(Of KeyTip)) As MenuBuilder Implements ISetKeyTip(Of MenuBuilder).WithKeyTip
            AddKeyTip(keyTip, callback:=callback)
            Return Me
        End Function

        Public Function InsertBeforeMSO(builtInControl As MSO) As MenuBuilder Implements ISetInsertionPoint(Of MenuBuilder).InsertBefore
            InsertBefore(builtInControl)
            Return Me
        End Function

        Public Function InsertBeforeQ(qualifiedControl As RibbonElement) As MenuBuilder Implements ISetInsertionPoint(Of MenuBuilder).InsertBefore
            InsertBefore(qualifiedControl)
            Return Me
        End Function

        Public Function InsertAfterMSO(builtInControl As MSO) As MenuBuilder Implements ISetInsertionPoint(Of MenuBuilder).InsertAfter
            InsertAfter(builtInControl)
            Return Me
        End Function

        Public Function InsertAfterQ(qualifiedControl As RibbonElement) As MenuBuilder Implements ISetInsertionPoint(Of MenuBuilder).InsertAfter
            InsertAfter(qualifiedControl)
            Return Me
        End Function
    End Class
End Namespace