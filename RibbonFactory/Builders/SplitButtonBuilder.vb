Imports RibbonFactory.BuilderInterfaces
Imports RibbonFactory.Containers
Imports RibbonFactory.Controls
Imports RibbonFactory.Enums
Imports RibbonFactory.Enums.MSO

Namespace Builders
    Public Class SplitButtonBuilder
        Inherits Builder(Of SplitButton)
        Implements ISetInsertionPoint(Of SplitButtonBuilder)
        Implements ISetVisibility(Of SplitButtonBuilder)
        Implements ISetEnabled(Of SplitButtonBuilder)
        Implements ISetKeyTip(Of SplitButtonBuilder)
        Implements ISetLabelVisibility(Of SplitButtonBuilder)
        Implements ISetSize(Of SplitButtonBuilder)

        Private ReadOnly _button As Button
        Private ReadOnly _toggleButton As ToggleButton
        Private ReadOnly _menu As Menu

        Public Sub New(button As Button, menu As Menu)
            _button = NullGuard.NotNull(button, argName:=NameOf(button))
            _menu = NullGuard.NotNull(menu, argName:=NameOf(menu))
        End Sub

        Public Sub New(button As ToggleButton, menu As Menu)
            _toggleButton = NullGuard.NotNull(button, argName:=NameOf(button))
            _menu = NullGuard.NotNull(menu, argName:=NameOf(menu))
        End Sub

        Public Overrides Function Build(tag As Object) As SplitButton
            If _button IsNot Nothing Then
                Return New SplitButton(button:=_button, menu:=_menu, attributes:=Attributes, tag:=tag)
            Else
                Return New SplitButton(button:=_toggleButton, menu:=_menu, attributes:=Attributes, tag:=tag)
            End If
        End Function

        Public Function Visible() As SplitButtonBuilder Implements ISetVisibility(Of SplitButtonBuilder).Visible
            AddVisible(visible:=True)
            Return Me
        End Function

        Public Function Visible(callback As FromControl(Of Boolean)) As SplitButtonBuilder Implements ISetVisibility(Of SplitButtonBuilder).Visible
            AddVisible(visible:=True, callback:=callback)
            Return Me
        End Function

        Public Function Invisible() As SplitButtonBuilder Implements ISetVisibility(Of SplitButtonBuilder).Invisible
            AddVisible(visible:=False)
            Return Me
        End Function

        Public Function Invisible(callback As FromControl(Of Boolean)) As SplitButtonBuilder Implements ISetVisibility(Of SplitButtonBuilder).Invisible
            AddVisible(visible:=False, callback:=callback)
            Return Me
        End Function

        Public Function Enabled() As SplitButtonBuilder Implements ISetEnabled(Of SplitButtonBuilder).Enabled
            AddEnabled(enabled:=True)
            Return Me
        End Function

        Public Function Enabled(callback As FromControl(Of Boolean)) As SplitButtonBuilder Implements ISetEnabled(Of SplitButtonBuilder).Enabled
            AddEnabled(enabled:=True, callback:=callback)
            Return Me
        End Function

        Public Function Disabled() As SplitButtonBuilder Implements ISetEnabled(Of SplitButtonBuilder).Disabled
            AddEnabled(enabled:=False)
            Return Me
        End Function

        Public Function Disabled(callback As FromControl(Of Boolean)) As SplitButtonBuilder Implements ISetEnabled(Of SplitButtonBuilder).Disabled
            AddEnabled(enabled:=False, callback:=callback)
            Return Me
        End Function

        Public Function WithKeyTip(keyTip As KeyTip) As SplitButtonBuilder Implements ISetKeyTip(Of SplitButtonBuilder).WithKeyTip
            AddKeyTip(keyTip)
            Return Me
        End Function

        Public Function WithKeyTip(keyTip As KeyTip, callback As FromControl(Of KeyTip)) As SplitButtonBuilder Implements ISetKeyTip(Of SplitButtonBuilder).WithKeyTip
            AddKeyTip(keyTip, callback:=callback)
            Return Me
        End Function

        Public Function ShowLabel() As SplitButtonBuilder Implements ISetLabelVisibility(Of SplitButtonBuilder).ShowLabel
            AddShowLabel(showLabel:=True)
            Return Me
        End Function

        Public Function ShowLabel(callback As FromControl(Of Boolean)) As SplitButtonBuilder Implements ISetLabelVisibility(Of SplitButtonBuilder).ShowLabel
            AddShowLabel(showLabel:=True, callback:=callback)
            Return Me
        End Function

        Public Function HideLabel() As SplitButtonBuilder Implements ISetLabelVisibility(Of SplitButtonBuilder).HideLabel
            AddShowLabel(showLabel:=False)
            Return Me
        End Function

        Public Function HideLabel(callback As FromControl(Of Boolean)) As SplitButtonBuilder Implements ISetLabelVisibility(Of SplitButtonBuilder).HideLabel
            AddShowLabel(showLabel:=False, callback:=callback)
            Return Me
        End Function

        Public Function Large() As SplitButtonBuilder Implements ISetSize(Of SplitButtonBuilder).Large
            AddLarge()
            Return Me
        End Function

        Public Function Large(callback As FromControl(Of ControlSize)) As SplitButtonBuilder Implements ISetSize(Of SplitButtonBuilder).Large
            AddLarge(callback)
            Return Me
        End Function

        Public Function Normal() As SplitButtonBuilder Implements ISetSize(Of SplitButtonBuilder).Normal
            AddNormal()
            Return Me
        End Function

        Public Function Normal(callback As FromControl(Of ControlSize)) As SplitButtonBuilder Implements ISetSize(Of SplitButtonBuilder).Normal
            AddNormal(callback)
            Return Me
        End Function

        Public Function InsertBeforeMSO(builtInControl As MSO) As SplitButtonBuilder Implements ISetInsertionPoint(Of SplitButtonBuilder).InsertBefore
            InsertBefore(builtInControl)
            Return Me
        End Function

        Public Function InsertBeforeQ(qualifiedControl As RibbonElement) As SplitButtonBuilder Implements ISetInsertionPoint(Of SplitButtonBuilder).InsertBefore
            InsertBefore(qualifiedControl)
            Return Me
        End Function

        Public Function InsertAfterMSO(builtInControl As MSO) As SplitButtonBuilder Implements ISetInsertionPoint(Of SplitButtonBuilder).InsertAfter
            InsertAfter(builtInControl)
            Return Me
        End Function

        Public Function InsertAfterQ(qualifiedControl As RibbonElement) As SplitButtonBuilder Implements ISetInsertionPoint(Of SplitButtonBuilder).InsertAfter
            InsertAfter(qualifiedControl)
            Return Me
        End Function
    End Class
End Namespace