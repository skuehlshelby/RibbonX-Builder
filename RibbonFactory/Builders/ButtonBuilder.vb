Imports RibbonFactory.Builder_Interfaces
Imports RibbonFactory.Controls
Imports RibbonFactory.Enums

Namespace Builders
    Public NotInheritable Class ButtonBuilder
        Inherits Builder(Of Button)
        Implements ISetEnabled(Of ButtonBuilder)
        Implements ISetVisibility(Of ButtonBuilder)
        Implements ISetLabelScreenTipAndSuperTip(Of ButtonBuilder)
        Implements ISetLabelVisibility(Of ButtonBuilder)
        Implements ISetAction(Of ButtonBuilder)
        Implements ISetSize(Of ButtonBuilder)

        Public Overrides Function Build() As Button
            Return Build(Nothing)
        End Function

        Public Overrides Function Build(tag As Object) As Button
            Return New Button(Attributes, tag)
        End Function

        Public Function Enabled() As ButtonBuilder Implements ISetEnabled(Of ButtonBuilder).Enabled
            Throw New NotImplementedException()
        End Function

        Public Function Disabled() As ButtonBuilder Implements ISetEnabled(Of ButtonBuilder).Disabled
            Throw New NotImplementedException()
        End Function

        Public Function Visible() As ButtonBuilder Implements ISetVisibility(Of ButtonBuilder).Visible
            Throw New NotImplementedException()
        End Function

        Public Function Visible(callback As FromControl(Of Boolean)) As ButtonBuilder Implements ISetVisibility(Of ButtonBuilder).Visible
            Throw New NotImplementedException()
        End Function

        Public Function Invisible() As ButtonBuilder Implements ISetVisibility(Of ButtonBuilder).Invisible
            Throw New NotImplementedException()
        End Function

        Public Function Invisible(callback As FromControl(Of Boolean)) As ButtonBuilder Implements ISetVisibility(Of ButtonBuilder).Invisible
            Throw New NotImplementedException()
        End Function

        Public Function WithLabel(label As String, Optional copyToScreenTip As Boolean = True) As ButtonBuilder Implements ISetLabelScreenTipAndSuperTip(Of ButtonBuilder).WithLabel
            Throw New NotImplementedException()
        End Function

        Public Function WithLabel(label As String, callback As FromControl(Of String), Optional copyToScreenTip As Boolean = True) As ButtonBuilder Implements ISetLabelScreenTipAndSuperTip(Of ButtonBuilder).WithLabel
            Throw New NotImplementedException()
        End Function

        Public Function WithScreenTip(screenTip As String) As ButtonBuilder Implements ISetLabelScreenTipAndSuperTip(Of ButtonBuilder).WithScreenTip
            Throw New NotImplementedException()
        End Function

        Public Function WithScreenTip(screenTip As String, callback As FromControl(Of String)) As ButtonBuilder Implements ISetLabelScreenTipAndSuperTip(Of ButtonBuilder).WithScreenTip
            Throw New NotImplementedException()
        End Function

        Public Function WithSuperTip(superTip As String) As ButtonBuilder Implements ISetLabelScreenTipAndSuperTip(Of ButtonBuilder).WithSuperTip
            Throw New NotImplementedException()
        End Function

        Public Function WithSuperTip(superTip As String, callback As FromControl(Of String)) As ButtonBuilder Implements ISetLabelScreenTipAndSuperTip(Of ButtonBuilder).WithSuperTip
            Throw New NotImplementedException()
        End Function

        Public Function ShowLabel() As ButtonBuilder Implements ISetLabelVisibility(Of ButtonBuilder).ShowLabel
            Throw New NotImplementedException()
        End Function

        Public Function HideLabel() As ButtonBuilder Implements ISetLabelVisibility(Of ButtonBuilder).HideLabel
            Throw New NotImplementedException()
        End Function

        Public Function ThatDoes(callback As OnAction, action As Action) As ButtonBuilder Implements ISetAction(Of ButtonBuilder).ThatDoes
            Throw New NotImplementedException()
        End Function

        Public Function Large() As ButtonBuilder Implements ISetSize(Of ButtonBuilder).Large
            Throw New NotImplementedException()
        End Function

        Public Function Large(callback As FromControl(Of ControlSize)) As ButtonBuilder Implements ISetSize(Of ButtonBuilder).Large
            Throw New NotImplementedException()
        End Function

        Public Function Normal() As ButtonBuilder Implements ISetSize(Of ButtonBuilder).Normal
            Throw New NotImplementedException()
        End Function

        Public Function Normal(callback As FromControl(Of ControlSize)) As ButtonBuilder Implements ISetSize(Of ButtonBuilder).Normal
            Throw New NotImplementedException()
        End Function
    End Class
End NameSpace