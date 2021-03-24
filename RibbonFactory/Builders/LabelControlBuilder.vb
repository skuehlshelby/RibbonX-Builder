Imports RibbonFactory.Builder_Interfaces
Imports RibbonFactory.Controls

Namespace Builders
    Public Class LabelControlBuilder
        Inherits Builder(Of LabelControl)
        Implements ISetLabelScreenTipAndSuperTip(Of LabelControl)
        Implements ISetVisibility(Of LabelControl)

        Public Overrides Function Build() As LabelControl
            Return Build(Nothing)
        End Function

        Public Overrides Function Build(tag As Object) As LabelControl
            Return New LabelControl(Attributes, tag)
        End Function

        Public Function WithLabel(label As String, Optional copyToScreenTip As Boolean = True) As Builder(Of LabelControl) Implements ISetLabelScreenTipAndSuperTip(Of LabelControl).WithLabel
            AddLabel(label)

            If copyToScreenTip Then
                AddScreenTip(label)
            End If

            Return Me
        End Function

        Public Function WithLabel(label As String, callback As FromControl(Of String), Optional copyToScreenTip As Boolean = True) As Builder(Of LabelControl) Implements ISetLabelScreenTipAndSuperTip(Of LabelControl).WithLabel
            AddLabel(label, callback)

            If copyToScreenTip Then
                AddScreenTip(label, callback)
            End If

            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String) As Builder(Of LabelControl) Implements ISetLabelScreenTipAndSuperTip(Of LabelControl).WithScreenTip
            AddScreenTip(screenTip)
            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String, callback As FromControl(Of String)) As Builder(Of LabelControl) Implements ISetLabelScreenTipAndSuperTip(Of LabelControl).WithScreenTip
            AddScreenTip(screenTip, callback)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String) As Builder(Of LabelControl) Implements ISetLabelScreenTipAndSuperTip(Of LabelControl).WithSuperTip
            AddScreenTip(superTip)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String, callback As FromControl(Of String)) As Builder(Of LabelControl) Implements ISetLabelScreenTipAndSuperTip(Of LabelControl).WithSuperTip
            AddScreenTip(superTip, callback)
            Return Me
        End Function

        Public Function Visible() As Builder(Of LabelControl) Implements ISetVisibility(Of LabelControl).Visible
            AddVisible(True)
            Return Me
        End Function

        Public Function Visible(callback As FromControl(Of Boolean)) As Builder(Of LabelControl) Implements ISetVisibility(Of LabelControl).Visible
            AddVisible(True, callback)
            Return Me
        End Function

        Public Function Invisible() As Builder(Of LabelControl) Implements ISetVisibility(Of LabelControl).Invisible
            AddVisible(False)
            Return Me
        End Function

        Public Function Invisible(callback As FromControl(Of Boolean)) As Builder(Of LabelControl) Implements ISetVisibility(Of LabelControl).Invisible
            AddVisible(False, callback)
            Return Me
        End Function
    End Class
End NameSpace