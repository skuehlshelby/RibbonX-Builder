Imports RibbonFactory.Builder_Interfaces
Imports RibbonFactory.Controls
Imports RibbonFactory.Enums
Imports RibbonFactory.RibbonAttributes
Imports RibbonFactory.RibbonAttributes.Categories.Label
Imports RibbonFactory.RibbonAttributes.Categories.Screentip
Imports RibbonFactory.RibbonAttributes.Categories.Supertip

Namespace Builders
    Public NotInheritable Class ButtonBuilder
        Inherits Builder(Of Button)
        Implements IDescribable(Of ButtonBuilder)

        Public Overrides Function Build() As Button
            Return Build(Nothing)
        End Function

        Public Overrides Function Build(tag As Object) As Button
            Return New Button(Attributes, tag)
        End Function

        Public Function WithLabel(label As String, Optional copyToScreenTip As Boolean = True) As ButtonBuilder Implements IDescribable(Of ButtonBuilder).WithLabel
            AddLabel(label)
            AddShowLabel(True)

            If copyToScreenTip Then
                AddScreenTip(label)
            End If

            Return Me
        End Function

        Public Function WithLabel(label As String, callback As FromControl(Of String), Optional copyToScreenTip As Boolean = True) As ButtonBuilder Implements IDescribable(Of ButtonBuilder).WithLabel
            AddLabel(label, callback)
            AddShowLabel(True)

            If copyToScreenTip Then
                AddScreenTip(label, callback)
            End If

            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String) As ButtonBuilder Implements IDescribable(Of ButtonBuilder).WithScreenTip
            AddScreenTip(screenTip)
            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String, callback As FromControl(Of String)) As ButtonBuilder Implements IDescribable(Of ButtonBuilder).WithScreenTip
            AddScreenTip(screenTip, callback)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String) As ButtonBuilder Implements IDescribable(Of ButtonBuilder).WithSupertip
            AddSuperTip(superTip)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String, callback As FromControl(Of String)) As ButtonBuilder Implements IDescribable(Of ButtonBuilder).WithSupertip
            AddSuperTip(superTip, callback)
            Return Me
        End Function

        Public Function WithVisibility(visible As Boolean) As ButtonBuilder
            AddVisible(visible)
            Return Me
        End Function

        Public Function WithVisibility(visible As Boolean, callback As FromControl(Of Boolean)) As ButtonBuilder
            AddVisible(visible, callback)
            Return Me
        End Function

        Public Function WithImage(image As ImageMSO) As ButtonBuilder
            AddImage(image)
            Return Me
        End Function

        Public Function ThatDoes(callback As OnAction, action As Action) As ButtonBuilder
            AddAction(callback, action)
            Return Me
        End Function
    End Class
End NameSpace