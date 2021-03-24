Imports RibbonFactory.Controls
Imports RibbonFactory.Enums

Namespace Builders
    Public NotInheritable Class ButtonBuilder
        Inherits Builder(Of Button)

        Public Overrides Function Build() As Button
            Return Build(Nothing)
        End Function

        Public Overrides Function Build(tag As Object) As Button
            Return New Button(Attributes, tag)
        End Function

        Public Function WithLabel(label As String, Optional copyToScreenTip As Boolean = True) As ButtonBuilder
            AddLabel(label)
            AddShowLabel(True)

            If copyToScreenTip Then
                AddScreenTip(label)
            End If

            Return Me
        End Function

        Public Function WithLabel(label As String, callback As FromControl(Of String), Optional copyToScreenTip As Boolean = True) As ButtonBuilder
            AddLabel(label, callback)
            AddShowLabel(True)

            If copyToScreenTip Then
                AddScreenTip(label, callback)
            End If

            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String) As ButtonBuilder
            AddScreenTip(screenTip)
            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String, callback As FromControl(Of String)) As ButtonBuilder
            AddScreenTip(screenTip, callback)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String) As ButtonBuilder
            AddSuperTip(superTip)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String, callback As FromControl(Of String)) As ButtonBuilder
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