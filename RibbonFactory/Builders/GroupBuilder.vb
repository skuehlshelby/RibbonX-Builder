Imports RibbonFactory.Containers
Imports RibbonFactory.Enums
Imports RibbonFactory.Enums.ImageMSO

Namespace Builders
    Public Class GroupBuilder
        Inherits Builder(Of Group)

        Public Overrides Function Build() As Group
            Return Build(Nothing)
        End Function

        Public Overrides Function Build(tag As Object) As Group
            Return New Group(Attributes, tag)
        End Function

        Public Function WithVisibility(visible As Boolean) As GroupBuilder
            AddVisible(visible)
            Return Me
        End Function

        Public Function WithVisibility(visible As Boolean, callback As FromControl(Of Boolean)) As GroupBuilder
            AddVisible(visible, callback)
            Return Me
        End Function

        Public Function WithLabel(label As String, Optional copyToScreenTip As Boolean = True) As GroupBuilder
            AddLabel(label)
            AddShowLabel(True)

            If copyToScreenTip Then
                AddScreenTip(label)
            End If

            Return Me
        End Function

        Public Function WithLabel(label As String, callback As FromControl(Of String), Optional copyToScreenTip As Boolean = True) As GroupBuilder
            AddLabel(label, callback)
            AddShowLabel(True)

            If copyToScreenTip Then
                AddScreenTip(label, callback)
            End If

            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String) As GroupBuilder
            AddScreenTip(screenTip)
            Return Me
        End Function

        Public Function WithScreenTip(screenTip As String, callback As FromControl(Of String)) As GroupBuilder 
            AddScreenTip(screenTip, callback)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String) As GroupBuilder
            AddSuperTip(superTip)
            Return Me
        End Function

        Public Function WithSuperTip(superTip As String, callback As FromControl(Of String)) As GroupBuilder
            AddSuperTip(superTip, callback)
            Return Me
        End Function

        Public Function WithImage(image As ImageMSO) As GroupBuilder
            AddImage(image)
            Return Me
        End Function
    End Class
End NameSpace