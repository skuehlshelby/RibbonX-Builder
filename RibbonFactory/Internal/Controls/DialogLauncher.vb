
Imports RibbonX.Controls.Base
Imports RibbonX.InternalApi

Namespace Controls

    Friend NotInheritable Class DialogLauncher
        Inherits ReadOnlyContainer(Of IButton)
        Implements IDialogLauncher

        Public Sub New(attributes As IPropertyCollection, button As IButton, Optional tag As Object = Nothing)
            MyBase.New(attributes, {button}, tag)
        End Sub

        Public Overrides Function ToXml(Optional exclude As ExcludedAttributes = ExcludedAttributes.None) As String
            Return $"<dialogBoxLauncher>{Items.Single().ToXml(ExcludedAttributes.SizeAndVisibility)}</dialogBoxLauncher>"
        End Function

        Public Overrides Function Clone() As Object
            Return New DialogLauncher(CType(Attributes.Clone(), IPropertyCollection), CType(Items.Single().Clone(), IButton), Tag)
        End Function

    End Class

End Namespace