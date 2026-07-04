Imports RibbonX.Api
Imports RibbonX.Controls
Imports RibbonX.InternalApi
Imports RibbonX.Utilities

Namespace Builders

    Friend NotInheritable Class DialogLauncherBuilder
        Inherits BuilderBase(Of DialogLauncher)
        Implements IDialogLauncherBuilder

        Private button As IButton

        Public Function WithButton(button As IButton) As IDialogLauncherBuilder Implements IDialogLauncherBuilder.WithButton
            Me.button = button
            Return Me
        End Function

        Public Function WithTag(tag As Object) As IDialogLauncherBuilder Implements IDialogLauncherBuilder.WithTag
            WithTagBase(tag)
            Return Me
        End Function

        Public Overrides Function Build() As DialogLauncher
            Return New DialogLauncher(ControlProperties, NotNull(button), Tag)
        End Function

        Public Shared Function FromAction(Optional action As Action(Of IDialogLauncherBuilder) = Nothing) As IDialogLauncher
            Dim instance As DialogLauncherBuilder = New DialogLauncherBuilder()

            action?.Invoke(instance)

            Return instance.Build()
        End Function

    End Class

End Namespace

