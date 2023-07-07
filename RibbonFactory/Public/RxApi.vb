Imports RibbonX.Api
Imports RibbonX.Builders
Imports RibbonX.InternalApi

Public Class RxApi

    Public Shared ReadOnly Property Instance As RxApi = New RxApi()

    Public Shared Function Box(Optional options As Action(Of IBoxBuilder) = Nothing) As IBox
        Return BoxBuilder.FromAction(options)
    End Function

    Public Shared Function Button(Optional builder As Action(Of IButtonBuilder) = Nothing) As IButton
        Return ButtonBuilder.FromAction(builder)
    End Function

    Public Shared Function ButtonGroup(Optional builder As Action(Of IButtonGroupBuilder) = Nothing) As IButtonGroup
        Return ButtonGroupBuilder.FromAction(builder)
    End Function

    Public Shared Function CheckBox(Optional builder As Action(Of ICheckBoxBuilder) = Nothing) As ICheckbox
        Return CheckBoxBuilder.FromAction(builder)
    End Function

    Public Shared Function ComboBox(Optional builder As Action(Of IComboBoxBuilder) = Nothing) As IComboBox
        Return ComboBoxBuilder.FromAction(builder)
    End Function

    Public Shared Function DialogLauncher(button As IButton, Optional tag As Object = Nothing) As IDialogLauncher
        Return DialogLauncherBuilder.FromAction(Sub(b) b.WithButton(button).WithTag(tag))
    End Function

    Public Shared Function DropDown(Optional builder As Action(Of IDropdownBuilder) = Nothing) As IDropDown
        Return DropDownBuilder.FromAction(builder)
    End Function

    Public Shared Function EditBox(Optional builder As Action(Of IEditBoxBuilder) = Nothing) As IEditBox
        Return EditBoxBuilder.FromAction(builder)
    End Function

    Public Shared Function Gallery(Optional builder As Action(Of IGalleryBuilder) = Nothing) As IGallery
        Return GalleryBuilder.FromAction(builder)
    End Function

    Public Shared Function Group(Optional builder As Action(Of IGroupBuilder) = Nothing) As IGroup
        Return GroupBuilder.FromAction(builder)
    End Function

    Public Shared Function Item(Optional builder As Action(Of IItemBuilder) = Nothing) As IItem
        Return ItemBuilder.FromAction(builder)
    End Function

    Public Shared Function LabelControl(Optional builder As Action(Of ILabelControlBuilder) = Nothing) As ILabelControl
        Return LabelControlBuilder.FromAction(builder)
    End Function

    Public Shared Function Menu(Optional builder As Action(Of IMenuBuilder) = Nothing) As IMenu
        Return MenuBuilder.FromAction(builder)
    End Function

    Public Shared Function MenuSeparator(Optional builder As Action(Of IMenuSeparatorBuilder) = Nothing) As IMenuSeparator
        Return MenuSeparatorBuilder.FromAction(builder)
    End Function

    Public Shared Function Ribbon(Optional builder As Action(Of IRibbonBuilder) = Nothing) As IRibbon
        Return RibbonBuilder.FromAction(builder)
    End Function

    Public Shared Function Separator(Optional builder As Action(Of ISeparatorBuilder) = Nothing) As ISeparator
        Return SeparatorBuilder.FromAction(builder)
    End Function

    Public Shared Function SplitButton(Optional builder As Action(Of ISplitButtonBuilder) = Nothing) As ISplitButton
        Return SplitButtonBuilder.FromAction(builder)
    End Function

    Public Shared Function Tab(Optional builder As Action(Of ITabBuilder) = Nothing) As ITab
        Return TabBuilder.FromAction(builder)
    End Function

    Public Shared Function ToggleButton(Optional builder As Action(Of IToggleButtonBuilder) = Nothing) As IToggleButton
        Return ToggleButtonBuilder.FromAction(builder)
    End Function

End Class
