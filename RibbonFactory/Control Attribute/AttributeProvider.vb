Imports RibbonFactory.Attributes

Friend NotInheritable Class AttributeProvider
    Friend Shared Function GetDefaults(Of T As RibbonElement)() As ControlPropertyList
        GetDefaults = New ControlPropertyList

        For Each InterfaceType As Type In GetType(T).GetInterfaces
            If InterfaceType Is GetType(IActionable) Then
                GetDefaults.Add(New ControlProperty(Of System.Action)(Action.OnAction, Sub() MsgBox("TODO: Supply actual action.")))

            ElseIf InterfaceType Is GetType(IDisablable) Then
                GetDefaults.Add(New ControlProperty(Of Boolean)(Enabled.Enabled, True))
                GetDefaults.Add(New ControlProperty(Of Boolean)(Visible.Visible, True))

            ElseIf InterfaceType Is GetType(IDescribable) Then
                GetDefaults.Add(New ControlProperty(Of Boolean)(ShowLabel.ShowLabel, False))
                GetDefaults.Add(New ControlProperty(Of String)(Label.Label, ""))
                GetDefaults.Add(New ControlProperty(Of String)(Screentip.Screentip, ""))
                GetDefaults.Add(New ControlProperty(Of String)(Supertip.Supertip, ""))

            ElseIf InterfaceType Is GetType(IEditable) Then
                GetDefaults.Add(New ControlProperty(Of String)(Text.SizeString, New String("F", 20)))
                GetDefaults.Add(New ControlProperty(Of Byte)(Text.MaxLength, 50))
                GetDefaults.Add(New DynamicControlProperty(Of String)(Text.GetText, "GetText", String.Empty))
                GetDefaults.Add(New ControlProperty(Of System.Action)(Text.OnChange, Sub() MsgBox("TODO: Supply actual action.")))

            ElseIf InterfaceType Is GetType(IGraphic) Then
                GetDefaults.Add(New ControlProperty(Of Boolean)(ShowImage.ShowImage, False))
                GetDefaults.Add(New ControlProperty(Of ImageMSO)(Image.ImageMso, ImageMSO.HappyFace))

            ElseIf InterfaceType Is GetType(IResizable) Then
                GetDefaults.Add(New ControlProperty(Of ControlSize)(Size.Size, ControlSize.large))

            ElseIf InterfaceType Is GetType(ISelectable) Then

            ElseIf InterfaceType Is GetType(IToggleable) Then
                GetDefaults.Add(New DynamicControlProperty(Of Boolean)(Toggleable.GetPressed, "GetPressed", False))

            ElseIf InterfaceType Is GetType(IVeryDescribable) Then
                GetDefaults.Add(New ControlProperty(Of String)(Description.Description, String.Empty))

            End If
        Next InterfaceType
    End Function
End Class
