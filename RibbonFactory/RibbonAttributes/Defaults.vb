Imports RibbonFactory.ControlInterfaces
Imports RibbonFactory.Controls
Imports RibbonFactory.Controls.Events
Imports RibbonFactory.Enums

Namespace RibbonAttributes

    Friend NotInheritable Class Defaults(Of T As RibbonElement)
        Implements IDefaultProvider

        Public Function GetDefaults() As AttributeSet Implements IDefaultProvider.GetDefaults
            Dim defaults As AttributeSet = New AttributeSet()

            defaults.Add(New RibbonAttributeReadOnly(Of String)(IdManager.GetID(Of T)(), AttributeName.Id, AttributeCategory.IdType))

            Dim implementedInterfaces As Type() = GetType(T).GetInterfaces().ToArray()

            For Each interfaceType As Type In implementedInterfaces
                Select Case interfaceType
                    Case GetType(IEnabled)
                        defaults.Add(New RibbonAttributeDefault(Of Boolean)(True, AttributeName.Enabled, AttributeCategory.Enabled))
                    Case GetType(IVisible)
                        defaults.Add(New RibbonAttributeDefault(Of Boolean)(True, AttributeName.Visible, AttributeCategory.Visibility))
                    Case GetType(ILabel)
                        defaults.Add(New RibbonAttributeDefault(Of String)(String.Empty, AttributeName.Label, AttributeCategory.Label))
                    Case GetType(IShowLabel)
                        defaults.Add(New RibbonAttributeDefault(Of Boolean)(False, AttributeName.ShowLabel, AttributeCategory.LabelVisibility))
                    Case GetType(IScreenTip)
                        defaults.Add(New RibbonAttributeDefault(Of String)(String.Empty, AttributeName.Screentip, AttributeCategory.ScreenTip))
                    Case GetType(ISuperTip)
                        defaults.Add(New RibbonAttributeDefault(Of String)(String.Empty, AttributeName.Supertip, AttributeCategory.SuperTip))
                    Case GetType(ISize)
                        defaults.Add(New RibbonAttributeDefault(Of ControlSize)(ControlSize.normal, AttributeName.Size, AttributeCategory.Size))
                    Case GetType(IDescription)
                        defaults.Add(New RibbonAttributeDefault(Of String)(String.Empty, AttributeName.Description, AttributeCategory.Description))
                    Case GetType(IKeyTip)
                        defaults.Add(New RibbonAttributeDefault(Of KeyTip)("A", AttributeName.Keytip, AttributeCategory.KeyTip))
                    Case GetType(IImage)
                        defaults.Add(New RibbonAttributeDefault(Of RibbonImage)(RibbonImage.Create(ImageMSO.Common.HappyFace), AttributeName.ImageMso, AttributeCategory.Image))
                    Case GetType(IShowImage)
                        defaults.Add(New RibbonAttributeDefault(Of Boolean)(False, AttributeName.ShowImage, AttributeCategory.ImageVisibility))
                    Case GetType(ITitle)
                        defaults.Add(New RibbonAttributeDefault(Of String)(String.Empty, AttributeName.Title, AttributeCategory.Title))
                    Case GetType(IClickable)
                        defaults.Add(New RibbonAttributeDefault(Of Action)(Sub() Return, AttributeName.OnAction, AttributeCategory.OnAction))
                        defaults.Add(New RibbonAttributeInvocationList(Of CancelableEventArgs)(AttributeName.BeforeClick, AttributeCategory.BeforeClick))
                        defaults.Add(New RibbonAttributeInvocationList(AttributeName.OnClick, AttributeCategory.OnClick))
                    Case GetType(IText)
                        defaults.Add(New RibbonAttributeDefault(Of String)(String.Empty, AttributeName.GetText, AttributeCategory.Text))
                        defaults.Add(New RibbonAttributeDefault(Of String)("WWWWWW", AttributeName.SizeString, AttributeCategory.SizeString))
                        defaults.Add(New RibbonAttributeDefault(Of Byte)(Byte.MaxValue, AttributeName.MaxLength, AttributeCategory.MaxLength))
                        defaults.Add(New RibbonAttributeInvocationList(Of BeforeTextChangeEventArgs)(AttributeName.BeforeTextChange, AttributeCategory.BeforeTextChange))
                        defaults.Add(New RibbonAttributeInvocationList(Of TextChangeEventArgs)(AttributeName.OnTextChange, AttributeCategory.OnTextChange))
                    Case GetType(IPressed)
                        defaults.Add(New RibbonAttributeDefault(Of Boolean)(False, AttributeName.GetPressed, AttributeCategory.Pressed))
                        defaults.Add(New RibbonAttributeInvocationList(Of BeforeToggleChangeEventArgs)(AttributeName.BeforeToggle, AttributeCategory.BeforeToggle))
                        defaults.Add(New RibbonAttributeInvocationList(Of ToggleChangeEventArgs)(AttributeName.OnToggle, AttributeCategory.OnToggle))
                    Case GetType(ISelect)
                        defaults.Add(New RibbonAttributeDefault(Of Item)(Item.Blank(), AttributeName.GetSelectedItemIndex, AttributeCategory.SelectedItemPosition))
                        defaults.Add(New RibbonAttributeInvocationList(Of BeforeSelectionChangeEventArgs)(AttributeName.BeforeSelectionChange, AttributeCategory.BeforeSelectionChange))
                        defaults.Add(New RibbonAttributeInvocationList(Of SelectionChangeEventArgs)(AttributeName.OnSelectionChange, AttributeCategory.OnSelectionChange))
                    Case GetType(IItemDimensions)
                        defaults.Add(New RibbonAttributeDefault(Of Integer)(1, AttributeName.ItemHeight, AttributeCategory.ItemHeight))
                        defaults.Add(New RibbonAttributeDefault(Of Integer)(1, AttributeName.ItemWidth, AttributeCategory.ItemWidth))
                    Case GetType(IItemSize)
                        defaults.Add(New RibbonAttributeDefault(Of ControlSize)(ControlSize.normal, AttributeName.ItemSize, AttributeCategory.ItemSize))
                    Case GetType(IBoxStyle)
                        defaults.Add(New RibbonAttributeDefault(Of BoxStyle)(BoxStyle.horizontal, AttributeName.BoxStyle, AttributeCategory.BoxStyle))
                End Select
            Next interfaceType

            Return defaults
        End Function

        Public Shared Function [Get]() As IEnumerable(Of IRibbonAttribute)
            Return New Defaults(Of T)().GetDefaults()
        End Function
    End Class

End Namespace