Imports System.Drawing
Imports RibbonX.ControlInterfaces
Imports RibbonX.Controls
Imports RibbonX.Controls.Events
Imports RibbonX.Enums
Imports RibbonX.Enums.ImageMSO
Imports RibbonX.Enums.MSO
Imports RibbonX.RibbonAttributes
Imports RibbonX.Utilities
Imports RibbonX.ComTypes.StdOle

Namespace Builders

    Friend Class BuilderBase(Of TElement As RibbonElement)
        Private ReadOnly _attributes As AttributeSet = GetDefaults()

        Public Sub New(Optional template As RibbonElement = Nothing)
            If template IsNot Nothing Then
                Dim defaultProvider As IDefaultProvider = TryCast(template, IDefaultProvider)

                If defaultProvider Is Nothing Then
                    Throw New ArgumentException($"Could not copy properties from '{template.ID}' to type '{GetType(TElement)}'")
                Else
                    _attributes.OverwriteWithIntersectionOf(defaultProvider.GetDefaults())
                End If
            End If
        End Sub

        Public Function Build() As AttributeSet
            Return _attributes
        End Function

#Region "Defaults"
        Private Shared Function GetNextElementId() As Integer
            Static id As Integer = 0
            id = If(id = Integer.MaxValue, 0, id + 1)
            Return id
        End Function

        Private Function GetDefaults() As AttributeSet
            Dim elementType As Type = GetType(TElement)

            Dim defaults As AttributeSet = New AttributeSet From {
                New RibbonAttributeReadOnly(Of String)(elementType.Name & GetNextElementId(), AttributeName.Id, AttributeCategory.IdType)
            }

            Dim implementedInterfaces As Type() = elementType.GetInterfaces()

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
                        defaults.Add(New RibbonAttributeDefault(Of ControlSize)(ControlSize.Normal, AttributeName.Size, AttributeCategory.Size))
                    Case GetType(IDescription)
                        defaults.Add(New RibbonAttributeDefault(Of String)(String.Empty, AttributeName.Description, AttributeCategory.Description))
                    Case GetType(IKeyTip)
                        defaults.Add(New RibbonAttributeDefault(Of KeyTip)("A", AttributeName.Keytip, AttributeCategory.KeyTip))
                    Case GetType(IImage)
                        defaults.Add(New RibbonAttributeDefault(Of RibbonImage)(RibbonImage.Create(Common.HappyFace), AttributeName.ImageMso, AttributeCategory.Image))
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
                        defaults.Add(New RibbonAttributeDefault(Of Item)(Nothing, AttributeName.GetSelectedItemIndex, AttributeCategory.SelectedItemPosition))
                        defaults.Add(New RibbonAttributeInvocationList(Of BeforeSelectionChangeEventArgs)(AttributeName.BeforeSelectionChange, AttributeCategory.BeforeSelectionChange))
                        defaults.Add(New RibbonAttributeInvocationList(Of SelectionChangeEventArgs)(AttributeName.OnSelectionChange, AttributeCategory.OnSelectionChange))
                    Case GetType(IItemDimensions)
                        defaults.Add(New RibbonAttributeDefault(Of Integer)(1, AttributeName.ItemHeight, AttributeCategory.ItemHeight))
                        defaults.Add(New RibbonAttributeDefault(Of Integer)(1, AttributeName.ItemWidth, AttributeCategory.ItemWidth))
                    Case GetType(IItemSize)
                        defaults.Add(New RibbonAttributeDefault(Of ControlSize)(ControlSize.Normal, AttributeName.ItemSize, AttributeCategory.ItemSize))
                    Case GetType(IBoxStyle)
                        defaults.Add(New RibbonAttributeDefault(Of BoxStyle)(BoxStyle.Horizontal, AttributeName.BoxStyle, AttributeCategory.BoxStyle))
                End Select
            Next interfaceType

            Return defaults
        End Function

#End Region

#Region "Enabled"

        Protected Sub EnabledBase()
            _attributes.Add(New RibbonAttributeReadOnly(Of Boolean)(True, AttributeName.Enabled, AttributeCategory.Enabled))
        End Sub

        Protected Sub EnabledBase(callback As FromControl(Of Boolean))
            _attributes.Add(New RibbonAttributeReadWrite(Of Boolean)(True, AttributeName.GetEnabled, AttributeCategory.Enabled, callback.Method.Name))
        End Sub

#End Region

#Region "Disabled"

        Protected Sub DisabledBase()
            _attributes.Add(New RibbonAttributeReadOnly(Of Boolean)(False, AttributeName.Enabled, AttributeCategory.Enabled))
        End Sub

        Protected Sub DisabledBase(callback As FromControl(Of Boolean))
            _attributes.Add(New RibbonAttributeReadWrite(Of Boolean)(False, AttributeName.GetEnabled, AttributeCategory.Enabled, callback.Method.Name))
        End Sub

#End Region

#Region "Visible"
        Protected Sub VisibleBase()
            _attributes.Add(New RibbonAttributeReadOnly(Of Boolean)(True, AttributeName.Visible, AttributeCategory.Visibility))
        End Sub

        Protected Sub VisibleBase(callback As FromControl(Of Boolean))
            _attributes.Add(New RibbonAttributeReadWrite(Of Boolean)(True, AttributeName.GetVisible, AttributeCategory.Visibility, callback.Method.Name))
        End Sub
#End Region

#Region "Invisible"
        Protected Sub InvisibleBase()
            _attributes.Add(New RibbonAttributeReadOnly(Of Boolean)(False, AttributeName.Visible, AttributeCategory.Visibility))
        End Sub

        Protected Sub InvisibleBase(callback As FromControl(Of Boolean))
            _attributes.Add(New RibbonAttributeReadWrite(Of Boolean)(False, AttributeName.GetVisible, AttributeCategory.Visibility, callback.Method.Name))
        End Sub
#End Region

#Region "Id"

        Protected Sub WithIdBase(id As String)
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(id, AttributeName.Id, AttributeCategory.IdType))
        End Sub

        Protected Sub WithIdBase([namespace] As String, id As String)
            _attributes.Add(New RibbonAttributeReadOnly(Of String)($"{[namespace]}:{id}", AttributeName.IdQ, AttributeCategory.IdType))
        End Sub

        Protected Sub WithIdBase(mso As MSO)
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(mso.ToString(), AttributeName.IdMso, AttributeCategory.IdType))
        End Sub

#End Region

#Region "Insert Before/After"

        Protected Sub InsertAfterMsoBase(mso As MSO)
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(mso.ToString(), AttributeName.InsertAfterMso, AttributeCategory.Insertion))
        End Sub

        Protected Sub InsertAfterQBase(control As RibbonElement)
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(control.ID, AttributeName.InsertAfterQ, AttributeCategory.Insertion))
        End Sub

        Protected Sub InsertBeforeMsoBase(mso As MSO)
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(mso.ToString(), AttributeName.InsertBeforeMso, AttributeCategory.Insertion))
        End Sub

        Protected Sub InsertBeforeQBase(control As RibbonElement)
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(control.ID, AttributeName.InsertBeforeQ, AttributeCategory.Insertion))
        End Sub

#End Region

#Region "Label"

        Protected Sub LabelBase(label As String)
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(label, AttributeName.Label, AttributeCategory.Label))
        End Sub

        Protected Sub LabelBase(label As String, callback As FromControl(Of String))
            _attributes.Add(New RibbonAttributeReadWrite(Of String)(label, AttributeName.GetLabel, AttributeCategory.Label, callback.Method.Name))
        End Sub

#End Region

#Region "Show Label"

        Protected Sub HideLabelBase()
            _attributes.Add(New RibbonAttributeReadOnly(Of Boolean)(False, AttributeName.ShowLabel, AttributeCategory.LabelVisibility))
        End Sub

        Protected Sub HideLabelBase(callback As FromControl(Of Boolean))
            _attributes.Add(New RibbonAttributeReadWrite(Of Boolean)(False, AttributeName.GetShowLabel, AttributeCategory.LabelVisibility, callback.Method.Name))
        End Sub

        Protected Sub ShowLabelBase()
            _attributes.Add(New RibbonAttributeReadOnly(Of Boolean)(True, AttributeName.ShowLabel, AttributeCategory.LabelVisibility))
        End Sub

        Protected Sub ShowLabelBase(callback As FromControl(Of Boolean))
            _attributes.Add(New RibbonAttributeReadWrite(Of Boolean)(True, AttributeName.GetShowLabel, AttributeCategory.LabelVisibility, callback.Method.Name))
        End Sub

#End Region

#Region "Description"

        Protected Sub DescriptionBase(description As String)
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(description, AttributeName.Description, AttributeCategory.Description))
        End Sub

        Protected Sub DescriptionBase(description As String, callback As FromControl(Of String))
            _attributes.Add(New RibbonAttributeReadWrite(Of String)(description, AttributeName.GetDescription, AttributeCategory.Description, callback.Method.Name))
        End Sub

#End Region

#Region "KeyTip"

        Protected Sub KeyTipBase(keyTip As KeyTip)
            _attributes.Add(New RibbonAttributeReadOnly(Of KeyTip)(keyTip, AttributeName.Keytip, AttributeCategory.KeyTip))
        End Sub

        Protected Sub KeyTipBase(keyTip As KeyTip, callback As FromControl(Of KeyTip))
            _attributes.Add(New RibbonAttributeReadWrite(Of KeyTip)(keyTip, AttributeName.GetKeytip, AttributeCategory.KeyTip, callback.Method.Name))
        End Sub

#End Region

#Region "Screentip"

        Protected Sub ScreentipBase(screentip As String)
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(screentip, AttributeName.Screentip, AttributeCategory.ScreenTip))
        End Sub

        Protected Sub ScreentipBase(screentip As String, callback As FromControl(Of String))
            _attributes.Add(New RibbonAttributeReadWrite(Of String)(screentip, AttributeName.GetScreentip, AttributeCategory.ScreenTip, callback.Method.Name))
        End Sub

#End Region

#Region "Supertip"

        Protected Sub SupertipBase(supertip As String)
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(supertip, AttributeName.Supertip, AttributeCategory.SuperTip))
        End Sub

        Protected Sub SupertipBase(supertip As String, callback As FromControl(Of String))
            _attributes.Add(New RibbonAttributeReadWrite(Of String)(supertip, AttributeName.GetSupertip, AttributeCategory.SuperTip, callback.Method.Name))
        End Sub

#End Region

#Region "Title"

        Protected Sub TitleBase(title As String)
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(title, AttributeName.Title, AttributeCategory.Title))
        End Sub

        Protected Sub TitleBase(title As String, callback As FromControl(Of String))
            _attributes.Add(New RibbonAttributeReadWrite(Of String)(title, AttributeName.GetTitle, AttributeCategory.Title, callback.Method.Name))
        End Sub

#End Region

#Region "Image"

        Protected Sub ImageBase(path As String)
            _attributes.Add(New RibbonAttributeReadOnly(Of RibbonImage)(RibbonImage.Create(path), AttributeName.Image, AttributeCategory.Image))
        End Sub

        Protected Sub ImageBase(image As ImageMSO)
            _attributes.Add(New RibbonAttributeReadOnly(Of RibbonImage)(RibbonImage.Create(image), AttributeName.ImageMso, AttributeCategory.Image))
        End Sub

        Protected Sub ImageBase(image As IPictureDisp, callback As FromControl(Of IPictureDisp))
            _attributes.Add(New RibbonAttributeReadWrite(Of RibbonImage)(RibbonImage.Create(image), AttributeName.GetImage, AttributeCategory.Image, callback.Method.Name))
        End Sub

        Protected Sub ImageBase(image As Bitmap, callback As FromControl(Of IPictureDisp))
            ImageBase(New AdapterForIPicture(image), callback)
        End Sub

        Protected Sub ImageBase(image As Icon, callback As FromControl(Of IPictureDisp))
            ImageBase(New AdapterForIPicture(image), callback)
        End Sub

#End Region

#Region "Show Image"

        Protected Sub ShowImageBase()
            _attributes.Add(New RibbonAttributeReadOnly(Of Boolean)(True, AttributeName.ShowImage, AttributeCategory.ImageVisibility))
        End Sub

        Protected Sub ShowImageBase(callback As FromControl(Of Boolean))
            _attributes.Add(New RibbonAttributeReadWrite(Of Boolean)(True, AttributeName.GetShowImage, AttributeCategory.ImageVisibility, callback.Method.Name))
        End Sub

#End Region

#Region "Hide Image"

        Protected Sub HideImageBase()
            _attributes.Add(New RibbonAttributeReadOnly(Of Boolean)(False, AttributeName.ShowImage, AttributeCategory.ImageVisibility))
        End Sub

        Protected Sub HideImageBase(callback As FromControl(Of Boolean))
            _attributes.Add(New RibbonAttributeReadWrite(Of Boolean)(False, AttributeName.GetShowImage, AttributeCategory.ImageVisibility, callback.Method.Name))
        End Sub

#End Region

#Region "Normal"

        Protected Sub NormalBase()
            _attributes.Add(New RibbonAttributeReadOnly(Of ControlSize)(ControlSize.Normal, AttributeName.Size, AttributeCategory.Size))
        End Sub

        Protected Sub NormalBase(callback As FromControl(Of ControlSize))
            _attributes.Add(New RibbonAttributeReadWrite(Of ControlSize)(ControlSize.Normal, AttributeName.GetSize, AttributeCategory.Size, callback.Method.Name))
        End Sub

#End Region

#Region "Large"

        Protected Sub LargeBase()
            _attributes.Add(New RibbonAttributeReadOnly(Of ControlSize)(ControlSize.Large, AttributeName.Size, AttributeCategory.Size))
        End Sub

        Protected Sub LargeBase(callback As FromControl(Of ControlSize))
            _attributes.Add(New RibbonAttributeReadWrite(Of ControlSize)(ControlSize.Large, AttributeName.GetSize, AttributeCategory.Size, callback.Method.Name))
        End Sub

#End Region

#Region "Size String"

        Protected Sub SizeStringBase(sizeString As String)
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(sizeString, AttributeName.SizeString, AttributeCategory.SizeString))
        End Sub

#End Region

#Region "Max Length"

        Protected Sub MaxLengthBase(maxLength As Byte)
            _attributes.Add(New RibbonAttributeReadOnly(Of Byte)(maxLength, AttributeName.MaxLength, AttributeCategory.MaxLength))
        End Sub

#End Region

#Region "Horizontal/Vertical"

        Protected Sub HorizontalBase()
            _attributes.Add(New RibbonAttributeReadOnly(Of BoxStyle)(BoxStyle.Horizontal, AttributeName.BoxStyle, AttributeCategory.BoxStyle))
        End Sub

        Protected Sub VerticalBase()
            _attributes.Add(New RibbonAttributeReadOnly(Of BoxStyle)(BoxStyle.Vertical, AttributeName.BoxStyle, AttributeCategory.BoxStyle))
        End Sub

#End Region

#Region "On Action"

        Protected Sub OnActionBase(callback As OnAction)
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(callback.Method.Name, AttributeName.OnAction, AttributeCategory.OnAction))
        End Sub

        Protected Sub OnActionBase(callback As SelectionChanged)
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(callback.Method.Name, AttributeName.OnAction, AttributeCategory.OnAction))
        End Sub

#End Region

#Region "Get/Set Text"

        Protected Sub TextBase(text As String, getText As FromControl(Of String), setText As TextChanged)
            Dim wrapper As ValueWrapper(Of String) = New ValueWrapper(Of String)(text)
            _attributes.Add(New RibbonAttributeWrappedValue(Of String)(wrapper, getText.Method.Name, AttributeName.GetText, AttributeCategory.Text))
            _attributes.Add(New RibbonAttributeWrappedValue(Of String)(wrapper, setText.Method.Name, AttributeName.OnChange, AttributeCategory.OnChange))
        End Sub

#End Region

#Region "Get/Set Pressed"

        Protected Sub PressedBase(pressed As Boolean, getPressed As FromControl(Of Boolean), setPressed As ButtonPressed)
            Dim wrapper As ValueWrapper(Of Boolean) = New ValueWrapper(Of Boolean)(pressed)
            _attributes.Add(New RibbonAttributeWrappedValue(Of Boolean)(wrapper, getPressed.Method.Name, AttributeName.GetPressed, AttributeCategory.Pressed))
            _attributes.Add(New RibbonAttributeWrappedValue(Of Boolean)(wrapper, setPressed.Method.Name, AttributeName.OnAction, AttributeCategory.OnAction))
        End Sub

#End Region

#Region "Add Handlers"

        Protected Sub AddBeforeClickHandlers(handlers As IEnumerable(Of Action(Of Button, CancelableEventArgs)))
            _attributes.Add(New RibbonAttributeInvocationList(Of CancelableEventArgs)(
                            New List(Of EventHandler(Of CancelableEventArgs))(FromElementAndEventArgs(handlers)), AttributeName.BeforeClick, AttributeCategory.BeforeClick))
        End Sub

        Protected Sub AddOnClickHandlers(handlers As IEnumerable(Of Action(Of Button)))
            _attributes.Add(New RibbonAttributeInvocationList(
                            New List(Of EventHandler)(FromElements(handlers)), AttributeName.OnClick, AttributeCategory.OnClick))
        End Sub

        Protected Sub AddBeforeTextChangeHandlers(handlers As IEnumerable(Of EventHandler(Of BeforeTextChangeEventArgs)))
            _attributes.Add(New RibbonAttributeInvocationList(Of BeforeTextChangeEventArgs)(
                            New List(Of EventHandler(Of BeforeTextChangeEventArgs))(handlers), AttributeName.BeforeTextChange, AttributeCategory.BeforeTextChange))
        End Sub

        Protected Sub AddOnTextChangeHandlers(handlers As IEnumerable(Of EventHandler(Of TextChangeEventArgs)))
            _attributes.Add(New RibbonAttributeInvocationList(Of TextChangeEventArgs)(
                            New List(Of EventHandler(Of TextChangeEventArgs))(handlers), AttributeName.OnTextChange, AttributeCategory.OnTextChange))
        End Sub

        Protected Sub AddBeforeToggleHandlers(handlers As IEnumerable(Of Action(Of BeforeToggleChangeEventArgs)))
            _attributes.Add(New RibbonAttributeInvocationList(Of BeforeToggleChangeEventArgs)(
                            New List(Of EventHandler(Of BeforeToggleChangeEventArgs))(FromEventArgs(handlers)), AttributeName.BeforeToggle, AttributeCategory.BeforeToggle))
        End Sub

        Protected Sub AddToggleHandlers(handlers As IEnumerable(Of Action(Of ToggleChangeEventArgs)))
            _attributes.Add(New RibbonAttributeInvocationList(Of ToggleChangeEventArgs)(
                            New List(Of EventHandler(Of ToggleChangeEventArgs))(FromEventArgs(handlers)), AttributeName.OnToggle, AttributeCategory.OnToggle))
        End Sub

        Protected Sub AddBeforeSelectionChangeHandlers(handlers As IEnumerable(Of Action(Of BeforeSelectionChangeEventArgs)))
            _attributes.Add(New RibbonAttributeInvocationList(Of BeforeSelectionChangeEventArgs)(
                            New List(Of EventHandler(Of BeforeSelectionChangeEventArgs))(FromEventArgs(handlers)), AttributeName.BeforeSelectionChange, AttributeCategory.BeforeSelectionChange))
        End Sub

        Protected Sub AddSelectionChangeHandlers(handlers As IEnumerable(Of Action(Of SelectionChangeEventArgs)))
            _attributes.Add(New RibbonAttributeInvocationList(Of SelectionChangeEventArgs)(
                            New List(Of EventHandler(Of SelectionChangeEventArgs))(FromEventArgs(handlers)), AttributeName.OnSelectionChange, AttributeCategory.OnSelectionChange))
        End Sub

        Private Function FromElements(Of T As RibbonElement)(actions As IEnumerable(Of Action(Of T))) As IEnumerable(Of EventHandler)
            Return actions.Select(Function(a) New EventHandler(Sub(sender, e) a.Invoke(CType(sender, T))))
        End Function

        Private Function FromEventArgs(Of T As EventArgs)(actions As IEnumerable(Of Action(Of T))) As IEnumerable(Of EventHandler(Of T))
            Return actions.Select(Function(a) New EventHandler(Of T)(Sub(sender, e) a.Invoke(e)))
        End Function

        Private Function FromElementAndEventArgs(Of T As RibbonElement, K As EventArgs)(actions As IEnumerable(Of Action(Of T, K))) As IEnumerable(Of EventHandler(Of K))
            Return actions.Select(Function(a) New EventHandler(Of K)(Sub(sender, e) a.Invoke(CType(sender, T), e)))
        End Function

#End Region

#Region "Invalidate Content On Drop"

        Protected Sub InvalidateContentOnDropBase()
            _attributes.Add(New RibbonAttributeReadOnly(Of Boolean)(True, AttributeName.InvalidateContentOnDrop, AttributeCategory.ContentInvalidation))
        End Sub

#End Region

#Region "Do Not Show In Ribbon"
        Protected Sub DoNotShowInRibbonBase()
            _attributes.Add(New RibbonAttributeReadOnly(Of Boolean)(False, AttributeName.ShowInRibbon, AttributeCategory.ShowInRibbon))
        End Sub

#End Region

        '' Child Item Methods ''

#Region "Selected Item"

        Protected Sub SelectedBase(getSelected As FromControl(Of String), setSelected As SelectionChanged)
            Dim wrapped As ValueWrapper(Of Item) = New ValueWrapper(Of Item)(Nothing)
            _attributes.Add(New RibbonAttributeWrappedValue(Of Item)(wrapped, getSelected.Method.Name, AttributeName.GetSelectedItemID, AttributeCategory.SelectedItemPosition))
            _attributes.Add(New RibbonAttributeWrappedValue(Of Item)(wrapped, setSelected.Method.Name, AttributeName.OnAction, AttributeCategory.OnAction))
        End Sub

        Protected Sub SelectedBase(getSelected As FromControl(Of Integer), setSelected As SelectionChanged)
            Dim wrapped As ValueWrapper(Of Item) = New ValueWrapper(Of Item)(Nothing)
            _attributes.Add(New RibbonAttributeWrappedValue(Of Item)(wrapped, getSelected.Method.Name, AttributeName.GetSelectedItemIndex, AttributeCategory.SelectedItemPosition))
            _attributes.Add(New RibbonAttributeWrappedValue(Of Item)(wrapped, setSelected.Method.Name, AttributeName.OnAction, AttributeCategory.OnAction))
        End Sub

#End Region

#Region "Item Count"

        Protected Sub ItemCountBase(callback As FromControl(Of Integer))
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(callback.Method.Name, AttributeName.GetItemCount, AttributeCategory.ItemCount))
        End Sub

#End Region

#Region "Item Id"

        Protected Sub ItemIDBase(callback As FromControlAndIndex(Of String))
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(callback.Method.Name, AttributeName.GetItemID, AttributeCategory.ItemId))
        End Sub

#End Region

#Region "Item Label"

        Protected Sub ItemLabelBase(callback As FromControlAndIndex(Of String))
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(callback.Method.Name, AttributeName.GetItemLabel, AttributeCategory.ItemLabel))
        End Sub

#End Region

#Region "Show/Hide Item Label"

        Protected Sub ShowItemLabelBase()
            _attributes.Add(New RibbonAttributeReadOnly(Of Boolean)(True, AttributeName.ShowItemLabel, AttributeCategory.ItemLabelVisibility))
        End Sub

        Protected Sub ShowItemLabelBase(callback As FromControl(Of Boolean))
            _attributes.Add(New RibbonAttributeReadWrite(Of Boolean)(True, AttributeName.ShowItemLabel, AttributeCategory.ItemLabelVisibility, callback.Method.Name))
        End Sub

        Protected Sub HideItemLabelBase()
            _attributes.Add(New RibbonAttributeReadOnly(Of Boolean)(False, AttributeName.ShowItemLabel, AttributeCategory.ItemLabelVisibility))
        End Sub

        Protected Sub HideItemLabelBase(callback As FromControl(Of Boolean))
            _attributes.Add(New RibbonAttributeReadWrite(Of Boolean)(False, AttributeName.ShowItemLabel, AttributeCategory.ItemLabelVisibility, callback.Method.Name))
        End Sub

#End Region

#Region "Item Screentip"

        Protected Sub ItemScreentipBase(callback As FromControlAndIndex(Of String))
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(callback.Method.Name, AttributeName.GetItemScreentip, AttributeCategory.ItemScreentip))
        End Sub

#End Region

#Region "Item Supertip"

        Protected Sub ItemSupertipBase(callback As FromControlAndIndex(Of String))
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(callback.Method.Name, AttributeName.GetItemSupertip, AttributeCategory.ItemSupertip))
        End Sub

#End Region

#Region "Item Image"

        Protected Sub ItemImageBase(callback As FromControlAndIndex(Of IPictureDisp))
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(callback.Method.Name, AttributeName.GetItemImage, AttributeCategory.ItemImage))
        End Sub

#End Region

#Region "Show/Hide Item Image"

        Protected Sub ShowItemImageBase()
            _attributes.Add(New RibbonAttributeReadOnly(Of Boolean)(True, AttributeName.ShowItemImage, AttributeCategory.ItemImageVisibility))
        End Sub

        Protected Sub HideItemImageBase()
            _attributes.Add(New RibbonAttributeReadOnly(Of Boolean)(False, AttributeName.ShowItemImage, AttributeCategory.ItemImageVisibility))
        End Sub

#End Region

#Region "Item Height/Width"

        Protected Sub ItemHeightBase(height As Integer)
            _attributes.Add(New RibbonAttributeReadOnly(Of Integer)(height, AttributeName.ItemHeight, AttributeCategory.ItemHeight))
        End Sub

        Protected Sub ItemHeightBase(height As Integer, callback As FromControl(Of Integer))
            _attributes.Add(New RibbonAttributeReadWrite(Of Integer)(height, AttributeName.GetItemHeight, AttributeCategory.ItemHeight, callback.Method.Name))
        End Sub

        Protected Sub ItemWidthBase(width As Integer)
            _attributes.Add(New RibbonAttributeReadOnly(Of Integer)(width, AttributeName.ItemWidth, AttributeCategory.ItemWidth))
        End Sub

        Protected Sub ItemWidthBase(width As Integer, callback As FromControl(Of Integer))
            _attributes.Add(New RibbonAttributeReadWrite(Of Integer)(width, AttributeName.GetItemWidth, AttributeCategory.ItemWidth, callback.Method.Name))
        End Sub

#End Region

#Region "Row/Column Count"

        Protected Sub RowCountBase(count As Integer)
            _attributes.Add(New RibbonAttributeReadOnly(Of Integer)(count, AttributeName.Rows, AttributeCategory.Rows))
        End Sub

        Protected Sub ColumnCountBase(count As Integer)
            _attributes.Add(New RibbonAttributeReadOnly(Of Integer)(count, AttributeName.Columns, AttributeCategory.Columns))
        End Sub

#End Region

#Region "Large/Normal Items"

        Protected Sub LargeItemsBase()
            _attributes.Add(New RibbonAttributeReadOnly(Of ControlSize)(ControlSize.Large, AttributeName.ItemSize, AttributeCategory.ItemSize))
        End Sub

        Protected Sub NormalItemsBase()
            _attributes.Add(New RibbonAttributeReadOnly(Of ControlSize)(ControlSize.Normal, AttributeName.ItemSize, AttributeCategory.ItemSize))
        End Sub

#End Region

    End Class

End Namespace