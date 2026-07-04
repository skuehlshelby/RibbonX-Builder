Imports System.Drawing
Imports RibbonX.Api.Internal
Imports RibbonX.Callbacks
Imports RibbonX.ComTypes.StdOle
Imports RibbonX.Controls.Attributes
Imports RibbonX.Controls.Base
Imports RibbonX.Controls.BuiltIn
Imports RibbonX.Images
Imports RibbonX.Images.BuiltIn
Imports RibbonX.Properties
Imports RibbonX.SimpleTypes

Namespace InternalApi
    Friend MustInherit Class BuilderBase(Of TElement As RibbonElement)
        Protected ReadOnly ControlProperties As IPropertyCollection = GetDefaults()
        Protected Tag As Object = Nothing

        Public MustOverride Function Build() As TElement

#Region "Defaults"
        Private Shared Function GetNextElementId() As Integer
            Static id As Integer = 0
            id = If(id = Integer.MaxValue, 0, id + 1)
            Return id
        End Function

        Private Shared Function GetDefaults() As IPropertyCollection
            Dim elementType As Type = GetType(TElement)

            Dim defaults As IPropertyCollection = New PropertyCollection(RibbonProperty.Create(Name.Id, Category.IdType, $"{elementType.Name}{GetNextElementId()}"))

            Dim implementedInterfaces As Type() = elementType.GetInterfaces()

            For Each interfaceType As Type In implementedInterfaces
                Select Case interfaceType
                    Case GetType(IEnabled)
                        defaults.Add(RibbonProperty.Create(Name.Enabled, Category.Enabled, True, True))
                    Case GetType(IVisible)
                        defaults.Add(RibbonProperty.Create(Name.Visible, Category.Visibility, True, True))
                    Case GetType(ILabel)
                        defaults.Add(RibbonProperty.Create(Name.Label, Category.Label, String.Empty, True))
                    Case GetType(IShowLabel)
                        defaults.Add(RibbonProperty.Create(Name.ShowLabel, Category.LabelVisibility, False, True))
                    Case GetType(IScreenTip)
                        defaults.Add(RibbonProperty.Create(Name.Screentip, Category.ScreenTip, String.Empty, True))
                    Case GetType(ISuperTip)
                        defaults.Add(RibbonProperty.Create(Name.Supertip, Category.SuperTip, String.Empty, True))
                    Case GetType(ISize)
                        defaults.Add(RibbonProperty.Create(Name.Size, Category.Size, ControlSize.Normal, True))
                    Case GetType(IDescription)
                        defaults.Add(RibbonProperty.Create(Name.Description, Category.Description, String.Empty, True))
                    Case GetType(IKeyTip)
                        defaults.Add(RibbonProperty.Create(Name.Keytip, Category.KeyTip, CType("A", KeyTip), True))
                    Case GetType(IImage)
                        defaults.Add(RibbonProperty.Create(Name.ImageMso, Category.Image, RibbonImage.Create(Common.HappyFace), True))
                    Case GetType(IShowImage)
                        defaults.Add(RibbonProperty.Create(Name.ShowImage, Category.ImageVisibility, False, True))
                    Case GetType(ITitle)
                        defaults.Add(RibbonProperty.Create(Name.Title, Category.Title, String.Empty, True))
                    Case GetType(IClickable)
                        defaults.Add(RibbonProperty.Create(Name.OnAction, Category.OnAction, Sub() Return, True))
                    Case GetType(IText)
                        defaults.Add(RibbonProperty.Create(Name.GetText, Category.Text, String.Empty, True))
                        defaults.Add(RibbonProperty.Create(Name.SizeString, Category.SizeString, "WWWWWW", True))
                        defaults.Add(RibbonProperty.Create(Name.MaxLength, Category.MaxLength, Byte.MaxValue, True))
                    Case GetType(IChecked)
                        defaults.Add(RibbonProperty.Create(Name.GetPressed, Category.Pressed, False, True))
                    Case GetType(ISelect)
                        defaults.Add(RibbonProperty.Create(Of IItem)(Name.GetSelectedItemIndex, Category.SelectedItemPosition, Nothing, True))
                    Case GetType(IItemDimensions)
                        defaults.Add(RibbonProperty.Create(Name.ItemHeight, Category.ItemHeight, 1, True))
                        defaults.Add(RibbonProperty.Create(Name.ItemWidth, Category.ItemWidth, 1, True))
                    Case GetType(IItemSize)
                        defaults.Add(RibbonProperty.Create(Name.ItemSize, Category.ItemSize, ControlSize.Normal, True))
                    Case GetType(IBoxStyle)
                        defaults.Add(RibbonProperty.Create(Name.BoxStyle, Category.BoxStyle, BoxStyle.Horizontal, True))
                    Case GetType(IRowsAndColumns)
                        defaults.Add(RibbonProperty.Create(Name.Rows, Category.Rows, 0, True))
                        defaults.Add(RibbonProperty.Create(Name.Columns, Category.Columns, 0, True))
                End Select
            Next interfaceType

            Return defaults
        End Function

#End Region

#Region "Tag"

        Protected Sub WithTagBase(tag As Object)
            Me.Tag = tag
        End Sub

#End Region

#Region "Templating"
        Protected Sub FromTemplateBase(template As IRibbonElement)
            If template IsNot Nothing AndAlso TypeOf template Is IAttributeSource Then
                ControlProperties.Merge(DirectCast(template, IAttributeSource).GetAttributes())
            End If
        End Sub
#End Region

#Region "Enabled"

        Protected Sub EnabledBase()
            ControlProperties.Add(RibbonProperty.Create(Name.Enabled, Category.Enabled, True))
        End Sub

        Protected Sub EnabledBase(callback As FromControl(Of Boolean))
            ControlProperties.Add(RibbonProperty.Create(Name.GetEnabled, Category.Enabled, True, callback.Method.Name))
        End Sub

#End Region

#Region "Disabled"

        Protected Sub DisabledBase()
            ControlProperties.Add(RibbonProperty.Create(Name.Enabled, Category.Enabled, False))
        End Sub

        Protected Sub DisabledBase(callback As FromControl(Of Boolean))
            ControlProperties.Add(RibbonProperty.Create(Name.GetEnabled, Category.Enabled, False, callback.Method.Name))
        End Sub

#End Region

#Region "Visible"
        Protected Sub VisibleBase()
            ControlProperties.Add(RibbonProperty.Create(Name.Visible, Category.Visibility, True))
        End Sub

        Protected Sub VisibleBase(callback As FromControl(Of Boolean))
            ControlProperties.Add(RibbonProperty.Create(Name.GetVisible, Category.Visibility, True, callback.Method.Name))
        End Sub
#End Region

#Region "Invisible"
        Protected Sub InvisibleBase()
            ControlProperties.Add(RibbonProperty.Create(Name.Visible, Category.Visibility, False))
        End Sub

        Protected Sub InvisibleBase(callback As FromControl(Of Boolean))
            ControlProperties.Add(RibbonProperty.Create(Name.GetVisible, Category.Visibility, False, callback.Method.Name))
        End Sub
#End Region

#Region "Id"

        Protected Sub WithIdBase(id As String)
            ControlProperties.Add(RibbonProperty.Create(Name.Id, Category.IdType, id))
        End Sub

        Protected Sub WithIdBase([namespace] As String, id As String)
            ControlProperties.Add(RibbonProperty.Create(Name.IdQ, Category.IdType, $"{[namespace]}:{id}"))
        End Sub

        Protected Sub WithIdBase(mso As MSO)
            ControlProperties.Add(RibbonProperty.Create(Name.IdMso, Category.IdType, mso.ToString()))
        End Sub

#End Region

#Region "Insert Before/After"

        Protected Sub InsertAfterMsoBase(mso As MSO)
            ControlProperties.Add(RibbonProperty.Create(Name.InsertAfterMso, Category.Insertion, mso.ToString()))
        End Sub

        Protected Sub InsertAfterQBase(control As IRibbonElement)
            ControlProperties.Add(RibbonProperty.Create(Name.InsertAfterQ, Category.Insertion, control.Id))
        End Sub

        Protected Sub InsertBeforeMsoBase(mso As MSO)
            ControlProperties.Add(RibbonProperty.Create(Name.InsertBeforeMso, Category.Insertion, mso.ToString()))
        End Sub

        Protected Sub InsertBeforeQBase(control As IRibbonElement)
            ControlProperties.Add(RibbonProperty.Create(Name.InsertBeforeQ, Category.Insertion, control.Id))
        End Sub

#End Region

#Region "Label"

        Protected Sub LabelBase(label As String)
            ControlProperties.Add(RibbonProperty.Create(Name.Label, Category.Label, label))
        End Sub

        Protected Sub LabelBase(label As String, callback As FromControl(Of String))
            ControlProperties.Add(RibbonProperty.Create(Name.GetLabel, Category.Label, label, callback.Method.Name))
        End Sub

#End Region

#Region "Show Label"

        Protected Sub HideLabelBase()
            ControlProperties.Add(RibbonProperty.Create(Name.ShowLabel, Category.LabelVisibility, False))
        End Sub

        Protected Sub HideLabelBase(callback As FromControl(Of Boolean))
            ControlProperties.Add(RibbonProperty.Create(Name.GetShowLabel, Category.LabelVisibility, False, callback.Method.Name))
        End Sub

        Protected Sub ShowLabelBase()
            ControlProperties.Add(RibbonProperty.Create(Name.ShowLabel, Category.LabelVisibility, True))
        End Sub

        Protected Sub ShowLabelBase(callback As FromControl(Of Boolean))
            ControlProperties.Add(RibbonProperty.Create(Name.GetShowLabel, Category.LabelVisibility, True, callback.Method.Name))
        End Sub

#End Region

#Region "Description"

        Protected Sub DescriptionBase(description As String)
            ControlProperties.Add(RibbonProperty.Create(Name.Description, Category.Description, description))
        End Sub

        Protected Sub DescriptionBase(description As String, callback As FromControl(Of String))
            ControlProperties.Add(RibbonProperty.Create(Name.GetDescription, Category.Description, description, callback.Method.Name))
        End Sub

#End Region

#Region "KeyTip"

        Protected Sub KeyTipBase(keyTip As KeyTip)
            ControlProperties.Add(RibbonProperty.Create(Name.Keytip, Category.KeyTip, keyTip))
        End Sub

        Protected Sub KeyTipBase(keyTip As KeyTip, callback As FromControl(Of KeyTip))
            ControlProperties.Add(RibbonProperty.Create(Name.GetKeytip, Category.KeyTip, keyTip, callback.Method.Name))
        End Sub

#End Region

#Region "Screentip"

        Protected Sub ScreentipBase(screentip As String)
            ControlProperties.Add(RibbonProperty.Create(Name.Screentip, Category.ScreenTip, screentip))
        End Sub

        Protected Sub ScreentipBase(screentip As String, callback As FromControl(Of String))
            ControlProperties.Add(RibbonProperty.Create(Name.GetScreentip, Category.ScreenTip, screentip, callback.Method.Name))
        End Sub

#End Region

#Region "Supertip"

        Protected Sub SupertipBase(supertip As String)
            ControlProperties.Add(RibbonProperty.Create(Name.Supertip, Category.SuperTip, supertip))
        End Sub

        Protected Sub SupertipBase(supertip As String, callback As FromControl(Of String))
            ControlProperties.Add(RibbonProperty.Create(Name.GetSupertip, Category.SuperTip, supertip, callback.Method.Name))
        End Sub

#End Region

#Region "Title"

        Protected Sub TitleBase(title As String)
            ControlProperties.Add(RibbonProperty.Create(Name.Title, Category.Title, title))
        End Sub

        Protected Sub TitleBase(title As String, callback As FromControl(Of String))
            ControlProperties.Add(RibbonProperty.Create(Name.GetTitle, Category.Title, callback.Method.Name, title))
        End Sub

#End Region

#Region "Image"

        Protected Sub ImageBase(image As ImageMSO)
            ControlProperties.Add(RibbonProperty.Create(Name.ImageMso, Category.Image, RibbonImage.Create(image)))
        End Sub

        Protected Sub ImageBase(image As ImageMSO, callback As FromControl(Of String))
            ControlProperties.Add(RibbonProperty.Create(Name.GetImage, Category.Image, RibbonImage.Create(image), callback.Method.Name))
        End Sub

        Protected Sub ImageBase(imageId As String, bitmap As Bitmap)
            ControlProperties.Add(RibbonProperty.Create(Name.Image, Category.Image, RibbonImage.Create(imageId, bitmap)))
        End Sub

        Protected Sub ImageBase(imageId As String, icon As Icon)
            ControlProperties.Add(RibbonProperty.Create(Name.Image, Category.Image, RibbonImage.Create(imageId, icon)))
        End Sub

        Protected Sub ImageBase(image As IPictureDisp, callback As FromControl(Of IPictureDisp))
            ControlProperties.Add(RibbonProperty.Create(Name.GetImage, Category.Image, RibbonImage.Create(image), callback.Method.Name))
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
            ControlProperties.Add(RibbonProperty.Create(Name.ShowImage, Category.ImageVisibility, True))
        End Sub

        Protected Sub ShowImageBase(callback As FromControl(Of Boolean))
            ControlProperties.Add(RibbonProperty.Create(Name.GetShowImage, Category.ImageVisibility, True, callback.Method.Name))
        End Sub

#End Region

#Region "Hide Image"

        Protected Sub HideImageBase()
            ControlProperties.Add(RibbonProperty.Create(Name.ShowImage, Category.ImageVisibility, False))
        End Sub

        Protected Sub HideImageBase(callback As FromControl(Of Boolean))
            ControlProperties.Add(RibbonProperty.Create(Name.GetShowImage, Category.ImageVisibility, False, callback.Method.Name))
        End Sub

#End Region

#Region "Normal"

        Protected Sub NormalBase()
            ControlProperties.Add(RibbonProperty.Create(Name.Size, Category.Size, ControlSize.Normal))
        End Sub

        Protected Sub NormalBase(callback As FromControl(Of ControlSize))
            ControlProperties.Add(RibbonProperty.Create(Name.GetSize, Category.Size, ControlSize.Normal, callback.Method.Name))
        End Sub

#End Region

#Region "Large"

        Protected Sub LargeBase()
            ControlProperties.Add(RibbonProperty.Create(Name.Size, Category.Size, ControlSize.Large))
        End Sub

        Protected Sub LargeBase(callback As FromControl(Of ControlSize))
            ControlProperties.Add(RibbonProperty.Create(Name.GetSize, Category.Size, ControlSize.Large, callback.Method.Name))
        End Sub

#End Region

#Region "Size String"

        Protected Sub SizeStringBase(sizeString As String)
            ControlProperties.Add(RibbonProperty.Create(Name.SizeString, Category.SizeString, sizeString))
        End Sub

#End Region

#Region "Max Length"

        Protected Sub MaxLengthBase(maxLength As Byte)
            ControlProperties.Add(RibbonProperty.Create(Name.MaxLength, Category.MaxLength, maxLength))
        End Sub

#End Region

#Region "Horizontal/Vertical"

        Protected Sub HorizontalBase()
            ControlProperties.Add(RibbonProperty.Create(Name.BoxStyle, Category.BoxStyle, BoxStyle.Horizontal))
        End Sub

        Protected Sub VerticalBase()
            ControlProperties.Add(RibbonProperty.Create(Name.BoxStyle, Category.BoxStyle, BoxStyle.Vertical))
        End Sub

#End Region

#Region "On Action"

        Protected Sub OnActionBase(callback As OnAction)
            ControlProperties.Add(RibbonProperty.Create(Name.OnAction, Category.OnAction, callback.Method.Name))
        End Sub

        Protected Sub OnActionBase(callback As SelectionChanged)
            ControlProperties.Add(RibbonProperty.Create(Name.OnAction, Category.OnAction, callback.Method.Name))
        End Sub

#End Region

#Region "Get/Set Text"

        Protected Sub TextBase(text As String, getText As FromControl(Of String), setText As TextChanged)
            ControlProperties.Add(RibbonProperty.Create(Name.GetText, Category.Text, text, getText.Method.Name))
            ControlProperties.Add(RibbonProperty.Create(Name.OnChange, Category.OnChange, text, setText.Method.Name))
            ControlProperties.Forward(Category.OnChange, Category.Text)
        End Sub

#End Region

#Region "Get/Set Pressed"

        Protected Sub PressedBase(pressed As Boolean, getPressed As FromControl(Of Boolean), setPressed As ButtonPressed)
            ControlProperties.Add(RibbonProperty.Create(Name.GetPressed, Category.Pressed, pressed, getPressed.Method.Name))
            ControlProperties.Add(RibbonProperty.Create(Name.OnAction, Category.OnAction, pressed, setPressed.Method.Name))
            ControlProperties.Forward(Category.OnAction, Category.Pressed)
        End Sub

#End Region

#Region "Add Handlers"

        Protected Sub OnActionBase(Of T As IRibbonElement)(action As Action(Of IActionBuilder(Of T)), before As String, after As String)
            Dim builder As IActionBuilder(Of T) = New ActionBuilder(Of T, Boolean)(ControlProperties, before, after)
            Utilities.NotNull(action).Invoke(builder)
        End Sub

        Protected Sub OnActionBase(Of T As IRibbonElement, K)(action As Action(Of IActionBuilder(Of T, K)), before As String, after As String)
            Dim builder As IActionBuilder(Of T, K) = New ActionBuilder(Of T, K)(ControlProperties, before, after)
            Utilities.NotNull(action).Invoke(builder)
        End Sub

        Private NotInheritable Class ActionBuilder(Of TRibbonElement As IRibbonElement, TData)
            Implements IActionBuilder(Of TRibbonElement)
            Implements IActionBuilder(Of TRibbonElement, TData)

            Private ReadOnly controlProperties As IPropertyCollection
            Private ReadOnly before As String
            Private ReadOnly after As String

            Public Sub New(controlProperties As IPropertyCollection, before As String, after As String)
                Me.controlProperties = controlProperties
                Me.before = before
                Me.after = after
            End Sub

            Public Function [Do](ParamArray actions() As Action) As IActionBuilder(Of TRibbonElement) Implements IActionBuilder(Of TRibbonElement).Do
                For Each a As Action In actions
                    controlProperties.AddHandler(after, New EventHandler(Sub(sender, e) a.Invoke()))
                Next
                Return Me
            End Function

            Public Function [Do](ParamArray actions() As Action(Of TRibbonElement)) As IActionBuilder(Of TRibbonElement) Implements IActionBuilder(Of TRibbonElement).Do
                For Each a As Action(Of TRibbonElement) In actions
                    controlProperties.AddHandler(after, New EventHandler(Sub(sender, e) a.Invoke(DirectCast(sender, TRibbonElement))))
                Next
                Return Me
            End Function

            Public Function [Do](ParamArray actions() As Action(Of TData)) As IActionBuilder(Of TRibbonElement) Implements IActionBuilder(Of TRibbonElement, TData).Do
                For Each a As Action(Of TData) In actions
                    controlProperties.AddHandler(after, New EventHandler(Of EventArgs(Of TData))(Sub(sender, e) a.Invoke(e.Current)))
                Next
                Return Me
            End Function

            Public Function [Do](ParamArray actions() As Action(Of TRibbonElement, TData)) As IActionBuilder(Of TRibbonElement) Implements IActionBuilder(Of TRibbonElement, TData).Do
                For Each a As Action(Of TRibbonElement, TData) In actions
                    controlProperties.AddHandler(after, New EventHandler(Of EventArgs(Of TData))(Sub(sender, e) a.Invoke(DirectCast(sender, TRibbonElement), e.Current)))
                Next
                Return Me
            End Function

            Public Function ButFirst(ParamArray actions() As Action) As IActionBuilder(Of TRibbonElement) Implements IActionBuilder(Of TRibbonElement).ButFirst
                For Each a As Action In actions
                    controlProperties.AddHandler(NameOf(before), New EventHandler(Of EventArgs)(Sub(sender, e) a.Invoke()))
                Next
                Return Me
            End Function

            Public Function ButFirst(ParamArray actions() As Func(Of Boolean)) As IActionBuilder(Of TRibbonElement) Implements IActionBuilder(Of TRibbonElement).ButFirst
                For Each a As Func(Of Boolean) In actions
                    controlProperties.AddHandler(NameOf(before), New EventHandler(Of CancelableEventArgs)(Sub(sender, e) If a.Invoke() Then e.Cancel()))
                Next
                Return Me
            End Function

            Public Function ButFirst(ParamArray actions() As Func(Of TRibbonElement, Boolean)) As IActionBuilder(Of TRibbonElement) Implements IActionBuilder(Of TRibbonElement).ButFirst
                For Each a As Func(Of TRibbonElement, Boolean) In actions
                    controlProperties.AddHandler(NameOf(before), New EventHandler(Of CancelableEventArgs)(Sub(sender, e) If a.Invoke(DirectCast(sender, TRibbonElement)) Then e.Cancel()))
                Next
                Return Me
            End Function

            Public Function ButFirst(ParamArray actions() As Func(Of TData, Boolean)) As IActionBuilder(Of TRibbonElement) Implements IActionBuilder(Of TRibbonElement, TData).ButFirst
                For Each a As Func(Of TData, Boolean) In actions
                    controlProperties.AddHandler(NameOf(before), New EventHandler(Of CancelableEventArgs(Of TData))(Sub(sender, e) If a.Invoke(e.Current) Then e.Cancel()))
                Next
                Return Me
            End Function

            Public Function ButFirst(ParamArray actions() As Func(Of TRibbonElement, TData, Boolean)) As IActionBuilder(Of TRibbonElement) Implements IActionBuilder(Of TRibbonElement, TData).ButFirst
                For Each a As Func(Of TRibbonElement, TData, Boolean) In actions
                    controlProperties.AddHandler(NameOf(before), New EventHandler(Of CancelableEventArgs(Of TData))(Sub(sender, e) If a.Invoke(DirectCast(sender, TRibbonElement), e.Current) Then e.Cancel()))
                Next
                Return Me
            End Function
        End Class

#End Region

#Region "Invalidate Content On Drop"

        Protected Sub InvalidateContentOnDropBase()
            ControlProperties.Add(RibbonProperty.Create(Name.InvalidateContentOnDrop, Category.ContentInvalidation, True))
        End Sub

#End Region

#Region "Do Not Show In Ribbon"
        Protected Sub DoNotShowInRibbonBase()
            ControlProperties.Add(RibbonProperty.Create(Name.ShowInRibbon, Category.ShowInRibbon, False))
        End Sub

#End Region

        '' Child Item Methods ''

#Region "Selected Item"

        Protected Sub SelectedBase(getSelected As FromControl(Of String), setSelected As SelectionChanged)
            ControlProperties.Add(RibbonProperty.Create(Of IItem)(Name.GetSelectedItemID, Category.SelectedItemPosition, Nothing, getSelected.Method.Name))
            ControlProperties.Add(RibbonProperty.Create(Of IItem)(Name.OnAction, Category.OnAction, Nothing, setSelected.Method.Name))
            ControlProperties.Forward(Category.OnAction, Category.SelectedItemPosition)
        End Sub

        Protected Sub SelectedBase(getSelected As FromControl(Of Integer), setSelected As SelectionChanged)
            ControlProperties.Add(RibbonProperty.Create(Of IItem)(Name.GetSelectedItemIndex, Category.SelectedItemPosition, Nothing, getSelected.Method.Name))
            ControlProperties.Add(RibbonProperty.Create(Of IItem)(Name.OnAction, Category.OnAction, Nothing, setSelected.Method.Name))
            ControlProperties.Forward(Category.OnAction, Category.SelectedItemPosition)
        End Sub

#End Region

#Region "Item Count"

        Protected Sub ItemCountBase(callback As FromControl(Of Integer))
            ControlProperties.Add(RibbonProperty.Create(Name.GetItemCount, Category.ItemCount, callback.Method.Name))
        End Sub

#End Region

#Region "Item Id"

        Protected Sub ItemIDBase(callback As FromControlAndIndex(Of String))
            ControlProperties.Add(RibbonProperty.Create(Name.GetItemID, Category.ItemId, callback.Method.Name))
        End Sub

#End Region

#Region "Item Label"

        Protected Sub ItemLabelBase(callback As FromControlAndIndex(Of String))
            ControlProperties.Add(RibbonProperty.Create(Name.GetItemLabel, Category.ItemLabel, callback.Method.Name))
        End Sub

#End Region

#Region "Show/Hide Item Label"

        Protected Sub ShowItemLabelBase()
            ControlProperties.Add(RibbonProperty.Create(Name.ShowItemLabel, Category.ItemLabelVisibility, True))
        End Sub

        Protected Sub ShowItemLabelBase(callback As FromControl(Of Boolean))
            ControlProperties.Add(RibbonProperty.Create(Name.ShowItemLabel, Category.ItemLabelVisibility, True, callback.Method.Name))
        End Sub

        Protected Sub HideItemLabelBase()
            ControlProperties.Add(RibbonProperty.Create(Name.ShowItemLabel, Category.ItemLabelVisibility, False))
        End Sub

        Protected Sub HideItemLabelBase(callback As FromControl(Of Boolean))
            ControlProperties.Add(RibbonProperty.Create(Name.ShowItemLabel, Category.ItemLabelVisibility, False, callback.Method.Name))
        End Sub

#End Region

#Region "Item Screentip"

        Protected Sub ItemScreentipBase(callback As FromControlAndIndex(Of String))
            ControlProperties.Add(RibbonProperty.Create(Name.GetItemScreentip, Category.ItemScreentip, callback.Method.Name))
        End Sub

#End Region

#Region "Item Supertip"

        Protected Sub ItemSupertipBase(callback As FromControlAndIndex(Of String))
            ControlProperties.Add(RibbonProperty.Create(Name.GetItemSupertip, Category.ItemSupertip, callback.Method.Name))
        End Sub

#End Region

#Region "Item Image"

        Protected Sub ItemImageBase(callback As FromControlAndIndex(Of IPictureDisp))
            ControlProperties.Add(RibbonProperty.Create(Name.GetItemImage, Category.ItemImage, callback.Method.Name))
        End Sub

        Protected Sub ItemImageBase(callback As FromControlAndIndex(Of String))
            ControlProperties.Add(RibbonProperty.Create(Name.GetItemImage, Category.ItemImage, callback.Method.Name))
        End Sub

#End Region

#Region "Show/Hide Item Image"

        Protected Sub ShowItemImageBase()
            ControlProperties.Add(RibbonProperty.Create(Name.ShowItemImage, Category.ItemImageVisibility, True))
        End Sub

        Protected Sub HideItemImageBase()
            ControlProperties.Add(RibbonProperty.Create(Name.ShowItemImage, Category.ItemImageVisibility, False))
        End Sub

#End Region

#Region "Item Height/Width"

        Protected Sub ItemHeightBase(height As Integer)
            ControlProperties.Add(RibbonProperty.Create(Name.ItemHeight, Category.ItemHeight, height))
        End Sub

        Protected Sub ItemHeightBase(height As Integer, callback As FromControl(Of Integer))
            ControlProperties.Add(RibbonProperty.Create(Name.GetItemHeight, Category.ItemHeight, height, callback.Method.Name))
        End Sub

        Protected Sub ItemWidthBase(width As Integer)
            ControlProperties.Add(RibbonProperty.Create(Name.ItemWidth, Category.ItemWidth, width))
        End Sub

        Protected Sub ItemWidthBase(width As Integer, callback As FromControl(Of Integer))
            ControlProperties.Add(RibbonProperty.Create(Name.GetItemWidth, Category.ItemWidth, width, callback.Method.Name))
        End Sub

#End Region

#Region "Row/Column Count"

        Protected Sub RowCountBase(count As Integer)
            ControlProperties.Add(RibbonProperty.Create(Name.Rows, Category.Rows, count))
        End Sub

        Protected Sub ColumnCountBase(count As Integer)
            ControlProperties.Add(RibbonProperty.Create(Name.Columns, Category.Columns, count))
        End Sub

#End Region

#Region "Large/Normal Items"

        Protected Sub LargeItemsBase()
            ControlProperties.Add(RibbonProperty.Create(Name.ItemSize, Category.ItemSize, ControlSize.Large))
        End Sub

        Protected Sub NormalItemsBase()
            ControlProperties.Add(RibbonProperty.Create(Name.ItemSize, Category.ItemSize, ControlSize.Normal))
        End Sub

#End Region

    End Class

End Namespace