Imports RibbonFactory.Controls
Imports RibbonFactory.Enums
Imports RibbonFactory.Enums.ImageMSO
Imports RibbonFactory.Enums.MSO
Imports stdole

Namespace RibbonAttributes

	''' <summary>
	''' The purpose of this class is to provide a central place where attributes are configured. It does not attempt
	''' to compose attributes in meaningful ways, only to configure individual attributes.
	''' </summary>
	Friend Class AttributeGroupBuilder
		Implements IEnumerable(Of IRibbonAttribute)

		Private ReadOnly _attributes As AttributeSet = New AttributeSet()

		Public Sub New()
			_attributes = New AttributeSet()
		End Sub

		Public Sub New(attributes As IEnumerable(Of IRibbonAttribute))
			_attributes = New AttributeSet(attributes)
		End Sub

		Public Sub SetDefaults(defaultProvider As IDefaultProvider)
			If _attributes.Any() Then
				Throw New InvalidOperationException("Attempted to set defaults when one or more attributes were already set.")
			End If

			For Each attribute As IRibbonAttribute In defaultProvider.GetDefaults()
				_attributes.Add(attribute)
			Next
		End Sub

		Public Function Build() As AttributeSet
			Return _attributes
		End Function

#Region "ID, IDMSO, and IDQ"

		Public Sub AddID(id As String)
			_attributes.Add(New RibbonAttributeReadOnly(Of String)(id, AttributeName.Id, AttributeCategory.IdType))
		End Sub

		Public Sub AddIdMso(idMso As MSO)
			_attributes.Clear()
			_attributes.Add(New RibbonAttributeReadOnly(Of String)(idMso.ToString(), AttributeName.IdMso, AttributeCategory.IdType))
		End Sub

		Public Sub AddIdQ([namespace] As String, id As String)
			_attributes.Add(New RibbonAttributeReadOnly(Of String)($"{[namespace]}:{id}", AttributeName.IdQ, AttributeCategory.IdType))
		End Sub

#End Region

#Region "InsetAfterMSO, InsertBeforeMSO, InsertAfterQ, and InsertBeforeQ"

		Public Sub AddInsertAfterMso(mso As MSO)
			_attributes.Add(New RibbonAttributeReadOnly(Of String)(mso.ToString(), AttributeName.InsertAfterMso, AttributeCategory.Insertion))
		End Sub

		Public Sub AddInsertAfterQ(control As RibbonElement)
			_attributes.Add(New RibbonAttributeReadOnly(Of String)(control.ID, AttributeName.InsertAfterQ, AttributeCategory.Insertion))
		End Sub

		Public Sub AddInsertBeforeMso(mso As MSO)
			_attributes.Add(New RibbonAttributeReadOnly(Of String)(mso.ToString(), AttributeName.InsertBeforeMso, AttributeCategory.Insertion))
		End Sub

		Public Sub AddInsertBeforeQ(control As RibbonElement)
			_attributes.Add(New RibbonAttributeReadOnly(Of String)(control.ID, AttributeName.InsertBeforeQ, AttributeCategory.Insertion))
		End Sub

#End Region

#Region "Enabled and GetEnabled"

		Public Sub AddEnabled(enabled As Boolean)
			_attributes.Add(New RibbonAttributeReadOnly(Of Boolean)(enabled, AttributeName.Enabled, AttributeCategory.Enabled))
		End Sub

		Public Sub AddGetEnabled(enabled As Boolean, callback As FromControl(Of Boolean))
			_attributes.Add(New RibbonAttributeReadWrite(Of Boolean)(enabled, AttributeName.GetEnabled, AttributeCategory.Enabled, callback.Method.Name))
		End Sub

#End Region

#Region "Visible and GetVisible"

		Public Sub AddVisible(visible As Boolean)
			_attributes.Add(New RibbonAttributeReadOnly(Of Boolean)(visible, AttributeName.Visible, AttributeCategory.Visibility))
		End Sub

		Public Sub AddGetVisible(visible As Boolean, callback As FromControl(Of Boolean))
			_attributes.Add(New RibbonAttributeReadWrite(Of Boolean)(visible, AttributeName.GetVisible, AttributeCategory.Visibility, callback.Method.Name))
		End Sub

#End Region

#Region "Label and GetLabel"

		Public Sub AddLabel(label As String)
			_attributes.Add(New RibbonAttributeReadOnly(Of String)(label, AttributeName.Label, AttributeCategory.Label))
		End Sub

		Public Sub AddGetLabel(label As String, callback As FromControl(Of String))
			_attributes.Add(New RibbonAttributeReadWrite(Of String)(label, AttributeName.GetLabel, AttributeCategory.Label, callback.Method.Name))
		End Sub

#End Region

#Region "ShowLabel and GetShowLabel"

		Public Sub AddShowLabel(showLabel As Boolean)
			_attributes.Add(New RibbonAttributeReadOnly(Of Boolean)(showLabel, AttributeName.ShowLabel, AttributeCategory.LabelVisibility))
		End Sub

		Public Sub AddGetShowLabel(showLabel As Boolean, callback As FromControl(Of Boolean))
			_attributes.Add(New RibbonAttributeReadWrite(Of Boolean)(showLabel, AttributeName.GetShowLabel, AttributeCategory.LabelVisibility, callback.Method.Name))
		End Sub

#End Region

#Region "Screentip and GetScreentip"

		Public Sub AddScreentip(screentip As String)
			_attributes.Add(New RibbonAttributeReadOnly(Of String)(screentip, AttributeName.Screentip, AttributeCategory.ScreenTip))
		End Sub

		Public Sub AddGetScreentip(screentip As String, callback As FromControl(Of String))
			_attributes.Add(New RibbonAttributeReadWrite(Of String)(screentip, AttributeName.GetScreentip, AttributeCategory.ScreenTip, callback.Method.Name))
		End Sub

#End Region

#Region "Supertip and GetSupertip"

		Public Sub AddSupertip(supertip As String)
			_attributes.Add(New RibbonAttributeReadOnly(Of String)(supertip, AttributeName.Supertip, AttributeCategory.SuperTip))
		End Sub

		Public Sub AddGetSupertip(supertip As String, callback As FromControl(Of String))
			_attributes.Add(New RibbonAttributeReadWrite(Of String)(supertip, AttributeName.GetSupertip, AttributeCategory.SuperTip, callback.Method.Name))
		End Sub

#End Region

#Region "Description and GetDescription"

		Public Sub AddDescription(description As String)
			_attributes.Add(New RibbonAttributeReadOnly(Of String)(description, AttributeName.Description, AttributeCategory.Description))
		End Sub

		Public Sub AddGetDescription(description As String, callback As FromControl(Of String))
			_attributes.Add(New RibbonAttributeReadWrite(Of String)(description, AttributeName.GetDescription, AttributeCategory.Description, callback.Method.Name))
		End Sub

#End Region

#Region "Keytip and GetKeytip"

		Public Sub AddKeytip(keyTip As KeyTip)
			_attributes.Add(New RibbonAttributeReadOnly(Of KeyTip)(keyTip, AttributeName.Keytip, AttributeCategory.KeyTip))
		End Sub

		Public Sub AddGetKeytip(keyTip As KeyTip, callback As FromControl(Of KeyTip))
			_attributes.Add(New RibbonAttributeReadWrite(Of KeyTip)(keyTip, AttributeName.GetKeytip, AttributeCategory.KeyTip, callback.Method.Name))
		End Sub

#End Region

#Region "Title and GetTitle"

		Public Sub AddTitle(title As String)
			_attributes.Add(New RibbonAttributeReadOnly(Of String)(title, AttributeName.Title, AttributeCategory.Title))
		End Sub

		Public Sub AddGetTitle(title As String, callback As FromControl(Of String))
			_attributes.Add(New RibbonAttributeReadWrite(Of String)(title, AttributeName.GetTitle, AttributeCategory.Title, callback.Method.Name))
		End Sub

#End Region

#Region "Image, ImageMSO, and GetImage"

		Public Sub AddImage(path As String)
			_attributes.Add(New RibbonAttributeReadOnly(Of String)(path, AttributeName.Image, AttributeCategory.Image))
		End Sub

		Public Sub AddImageMso(image As ImageMSO)
			_attributes.Add(New RibbonAttributeReadOnly(Of ImageMSO)(image, AttributeName.ImageMso, AttributeCategory.Image))
		End Sub

		Public Sub AddGetImage(image As IPictureDisp, callback As FromControl(Of IPictureDisp))
			_attributes.Add(New RibbonAttributeReadWrite(Of IPictureDisp)(image, AttributeName.GetImage, AttributeCategory.Image, callback.Method.Name))
		End Sub

#End Region

#Region "ShowImage and GetShowImage"

		Public Sub AddShowImage(showImage As Boolean)
			_attributes.Add(New RibbonAttributeReadOnly(Of Boolean)(showImage, AttributeName.ShowImage, AttributeCategory.ImageVisibility))
		End Sub

		Public Sub AddGetShowImage(showImage As Boolean, callback As FromControl(Of Boolean))
			_attributes.Add(New RibbonAttributeReadWrite(Of Boolean)(showImage, AttributeName.GetShowImage, AttributeCategory.ImageVisibility, callback.Method.Name))
		End Sub

#End Region

#Region "Size and GetSize"

		Public Sub AddSize(size As ControlSize)
			_attributes.Add(New RibbonAttributeReadOnly(Of ControlSize)(size, AttributeName.Size, AttributeCategory.Size))
		End Sub

		Public Sub AddGetSize(size As ControlSize, callback As FromControl(Of ControlSize))
			_attributes.Add(New RibbonAttributeReadWrite(Of ControlSize)(size, AttributeName.GetSize, AttributeCategory.Size, callback.Method.Name))
		End Sub

#End Region

#Region "OnAction"

		Public Sub AddOnAction(callback As OnAction)
			_attributes.Add(New RibbonAttributeReadOnly(Of String)(callback.Method.Name, AttributeName.OnAction, AttributeCategory.OnAction))
		End Sub

		Public Sub AddOnAction(callback As ButtonPressed)
			_attributes.Add(New RibbonAttributeReadOnly(Of String)(callback.Method.Name, AttributeName.OnAction, AttributeCategory.OnAction))
		End Sub

#End Region

#Region "GetText And OnChange"

		Public Sub AddGetText(text As String, getText As FromControl(Of String), setText As TextChanged)
			_attributes.Add(New RibbonAttributeComposite(Of String)(text, AttributeName.GetText, AttributeCategory.Text, getText.Method.Name, AttributeName.OnChange, AttributeCategory.OnAction, setText.Method.Name))
		End Sub

#End Region

#Region "GetPressed And OnAction"

		Public Sub AddGetPressed(pressed As Boolean, getPressed As FromControl(Of Boolean), setPressed As ButtonPressed)
			_attributes.Add(New RibbonAttributeComposite(Of Boolean)(pressed, AttributeName.GetPressed, AttributeCategory.Pressed, getPressed.Method.Name, AttributeName.OnAction, AttributeCategory.OnAction, setPressed.Method.Name))
		End Sub

#End Region

#Region "GetItemId, GetItemLabel, GetItemScreentip, GetItemSupertip, and GetItemImage"

		Public Sub AddGetItemID(callback As FromControlAndIndex(Of String))
			_attributes.Add(New RibbonAttributeReadOnly(Of String)(callback.Method.Name, AttributeName.GetItemID, AttributeCategory.ItemId))
		End Sub

		Public Sub AddGetItemLabel(callback As FromControlAndIndex(Of String))
			_attributes.Add(New RibbonAttributeReadOnly(Of String)(callback.Method.Name, AttributeName.GetItemLabel, AttributeCategory.ItemLabel))
		End Sub

		Public Sub AddGetItemScreentip(callback As FromControlAndIndex(Of String))
			_attributes.Add(New RibbonAttributeReadOnly(Of String)(callback.Method.Name, AttributeName.GetItemScreentip, AttributeCategory.ItemScreentip))
		End Sub

		Public Sub AddGetItemSupertip(callback As FromControlAndIndex(Of String))
			_attributes.Add(New RibbonAttributeReadOnly(Of String)(callback.Method.Name, AttributeName.GetItemSupertip, AttributeCategory.ItemSupertip))
		End Sub

		Public Sub AddGetItemImage(callback As FromControlAndIndex(Of IPictureDisp))
			_attributes.Add(New RibbonAttributeReadOnly(Of String)(callback.Method.Name, AttributeName.GetItemImage, AttributeCategory.ItemImage))
		End Sub

#End Region

#Region "SelectedItemID, SelectedItemIndex, and OnAction"

		Public Sub AddGetSelected(getSelected As FromControl(Of String), setSelected As SelectionChanged)
			_attributes.Add(New RibbonAttributeComposite(Of Item)(Item.Dummy(), AttributeName.GetSelectedItemID, AttributeCategory.SelectedItemPosition, getSelected.Method.Name, AttributeName.OnAction, AttributeCategory.OnAction, setSelected.Method.Name))
		End Sub

		Public Sub AddGetSelected(getSelected As FromControl(Of Integer), setSelected As SelectionChanged)
			_attributes.Add(New RibbonAttributeComposite(Of Item)(Item.Dummy(), AttributeName.GetSelectedItemIndex, AttributeCategory.SelectedItemPosition, getSelected.Method.Name, AttributeName.OnAction, AttributeCategory.OnAction, setSelected.Method.Name))
		End Sub

#End Region

#Region "Rows and Columns"

		Public Sub AddRows(numberOfRows As Integer)
			_attributes.Add(New RibbonAttributeReadOnly(Of Integer)(numberOfRows, AttributeName.Rows, AttributeCategory.Rows))
		End Sub

		Public Sub AddColumns(numberOfColumns As Integer)
			_attributes.Add(New RibbonAttributeReadOnly(Of Integer)(numberOfColumns, AttributeName.Columns, AttributeCategory.Columns))
		End Sub

#End Region

#Region "ItemHeight and GetItemHeight"

		Public Sub AddItemHeight(itemHeight As Integer)
			_attributes.Add(New RibbonAttributeReadOnly(Of Integer)(itemHeight, AttributeName.ItemHeight, AttributeCategory.ItemHeight))
		End Sub

		Public Sub AddGetItemHeight(itemHeight As Integer, callback As FromControl(Of Integer))
			_attributes.Add(New RibbonAttributeReadWrite(Of Integer)(itemHeight, AttributeName.GetItemHeight, AttributeCategory.ItemHeight, callback.Method.Name))
		End Sub

#End Region

#Region "ItemWidth and GetItemWidth"

		Public Sub AddItemWidth(itemWidth As Integer)
			_attributes.Add(New RibbonAttributeReadOnly(Of Integer)(itemWidth, AttributeName.ItemWidth, AttributeCategory.ItemWidth))
		End Sub

		Public Sub AddGetItemWidth(itemWidth As Integer, callback As FromControl(Of Integer))
			_attributes.Add(New RibbonAttributeReadWrite(Of Integer)(itemWidth, AttributeName.GetItemWidth, AttributeCategory.ItemWidth, callback.Method.Name))
		End Sub

#End Region

#Region "Ribbon Only: LoadImage, OnLoad, and StartFromScratch"

		Public Sub AddLoadImage(callback As LoadImage)
			_attributes.Add(New RibbonAttributeReadOnly(Of String)(callback.Method.Name, AttributeName.LoadImage, AttributeCategory.LoadImage))
		End Sub

		Public Sub AddOnLoad(callback As OnLoad)
			_attributes.Add(New RibbonAttributeReadOnly(Of String)(callback.Method.Name, AttributeName.OnLoad, AttributeCategory.OnLoad))
		End Sub

		Public Sub AddStartFromScratch(startFromScratch As Boolean)
			_attributes.Add(New RibbonAttributeReadOnly(Of Boolean)(startFromScratch, AttributeName.StartFromScratch, AttributeCategory.StartFromScratch))
		End Sub

#End Region

		Public Sub AddMaxLength(maxLength As Byte)
			_attributes.Add(New RibbonAttributeReadOnly(Of Byte)(maxLength, AttributeName.MaxLength, AttributeCategory.MaxLength))
		End Sub

		Public Sub AddSizeString(sizeString As String)
			_attributes.Add(New RibbonAttributeReadOnly(Of String)(sizeString, AttributeName.SizeString, AttributeCategory.SizeString))
		End Sub

		Public Sub AddInvalidateContentOnDrop(invalidateOnDrop As Boolean)
			_attributes.Add(New RibbonAttributeReadOnly(Of Boolean)(invalidateOnDrop, AttributeName.InvalidateContentOnDrop, AttributeCategory.ContentInvalidation))
		End Sub

		Public Sub AddShowItemImage(showItemImage As Boolean)
			_attributes.Add(New RibbonAttributeReadOnly(Of Boolean)(showItemImage, AttributeName.ShowItemImage, AttributeCategory.ItemImageVisibility))
		End Sub

		Public Sub AddShowItemLabel(showItemLabel As Boolean)
			_attributes.Add(New RibbonAttributeReadOnly(Of Boolean)(showItemLabel, AttributeName.ShowItemLabel, AttributeCategory.ItemLabelVisibility))
		End Sub

		Public Sub AddGetItemCount(callback As FromControl(Of Integer))
			_attributes.Add(New RibbonAttributeReadOnly(Of String)(callback.Method.Name, AttributeName.GetItemCount, AttributeCategory.ItemCount))
		End Sub

		Public Sub AddItemSize(size As ControlSize)
			_attributes.Add(New RibbonAttributeReadOnly(Of ControlSize)(size, AttributeName.ItemSize, AttributeCategory.ItemSize))
		End Sub

		Public Sub AddBoxStyle(style As BoxStyle)
			_attributes.Add(New RibbonAttributeReadOnly(Of BoxStyle)(style, AttributeName.BoxStyle, AttributeCategory.BoxStyle))
		End Sub

		Public Sub AddShowInRibbon(showInRibbon As Boolean)
			_attributes.Add(New RibbonAttributeReadOnly(Of Boolean)(showInRibbon, AttributeName.ShowInRibbon, AttributeCategory.ShowInRibbon))
		End Sub

		Public Function GetEnumerator() As IEnumerator(Of IRibbonAttribute) Implements IEnumerable(Of IRibbonAttribute).GetEnumerator
			Return _attributes.GetEnumerator()
		End Function

		Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
			Return _attributes.GetEnumerator()
		End Function

	End Class

End Namespace