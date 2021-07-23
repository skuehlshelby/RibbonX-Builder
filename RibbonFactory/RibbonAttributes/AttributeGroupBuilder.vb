Imports RibbonFactory.Enums
Imports RibbonFactory.Enums.ImageMSO
Imports RibbonFactory.Enums.MSO
Imports stdole

Namespace RibbonAttributes

    Friend Class AttributeGroupBuilder

        Private ReadOnly _attributes As AttributeGroup = New AttributeGroup()

        Public Sub SetDefaults(defaultProvider As IDefaultProvider)
            If _attributes.Any() Then
                Throw New InvalidOperationException("Attempted to set defaults when one or more attributes were already set.")
            End If

            For Each attribute As RibbonAttribute In defaultProvider.GetDefaults()
                _attributes.Add(attribute)
            Next
        End Sub

        Public Function Build() As AttributeGroup 
            Return _attributes
        End Function

        Public Sub AddEnabled(enabled As Boolean)
            _attributes.Add(New RibbonAttributeReadOnly(Of Boolean)(enabled, AttributeName.Enabled, AttributeName.GetEnabled))
        End Sub

        Public Sub AddGetEnabled(enabled As Boolean, callback As FromControl(Of Boolean))
            _attributes.Add(New RibbonAttributeReadWrite(Of Boolean)(enabled, callback.Method.Name, AttributeName.GetEnabled, AttributeName.Enabled))
        End Sub

        Public Sub AddID(id As String)
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(id, AttributeName.Id, AttributeName.IdMso, AttributeName.IdQ))
        End Sub

        Public Sub AddIdMso(idMso As MSO)
            _attributes.Add(New RibbonAttributeReadOnly(Of MSO)(idMso, AttributeName.IdMso, AttributeName.Id, AttributeName.IdQ))
        End Sub

        Public Sub AddIdQ([namespace] As String, id As String)
            _attributes.Add(New RibbonAttributeReadOnly(Of String)($"{[namespace]}:{id}", AttributeName.IdQ, AttributeName.Id, AttributeName.IdMso))
        End Sub

        Public Sub AddInsertAfterMso(mso As MSO)
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(mso.ToString(), AttributeName.InsertAfterMso, AttributeName.InsertAfterQ, AttributeName.InsertBeforeQ, AttributeName.InsertBeforeMso))
        End Sub

        Public Sub AddInsertAfterQ(control As RibbonElement)
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(control.ID, AttributeName.InsertAfterQ, AttributeName.InsertAfterMso, AttributeName.InsertBeforeQ, AttributeName.InsertBeforeMso))
        End Sub

        Public Sub AddInsertBeforeMso(mso As MSO)
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(mso.ToString(), AttributeName.InsertBeforeMso, AttributeName.InsertAfterQ, AttributeName.InsertBeforeQ, AttributeName.InsertAfterMso))
        End Sub

        Public Sub AddInsertBeforeQ(control As RibbonElement)
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(control.ID, AttributeName.InsertBeforeQ, AttributeName.InsertAfterQ, AttributeName.InsertAfterMso, AttributeName.InsertBeforeMso))
        End Sub

        Public Sub AddVisible(visible As Boolean)
            _attributes.Add(New RibbonAttributeReadOnly(Of Boolean)(visible, AttributeName.Visible, AttributeName.GetVisible))
        End Sub

        Public Sub AddGetVisible(visible As Boolean, callback As FromControl(Of Boolean))
            _attributes.Add(New RibbonAttributeReadWrite(Of Boolean)(visible, callback.Method.Name, AttributeName.GetVisible, AttributeName.Visible))
        End Sub

        Public Sub AddLabel(label As String)
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(label, AttributeName.Label, AttributeName.GetLabel))
        End Sub

        Public Sub AddGetLabel(label As String, callback As FromControl(Of String))
            _attributes.Add(New RibbonAttributeReadWrite(Of String)(label, callback.Method.Name, AttributeName.GetLabel, AttributeName.Label))
        End Sub

        Public Sub AddShowLabel(showLabel As Boolean)
            _attributes.Add(New RibbonAttributeReadOnly(Of Boolean)(showLabel, AttributeName.ShowLabel, AttributeName.GetShowLabel))
        End Sub

        Public Sub AddGetShowLabel(showLabel As Boolean, callback As FromControl(Of Boolean))
            _attributes.Add(New RibbonAttributeReadWrite(Of Boolean)(showLabel, callback.Method.Name, AttributeName.GetShowLabel, AttributeName.ShowLabel))
        End Sub

        Public Sub AddScreentip(screentip As String)
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(screentip, AttributeName.Screentip, AttributeName.GetScreentip))
        End Sub

        Public Sub AddGetScreentip(screentip As String, callback As FromControl(Of String))
            _attributes.Add(New RibbonAttributeReadWrite(Of String)(screentip, callback.Method.Name, AttributeName.GetScreentip, AttributeName.Screentip))
        End Sub

        Public Sub AddSupertip(supertip As String)
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(supertip, AttributeName.Supertip, AttributeName.GetSupertip))
        End Sub

        Public Sub AddGetSupertip(supertip As String, callback As FromControl(Of String))
            _attributes.Add(New RibbonAttributeReadWrite(Of String)(supertip, callback.Method.Name, AttributeName.GetSupertip, AttributeName.Supertip))
        End Sub

        Public Sub AddDescription(description As String)
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(description, AttributeName.Description, AttributeName.GetDescription))
        End Sub

        Public Sub AddGetDescription(description As String, callback As FromControl(Of String))
            _attributes.Add(New RibbonAttributeReadWrite(Of String)(description, callback.Method.Name, AttributeName.GetDescription, AttributeName.Description))
        End Sub

        Public Sub AddKeytip(keyTip As KeyTip)
            _attributes.Add(New RibbonAttributeReadOnly(Of KeyTip)(keyTip, AttributeName.Keytip, AttributeName.GetKeytip))
        End Sub

        Public Sub AddGetKeytip(keyTip As KeyTip, callback As FromControl(Of KeyTip))
            _attributes.Add(New RibbonAttributeReadWrite(Of KeyTip)(keyTip, callback.Method.Name, AttributeName.GetKeytip, AttributeName.Keytip))
        End Sub

        Public Sub AddImage(path As String)
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(path, AttributeName.Image, AttributeName.ImageMso, AttributeName.GetImage))
        End Sub

        Public Sub AddImageMso(image As ImageMSO)
            _attributes.Add(New RibbonAttributeReadOnly(Of ImageMSO)(image, AttributeName.ImageMso, AttributeName.Image, AttributeName.GetImage))
        End Sub

        Public Sub AddGetImage(image As IPictureDisp, callback As FromControl(Of IPictureDisp))
            _attributes.Add(New RibbonAttributeReadWrite(Of IPictureDisp)(image, callback.Method.Name, AttributeName.GetImage, AttributeName.Image, AttributeName.ImageMso))
        End Sub

        Public Sub AddShowImage(showImage As Boolean)
            _attributes.Add(New RibbonAttributeReadOnly(Of Boolean)(showImage, AttributeName.ShowImage, AttributeName.GetShowImage))
        End Sub

        Public Sub AddGetShowImage(showImage As Boolean, callback As FromControl(Of Boolean))
            _attributes.Add(New RibbonAttributeReadWrite(Of Boolean)(showImage, callback.Method.Name, AttributeName.GetShowImage, AttributeName.ShowImage))
        End Sub

        Public Sub AddOnAction(action As Action, callback As OnAction)
            _attributes.Add(New RibbonAttributeReadWrite(Of Action)(action, callback.Method.Name, AttributeName.OnAction))
        End Sub

        Public Sub AddOnAction(action As Action, callback As ButtonPressed)
            _attributes.Add(New RibbonAttributeReadWrite(Of Action)(action, callback.Method.Name, AttributeName.OnAction))
        End Sub

        Public Sub AddOnAction(action As Action, callback As SelectionChanged)
            _attributes.Add(New RibbonAttributeReadWrite(Of Action)(action, callback.Method.Name, AttributeName.OnAction))
        End Sub

        Public Sub AddOnChange(action As Action, callback As TextChanged)
            _attributes.Add(New RibbonAttributeReadWrite(Of Action)(action, callback.Method.Name, AttributeName.OnChange))
        End Sub

        Public Sub AddGetText(text As String, callback As FromControl(Of String))
            _attributes.Add(New RibbonAttributeReadWrite(Of String)(text, callback.Method.Name, AttributeName.GetText))
        End Sub

        Public Sub AddGetPressed(pressed As Boolean, callback As FromControl(Of Boolean))
            _attributes.Add(New RibbonAttributeReadWrite(Of Boolean)(pressed, callback.Method.Name, AttributeName.GetPressed))
        End Sub

        Public Sub AddSize(size As ControlSize)
            _attributes.Add(New RibbonAttributeReadOnly(Of ControlSize)(size, AttributeName.Size, AttributeName.GetSize))
        End Sub

        Public Sub AddGetSize(size As ControlSize, callback As FromControl(Of ControlSize))
            _attributes.Add(New RibbonAttributeReadWrite(Of ControlSize)(size, callback.Method.Name, AttributeName.GetSize, AttributeName.Size))
        End Sub

        Public Sub AddMaxLength(maxLength As Byte)
            _attributes.Add(New RibbonAttributeReadOnly(Of Byte)(maxLength, AttributeName.MaxLength))
        End Sub

        Public Sub AddSizeString(sizeString As String)
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(sizeString, AttributeName.SizeString))
        End Sub

        Public Sub AddInvalidateContentOnDrop(invalidateOnDrop As Boolean)
            _attributes.Add(New RibbonAttributeReadOnly(Of Boolean)(invalidateOnDrop, AttributeName.InvalidateContentOnDrop))
        End Sub

        Public Sub AddShowItemImage(showItemImage As Boolean)
            _attributes.Add(New RibbonAttributeReadOnly(Of Boolean)(showItemImage, AttributeName.ShowItemImage))
        End Sub

        Public Sub AddShowItemLabel(showItemLabel As Boolean)
            _attributes.Add(New RibbonAttributeReadOnly(Of Boolean)(showItemLabel, AttributeName.ShowItemLabel))
        End Sub

        Public Sub AddGetItemCount(callback As FromControl(Of Integer))
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(callback.Method.Name, AttributeName.GetItemCount))
        End Sub

        Public Sub AddGetItemID(callback As FromControlAndIndex(Of String))
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(callback.Method.Name, AttributeName.GetItemID))
        End Sub

        Public Sub AddGetItemImage(callback As FromControlAndIndex(Of IPictureDisp))
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(callback.Method.Name, AttributeName.GetItemImage))
        End Sub

        Public Sub AddGetItemLabel(callback As FromControlAndIndex(Of String))
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(callback.Method.Name, AttributeName.GetItemLabel))
        End Sub

        Public Sub AddGetItemScreentip(callback As FromControlAndIndex(Of String))
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(callback.Method.Name, AttributeName.GetItemScreentip))
        End Sub

        Public Sub AddGetItemSupertip(callback As FromControlAndIndex(Of String))
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(callback.Method.Name, AttributeName.GetItemSupertip))
        End Sub

        Public Sub AddGetSelectedItemID(callback As FromControl(Of String))
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(callback.Method.Name, AttributeName.GetSelectedItemID))
        End Sub

        Public Sub AddGetSelectedItemIndex(callback As FromControl(Of Integer))
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(callback.Method.Name, AttributeName.GetSelectedItemIndex))
        End Sub

        Public Sub AddColumns(numberOfColumns As Integer)
            _attributes.Add(New RibbonAttributeReadOnly(Of Integer)(numberOfColumns, AttributeName.Columns))
        End Sub

        Public Sub AddRows(numberOfRows As Integer)
            _attributes.Add(New RibbonAttributeReadOnly(Of Integer)(numberOfRows, AttributeName.Rows))
        End Sub

        Public Sub AddItemHeight(itemHeight As Integer)
            _attributes.Add(New RibbonAttributeReadOnly(Of Integer)(itemHeight, AttributeName.ItemHeight, AttributeName.GetItemHeight))
        End Sub

        Public Sub AddGetItemHeight(itemHeight As Integer, callback As FromControl(Of Integer))
            _attributes.Add(New RibbonAttributeReadWrite(Of Integer)(itemHeight, callback.Method.Name, AttributeName.GetItemHeight, AttributeName.ItemHeight))
        End Sub

        Public Sub AddItemWidth(itemWidth As Integer)
            _attributes.Add(New RibbonAttributeReadOnly(Of Integer)(itemWidth, AttributeName.ItemWidth, AttributeName.GetItemWidth))
        End Sub

        Public Sub AddGetItemWidth(itemWidth As Integer, callback As FromControl(Of Integer))
            _attributes.Add(New RibbonAttributeReadWrite(Of Integer)(itemWidth, callback.Method.Name, AttributeName.GetItemWidth, AttributeName.ItemWidth))
        End Sub

        Public Sub AddStartFromScratch(startFromScratch As Boolean)
            _attributes.Add(New RibbonAttributeReadOnly(Of Boolean)(startFromScratch, AttributeName.StartFromScratch))
        End Sub

        Public Sub AddOnLoad(callback As OnLoad)
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(callback.Method.Name, AttributeName.OnLoad))
        End Sub

        Public Sub AddLoadImage(callback As LoadImage)
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(callback.Method.Name, AttributeName.LoadImage))
        End Sub

        Public Sub AddBoxStyle(style As BoxStyle)
            _attributes.Add(New RibbonAttributeReadOnly(Of BoxStyle)(style, AttributeName.BoxStyle))
        End Sub

        Public Sub AddTitle(title As String)
            _attributes.Add(New RibbonAttributeReadOnly(Of String)(title, AttributeName.Title, AttributeName.GetTitle))
        End Sub

        Public Sub AddGetTitle(title As String, callback As FromControl(Of String))
            _attributes.Add(New RibbonAttributeReadWrite(Of String)(title, callback.Method.Name,  AttributeName.GetTitle, AttributeName.Title))
        End Sub

    End Class

End NameSpace