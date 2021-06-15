Imports RibbonFactory.Component_Interfaces
Imports RibbonFactory.RibbonAttributes
Imports RibbonFactory.RibbonAttributes.Categories.Enabled
Imports RibbonFactory.RibbonAttributes.Categories.ID
Imports RibbonFactory.RibbonAttributes.Categories.Image
Imports RibbonFactory.RibbonAttributes.Categories.KeyTip
Imports RibbonFactory.RibbonAttributes.Categories.Label
Imports RibbonFactory.RibbonAttributes.Categories.MaxLength
Imports RibbonFactory.RibbonAttributes.Categories.Screentip
Imports RibbonFactory.RibbonAttributes.Categories.SuperTip
Imports RibbonFactory.RibbonAttributes.Categories.Text
Imports RibbonFactory.RibbonAttributes.Categories.Visible
Imports stdole

Namespace Controls

    Public NotInheritable Class ComboBox
        Inherits RibbonElement
        Implements IList(Of DropdownItem)
        Implements IEnabled
        Implements IVisible
        Implements ILabel
        Implements IScreenTip
        Implements ISuperTip
        Implements IShowLabel
        Implements IKeyTip
        Implements IImage
        Implements IShowImage
        Implements IMaxLength
        Implements IOnChange
        Implements IText

        Private ReadOnly _items As List(Of DropdownItem) = New List(Of DropdownItem)()
        Private ReadOnly _attributes As AttributeGroup

        Public Sub New(attributes As AttributeGroup, Optional tag As Object = Nothing)
            MyBase.New(tag)
            _attributes = attributes
        End Sub

        Public Overrides ReadOnly Property ID As String
            Get
                Return _attributes.Lookup(Of Id).GetValue()
            End Get
        End Property

        Public Overrides ReadOnly Property XML As String
            Get
                Return $"<{NameOf(ComboBox).CamelCase()} { String.Join(" ", _attributes) }/>"
            End Get
        End Property

        Public Property Enabled As Boolean Implements IEnabled.Enabled
            Get
                Return _attributes.Lookup(Of Enabled).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetEnabled).SetValue(Value)
                Refresh()
            End Set
        End Property

        Public Property Visible As Boolean Implements IVisible.Visible
            Get
                Return _attributes.Lookup(Of Visible).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetVisible).SetValue(Value)
                Refresh()
            End Set
        End Property

        Public Property Label As String Implements ILabel.Label
            Get
                Return _attributes.Lookup(Of Label).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetLabel).SetValue(Value)
                Refresh()
            End Set
        End Property

        Public Property ScreenTip As String Implements IScreenTip.ScreenTip
            Get
                Return _attributes.Lookup(Of Screentip).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetScreentip).SetValue(Value)
                Refresh()
            End Set
        End Property

        Public Property SuperTip As String Implements ISuperTip.SuperTip
            Get
                Return _attributes.Lookup(Of Supertip).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetSuperTip).SetValue(Value)
                Refresh()
            End Set
        End Property

        Public Property ShowLabel As Boolean Implements IShowLabel.ShowLabel
            Get
                Return _attributes.Lookup(Of ShowLabel).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetShowLabel).SetValue(Value)
                Refresh()
            End Set
        End Property

        Public Property KeyTip As KeyTip Implements IKeyTip.KeyTip

            Get
                Return _attributes.Lookup(Of Categories.KeyTip.Keytip).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetKeytip).SetValue(Value)
                Refresh()
            End Set
        End Property

        Public Property Image As IPictureDisp Implements IImage.Image
            Get
                Return _attributes.Lookup(Of GetImage).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetImage).SetValue(value)
                Refresh()
            End Set
        End Property

        Public ReadOnly Property IsCustom As Boolean Implements IImage.IsCustom
            Get
                Return TypeOf _attributes.Lookup(Of ImageBase) Is GetImage
            End Get
        End Property

        Public Property ShowImage As Boolean Implements IShowImage.ShowImage
            Get
                Return _attributes.Lookup(Of ShowImage).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetShowImage).SetValue(Value)
                Refresh()
            End Set
        End Property

        Public ReadOnly Property MaxLength As UShort Implements IMaxLength.MaxLength
            Get
                Return _attributes.Lookup(Of MaxLength).GetValue()
            End Get
        End Property

        Public Property Text As String Implements IText.Text

            Get
                Return _attributes.Lookup(Of GetText).GetValue()
            End Get
            Set
                _attributes.Lookup(Of GetText).SetValue(Value)
                Refresh()
            End Set
        End Property

        Public Sub Execute() Implements IOnChange.Execute
            _attributes.Lookup(Of Categories.OnChange.OnChange).GetValue().Invoke()
        End Sub

        #Region "Methods Pertaining to Dropdown Items"

        Public Sub Add(element As DropdownItem) Implements ICollection(Of DropdownItem).Add
            _items.Add(element)
            Refresh()
        End Sub

        Public Sub Clear() Implements ICollection(Of DropdownItem).Clear
            _items.Clear()
            Refresh()
        End Sub

        Public Sub CopyTo(array() As DropdownItem, arrayIndex As Integer) Implements IList(Of DropdownItem).CopyTo
            _items.CopyTo(array, arrayIndex)
        End Sub

        Public Function Contains(element As DropdownItem) As Boolean Implements IList(Of DropdownItem).Contains
            Return _items.Contains(element)
        End Function

        Public ReadOnly Property Count As Integer Implements IList(Of DropdownItem).Count
            Get
                Return _items.Count
            End Get
        End Property

        Public ReadOnly Property IsReadOnly As Boolean Implements IList(Of DropdownItem).IsReadOnly
            Get
                Return False
            End Get
        End Property

        Default Public Property Item(index As Integer) As DropdownItem Implements IList(Of DropdownItem).Item
            Get
                Return DirectCast(_items, IList(Of DropdownItem))(index)
            End Get
            Set
                DirectCast(_items, IList(Of DropdownItem))(index) = value
            End Set
        End Property

        Public Function Remove(element As DropdownItem) As Boolean Implements IList(Of DropdownItem).Remove
            If _items.Remove(element) Then
                Refresh()
                Return True
            Else
                Return False
            End If
        End Function

        Public Function GetEnumerator() As IEnumerator(Of DropdownItem) Implements IEnumerable(Of DropdownItem).GetEnumerator
            Return _items.GetEnumerator()
        End Function

        Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return _items.GetEnumerator()
        End Function

        Public Function IndexOf(element As DropdownItem) As Integer Implements IList(Of DropdownItem).IndexOf
            Return DirectCast(_items, IList(Of DropdownItem)).IndexOf(element)
        End Function

        Public Sub Insert(index As Integer, element As DropdownItem) Implements IList(Of DropdownItem).Insert
            DirectCast(_items, IList(Of DropdownItem)).Insert(index, element)
        End Sub

        Public Sub RemoveAt(index As Integer) Implements IList(Of DropdownItem).RemoveAt
            DirectCast(_items, IList(Of DropdownItem)).RemoveAt(index)
        End Sub

#End Region

    End Class

End Namespace