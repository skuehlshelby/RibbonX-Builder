Imports RibbonFactory.ComponentInterfaces
Imports RibbonFactory.RibbonAttributes
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

        Private ReadOnly _items As IList(Of DropdownItem) = New List(Of DropdownItem)()
        Private ReadOnly _attributes As AttributeGroup

        Friend Sub New(attributes As AttributeGroup, Optional tag As Object = Nothing)
            MyBase.New(tag)
            _attributes = attributes
        End Sub

        Public Overrides ReadOnly Property ID As String
            Get
                Return _attributes.ReadOnlyLookup(Of String)(AttributeName.Id).GetValue()
            End Get
        End Property

        Public Overrides ReadOnly Property XML As String
            Get
                Return $"<comboBox { _attributes }/>"
            End Get
        End Property

        Public Property Enabled As Boolean Implements IEnabled.Enabled
            Get
                Return _attributes.ReadOnlyLookup(Of Boolean)(AttributeName.Enabled).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of Boolean)(AttributeName.GetEnabled).SetValue(Value)
            End Set
        End Property

        Public Property Visible As Boolean Implements IVisible.Visible
            Get
                Return _attributes.ReadOnlyLookup(Of Boolean)(AttributeName.Visible).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of Boolean)(AttributeName.GetVisible).SetValue(Value)
            End Set
        End Property

        Public Property Label As String Implements ILabel.Label
            Get
                Return _attributes.ReadOnlyLookup(Of String)(AttributeName.Label).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of String)(AttributeName.GetLabel).SetValue(Value)
            End Set
        End Property

        Public Property ScreenTip As String Implements IScreenTip.ScreenTip
            Get
                Return _attributes.ReadOnlyLookup(Of String)(AttributeName.Screentip).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of String)(AttributeName.GetScreentip).SetValue(Value)
            End Set
        End Property

        Public Property SuperTip As String Implements ISuperTip.SuperTip
            Get
                Return _attributes.ReadOnlyLookup(Of String)(AttributeName.Supertip).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of String)(AttributeName.GetSupertip).SetValue(Value)
            End Set
        End Property

        Public Property ShowLabel As Boolean Implements IShowLabel.ShowLabel
            Get
                Return _attributes.ReadOnlyLookup(Of Boolean)(AttributeName.ShowLabel).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of Boolean)(AttributeName.GetShowLabel).SetValue(Value)
            End Set
        End Property

        Public Property KeyTip As KeyTip Implements IKeyTip.KeyTip
            Get
                Return _attributes.ReadOnlyLookup(Of KeyTip)(AttributeName.Keytip).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of KeyTip)(AttributeName.GetKeytip).SetValue(Value)
            End Set
        End Property

        Public Property Image As IPictureDisp Implements IImage.Image
            Get
                Return _attributes.ReadOnlyLookup(Of IPictureDisp)(AttributeName.GetImage).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of IPictureDisp)(AttributeName.GetImage).SetValue(Value)
            End Set
        End Property

        Public Property ShowImage As Boolean Implements IShowImage.ShowImage
            Get
                Return _attributes.ReadOnlyLookup(Of Boolean)(AttributeName.ShowImage).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of Boolean)(AttributeName.GetShowImage).SetValue(Value)
            End Set
        End Property

        Public ReadOnly Property MaxLength As Integer Implements IMaxLength.MaxLength
            Get
                Return _attributes.ReadOnlyLookup(Of Integer)(AttributeName.MaxLength).GetValue()
            End Get
        End Property

        Public Property Text As String Implements IText.Text
            Get
                Return _attributes.ReadOnlyLookup(Of String)(AttributeName.GetText).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of String)(AttributeName.GetText).SetValue(Value)
            End Set
        End Property

        Public Sub Execute() Implements IOnChange.Execute
            _attributes.ReadOnlyLookup(Of Action)(AttributeName.OnAction).GetValue().Invoke()
        End Sub

        #Region "Methods Pertaining to Dropdown Items"

        Public Sub Add(element As DropdownItem) Implements ICollection(Of DropdownItem).Add
            _items.Add(element)
        End Sub

        Public Sub Clear() Implements ICollection(Of DropdownItem).Clear
            _items.Clear()
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
                Return _items(index)
            End Get
            Set
                _items(index) = value
            End Set
        End Property

        Public Function Remove(element As DropdownItem) As Boolean Implements IList(Of DropdownItem).Remove
            If _items.Remove(element) Then
                
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
            Return _items.IndexOf(element)
        End Function

        Public Sub Insert(index As Integer, element As DropdownItem) Implements IList(Of DropdownItem).Insert
            _items.Insert(index, element)
        End Sub

        Public Sub RemoveAt(index As Integer) Implements IList(Of DropdownItem).RemoveAt
            _items.RemoveAt(index)
        End Sub

#End Region

    End Class

End Namespace