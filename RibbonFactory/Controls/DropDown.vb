Imports RibbonFactory.Component_Interfaces
Imports RibbonFactory.RibbonAttributes


Imports stdole

Namespace Controls

    Public NotInheritable Class DropDown
        Inherits RibbonElement
        Implements ICollection(Of DropdownItem)
        Implements ISelect
        Implements IEnabled
        Implements IVisible
        Implements ILabel
        Implements IScreenTip
        Implements ISuperTip
        Implements IShowLabel
        Implements IImage
        Implements IShowImage
        Implements IOnAction

        Private _selected As DropdownItem
        Private _selectedItemIndex As Integer
        Private ReadOnly _items As List(Of DropdownItem) = New List(Of DropdownItem)()
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
                Return $"<dropDown { _attributes }/>"
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

        Public Property SuperTip As String Implements ISupertip.Supertip
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

        Public Property Selected As DropdownItem Implements ISelect.Selected
            Get
                Return _selected
            End Get
            Set
                If _selected IsNot Value Then
                    _selected = Value
                    _selectedItemIndex = _items.IndexOf(_selected)
                    
                End If
            End Set
        End Property

        Public ReadOnly Property SelectedItemIndex As Integer Implements ISelect.SelectedItemIndex
            Get
                Return _selectedItemIndex
            End Get
        End Property

        Public Sub Execute() Implements IOnAction.Execute
            _attributes.ReadOnlyLookup(Of Action)(AttributeName.OnAction).GetValue().Invoke()
        End Sub

#Region "Methods Pertaining to Dropdown Items"

        Public Sub Add(item As DropdownItem) Implements ICollection(Of DropdownItem).Add
            _items.Add(item)
        End Sub

        Public Sub Add(ParamArray items As DropdownItem())
            _items.AddRange(items)
        End Sub

        Public Sub Clear() Implements ICollection(Of DropdownItem).Clear
            _items.Clear()
        End Sub

        Public Sub CopyTo(array() As DropdownItem, arrayIndex As Integer) Implements ICollection(Of DropdownItem).CopyTo
            _items.CopyTo(array, arrayIndex)
        End Sub

        Public Function Contains(item As DropdownItem) As Boolean Implements ICollection(Of DropdownItem).Contains
            Return _items.Contains(item)
        End Function

        Public ReadOnly Property Count As Integer Implements ICollection(Of DropdownItem).Count
            Get
                Return _items.Count
            End Get
        End Property

        Public ReadOnly Property IsReadOnly As Boolean Implements ICollection(Of DropdownItem).IsReadOnly
            Get
                Return False
            End Get
        End Property

        Public Function Remove(item As DropdownItem) As Boolean Implements ICollection(Of DropdownItem).Remove
            If _items.Remove(item) Then

                Return True
            Else
                Return False
            End If
        End Function

        Public Function GetEnumerator() As IEnumerator(Of DropdownItem) _
            Implements IEnumerable(Of DropdownItem).GetEnumerator
            Return _items.GetEnumerator()
        End Function

        Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return _items.GetEnumerator()
        End Function

#End Region
        
    End Class

End Namespace