Imports RibbonFactory.Component_Interfaces
Imports RibbonFactory.RibbonAttributes
Imports RibbonFactory.RibbonAttributes.Categories.Enabled
Imports RibbonFactory.RibbonAttributes.Categories.ID
Imports RibbonFactory.RibbonAttributes.Categories.Image
Imports RibbonFactory.RibbonAttributes.Categories.Label
Imports RibbonFactory.RibbonAttributes.Categories.Screentip
Imports RibbonFactory.RibbonAttributes.Categories.SuperTip
Imports RibbonFactory.RibbonAttributes.Categories.Visible
Imports stdole

Namespace Controls

    Public Class DropDown
        Inherits RibbonElement
        Implements ICollection(Of DropdownItem)
        Implements ISelectable
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
        Private ReadOnly _items As List(Of DropdownItem) = New List(Of DropdownItem)()
        Private ReadOnly _attributes As AttributeGroup

        Friend Sub New(attributes As AttributeGroup, Optional tag As Object = Nothing)
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
                Return $"<dropDown { String.Join(" ", _attributes) }/>"
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

        Public Property Image As Object Implements IImage.Image
            Get
                If IsCustom Then
                    Return _attributes.Lookup(Of GetImage).GetValue()
                Else
                    Return _attributes.Lookup(Of ImageMso).GetValue()
                End If
            End Get
            Set
                If IsCustom Then
                    _attributes.Lookup(Of GetImage).SetValue(CType(Value, IPictureDisp))
                Else
                    _attributes.Lookup(Of ImageMso).GetValue()
                End If
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

        Public Property Selected As DropdownItem Implements ISelectable.Selected
            Get
                Return _selected
            End Get
            Set
                _selected = Value
                Refresh()
            End Set
        End Property

        Public Sub Execute() Implements IOnAction.Execute
            _attributes.Lookup(Of Categories.OnAction.OnAction).GetValue().Invoke()
        End Sub

#Region "Methods Pertaining to Dropdown Items"

        Public Sub Add(item As DropdownItem) Implements ICollection(Of DropdownItem).Add
            _items.Add(item)
            Refresh()
        End Sub

        Public Sub Clear() Implements ICollection(Of DropdownItem).Clear
            _items.Clear()
            Refresh()
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

#End Region

#Region "Overriden Base Class Methods"

        Public Overrides Function Equals(obj As Object) As Boolean
            Return obj.GetHashCode() = GetHashCode() AndAlso TypeOf obj Is ComboBox
        End Function

#End Region
    End Class

End Namespace