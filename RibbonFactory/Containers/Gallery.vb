Imports RibbonFactory.ControlInterfaces
Imports RibbonFactory.Controls
Imports RibbonFactory.Enums
Imports RibbonFactory.RibbonAttributes
Imports RibbonFactory.Utilities
Imports stdole

Namespace Containers
    Public NotInheritable Class Gallery
        Inherits Container(Of Item)
        Implements IDescription
        Implements IEnabled
        Implements IVisible
        Implements IImage
        Implements IItemDimensions
        Implements IKeyTip
        Implements ILabel
        Implements IShowLabel
        Implements IScreenTip
        Implements ISuperTip
        Implements ISelect
        Implements IOnAction
        Implements ISize
        Implements IDefaultProvider

        Private _selected As Item
        Private _selectedItemIndex As Integer
        Private ReadOnly _attributes As AttributeSet

        Friend Sub New(attributes As AttributeSet, buttons As ICollection(Of Button), Optional tag As Object = Nothing)
            MyBase.New(New List(Of Item), tag)
            _attributes = attributes
            AddHandler _attributes.AttributeChanged, AddressOf RefreshNeeded
            Me.Buttons = buttons
        End Sub

        Public Overrides ReadOnly Property ID As String
            Get
                Return _attributes.ReadOnlyLookup(Of String)(AttributeCategory.IdType).GetValue()
            End Get
        End Property

        Public Overrides ReadOnly Property XML As String
            Get
                If Buttons IsNot Nothing AndAlso Buttons.Any() Then
                    Return String.Join(Environment.NewLine, $"<gallery { _attributes }>", String.Join(Environment.NewLine, Buttons), "</gallery>")
                Else
                    Return $"<gallery { _attributes }/>"
                End If
            End Get
        End Property

        Friend Overrides Sub Flatten(results As ICollection(Of RibbonElement))
            results.Add(Me)

            For Each button As Button In Buttons
                results.Add(button)
            Next
        End Sub

        Public Overrides Sub Add(item As Item)
            If Count = 0 Then
                _selected = item
                _selectedItemIndex = 0
            End If

            AddHandler item.ValueChanged, AddressOf OnChildItemChange

            MyBase.Add(item)
        End Sub

        Public Overrides Function Remove(item As Item) As Boolean
            If Items.Remove(item) Then
                RemoveHandler item.ValueChanged, AddressOf OnChildItemChange

                If _selected.Equals(item) Then
                    _selected = Items.FirstOrDefault()
                    _selectedItemIndex = 0
                End If

                RefreshNeeded()

                Return True
            Else
                Return False
            End If
        End Function

        Public Overrides Sub Clear()
            If Items.Any() Then
                _selected = Nothing
                _selectedItemIndex = 0

                For Each dropdownItem As Item In Me
                    RemoveHandler dropdownItem.ValueChanged, AddressOf OnChildItemChange
                Next

                Items.Clear()
                RefreshNeeded()
            End If
        End Sub

        Private Sub OnChildItemChange(sender As Object, e As ValueChangedEventArgs)
            RefreshNeeded()
        End Sub

        Public ReadOnly Property Buttons As ICollection(Of Button)

        Public Property Enabled As Boolean Implements IEnabled.Enabled
            Get
                Return _attributes.ReadOnlyLookup(Of Boolean)(AttributeCategory.Enabled).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of Boolean)(AttributeName.GetEnabled).SetValue(Value)
            End Set
        End Property

        Public Property Visible As Boolean Implements IVisible.Visible
            Get
                Return _attributes.ReadOnlyLookup(Of Boolean)(AttributeCategory.Visibility).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of Boolean)(AttributeCategory.Visibility).SetValue(Value)
            End Set
        End Property

        Public Property Description As String Implements IDescription.Description
            Get
                Return _attributes.ReadOnlyLookup(Of String)(AttributeCategory.Description).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of String)(AttributeCategory.Description).SetValue(Value)
            End Set
        End Property

        Public Property ItemHeight As Integer Implements IItemDimensions.ItemHeight
            Get
                Return _attributes.ReadOnlyLookup(Of Integer)(AttributeName.ItemHeight).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of Integer)(AttributeName.ItemHeight).SetValue(Value)
            End Set
        End Property

        Public Property ItemWidth As Integer Implements IItemDimensions.ItemWidth
            Get
                Return _attributes.ReadOnlyLookup(Of Integer)(AttributeName.ItemWidth).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of Integer)(AttributeName.ItemWidth).SetValue(Value)
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

        Public Property KeyTip As KeyTip Implements IKeyTip.KeyTip
            Get
                Return _attributes.ReadOnlyLookup(Of KeyTip)(AttributeName.Keytip).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of KeyTip)(AttributeName.GetKeytip).SetValue(Value)
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

        Public Property ShowLabel As Boolean Implements IShowLabel.ShowLabel
            Get
                Return _attributes.ReadOnlyLookup(Of Boolean)(AttributeName.ShowLabel).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of Boolean)(AttributeName.GetShowLabel).SetValue(Value)
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

        Public Property Selected As Item Implements ISelect.Selected
            Get
                Return _selected
            End Get
            Set
                If Value IsNot Nothing AndAlso Not _selected.Equals(Value) Then
                    _selected = Value
                    _selectedItemIndex = Items.IndexOf(Value)
                End If
            End Set
        End Property

        Public Property SelectedItemIndex As Integer Implements ISelect.SelectedItemIndex
            Get
                Return _selectedItemIndex
            End Get
            Set
                _selected = Items(Value)
                _selectedItemIndex = Value
            End Set
        End Property

        Public Property Size As ControlSize Implements ISize.Size
            Get
                Return _attributes.ReadOnlyLookup(Of ControlSize)(AttributeCategory.Size).GetValue()
            End Get
            Set
                _attributes.ReadWriteLookup(Of ControlSize)(AttributeCategory.Size).SetValue(Value)
            End Set
        End Property

        Public Sub Execute() Implements IOnAction.Execute
            _attributes.ReadOnlyLookup(Of Action)(AttributeName.OnAction).GetValue().Invoke()
        End Sub

        Private Function GetDefaults() As AttributeSet Implements IDefaultProvider.GetDefaults
            Return _attributes
        End Function

    End Class

End NameSpace