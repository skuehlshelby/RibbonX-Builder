Imports System.Reflection

Namespace RibbonAttributes
    ''' <summary>
    ''' This class provides a single place to specify which RibbonX attributes are members of the same category.
    ''' Members of the same category are mutually-exclusive.
    ''' This class has semantics similar to an enum.
    ''' </summary>
    Friend NotInheritable Class AttributeCategory
        Implements IEquatable(Of AttributeCategory)
        Implements ICollection(Of AttributeName)

        Private ReadOnly _id As Byte
        Private ReadOnly _name As String
        Private ReadOnly _members As ICollection(Of AttributeName)

        Private Sub New(id As Byte, name As String, Paramarray members() As AttributeName)
            _id = id
            _name = name
            _members = members
        End Sub

#Region "ICollection Members"

        Public ReadOnly Property Count As Integer Implements ICollection(Of AttributeName).Count
            Get
                Return _members.Count
            End Get
        End Property

        Public ReadOnly Property IsReadOnly As Boolean Implements ICollection(Of AttributeName).IsReadOnly
            Get
                Return True
            End Get
        End Property

        Public Sub Add(item As AttributeName) Implements ICollection(Of AttributeName).Add
            Throw New NotSupportedException($"The operation '{NameOf(Add)}' is not supported.")
        End Sub

        Public Sub Clear() Implements ICollection(Of AttributeName).Clear
            Throw New NotSupportedException($"The operation '{NameOf(Clear)}' is not supported.")
        End Sub

        Public Function Contains(item As AttributeName) As Boolean Implements ICollection(Of AttributeName).Contains
            Return _members.Contains(item)
        End Function

        Public Sub CopyTo(array() As AttributeName, arrayIndex As Integer) Implements ICollection(Of AttributeName).CopyTo
            _members.CopyTo(array, arrayIndex)
        End Sub

        Public Function Remove(item As AttributeName) As Boolean Implements ICollection(Of AttributeName).Remove
            Throw New NotSupportedException($"The operation '{NameOf(Remove)}' is not supported.")
        End Function

        Public Function GetEnumerator() As IEnumerator(Of AttributeName) Implements IEnumerable(Of AttributeName).GetEnumerator
            Return _members.GetEnumerator()
        End Function

        Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return DirectCast(_members, IEnumerable).GetEnumerator()
        End Function

#End Region

#Region "Categories"

        Public Shared Function Values() As IEnumerable(Of AttributeCategory)
            Return _
                GetType(AttributeCategory).GetProperties(
                    BindingFlags.Public Or BindingFlags.GetProperty Or BindingFlags.Static).Select(
                        Function(info) info.GetValue(Nothing)).Cast(Of AttributeCategory)()
        End Function

        Public Shared ReadOnly Property Enabled As AttributeCategory = New AttributeCategory(0, NameOf(Enabled), AttributeName.Enabled, AttributeName.GetEnabled)

        Public Shared ReadOnly Property IdType As AttributeCategory = New AttributeCategory(1, NameOf(IdType), AttributeName.Id, AttributeName.IdMso, AttributeName.IdQ)

        Public Shared ReadOnly Property Insertion As AttributeCategory = New AttributeCategory(2, NameOf(Insertion), AttributeName.InsertAfterMso, AttributeName.InsertAfterQ, AttributeName.InsertBeforeQ, AttributeName.InsertBeforeMso)

        Public Shared ReadOnly Property Visibility As AttributeCategory = New AttributeCategory(3, NameOf(Visibility), AttributeName.Visible, AttributeName.GetVisible)

        Public Shared ReadOnly Property Label As AttributeCategory = New AttributeCategory(4, NameOf(Label), AttributeName.Label, AttributeName.GetLabel)

        Public Shared ReadOnly Property LabelVisibility As AttributeCategory = New AttributeCategory(5, NameOf(LabelVisibility), AttributeName.ShowLabel, AttributeName.GetShowLabel)

        Public Shared ReadOnly Property ScreenTip As AttributeCategory = New AttributeCategory(6, NameOf(ScreenTip), AttributeName.Screentip, AttributeName.GetScreentip)

        Public Shared ReadOnly Property SuperTip As AttributeCategory = New AttributeCategory(7, NameOf(SuperTip), AttributeName.Supertip, AttributeName.GetSupertip)

        Public Shared ReadOnly Property Description As AttributeCategory = New AttributeCategory(8, NameOf(Description), AttributeName.Description, AttributeName.GetDescription)

        Public Shared ReadOnly Property KeyTip As AttributeCategory = New AttributeCategory(9, NameOf(KeyTip), AttributeName.Keytip, AttributeName.GetKeytip)

        Public Shared ReadOnly Property Image As AttributeCategory = New AttributeCategory(10, NameOf(Image), AttributeName.Image, AttributeName.ImageMso, AttributeName.GetImage)

        Public Shared ReadOnly Property ImageVisibility As AttributeCategory = New AttributeCategory(11, NameOf(ImageVisibility), AttributeName.ShowImage, AttributeName.GetShowImage)

        Public Shared ReadOnly Property OnAction As AttributeCategory = New AttributeCategory(12, NameOf(OnAction), AttributeName.OnAction)

        Public Shared ReadOnly Property OnChange As AttributeCategory = New AttributeCategory(13, NameOf(OnChange), AttributeName.OnChange)

        Public Shared ReadOnly Property Text As AttributeCategory = New AttributeCategory(14, NameOf(Text), AttributeName.GetText)

        Public Shared ReadOnly Property Pressed As AttributeCategory = New AttributeCategory(15, NameOf(Pressed), AttributeName.GetPressed) 

        Public Shared ReadOnly Property Size As AttributeCategory = New AttributeCategory(16, NameOf(Size), AttributeName.Size, AttributeName.GetSize) 

        Public Shared ReadOnly Property MaxLength As AttributeCategory = New AttributeCategory(17, NameOf(MaxLength), AttributeName.MaxLength) 

        Public Shared ReadOnly Property SizeString As AttributeCategory = New AttributeCategory(18, NameOf(SizeString), AttributeName.SizeString) 

        Public Shared ReadOnly Property ContentInvalidation As AttributeCategory = New AttributeCategory(19, NameOf(ContentInvalidation), AttributeName.InvalidateContentOnDrop)

        Public Shared ReadOnly Property ItemImageVisibility As AttributeCategory = New AttributeCategory(20, NameOf(ItemImageVisibility), AttributeName.ShowItemImage)

        Public Shared ReadOnly Property ItemLabelVisibility As AttributeCategory = New AttributeCategory(21, NameOf(ItemLabelVisibility), AttributeName.ShowItemLabel)

        Public Shared ReadOnly Property ItemCount As AttributeCategory = New AttributeCategory(22, NameOf(ItemCount), AttributeName.GetItemCount)

        Public Shared ReadOnly Property ItemId As AttributeCategory = New AttributeCategory(23, NameOf(ItemId), AttributeName.GetItemID)

        Public Shared ReadOnly Property ItemImage As AttributeCategory = New AttributeCategory(24, NameOf(ItemImage), AttributeName.GetItemImage)

        Public Shared ReadOnly Property ItemLabel As AttributeCategory = New AttributeCategory(25, NameOf(ItemLabel), AttributeName.GetItemLabel)

        Public Shared ReadOnly Property ItemScreentip As AttributeCategory = New AttributeCategory(26, NameOf(ItemScreentip), AttributeName.GetItemScreentip)

        Public Shared ReadOnly Property ItemSupertip As AttributeCategory = New AttributeCategory(27, NameOf(ItemSupertip), AttributeName.GetItemSupertip)

        Public Shared ReadOnly Property SelectedItemId As AttributeCategory = New AttributeCategory(28, NameOf(SelectedItemId), AttributeName.GetSelectedItemID)

        Public Shared ReadOnly Property SelectedItemIndex As AttributeCategory = New AttributeCategory(29, NameOf(SelectedItemIndex), AttributeName.GetSelectedItemIndex)

        Public Shared ReadOnly Property Columns As AttributeCategory = New AttributeCategory(30, NameOf(Columns), AttributeName.Columns)

        Public Shared ReadOnly Property Rows As AttributeCategory = New AttributeCategory(31, NameOf(Rows), AttributeName.Rows)

        Public Shared ReadOnly Property ItemHeight As AttributeCategory = New AttributeCategory(32, NameOf(ItemHeight), AttributeName.ItemHeight, AttributeName.GetItemHeight)

        Public Shared ReadOnly Property ItemWidth As AttributeCategory = New AttributeCategory(33, NameOf(ItemWidth), AttributeName.ItemWidth, AttributeName.GetItemWidth)
        
        Public Shared ReadOnly Property ItemSize As AttributeCategory = New AttributeCategory(34, NameOf(ItemSize), AttributeName.ItemSize)

        Public Shared ReadOnly Property StartFromScratch As AttributeCategory = New AttributeCategory(35, NameOf(StartFromScratch), AttributeName.StartFromScratch)

        Public Shared ReadOnly Property OnLoad As AttributeCategory = New AttributeCategory(36, NameOf(OnLoad), AttributeName.OnLoad)

        Public Shared ReadOnly Property LoadImage As AttributeCategory = New AttributeCategory(37, NameOf(LoadImage), AttributeName.LoadImage)

        Public Shared ReadOnly Property BoxStyle As AttributeCategory = New AttributeCategory(38, NameOf(BoxStyle), AttributeName.BoxStyle)

        Public Shared ReadOnly Property Title As AttributeCategory = New AttributeCategory(39, NameOf(Title), AttributeName.Title, AttributeName.GetTitle)

#End Region

#Region "Overrides and Equality Comparison"

        Public Overrides Function ToString() As String
            Return $"Category '{_name}': {String.Join(", ", _members)}"
        End Function

        Public Overloads Overrides Function Equals(obj As Object) As Boolean
            Return Equals(TryCast(obj, AttributeCategory))
        End Function

        Public Overloads Function Equals(other As AttributeCategory) As Boolean Implements IEquatable(Of AttributeCategory).Equals
            Return other IsNot Nothing AndAlso _id = other._id
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return _id
        End Function

        Public Shared Operator =(left as AttributeCategory, right as AttributeCategory) as Boolean
            Return Equals(left, right)
        End Operator

        Public Shared Operator <>(left as AttributeCategory, right as AttributeCategory) as Boolean
            Return Not Equals(left, right)
        End Operator

#End Region

    End Class
End NameSpace