Imports RibbonX.Images
Imports RibbonX.SimpleTypes

Namespace InternalApi

    Friend NotInheritable Class Category
        Private Sub New()
        End Sub

        Private Class PropertyCategory
            Implements IPropertyCategory

            Private ReadOnly name As String
            Private ReadOnly members As ICollection(Of IPropertyName)

            Public Sub New(name As String, ParamArray members() As IPropertyName)
                Me.name = name
                Me.members = members
            End Sub

            Public ReadOnly Property Count As Integer Implements IReadOnlyCollection(Of IPropertyName).Count
                Get
                    Return members.Count
                End Get
            End Property

            Public Overrides Function Equals(obj As Object) As Boolean
                Return Equals(TryCast(obj, IPropertyCategory))
            End Function

            Public Overloads Function Equals(other As IPropertyCategory) As Boolean Implements IEquatable(Of IPropertyCategory).Equals
                Return other IsNot Nothing AndAlso members.Intersect(other).Count() = Count
            End Function

            Public Overrides Function GetHashCode() As Integer
                Return members.Select(Function(m) m.GetHashCode()).Aggregate(Function(a, b) a Xor b)
            End Function

            Public Overrides Function ToString() As String
                Return name
            End Function

            Public Function GetEnumerator() As IEnumerator(Of IPropertyName) Implements IEnumerable(Of IPropertyName).GetEnumerator
                Return members.GetEnumerator()
            End Function

            Public Function Clone() As Object Implements ICloneable.Clone
                Return New PropertyCategory(name, members.ToArray())
            End Function

            Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
                Return members.GetEnumerator()
            End Function
        End Class

        Private NotInheritable Class PropertyCategory(Of T)
            Inherits PropertyCategory
            Implements IPropertyCategory(Of T)

            Public Sub New(name As String, ParamArray members() As IPropertyName)
                MyBase.New(name, members)
            End Sub
        End Class

        Public Shared ReadOnly Property Enabled As IPropertyCategory(Of Boolean) = New PropertyCategory(Of Boolean)(NameOf(Enabled), Name.Enabled, Name.GetEnabled)

        Public Shared ReadOnly Property IdType As IPropertyCategory(Of String) = New PropertyCategory(Of String)(NameOf(IdType), Name.Id, Name.IdMso, Name.IdQ)

        Public Shared ReadOnly Property Insertion As IPropertyCategory = New PropertyCategory(NameOf(Insertion), Name.InsertAfterMso, Name.InsertAfterQ, Name.InsertBeforeQ, Name.InsertBeforeMso)

        Public Shared ReadOnly Property Visibility As IPropertyCategory(Of Boolean) = New PropertyCategory(Of Boolean)(NameOf(Visibility), Name.Visible, Name.GetVisible)

        Public Shared ReadOnly Property Label As IPropertyCategory(Of String) = New PropertyCategory(Of String)(NameOf(Label), Name.Label, Name.GetLabel)

        Public Shared ReadOnly Property LabelVisibility As IPropertyCategory(Of Boolean) = New PropertyCategory(Of Boolean)(NameOf(LabelVisibility), Name.ShowLabel, Name.GetShowLabel)

        Public Shared ReadOnly Property ScreenTip As IPropertyCategory(Of String) = New PropertyCategory(Of String)(NameOf(ScreenTip), Name.Screentip, Name.GetScreentip)

        Public Shared ReadOnly Property SuperTip As IPropertyCategory(Of String) = New PropertyCategory(Of String)(NameOf(SuperTip), Name.Supertip, Name.GetSupertip)

        Public Shared ReadOnly Property Description As IPropertyCategory(Of String) = New PropertyCategory(Of String)(NameOf(Description), Name.Description, Name.GetDescription)

        Public Shared ReadOnly Property KeyTip As IPropertyCategory(Of KeyTip) = New PropertyCategory(Of KeyTip)(NameOf(KeyTip), Name.Keytip, Name.GetKeytip)

        Public Shared ReadOnly Property Image As IPropertyCategory(Of RibbonImage) = New PropertyCategory(Of RibbonImage)(NameOf(Image), Name.Image, Name.ImageMso, Name.GetImage)

        Public Shared ReadOnly Property ImageVisibility As IPropertyCategory(Of Boolean) = New PropertyCategory(Of Boolean)(NameOf(ImageVisibility), Name.ShowImage, Name.GetShowImage)

        Public Shared ReadOnly Property OnAction As IPropertyCategory = New PropertyCategory(NameOf(OnAction), Name.OnAction)

        Public Shared ReadOnly Property OnChange As IPropertyCategory = New PropertyCategory(NameOf(OnChange), Name.OnChange)

        Public Shared ReadOnly Property Text As IPropertyCategory(Of String) = New PropertyCategory(Of String)(NameOf(Text), Name.GetText)

        Public Shared ReadOnly Property Pressed As IPropertyCategory(Of Boolean) = New PropertyCategory(Of Boolean)(NameOf(Pressed), Name.GetPressed)

        Public Shared ReadOnly Property Size As IPropertyCategory(Of ControlSize) = New PropertyCategory(Of ControlSize)(NameOf(Size), Name.Size, Name.GetSize)

        Public Shared ReadOnly Property MaxLength As IPropertyCategory(Of Byte) = New PropertyCategory(Of Byte)(NameOf(MaxLength), Name.MaxLength)

        Public Shared ReadOnly Property SizeString As IPropertyCategory = New PropertyCategory(NameOf(SizeString), Name.SizeString)

        Public Shared ReadOnly Property ContentInvalidation As IPropertyCategory = New PropertyCategory(NameOf(ContentInvalidation), Name.InvalidateContentOnDrop)

        Public Shared ReadOnly Property ItemImageVisibility As IPropertyCategory = New PropertyCategory(NameOf(ItemImageVisibility), Name.ShowItemImage)

        Public Shared ReadOnly Property ItemLabelVisibility As IPropertyCategory = New PropertyCategory(NameOf(ItemLabelVisibility), Name.ShowItemLabel)

        Public Shared ReadOnly Property ItemCount As IPropertyCategory = New PropertyCategory(NameOf(ItemCount), Name.GetItemCount)

        Public Shared ReadOnly Property ItemId As IPropertyCategory(Of String) = New PropertyCategory(Of String)(NameOf(ItemId), Name.GetItemID)

        Public Shared ReadOnly Property ItemImage As IPropertyCategory = New PropertyCategory(NameOf(ItemImage), Name.GetItemImage)

        Public Shared ReadOnly Property ItemLabel As IPropertyCategory = New PropertyCategory(NameOf(ItemLabel), Name.GetItemLabel)

        Public Shared ReadOnly Property ItemScreentip As IPropertyCategory = New PropertyCategory(NameOf(ItemScreentip), Name.GetItemScreentip)

        Public Shared ReadOnly Property ItemSupertip As IPropertyCategory = New PropertyCategory(NameOf(ItemSupertip), Name.GetItemSupertip)

        Public Shared ReadOnly Property SelectedItemPosition As IPropertyCategory(Of IItem) = New PropertyCategory(Of IItem)(NameOf(SelectedItemPosition), Name.GetSelectedItemIndex, Name.GetSelectedItemID)

        Public Shared ReadOnly Property Columns As IPropertyCategory(Of Integer) = New PropertyCategory(Of Integer)(NameOf(Columns), Name.Columns)

        Public Shared ReadOnly Property Rows As IPropertyCategory(Of Integer) = New PropertyCategory(Of Integer)(NameOf(Rows), Name.Rows)

        Public Shared ReadOnly Property ItemHeight As IPropertyCategory(Of Integer) = New PropertyCategory(Of Integer)(NameOf(ItemHeight), Name.ItemHeight, Name.GetItemHeight)

        Public Shared ReadOnly Property ItemWidth As IPropertyCategory(Of Integer) = New PropertyCategory(Of Integer)(NameOf(ItemWidth), Name.ItemWidth, Name.GetItemWidth)

        Public Shared ReadOnly Property ItemSize As IPropertyCategory(Of ControlSize) = New PropertyCategory(Of ControlSize)(NameOf(ItemSize), Name.ItemSize)

        Public Shared ReadOnly Property StartFromScratch As IPropertyCategory(Of Boolean) = New PropertyCategory(Of Boolean)(NameOf(StartFromScratch), Name.StartFromScratch)

        Public Shared ReadOnly Property OnLoad As IPropertyCategory = New PropertyCategory(NameOf(OnLoad), Name.OnLoad)

        Public Shared ReadOnly Property LoadImage As IPropertyCategory = New PropertyCategory(NameOf(LoadImage), Name.LoadImage)

        Public Shared ReadOnly Property BoxStyle As IPropertyCategory(Of BoxStyle) = New PropertyCategory(Of BoxStyle)(NameOf(BoxStyle), Name.BoxStyle)

        Public Shared ReadOnly Property Title As IPropertyCategory(Of String) = New PropertyCategory(Of String)(NameOf(Title), Name.Title, Name.GetTitle)

        Public Shared ReadOnly Property ShowInRibbon As IPropertyCategory = New PropertyCategory(NameOf(ShowInRibbon), Name.ShowInRibbon)
    End Class

End Namespace