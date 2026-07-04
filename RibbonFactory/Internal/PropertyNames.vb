Imports RibbonX.Utilities

Namespace InternalApi
    Friend NotInheritable Class Name

        Private Sub New()
        End Sub

        Private NotInheritable Class PropertyName
            Implements IPropertyName

            Private ReadOnly name As String
            Private ReadOnly hashCode As Integer

            Public Sub New(name As String)
                Me.name = NotNullOrWhiteSpace(name)
                hashCode = name.ToLower().GetHashCode()
            End Sub

            Public Overrides Function Equals(obj As Object) As Boolean
                Return Equals(TryCast(obj, IPropertyName))
            End Function

            Public Overloads Function Equals(other As IPropertyName) As Boolean Implements IEquatable(Of IPropertyName).Equals
                Return other IsNot Nothing AndAlso name.Equals(other.ToString())
            End Function

            Public Overrides Function GetHashCode() As Integer
                Return hashCode
            End Function

            Public Overrides Function ToString() As String
                Return name
            End Function

            Public Function ToXml() As String Implements IPropertyName.ToXml
                Return name.CamelCase()
            End Function

            Public Function Clone() As Object Implements ICloneable.Clone
                Return New PropertyName(name)
            End Function

        End Class

        Public Shared ReadOnly Property Enabled As IPropertyName = New PropertyName(NameOf(Enabled))

        Public Shared ReadOnly Property GetEnabled As IPropertyName = New PropertyName(NameOf(GetEnabled))

        Public Shared ReadOnly Property Id As IPropertyName = New PropertyName(NameOf(Id))

        Public Shared ReadOnly Property IdMso As IPropertyName = New PropertyName(NameOf(IdMso))

        Public Shared ReadOnly Property IdQ As IPropertyName = New PropertyName(NameOf(IdQ))

        Public Shared ReadOnly Property InsertAfterMso As IPropertyName = New PropertyName(NameOf(InsertAfterMso))

        Public Shared ReadOnly Property InsertAfterQ As IPropertyName = New PropertyName(NameOf(InsertAfterQ))

        Public Shared ReadOnly Property InsertBeforeMso As IPropertyName = New PropertyName(NameOf(InsertBeforeMso))

        Public Shared ReadOnly Property InsertBeforeQ As IPropertyName = New PropertyName(NameOf(InsertBeforeQ))

        Public Shared ReadOnly Property Tag As IPropertyName = New PropertyName(NameOf(Tag))

        Public Shared ReadOnly Property Visible As IPropertyName = New PropertyName(NameOf(Visible))

        Public Shared ReadOnly Property GetVisible As IPropertyName = New PropertyName(NameOf(GetVisible))

        Public Shared ReadOnly Property Label As IPropertyName = New PropertyName(NameOf(Label))

        Public Shared ReadOnly Property Screentip As IPropertyName = New PropertyName(NameOf(Screentip))

        Public Shared ReadOnly Property Supertip As IPropertyName = New PropertyName(NameOf(Supertip))

        Public Shared ReadOnly Property GetLabel As IPropertyName = New PropertyName(NameOf(GetLabel))

        Public Shared ReadOnly Property GetScreentip As IPropertyName = New PropertyName(NameOf(GetScreentip))

        Public Shared ReadOnly Property GetSupertip As IPropertyName = New PropertyName(NameOf(GetSupertip))

        Public Shared ReadOnly Property ShowLabel As IPropertyName = New PropertyName(NameOf(ShowLabel))

        Public Shared ReadOnly Property Keytip As IPropertyName = New PropertyName(NameOf(Keytip))

        Public Shared ReadOnly Property GetKeytip As IPropertyName = New PropertyName(NameOf(GetKeytip))

        Public Shared ReadOnly Property GetShowLabel As IPropertyName = New PropertyName(NameOf(GetShowLabel))

        Public Shared ReadOnly Property Image As IPropertyName = New PropertyName(NameOf(Image))

        Public Shared ReadOnly Property ImageMso As IPropertyName = New PropertyName(NameOf(ImageMso))

        Public Shared ReadOnly Property ShowImage As IPropertyName = New PropertyName(NameOf(ShowImage))

        Public Shared ReadOnly Property GetImage As IPropertyName = New PropertyName(NameOf(GetImage))

        Public Shared ReadOnly Property GetShowImage As IPropertyName = New PropertyName(NameOf(GetShowImage))

        Public Shared ReadOnly Property Description As IPropertyName = New PropertyName(NameOf(Description))

        Public Shared ReadOnly Property OnAction As IPropertyName = New PropertyName(NameOf(OnAction))

        Public Shared ReadOnly Property GetDescription As IPropertyName = New PropertyName(NameOf(GetDescription))

        Public Shared ReadOnly Property Size As IPropertyName = New PropertyName(NameOf(Size))

        Public Shared ReadOnly Property GetSize As IPropertyName = New PropertyName(NameOf(GetSize))

        Public Shared ReadOnly Property SizeString As IPropertyName = New PropertyName(NameOf(SizeString))

        Public Shared ReadOnly Property InvalidateContentOnDrop As IPropertyName = New PropertyName(NameOf(InvalidateContentOnDrop))

        Public Shared ReadOnly Property ShowItemImage As IPropertyName = New PropertyName(NameOf(ShowItemImage))

        Public Shared ReadOnly Property GetItemCount As IPropertyName = New PropertyName(NameOf(GetItemCount))

        Public Shared ReadOnly Property GetItemID As IPropertyName = New PropertyName(NameOf(GetItemID))

        Public Shared ReadOnly Property GetItemImage As IPropertyName = New PropertyName(NameOf(GetItemImage))

        Public Shared ReadOnly Property GetItemLabel As IPropertyName = New PropertyName(NameOf(GetItemLabel))

        Public Shared ReadOnly Property GetItemScreentip As IPropertyName = New PropertyName(NameOf(GetItemScreentip))

        Public Shared ReadOnly Property GetItemSupertip As IPropertyName = New PropertyName(NameOf(GetItemSupertip))

        Public Shared ReadOnly Property MaxLength As IPropertyName = New PropertyName(NameOf(MaxLength))

        Public Shared ReadOnly Property OnChange As IPropertyName = New PropertyName(NameOf(OnChange))

        Public Shared ReadOnly Property ShowItemLabel As IPropertyName = New PropertyName(NameOf(ShowItemLabel))

        Public Shared ReadOnly Property GetPressed As IPropertyName = New PropertyName(NameOf(GetPressed))

        Public Shared ReadOnly Property GetSelectedItemID As IPropertyName = New PropertyName(NameOf(GetSelectedItemID))

        Public Shared ReadOnly Property GetSelectedItemIndex As IPropertyName = New PropertyName(NameOf(GetSelectedItemIndex))

        Public Shared ReadOnly Property GetText As IPropertyName = New PropertyName(NameOf(GetText))

        Public Shared ReadOnly Property Columns As IPropertyName = New PropertyName(NameOf(Columns))

        Public Shared ReadOnly Property ItemHeight As IPropertyName = New PropertyName(NameOf(ItemHeight))

        Public Shared ReadOnly Property ItemWidth As IPropertyName = New PropertyName(NameOf(ItemWidth))

        Public Shared ReadOnly Property Rows As IPropertyName = New PropertyName(NameOf(Rows))

        Public Shared ReadOnly Property GetItemHeight As IPropertyName = New PropertyName(NameOf(GetItemHeight))

        Public Shared ReadOnly Property GetItemWidth As IPropertyName = New PropertyName(NameOf(GetItemWidth))

        Public Shared ReadOnly Property StartFromScratch As IPropertyName = New PropertyName(NameOf(StartFromScratch))

        Public Shared ReadOnly Property OnLoad As IPropertyName = New PropertyName(NameOf(OnLoad))

        Public Shared ReadOnly Property LoadImage As IPropertyName = New PropertyName(NameOf(LoadImage))

        Public Shared ReadOnly Property BoxStyle As IPropertyName = New PropertyName(NameOf(BoxStyle))

        Public Shared ReadOnly Property Title As IPropertyName = New PropertyName(NameOf(Title))

        Public Shared ReadOnly Property GetTitle As IPropertyName = New PropertyName(NameOf(GetTitle))

        Public Shared ReadOnly Property ItemSize As IPropertyName = New PropertyName(NameOf(ItemSize))

        Public Shared ReadOnly Property ShowInRibbon As IPropertyName = New PropertyName(NameOf(ShowInRibbon))
    End Class
End Namespace