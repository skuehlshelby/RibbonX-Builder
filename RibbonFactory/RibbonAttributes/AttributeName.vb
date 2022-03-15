Imports System.Reflection
Imports RibbonFactory.Utilities

Namespace RibbonAttributes

	''' <summary>
	''' The purpose of this class is to provide a single place where all attribute names are listed, along with their spelling.
	''' </summary>
	Friend NotInheritable Class AttributeName
		Implements IEquatable(Of AttributeName)

		Private Sub New(value As Byte, name As String)
			Me.Name = name
			Me.Value = value
		End Sub

#Region "Instance Members"

		Public ReadOnly Property Name As String

		Public ReadOnly Property Value As Byte

		Public Overrides Function ToString() As String
			Return Name
		End Function

#End Region

#Region "Equality Operations"

		Public Overloads Overrides Function Equals(other As Object) As Boolean
			Return Equals(TryCast(other, AttributeName))
		End Function

		Public Overloads Function Equals(other As AttributeName) As Boolean Implements IEquatable(Of AttributeName).Equals
			Return other IsNot Nothing AndAlso other.Value = Value
		End Function

		Public Overrides Function GetHashCode() As Integer
			Return Value
		End Function

		Public Shared Operator =(left As AttributeName, right As AttributeName) As Boolean
			Return Equals(left, right)
		End Operator

		Public Shared Operator <>(left As AttributeName, right As AttributeName) As Boolean
			Return Not Equals(left, right)
		End Operator

#End Region

		Public Shared Function Values() As IEnumerable(Of AttributeName)
			Return GetType(AttributeName).
				GetProperties(BindingFlags.Public Or BindingFlags.GetProperty Or BindingFlags.Static).
				Select(Function(info) info.GetValue(Nothing)).
				Cast(Of AttributeName)()
		End Function

		Public Shared Property Enabled As AttributeName                 = New AttributeName(value:=0, name:=NameOf(Enabled).CamelCase())

		Public Shared Property GetEnabled As AttributeName              = New AttributeName(value:=1, name:=NameOf(GetEnabled).CamelCase())

		Public Shared Property Id As AttributeName                      = New AttributeName(value:=2, name:=NameOf(Id).CamelCase())

		Public Shared Property IdMso As AttributeName                   = New AttributeName(value:=3, name:=NameOf(IdMso).CamelCase())

		Public Shared Property IdQ As AttributeName                     = New AttributeName(value:=4, name:=NameOf(IdQ).CamelCase())

		Public Shared Property InsertAfterMso As AttributeName          = New AttributeName(value:=5, name:=NameOf(InsertAfterMso).CamelCase())

		Public Shared Property InsertAfterQ As AttributeName            = New AttributeName(value:=6, name:=NameOf(InsertAfterQ).CamelCase())

		Public Shared Property InsertBeforeMso As AttributeName         = New AttributeName(value:=7, name:=NameOf(InsertBeforeMso).CamelCase())

		Public Shared Property InsertBeforeQ As AttributeName           = New AttributeName(value:=8, name:=NameOf(InsertBeforeQ).CamelCase())

		Public Shared Property Tag As AttributeName                     = New AttributeName(value:=9, name:=NameOf(Tag).CamelCase())

		Public Shared Property Visible As AttributeName                 = New AttributeName(value:=10, name:=NameOf(Visible).CamelCase())

		Public Shared Property GetVisible As AttributeName              = New AttributeName(value:=11, name:=NameOf(GetVisible).CamelCase())

		Public Shared Property Label As AttributeName                   = New AttributeName(value:=12, name:=NameOf(Label).CamelCase())

		Public Shared Property Screentip As AttributeName               = New AttributeName(value:=13, name:=NameOf(Screentip).CamelCase())

		Public Shared Property Supertip As AttributeName                = New AttributeName(value:=14, name:=NameOf(Supertip).CamelCase())

		Public Shared Property GetLabel As AttributeName                = New AttributeName(value:=15, name:=NameOf(GetLabel).CamelCase())

		Public Shared Property GetScreentip As AttributeName            = New AttributeName(value:=16, name:=NameOf(GetScreentip).CamelCase())

		Public Shared Property GetSupertip As AttributeName             = New AttributeName(value:=17, name:=NameOf(GetSupertip).CamelCase())

		Public Shared Property ShowLabel As AttributeName               = New AttributeName(value:=18, name:=NameOf(ShowLabel).CamelCase())

		Public Shared Property Keytip As AttributeName                  = New AttributeName(value:=19, name:=NameOf(Keytip).CamelCase())

		Public Shared Property GetKeytip As AttributeName               = New AttributeName(value:=20, name:=NameOf(GetKeytip).CamelCase())

		Public Shared Property GetShowLabel As AttributeName            = New AttributeName(value:=21, name:=NameOf(GetShowLabel).CamelCase())

		Public Shared Property Image As AttributeName                   = New AttributeName(value:=22, name:=NameOf(Image).CamelCase())

		Public Shared Property ImageMso As AttributeName                = New AttributeName(value:=23, name:=NameOf(ImageMso).CamelCase())

		Public Shared Property ShowImage As AttributeName               = New AttributeName(value:=24, name:=NameOf(ShowImage).CamelCase())

		Public Shared Property GetImage As AttributeName                = New AttributeName(value:=25, name:=NameOf(GetImage).CamelCase())

		Public Shared Property GetShowImage As AttributeName            = New AttributeName(value:=26, name:=NameOf(GetShowImage).CamelCase())

		Public Shared Property Description As AttributeName             = New AttributeName(value:=27, name:=NameOf(Description).CamelCase())

		Public Shared Property OnAction As AttributeName                = New AttributeName(value:=28, name:=NameOf(OnAction).CamelCase())

		Public Shared Property GetDescription As AttributeName          = New AttributeName(value:=29, name:=NameOf(GetDescription).CamelCase())

		Public Shared Property Size As AttributeName                    = New AttributeName(value:=30, name:=NameOf(Size).CamelCase())

		Public Shared Property GetSize As AttributeName                 = New AttributeName(value:=31, name:=NameOf(GetSize).CamelCase())

		Public Shared Property SizeString As AttributeName              = New AttributeName(value:=32, name:=NameOf(SizeString).CamelCase())

		Public Shared Property InvalidateContentOnDrop As AttributeName = New AttributeName(value:=33, name:=NameOf(InvalidateContentOnDrop).CamelCase())

		Public Shared Property ShowItemImage As AttributeName           = New AttributeName(value:=34, name:=NameOf(ShowItemImage).CamelCase())

		Public Shared Property GetItemCount As AttributeName            = New AttributeName(value:=35, name:=NameOf(GetItemCount).CamelCase())

		Public Shared Property GetItemID As AttributeName               = New AttributeName(value:=36, name:=NameOf(GetItemID).CamelCase())

		Public Shared Property GetItemImage As AttributeName            = New AttributeName(value:=37, name:=NameOf(GetItemImage).CamelCase())

		Public Shared Property GetItemLabel As AttributeName            = New AttributeName(value:=38, name:=NameOf(GetItemLabel).CamelCase())

		Public Shared Property GetItemScreentip As AttributeName        = New AttributeName(value:=39, name:=NameOf(GetItemScreentip).CamelCase())

		Public Shared Property GetItemSupertip As AttributeName         = New AttributeName(value:=40, name:=NameOf(GetItemSupertip).CamelCase())

		Public Shared Property MaxLength As AttributeName               = New AttributeName(value:=41, name:=NameOf(MaxLength).CamelCase())

		Public Shared Property OnChange As AttributeName                = New AttributeName(value:=42, name:=NameOf(OnChange).CamelCase())

		Public Shared Property ShowItemLabel As AttributeName           = New AttributeName(value:=43, name:=NameOf(ShowItemLabel).CamelCase())

		Public Shared Property GetPressed As AttributeName              = New AttributeName(value:=44, name:=NameOf(GetPressed).CamelCase())

		Public Shared Property GetSelectedItemID As AttributeName       = New AttributeName(value:=45, name:=NameOf(GetSelectedItemID).CamelCase())

		Public Shared Property GetSelectedItemIndex As AttributeName    = New AttributeName(value:=46, name:=NameOf(GetSelectedItemIndex).CamelCase())

		Public Shared Property GetText As AttributeName                 = New AttributeName(value:=47, name:=NameOf(GetText).CamelCase())

		Public Shared Property Columns As AttributeName                 = New AttributeName(value:=48, name:=NameOf(Columns).CamelCase())

		Public Shared Property ItemHeight As AttributeName              = New AttributeName(value:=49, name:=NameOf(ItemHeight).CamelCase())

		Public Shared Property ItemWidth As AttributeName               = New AttributeName(value:=50, name:=NameOf(ItemWidth).CamelCase())

		Public Shared Property Rows As AttributeName                    = New AttributeName(value:=51, name:=NameOf(Rows).CamelCase())

		Public Shared Property GetItemHeight As AttributeName           = New AttributeName(value:=52, name:=NameOf(GetItemHeight).CamelCase())

		Public Shared Property GetItemWidth As AttributeName            = New AttributeName(value:=53, name:=NameOf(GetItemWidth).CamelCase())

		Public Shared Property StartFromScratch As AttributeName        = New AttributeName(value:=54, name:=NameOf(StartFromScratch).CamelCase())

		Public Shared Property OnLoad As AttributeName                  = New AttributeName(value:=55, name:=NameOf(OnLoad).CamelCase())

		Public Shared Property LoadImage As AttributeName               = New AttributeName(value:=56, name:=NameOf(LoadImage).CamelCase())

		Public Shared Property BoxStyle As AttributeName                = New AttributeName(value:=57, name:=NameOf(BoxStyle).CamelCase())

		Public Shared Property Title As AttributeName                   = New AttributeName(value:=58, name:=NameOf(Title).CamelCase())

		Public Shared Property GetTitle As AttributeName                = New AttributeName(value:=59, name:=NameOf(GetTitle).CamelCase())

		Public Shared Property ItemSize As AttributeName                = New AttributeName(value:=60, name:=NameOf(ItemSize).CamelCase())

		Public Shared Property ShowInRibbon As AttributeName            = New AttributeName(value:=61, name:=NameOf(ShowInRibbon).CamelCase())

		Public Shared Property FrameworkBeforeEvent As AttributeName    = New AttributeName(value:=62, name:=NameOf(FrameworkBeforeEvent))

		Public Shared Property FrameworkOnEvent As AttributeName		= New AttributeName(value:=62, name:=NameOf(FrameworkOnEvent))

	End Class

End Namespace