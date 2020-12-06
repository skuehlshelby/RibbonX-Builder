Imports System.Text

Namespace Attributes
    Public MustInherit Class AttributeGroup
        Public MustOverride Overrides Function GetHashCode() As Integer
        Public MustOverride Overrides Function Equals(obj As Object) As Boolean
        Public MustOverride Overrides Function ToString() As String

        Public Overloads Shared Operator =(ByVal Left As AttributeGroup, ByVal Right As AttributeGroup) As Boolean
            Return Left.Equals(Right)
        End Operator

        Public Overloads Shared Operator =(ByVal Left As AttributeGroup, ByVal Right As Object) As Boolean
            Return Left.Equals(Right)
        End Operator

        Public Overloads Shared Operator <>(ByVal Left As AttributeGroup, ByVal Right As AttributeGroup) As Boolean
            Return Not Left.Equals(Right)
        End Operator

        Public Overloads Shared Operator <>(ByVal Left As AttributeGroup, ByVal Right As Object) As Boolean
            Return Not Left.Equals(Right)
        End Operator

        Protected Shared Function CamelCase(ByVal Input As String) As String
            Dim Buffer As StringBuilder = New StringBuilder

            For Each Word As String In SplitByCase(Input)
                Buffer.Append(CapitalizeOnlyFirstLetter(Word))
            Next Word

            If Buffer.Length > 0 Then
                Buffer.Chars(0) = Char.ToLower(Buffer.Chars(0))
            End If

            Return Buffer.ToString
        End Function

        Private Shared Function SplitByCase(ByVal Input As String) As List(Of String)
            SplitByCase = New List(Of String)
            Dim Buffer As StringBuilder = New StringBuilder
            Dim LastLetter As Char = "A"c

            For Each Letter As Char In Input
                If Char.IsUpper(Letter) AndAlso Char.IsLower(LastLetter) Then
                    SplitByCase.Add(Buffer.ToString)
                    Buffer = New StringBuilder
                End If

                Buffer.Append(Letter)
                LastLetter = Letter
            Next Letter

            SplitByCase.Add(Buffer.ToString)
        End Function

        Private Shared Function CapitalizeOnlyFirstLetter(ByVal Input As String) As String
            If Input.Length > 0 Then
                Return Char.ToUpper(Input.ElementAt(0)) & LCase$(Input.Substring(1))
            Else
                Return Input
            End If
        End Function
    End Class

    Public MustInherit Class AttributeGroup(Of T As AttributeGroup)
        Inherits AttributeGroup

        Protected Sub New(ByVal AttributeName As String)
            Me.AttributeName = CamelCase(AttributeName)
        End Sub

        Public ReadOnly Property AttributeName As String

        Public Overrides Function GetHashCode() As Integer
            Return GetType(T).GetHashCode
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            If obj.GetHashCode = GetHashCode() Then
                Return TypeOf obj Is T
            Else
                Return False
            End If
        End Function

        Public Overrides Function ToString() As String
            Return AttributeName
        End Function
    End Class

    Friend NotInheritable Class Enabled
        Inherits AttributeGroup(Of Enabled)
        Protected Sub New(AttributeName As String)
            MyBase.New(AttributeName)
        End Sub
        Friend Shared ReadOnly Property Enabled As Enabled = New Enabled(NameOf(Enabled))
        Friend Shared ReadOnly Property GetEnabled As Enabled = New Enabled(NameOf(GetEnabled))
    End Class

    Friend NotInheritable Class Visible
        Inherits AttributeGroup(Of Visible)
        Protected Sub New(AttributeName As String)
            MyBase.New(AttributeName)
        End Sub
        Friend Shared ReadOnly Property Visible As Visible = New Visible(NameOf(Visible))
        Friend Shared ReadOnly Property GetVisible As Visible = New Visible(NameOf(GetVisible))
    End Class

    Friend NotInheritable Class Description
        Inherits AttributeGroup(Of Visible)
        Protected Sub New(AttributeName As String)
            MyBase.New(AttributeName)
        End Sub
        Friend Shared ReadOnly Property Description As Description = New Description(NameOf(Description))
        Friend Shared ReadOnly Property GetDescription As Description = New Description(NameOf(GetDescription))
    End Class

    Friend NotInheritable Class Position
        Inherits AttributeGroup(Of Position)
        Protected Sub New(AttributeName As String)
            MyBase.New(AttributeName)
        End Sub
        Friend Shared ReadOnly Property InsertAfterMso As Position = New Position(NameOf(InsertAfterMso))
        Friend Shared ReadOnly Property InsertBeforeMso As Position = New Position(NameOf(InsertBeforeMso))
        Friend Shared ReadOnly Property InsertAfterQ As Position = New Position(NameOf(InsertAfterQ))
        Friend Shared ReadOnly Property InsertBeforeQ As Position = New Position(NameOf(InsertBeforeQ))
    End Class

    Friend NotInheritable Class Label
        Inherits AttributeGroup(Of Label)
        Protected Sub New(AttributeName As String)
            MyBase.New(AttributeName)
        End Sub
        Friend Shared ReadOnly Property Label As Label = New Label(NameOf(Label))
        Friend Shared ReadOnly Property GetLabel As Label = New Label(NameOf(GetLabel))
    End Class

    Friend NotInheritable Class ShowLabel
        Inherits AttributeGroup(Of ShowLabel)
        Protected Sub New(AttributeName As String)
            MyBase.New(AttributeName)
        End Sub
        Friend Shared ReadOnly Property ShowLabel As ShowLabel = New ShowLabel(NameOf(ShowLabel))
        Friend Shared ReadOnly Property GetShowLabel As ShowLabel = New ShowLabel(NameOf(GetShowLabel))
    End Class

    Friend NotInheritable Class Screentip
        Inherits AttributeGroup(Of Screentip)
        Protected Sub New(AttributeName As String)
            MyBase.New(AttributeName)
        End Sub
        Friend Shared ReadOnly Property Screentip As Screentip = New Screentip(NameOf(Screentip))
        Friend Shared ReadOnly Property GetScreentip As Screentip = New Screentip(NameOf(GetScreentip))
    End Class

    Friend NotInheritable Class Supertip
        Inherits AttributeGroup(Of Supertip)
        Protected Sub New(AttributeName As String)
            MyBase.New(AttributeName)
        End Sub
        Friend Shared ReadOnly Property Supertip As Supertip = New Supertip(NameOf(Supertip))
        Friend Shared ReadOnly Property GetSupertip As Supertip = New Supertip(NameOf(GetSupertip))
    End Class

    Friend NotInheritable Class Size
        Inherits AttributeGroup(Of Size)
        Protected Sub New(AttributeName As String)
            MyBase.New(AttributeName)
        End Sub
        Friend Shared ReadOnly Property Size As Size = New Size(NameOf(Size))
        Friend Shared ReadOnly Property GetSize As Size = New Size(NameOf(GetSize))
    End Class

    Friend NotInheritable Class Box
        Inherits AttributeGroup(Of Box)
        Protected Sub New(AttributeName As String)
            MyBase.New(AttributeName)
        End Sub
        Friend Shared ReadOnly Property BoxStyle As Box = New Box(NameOf(BoxStyle))
    End Class

    Friend NotInheritable Class Image
        Inherits AttributeGroup(Of Image)
        Protected Sub New(AttributeName As String)
            MyBase.New(AttributeName)
        End Sub
        Friend Shared ReadOnly Property Image As Image = New Image(NameOf(Image))
        Friend Shared ReadOnly Property ImageMso As Image = New Image(NameOf(ImageMso))
        Friend Shared ReadOnly Property GetImage As Image = New Image(NameOf(GetImage))
    End Class

    Friend NotInheritable Class ShowImage
        Inherits AttributeGroup(Of ShowImage)
        Protected Sub New(AttributeName As String)
            MyBase.New(AttributeName)
        End Sub
        Friend Shared ReadOnly Property ShowImage As ShowImage = New ShowImage(NameOf(ShowImage))
    End Class

    Friend NotInheritable Class Action
        Inherits AttributeGroup(Of Action)
        Protected Sub New(AttributeName As String)
            MyBase.New(AttributeName)
        End Sub
        Friend Shared ReadOnly Property OnAction As Action = New Action(NameOf(OnAction))
    End Class

    Friend NotInheritable Class Toggleable
        Inherits AttributeGroup(Of Toggleable)
        Protected Sub New(AttributeName As String)
            MyBase.New(AttributeName)
        End Sub
        Friend Shared ReadOnly Property GetPressed As Toggleable = New Toggleable(NameOf(GetPressed))
    End Class

    Friend NotInheritable Class Text
        Inherits AttributeGroup(Of Text)
        Protected Sub New(AttributeName As String)
            MyBase.New(AttributeName)
        End Sub
        Friend Shared ReadOnly Property GetText As Text = New Text(NameOf(GetText))
        Friend Shared ReadOnly Property OnChange As Text = New Text(NameOf(OnChange))
        Friend Shared ReadOnly Property MaxLength As Text = New Text(NameOf(MaxLength))
        Friend Shared ReadOnly Property SizeString As Text = New Text(NameOf(SizeString))
    End Class
End Namespace