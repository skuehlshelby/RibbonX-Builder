Imports System.Text
Imports NUnit.Framework

Namespace RibbonFactoryTests

    Public Class Tests

        <SetUp>
        Public Sub Setup()
        End Sub

        <Test>
        Public Sub Test1()
            Assert.AreEqual(RibbonFactory.RibbonElement.FabricateID(Of RibbonFactory.Button), "Button1")
        End Sub


        <Test>
        Public Sub AttributeGroupEquality()
            Assert.AreEqual(Label.Label, Label.GetLabel)
            Assert.AreNotEqual(Label.GetLabel, Screentip.GetScreentip)
            Assert.AreEqual(Screentip.GetScreentip, Screentip.Screentip)
        End Sub

        <Test>
        Public Sub ControlListAddsWithReplacement()
            With New ControlList() From {New RibbonFactory.ControlProperty(Of String)(RibbonFactory.Attributes.Label.GetLabel, "My Label")}

            End With
        End Sub
    End Class

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
            If obj.GetHashCode = Me.GetHashCode Then
                Return obj IsNot Nothing AndAlso TypeOf obj Is T
            Else
                Return False
            End If
        End Function

        Public Overrides Function ToString() As String
            Return AttributeName
        End Function
    End Class

    Public NotInheritable Class Label
        Inherits AttributeGroup(Of Label)
        Private Sub New(ByVal AttributeName As String)
            MyBase.New(AttributeName)
        End Sub
        Public Shared ReadOnly Property Label As Label = New Label(NameOf(Label))
        Public Shared ReadOnly Property GetLabel As Label = New Label(NameOf(GetLabel))
    End Class

    Public NotInheritable Class Screentip
        Inherits AttributeGroup(Of Screentip)
        Private Sub New(ByVal AttributeName As String)
            MyBase.New(AttributeName)
        End Sub
        Public Shared ReadOnly Property Screentip As Screentip = New Screentip(NameOf(Screentip))
        Public Shared ReadOnly Property GetScreentip As Screentip = New Screentip(NameOf(GetScreentip))
    End Class

    Friend NotInheritable Class ControlList
        Private This As HashSet(Of RibbonFactory.IControlProperty) = New HashSet(Of RibbonFactory.IControlProperty)

        Public ReadOnly Property Count As Integer
            Get
                Return This.Count
            End Get
        End Property
        Public Function Add(item As RibbonFactory.IControlProperty) As Boolean
            If This.Contains(item) Then
                This.Remove(item)
                Return This.Add(item)
            Else
                Return This.Add(item)
            End If
        End Function
        Public Sub Clear()
            This.Clear()
        End Sub

        Public Sub CopyTo(array() As RibbonFactory.IControlProperty, arrayIndex As Integer)
            This.CopyTo(array, arrayIndex)
        End Sub

        Public Function Contains(item As RibbonFactory.IControlProperty) As Boolean
            Return This.Contains(item)
        End Function

        Public Function Remove(item As RibbonFactory.IControlProperty) As Boolean
            Return This.Remove(item)
        End Function

        Public Function GetEnumerator() As IEnumerator(Of RibbonFactory.IControlProperty)
            Return This.GetEnumerator
        End Function
    End Class
End Namespace