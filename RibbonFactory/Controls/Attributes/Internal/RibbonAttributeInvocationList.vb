Namespace Controls.Attributes.Internal

    <DebuggerDisplay("{GetDebuggerDisplay(),nq}")>
    Friend Class RibbonAttributeInvocationList
        Implements IRibbonAttributeReadOnly(Of ICollection(Of EventHandler))
        Implements ICloneable

        Public Event ValueChanged() Implements IRibbonAttribute.ValueChanged

        Private ReadOnly _value As ICollection(Of EventHandler)

        Public Sub New(name As AttributeName, category As AttributeCategory)
            Me.New(Nothing, name, category)
        End Sub

        Public Sub New(handlers As ICollection(Of EventHandler), name As AttributeName, category As AttributeCategory)
            _value = If(handlers, New List(Of EventHandler)())
            Me.Name = name
            Me.Category = category
        End Sub

        Public ReadOnly Property Name As AttributeName Implements IRibbonAttribute.Name

        Public ReadOnly Property Category As AttributeCategory Implements IRibbonAttribute.Category

        Public Function GetValue() As ICollection(Of EventHandler) Implements IRibbonAttributeReadOnly(Of ICollection(Of EventHandler)).GetValue
            Return _value
        End Function

        Public Overrides Function ToString() As String
            Return String.Empty
        End Function

        Private Function GetDebuggerDisplay() As String
            Return $"{NameOf(RibbonAttributeInvocationList)} ({Name}): {_value.Count} subscribers"
        End Function

        Private Function Clone() As Object Implements ICloneable.Clone
            Dim copy As EventHandler()
            ReDim copy(0 To _value.Count - 1)

            _value.CopyTo(copy, 0)

            Return New RibbonAttributeInvocationList(copy.ToList(), Name.Clone(), Category.Clone())
        End Function
    End Class


    <DebuggerDisplay("{GetDebuggerDisplay(),nq}")>
    Friend Class RibbonAttributeInvocationList(Of T)
        Implements IRibbonAttributeReadOnly(Of ICollection(Of EventHandler(Of T)))
        Implements ICloneable

        Public Event ValueChanged() Implements IRibbonAttribute.ValueChanged

        Private ReadOnly _value As ICollection(Of EventHandler(Of T))

        Public Sub New(name As AttributeName, category As AttributeCategory)
            Me.New(Nothing, name, category)
        End Sub

        Public Sub New(handlers As ICollection(Of EventHandler(Of T)), name As AttributeName, category As AttributeCategory)
            _value = If(handlers, New List(Of EventHandler(Of T))())
            Me.Name = name
            Me.Category = category
        End Sub

        Public ReadOnly Property Name As AttributeName Implements IRibbonAttribute.Name

        Public ReadOnly Property Category As AttributeCategory Implements IRibbonAttribute.Category

        Public Function GetValue() As ICollection(Of EventHandler(Of T)) Implements IRibbonAttributeReadOnly(Of ICollection(Of EventHandler(Of T))).GetValue
            Return _value
        End Function

        Public Overrides Function ToString() As String
            Return String.Empty
        End Function

        Private Function GetDebuggerDisplay() As String
            Return $"{NameOf(RibbonAttributeInvocationList(Of T))} ({Name}): {_value.Count} subscribers"
        End Function

        Private Function Clone() As Object Implements ICloneable.Clone
            Dim copy As EventHandler(Of T)()
            ReDim copy(0 To _value.Count - 1)

            _value.CopyTo(copy, 0)

            Return New RibbonAttributeInvocationList(Of T)(copy.ToList(), Name.Clone(), Category.Clone())
        End Function

    End Class

End Namespace

