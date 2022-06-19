Namespace Controls.Attributes.Internal

    Friend Class ValueWrapper(Of T)
        Implements ICloneable

        Public Sub New(value As T)
            Me.Value = value
        End Sub

        Public Property Value As T

        Public Function Clone() As ValueWrapper(Of T)
            Dim valueCopy As T = If(TypeOf Value Is ICloneable, DirectCast(DirectCast(Value, ICloneable).Clone(), T), Value)

            Return New ValueWrapper(Of T)(valueCopy)
        End Function

        Private Function CloneII() As Object Implements ICloneable.Clone
            Return Clone()
        End Function
    End Class

End Namespace

