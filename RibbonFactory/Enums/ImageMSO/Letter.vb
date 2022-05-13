Namespace Enums.ImageMSO

    Public NotInheritable Class Letter
        Inherits ImageMSO

        Private Sub New(value As Integer, name As String)
            MyBase.New(value, name)
        End Sub

        Public Shared ReadOnly Property A As Letter = New Letter(1, NameOf(A))
        Public Shared ReadOnly Property B As Letter = New Letter(2, NameOf(B))
        Public Shared ReadOnly Property C As Letter = New Letter(3, NameOf(C))
        Public Shared ReadOnly Property D As Letter = New Letter(4, NameOf(D))
        Public Shared ReadOnly Property E As Letter = New Letter(5, NameOf(E))
        Public Shared ReadOnly Property F As Letter = New Letter(6, NameOf(F))
        Public Shared ReadOnly Property G As Letter = New Letter(7, NameOf(G))
        Public Shared ReadOnly Property H As Letter = New Letter(8, NameOf(H))
        Public Shared ReadOnly Property I As Letter = New Letter(9, NameOf(I))
        Public Shared ReadOnly Property J As Letter = New Letter(10, NameOf(J))
        Public Shared ReadOnly Property K As Letter = New Letter(11, NameOf(K))
        Public Shared ReadOnly Property L As Letter = New Letter(12, NameOf(L))
        Public Shared ReadOnly Property M As Letter = New Letter(13, NameOf(M))
        Public Shared ReadOnly Property N As Letter = New Letter(14, NameOf(N))
        Public Shared ReadOnly Property O As Letter = New Letter(15, NameOf(O))
        Public Shared ReadOnly Property P As Letter = New Letter(16, NameOf(P))
        Public Shared ReadOnly Property Q As Letter = New Letter(17, NameOf(Q))
        Public Shared ReadOnly Property R As Letter = New Letter(18, NameOf(R))
        Public Shared ReadOnly Property S As Letter = New Letter(19, NameOf(S))
        Public Shared ReadOnly Property T As Letter = New Letter(20, NameOf(T))
        Public Shared ReadOnly Property U As Letter = New Letter(21, NameOf(U))
        Public Shared ReadOnly Property V As Letter = New Letter(22, NameOf(V))
        Public Shared ReadOnly Property W As Letter = New Letter(23, NameOf(W))
        Public Shared ReadOnly Property X As Letter = New Letter(24, NameOf(X))
        Public Shared ReadOnly Property Y As Letter = New Letter(25, NameOf(Y))
        Public Shared ReadOnly Property Z As Letter = New Letter(26, NameOf(Z))

        Public Overrides Function Clone() As Object
            Return New Letter(value, name)
        End Function

    End Class

End Namespace