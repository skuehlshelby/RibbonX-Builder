Imports System.ComponentModel
Imports RibbonX

<TestClass()>
Public Class BindingTests
    Const Expected As String = "New Description"
    Private button As IButton
    Private thingy As Thing

    <TestInitialize>
    Public Sub Setup()
        button = RxApi.Button(Sub(b) b.WithDescription("Original Description On Button", AddressOf TestBase.GetDescriptionShared))
        thingy = New Thing() With {
            .Description = "Original Description On Thing"
        }
    End Sub

    Private Class Thing
        Implements INotifyPropertyChanged

        Private _description As String = String.Empty

        Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

        Public Property Description As String
            Get
                Return _description
            End Get
            Set(value As String)
                If Not Equals(_description, value) Then
                    _description = value
                    RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(NameOf(Description)))
                End If
            End Set
        End Property

        Public Sub SetDescription(value As String)
            Description = value
        End Sub

        Public Sub SetDescription(value As String, value2 As String)
            Description = value & value2
        End Sub
    End Class

    <TestMethod("Can Bind IButton.Description->Thing.Description")>
    Public Sub CanBindFullPropertyOnClassWhichImplements()
        Assert.AreNotEqual(button.Description, thingy.Description)

        button.Bind(thingy, Sub(btn, thng) thng.Description = btn.Description)

        Assert.AreEqual(button.Description, thingy.Description)

        button.Description = Expected

        Assert.AreEqual(Expected, thingy.Description)
        Assert.AreEqual(button.Description, thingy.Description)
    End Sub

    <TestMethod("Can Bind Thing.Description->IButton.Description")>
    Public Sub BindingTestII()
        Assert.AreNotEqual(button.Description, thingy.Description)

        button.Bind(thingy, Sub(btn, r) btn.Description = r.Description)

        Assert.AreEqual(button.Description, thingy.Description)

        thingy.Description = Expected

        Assert.AreEqual(Expected, thingy.Description)
        Assert.AreEqual(Expected, thingy.Description)
    End Sub

    <TestMethod("Can Bind Thing.SetDescription(IButton.Description)")>
    Public Sub BindingTestIII()
        Assert.AreNotEqual(button.Description, thingy.Description)

        button.Bind(thingy, Sub(btn, r) r.SetDescription(btn.Description))

        Assert.AreEqual(button.Description, thingy.Description)

        button.Description = Expected

        Assert.AreEqual(Expected, thingy.Description)
        Assert.AreEqual(Expected, button.Description)
    End Sub

    <TestMethod("Can Bind Thing.SetDescription(IButton.Description, a_bit_extra)")>
    Public Sub BindingTestIV()
        Assert.AreNotEqual(button.Description, thingy.Description)

        button.Bind(thingy, Sub(btn, r) r.SetDescription(btn.Description, "a_bit_extra"))

        Assert.AreEqual(button.Description & "a_bit_extra", thingy.Description)

        button.Description = Expected

        Assert.AreEqual(Expected & "a_bit_extra", thingy.Description)
        Assert.AreEqual(Expected, button.Description)
    End Sub

    <TestMethod("Garbage Collection Works")>
    Public Sub BindingTestV()
        Assert.AreNotEqual(button.Description, thingy.Description)

        button.Bind(thingy, Sub(btn, r) r.Description = btn.Description)

        thingy = Nothing

        GC.Collect()
        GC.WaitForPendingFinalizers()
        GC.Collect()
        GC.WaitForPendingFinalizers()

        Dim target2 As Thing = New Thing() With {
            .Description = "Another Description"
        }

        button.Bind(target2, Sub(btn, r) r.Description = btn.Description)

        button.Description = Expected

        Assert.AreEqual(Expected, button.Description)
    End Sub

End Class
