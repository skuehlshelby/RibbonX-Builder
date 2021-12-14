Imports System.Drawing
Imports System.Reflection
Imports RibbonFactory.Builders
Imports RibbonFactory.Containers
Imports RibbonFactory.Controls

Public Class ButtonsGroup
    Private ReadOnly _helloWorld As Button
    Private ReadOnly _calculatePi As Button

    Public Sub New(ribbon As Ribbon)
        _calculatePi = New ButtonBuilder().
                WithLabel("Calculate Pi", copyToScreenTip:= True).
                WithSuperTip("So that you too can know the value of Pi.").
                WithImage(New Bitmap(Assembly.GetExecutingAssembly().GetManifestResourceStream("ExampleRibbon.pi.png")), AddressOf ribbon.GetImage).
                ThatDoes(AddressOf ribbon.OnAction, Sub() MessageBox($"The value of Pi is {Math.PI:#.#####}...")).
                Build()

        _helloWorld = New ButtonBuilder().
                WithLabel("Hello World", copyToScreenTip:= True).
                WithSuperTip("The classic introductory exercise. Click me!").
                WithImage(New Bitmap(Assembly.GetExecutingAssembly().GetManifestResourceStream("ExampleRibbon.hello.png")), AddressOf ribbon.GetImage).
                ThatDoes(AddressOf ribbon.OnAction, Sub() MessageBox("Hello World!")).
                Build()
    End Sub

    Public Function AsGroup() As Group
        Return New GroupBuilder().
            WithControls(_helloWorld, _calculatePi).
            WithLabel("Example Group").
            Build()
    End Function

    Private Shared Sub MessageBox(prompt As String)
        MsgBox(prompt, MsgBoxStyle.OkOnly, My.Application.Info.Title)
    End Sub
End Class