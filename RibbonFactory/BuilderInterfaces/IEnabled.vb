Namespace BuilderInterfaces

    Public Interface IEnabled(Of T)

        Function Enabled() As T
        Function Enabled(callback As FromControl(Of Boolean)) As T
        Function Disabled() As T
        Function Disabled(callback As FromControl(Of Boolean)) As T

    End Interface

End NameSpace