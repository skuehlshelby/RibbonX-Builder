Imports RibbonFactory.Enums.MSO

Namespace BuilderInterfaces

    Public Interface IID(Of TBuilder)
        ''' <summary>
        ''' Add a specific ID to a control. This is not normally necessary, as IDs are automatically generated and assigned.
        ''' </summary>
        ''' <param name="id">The custom ID to add.</param>
        ''' <returns>The original builder instance, so object construction may continue.</returns>
        Function WithId(id As String) As TBuilder

        ''' <summary>
        ''' Add a namespace to a control. Ribbons whose controls share namespaces are merged by the Office host application.
        ''' </summary>
        ''' <param name="[namespace]">The namespace that this control should occupy.</param>
        ''' <param name="id">The custom ID to add. A unique ID may be obtained using RibbonAttributes.IdManager.GetID(), if you do not wish to think of one yourself.</param>
        ''' <returns>The original builder instance, so object construction may continue.</returns>
        Function WithIdQ([namespace] As String, id As String) As TBuilder

        ''' <summary>
        ''' Specify that the control under construction should be a copy of a control supplied by the host application.
        ''' </summary>
        ''' <param name="mso">The ID of the Office control that should be copied.</param>
        ''' <returns>The original builder instance. Object construction should be terminated after this call.</returns>
        Function WithIdMso(mso As MSO) As TBuilder

    End Interface

End NameSpace