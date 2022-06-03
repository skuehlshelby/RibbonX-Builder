Imports RibbonX.Controls
Imports RibbonX.Utilities.Testing

<TestClass()> Public Class ValidationWithXsdTests

    <TestMethod()>
    Public Sub RibbonXWithShowLabelAttributeOnGroupElementIsInvalid()
        Dim invalidRibbonX As String = <?xml version="1.0" encoding="utf-8"?>

                                       <customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
                                           <ribbon>
                                               <tabs>
                                                   <tab>
                                                       <group label="group1" showLabel="true">
                                                           <button id="button1" description="test" label="test" showLabel="true"/>
                                                       </group>
                                                   </tab>
                                               </tabs>
                                           </ribbon>
                                       </customUI>.ToString()

        Dim errorLog As RibbonErrorLog = Ribbon.GetErrors(invalidRibbonX)

        Assert.AreEqual(1, errorLog.Errors.Count()) 'showLabel is not defined for the element 'group'
    End Sub

End Class