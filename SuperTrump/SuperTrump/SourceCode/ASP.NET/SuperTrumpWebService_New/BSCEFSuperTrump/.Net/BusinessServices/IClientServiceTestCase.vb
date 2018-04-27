Imports NUnit.Framework

<TestFixture()> _
Public Class IClientServiceTestCase

    <Test()> _
            Public Sub TestProcessMQMessage()
        Dim obj As New IClientService
        Dim objInput, objOutPut As New Xml.XmlDocument
        Try
            objInput.Load("XML_IN\ProcessMQMessagetest.xml")
            objOutPut.Load("XML_OUT\ProcessMQMessagetest.xml")
            Assert.Greater(objOutPut.InnerText, obj.ProcessMQMessage(objInput.InnerXml))
        Catch ex As Exception
            Throw ex
        Finally
            obj.Dispose()
            objInput = Nothing
            objOutPut = Nothing
        End Try
    End Sub

    <Test()> _
           Public Sub ProcessPricingRequestTEST()
        Dim obj As New IClientService
        Dim objInput, objOutPut As New Xml.XmlDocument
        Try
            objInput.Load("XML_IN\ProcessPricingRequestTEST.xml")
            objOutPut.Load("XML_OUT\ProcessPricingRequestTEST.xml")
            Assert.Greater(objOutPut.InnerText, obj.ProcessPricingRequest(objInput.InnerXml))
        Catch ex As Exception
            Throw ex
        Finally
            obj.Dispose()
            objInput = Nothing
            objOutPut = Nothing
        End Try
    End Sub

End Class
