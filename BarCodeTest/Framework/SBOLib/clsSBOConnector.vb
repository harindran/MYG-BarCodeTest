Imports SAPbouiCOM.Framework
Namespace Mukesh.SBOLib
    Public Class SBOConnector
        Public Function GetApplication(ByVal ConnectionStr As String) As SAPbouiCOM.Application
            Dim objGUIAPI As SAPbouiCOM.SboGuiApi
            Dim objApp As SAPbouiCOM.Application
            Try
                objGUIAPI = New SAPbouiCOM.SboGuiApi
                objGUIAPI.Connect(ConnectionStr)
                objApp = objGUIAPI.GetApplication(-1)

                If Not objApp Is Nothing Then Return objApp
            Catch ex As Exception
                MsgBox(ex.Message)
                End
            End Try
            Return Nothing
        End Function
        Public Function GetCompany(ByVal SBOApplication As SAPbouiCOM.Application) As SAPbobsCOM.Company
            Dim objCompany As New SAPbobsCOM.Company
            Dim strCookie As String
            Dim strConContext As String
            SBOApplication.SetStatusBarMessage("Connecting to Company... Please wait", SAPbouiCOM.BoMessageTime.bmt_Short, False)
            Try
                strCookie = objCompany.GetContextCookie()
                strConContext = SBOApplication.Company.GetConnectionContext(strCookie)
                objCompany.SetSboLoginContext(strConContext)
                objCompany.Connect()
                If objCompany.Connected = True Then
                    objAddOn.objApplication.SetStatusBarMessage("Company has been connected Sucessfully", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                Else
                    objAddOn.objApplication.SetStatusBarMessage(objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                End If
                Return objCompany
            Catch ex As Exception
                MsgBox(ex.Message & vbLf & ex.StackTrace)
            End Try
            Return Nothing
        End Function
        Private Sub loadCompanylist(ByVal FormUID As String)
            Dim objRS As SAPbobsCOM.Recordset
            objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objRS = objAddOn.objCompany.GetCompanyList()
            Dim str As String
            While Not objRS.EoF
                str = objRS.Fields.Item(0).Value
                objRS.MoveNext()
            End While
        End Sub
    End Class




End Namespace