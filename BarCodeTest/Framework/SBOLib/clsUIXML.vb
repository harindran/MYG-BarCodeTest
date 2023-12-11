Namespace Mukesh.SBOLib
    Public Class UIXML
        Private Shared intTotalFormCount As Integer = 0
        Private objApplication As SAPbouiCOM.Application

        Public Sub New(ByVal SBOApplication As SAPbouiCOM.Application)
            objApplication = SBOApplication
        End Sub

        Public Function LoadScreenXML(ByVal FileName As String, ByVal Type As enuResourceType, ByVal FormType As String) As SAPbouiCOM.Form
            intTotalFormCount += 1
            Return LoadScreenXML(FileName, Type, FormType, FormType & intTotalFormCount)
        End Function

        Public Function LoadScreenXML(ByVal FileName As String, ByVal Type As enuResourceType, ByVal FormType As String, ByVal FormUID As String) As SAPbouiCOM.Form
            Try
                Dim objForm As SAPbouiCOM.Form
                Dim objXML As New Xml.XmlDocument
                Dim strResource As String
                Dim objFrmCreationPrams As SAPbouiCOM.FormCreationParams

                If Type = enuResourceType.Content Then
                    objXML.Load(FileName)
                    objFrmCreationPrams = objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
                    objFrmCreationPrams.FormType = FormType
                    objFrmCreationPrams.UniqueID = FormUID
                    objFrmCreationPrams.XmlData = objXML.InnerXml
                    objForm = objApplication.Forms.AddEx(objFrmCreationPrams)
                Else
                    strResource = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name & "." & FileName
                    objXML.Load(System.Reflection.Assembly.GetExecutingAssembly.GetManifestResourceStream(strResource))
                    objFrmCreationPrams = objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
                    objFrmCreationPrams.FormType = FormType
                    objFrmCreationPrams.UniqueID = FormUID
                    objFrmCreationPrams.XmlData = objXML.InnerXml
                    objForm = objApplication.Forms.AddEx(objFrmCreationPrams)
                End If
                Return objForm
            Catch ex As Exception

            End Try
        End Function

        Public Sub LoadMenuXML(ByVal FileName As String, ByVal Type As enuResourceType)
            Dim objXML As New Xml.XmlDocument
            Dim strResource As String

            If Type = enuResourceType.Content Then
                objXML.Load(FileName)
                objApplication.LoadBatchActions(objXML.InnerXml)
            Else
                strResource = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name & "." & FileName
                objXML.Load(System.Reflection.Assembly.GetExecutingAssembly.GetManifestResourceStream(strResource))
                objApplication.LoadBatchActions(objXML.InnerXml)
            End If
        End Sub

        Public Enum enuResourceType
            Embeded
            Content
        End Enum
    End Class
End Namespace