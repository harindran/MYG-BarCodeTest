Imports SAPbouiCOM.Framework
Namespace Mukesh.SBOLib
    Public Class GeneralFunctions
        Private objCompany As SAPbobsCOM.Company
        Private strThousSep As String = ","
        Private strDecSep As String = "."
        Private intQtyDec As Integer = 3
        'Public con As New Sap.Data.Hana.HanaConnection

        Public Sub New(ByVal Company As SAPbobsCOM.Company)
            Dim objRS As SAPbobsCOM.Recordset
            objCompany = Company

            objRS = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objRS.DoQuery("SELECT * FROM OADM")
            If Not objRS.EoF Then
                strThousSep = objRS.Fields.Item("ThousSep").Value
                strDecSep = objRS.Fields.Item("DecSep").Value
                intQtyDec = objRS.Fields.Item("QtyDec").Value
            End If
        End Sub

        Public Function GetDateTimeValue(ByVal SBODaMIPLAGNTMASring As String) As DateTime
            Dim objBridge As SAPbobsCOM.SBObob
            objBridge = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
            objBridge.Format_StringToDate("")
            Return objBridge.Format_StringToDate(SBODaMIPLAGNTMASring).Fields.Item(0).Value
        End Function
        Public Function GetSBODateString(ByVal DateVal As DateTime) As String

            Dim objBridge As SAPbobsCOM.SBObob
            objBridge = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
            Return objBridge.Format_DateToString(DateVal).Fields.Item(0).Value
        End Function

        Public Function GetSBODaMIPLAGNTMASring(ByVal DateVal As DateTime) As String
            Dim objBridge As SAPbobsCOM.SBObob
            objBridge = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
            Return objBridge.Format_DateToString(DateVal).Fields.Item(0).Value
        End Function
        Public Function GetQtyValue(ByVal QtyString As String) As Double
            Dim dblValue As Double
            QtyString = QtyString.Replace(strThousSep, "")
            QtyString = QtyString.Replace(strDecSep, System.Globalization.NumberFormatInfo.CurrentInfo.NumberDecimalSeparator)
            dblValue = Convert.ToDouble(QtyString)
            Return dblValue
        End Function

        Public Function GetQtyString(ByVal QtyVal As Double) As String
            GetQtyString = QtyVal.ToString()
            GetQtyString.Replace(",", strDecSep)
        End Function

        Public Function GetCode(ByVal sTableName As String) As String
            Dim objRS As SAPbobsCOM.Recordset
            objRS = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If objAddOn.HANA Then
                sTableName = Replace(sTableName, "[", "")
                sTableName = Replace(sTableName, "]", "")
                objRS.DoQuery("SELECT TOP 1 ""Code"" FROM " & Chr(34) & sTableName & Chr(34) & " ORDER BY CAST(""Code"" AS integer) DESC;")
            Else
                objRS.DoQuery("SELECT Top 1 Code FROM " & sTableName & " ORDER BY Convert(INT,Code) DESC")
            End If

            If Not objRS.EoF Then
                Return Convert.ToInt32(objRS.Fields.Item(0).Value.ToString()) + 1
            Else
                GetCode = "1"
            End If
        End Function

        Public Function GetDocNum(ByVal sUDOName As String, ByVal Series As Integer) As String
            Dim StrSQL As String
            Dim objRS As SAPbobsCOM.Recordset
            objRS = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If objAddOn.HANA Then
                '  StrSQL = "select ""AutoKey"" from ONNM where ""ObjectCode""='" & sUDOName & "'"
                If Series = 0 Then
                    StrSQL = " select  ""NextNumber""  from NNM1 where ""ObjectCode""='" & sUDOName & "'"
                Else
                    StrSQL = " select  ""NextNumber""  from NNM1 where ""ObjectCode""='" & sUDOName & "' and ""Series"" = " & Series
                End If

            Else
                StrSQL = "select Autokey from onnm where objectcode='" & sUDOName & "'"
            End If
            objRS.DoQuery(StrSQL)
            objRS.MoveFirst()
            If Not objRS.EoF Then
                Return Convert.ToInt32(objRS.Fields.Item(0).Value.ToString())
            Else
                GetDocNum = "1"
            End If
        End Function

        'Public Function Connection(ByVal user As String) As Boolean
        '    Try

        '        '   con = New Sap.Data.Hana.HanaConnection("Server=192.168.168.249:30015;Catalog=PREETHI_ADDON;UserID=SYSTEM;Password=Miplive2017")

        '        'con = New SqlConnection("DATA SOURCE = " + objGVar.MServerName + ";INITIAL CATALOG = " + objGVar.MDBName + "; USER ID=" + objGVar.MUID + "; PASSWORD=" + objGVar.MPWD + "; Connect Timeout=120")
        '        con = New Sap.Data.Hana.HanaConnection("Server = " + objGVar.MServerName + ":30015;Catalog = " + objGVar.MDBName + "; UserID=" + objGVar.MUID + "; Password=" + objGVar.MPWD + "")

        '        'Maincon.Open()

        '        con.Open()
        '        '  MsgBox("HanaClient Connected")

        '        Return True
        '    Catch ex As Exception

        '        '  MsgBox("Error : " + ex.Message + vbCrLf + "Position : " + ex.StackTrace)
        '        Return Replace(ex.ToString, "'", "")
        '    End Try
        'End Function
        'Public Function GetDocNum_Mbook(ByVal sUDOName As String) As String
        '    objRS = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        '    StrSQL = "select Autokey from onnm where objectcode='" & sUDOName & "'"
        '    objRS.DoQuery(StrSQL)
        '    objRS.MoveFirst()
        '    objAddOn.objApplication.MessageBox(objRS.RecordCount)
        '    If objRS.RecordCount > 0 Then
        '        Return objRS.Fields.Item(0).Value.ToString
        '    Else
        '        Return "1"
        '    End If
        'End Function

        Public Function getSingleValue(ByVal StrSQL As String) As String
            Try
                Dim rset As SAPbobsCOM.Recordset = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim strReturnVal As String = ""
                rset.DoQuery(StrSQL)
                Return IIf(rset.RecordCount > 0, rset.Fields.Item(0).Value.ToString(), "")
            Catch ex As Exception
                objAddOn.objApplication.StatusBar.SetText(" Get Single Value Function Failed :  " & ex.Message + StrSQL, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Return ""
            End Try
        End Function
    End Class
End Namespace