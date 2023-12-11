Imports System.IO
Imports System.Drawing.Printing
Imports System.Drawing
Imports SAPbouiCOM.Framework
Public Class ClsbarcodeReprint
    Public Const formtype = "BarRprt"
    Dim objForm As SAPbouiCOM.Form
    Dim objRS As SAPbobsCOM.Recordset
    Dim strSQL As String
    Dim checkbox As SAPbouiCOM.CheckBox
    Dim objMatrix As SAPbouiCOM.Matrix
    Public Sub LoadScreen()
        objForm = objAddOn.objUIXml.LoadScreenXML("BarcodeReprint.xml", Mukesh.SBOLib.UIXML.enuResourceType.Embeded, formtype)
        objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
        AddPrinters(objForm.UniqueID)
        addConditions(objForm.UniqueID)
        objMatrix = objForm.Items.Item("14").Specific
        'objMatrix.AddRow(1)

    End Sub
    Public Sub ItemEvent(ByVal FormUID As String, ByRef Pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        Try
            If Pval.BeforeAction = True Then
                Select Case Pval.EventType
                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        If Pval.ItemUID = "3" Then
                            If validate(FormUID) = False Then
                                BubbleEvent = False : Exit Sub
                            End If
                        End If
                End Select

            ElseIf Pval.BeforeAction = False Then
                Select Case Pval.EventType
                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                        If Pval.ItemUID = "7" Or Pval.ItemUID = "9" Or Pval.ItemUID = "11" Or Pval.ItemUID = "13" Then
                            BarcodeCFL(FormUID, Pval)
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If Pval.ItemUID = "4" Then
                            loadmatrix(FormUID)
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        objForm = objAddOn.objApplication.Forms.Item(FormUID)
                        If Pval.ItemUID = "3" Then
                            If validate(FormUID) Then
                                ' CreateFile(FormUID)
                                'CreateFileNew(FormUID, "Rs. ")
                                SelectPrinterType(FormUID)
                            End If
                        ElseIf Pval.ItemUID = "5" Then
                            objMatrix = objForm.Items.Item("14").Specific
                            objMatrix.Clear()
                        ElseIf Pval.ItemUID = "101" Then
                            Try
                                objForm.Close()
                            Catch ex As Exception
                            End Try

                        End If

                    Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                        objForm = objAddOn.objApplication.Forms.Item(FormUID)
                        If Pval.ItemUID = "14" And Pval.ColUID = "1" Then
                            objMatrix = objForm.Items.Item("14").Specific
                            objForm.Freeze(True)
                            For i As Integer = 1 To objMatrix.RowCount
                                objMatrix.Columns.Item("1").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            Next
                            objForm.Freeze(False)
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                        If Pval.ItemUID = "7B" Then
                            If objForm.Items.Item("7B").Specific.String <> "" Then
                                LoadGRPOQuantity(FormUID)
                            End If

                        End If
                End Select
            End If
        Catch ex As Exception
            objAddOn.objApplication.MessageBox(ex.Message)
        End Try
    End Sub
    Private Sub LoadGRPOQuantity(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        Dim objMatrix As SAPbouiCOM.Matrix
        objMatrix = objForm.Items.Item("14").Specific
        For i As Integer = 1 To objMatrix.RowCount
            objMatrix.Columns.Item("4A").Cells.Item(i).Specific.String = objForm.Items.Item("7B").Specific.String
        Next
    End Sub

    Private Sub SelectPrinterType(ByVal FormUID As String)
        Dim objCombo As SAPbouiCOM.ComboBox
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objCombo = objForm.Items.Item("7D").Specific
        'Dim FC As String = "F"
        'objCombo = objForm.Items.Item("16").Specific
        Select Case objCombo.Selected.Value
            Case "BR"
                'CreateFile(FormUID, "Rs. ")
                CreateFileNew(FormUID, "Rs. ")
            Case "SM"
                'CreateFile2(FormUID, "Rs. ")
                'CreateFileNew(FormUID, "Rs. ")
                CreateFileUpdatedPRN(FormUID, "Rs. ")
        End Select

    End Sub
    Public Sub menuevent(ByVal pVal As SAPbouiCOM.MenuEvent, ByVal BubbleEvent As Boolean)
        If pVal.MenuUID = "1281" Then 'find
        ElseIf pVal.MenuUID = "1282" Then 'add
        End If
    End Sub
    Private Sub AddPrinters(ByVal FormUID As String)
        Try
            Dim objCombo As SAPbouiCOM.ComboBox
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            objCombo = objForm.Items.Item("7D").Specific
            objCombo.ValidValues.Add("BR", "MyG")
            objCombo.ValidValues.Add("SM", "3G")
            objCombo.Select("SM", SAPbouiCOM.BoSearchKey.psk_ByDescription)
        Catch ex As Exception

        End Try
    End Sub
    Public Sub loadmatrix(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("14").Specific
        objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Dim SerialItem As String = ""
        SerialItem = objAddOn.objGenFunc.getSingleValue("select ""ManSerNum"" from OITM where ""ItemCode""='" & objForm.Items.Item("7").Specific.value & "'")


        'strSQL = " select '' as No,'Y' as sel,Intrserial,T1.Price,T0.ItemCode,T2.sww,T3.U_shtname as FirmName,T0.ItemName,T4.whsname, BaseNum,T5.U_o7_color as Color,T5.U_o8_siz as Size  from OSRI T0" & _
        '     " join PDN1 T5 on T5.docentry=T0.baseentry and T5.linenum=T0.Baselinnum " & _
        '      " join ITM1 T1 on T1.ItemCode = T0.itemcode join OITM T2 on T2.ItemCode =T0.ItemCode " & _
        '      " join OMRC T3 on T3.FirmCode =T2.FirmCode join OWHS T4 on T4.WhsCode =T0.WhsCode " & _
        '      " where T1.PriceList =2 and T0.U_printed='Y' and T0.status=0 "

        If SerialItem = "Y" Then
            If objAddOn.HANA Then
               
                'strSQL = "select '' as No,'Y' as sel,T4.""IntrSerial"",(select top 1 ""Price"" from ITM1 where ""ItemCode""='AZRD09784' and ""PriceList""=5) AS ""price"","
                'strSQL += vbCrLf + " T1.""ItemCode"",CAST(T1.""Quantity"" AS varchar(10)) AS ""qty"",T5.""SWW"",'' ""mfr"" ,I1.""ItemName"",(select ""WhsName"" from OWHS where ""WhsCode""='WHO2') as ""WhsName"","
                'strSQL += vbCrLf + " I1.""BaseNum"",To_varchar(T0.""DocDate"",'Mon-yyyy') AS year1 from OPDN T0 left join PDN1 T1 on T0.""DocEntry""=T1.""DocEntry"""
                'strSQL += vbCrLf + " left outer join SRI1 I1 on T1.""ItemCode""=I1.""ItemCode"" and (T1.""DocEntry""=I1.""BaseEntry"" and T1.""ObjType""=I1.""BaseType"") and T1.""LineNum""=I1.""BaseLinNum"""
                'strSQL += vbCrLf + " left outer join OSRI T4 on T4.""ItemCode""=I1.""ItemCode"" and I1.""SysSerial""=T4.""SysSerial"" and I1.""WhsCode"" = T4.""WhsCode"" left outer join OITM T5 on T5.""ItemCode"" =T4.""ItemCode""  "
                'strSQL += vbCrLf + " where  T4.""U_printed""='Y' and T4.""Status""=0 "

                strSQL = "select '' as No,'Y' as sel,T4.""IntrSerial"",(select top 1 ""Price"" from ITM1 where ""ItemCode""='048107108267' and ""PriceList""=5) AS ""price"","
                strSQL += vbCrLf + "  I1.""ItemCode"",CAST(T4.""Quantity"" AS varchar(10)) AS ""qty"",T5.""SWW"",'' ""mfr"" ,I1.""ItemName"",(select ""WhsName"" from OWHS where ""WhsCode""='CRHQ') as ""WhsName"","
                strSQL += vbCrLf + "  I1.""BaseNum"",To_varchar(I1.""DocDate"",'Mon-yyyy') AS year1"
                strSQL += vbCrLf + "  from SRI1 I1 left outer join OSRI T4 on T4.""ItemCode""=I1.""ItemCode"" and I1.""SysSerial""=T4.""SysSerial"" and I1.""WhsCode"" = T4.""WhsCode"" left outer join OITM T5 on T5.""ItemCode"" =T4.""ItemCode""  "
                strSQL += vbCrLf + " where  (T4.""U_printed""='Y' or T4.""U_printed""<>'Y') and T4.""Status""=0 "

            Else
                strSQL = " select '' as No,'Y' as sel,Intrserial,isnull(T0.U_MRP,T5.U_MRP ) as price,T0.ItemCode,T2.sww,T3.U_shtname as FirmName,T0.ItemName,T4.whsname, BaseNum  from OSRI T0" & _
            " join PDN1 T5 on T5.DocEntry=T0.BaseEntry and T5.LineNum=T0.BaseLinNum " & _
             "  join OITM T2 on T2.ItemCode =T0.ItemCode " & _
             " join OMRC T3 on T3.FirmCode =T2.FirmCode join OWHS T4 on T4.WhsCode =T0.WhsCode " & _
             " where  T0.U_printed='Y' and T0.Status=0 "
            End If

            '   strSQL = " select '' as No,'Y' as sel,Intrserial,isnull(T0.U_MRP,T5.U_MRP ) as price,T0.ItemCode,T2.sww,T3.U_shtname as FirmName,T0.ItemName,T4.whsname, BaseNum,T5.U_o7_color as Color,T5.U_o8_siz as Size  from OSRI T0" & _
            '" join PDN1 T5 on T5.DocEntry=T0.BaseEntry and T5.LineNum=T0.BaseLinNum " & _
            ' "  join OITM T2 on T2.ItemCode =T0.ItemCode " & _
            ' " join OMRC T3 on T3.FirmCode =T2.FirmCode join OWHS T4 on T4.WhsCode =T0.WhsCode " & _
            ' " where  T0.U_printed='Y' and T0.Status=0 "

            If objForm.Items.Item("7").Specific.string <> "" Then
                If objAddOn.HANA Then
                    strSQL = strSQL + " and I1.""ItemCode"" ='" & objForm.Items.Item("7").Specific.value & "'"
                Else
                    strSQL = strSQL + " and T0.ItemCode ='" & objForm.Items.Item("7").Specific.value & "'"
                End If

            End If
            If objForm.Items.Item("11").Specific.string <> "" Then
                If objAddOn.HANA Then
                    strSQL = strSQL + " and I1.""WhsCode"" ='" & objForm.Items.Item("11").Specific.value & "'"
                Else
                    strSQL = strSQL + " and T4.WhsName ='" & objForm.Items.Item("11").Specific.value & "'"
                End If

            End If
            'If objForm.Items.Item("9").Specific.value <> "" Then
            '    If objAddOn.HANA Then
            '        strSQL = strSQL + " and T0.""IntrSerial"" >='" & objForm.Items.Item("9").Specific.value & "'"
            '    Else
            '        strSQL = strSQL + " and T0.IntrSerial >='" & objForm.Items.Item("9").Specific.value & "'"
            '    End If

            'End If
            'If objForm.Items.Item("13").Specific.value <> "" Then
            '    If objAddOn.HANA Then
            '        strSQL = strSQL + " and T0.""IntrSerial"" <='" & objForm.Items.Item("13").Specific.value & "'"
            '    Else
            '        strSQL = strSQL + " and T0.IntrSerial <='" & objForm.Items.Item("13").Specific.value & "'"
            '    End If

            'End If
        Else
            If objAddOn.HANA Then
                'strSQL = "select '' as No,'Y' as sel,I1.""BatchNum"",(select top 1 ""Price"" from ITM1 where ""ItemCode""='" & objForm.Items.Item("7").Specific.value & "' and ""PriceList""=5) AS ""price"","
                'strSQL += vbCrLf + " T1.""ItemCode"",CAST(T1.""Quantity"" AS varchar(10)) AS ""qty"",T5.""SWW"",'' ""mfr"" ,I1.""ItemName"",(select ""WhsName"" from OWHS where ""WhsCode""='" & objForm.Items.Item("11").Specific.value & "') as ""WhsName"","
                'strSQL += vbCrLf + " I1.""BaseNum"",To_varchar(T0.""DocDate"",'Mon-yyyy') AS year1 from OPDN T0 left join PDN1 T1 on T0.""DocEntry""=T1.""DocEntry"""
                'strSQL += vbCrLf + " left outer join IBT1 I1 on T1.""ItemCode""=I1.""ItemCode"" and (T1.""DocEntry""=I1.""BaseEntry"" and T1.""ObjType""=I1.""BaseType"") and T1.""LineNum""=I1.""BaseLinNum"""
                'strSQL += vbCrLf + " left outer join OIBT T4 on T4.""ItemCode""=I1.""ItemCode"" and I1.""BatchNum""=T4.""BatchNum"" and I1.""WhsCode"" = T4.""WhsCode"" left outer join OITM T5 on T5.""ItemCode"" =T4.""ItemCode""  "
                'strSQL += vbCrLf + " where  T4.""U_printed""='Y' and T4.""Status""=0  "

                strSQL = "select '' as No,'Y' as sel,I1.""BatchNum"",(select top 1 ""Price"" from ITM1 where ""ItemCode""='048107108267' and ""PriceList""=5) AS ""price"","
                strSQL += vbCrLf + "  I1.""ItemCode"",CAST(I1.""Quantity"" AS varchar(10)) AS ""qty"",T5.""SWW"",'' ""mfr"" ,I1.""ItemName"",(select ""WhsName"" from OWHS where ""WhsCode""='CRHQ') as ""WhsName"","
                strSQL += vbCrLf + "  I1.""BaseNum"",To_varchar(I1.""DocDate"",'Mon-yyyy') AS year1"
                strSQL += vbCrLf + "  from IBT1 I1 left outer join OIBT T4 on T4.""ItemCode""=I1.""ItemCode"" and I1.""BatchNum""=T4.""BatchNum"" and I1.""WhsCode"" = T4.""WhsCode"" left outer join OITM T5 on T5.""ItemCode"" =T4.""ItemCode""  "
                strSQL += vbCrLf + " where  (T4.""U_printed""='Y' or T4.""U_printed""<>'Y') and T4.""Status""=0"
            Else
                strSQL = " select '' as No,'Y' as sel,Intrserial,isnull(T0.U_MRP,T5.U_MRP ) as price,T0.ItemCode,T2.sww,T3.U_shtname as FirmName,T0.ItemName,T4.whsname, BaseNum  from OSRI T0" & _
            " join PDN1 T5 on T5.DocEntry=T0.BaseEntry and T5.LineNum=T0.BaseLinNum " & _
             "  join OITM T2 on T2.ItemCode =T0.ItemCode " & _
             " join OMRC T3 on T3.FirmCode =T2.FirmCode join OWHS T4 on T4.WhsCode =T0.WhsCode " & _
             " where  T0.U_printed='Y' and T0.Status=0 "
            End If
            If objForm.Items.Item("7").Specific.string <> "" Then
                If objAddOn.HANA Then
                    strSQL = strSQL + " and I1.""ItemCode"" ='" & objForm.Items.Item("7").Specific.value & "'"
                Else
                    strSQL = strSQL + " and T0.ItemCode ='" & objForm.Items.Item("7").Specific.value & "'"
                End If

            End If
            If objForm.Items.Item("11").Specific.string <> "" Then
                If objAddOn.HANA Then
                    strSQL = strSQL + " and I1.""WhsCode"" ='" & objForm.Items.Item("11").Specific.value & "'"
                Else
                    strSQL = strSQL + " and T4.WhsName ='" & objForm.Items.Item("11").Specific.value & "'"
                End If

            End If
        End If

        objMatrix.Clear()
        objForm.DataSources.DataTables.Item("DT").ExecuteQuery(strSQL)
        objMatrix.LoadFromDataSource()
        objAddOn.objApplication.Menus.Item("1300").Activate()
        If objMatrix.RowCount = 0 Then
            objAddOn.objApplication.SetStatusBarMessage("No barcoded items exist in the selected GRPO")
        End If
    End Sub
    Private Sub addConditions(ByVal FormUID As String)
        Dim objCFL As SAPbouiCOM.ChooseFromList
        Dim objconditions As SAPbouiCOM.Conditions
        Dim objcond As SAPbouiCOM.Condition
        objForm = objAddOn.objApplication.Forms.Item(FormUID)

        objCFL = objForm.ChooseFromLists.Item("CFL_2")
        objconditions = objCFL.GetConditions
        objcond = objconditions.Add
        objcond.Alias = "ManSerNum"
        objcond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
        objcond.CondVal = "Y"
        objcond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
        objcond = objconditions.Add
        objcond.Alias = "ManBtchNum"
        objcond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
        objcond.CondVal = "Y"
        objCFL.SetConditions(objconditions)

        objCFL = objForm.ChooseFromLists.Item("CFL_4")
        objconditions = objCFL.GetConditions
        objcond = objconditions.Add
        objcond.Alias = "BaseType"
        objcond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
        objcond.CondVal = "20"
        objCFL.SetConditions(objconditions)

        objCFL = objForm.ChooseFromLists.Item("CFL_5")
        objconditions = objCFL.GetConditions
        objcond = objconditions.Add
        objcond.Alias = "BaseType"
        objcond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
        objcond.CondVal = "20"
        objCFL.SetConditions(objconditions)
    End Sub
    Private Sub CreateFile(ByVal FormUID As String)
        Dim intloop As Integer
        Dim model As String
        Dim desc As String
        Dim price As String
        Dim barcode As String
        Dim color As String
        Dim size As String
        Dim colorsize As String = ""
        Dim count As Integer = 0
        Dim path As String = ""
        Dim fs As FileStream
        Dim rName As String = ""
        rName = SystemInformation.UserName
        If Directory.Exists(System.Windows.Forms.Application.StartupPath + "\" + rName) Then
        Else
            Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + "\" + rName)
        End If
        path = System.Windows.Forms.Application.StartupPath + "\" + rName + "\mylabel.prn"
        ' Delete the file if it exists.
        If File.Exists(path) Then
            File.Delete(path)
        End If
        'Create the file.
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("14").Specific
        While objMatrix.RowCount > 0
            intloop = 1
            If objMatrix.Columns.Item("1").Cells.Item(intloop).Specific.checked = True Then
                barcode = objMatrix.Columns.Item("2").Cells.Item(intloop).Specific.string
                price = objMatrix.Columns.Item("3").Cells.Item(intloop).Specific.string
                desc = objMatrix.Columns.Item("6").Cells.Item(intloop).Specific.string
                model = objMatrix.Columns.Item("5").Cells.Item(intloop).Specific.string
                color = objMatrix.Columns.Item("10").Cells.Item(intloop).Specific.string
                size = objMatrix.Columns.Item("11").Cells.Item(intloop).Specific.string
                barcode = Left(barcode, 12)
                price = price.Replace(",", "")
                price = Left(price, 7)
                desc = Left(desc, 3)
                model = Left(model, 10)
                desc = desc + "-" + model
                color = Left(color, 9)
                size = Left(size, 2)
                colorsize = color + "-" + size

                fs = New FileStream(path, FileMode.Create, FileAccess.Write)
                Dim s As New StreamWriter(fs)
                's.WriteLine()
                's.WriteLine()
                's.WriteLine("Q0001,0")
                's.WriteLine("q831")
                's.WriteLine("rN")
                's.WriteLine("S1")
                's.WriteLine("D7")
                's.WriteLine("ZT")
                's.WriteLine("JB")
                's.WriteLine("OD")
                's.WriteLine("R64,0")
                's.WriteLine("N")
                's.WriteLine("A665,34,2,1,1,1,N," & Chr(34) & "MRP:Rs." & price & Chr(34))
                's.WriteLine("A480,34,2,2,1,1,N," & Chr(34) & barcode & Chr(34))
                's.WriteLine("A665,66,2,1,1,1,N," & Chr(34) & colorsize & Chr(34))
                's.WriteLine("A665,98,2,1,1,1,N," & Chr(34) & desc & Chr(34))
                's.WriteLine("B480,113,2,1A,1,3,60,N," & Chr(34) & barcode & Chr(34))
                's.WriteLine("X0,128,1,0,129")
                's.WriteLine("P1")


                s.WriteLine("^XA")
                s.WriteLine("^MMT")
                s.WriteLine("^PW607")
                s.WriteLine("^LL0200")
                s.WriteLine("^LS0")
                s.WriteLine("^BY2,3,68^FT585,97^BCI,,N,N")
                s.WriteLine("^FD>:" & barcode & "^FS")
                s.WriteLine("^FT572,81^A0I,15,28^FH\^FD" & barcode & "^FS")
                s.WriteLine("^FT576,58^A0I,21,21^FH\^FD" & desc & "^FS")
                s.WriteLine("^FT574,30^A0I,19,24^FH\^FD" & price & "^FS")
                s.WriteLine("^BY2,3,68^FT280,97^BCI,,N,N")
                s.WriteLine("^FD>:" & barcode & "^FS")
                s.WriteLine("^FT268,81^A0I,15,28^FH\^FD" & barcode & "^FS")
                s.WriteLine("^FT272,58^A0I,21,21^FH\^FD" & desc & "^FS")
                s.WriteLine("^FT270,30^A0I,19,24^FH\^FD" & price & "^FS")
                s.WriteLine("^PQ1,0,1,Y^XZ")
                s.Close()
                fs.Close()
                Try
                    File.Copy(path, getDefaultPrinter())
                Catch ex As Exception

                    RawPrinterHelper.SendFileToPrinter(getDefaultPrinter(), path)
                End Try
                count += 1
            End If
            updatePrice(FormUID, intloop)
            objMatrix.DeleteRow(intloop)
        End While
        objAddOn.objApplication.MessageBox(CStr(count) + " Barcode(s) Printed")
    End Sub
    Private Sub CreateFileNew(ByVal FormUID As String, ByVal Prefix As String)
        ' big
        Dim intloop As Integer
        Dim qty As String
        Dim desc As String
        Dim price As String
        ' Dim price2 As Decimal
        Dim barcode As String
        Dim barcode1 As String = ""
        Dim vendor As String
        Dim year1 As String
        'Dim year2 As String
        ' Dim year3 As Date
        Dim colorsize As String = ""
        Dim count As Integer = 0
        Dim itemcode As String
        Dim itemcode1 As String = ""
        Dim path As String = ""
        Dim fs As FileStream
        Dim rName As String = ""
        rName = SystemInformation.UserName
        Dim Foldername As String
        'Dim Total As Integer
        'Dim RowCount As Integer
        'Foldername = "D:" + "\" + rName + "\MyG"
        Foldername = "E:" + "\" + rName + "\MyG"
        If Directory.Exists(Foldername) Then
        Else
            Directory.CreateDirectory(Foldername)
        End If


        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("14").Specific
        'While objMatrix.RowCount > 0
        intloop = 1
        '--------------File Creation--------------------
        Dim Filetype As String = "BIG"

        Filetype = Filetype.Replace(" ", "")
        'Dim Filename As String = Trim(objMatrix.Columns.Item("6").Cells.Item(intloop).Specific.string) + "_" + Filetype + "_" + (desc.Replace(" ", "")).Replace("/", "").Replace(".", "").Replace(":", "") + "_" + CStr(CInt(objMatrix.Columns.Item("7").Cells.Item(intloop).Specific.string) + 1) + "_" + CStr(count) + ".prn"
        Dim Filename As String = Trim(objForm.Items.Item("7").Specific.String) + "_" + Filetype + "_" + System.DateTime.Now.ToString("yyMMddHHmmss") + "_" + CStr(count) + ".prn"
        'Create the file.
        ' path = Application.StartupPath + "\" + rName + "\mylabel.prn"
        path = Foldername + "\" + Filename
        ' Delete the file if it exists.
        If File.Exists(path) Then
            File.Delete(path)
        End If
        fs = New FileStream(path, FileMode.Create, FileAccess.Write)

        Dim s As New StreamWriter(fs)

        For i As Integer = 1 To objMatrix.RowCount
            If objMatrix.Columns.Item("1").Cells.Item(i).Specific.checked = True Then
                itemcode = objForm.Items.Item("7").Specific.String 'objMatrix.Columns.Item("8").Cells.Item(i).Specific.string
                barcode = objMatrix.Columns.Item("2").Cells.Item(i).Specific.string
                price = objMatrix.Columns.Item("3").Cells.Item(i).Specific.string
                desc = objMatrix.Columns.Item("7").Cells.Item(i).Specific.string
                qty = objMatrix.Columns.Item("4A").Cells.Item(i).Specific.string
                'vendor = objMatrix.Columns.Item("5A").Cells.Item(i).Specific.string
                year1 = objMatrix.Columns.Item("9A").Cells.Item(i).Specific.string

                'barcode = Left(barcode, 10)
                ' qty = CInt(qty)
                desc = Left(desc, 20)
                'desc = Strings.Left(desc, desc.LastIndexOf(":"))
                price = price.Replace(",", "")
                price = Prefix + Mid(price, 1, (price.IndexOf(".") + 3))
                price = price.Replace("s", "")
                price = price.Replace("R", "")
                price = price.Replace(" ", "")
                For k As Integer = 1 To qty
                    If k Mod 2 = 1 Then
                        'RowCount = 2
                        s.WriteLine("^XA")
                        s.WriteLine("^MMT")
                        s.WriteLine("^PW799")
                        s.WriteLine("^LL0400")
                        s.WriteLine("^LS0")
                        s.WriteLine("^FT788,275^A0I,21,24^FH\^FDPacked by : MYG,CORP.OFFICE^FS")
                        s.WriteLine("^FT788,250^A0I,16,28^FH\^FDCALICUT,KERALA,INDIA^FS")
                        s.WriteLine("^FT788,222^A0I,19,24^FH\^FDpacked date : " & year1 & " QTY : " & qty & " nos^FS")
                        s.WriteLine("^FT788,193^A0I,19,24^FH\^FDCommodity : " & desc & "^FS")
                        s.WriteLine("^FT788,165^A0I,20,26^FH\^FDCustomer Care No. 18001232006^FS")
                        s.WriteLine("^FT789,145^A0I,21,28^FH\^FDEmail : Info@myg.in^FS")
                        s.WriteLine("^BY2,3,49^FT787,89^BCI,,N,N")
                        s.WriteLine("^FD>:" & itemcode & "^FS")
                        s.WriteLine("^FT788,67^A0I,21,28^FH\^FD" & itemcode & "^FS")         'ItemNumber 1
                        s.WriteLine("^FT788,32^A0I,25,33^FH\^FDMRP Rs" & price & "^FS")
                        s.WriteLine("^FT572,32^A0I,14,28^FH\^FDinc all taxes^FS")
                        s.WriteLine("^BY2,3,49^FT787,320^BCI,,N,N")
                        s.WriteLine("^FD>:" & barcode & "^FS")              'Serial number 1    left 
                        s.WriteLine("^FT788,300^A0I,21,28^FH\^FD" & barcode & "^FS")
                        'If i = objMatrix.RowCount Then
                        '    s.WriteLine("^PQ1,0,1,Y^XZ")
                        'End If
                        If k = qty Then
                            s.WriteLine("^PQ1,0,1,Y^XZ")
                        End If
                    Else
                        'RowCount = 1
                        s.WriteLine("^FT388,275^A0I,21,24^FH\^FDPacked by : MYG,CORP.OFFICE^FS")
                        s.WriteLine("^FT388,250^A0I,16,28^FH\^FDCALICUT,KERALA,INDIA^FS")
                        s.WriteLine("^FT388,222^A0I,19,24^FH\^FDpacked date : " & year1 & " QTY : " & qty & " nos^FS")
                        s.WriteLine("^FT388,193^A0I,19,24^FH\^FDCommodity : " & desc & "^FS")
                        s.WriteLine("^FT388,165^A0I,20,26^FH\^FDCustomer Care No. 18001232006^FS")
                        s.WriteLine("^FT389,145^A0I,21,28^FH\^FDEmail : Info@myg.in^FS")
                        s.WriteLine("^BY2,3,49^FT387,89^BCI,,N,N")
                        s.WriteLine("^FD>:" & itemcode & "^FS")            'Right
                        s.WriteLine("^FT388,67^A0I,21,28^FH\^FD" & itemcode & "^FS")        'ItemNumber 2
                        s.WriteLine("^FT388,32^A0I,25,33^FH\^FDMRP Rs" & price & "^FS")
                        s.WriteLine("^FT172,32^A0I,14,28^FH\^FDinc all taxes^FS")
                        s.WriteLine("^BY2,3,49^FT387,320^BCI,,N,N")
                        s.WriteLine("^FD>:" & barcode & "^FS")             'Serial number 2   Right
                        s.WriteLine("^FT388,300^A0I,21,28^FH\^FD" & barcode & "^FS")
                        s.WriteLine("^PQ1,0,1,Y^XZ")
                    End If
                Next

               
            End If
        Next i

        s.Close()
        fs.Close()

        Dim ToPath As String = ""
        Try
            Dim systemname As String = ""
            ' Dim Str_FromPath As String = "C:\" & Environment.UserName & "\" & txtBarcode.Text & "_" & txtbatchtag.Text & ".txt"
            Dim str_fromPath As String = path
            systemname = Environment.MachineName
            ToPath = "\\" & systemname & "\\" & getDefaultPrinter()
            objAddOn.objApplication.SetStatusBarMessage(ToPath, SAPbouiCOM.BoMessageTime.bmt_Short, False)
            File.Copy(str_fromPath, ToPath)
        Catch ex As Exception
            objAddOn.objApplication.MessageBox(ToPath & " Ex: " & ex.Message)
        End Try
        'Dim objCombo As SAPbouiCOM.ComboBox
        'objCombo = objForm.Items.Item("17").Specific

        Dim j As Integer = 1

        While j <= objMatrix.RowCount
            'updateStatus(FormUID, j, j + 1)
            objMatrix.DeleteRow(j)
        End While
        'End While
        ' objAddOn.objApplication.MessageBox(CStr(count) + " Barcode(s) printed")
        objAddOn.objApplication.MessageBox("Barcode file(s) generated")
    End Sub

    Private Sub CreateFileUpdatedPRN(ByVal FormUID As String, ByVal Prefix As String)
        ' big
        Dim intloop As Integer
        Dim qty As String
        Dim desc As String
        Dim price As String
        ' Dim price2 As Decimal
        Dim barcode As String
        Dim barcode1 As String = ""
        'Dim vendor As String
        Dim year1 As String
        'Dim year2 As String
        ' Dim year3 As Date
        Dim colorsize As String = ""
        Dim count As Integer = 0
        Dim itemcode As String
        Dim itemcode1 As String = ""
        Dim path As String = ""
        Dim fs As FileStream
        Dim rName As String = ""
        rName = SystemInformation.UserName
        Dim Foldername As String
        'Dim Total As Integer
        'Dim RowCount As Integer
        'Foldername = "D:" + "\" + rName + "\MyG"
        Foldername = "E:" + "\" + rName + "\MyG"
        If Directory.Exists(Foldername) Then
        Else
            Directory.CreateDirectory(Foldername)
        End If


        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("14").Specific
        'While objMatrix.RowCount > 0
        intloop = 1
        '--------------File Creation--------------------
        Dim Filetype As String = "BIG"

        Filetype = Filetype.Replace(" ", "")
        'Dim Filename As String = Trim(objMatrix.Columns.Item("6").Cells.Item(intloop).Specific.string) + "_" + Filetype + "_" + (desc.Replace(" ", "")).Replace("/", "").Replace(".", "").Replace(":", "") + "_" + CStr(CInt(objMatrix.Columns.Item("7").Cells.Item(intloop).Specific.string) + 1) + "_" + CStr(count) + ".prn"
        Dim Filename As String = Trim(objForm.Items.Item("7").Specific.String) + "_" + Filetype + "_" + System.DateTime.Now.ToString("yyMMddHHmmss") + "_" + CStr(count) + ".prn"
        'Create the file.
        ' path = Application.StartupPath + "\" + rName + "\mylabel.prn"
        path = Foldername + "\" + Filename
        ' Delete the file if it exists.
        If File.Exists(path) Then
            File.Delete(path)
        End If
        fs = New FileStream(path, FileMode.Create, FileAccess.Write)

        Dim s As New StreamWriter(fs)


        For i As Integer = 1 To objMatrix.RowCount

            If objMatrix.Columns.Item("1").Cells.Item(i).Specific.checked = True Then
                itemcode = objForm.Items.Item("7").Specific.String 'objMatrix.Columns.Item("8").Cells.Item(i).Specific.string
                barcode = objMatrix.Columns.Item("2").Cells.Item(i).Specific.string
                price = objMatrix.Columns.Item("3").Cells.Item(i).Specific.string
                desc = objMatrix.Columns.Item("7").Cells.Item(i).Specific.string
                qty = objMatrix.Columns.Item("4A").Cells.Item(i).Specific.string
                'vendor = objMatrix.Columns.Item("5A").Cells.Item(i).Specific.string
                year1 = objMatrix.Columns.Item("9A").Cells.Item(i).Specific.string

                'barcode = Left(barcode, 10)
                ' qty = CInt(qty)
                desc = Left(desc, 20)
                'desc = Strings.Left(desc, desc.LastIndexOf(":"))
                price = price.Replace(",", "")
                price = Prefix + Mid(price, 1, (price.IndexOf(".") + 3))
                price = price.Replace("s", "")
                price = price.Replace("R", "")
                price = price.Replace(" ", "")

                For k As Integer = 1 To qty
                    If k Mod 2 = 1 Then
                        'RowCount = 2
                        s.WriteLine("^XA")
                        s.WriteLine("^MMT")
                        s.WriteLine("^PW799")
                        s.WriteLine("^LL0400")
                        s.WriteLine("^LS0")
                        s.WriteLine("^FT788,275^A0I,21,24^FH\^FD ^FS")
                        s.WriteLine("^FT788,250^A0I,16,28^FH\^FD ^FS")
                        s.WriteLine("^FT788,222^A0I,19,24^FH\^FD ^FS")
                        s.WriteLine("^FT788,193^A0I,19,24^FH\^FDCommodity : " & desc & "^FS")
                        s.WriteLine("^FT788,165^A0I,20,26^FH\^FD ^FS")
                        s.WriteLine("^FT789,145^A0I,21,28^FH\^FD ^FS")
                        s.WriteLine("^BY2,3,49^FT787,89^BCI,,N,N")
                        s.WriteLine("^FD>:" & itemcode & "^FS")
                        s.WriteLine("^FT788,67^A0I,21,28^FH\^FD" & itemcode & "^FS")         'ItemNumber 1
                        s.WriteLine("^FT788,32^A0I,25,33^FH\^FD ^FS")
                        s.WriteLine("^FT572,32^A0I,14,28^FH\^FD ^FS")
                        s.WriteLine("^BY2,3,49^FT787,320^BCI,,N,N")
                        s.WriteLine("^FD>:" & barcode & "^FS")              'Serial number 1    left 
                        s.WriteLine("^FT788,300^A0I,21,28^FH\^FD" & barcode & "^FS")
                        'If i = objMatrix.RowCount Then
                        '    s.WriteLine("^PQ1,0,1,Y^XZ")
                        'End If
                        If k = qty Then
                            s.WriteLine("^PQ1,0,1,Y^XZ")
                        End If
                    Else
                        'RowCount = 1
                        s.WriteLine("^FT388,275^A0I,21,24^FH\^FD ^FS")
                        s.WriteLine("^FT388,250^A0I,16,28^FH\^FD ^FS")
                        s.WriteLine("^FT388,222^A0I,19,24^FH\^FD ^FS")
                        s.WriteLine("^FT388,193^A0I,19,24^FH\^FDCommodity : " & desc & "^FS")
                        s.WriteLine("^FT388,165^A0I,20,26^FH\^FD ^FS")
                        s.WriteLine("^FT389,145^A0I,21,28^FH\^FD ^FS")
                        s.WriteLine("^BY2,3,49^FT387,89^BCI,,N,N")
                        s.WriteLine("^FD>:" & itemcode & "^FS")            'Right
                        s.WriteLine("^FT388,67^A0I,21,28^FH\^FD" & itemcode & "^FS")        'ItemNumber 2
                        s.WriteLine("^FT388,32^A0I,25,33^FH\^FD ^FS")
                        s.WriteLine("^FT172,32^A0I,14,28^FH\^FD ^FS")
                        s.WriteLine("^BY2,3,49^FT387,320^BCI,,N,N")
                        s.WriteLine("^FD>:" & barcode & "^FS")             'Serial number 2   Right
                        s.WriteLine("^FT388,300^A0I,21,28^FH\^FD" & barcode & "^FS")
                        s.WriteLine("^PQ1,0,1,Y^XZ")
                    End If
                Next
            End If
        Next i

        s.Close()
        fs.Close()

        Dim ToPath As String = ""
        Try
            Dim systemname As String = ""
            ' Dim Str_FromPath As String = "C:\" & Environment.UserName & "\" & txtBarcode.Text & "_" & txtbatchtag.Text & ".txt"
            Dim str_fromPath As String = path
            systemname = Environment.MachineName
            ToPath = "\\" & systemname & "\\" & getDefaultPrinter()
            objAddOn.objApplication.SetStatusBarMessage(ToPath, SAPbouiCOM.BoMessageTime.bmt_Short, False)
            File.Copy(str_fromPath, ToPath)
        Catch ex As Exception
            objAddOn.objApplication.MessageBox(ToPath & " Ex: " & ex.Message)
        End Try
        Dim j As Integer = 1

        While j <= objMatrix.RowCount
            'updateStatus(FormUID, j, j + 1)
            objMatrix.DeleteRow(j)
        End While
        'End While
        ' objAddOn.objApplication.MessageBox(CStr(count) + " Barcode(s) printed")
        objAddOn.objApplication.MessageBox("Barcode file(s) generated")
    End Sub
    Private Sub updatePrice(ByVal FormUID As String, ByVal intloop As Integer)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("14").Specific
        objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        strSQL = "update OSRN set U_MRP =" & CDbl(objMatrix.Columns.Item("3").Cells.Item(intloop).Specific.string) & _
        " where distnumber='" & objMatrix.Columns.Item("2").Cells.Item(intloop).Specific.string & "'"
        objRS.DoQuery(strSQL)
        objRS = Nothing
    End Sub
  
    'Private Function getDefaultPrinter() As String
    '    Dim printer As String
    '    Dim settings As PrinterSettings
    '    settings = New PrinterSettings

    '    For Each printer In PrinterSettings.InstalledPrinters
    '        settings.PrinterName = printer
    '        If settings.IsDefaultPrinter Then
    '            Return printer
    '        End If
    '    Next
    '    Return ""
    'End Function
    Private Function getDefaultPrinter() As String
        Dim printer As String = ""
        Dim settings As PrinterSettings
        ' Dim prinqu As New PrintQueue

        printer = "TSP"
        settings = New PrinterSettings
        'Dim prThis As ApiPrinter
        'For Each printer In PrinterSettings.InstalledPrinters
        '    settings.PrinterName = printer

        '    If settings.IsDefaultPrinter Then
        '        Return printer
        '    End If
        'Next



        Return printer

        '--------------


    End Function
    Private Sub BarcodeCFL(ByVal FormUID As String, ByVal pval As SAPbouiCOM.ItemEvent)
        Dim CFLEvent As SAPbouiCOM.ChooseFromListEvent
        Dim objDATATable As SAPbouiCOM.DataTable
        CFLEvent = pval
        objDATATable = CFLEvent.SelectedObjects
        Try
            If Not objDATATable Is Nothing Then
                If pval.ItemUID = "7" Then
                    objForm.DataSources.UserDataSources.Item("US1").ValueEx = objDATATable.GetValue("ItemCode", 0)
                ElseIf pval.ItemUID = "11" Then
                    objForm.DataSources.UserDataSources.Item("US3").ValueEx = objDATATable.GetValue("WhsCode", 0)
                ElseIf pval.ItemUID = "9" Then
                    objForm.DataSources.UserDataSources.Item("US2").ValueEx = objDATATable.GetValue("IntrSerial", 0)
                ElseIf pval.ItemUID = "13" Then
                    objForm.DataSources.UserDataSources.Item("US4").ValueEx = objDATATable.GetValue("IntrSerial", 0)
                End If
            End If
        Catch ex As Exception
            ' MsgBox(ex.ToString)
        End Try
    End Sub
    Private Function validate(ByVal FormUID As String) As Boolean
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("14").Specific
        If objMatrix.VisualRowCount = 0 Then
            objAddOn.objApplication.SetStatusBarMessage("Atleast one barcode should be selected to print", SAPbouiCOM.BoMessageTime.bmt_Short)
            Return False
        End If
        Return True
    End Function
End Class
