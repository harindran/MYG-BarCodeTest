Imports System.IO
Imports System.Drawing.Printing
Imports System.printing
Imports BarcodeLib
Imports System.Text.RegularExpressions
Imports SAPbouiCOM.Framework

Public Class ClsPrintBarcode
    Public Const formtype = "BarPrint"
    Dim objForm As SAPbouiCOM.Form
    Dim objRS As SAPbobsCOM.Recordset
    Dim strSQL As String
    Dim objRS1 As SAPbobsCOM.Recordset
    Dim strSQL1 As String
    Dim objCombo As SAPbouiCOM.ComboBox
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim GRPODocNum As String
    Dim objComboType As SAPbouiCOM.ComboBox
    Public Sub LoadScreen()
        Try
            objForm = objAddOn.objUIXml.LoadScreenXML("BarCodePrint.xml", Mukesh.SBOLib.UIXML.enuResourceType.Embeded, formtype)
            AddCFLCondition(objForm.UniqueID)
            AddCFLCondition1(objForm.UniqueID)
            AddPrinters(objForm.UniqueID)
            objMatrix = objForm.Items.Item("13").Specific
            objMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            objMatrix = objForm.Items.Item("6").Specific
            objMatrix.Columns.Item("5").Visible = True
            ' objForm.Visible = True
        Catch ex As Exception
            objAddOn.objApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, False)
        End Try
    End Sub
    Public Sub ItemEvent(ByVal FormUID As String, ByRef Pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If Pval.BeforeAction = True Then
                Select Case Pval.EventType
                    Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                        'objForm = objAddOn.objApplication.Forms.Item(FormUID) ' to be rectified 27-Dec-2018
                        'If Pval.ItemUID = "8" Then
                        '    objMatrix = objForm.Items.Item("6").Specific
                        '    If objMatrix.RowCount > 0 Then
                        '        If CInt(objForm.Items.Item("8").Specific.string) > CInt(objMatrix.RowCount) Then
                        '            objAddOn.objApplication.SetStatusBarMessage("Cannot be more than Pending quantity", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        '            BubbleEvent = False
                        '        End If
                        '    End If
                        'End If
                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                        Dim objEdit As SAPbouiCOM.EditText
                        Dim oEmptyConds As New SAPbouiCOM.Conditions
                        Dim objComboType As SAPbouiCOM.ComboBox
                        objComboType = objForm.Items.Item("6B").Specific
                        objEdit = objForm.Items.Item("4").Specific
                        Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
                        oCFLCreationParams = objAddOn.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
                        Try
                            If objComboType.Selected.Value = "OPDN" Then
                                objEdit.ChooseFromListUID.Remove(Pval.Row)
                                objEdit.ChooseFromListUID = "CFL_2"
                                objEdit.ChooseFromListAlias = "DocNum"
                            ElseIf objComboType.Selected.Value = "OIGN" Then
                                objEdit.ChooseFromListUID.Remove(Pval.Row)
                                objEdit.ChooseFromListUID = "CFL_3"
                                objEdit.ChooseFromListAlias = "DocNum"

                            End If
                        Catch ex As Exception
                        End Try
                    Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                        objMatrix = objForm.Items.Item("13").Specific
                        objComboType = objForm.Items.Item("6B").Specific
                        Dim ColItem As SAPbouiCOM.Column = objMatrix.Columns.Item("1")
                        Dim objlink As SAPbouiCOM.LinkedButton = ColItem.ExtendedObject
                        Try
                            If objComboType.Selected.Value = "OPDN" Then
                                objlink.LinkedObjectType = "20"
                                objlink.Item.LinkTo = "DocEntry"
                            ElseIf objComboType.Selected.Value = "OIGN" Then
                                objlink.LinkedObjectType = "59"
                                objlink.Item.LinkTo = "DocEntry"
                            End If
                        Catch ex As Exception
                        End Try
                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        If Pval.ItemUID = "11" Then
                            If validate(FormUID) = False Then
                                BubbleEvent = False : Exit Sub
                            End If
                        End If

                End Select
            ElseIf Pval.BeforeAction = False Then
                Select Case Pval.EventType

                    Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                        If objForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized Then
                            Resize(FormUID)
                            BubbleEvent = False
                        ElseIf objForm.State = SAPbouiCOM.BoFormStateEnum.fs_Restore Then
                            Restore(FormUID)
                            BubbleEvent = False
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                        objComboType = objForm.Items.Item("6B").Specific
                        If Pval.ItemUID = "4" Then
                            'cfl(FormUID, Pval)
                            LoadGRPO(FormUID, Pval, objComboType.Selected.Value)
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        objForm = objAddOn.objApplication.Forms.Item(FormUID)
                        objComboType = objForm.Items.Item("6B").Specific
                        If Pval.ItemUID = "11" Then
                            If validate(FormUID) Then
                                Dim count As Integer = 0
                                objMatrix = objForm.Items.Item("6").Specific
                                count = objMatrix.RowCount
                                SelectPrinterType(FormUID)
                                ' objAddOn.objApplication.MessageBox(CStr(count) + " Barcode(s) printed")
                            End If
                        ElseIf Pval.ItemUID = "101" Then
                            objForm.Close()
                        End If
                        If Pval.ItemUID = "13" And Pval.ColUID = "0" Then
                            LoadBarcode(FormUID, objComboType.Selected.Value)
                        End If
                        If Pval.ItemUID = "8A" Then
                            PrintGRPOBarcode(FormUID, objComboType.Selected.Value)
                        End If
                        If Pval.ItemUID = "8" Then
                            If objForm.Items.Item("8").Specific.String <> "" Then
                                LoadQuantity(FormUID)
                            End If
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                        If Pval.ItemUID = "8" Then
                            If objForm.Items.Item("8").Specific.String <> "" Then
                                LoadQuantity(FormUID)
                            End If
                            'RemoveRows(FormUID)
                        End If
                End Select
            End If
        Catch ex As Exception
            objAddOn.objApplication.MessageBox(ex.Message)
        End Try
    End Sub
    Private Sub SelectPrinterType(ByVal FormUID As String)
        Dim objCombo As SAPbouiCOM.ComboBox
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objCombo = objForm.Items.Item("18").Specific
        Dim FC As String = "F"
        objCombo = objForm.Items.Item("16").Specific
        If UCase(FC) = "C" Then
            ' Barcodeprint(FormUID, objCombo.Selected.Value) ' commented since we are not using crystal report
        ElseIf UCase(FC) = "F" Then
            Select Case objCombo.Selected.Value
                Case "BR"
                    'CreateFile(FormUID, "Rs. ")
                    CreateFileNew(FormUID, "Rs. ")
                Case "SM"
                    'CreateFile2(FormUID, "Rs. ")
                    'CreateFileNew(FormUID, "Rs. ")
                    CreateFileUpdatedPRN(FormUID, "Rs. ")
            End Select
        End If
    End Sub
    Private Sub LoadQuantity(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        Dim objMatrix As SAPbouiCOM.Matrix
        objMatrix = objForm.Items.Item("6").Specific
        For i As Integer = 1 To objMatrix.RowCount
            objMatrix.Columns.Item("5").Cells.Item(i).Specific.String = objForm.Items.Item("8").Specific.String
        Next
    End Sub
    Sub Barcodeprint(ByVal FormUID As String, ByVal BType As String)

    End Sub

    Private Sub code1()
        '     System.Drawing.Printing.PrinterSettings printerSettings = new System.Drawing.Printing.PrinterSettings();

        'printerSettings.PrinterName = cboCurrentPrinters.SelectedItem.ToString();

        '// don't use this, use the new button
        '//PrintLayout.Scaling = PrintLayoutSettings.PrintScaling.DoNotScale;

        'System.Drawing.Printing.PageSettings pSettings = new System.Drawing.Printing.PageSettings(printerSettings);
    End Sub
    Private Sub AddPrinters(ByVal FormUID As String)
        Try
            Dim objCombo, ObjComboType As SAPbouiCOM.ComboBox
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            objCombo = objForm.Items.Item("16").Specific
            ObjComboType = objForm.Items.Item("6B").Specific
            ObjComboType.ValidValues.Add("OPDN", "GRPO")
            ObjComboType.ValidValues.Add("OIGN", "GoodsReceipt")
            'ObjComboType.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            ObjComboType.Select("OPDN", SAPbouiCOM.BoSearchKey.psk_ByDescription)
            objCombo.ValidValues.Add("BR", "MyG")
            objCombo.ValidValues.Add("SM", "3G")
            objCombo.Select("SM", SAPbouiCOM.BoSearchKey.psk_ByDescription)
            objCombo = objForm.Items.Item("17").Specific
            Dim printer As String
            Dim settings As PrinterSettings
            settings = New PrinterSettings
            Dim CP As Integer = 0
            For Each printer In PrinterSettings.InstalledPrinters
                settings.PrinterName = printer
                CP += 1
                objCombo.ValidValues.Add(CStr(CP), settings.PrinterName)
                'If settings.IsDefaultPrinter Then
                '    Return printer
                'End If
            Next

            objCombo = objForm.Items.Item("18").Specific
            objCombo.ValidValues.Add("F", "File")
            objCombo.ValidValues.Add("C", "Crystal")
            objCombo.Select("F", SAPbouiCOM.BoSearchKey.psk_ByValue)

        Catch ex As Exception

        End Try
    End Sub
    Private Sub AddCFLCondition(ByVal FormUID As String)
        Try
            Dim conditions As SAPbouiCOM.Conditions
            Dim condition As SAPbouiCOM.Condition
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            conditions = objForm.ChooseFromLists.Item("CFL_2").GetConditions()
            If conditions.Count > 0 Then
                objForm.ChooseFromLists.Item("CFL_2").SetConditions(Nothing)
            End If
            'condition = conditions.Add
            'condition.Alias = "U_serial"
            'condition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'condition.CondVal = "Y"
            'condition.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            condition = conditions.Add
            condition.Alias = "U_bstatus"
            condition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            condition.CondVal = "Y"
            objForm.ChooseFromLists.Item("CFL_2").SetConditions(conditions)
        Catch ex As Exception

        End Try
       
    End Sub
    Private Sub AddCFLCondition1(ByVal FormUID As String)
        Try
            Dim conditions As SAPbouiCOM.Conditions
            Dim condition As SAPbouiCOM.Condition
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            conditions = objForm.ChooseFromLists.Item("CFL_3").GetConditions()
            If conditions.Count > 0 Then
                objForm.ChooseFromLists.Item("CFL_3").SetConditions(Nothing)
            End If
            'condition = conditions.Add
            'condition.Alias = "U_serial"
            'condition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'condition.CondVal = "Y"
            'condition.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            condition = conditions.Add
            condition.Alias = "U_bstatus"
            condition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            condition.CondVal = "Y"
            objForm.ChooseFromLists.Item("CFL_3").SetConditions(conditions)
        Catch ex As Exception

        End Try

    End Sub
    Private Sub Resize(ByVal FormUID As String)
        'maximize
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objForm.Freeze(True)
        'width="751" top="74" height="117"
        '  objForm.Items.Item("13").Width = objForm.Width - 30
        objForm.Items.Item("13").Top = 74
        objForm.Items.Item("13").Height = 117

        'width="270" top="280" height="130"
        ' objForm.Items.Item("6").Width = objForm.Width - 30
        objForm.Items.Item("6").Top = 280
        '  objForm.Items.Item("6").Height = 130
        objForm.Freeze(False)
    End Sub
    Private Sub Restore(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objForm.Items.Item("13").Width = 751
        objForm.Items.Item("13").Height = 117

        objForm.Items.Item("6").Top = 280
        objForm.Items.Item("6").Width = 270
        objForm.Items.Item("6").Height = 130
        objMatrix = objForm.Items.Item("6").Specific
        objMatrix.AutoResizeColumns()
    End Sub
    Private Sub RemoveRows(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("6").Specific
        If CInt(objForm.Items.Item("8").Specific.string) < CInt(objMatrix.RowCount) Then
            While objMatrix.RowCount > CInt(objForm.Items.Item("8").Specific.string)
                objMatrix.DeleteRow(objMatrix.RowCount)
            End While
        End If
    End Sub
    Private Sub PrintGRPOBarcode(ByVal FormUID As String, ByVal DocType As String)
       
        Dim strSQL As String = ""
        Dim objmatrix As SAPbouiCOM.Matrix
        Dim objDTable As SAPbouiCOM.DataTable
        Dim BaseEntry As Integer
        Dim Itemcode As String = ""
        Dim Header, Line As String
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objmatrix = objForm.Items.Item("13").Specific
        If DocType = "OPDN" Then
            Header = "OPDN"
            Line = "PDN1"
        Else
            Header = "OIGN"
            Line = "IGN1"
        End If
        Try
            'Itemcode = objmatrix.Columns.Item("4").Cells.Item(1).Specific.string
            BaseEntry = CInt(objmatrix.Columns.Item("1").Cells.Item(1).Specific.string)
            Dim Linenum As Integer
            ' Linenum = CInt(objmatrix.Columns.Item("12").Cells.Item(1).Specific.string)
            Dim ManagedBy As String = Trim(objmatrix.Columns.Item("13").Cells.Item(1).Specific.String)
            objmatrix = objForm.Items.Item("6").Specific
            objmatrix.Clear()
            objDTable = objForm.DataSources.DataTables.Item("T3")
            If ManagedBy = "S" Then
                If objAddOn.HANA Then
                    strSQL = " Select * from (SELECT distinct T4.""IntrSerial"" ""BatchNum"",T5.""Price"" as ""price"", 'Y' AS ""sel"",T1.""Dscription"", "
                    strSQL += vbCrLf + "  CAST(T4.""Quantity"" AS varchar(10)) AS ""qty"", T0.""CardCode"", To_varchar(T0.""DocDate"",'Mon-yyyy') AS ""year1"",I1.""BaseEntry"" , I1.""BaseLinNum"", T1.""ItemCode"",T4.""U_printed"",T4.""Status""  "
                    strSQL += vbCrLf + "   from " & Header & " T0 left join " & Line & " T1 on T0.""DocEntry""=T1.""DocEntry"" left outer join SRI1 I1 on T1.""ItemCode""=I1.""ItemCode"" and (T1.""DocEntry""=I1.""BaseEntry"" and T1.""ObjType""=I1.""BaseType"") and T1.""LineNum""=I1.""BaseLinNum"""
                    strSQL += vbCrLf + "  left outer join OSRI T4 on T4.""ItemCode""=I1.""ItemCode"" and I1.""SysSerial""=T4.""SysSerial"" and I1.""WhsCode"" = T4.""WhsCode"" left outer join ITM1 T5 on T5.""ItemCode""=I1.""ItemCode"" and T5.""PriceList""='5')A "
                    strSQL += vbCrLf + " Where A.""BaseEntry"" = " & BaseEntry & "  AND A.""Status"" IN (0,2)  ;    "
                Else
                    strSQL = "select T0.intrserial,(select top 1 price from ITM1 where PriceList=5 and ItemCode='" & Itemcode & "') as price,'Y' as sel,T5.Dscription ,convert(varchar(10),T5.Quantity) as qty,T6.CardCode ,format(T6.DocDate,'MMM-yyyy') as year1 ,T0.BaseEntry,T0.BaseLinNum,T5.ItemCode from OSRI T0 " &
                " join PDN1 T5 on T5.docentry=T0.baseentry and T5.linenum=T0.Baselinnum and T5.whscode= T0.Whscode " & _
               " join OPDN T6 on T6.DocEntry=T5.DocEntry " & _
                " where T0.BaseEntry =" & BaseEntry & " And T0.status = 0 and (T0.U_printed<>'Y' or T0.U_printed is null) and T0.BaseLinnum=" & Linenum
                End If
                objDTable.ExecuteQuery(strSQL)
            End If

            If ManagedBy = "B" Then
                If objAddOn.HANA Then
                    strSQL = " Select * from (SELECT distinct I1.""BatchNum"" ""BatchNum"",T5.""Price"" as ""price"", 'Y' AS ""sel"",T1.""Dscription"", "
                    strSQL += vbCrLf + "  CAST(I1.""Quantity"" AS varchar(10)) AS ""qty"", T0.""CardCode"", To_varchar(T0.""DocDate"",'Mon-yyyy') AS ""year1"",I1.""BaseEntry"" , I1.""BaseLinNum"", T1.""ItemCode"",T4.""U_printed"",T4.""Status""  "
                    strSQL += vbCrLf + "  from " & Header & " T0 left join " & Line & " T1 on T0.""DocEntry""=T1.""DocEntry"" left outer join IBT1 I1 on T1.""ItemCode""=I1.""ItemCode"" and (T1.""DocEntry""=I1.""BaseEntry"" and T1.""ObjType""=I1.""BaseType"") and T1.""LineNum""=I1.""BaseLinNum"""
                    strSQL += vbCrLf + " left outer join OIBT T4 on T4.""ItemCode""=I1.""ItemCode"" and I1.""BatchNum""=T4.""BatchNum"" and I1.""WhsCode"" = T4.""WhsCode"" left outer join ITM1 T5 on T5.""ItemCode""=I1.""ItemCode"" and T5.""PriceList""='5') A "
                    strSQL += vbCrLf + " Where A.""BaseEntry"" = " & BaseEntry & "  AND A.""Status"" IN (0,2)  ;    "
                Else
                    strSQL = "select case when isnull(T0.U_EAN,'')='' then T0.BatchNum else T0.U_EAN end as BatchNum,(select top 1 price from ITM1 where PriceList=5 and ItemCode='" & Itemcode & "') as price,'Y' as sel,T5.Dscription ,convert(varchar(10),T5.Quantity) as qty,T6.CardCode ,format(T6.DocDate,'MMM-yyyy') as year1  ,T0.BaseEntry,T0.BaseLinNum, T5.ItemCode from OIBT T0 " & _
                " join PDN1 T5 on T5.docentry=T0.baseentry and T5.linenum=T0.Baselinnum and T5.whscode= T0.Whscode " & _
               " join OPDN T6 on T6.DocEntry=T5.DocEntry " & _
                " where T0.BaseEntry =" & BaseEntry & " And T0.status = 0 and (T0.U_printed<>'Y' or T0.U_printed is null) and T0.BaseLinNum=" & Linenum
                End If

                Dim objRS1 As SAPbobsCOM.Recordset
                objRS1 = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                objRS1.DoQuery(strSQL)

                objDTable.Rows.Clear()
                While Not objRS1.EoF
                    objDTable.Rows.Add()

                    objDTable.SetValue(0, objDTable.Rows.Count - 1, objRS1.Fields.Item("BatchNum").Value)
                    objDTable.SetValue(1, objDTable.Rows.Count - 1, objRS1.Fields.Item("price").Value)
                    objDTable.SetValue(2, objDTable.Rows.Count - 1, objRS1.Fields.Item("sel").Value)
                    objDTable.SetValue(3, objDTable.Rows.Count - 1, objRS1.Fields.Item("Dscription").Value)
                    objDTable.SetValue(4, objDTable.Rows.Count - 1, objRS1.Fields.Item("qty").Value)
                    objDTable.SetValue(5, objDTable.Rows.Count - 1, objRS1.Fields.Item("CardCode").Value)
                    objDTable.SetValue(6, objDTable.Rows.Count - 1, objRS1.Fields.Item("year1").Value)
                    objDTable.SetValue(7, objDTable.Rows.Count - 1, objRS1.Fields.Item("BaseEntry").Value)
                    objDTable.SetValue(8, objDTable.Rows.Count - 1, objRS1.Fields.Item("BaseLinNum").Value)
                    objDTable.SetValue(9, objDTable.Rows.Count - 1, objRS1.Fields.Item("ItemCode").Value)

                    objRS1.MoveNext()
                End While
            End If
            objmatrix.LoadFromDataSource()
            If objmatrix.RowCount = 0 Then
                objAddOn.objApplication.SetStatusBarMessage("No barcoded items exist in the selected GRPO")
                Exit Sub
            End If
            objmatrix = objForm.Items.Item("13").Specific
            objForm.Items.Item("8").Specific.string = ""
        Catch ex As Exception
            '   MsgBox(ex.ToString)
        End Try
    End Sub
    Public Sub LoadBarcode(ByVal FormUID As String, ByVal DocType As String)
        Dim strSQL As String = ""
        Dim objmatrix As SAPbouiCOM.Matrix
        Dim objDTable As SAPbouiCOM.DataTable
        Dim BaseEntry As Integer
        Dim Itemcode As String = ""
        Dim intloop As Integer
        Dim Header, Line As String
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objmatrix = objForm.Items.Item("13").Specific
        For intloop = 1 To objmatrix.RowCount
            If objmatrix.IsRowSelected(intloop) Then
                GRPODocNum = objmatrix.Columns.Item("1").Cells.Item(intloop).Specific.String
                Exit For
            End If
        Next
        If DocType = "OPDN" Then
            Header = "OPDN"
            Line = "PDN1"
        Else
            Header = "OIGN"
            Line = "IGN1"
        End If
        Try
            Itemcode = objmatrix.Columns.Item("4").Cells.Item(intloop).Specific.string
            BaseEntry = CInt(objmatrix.Columns.Item("1").Cells.Item(intloop).Specific.string)
            Dim Linenum As Integer
            Linenum = CInt(objmatrix.Columns.Item("12").Cells.Item(intloop).Specific.string)
            Dim ManagedBy As String = Trim(objmatrix.Columns.Item("13").Cells.Item(intloop).Specific.String)
            objmatrix = objForm.Items.Item("6").Specific
            objmatrix.Clear()
            objDTable = objForm.DataSources.DataTables.Item("T3")

            If ManagedBy = "S" Then
                If objAddOn.HANA Then
                    'strSQL = "SELECT T0.""IntrSerial"", (select top 1 ""Price"" from ITM1 where ""ItemCode""='" & Itemcode & "' and ""PriceList""=5) AS ""price"", 'Y' AS ""sel"", T5.""Dscription"", CAST(T5.""Quantity"" AS varchar(10)) AS ""qty"", T6.""CardCode"", To_varchar(T6.""DocDate"",'Mon-yyyy') AS year1, T0.""BaseEntry"", T0.""BaseLinNum"",T5.""ItemCode"" FROM OSRI T0 " & _
                    '    " INNER JOIN PDN1 T5 ON T5.""DocEntry"" = T0.""BaseEntry"" AND T5.""LineNum"" = T0.""BaseLinNum"" AND T5.""WhsCode"" = T0.""WhsCode"" " & _
                    '    " INNER JOIN OPDN T6 ON T6.""DocEntry"" = T5.""DocEntry"" WHERE T0.""BaseEntry"" = " & BaseEntry & " AND T0.""Status"" IN (0,2) " & _
                    '    " AND (T0.""U_printed"" <> 'Y' OR T0.""U_printed"" IS NULL) AND T0.""BaseLinNum"" = " & Linenum & ";"

                    strSQL = " Select * from (SELECT distinct T4.""IntrSerial"" ""IntrSerial"",(select top 1 ""Price"" from ITM1 where ""ItemCode""='" & Itemcode & "' and ""PriceList""=5) AS ""price"", 'Y' AS ""sel"",T1.""Dscription"","
                    strSQL += vbCrLf + " CAST(T1.""Quantity"" AS varchar(10)) AS ""qty"", T0.""CardCode"", To_varchar(T0.""DocDate"",'Mon-yyyy') AS year1,T4.""BaseEntry"", T4.""BaseLinNum"",T1.""ItemCode"",T4.""U_printed"",T4.""Status"" "
                    strSQL += vbCrLf + " from " & Header & " T0 inner join " & Line & " T1 on T0.""DocEntry""=T1.""DocEntry"" left outer join SRI1 I1 on T1.""ItemCode""=I1.""ItemCode"" and (T1.""DocEntry""=I1.""BaseEntry"" and T1.""ObjType""=I1.""BaseType"") and T1.""LineNum""=I1.""BaseLinNum"" "
                    strSQL += vbCrLf + " left outer join OSRI T4 on T4.""ItemCode""=I1.""ItemCode"" and I1.""SysSerial""=T4.""SysSerial"" and I1.""WhsCode"" = T4.""WhsCode"" ) A   "
                    strSQL += vbCrLf + " Where A.""BaseEntry"" = " & BaseEntry & " and A.""BaseLinNum""=" & Linenum & " AND A.""Status"" IN (0,2)  ;    "
                Else
                    strSQL = "select T0.intrserial,(select top 1 price from ITM1 where PriceList=5 and ItemCode='" & Itemcode & "') as price,'Y' as sel,T5.Dscription ,convert(varchar(10),T5.Quantity) as qty,T6.CardCode ,format(T6.DocDate,'MMM-yyyy') as year1 ,T0.BaseEntry,T0.BaseLinNum,T5.ItemCode from OSRI T0 " &
                " join PDN1 T5 on T5.docentry=T0.baseentry and T5.linenum=T0.Baselinnum and T5.whscode= T0.Whscode " & _
               " join OPDN T6 on T6.DocEntry=T5.DocEntry " & _
                " where T0.BaseEntry =" & BaseEntry & " And T0.status = 0 and (T0.U_printed<>'Y' or T0.U_printed is null) and T0.BaseLinnum=" & Linenum
                End If

                ' If objAddOn.HANA Then
                '     strSQL = "SELECT T0.""IntrSerial"", IFNULL(T0.U_MRP, T5.U_MRP) AS ""price"", 'Y' AS ""sel"", T5.""Dscription"", CAST(T5.""Quantity"" AS varchar(10)) AS ""qty"", T6.""CardCode"", CAST(YEAR(T6.""DocDate"") AS varchar()) AS ""year1"", T0.""BaseEntry"", T0.""BaseLinNum"" FROM OSRI T0 " & _
                '         " INNER JOIN PDN1 T5 ON T5.""DocEntry"" = T0.""BaseEntry"" AND T5.""LineNum"" = T0.""BaseLinNum"" AND T5.""WhsCode"" = T0.""WhsCode"" " & _
                '         " INNER JOIN OPDN T6 ON T6.""DocEntry"" = T5.""DocEntry"" WHERE T0.""BaseEntry"" = " & BaseEntry & " AND T0.""Status"" IN (0,2) " & _
                '         " AND (T0.""U_printed"" <> 'Y' OR T0.""U_printed"" IS NULL) AND T0.""BaseLinNum"" = " & Linenum & ";"
                ' Else
                '     strSQL = "select T0.intrserial,isnull(T0.U_MRP,T5.U_MRP ) as price,'Y' as sel,T5.Dscription ,convert(varchar(10),T5.Quantity) as qty,T6.CardCode ,convert(varchar(5),YEAR(T6.DocDate)) as year1 ,T0.BaseEntry,T0.BaseLinNum from OSRI T0 " & _
                ' " join PDN1 T5 on T5.docentry=T0.baseentry and T5.linenum=T0.Baselinnum and T5.whscode= T0.Whscode " & _
                '" join OPDN T6 on T6.DocEntry=T5.DocEntry " & _
                ' " where T0.BaseEntry =" & BaseEntry & " And T0.status = 0 and (T0.U_printed<>'Y' or T0.U_printed is null) and T0.BaseLinnum=" & Linenum
                ' End If

                objDTable.ExecuteQuery(strSQL)

                'Dim objRS1 As SAPbobsCOM.Recordset
                'objRS1 = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                'objRS1.DoQuery(strSQL)
                'objDTable.Rows.Clear()
                'While Not objRS1.EoF
                '    For i As Integer = 1 To CLng(objRS1.Fields.Item("qty").Value)
                '        'For i As Integer = 1 To objRS1.RecordCount
                '        objDTable.Rows.Add()

                '        objDTable.SetValue(0, objDTable.Rows.Count - 1, objRS1.Fields.Item("BatchNum").Value)
                '        objDTable.SetValue(1, objDTable.Rows.Count - 1, objRS1.Fields.Item("price").Value)
                '        objDTable.SetValue(2, objDTable.Rows.Count - 1, objRS1.Fields.Item("sel").Value)
                '        objDTable.SetValue(3, objDTable.Rows.Count - 1, objRS1.Fields.Item("Dscription").Value)
                '        objDTable.SetValue(4, objDTable.Rows.Count - 1, objRS1.Fields.Item("qty").Value)
                '        objDTable.SetValue(5, objDTable.Rows.Count - 1, objRS1.Fields.Item("CardCode").Value)
                '        objDTable.SetValue(6, objDTable.Rows.Count - 1, objRS1.Fields.Item("year1").Value)
                '        objDTable.SetValue(7, objDTable.Rows.Count - 1, objRS1.Fields.Item("BaseEntry").Value)
                '        objDTable.SetValue(8, objDTable.Rows.Count - 1, objRS1.Fields.Item("BaseLinNum").Value)
                '        objDTable.SetValue(9, objDTable.Rows.Count - 1, objRS1.Fields.Item("ItemCode").Value)
                '    Next
                '    objRS1.MoveNext()
                'End While
            End If

            ' Select * from (SELECT distinct I1."BatchNum" "BatchNum",(select top 1 "Price" from ITM1 where "ItemCode"='PRINT0009' and "PriceList"=5) AS "price", 'Y' AS "sel",T1."Dscription",
            ' CAST(T1."Quantity" AS varchar(10)) AS "qty", T0."CardCode", To_varchar(T0."DocDate",'Mon-yyyy') AS "year1",I1."BaseEntry" , I1."BaseLinNum", T1."ItemCode",T0."U_printed"  from 
            ' OPDN T0 left join PDN1 T1 on T0."DocEntry"=T1."DocEntry" left outer join IBT1 I1 on T1."ItemCode"=I1."ItemCode" and (T1."DocEntry"=I1."BaseEntry" and T1."ObjType"=I1."BaseType") and T1."LineNum"=I1."BaseLinNum"
            ' left outer join OIBT T4 on T4."ItemCode"=I1."ItemCode" and I1."BatchNum"=T4."BatchNum" and I1."WhsCode" = T4."WhsCode"
            ')A Where A."BaseEntry" = 97 and A."BaseLinNum"=0 AND T4."Status" IN (0,2) AND (A."U_printed" <> 'Y' OR A."U_printed" IS NULL)  ;  
            If ManagedBy = "B" Then
                If objAddOn.HANA Then
                    'strSQL = "SELECT CASE WHEN IFNULL(T0.""U_EAN"",'')='' THEN T0.""BatchNum"" ELSE T0.""U_EAN"" END AS ""BatchNum"" , (select top 1 ""Price"" from ITM1 where ""ItemCode""='" & Itemcode & "' and ""PriceList""=5) AS ""price"", 'Y' AS ""sel"",T5.""Dscription"", CAST(T5.""Quantity"" AS varchar(10)) AS ""qty"", T6.""CardCode"", To_varchar(T6.""DocDate"",'Mon-yyyy') AS ""year1"", T0.""BaseEntry"", T0.""BaseLinNum"", T5.""ItemCode"" FROM OIBT T0 " & _
                    '    " INNER JOIN PDN1 T5 ON T5.""DocEntry"" = T0.""BaseEntry"" AND T5.""LineNum"" = T0.""BaseLinNum"" AND T5.""WhsCode"" = T0.""WhsCode"" " & _
                    '    " INNER JOIN OPDN T6 ON T6.""DocEntry"" = T5.""DocEntry"" WHERE T0.""BaseEntry"" = " & BaseEntry & " AND T0.""Status"" IN (0,2) " & _
                    '    " AND (T0.""U_printed"" <> 'Y' OR T0.""U_printed"" IS NULL) AND T0.""BaseLinNum"" = " & Linenum & ";"

                    strSQL = " Select * from (SELECT distinct I1.""BatchNum"" ""BatchNum"",(select top 1 ""Price"" from ITM1 where ""ItemCode""='" & Itemcode & "' and ""PriceList""=5) AS ""price"", 'Y' AS ""sel"",T1.""Dscription"", "
                    strSQL += vbCrLf + " CAST(T1.""Quantity"" AS varchar(10)) AS ""qty"", T0.""CardCode"", To_varchar(T0.""DocDate"",'Mon-yyyy') AS ""year1"",I1.""BaseEntry"" , I1.""BaseLinNum"", T1.""ItemCode"",T4.""U_printed"",T4.""Status""  "
                    strSQL += vbCrLf + " from " & Header & " T0 left join " & Line & " T1 on T0.""DocEntry""=T1.""DocEntry"" left outer join IBT1 I1 on T1.""ItemCode""=I1.""ItemCode"" and (T1.""DocEntry""=I1.""BaseEntry"" and T1.""ObjType""=I1.""BaseType"") and T1.""LineNum""=I1.""BaseLinNum"""
                    strSQL += vbCrLf + " left outer join OIBT T4 on T4.""ItemCode""=I1.""ItemCode"" and I1.""BatchNum""=T4.""BatchNum"" and I1.""WhsCode"" = T4.""WhsCode"")A "
                    strSQL += vbCrLf + " Where A.""BaseEntry"" = " & BaseEntry & " and A.""BaseLinNum""=" & Linenum & " AND A.""Status"" IN (0,2)  ;  "
                Else
                    strSQL = "select case when isnull(T0.U_EAN,'')='' then T0.BatchNum else T0.U_EAN end as BatchNum,(select top 1 price from ITM1 where PriceList=5 and ItemCode='" & Itemcode & "') as price,'Y' as sel,T5.Dscription ,convert(varchar(10),T5.Quantity) as qty,T6.CardCode ,format(T6.DocDate,'MMM-yyyy') as year1  ,T0.BaseEntry,T0.BaseLinNum, T5.ItemCode from OIBT T0 " & _
                " join PDN1 T5 on T5.docentry=T0.baseentry and T5.linenum=T0.Baselinnum and T5.whscode= T0.Whscode " & _
               " join OPDN T6 on T6.DocEntry=T5.DocEntry " & _
                " where T0.BaseEntry =" & BaseEntry & " And T0.status = 0 and (T0.U_printed<>'Y' or T0.U_printed is null) and T0.BaseLinNum=" & Linenum
                End If

                ' If objAddOn.HANA Then
                '     strSQL = "SELECT CASE WHEN IFNULL(T0.""U_EAN"",'')='' THEN T0.""BatchNum"" ELSE T0.""U_EAN"" END AS ""BatchNum"" , IFNULL(T0.U_MRP, T5.U_MRP) AS ""price"", 'Y' AS ""sel"", T5.""Dscription"", CAST(T5.""Quantity"" AS varchar(10)) AS ""qty"", T6.""CardCode"", CAST(YEAR(T6.""DocDate"") AS varchar(5)) AS ""year1"", T0.""BaseEntry"", T0.""BaseLinNum"" FROM OIBT T0 " & _
                '         " INNER JOIN PDN1 T5 ON T5.""DocEntry"" = T0.""BaseEntry"" AND T5.""LineNum"" = T0.""BaseLinNum"" AND T5.""WhsCode"" = T0.""WhsCode"" " & _
                '         " INNER JOIN OPDN T6 ON T6.""DocEntry"" = T5.""DocEntry"" WHERE T0.""BaseEntry"" = " & BaseEntry & " AND T0.""Status"" IN (0,2) " & _
                '         " AND (T0.""U_printed"" <> 'Y' OR T0.""U_printed"" IS NULL) AND T0.""BaseLinNum"" = " & Linenum & ";"
                ' Else
                '     strSQL = "select case when isnull(T0.U_EAN,'')='' then T0.BatchNum else T0.U_EAN end as BatchNum,isnull(T0.U_MRP,T5.U_MRP ) as price,'Y' as sel,T5.Dscription ,convert(varchar(10),T5.Quantity) as qty,T6.CardCode ,convert(varchar(5),YEAR(T6.DocDate)) as year1 ,T0.BaseEntry,T0.BaseLinNum from OIBT T0 " & _
                ' " join PDN1 T5 on T5.docentry=T0.baseentry and T5.linenum=T0.Baselinnum and T5.whscode= T0.Whscode " & _
                '" join OPDN T6 on T6.DocEntry=T5.DocEntry " & _
                ' " where T0.BaseEntry =" & BaseEntry & " And T0.status = 0 and (T0.U_printed<>'Y' or T0.U_printed is null) and T0.BaseLinNum=" & Linenum
                ' End If
                Dim objRS1 As SAPbobsCOM.Recordset
                objRS1 = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                objRS1.DoQuery(strSQL)
                'For i As Integer = 0 To objDTable.Columns.Count - 1
                '    MsgBox(objDTable.Columns.Item(i).Name)
                'Next
                objDTable.Rows.Clear()
                While Not objRS1.EoF
                    'For i As Integer = 1 To CLng(objRS1.Fields.Item("qty").Value)

                    'For i As Integer = 1 To objRS1.RecordCount
                    objDTable.Rows.Add()

                    objDTable.SetValue(0, objDTable.Rows.Count - 1, objRS1.Fields.Item("BatchNum").Value)
                    objDTable.SetValue(1, objDTable.Rows.Count - 1, objRS1.Fields.Item("price").Value)
                    objDTable.SetValue(2, objDTable.Rows.Count - 1, objRS1.Fields.Item("sel").Value)
                    objDTable.SetValue(3, objDTable.Rows.Count - 1, objRS1.Fields.Item("Dscription").Value)
                    objDTable.SetValue(4, objDTable.Rows.Count - 1, objRS1.Fields.Item("qty").Value)
                    objDTable.SetValue(5, objDTable.Rows.Count - 1, objRS1.Fields.Item("CardCode").Value)
                    objDTable.SetValue(6, objDTable.Rows.Count - 1, objRS1.Fields.Item("year1").Value)
                    objDTable.SetValue(7, objDTable.Rows.Count - 1, objRS1.Fields.Item("BaseEntry").Value)
                    objDTable.SetValue(8, objDTable.Rows.Count - 1, objRS1.Fields.Item("BaseLinNum").Value)
                    objDTable.SetValue(9, objDTable.Rows.Count - 1, objRS1.Fields.Item("ItemCode").Value)

                    'objDTable.SetValue("price", objDTable.Rows.Count - 1, objRS1.Fields.Item("price").Value)
                    'objDTable.SetValue("select", objDTable.Rows.Count - 1, objRS1.Fields.Item("sel").Value)
                    'objDTable.SetValue("desc", objDTable.Rows.Count - 1, objRS1.Fields.Item("Dscription").Value)
                    'objDTable.SetValue("qty", objDTable.Rows.Count - 1, "1")
                    'objDTable.SetValue("vendor", objDTable.Rows.Count - 1, objRS1.Fields.Item("CardCode").Value)
                    'objDTable.SetValue("year1", objDTable.Rows.Count - 1, objRS1.Fields.Item("year1").Value)
                    'objDTable.SetValue("docentry", objDTable.Rows.Count - 1, objRS1.Fields.Item("BaseEntry").Value)
                    'objDTable.SetValue("lineid", objDTable.Rows.Count - 1, objRS1.Fields.Item("BaseLinNum").Value)
                    ' Next
                    objRS1.MoveNext()
                End While
            End If


            objmatrix.LoadFromDataSource()
            If objmatrix.RowCount = 0 Then
                objAddOn.objApplication.SetStatusBarMessage("No barcoded items exist in the selected GRPO")
                Exit Sub
            End If
            objmatrix = objForm.Items.Item("13").Specific
            objForm.Items.Item("8").Specific.string = objmatrix.Columns.Item("10").Cells.Item(intloop).Specific.string
        Catch ex As Exception
            '   MsgBox(ex.ToString)
        End Try
    End Sub
    Private Sub LoadGRPO(ByVal FormUID As String, ByVal pval As SAPbouiCOM.ItemEvent, ByVal DocType As String)
        Dim CFLEvent As SAPbouiCOM.ChooseFromListEvent
        Dim DataTable As SAPbouiCOM.DataTable
        Dim strSQL As String = ""
        Dim strDocNum As String = ""
        Dim intloop As Integer
        Dim Header, Line As String
        Try
            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            objMatrix = objForm.Items.Item("13").Specific
            CFLEvent = pval
            DataTable = CFLEvent.SelectedObjects
            If Not (DataTable Is Nothing) Then
                For intloop = 1 To DataTable.Rows.Count
                    strDocNum = strDocNum + CStr(DataTable.GetValue("DocEntry", intloop - 1)) + ","
                Next
            Else
                Exit Sub
            End If
            strDocNum = strDocNum.Remove(Len(strDocNum) - 1, 1)
            If DocType = "OPDN" Then
                Header = "OPDN"
                Line = "PDN1"
            Else
                Header = "OIGN"
                Line = "IGN1"
            End If
            If objAddOn.HANA Then
                strSQL = "SELECT T0.""DocEntry"", T0.""DocDate"", T4.""FirmName"", T1.""ItemCode"", T2.SWW, T1.""Dscription"", T1.""WhsCode"", CAST(T1.""Quantity"" - SUM(IFNULL(T3.""Quantity"", 0)) AS integer) AS ""Quantity"", " & _
                    " IFNULL(T1.""U_bqty"", 0) AS ""bqty"", CAST((T1.""Quantity"" - SUM(IFNULL(T3.""Quantity"", 0))) - IFNULL(T1.""U_bqty"", 0) AS integer) AS ""pending"", " & _
                    " T0.""DocNum"", T1.""LineNum"", CASE WHEN T2.""ManBtchNum"" = 'Y' THEN 'B' WHEN T2.""ManSerNum"" = 'Y' THEN 'S' ELSE 'N' END AS ""ManagedBy""  FROM " & Header & " T0 INNER JOIN " & Line & " T1 ON T1.""DocEntry"" = T0.""DocEntry"" INNER JOIN OITM T2 ON T2.""ItemCode"" = T1.""ItemCode"" " & _
                    " INNER JOIN OMRC T4 ON T4.""FirmCode"" = T2.""FirmCode"" LEFT OUTER JOIN ""RPD1"" T3 ON T3.""BaseEntry"" = T1.""DocEntry"" AND T3.""BaseLine"" = T1.""LineNum"" " & _
                    " WHERE T0.""DocEntry"" IN (" & strDocNum & ") AND IFNULL(T1.""U_bqty"", 0) <> T1.""Quantity"" " & _
                    " GROUP BY T0.""DocNum"", T0.""DocDate"", T4.""FirmName"", T1.""ItemCode"", T2.""SWW"", T1.""Dscription"", T1.""WhsCode"", T1.""Quantity"", IFNULL(T1.""U_bqty"", 0), CAST(T1.""Quantity"" - IFNULL(T1.""U_bqty"", 0) AS integer), T0.""DocEntry"", T1.""LineNum"",T2.""ManBtchNum"",T2.""ManSerNum""  " & _
                    " ORDER BY T0.""DocNum"", T1.""LineNum"" ;"
            Else
                strSQL = "select T0.DocEntry,T0.DocDate,T4.FirmName ,T1.ItemCode ,T2.SWW,T1.Dscription ,T1.Whscode ,convert(int,T1.Quantity-sum(isnull(t3.quantity,0))) as Quantity," & _
           " isnull(T1.U_bqty,0) as bqty,convert(int,(T1.Quantity-sum(isnull(t3.quantity,0)))-isnull(T1.U_bqty,0)) as pending,T0.docNum,T1.linenum,  case when ManBtchNum='Y' then 'B' when ManSerNum='Y' then 'S' else 'N' end as 'ManagedBy' from OPDN T0" & _
     " join PDN1 T1 on T1 .DocEntry =T0.DocEntry " & _
    " join OITM T2 on T2.ItemCode =T1.ItemCode join OMRC T4 on T4.FirmCode=T2.FirmCode " & _
    " left outer join rpd1 T3 on T3.baseentry = T1.docentry and T3.baseline=T1.linenum " & _
    " where T0.docentry in(" & strDocNum & ") and isnull(T1.U_bqty,0)<>T1.quantity" & _
    " group by T0.DocNum,T0.DocDate,T4.FirmName ,T1.ItemCode ,T2.SWW,T1.Dscription ,T1.Whscode ,T1.Quantity," & _
     " isnull(T1.U_bqty,0) ,convert(int,T1.quantity-isnull(T1.U_bqty,0)) ,T0.docentry,T1.linenum ,T2.ManBtchNum,T2.ManSerNum order by T0.docnum, T1.linenum "
            End If

            objForm.DataSources.DataTables.Item("T2").ExecuteQuery(strSQL)
            objMatrix.LoadFromDataSource()
            Dim objcombo As SAPbouiCOM.ComboBox = objForm.Items.Item("16").Specific
            Dim objcomboFile As SAPbouiCOM.ComboBox = objForm.Items.Item("18").Specific
            objcombo.Select("SM", SAPbouiCOM.BoSearchKey.psk_ByDescription)
            objcomboFile.Select("F", SAPbouiCOM.BoSearchKey.psk_ByDescription)
        Catch ex As Exception

        End Try
    End Sub
    Public Sub cfl(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent)
        Try
            Dim CFLEvent As SAPbouiCOM.ChooseFromListEvent
            Dim DataTable As SAPbouiCOM.DataTable
            objForm = objAddOn.objApplication.Forms.Item(FormUID)

            CFLEvent = pVal
            DataTable = CFLEvent.SelectedObjects
            Try
                If Not (DataTable Is Nothing) Then
                    objForm.Items.Item("4").Specific.value = DataTable.GetValue("DocNum", 0)
                End If
            Catch ex As Exception
                objForm.Items.Item("4").Specific.value = DataTable.GetValue("DocNum", 0)
            End Try

        Catch ex As Exception
            objAddOn.objApplication.SetStatusBarMessage(ex.Message)
        End Try
    End Sub

    Private Sub CreateFile(ByVal FormUID As String, ByVal Prefix As String)
        ' big
        Dim intloop As Integer
        Dim qty As String
        Dim desc As String
        Dim price As String
        'Dim price2 As Decimal
        Dim barcode As String
        Dim vendor As String
        Dim year1 As String
        Dim colorsize As String = ""
        Dim count As Integer = 0

        Dim path As String = ""
        Dim fs As FileStream
        Dim rName As String = ""
        rName = SystemInformation.UserName
        Dim Foldername As String


        Foldername = "D:" + "\" + rName + "\" + GRPODocNum + "\MyG"
        If Directory.Exists(Foldername) Then
        Else
            Directory.CreateDirectory(Foldername)
        End If


        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("6").Specific
        While objMatrix.RowCount > 0
            intloop = 1
            If objMatrix.Columns.Item("3").Cells.Item(intloop).Specific.checked = True Then
                barcode = objMatrix.Columns.Item("1").Cells.Item(intloop).Specific.string
                price = objMatrix.Columns.Item("2").Cells.Item(intloop).Specific.string
                desc = objMatrix.Columns.Item("4").Cells.Item(intloop).Specific.string
                'qty = objMatrix.Columns.Item("5").Cells.Item(intloop).Specific.string
                'vendor = objMatrix.Columns.Item("5A").Cells.Item(intloop).Specific.string
                'year1 = objMatrix.Columns.Item("5B").Cells.Item(intloop).Specific.string
                'year1 = Right(year1, 2)
                'barcode = Left(barcode, 10)
                'convert_value(price)
                price = price.Replace(",", "")
                price = Prefix + Mid(price, 1, (price.IndexOf(".") + 3))
                price = convert_value(price.Replace(" ", ""))
                ' price2 = CDbl(price)

                'qty = CStr(objMatrix.Columns.Item("7").Cells.Item(intloop).Specific.string) + "-" + CStr(CInt(qty))
                '--------------File Creation--------------------
                Dim Filetype As String = "BIG"

                Filetype = Filetype.Replace(" ", "")
                Dim Filename As String = Trim(objMatrix.Columns.Item("6").Cells.Item(intloop).Specific.string) + "_" + Filetype + "_" + (desc.Replace(" ", "")).Replace("/", "").Replace(".", "").Replace(":", "") + "_" + CStr(CInt(objMatrix.Columns.Item("7").Cells.Item(intloop).Specific.string) + 1) + "_" + CStr(count) + ".prn"

                'Create the file.
                ' path = Application.StartupPath + "\" + rName + "\mylabel.prn"
                path = Foldername + "\" + Filename
                ' Delete the file if it exists.
                If File.Exists(path) Then
                    File.Delete(path)
                End If
                fs = New FileStream(path, FileMode.Create, FileAccess.Write)

                '-----------------------------------------------------------
                Dim s As New StreamWriter(fs)
                ' s.WriteLine("")
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
                Dim systemname As String = ""
                ' Dim Str_FromPath As String = "C:\" & Environment.UserName & "\" & txtBarcode.Text & "_" & txtbatchtag.Text & ".txt"
                Dim str_fromPath As String = path
                systemname = Environment.MachineName
                Dim ToPath As String = "\\" & systemname & "\\" & getDefaultPrinter()

                File.Copy(Str_FromPath, ToPath)

                Dim objCombo As SAPbouiCOM.ComboBox
                objCombo = objForm.Items.Item("17").Specific

                updateStatus(FormUID, intloop, count + 1)
                count += 1
            End If
            objMatrix.DeleteRow(intloop)
        End While
        ' objAddOn.objApplication.MessageBox(CStr(count) + " Barcode(s) printed")
        objAddOn.objApplication.MessageBox("Barcode file(s) generated")
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
        'Foldername = "D:" + "\" + rName + "\" + GRPODocNum + "\MyG"
        Foldername = "E:" + "\" + rName + "\" + GRPODocNum + "\MyG"
        If Directory.Exists(Foldername) Then
        Else
            Directory.CreateDirectory(Foldername)
        End If


        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("6").Specific
        'While objMatrix.RowCount > 0
        intloop = 1

        ' If objMatrix.Columns.Item("3").Cells.Item(intloop).Specific.checked = True Then

        'year1 = year1.Replace("/", "-")
        ' year2 = convert_date(year1)
        'year1 = Right(year1, 5)
        'year3 = Date.ParseExact(year1, "dd-mm-yyyy",
        'System.Globalization.DateTimeFormatInfo.InvariantInfo)
        '                year2 = Format(year3, "MMM-yyyy")
        ' year1 = Format(objMatrix.Columns.Item("5B").Cells.Item(intloop).Specific, "mmm yyyy")
        ' year1 = Right(year1, 4)
        'year3 = Date.ParseExact(year1, "dd-MMM-yyyy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
        'year2 = Format(year1, "MMM-yyyy")

        'price = convert_value(price.Replace(" ", ""))
        'convert_val(price)
        ' price = CStr(price)
        'price = convert_val(price)
        'price2 = CDbl(price)

        'qty = CStr(objMatrix.Columns.Item("7").Cells.Item(intloop).Specific.string) + "-" + CStr(CInt(qty))
        '--------------File Creation--------------------
        Dim Filetype As String = "BIG"

        Filetype = Filetype.Replace(" ", "")
        'Dim Filename As String = Trim(objMatrix.Columns.Item("6").Cells.Item(intloop).Specific.string) + "_" + Filetype + "_" + (desc.Replace(" ", "")).Replace("/", "").Replace(".", "").Replace(":", "") + "_" + CStr(CInt(objMatrix.Columns.Item("7").Cells.Item(intloop).Specific.string) + 1) + "_" + CStr(count) + ".prn"
        Dim Filename As String = Trim(objMatrix.Columns.Item("8").Cells.Item(1).Specific.string) + "_" + Filetype + "_" + System.DateTime.Now.ToString("yyMMddHHmmss") + "_" + CStr(count) + ".prn"
        'Create the file.
        ' path = Application.StartupPath + "\" + rName + "\mylabel.prn"
        path = Foldername + "\" + Filename
        ' Delete the file if it exists.
        If File.Exists(path) Then
            File.Delete(path)
        End If
        fs = New FileStream(path, FileMode.Create, FileAccess.Write)

        Dim s As New StreamWriter(fs)



        '-----------------------------------------------------------

        ' s.WriteLine("")
        's.WriteLine("^XA")
        's.WriteLine("^MMT")
        's.WriteLine("^PW799")
        's.WriteLine("^LL0400")
        's.WriteLine("^LS0")
        's.WriteLine("^FT788,275^A0I,21,24^FH\^FDPacked by : MYG,CORP.OFFICE^FS")
        's.WriteLine("^FT788,250^A0I,16,28^FH\^FDCALICUT,KERALA,INDIA^FS")
        's.WriteLine("^FT788,222^A0I,19,24^FH\^FDpacked date : " & year1 & " QTY : " & qty & " nos^FS")
        's.WriteLine("^FT788,193^A0I,19,24^FH\^FDCommodity : " & desc & "^FS")
        's.WriteLine("^FT788,165^A0I,20,26^FH\^FDCustomer Care No. 18001232006^FS")
        's.WriteLine("^FT789,145^A0I,21,28^FH\^FDEmail : Info@myg.in^FS")
        's.WriteLine("^BY2,3,49^FT787,89^BCI,,N,N")
        's.WriteLine("^FD>:" & itemcode & "^FS")                             'left
        's.WriteLine("^FT788,67^A0I,21,28^FH\^FD" & itemcode & "^FS")         'ItemNumber 1
        's.WriteLine("^FT788,32^A0I,25,33^FH\^FDMRP Rs" & price & "^FS")
        's.WriteLine("^FT572,32^A0I,14,28^FH\^FDinc all taxes^FS")
        's.WriteLine("^FT388,275^A0I,21,24^FH\^FDPacked by : MYG,CORP.OFFICE^FS")
        's.WriteLine("^FT388,250^A0I,16,28^FH\^FDCALICUT,KERALA,INDIA^FS")
        's.WriteLine("^FT388,222^A0I,19,24^FH\^FDpacked date : " & year1 & " QTY : " & qty & " nos^FS")
        's.WriteLine("^FT388,193^A0I,19,24^FH\^FDCommodity : " & desc & "^FS")
        's.WriteLine("^FT388,165^A0I,20,26^FH\^FDCustomer Care No. 18001232006^FS")
        's.WriteLine("^FT389,145^A0I,21,28^FH\^FDEmail : Info@myg.in^FS")
        's.WriteLine("^BY2,3,49^FT387,89^BCI,,N,N")
        's.WriteLine("^FD>:" & itemcode1 & "^FS")            'Right
        's.WriteLine("^FT388,67^A0I,21,28^FH\^FD" & itemcode1 & "^FS")        'ItemNumber 2
        's.WriteLine("^FT388,32^A0I,25,33^FH\^FDMRP Rs" & price & "^FS")
        's.WriteLine("^FT172,32^A0I,14,28^FH\^FDinc all taxes^FS")
        's.WriteLine("^BY2,3,49^FT787,320^BCI,,N,N")
        's.WriteLine("^FD>:" & barcode & "^FS")              'Serial number 1    left 
        's.WriteLine("^FT788,300^A0I,21,28^FH\^FD" & barcode & "^FS")
        's.WriteLine("^BY2,3,49^FT387,320^BCI,,N,N")
        's.WriteLine("^FD>:" & barcode1 & "^FS")             'Serial number 2   Right
        's.WriteLine("^FT388,300^A0I,21,28^FH\^FD" & barcode1 & "^FS")
        's.WriteLine("^PQ1,0,1,Y^XZ")
        ' RowCount += 1

        For i As Integer = 1 To objMatrix.RowCount

            If objMatrix.Columns.Item("3").Cells.Item(i).Specific.checked = True Then
                itemcode = objMatrix.Columns.Item("8").Cells.Item(i).Specific.string
                barcode = objMatrix.Columns.Item("1").Cells.Item(i).Specific.string
                price = objMatrix.Columns.Item("2").Cells.Item(i).Specific.string
                desc = objMatrix.Columns.Item("4").Cells.Item(i).Specific.string
                qty = objMatrix.Columns.Item("5").Cells.Item(i).Specific.string
                vendor = objMatrix.Columns.Item("5A").Cells.Item(i).Specific.string
                year1 = objMatrix.Columns.Item("5B").Cells.Item(i).Specific.string

                'barcode = Left(barcode, 10)
                qty = CInt(qty)
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
        Dim objCombo As SAPbouiCOM.ComboBox
        objCombo = objForm.Items.Item("17").Specific
       
        Dim j As Integer = 1

        While j <= objMatrix.VisualRowCount
            If objMatrix.Columns.Item("3").Cells.Item(j).Specific.checked = True Then
                updateStatus(FormUID, j, j + 1)
                objMatrix.DeleteRow(j)
            End If

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
        'Foldername = "D:" + "\" + rName + "\" + GRPODocNum + "\MyG"
        Foldername = "E:" + "\" + rName + "\" + GRPODocNum + "\MyG"
        If Directory.Exists(Foldername) Then
        Else
            Directory.CreateDirectory(Foldername)
        End If


        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("6").Specific
        'While objMatrix.RowCount > 0
        intloop = 1

        ' If objMatrix.Columns.Item("3").Cells.Item(intloop).Specific.checked = True Then

        'year1 = year1.Replace("/", "-")
        ' year2 = convert_date(year1)
        'year1 = Right(year1, 5)
        'year3 = Date.ParseExact(year1, "dd-mm-yyyy",
        'System.Globalization.DateTimeFormatInfo.InvariantInfo)
        '                year2 = Format(year3, "MMM-yyyy")
        ' year1 = Format(objMatrix.Columns.Item("5B").Cells.Item(intloop).Specific, "mmm yyyy")
        ' year1 = Right(year1, 4)
        'year3 = Date.ParseExact(year1, "dd-MMM-yyyy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
        'year2 = Format(year1, "MMM-yyyy")

        'price = convert_value(price.Replace(" ", ""))
        'convert_val(price)
        ' price = CStr(price)
        'price = convert_val(price)
        'price2 = CDbl(price)

        'qty = CStr(objMatrix.Columns.Item("7").Cells.Item(intloop).Specific.string) + "-" + CStr(CInt(qty))
        '--------------File Creation--------------------
        Dim Filetype As String = "BIG"

        Filetype = Filetype.Replace(" ", "")
        'Dim Filename As String = Trim(objMatrix.Columns.Item("6").Cells.Item(intloop).Specific.string) + "_" + Filetype + "_" + (desc.Replace(" ", "")).Replace("/", "").Replace(".", "").Replace(":", "") + "_" + CStr(CInt(objMatrix.Columns.Item("7").Cells.Item(intloop).Specific.string) + 1) + "_" + CStr(count) + ".prn"
        Dim Filename As String = Trim(objMatrix.Columns.Item("8").Cells.Item(1).Specific.string) + "_" + Filetype + "_" + System.DateTime.Now.ToString("yyMMddHHmmss") + "_" + CStr(count) + ".prn"
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

            If objMatrix.Columns.Item("3").Cells.Item(i).Specific.checked = True Then
                itemcode = objMatrix.Columns.Item("8").Cells.Item(i).Specific.string
                barcode = objMatrix.Columns.Item("1").Cells.Item(i).Specific.string
                price = objMatrix.Columns.Item("2").Cells.Item(i).Specific.string
                desc = objMatrix.Columns.Item("4").Cells.Item(i).Specific.string
                qty = objMatrix.Columns.Item("5").Cells.Item(i).Specific.string
                vendor = objMatrix.Columns.Item("5A").Cells.Item(i).Specific.string
                year1 = objMatrix.Columns.Item("5B").Cells.Item(i).Specific.string

                'barcode = Left(barcode, 10)
                qty = CInt(qty)
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
        Dim objCombo As SAPbouiCOM.ComboBox
        objCombo = objForm.Items.Item("17").Specific

        Dim j As Integer = 1

        While j <= objMatrix.RowCount
            updateStatus(FormUID, j, j + 1)
            objMatrix.DeleteRow(j)
        End While
        'End While
        ' objAddOn.objApplication.MessageBox(CStr(count) + " Barcode(s) printed")
        'objAddOn.objApplication.MessageBox("Barcode file(s) generated")
        objAddOn.objApplication.StatusBar.SetText("Barcode file(s) generated", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    End Sub
    Private Sub CreateFile2(ByVal FormUID As String, ByVal Prefix As String)
        ' big
        Dim intloop As Integer
        Dim qty As String
        Dim desc As String
        Dim price As String
        ' Dim price2 As Decimal
        Dim barcode As String
        Dim vendor As String
        Dim year1 As String
        Dim year2 As String
        ' Dim year3 As Date
        Dim colorsize As String = ""
        Dim count As Integer = 0

        Dim path As String = ""
        Dim fs As FileStream
        Dim rName As String = ""
        rName = SystemInformation.UserName
        Dim Foldername As String


        Foldername = "D:" + "\" + rName + "\" + GRPODocNum + "\MyG"
        If Directory.Exists(Foldername) Then
        Else
            Directory.CreateDirectory(Foldername)
        End If


        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("6").Specific
        While objMatrix.RowCount > 0
            intloop = 1
            If objMatrix.Columns.Item("3").Cells.Item(intloop).Specific.checked = True Then
                barcode = objMatrix.Columns.Item("1").Cells.Item(intloop).Specific.string
                price = objMatrix.Columns.Item("2").Cells.Item(intloop).Specific.string
                desc = Trim(objMatrix.Columns.Item("4").Cells.Item(intloop).Specific.string)
                qty = objMatrix.Columns.Item("5").Cells.Item(intloop).Specific.string
                vendor = objMatrix.Columns.Item("5A").Cells.Item(intloop).Specific.string
                year1 = objMatrix.Columns.Item("5B").Cells.Item(intloop).Specific.string
                'year1 = year1.Replace("/", "-")
                ' year2 = convert_date(year1)
                'year1 = Right(year1, 5)
                'year3 = Date.ParseExact(year1, "dd-mm-yyyy",
                'System.Globalization.DateTimeFormatInfo.InvariantInfo)
                '                year2 = Format(year3, "MMM-yyyy")
                ' year1 = Format(objMatrix.Columns.Item("5B").Cells.Item(intloop).Specific, "mmm yyyy")
                ' year1 = Right(year1, 4)
                'year3 = Date.ParseExact(year1, "dd-MMM-yyyy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                'year2 = Format(year1, "MMM-yyyy")
                barcode = Left(barcode, 10)
                qty = CInt(qty)


                price = price.Replace(",", "")
                price = Prefix + Mid(price, 1, (price.IndexOf(".") + 3))
                price = price.Replace("s", "")
                price = price.Replace("R", "")
                price = price.Replace(" ", "")
                'price = convert_value(price.Replace(" ", ""))
                'convert_val(price)
                ' price = CStr(price)
                'price = convert_val(price)
                'price2 = CDbl(price)

                'qty = CStr(objMatrix.Columns.Item("7").Cells.Item(intloop).Specific.string) + "-" + CStr(CInt(qty))
                '--------------File Creation--------------------
                Dim Filetype As String = "BIG"

                Filetype = Filetype.Replace(" ", "")
                Dim Filename As String = Trim(objMatrix.Columns.Item("6").Cells.Item(intloop).Specific.string) + "_" + Filetype + "_" + (desc.Replace(" ", "")).Replace("/", "").Replace(".", "").Replace(":", "") + "_" + CStr(CInt(objMatrix.Columns.Item("7").Cells.Item(intloop).Specific.string) + 1) + "_" + CStr(count) + ".prn"

                'Create the file.
                ' path = Application.StartupPath + "\" + rName + "\mylabel.prn"
                path = Foldername + "\" + Filename
                ' Delete the file if it exists.
                If File.Exists(path) Then
                    File.Delete(path)
                End If
                fs = New FileStream(path, FileMode.Create, FileAccess.Write)

                '-----------------------------------------------------------
                Dim s As New StreamWriter(fs)
                s.WriteLine("")
                s.WriteLine("^XA")
                s.WriteLine("^MMT")
                s.WriteLine("^PW799")
                s.WriteLine("^LL0400")
                s.WriteLine("^LS0")
                s.WriteLine("^FT788,297^A0I,21,24^FH\^FDPacked by : MYG,CORP.OFFICE^FS")
                s.WriteLine("^FT788,270^A0I,16,28^FH\^FDCALICUT,KERALA,INDIA^FS")
                s.WriteLine("^FT788,245^A0I,19,24^FH\^FDpacked date : " & year1 & " QTY : " & qty & " nos^FS")
                s.WriteLine("^FT788,217^A0I,19,24^FH\^FDCommodity : " & desc & "^FS")
                s.WriteLine("^FT788,188^A0I,20,26^FH\^FDCustomer Care No. 18001232006^FS")
                s.WriteLine("^FT789,160^A0I,21,28^FH\^FDEmail : Info@myg.in^FS")
                s.WriteLine("^BY2,3,61^FT787,89^BCI,,N,N")
                s.WriteLine("^FD>:" & barcode & "^FS")
                s.WriteLine("^FT788,67^A0I,21,28^FH\^FD" & barcode & "^FS")
                s.WriteLine("^FT788,32^A0I,25,33^FH\^FDMRP Rs" & price & "^FS")
                s.WriteLine("^FT530,32^A0I,14,28^FH\^FDinc all taxes^FS")
                s.WriteLine("^FT388,297^A0I,21,24^FH\^FDPacked by : MYG,CORP.OFFICE^FS")
                s.WriteLine("^FT388,270^A0I,16,28^FH\^FDCALICUT,KERALA,INDIA^FS")
                s.WriteLine("^FT388,245^A0I,19,24^FH\^FDpacked date : " & year1 & " QTY : " & qty & " nos^FS")
                s.WriteLine("^FT388,217^A0I,19,24^FH\^FDCommodity : " & desc & "^FS")
                s.WriteLine("^FT388,188^A0I,20,26^FH\^FDCustomer Care No. 18001232006^FS")
                s.WriteLine("^FT388,160^A0I,21,28^FH\^FDEmail : Info@myg.in^FS")
                s.WriteLine("^BY2,3,61^FT387,89^BCI,,N,N")
                s.WriteLine("^FD>:" & barcode & "^FS")
                s.WriteLine("^FT388,67^A0I,21,28^FH\^FD" & barcode & "^FS")
                s.WriteLine("^FT388,32^A0I,25,33^FH\^FDMRP Rs" & price & "^FS")
                s.WriteLine("^FT130,32^A0I,14,28^FH\^FDinc all taxes^FS")
                s.WriteLine("^PQ1,0,1,Y^XZ")

                s.Close()
                fs.Close()


                Dim systemname As String = ""
                ' Dim Str_FromPath As String = "C:\" & Environment.UserName & "\" & txtBarcode.Text & "_" & txtbatchtag.Text & ".txt"
                Dim str_fromPath As String = path
                systemname = Environment.MachineName
                Dim ToPath As String = "\\" & systemname & "\\" & getDefaultPrinter()

                File.Copy(str_fromPath, ToPath)

                Dim objCombo As SAPbouiCOM.ComboBox
                objCombo = objForm.Items.Item("17").Specific

                updateStatus(FormUID, intloop, count + 1)
                count += 1
            End If
            objMatrix.DeleteRow(intloop)
        End While
        ' objAddOn.objApplication.MessageBox(CStr(count) + " Barcode(s) printed")
        objAddOn.objApplication.MessageBox("Barcode file(s) genereted")
    End Sub
    'Function retWord(ByVal Num As Decimal) As String
    '    'This two dimensional array store the primary word convertion of number.
    '    retWord = ""
    '    Dim ArrWordList(,) As Object = {{0, "M"}, {1, "B"}, {2, "C"}, {3, "A"}, {4, "S"}, _
    '                                    {5, "R"}, {6, "D"}, {7, "E"}, {8, "U"}, {9, "N"}}

    '    Dim i As Integer
    '    For i = 0 To UBound(ArrWordList)
    '        If Num = ArrWordList(i, 0) Then
    '            retWord = ArrWordList(i, 1)
    '            Exit For
    '        End If
    '    Next
    '    Return retWord
    'End Function

    'Function convert_val(ByVal str As String) As String
    '    Dim a As Array
    '    Dim TEMP As String = ""
    '    Dim val As String = ""
    '    a = str.ToCharArray()
    '    For i = 0 To a.Length
    '        If a(i) = "0" Then
    '            TEMP = "M"
    '        ElseIf a(i) = "1" Then
    '            TEMP = "B"
    '        ElseIf a(i) = "2" Then
    '            TEMP = "C"
    '        ElseIf a(i) = "3" Then
    '            TEMP = "A"
    '        ElseIf a(i) = "4" Then
    '            TEMP = "S"
    '        ElseIf a(i) = "5" Then
    '            TEMP = "R"
    '        ElseIf a(i) = "6" Then
    '            TEMP = "D"
    '        ElseIf a(i) = "7" Then
    '            TEMP = "E"
    '        ElseIf a(i) = "8" Then
    '            TEMP = "U"
    '        ElseIf a(i) = "9" Then
    '            TEMP = "N"
    '        ElseIf a(i) = "." Then
    '            TEMP = "."
    '        ElseIf a(i) = "R" Then
    '            TEMP = "R"
    '        ElseIf a(i) = "s" Then
    '            TEMP = "s"
    '        ElseIf a(i) = " " Then
    '            TEMP = " "
    '        End If
    '        val = val + TEMP
    '    Next

    '    Return val
    'End Function


    'Private Sub CreateFile_3G(ByVal FormUID As String, ByVal Prefix As String)
    '    'sm
    '    Dim intloop As Integer
    '    Dim qty As String
    '    Dim desc As String
    '    Dim price As String
    '    Dim barcode As String
    '    Dim vendor As String
    '    Dim year1 As String
    '    Dim colorsize As String = ""
    '    Dim count As Integer = 0

    '    Dim path As String = ""
    '    Dim fs As FileStream
    '    Dim rName As String = ""
    '    rName = SystemInformation.UserName
    '    Dim Foldername As String
    '    Foldername = Application.StartupPath + "\" + rName + "\" + GRPODocNum + "\3G"
    '    If Directory.Exists(Foldername) Then
    '    Else
    '        Directory.CreateDirectory(Foldername)
    '    End If


    '    objForm = objAddOn.objApplication.Forms.Item(FormUID)
    '    objMatrix = objForm.Items.Item("6").Specific
    '    While objMatrix.RowCount > 0
    '        intloop = 1
    '        If objMatrix.Columns.Item("3").Cells.Item(intloop).Specific.checked = True Then
    '            barcode = objMatrix.Columns.Item("1").Cells.Item(intloop).Specific.string
    '            price = objMatrix.Columns.Item("2").Cells.Item(intloop).Specific.string
    '            desc = objMatrix.Columns.Item("4").Cells.Item(intloop).Specific.string
    '            'qty = objMatrix.Columns.Item("5").Cells.Item(intloop).Specific.string
    '            'vendor = objMatrix.Columns.Item("5A").Cells.Item(intloop).Specific.string
    '            'year1 = objMatrix.Columns.Item("5B").Cells.Item(intloop).Specific.string
    '            'year1 = Right(year1, 2)
    '            'barcode = Left(barcode, 10)
    '            'price = price.Replace(",", "")
    '            'price = Prefix + Mid(price, 1, (price.IndexOf(".") + 3))
    '            'qty = CStr(objMatrix.Columns.Item("7").Cells.Item(intloop).Specific.string) + "-" + CStr(CInt(qty))
    '            '--------------File Creation--------------------
    '            Dim Filetype As String = "SMALL"

    '            Filetype = Filetype.Replace(" ", "")
    '            Dim Filename As String = Trim(objMatrix.Columns.Item("6").Cells.Item(intloop).Specific.string) + "_" + Filetype + "_" + (desc.Replace(" ", "")).Replace("/", "") + "_" + CStr(CInt(objMatrix.Columns.Item("7").Cells.Item(intloop).Specific.string) + 1) + "_" + CStr(count) + ".prn"

    '            'Create the file.
    '            ' path = Application.StartupPath + "\" + rName + "\mylabel.prn"
    '            path = Foldername + "\" + Filename
    '            ' Delete the file if it exists.
    '            If File.Exists(path) Then
    '                File.Delete(path)
    '            End If
    '            fs = New FileStream(path, FileMode.Create, FileAccess.Write)

    '            '-----------------------------------------------------------
    '            Dim s As New StreamWriter(fs)
    '            s.WriteLine("")
    '            s.WriteLine("^XA")
    '            s.WriteLine("^MMT")
    '            s.WriteLine("^PW607")
    '            s.WriteLine("^LL0200")
    '            s.WriteLine("^LS0")
    '            s.WriteLine("^BY2,3,68^FT575,97^BCI,,N,N")
    '            s.WriteLine("^FD>:" & barcode & "^FS")
    '            s.WriteLine("^FT572,81^A0I,15,28^FH\^FD" & barcode & "^FS")
    '            s.WriteLine("^FT576,58^A0I,21,21^FH\^FD" & desc & "^FS")
    '            s.WriteLine("^FT574,30^A0I,19,24^FH\^FD" & price & "^FS")
    '            s.WriteLine("^BY2,3,68^FT271,97^BCI,,N,N")
    '            s.WriteLine("^FD>:" & barcode & "^FS")
    '            s.WriteLine("^FT268,81^A0I,15,28^FH\^FD" & barcode & "^FS")
    '            s.WriteLine("^FT272,58^A0I,21,21^FH\^FD" & desc & "^FS")
    '            s.WriteLine("^FT270,30^A0I,19,24^FH\^FD" & price & "^FS")
    '            s.WriteLine("^PQ1,0,1,Y^XZ")

    '            s.Close()
    '            fs.Close()

    '            Dim objCombo As SAPbouiCOM.ComboBox
    '            objCombo = objForm.Items.Item("17").Specific

    '            updateStatus(FormUID, intloop, count + 1)
    '            count += 1
    '        End If
    '        objMatrix.DeleteRow(intloop)
    '    End While
    '    ' objAddOn.objApplication.MessageBox(CStr(count) + " Barcode(s) printed")
    '    objAddOn.objApplication.MessageBox("Barcode file(s) genereted")
    'End Sub
    
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
    Private Sub updateStatus(ByVal FormUID As String, ByVal rowid As Integer, ByVal count1 As Integer)
        Dim docentry As Integer
        Dim lineid As Integer
        Dim barcode As String
        Dim bqty As Integer = 0

        Try

            objForm = objAddOn.objApplication.Forms.Item(FormUID)
            objMatrix = objForm.Items.Item("6").Specific
            docentry = CInt(objMatrix.Columns.Item("6").Cells.Item(rowid).Specific.string)
            lineid = CInt(objMatrix.Columns.Item("7").Cells.Item(rowid).Specific.string)
            barcode = objMatrix.Columns.Item("1").Cells.Item(rowid).Specific.string
            objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If objAddOn.HANA Then
                strSQL = "UPDATE OSRN SET ""U_printed"" = 'Y' WHERE ""DistNumber"" = '" & barcode & "';" ' AND ""BaseEntry"" =" & docentry & " AND ""BaseLinNum""= " & lineid & " ;"
            Else
                strSQL = "update OSRN set U_printed='Y' where distnumber='" & barcode & "'" ' AND BaseEntry =" & docentry & " AND BaseLinNum= " & lineid
            End If

            objRS.DoQuery(strSQL)
            objRS = Nothing
            objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If objAddOn.HANA Then
                strSQL = "UPDATE OBTN SET ""U_printed"" = 'Y' WHERE ""DistNumber"" = '" & barcode & "';" ' AND ""BaseEntry"" =" & docentry & " AND ""BaseLinNum""= " & lineid & " ;"
            Else
                strSQL = "update OBTN set U_printed='Y' where distnumber='" & barcode & "' " 'AND BaseEntry =" & docentry & " AND BaseLinNum= " & lineid
            End If

            objRS.DoQuery(strSQL)
            objRS = Nothing
            objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If objAddOn.HANA Then
                strSQL = "select IFNULL(COUNT(""U_printed""),0) as ""bqty"" from OSRI where ""BaseEntry"" = " & docentry & " and ""BaseLinNum"" = " & lineid & " and ""U_printed""='Y';"
            Else
                strSQL = "select isnull(count(U_printed),0) as bqty from OSRI where baseentry=" & docentry & " and baselinnum=" & lineid & " and U_printed='Y'"
            End If

            objRS.DoQuery(strSQL)
            If Not objRS.EoF Then
                bqty = objRS.Fields.Item("bqty").Value
            End If
            objRS = Nothing
            If bqty = 0 Then
                objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If objAddOn.HANA Then
                    strSQL = "select IFNULL(COUNT(""U_printed""),0) as ""bqty"" from OIBT where ""BaseEntry"" = " & docentry & " and ""BaseLinNum""=" & lineid & " and ""U_printed""='Y';"
                Else
                    strSQL = "select isnull(count(U_printed),0) as bqty from OIBT where baseentry=" & docentry & " and baselinnum=" & lineid & " and U_printed='Y'"
                End If
                objRS.DoQuery(strSQL)
                If Not objRS.EoF Then
                    bqty = count1
                End If
                objRS = Nothing
            End If
            objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If objAddOn.HANA Then
                strSQL = "update PDN1 set ""U_bqty""= ""U_bqty""+" & bqty & " where ""DocEntry"" =" & docentry & " and ""LineNum"" = " & lineid & " ;"
            Else
                strSQL = "update PDN1 set U_bqty=" & bqty & " where docentry=" & docentry & " and linenum=" & lineid
            End If

            objRS.DoQuery(strSQL)
            objRS = Nothing
            objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            ' strSQL = "Select U_bqty, quantity from pdn1 where docentry=" & docentry & " and linenum=" & lineid

            If objAddOn.HANA Then
                strSQL = "select sum(""U_bqty"") as ""U_bqty"", sum(""Quantity"") as quantity from PDN1 where ""DocEntry"" = " & docentry & " ;"
            Else
                strSQL = "Select sum(U_bqty) as U_bqty, sum(quantity) as quantity from pdn1 where docentry=" & docentry
            End If

            objRS.DoQuery(strSQL)
            'Dim bqty As Long = 0
            Dim qty As Long = 0
            bqty = objRS.Fields.Item("U_bqty").Value
            qty = objRS.Fields.Item("quantity").Value
            objRS = Nothing
            objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objRS1 = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If bqty = qty Then
                If objAddOn.HANA Then
                    strSQL = "update OPDN set ""U_bstatus""='Y' where ""DocEntry"" =" & docentry & ";"
                Else
                    strSQL = "update opdn set U_bstatus='Y' where docentry=" & docentry
                End If
            Else
                If objAddOn.HANA Then
                    strSQL = "UPDATE OSRN SET ""U_printed"" = '' WHERE ""DistNumber"" = '" & barcode & "';"
                    strSQL1 = "UPDATE OBTN SET ""U_printed"" = '' WHERE ""DistNumber"" = '" & barcode & "';"
                Else
                    strSQL = "update opdn set U_bstatus='Y' where docentry=" & docentry
                End If
            End If
            If strSQL1 <> "" Then objRS1.DoQuery(strSQL1)
            If strSQL <> "" Then objRS.DoQuery(strSQL)
            objRS = Nothing
            objRS1 = Nothing
        Catch ex As Exception
            objAddOn.objApplication.MessageBox(ex.Message)
        End Try
    End Sub
    Public Sub menuevent(ByVal pVal As SAPbouiCOM.MenuEvent, ByVal BubbleEvent As Boolean)
        If pVal.MenuUID = "1281" Then 'find
        ElseIf pVal.MenuUID = "1282" Then 'add
        End If
    End Sub
    Private Function validate(ByVal FormUID As String) As Boolean
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("6").Specific
        If objMatrix.RowCount = 0 Then
            objAddOn.objApplication.SetStatusBarMessage("Atleast one barcode should be selected to print", SAPbouiCOM.BoMessageTime.bmt_Short)
            Return False
        End If
        Return True
    End Function
    Function imageToByteArray(ByVal imageIn As System.Drawing.Image) As Byte()
        Dim ms As IO.MemoryStream = New IO.MemoryStream()
        imageIn.Save(ms, System.Drawing.Imaging.ImageFormat.Png)
        Return ms.ToArray()
    End Function


    Function convert_value(ByVal str As String) As String
        Dim objRS1 As SAPbobsCOM.Recordset = Nothing
        Dim DT As New DataTable
        Dim strsql As String = ""
        Dim a As Array
        Dim i As Integer
        Dim str2, str3, str4 As String
        Dim value As String = ""
        a = str.ToCharArray()
        objRS1 = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If objAddOn.HANA Then
            strsql = "select ""Code"",""Name"" from ""@PRICECODE"""
        Else
            strsql = "Select Code,Name from @PRICECODE"
        End If
        objRS1.DoQuery(strsql)
        DT.Columns.Add("code")
        DT.Columns.Add("Name")
        While Not objRS1.EoF
            
            Dim dr = DT.NewRow()
            str4 = objRS1.Fields.Item(0).Value
            dr.Item(0) = str4
            str2 = objRS1.Fields.Item(1).Value
            dr.Item(1) = str2
            DT.Rows.Add(dr)
            objRS1.MoveNext()
        End While
        str3 = ""
        For i = 0 To a.Length - 1
            For j = 0 To DT.Rows.Count - 1
                If DT.Rows(j).Item(0) = a(i).ToString Then
                    value = DT.Rows(j).Item(1).ToString
                    str3 = str3 + value
                End If
            Next
        Next

        Return str3
    End Function

    'Function convert_date(ByVal date1 As String) As String
    '    Dim result As String = ""
    '    Dim val As String = ""
    '    Dim temp As String = ""
    '    Dim temp2 As String = ""
    '    temp = Right(date1, 2)
    '    temp2 = date1.Substring(3, 2)
    '    If temp2 = "01" Then
    '        val = "Jan"
    '    ElseIf temp2 = "02" Then
    '        val = "Feb"
    '    ElseIf temp2 = "03" Then
    '        val = "Mar"
    '    ElseIf temp2 = "04" Then
    '        val = "Apr"
    '    ElseIf temp2 = "05" Then
    '        val = "May"
    '    ElseIf temp2 = "06" Then
    '        val = "Jun"
    '    ElseIf temp2 = "07" Then
    '        val = "Jul"
    '    ElseIf temp2 = "08" Then
    '        val = "Aug"
    '    ElseIf temp2 = "09" Then
    '        val = "Sep"
    '    ElseIf temp2 = "10" Then
    '        val = "Oct"
    '    ElseIf temp2 = "11" Then
    '        val = "Nov"
    '    ElseIf temp2 = "12" Then
    '        val = "Dec"
    '    End If
    '    result = val + "-" + "20" + temp
    '    Return result
    'End Function
End Class
