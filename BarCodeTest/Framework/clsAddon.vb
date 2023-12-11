Imports System.IO
Imports SAPbouiCOM.Framework
Public Class clsAddOn
    Public WithEvents objApplication As SAPbouiCOM.Application
    Public objCompany As SAPbobsCOM.Company
    Dim oProgBarx As SAPbouiCOM.ProgressBar
    Public objGenFunc As Mukesh.SBOLib.GeneralFunctions
    Public objUIXml As Mukesh.SBOLib.UIXML
    Public ZB_row As Integer = 0
    Public Warranty As Boolean = True
    Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
    Dim oUserTablesMD As SAPbobsCOM.UserTablesMD
    Dim ret As Long
    Public Flagnew As Boolean = False
    'Dim str As String
    Dim objForm As SAPbouiCOM.Form
    Dim objUDFEngine As Mukesh.SBOLib.UDFEngine
    Dim MenuCount As Integer = 0
    Public storeserial As String
    Public CancelCode As Integer
    Public HANA As Boolean = True

    '---
    Dim objBarcodePrint As ClsPrintBarcode
    Dim objReprint As ClsbarcodeReprint
    Public tempTotal As Long = 0
    Public SerialArray(10, 5) As String
    Public ItemRowCount As Integer

    Dim Filters As SAPbouiCOM.EventFilters
    Public HWKEY() As String = New String() {"N0498304534", "L1653539483", "S1020319487", "W1831256098", "R1574408489", "H0816777137", "A1836445156", "A0061802481", "M0394249985", "V0913316776", "F0123559701", "L1552968038", "M0090876837", "H0922924113", "Y1334940735", "B0241390111"}
    Private Sub CheckLicense()

    End Sub
    Function isValidLicense() As Boolean
        Try
            objApplication.Menus.Item("257").Activate()
            Dim CrrHWKEY As String = objApplication.Forms.ActiveForm.Items.Item("79").Specific.Value.ToString.Trim
            objApplication.Forms.ActiveForm.Close()

            For i As Integer = 0 To HWKEY.Length - 1
                If HWKEY(i).Trim = CrrHWKEY.Trim Then
                    Return True
                End If
            Next
            MsgBox("Add-on installation failed due to license mismatch", MsgBoxStyle.OkOnly, "License Management")
            Return False
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return True
    End Function
    Public Sub Intialize(ByVal args() As String)
        Try
            Dim oapplication As Application
            If (args.Length < 1) Then oapplication = New Application Else oapplication = New Application(args(0))
            objapplication = Application.SBO_Application
            If isValidLicense() Then
                objapplication.StatusBar.SetText("Establishing Company Connection Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                objcompany = Application.SBO_Application.Company.GetDICompany()
                Try
                    createObjects()
                    createTables()
                    createUDOs()
                    loadMenu()
                Catch ex As Exception
                    objAddOn.objApplication.MessageBox(ex.ToString)
                    End
                End Try
                objApplication.StatusBar.SetText("Addon Connected Successfully..!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                oapplication.Run()
            Else
                objapplication.StatusBar.SetText("Addon Disconnected due to license mismatch..!!!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            'System.Windows.Forms.Application.Run()
        Catch ex As Exception
            objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub

    Public Sub Intialize()
        'If UCase(My.Settings.HANAAddon).Contains("TRUE") Then
        '    HANA = True
        'Else
        '    HANA = False
        'End If
        ' HANA = False
        Dim objSBOConnector As New Mukesh.SBOLib.SBOConnector
        ' Dim constr As String = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056"
        objApplication = objSBOConnector.GetApplication(System.Environment.GetCommandLineArgs.GetValue(1))

        'objApplication = objSBOConnector.GetApplication(constr)
        objCompany = objSBOConnector.GetCompany(objApplication)
        Try
            createTables()
            createUDOs()
            createObjects()
            '  SetFilters()
            loadMenu()
        Catch ex As Exception
            objAddOn.objApplication.MessageBox(ex.ToString)
            End
        End Try
        If isValidLicense() Then
            objApplication.SetStatusBarMessage("Barcode Add-On connected  successfully!", SAPbouiCOM.BoMessageTime.bmt_Short, False)

        Else
            objApplication.SetStatusBarMessage("Failed To Connect, Please Check The License Configuration", SAPbouiCOM.BoMessageTime.bmt_Short, False)
            objCompany.Disconnect()
            objApplication = Nothing
            objCompany = Nothing
            End
        End If


    End Sub
    Private Sub createUDOs()
        'Dim ct1(2) As String
        'ct1(0) = "MI_INVTR1"
        'ct1(1) = "MI_INVTR2"
        'createUDOC("MI_INVTRH", "MI_INVTR", "StockTransfer", ct1, SAPbobsCOM.BoUDOObjType.boud_Document, False, True)
    End Sub
    Private Sub createObjects()
        'Library Object Initilisation
        objGenFunc = New Mukesh.SBOLib.GeneralFunctions(objCompany)
        objUIXml = New Mukesh.SBOLib.UIXML(objApplication)
        'Business Object Initialisation
        objBarcodePrint = New ClsPrintBarcode
        objReprint = New ClsbarcodeReprint

    End Sub
    Private Sub objApplication_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles objApplication.ItemEvent
        Try
            Select Case pVal.FormTypeEx
                Case ClsPrintBarcode.formtype
                    objBarcodePrint.ItemEvent(FormUID, pVal, BubbleEvent)
                Case ClsbarcodeReprint.formtype
                    objReprint.ItemEvent(FormUID, pVal, BubbleEvent)

                Case "24"
                    'Try
                    '    objAddOn.objApplication.Forms.GetForm("41", 1)

                    '    objStringBatch.ItemEvent(FormUID, pVal, BubbleEvent)
                    'Catch ex As Exception
                    '    objStringSerialDate.ItemEvent(FormUID, pVal, BubbleEvent)
                    'End Try
                    'If objAddOn.objApplication.Forms.Item(FormUID).Title.Contains("Batches") Then
                    '    objStringBatch.ItemEvent(FormUID, pVal, BubbleEvent)
                    'Else
                    '    objStringSerialDate.ItemEvent(FormUID, pVal, BubbleEvent)
                    'End If

                    'Case "65051"
                    '    objAutoCreation.ItemEvent(FormUID, pVal, BubbleEvent)
                    'Case "143"
                    '    objGRPO.ItemEvent(FormUID, pVal, BubbleEvent)
                    'Case "721"
                    '    objGoodsReceipt.ItemEvent(FormUID, pVal, BubbleEvent)

                    'Case "41"
                    '    objbatchmrp.ItemEvent(FormUID, pVal, BubbleEvent)
                    'Case "65053"
                    '    objAutoBatchCreation.ItemEvent(FormUID, pVal, BubbleEvent)
            End Select
        Catch ex As Exception
            objAddOn.objApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short)
        End Try
    End Sub
    Public Shared Function ErrorHandler(ByVal p_ex As Exception, ByVal objApplication As SAPbouiCOM.Application)
        Dim sMsg As String = Nothing
        If p_ex.Message = "Form - already exists [66000-11]" Then
            Return True
            Exit Function  'ignore error
        End If
        Return False
    End Function
    Private Sub objApplication_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles objApplication.MenuEvent

        If pVal.BeforeAction Then
            If pVal.MenuUID = btnDelRow Then
                If objAddOn.objApplication.Forms.ActiveForm.UniqueID.Contains("invtransudo") Then
                    ' objInvTransferUDO.MenuEvent(pVal, BubbleEvent)
                End If
            End If
        Else

            Try
                Select Case pVal.MenuUID
                    Case ClsPrintBarcode.formtype
                        objBarcodePrint.LoadScreen()
                    Case ClsbarcodeReprint.formtype
                        objReprint.LoadScreen()

                End Select
                If pVal.MenuUID = btnFirst Or pVal.MenuUID = btnLast Or pVal.MenuUID = btnNext Or pVal.MenuUID = btnPrevious Or pVal.MenuUID = btnAdd Or pVal.MenuUID = btnDelRow Then
                    If objAddOn.objApplication.Forms.ActiveForm.UniqueID.Contains("invtransudo") Then
                        ' objInvTransferUDO.MenuEvent(pVal, BubbleEvent)
                    End If
                End If
                'If pVal.MenuUID = btnAdd Then
                '    If objAddOn.objApplication.Forms.ActiveForm.UniqueID.Contains("invtransudo") Then
                '        objInvTransferUDO.MenuEvent(pVal, BubbleEvent)
                '    End If
                'End If
            Catch ex As Exception
                ' objAddOn.objApplication.MessageBox(ex.ToString)
            End Try
        End If

    End Sub
    Private Sub loadMenu()
        Dim menuitem
        If objApplication.Menus.Item("43520").SubMenus.Exists("Barcode") Then Return
        MenuCount = objApplication.Menus.Item("43520").SubMenus.Count
        menuitem = CreateMenu("", MenuCount, "Barcode", SAPbouiCOM.BoMenuType.mt_POPUP, "Barcode", objApplication.Menus.Item("43520"))
        CreateMenu("", 1, "Print", SAPbouiCOM.BoMenuType.mt_STRING, "BarPrint", menuitem)
        CreateMenu("", 2, "Reprint", SAPbouiCOM.BoMenuType.mt_STRING, "BarRprt", menuitem)
        ' CreateMenu("", 3, "Stock Transfer", SAPbouiCOM.BoMenuType.mt_STRING, "invtransudo", menuitem)
        '  CreateMenu("", 4, "Stock Transfer New", SAPbouiCOM.BoMenuType.mt_STRING, "invtransudo", menuitem)
    End Sub
   
    Public Sub SetFilters()
        Dim objFilters As SAPbouiCOM.EventFilters
        Dim objFilter As SAPbouiCOM.EventFilter
        objFilters = New SAPbouiCOM.EventFilters


        objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
        objFilter.AddEx("BarcodePrint")
        objFilter.AddEx("invtransudo")
        objFilter.AddEx("BarRprt")
        objFilter.AddEx("PurchaseWarranty")
        objFilter.AddEx("batchselection")
        objFilter.AddEx("barcodeselection")
        objFilter.AddEx("65051")
        objFilter.AddEx("21")
        objFilter.AddEx("63")
        objFilter.AddEx("24")
        objFilter.AddEx("143")
        objFilter.AddEx("721")
        objFilter.AddEx("65053")

        objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK)
        objFilter.AddEx("batchselection")
        objFilter.AddEx("barcodeselection")
        objFilter.AddEx("invtransudo")


        objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK)
        objFilter.AddEx("BarcodePrint")
        objFilter.AddEx("invtransudo")
        objFilter.AddEx("BarRprt")

        objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_CLICK)
        objFilter.AddEx("BarcodePrint")
        objFilter.AddEx("invtransudo")
        objFilter.AddEx("BarRprt")
        objFilter.AddEx("PurchaseWarranty")
        objFilter.AddEx("65051")
        objFilter.AddEx("batchselection")
        objFilter.AddEx("barcodeselection")
        objFilter.AddEx("143")
        objFilter.AddEx("63")
        objFilter.AddEx("21")
        objFilter.AddEx("721")


        objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
        objFilter.AddEx("BarcodePrint")
        objFilter.AddEx("invtransudo")
        objFilter.AddEx("BarRprt")
        '  objFilter.AddEx("batchselection")


        objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
        objFilter.AddEx("invtransudo")
        objFilter.AddEx("BarRprt")
        ' objFilter.AddEx("batchselection")


        objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK)
        objFilter.AddEx("BarRprt")
        '  objFilter.AddEx("batchselection")


        objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT)
        objFilter.AddEx("invtrans")

        '  objFilter.AddEx("batchselection")


        objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_VALIDATE)
        objFilter.AddEx("BarcodePrint")
        objFilter.AddEx("batchselection")
        objFilter.AddEx("invtransudo")
        '  objFilter.AddEx("barcodeselection")

        objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_RESIZE)
        objFilter.AddEx("BarcodePrint")
        ' objFilter.AddEx("batchselection")


        objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_LOST_FOCUS)
        objFilter.AddEx("BarcodePrint")
        objFilter.AddEx("batchselection")
        'objFilter.AddEx("barcodeselection")




        objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_KEY_DOWN)

        'objFilter.AddEx("batchselection")
        objFilter.AddEx("barcodeselection")
        objFilter.AddEx("65051")
        objFilter.AddEx("65053")

        objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_CLOSE)
        objFilter.AddEx("143")
        objFilter.AddEx("721")
        ' objFilter.AddEx("batchselection")


        '  objApplication.SetFilter(objFilters)

    End Sub
    ' For Menu Creation
    Private Function CreateMenu(ByVal ImagePath As String, ByVal Position As Int32, ByVal DisplayName As String, ByVal MenuType As SAPbouiCOM.BoMenuType, ByVal UniqueID As String, ByVal ParentMenu As SAPbouiCOM.MenuItem) As SAPbouiCOM.MenuItem
        Try
            Dim oMenuPackage As SAPbouiCOM.MenuCreationParams
            oMenuPackage = objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
            oMenuPackage.Image = ImagePath
            oMenuPackage.Position = Position
            oMenuPackage.Type = MenuType
            oMenuPackage.UniqueID = UniqueID
            oMenuPackage.String = DisplayName
            ParentMenu.SubMenus.AddEx(oMenuPackage)
        Catch ex As Exception
            objApplication.StatusBar.SetText("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
        End Try
        Return ParentMenu.SubMenus.Item(UniqueID)
    End Function
    Private Sub objApplication_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles objApplication.RightClickEvent
        If eventInfo.BeforeAction Then
            Select Case objAddOn.objApplication.Forms.ActiveForm.TypeEx
                'Case clsInvTransferUDO.formtype
                '    objInvTransferUDO.RightClickEvent(eventInfo, BubbleEvent)
            End Select
        Else

            'Select Case objAddOn.objApplication.Forms.ActiveForm.TypeEx
            '    Case clsInvTransferUDO.formtype
            '        objInvTransferUDO.RightClickEvent(eventInfo, BubbleEvent)
            'End Select

        End If

    End Sub
    Private Sub createUDO(ByVal tblname As String, ByVal udocode As String, ByVal udoname As String, ByVal type As SAPbobsCOM.BoUDOObjType, Optional ByVal DfltForm As Boolean = False, Optional ByVal FindForm As Boolean = False)
        Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
        Dim creationPackage As SAPbouiCOM.FormCreationParams
        Dim objform As SAPbouiCOM.Form
        Dim lRetCode As Long
        oUserObjectMD = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
        If Not oUserObjectMD.GetByKey(udocode) Then
            oUserObjectMD.Code = udocode
            oUserObjectMD.Name = udoname
            oUserObjectMD.ObjectType = type
            oUserObjectMD.TableName = tblname
            oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
            oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
            If DfltForm = True Then
                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.FormColumns.FormColumnAlias = "Code"
                oUserObjectMD.FormColumns.FormColumnDescription = "Code"
                oUserObjectMD.FormColumns.Add()
                oUserObjectMD.FormColumns.FormColumnAlias = "Name"
                oUserObjectMD.FormColumns.FormColumnDescription = "Name"
                oUserObjectMD.FormColumns.Add()
            End If
            If FindForm = True Then
                If type = SAPbobsCOM.BoUDOObjType.boud_MasterData Then
                    oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                    Select Case udoname
                        Case "VehicleTypeMaster"
                            oUserObjectMD.FindColumns.ColumnAlias = "Code"
                            oUserObjectMD.FindColumns.ColumnDescription = "Code"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "Name"
                            oUserObjectMD.FindColumns.ColumnDescription = "Name"
                            oUserObjectMD.FindColumns.Add()
                    End Select
                ElseIf type = SAPbobsCOM.BoUDOObjType.boud_Document Then
                    oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                    Select Case udoname
                    End Select
                End If
            End If

            lRetCode = oUserObjectMD.Add()
            If lRetCode <> 0 Then
                ' MsgBox("error" + CStr(lRetCode))
                'MsgBox(objAddOn.objCompany.GetLastErrorDescription)
            Else

            End If
            If DfltForm = True Then
                creationPackage = objAddOn.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
                creationPackage.ObjectType = "1"
                creationPackage.UniqueID = udoname
                creationPackage.FormType = udoname
                creationPackage.BorderStyle = SAPbouiCOM.BoFormTypes.ft_Fixed
                objform = objAddOn.objApplication.Forms.AddEx(creationPackage)
            End If
        End If

    End Sub
    Private Sub objApplication_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles objApplication.FormDataEvent
        If BusinessObjectInfo.FormTypeEx = "940" And BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD And BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
            'If objAddOn.RetValue1 = "2" Then
            '    ' objAddOn.objgenerate.Getinfo()
            'End If
        End If
    End Sub
    Private Sub createTables()
        Try
            Dim objUDFEngine As New Mukesh.SBOLib.UDFEngine(objCompany)
            objAddOn.objApplication.SetStatusBarMessage("Creating Tables Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Long, False)
            objUDFEngine.AddAlphaField("OSRN", "printed", "Printed", 1)

            'objUDFEngine.AddAlphaField("OSRN", "boxno", "BoxNo", 10)
            'objUDFEngine.AddAlphaField("OITB", "prefix", "Barcode Prefix", 10)
            'objUDFEngine.AddNumericField("PDN1", "bqty", "Barcoded Qty", 10)
            objUDFEngine.AddAlphaField("OPDN", "bstatus", "BarcodeStatus", 5, "Y")
            objUDFEngine.AddAlphaField("OPDN", "serial", "Serial", 2, "N")
            'objUDFEngine.AddAlphaField("PDN1", "o7_color", "Color", 30)
            'objUDFEngine.AddAlphaField("PDN1", "o8_siz", "Size", 5)
            '---------- Preethi ---------------

            'objUDFEngine.AddAlphaField("PDN1", "Brand", "Brand", 20)
            'objUDFEngine.AddAlphaField("PDN1", "Catagory", "Category", 20)
            'objUDFEngine.AddAlphaField("PDN1", "Sleeve", "Sleeve", 20)
            'objUDFEngine.AddAlphaField("PDN1", "Fit", "Fit", 20)
            'objUDFEngine.AddAlphaField("PDN1", "Size", "Size", 5)
            'objUDFEngine.AddAlphaField("PDN1", "Style", "Style", 20)
            'objUDFEngine.AddAlphaField("PDN1", "Colour", "Colour", 20)
            'objUDFEngine.AddAlphaField("PDN1", "Mode", "Mode", 20)
            'objUDFEngine.AddAlphaField("PDN1", "City", "City", 20)
            'objUDFEngine.AddAlphaField("PDN1", "EAN", "EAN", 25)

            objUDFEngine.AddFloatField("PDN1", "bqty", "BarcodedQty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
            objUDFEngine.AddFloatField("PDN1", "MRP", "MRP", SAPbobsCOM.BoFldSubTypes.st_Price)
            objUDFEngine.AddAlphaField("OPDN", "btstatus", "BatchStatus", 5, "N") ' Batch print
            objUDFEngine.AddAlphaField("OPDN", "batch", "Batch", 2, "N") ' Batch print
            objUDFEngine.AddAlphaField("OMRC", "shtname", "ShortName", 3)

            'objUDFEngine.AddAlphaField("OSRN", "Brand", "Brand", 20)
            'objUDFEngine.AddAlphaField("OSRN", "Catagory", "", 20)
            'objUDFEngine.AddAlphaField("OSRN", "Sleeve", "Sleeve", 20)
            'objUDFEngine.AddAlphaField("OSRN", "Fit", "Fit", 20)
            'objUDFEngine.AddAlphaField("OSRN", "Size", "Size", 5)
            'objUDFEngine.AddAlphaField("OSRN", "Style", "Style", 20)
            'objUDFEngine.AddAlphaField("OSRN", "Colour", "Colour", 20)
            'objUDFEngine.AddAlphaField("OSRN", "Mode", "Mode", 20)
            'objUDFEngine.AddAlphaField("OSRN", "City", "City", 20)
            'objUDFEngine.AddFloatField("OSRN", "MRP", "MRP", SAPbobsCOM.BoFldSubTypes.st_Price)
            'objUDFEngine.AddAlphaField("OSRN", "EAN", "EAN", 25)

            'MI_INVTRH
            '******************Added By Bowya************************************************************''''''''''''''''''''''''''
            'objUDFEngine.AddNumericField("IGN1", "bqty", "Barcoded Qty", 10)
            'objUDFEngine.AddAlphaField("OIGN", "bstatus", "BarcodeStatus", 5, "N")
            'objUDFEngine.AddAlphaField("OIGN", "serial", "Serial", 2, "N")
            'objUDFEngine.AddAlphaField("OIGN", "btstatus", "BarcodeStatus", 5, "N")
            'objUDFEngine.AddAlphaField("OIGN", "batch", "Serial", 2, "N")

            ''objUDFEngine.AddAlphaField("IGN1", "o7_color", "Color", 30)
            ''objUDFEngine.AddAlphaField("IGN1", "o8_siz", "Size", 5)
            'objUDFEngine.AddFloatField("IGN1", "bqty", "BarcodedQty", SAPbobsCOM.BoFldSubTypes.st_Quantity)
            'objUDFEngine.AddFloatField("IGN1", "mrp", "MRP", SAPbobsCOM.BoFldSubTypes.st_Price)
            '************************************************************************************************************
            
            'objUDFEngine.CreateTable("TMPSERI", "Serial Temp", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            'objUDFEngine.AddAlphaField("@TMPSERI", "itemcode", "ItemCode", 50)
            'objUDFEngine.AddNumericField("@TMPSERI", "qty", "Qty", 10)
            'objUDFEngine.AddNumericField("@TMPSERI", "actqty", "ActQty", 10)
            'objUDFEngine.AddAlphaField("@TMPSERI", "prefix", "Prefix", 10)
            'objUDFEngine.AddNumericField("@TMPSERI", "serial", "Serial", 10)
            'objUDFEngine.AddAlphaField("@TMPSERI", "serialno", "SerialNo", 20)
            'objUDFEngine.AddAlphaField("@TMPSERI", "ltserial", "LastSerialNo", 20)
            'objUDFEngine.AddDateField("@TMPSERI", "date", "Date", SAPbobsCOM.BoFldSubTypes.st_None)
            'objUDFEngine.AddAlphaField("@TMPSERI", "stserial", "StoreSerial", 20)
            'objUDFEngine.AddNumericField("@TMPSERI", "docnum", "DocNum", 10)
            'objUDFEngine.AddAlphaField("@TMPSERI", "hostname", "HostName", 20)
            'objUDFEngine.AddAlphaField("@TMPSERI", "status", "Status", 2)


            'objUDFEngine.CreateTable("TMPBAT", "Batch Temp", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            'objUDFEngine.AddAlphaField("@TMPBAT", "itemcode", "ItemCode", 50)
            'objUDFEngine.AddNumericField("@TMPBAT", "qty", "Qty", 10)
            'objUDFEngine.AddNumericField("@TMPBAT", "actqty", "ActQty", 10)
            'objUDFEngine.AddAlphaField("@TMPBAT", "prefix", "Prefix", 10)
            'objUDFEngine.AddNumericField("@TMPBAT", "batch", "Batch", 10)
            'objUDFEngine.AddAlphaField("@TMPBAT", "batchno", "BatchNo", 20)
            'objUDFEngine.AddAlphaField("@TMPBAT", "ltbatch", "LastBatchNo", 20)
            'objUDFEngine.AddDateField("@TMPBAT", "date", "Date", SAPbobsCOM.BoFldSubTypes.st_None)
            'objUDFEngine.AddAlphaField("@TMPBAT", "stbatch", "StoreBatch", 20)
            'objUDFEngine.AddNumericField("@TMPBAT", "docnum", "DocNum", 10)
            'objUDFEngine.AddAlphaField("@TMPBAT", "hostname", "HostName", 20)
            'objUDFEngine.AddAlphaField("@TMPBAT", "status", "Status", 2)

            ' Internal whse selection Damage stores/Reconcilation
            objUDFEngine.addField("OLCT", "internal", "Internal", SAPbobsCOM.BoFieldTypes.db_Alpha, 5, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "N")
            objApplication.SetStatusBarMessage("Created Tables Successfully!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        Catch ex As Exception
            objAddOn.objApplication.SetStatusBarMessage(objAddOn.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub
    Private Sub createUDOC(ByVal tblname As String, ByVal udocode As String, ByVal udoname As String, ByVal childTable() As String, ByVal type As SAPbobsCOM.BoUDOObjType, Optional ByVal DfltForm As Boolean = False, Optional ByVal FindForm As Boolean = False)
        objAddOn.objApplication.SetStatusBarMessage("UDO Created Please Wait..", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
        Dim creationPackage As SAPbouiCOM.FormCreationParams
        Dim objform As SAPbouiCOM.Form
        Dim i As Integer
        Dim c_Yes As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tYES
        Dim lRetCode As Long
        oUserObjectMD = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
        If Not oUserObjectMD.GetByKey(udocode) Then
            oUserObjectMD.Code = udocode
            oUserObjectMD.Name = udoname
            oUserObjectMD.ObjectType = type
            oUserObjectMD.TableName = tblname
            oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
            oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
            If DfltForm = True Then
                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
            End If
            If FindForm = True Then
                If type = SAPbobsCOM.BoUDOObjType.boud_Document Then
                    oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                   
                Else
                    oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                    Select Case udoname
                        'Case "MIPLMHR"
                        '    oUserObjectMD.FindColumns.ColumnAlias = "Code"
                        '    oUserObjectMD.FindColumns.ColumnDescription = "Code"
                        '    oUserObjectMD.FindColumns.Add()
                        '    oUserObjectMD.FindColumns.ColumnAlias = "Name"
                        '    oUserObjectMD.FindColumns.ColumnDescription = "Name"
                        '    oUserObjectMD.FindColumns.Add()
                        '    oUserObjectMD.FindColumns.ColumnAlias = "U_Empno"
                        '    oUserObjectMD.FindColumns.ColumnDescription = "Employee.No"
                        '    oUserObjectMD.FindColumns.Add()
                        '    oUserObjectMD.FindColumns.ColumnAlias = "U_Frstnam"
                        '    oUserObjectMD.FindColumns.ColumnDescription = "First Name"
                        '    oUserObjectMD.FindColumns.Add()
                    End Select
                End If
            End If
            If childTable.Length > 0 Then
                For i = 0 To childTable.Length - 2
                    If Trim(childTable(i)) <> "" Then
                        oUserObjectMD.ChildTables.TableName = childTable(i)
                        oUserObjectMD.ChildTables.Add()
                    End If
                Next
            End If
            lRetCode = oUserObjectMD.Add()
            If lRetCode <> 0 Then
                MsgBox("error" + CStr(lRetCode))
                MsgBox(objAddOn.objCompany.GetLastErrorDescription)
            Else
            End If
            If DfltForm = True Then
                creationPackage = objAddOn.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
                ' Need to set the parameter with the object unique ID
                creationPackage.ObjectType = "1"
                creationPackage.UniqueID = udoname
                creationPackage.FormType = udoname
                creationPackage.BorderStyle = SAPbouiCOM.BoFormTypes.ft_Fixed
                objform = objAddOn.objApplication.Forms.AddEx(creationPackage)
            End If
        End If
        objAddOn.objApplication.SetStatusBarMessage("UDO Created Suceessfully..", SAPbouiCOM.BoMessageTime.bmt_Short, False)
    End Sub

    Private Sub objApplication_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles objApplication.AppEvent
        If EventType = SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged Or EventType = SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition Or EventType = SAPbouiCOM.BoAppEventTypes.aet_ShutDown Then
            Try
                If objCompany.Connected Then objCompany.Disconnect()
                objCompany = Nothing
                objApplication = Nothing
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objCompany)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objApplication)
                GC.Collect()
            Catch ex As Exception
            End Try
            End
        End If
    End Sub
    Private Sub applyFilter()
        Dim oFilters As SAPbouiCOM.EventFilters
        Dim oFilter As SAPbouiCOM.EventFilter
        oFilters = New SAPbouiCOM.EventFilters
        'Item Master Data 
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_GOT_FOCUS)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK)
    End Sub

    Public Sub New()

    End Sub

    Private Sub objApplication_ProgressBarEvent(ByRef pVal As SAPbouiCOM.ProgressBarEvent, ByRef BubbleEvent As Boolean) Handles objApplication.ProgressBarEvent

    End Sub
End Class


