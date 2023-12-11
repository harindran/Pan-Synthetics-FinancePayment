Imports SAPbouiCOM.Framework
Imports System.IO

Namespace Finance_Payment
    Public Class clsAddon
        Public WithEvents objapplication As SAPbouiCOM.Application
        Public objcompany As SAPbobsCOM.Company
        Public objmenuevent As clsMenuEvent
        Public objrightclickevent As clsRightClickEvent
        Public objglobalmethods As clsGlobalMethods
        Public objAP As SysAPInvoice
        Public objIncPayment As ClsIncomingPayments
        Public objOutgoingPayment As ClsOutgoingPayments
        Public objJE As ClsJE
        Public objUDFJE As ClsUDF_JE
        Dim objform As SAPbouiCOM.Form
        Dim strsql As String = ""
        Dim objrs As SAPbobsCOM.Recordset
        Dim print_close As Boolean = False
        Public HANA As Boolean = True
        'Public HANA As Boolean = False
        Public HWKEY() As String = New String() {"L1653539483", "D0872452844"}

        Public Sub Intialize(ByVal args() As String)
            Try
                Dim oapplication As Application
                If (args.Length < 1) Then oapplication = New Application Else oapplication = New Application(args(0))
                objapplication = Application.SBO_Application
                If isValidLicense() Then
                    objapplication.StatusBar.SetText("Establishing Company Connection Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    objcompany = Application.SBO_Application.Company.GetDICompany()

                    Create_DatabaseFields() 'UDF & UDO Creation Part    
                    Menu() 'Menu Creation Part
                    Create_Objects() 'Object Creation Part
                    If HANA Then
                        Localization = objglobalmethods.getSingleValue("select ""LawsSet"" from CINF")
                        CostCenter = objaddon.objglobalmethods.getSingleValue("select ""MDStyle"" from OADM")
                        MainCurr = objaddon.objglobalmethods.getSingleValue("select ""MainCurncy"" from OADM")
                        BranchSplitReconciliation = objaddon.objglobalmethods.getSingleValue("select ""U_BranchSplit"" from OADM")
                        PaymentWithReco = objaddon.objglobalmethods.getSingleValue("select ""U_PayWithReco"" from OUSR where ""USER_CODE""='" & objaddon.objcompany.UserName & "'")
                    Else
                        Localization = objglobalmethods.getSingleValue("select LawsSet from CINF")
                        CostCenter = objaddon.objglobalmethods.getSingleValue("select MDStyle from OADM")
                        MainCurr = objaddon.objglobalmethods.getSingleValue("select MainCurncy from OADM")
                        BranchSplitReconciliation = objaddon.objglobalmethods.getSingleValue("select U_BranchSplit from OADM")
                        PaymentWithReco = objaddon.objglobalmethods.getSingleValue("select U_PayWithReco from OUSR where USER_CODE='" & objaddon.objcompany.UserName & "'")
                    End If
                    Add_Authorizations() 'User Permissions
                    objapplication.StatusBar.SetText("Add-on Connected Successfully..!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    oapplication.Run()

                Else
                    objapplication.StatusBar.SetText("Add-on Disconnected due to license mismatch..!!!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                End If
                'System.Windows.Forms.Application.Run()
            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try
        End Sub

        Function isValidLicense() As Boolean
            Try
                Try
                    If objapplication.Forms.ActiveForm.TypeCount > 0 Then
                        For i As Integer = 0 To objapplication.Forms.ActiveForm.TypeCount - 1
                            objapplication.Forms.ActiveForm.Close()
                        Next
                    End If
                Catch ex As Exception
                End Try

                'If Not HANA Then
                '    objapplication.Menus.Item("1030").Activate()
                'End If
                objapplication.Menus.Item("257").Activate()
                Dim CrrHWKEY As String = objapplication.Forms.ActiveForm.Items.Item("79").Specific.Value.ToString.Trim
                objapplication.Forms.ActiveForm.Close()

                For i As Integer = 0 To HWKEY.Length - 1
                    If HWKEY(i).Trim = CrrHWKEY.Trim Then
                        Return True
                    End If
                Next
                MsgBox("Installing Add-On failed due to License mismatch", MsgBoxStyle.OkOnly, "License Management")
                Return False
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                'MsgBox(ex.ToString)
            End Try
            Return True
        End Function

        Private Sub Create_Objects()
            objmenuevent = New clsMenuEvent
            objrightclickevent = New clsRightClickEvent
            objglobalmethods = New clsGlobalMethods
            objAP = New SysAPInvoice
            objIncPayment = New ClsIncomingPayments
            objOutgoingPayment = New ClsOutgoingPayments
            objJE = New ClsJE
            objUDFJE = New ClsUDF_JE
        End Sub

        Private Sub Create_DatabaseFields()
            'If objapplication.Company.UserName.ToString.ToUpper <> "MANAGER" Then

            'If objapplication.MessageBox("Do you want to execute the field Creations?", 2, "Yes", "No") <> 1 Then Exit Sub
            objapplication.StatusBar.SetText("Creating Database Fields.Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            Dim objtable As New clsTable
            objtable.FieldCreation()
            'End If

        End Sub

        Public Sub Add_Authorizations()
            Try
                objaddon.objglobalmethods.AddToPermissionTree("Altrocks Tech", "ATPL_ADD-ON", "", "", "Y"c) 'Level 1 - Company Name
                objaddon.objglobalmethods.AddToPermissionTree("Customized Payment", "ATPL_CUSTPAY", "PAYINIT", "ATPL_ADD-ON", "Y"c)
            Catch ex As Exception
            End Try
        End Sub

#Region "Menu Creation Details"

        Private Sub Menu()
            Dim Menucount As Integer = 1

            'CreateMenu("", Menucount, "Multi-Branch A/P Service Invoice", SAPbouiCOM.BoMenuType.mt_STRING, "MBAPSI", "2304") : Menucount += 1
            Menucount = 1
            CreateMenu("", Menucount, "Customized Payment", SAPbouiCOM.BoMenuType.mt_STRING, "PAYINOUT", "43537") : Menucount += 1 ' "43537"


        End Sub

        Private Sub CreateMenu(ByVal ImagePath As String, ByVal Position As Int32, ByVal DisplayName As String, ByVal MenuType As SAPbouiCOM.BoMenuType, ByVal UniqueID As String, ByVal ParentMenuID As String)
            Try
                Dim oMenuPackage As SAPbouiCOM.MenuCreationParams
                Dim parentmenu As SAPbouiCOM.MenuItem
                parentmenu = objapplication.Menus.Item(ParentMenuID)
                If parentmenu.SubMenus.Exists(UniqueID.ToString) Then Exit Sub
                oMenuPackage = objapplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                oMenuPackage.Image = ImagePath
                oMenuPackage.Position = Position
                oMenuPackage.Type = MenuType
                oMenuPackage.UniqueID = UniqueID
                oMenuPackage.String = DisplayName
                parentmenu.SubMenus.AddEx(oMenuPackage)
            Catch ex As Exception
                objapplication.StatusBar.SetText("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
            End Try
            'Return ParentMenu.SubMenus.Item(UniqueID)
        End Sub

#End Region

#Region "ItemEvent_Link Button"

        Private Sub objapplication_ItemEvent(FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles objapplication.ItemEvent
            Try
                Dim oform As SAPbouiCOM.Form
                'If objapplication.Forms.Count > 0 Then oform = objapplication.Forms.ActiveForm
                Select Case pVal.FormTypeEx
                    Case "141"
                        objAP.ItemEvent(FormUID, pVal, BubbleEvent)
                    Case "170"
                        objIncPayment.ItemEvent(FormUID, pVal, BubbleEvent)
                    Case "426"
                        objOutgoingPayment.ItemEvent(FormUID, pVal, BubbleEvent)
                    Case "392"
                        objJE.ItemEvent(FormUID, pVal, BubbleEvent)
                    Case "-392"
                        objUDFJE.ItemEvent(FormUID, pVal, BubbleEvent)
                End Select
                If pVal.BeforeAction Then
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_CLICK
                            If objaddon.objapplication.Forms.ActiveForm.TypeEx = "0" Then
                                If pVal.ItemUID = "1" Then
                                    TempForm = True
                                Else
                                    TempForm = False
                                End If
                            End If
                            If bModal And (objaddon.objapplication.Forms.ActiveForm.TypeEx = "MBAPSI" Or objaddon.objapplication.Forms.ActiveForm.TypeEx = "FINPAY" Or objaddon.objapplication.Forms.ActiveForm.TypeEx = "FOUTPAY") Then
                                BubbleEvent = False
                                objapplication.Forms.Item("DistRule").Select()
                            ElseIf pModal And (objaddon.objapplication.Forms.ActiveForm.TypeEx = "FINPAY" Or objaddon.objapplication.Forms.ActiveForm.TypeEx = "FOUTPAY") Then
                                BubbleEvent = False
                                'oform = objapplication.Forms.GetFormByTypeAndCount("PAYM", 0)
                                'oform.Select()
                                objapplication.Forms.Item("PAYM").Select()
                            ElseIf objaddon.objapplication.Forms.ActiveForm.TypeEx = "PAYINIT" Then
                                If objaddon.objapplication.Forms.ActiveForm.Items.Item("opinrec").Specific.Selected = True Then
                                    If FormExist("FOITR") Then
                                        BubbleEvent = False
                                    End If
                                    'objapplication.Forms.Item("FOITR").Select()
                                ElseIf objaddon.objapplication.Forms.ActiveForm.Items.Item("opincpay").Specific.Selected = True Then
                                    If FormExist("FINPAY") Then
                                        BubbleEvent = False
                                    End If
                                    'objapplication.Forms.Item("FINPAY").Select()
                                ElseIf objaddon.objapplication.Forms.ActiveForm.Items.Item("opoutpay").Specific.Selected = True Then
                                    If FormExist("FOUTPAY") Then
                                        BubbleEvent = False
                                    End If
                                    'objapplication.Forms.Item("FOUTPAY").Select()
                                End If
                            ElseIf objaddon.objapplication.Forms.ActiveForm.TypeEx = "FOITR" Then
                                If FormExist("AdjTran") Then
                                    BubbleEvent = False
                                End If
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                            Dim EventEnum As SAPbouiCOM.BoEventTypes
                            EventEnum = pVal.EventType
                            If FormUID = "DistRule" And (EventEnum = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE) And bModal Then
                                bModal = False
                            ElseIf FormUID = "PAYM" And (EventEnum = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE) And pModal Then
                                pModal = False
                            ElseIf FormUID = "AdjTran" And (EventEnum = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE) And bModal Then
                                bModal = False
                            End If

                    End Select
                Else
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_CLICK
                            If pVal.ItemUID = "1" And objaddon.objapplication.Forms.ActiveForm.TypeEx = "866" Then
                                Try
                                    oform = objapplication.Forms.ActiveForm
                                    oform = objaddon.objapplication.Forms.GetForm("PAYM", 1)
                                    strsql = objaddon.objglobalmethods.getSingleValue("Select ""Rate"" from ORTT where ""RateDate""= '" & Now.Date.ToString("yyyyMMdd") & "' and ""Currency""='" & oform.Items.Item("tcurr").Specific.Selected.Value & "' ")
                                    If strsql <> "" Then oform.Items.Item("trate").Specific.String = strsql
                                Catch ex As Exception

                                End Try
                            End If
                            'Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            '    If pVal.FormTypeEx = "141" Then
                            '        Try
                            '            objAP.ItemEvent(FormUID, pVal, BubbleEvent)
                            '        Catch ex As Exception
                            '        End Try
                            '    ElseIf pVal.FormTypeEx = "170" Then
                            '        objIncPayment.ItemEvent(FormUID, pVal, BubbleEvent)
                            '    End If
                    End Select
                End If

            Catch ex As Exception

            End Try
        End Sub

#End Region

#Region "Form Data event"
        Private Sub objApplication_FormDataEvent(ByRef pVal As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles objapplication.FormDataEvent
            Try
                'If BusinessObjectInfo.BeforeAction = True Then
                Select Case pVal.FormTypeEx
                    Case "141"
                        objAP.FormDataEvent(pVal, BubbleEvent)
                    Case "170"
                        objIncPayment.FormDataEvent(pVal, BubbleEvent)
                    Case "426"
                        objOutgoingPayment.FormDataEvent(pVal, BubbleEvent)
                    Case "392"
                        objJE.FormDataEvent(pVal, BubbleEvent)
                End Select
                'End If

            Catch ex As Exception
                'objAddOn.objApplication.StatusBar.SetText(ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End Try
        End Sub
#End Region

#Region "Menu Event"

        Public Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles objapplication.MenuEvent
            Try
                Select Case pVal.MenuUID
                    Case "1281", "1282", "1283", "1284", "1285", "1286", "1287", "1300", "1288", "1289", "1290", "1291", "1304", "1292", "1293", "CPYD", "FINPAY", "FOUTPAY", "6005", "6913"
                        objmenuevent.MenuEvent_For_StandardMenu(pVal, BubbleEvent)
                    Case "MBAPSI", "PAYINOUT"
                        MenuEvent_For_FormOpening(pVal, BubbleEvent)
                        'Case "1293"
                        '    BubbleEvent = False
                    Case "519"
                        MenuEvent_For_Preview(pVal, BubbleEvent)
                End Select
            Catch ex As Exception
                'objaddon.objapplication.SetStatusBarMessage("Error in SBO_Application MenuEvent" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub

        Public Sub MenuEvent_For_Preview(ByRef pval As SAPbouiCOM.MenuEvent, ByRef bubbleevent As Boolean)
            Dim oform = objaddon.objapplication.Forms.ActiveForm()
            'If pval.BeforeAction Then
            '    If oform.TypeEx = "TRANOLVA" Then MenuEvent_For_PrintPreview(oform, "8f481d5cf08e494f9a83e1e46ab2299e", "txtentry") : bubbleevent = False
            '    If oform.TypeEx = "TRANOLAP" Then MenuEvent_For_PrintPreview(oform, "f15ee526ac514070a9d546cda7f94daf", "txtentry") : bubbleevent = False
            '    If oform.TypeEx = "OLSE" Then MenuEvent_For_PrintPreview(oform, "e47ed373e0cc48efb47c9773fba64fc3", "txtentry") : bubbleevent = False
            'End If
        End Sub

        Private Sub MenuEvent_For_PrintPreview(ByVal oform As SAPbouiCOM.Form, ByVal Menuid As String, ByVal Docentry_field As String)
            'Try
            '    Dim Docentry_Est As String = oform.Items.Item(Docentry_field).Specific.String
            '    If Docentry_Est = "" Then Exit Sub
            '    print_close = False
            '    objaddon.objapplication.Menus.Item(Menuid).Activate()
            '    oform = objaddon.objapplication.Forms.ActiveForm()
            '    oform.Items.Item("1000003").Specific.string = Docentry_Est
            '    oform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            '    print_close = True
            'Catch ex As Exception
            'End Try
        End Sub
        Public Function FormExist(ByVal FormID As String) As Boolean
            Try
                FormExist = False
                For Each uid As SAPbouiCOM.Form In objaddon.objapplication.Forms
                    If uid.TypeEx = FormID Then
                        FormExist = True
                        Exit For
                    End If
                Next
                If FormExist Then
                    objaddon.objapplication.Forms.Item(FormID).Visible = True
                    objaddon.objapplication.Forms.Item(FormID).Select()
                End If
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        End Function

        Public Sub MenuEvent_For_FormOpening(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Try
                If pVal.BeforeAction = False Then
                    Select Case pVal.MenuUID
                        Case "MBAPSI"
                            Dim activeform As New FrmAPService_Invoice
                            activeform.Show()
                        Case "PAYINOUT"
                            If Not FormExist("PAYINIT") Then
                                Dim activeform As New FrmPayInitialize
                                activeform.Show()
                            End If

                    End Select

                End If
            Catch ex As Exception
                'objaddon.objapplication.SetStatusBarMessage("Error in Form Opening MenuEvent" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub

#End Region

#Region "LayoutKeyEvent"

        Public Sub SBO_Application_LayoutKeyEvent(ByRef eventInfo As SAPbouiCOM.LayoutKeyInfo, ByRef BubbleEvent As Boolean) Handles objapplication.LayoutKeyEvent
            'Dim oForm_Layout As SAPbouiCOM.Form = Nothing
            'If SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.BusinessObject.Type = "NJT_CES" Then
            '    oForm_Layout = SAPbouiCOM.Framework.Application.SBO_Application.Forms.Item(eventInfo.FormUID)
            'End If
        End Sub

#End Region

#Region "Application Event"

        Public Sub SBO_Application_AppEvent(EventType As SAPbouiCOM.BoAppEventTypes) Handles objapplication.AppEvent
            If EventType = SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged Or EventType = SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition Or EventType = SAPbouiCOM.BoAppEventTypes.aet_ShutDown Then
                Try

                    If objcompany.Connected Then objcompany.Disconnect()
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objcompany)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objapplication)
                    objcompany = Nothing
                    'objapplication = Nothing
                    GC.Collect()
                    System.Windows.Forms.Application.Exit()
                    End
                Catch ex As Exception
                End Try
            End If
            'Select Case EventType
            '    Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown Or SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged Or SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition
            '        Try
            '            'If objcompany.Connected Then objcompany.Disconnect()
            '            ''System.Runtime.InteropServices.Marshal.ReleaseComObject(objcompany)
            '            ''System.Runtime.InteropServices.Marshal.ReleaseComObject(objapplication)
            '            'objcompany = Nothing
            '            ''objapplication = Nothing
            '            'GC.Collect()
            '            'System.Windows.Forms.Application.Exit()
            '            'End
            '        Catch ex As Exception
            '        End Try
            '        'Case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged
            '        '    End
            '        '    'Case SAPbouiCOM.BoAppEventTypes.aet_FontChanged
            '        '    '    End
            '        '    'Case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged
            '        '    '    End
            '        'Case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition
            '        '    End
            'End Select
        End Sub

#End Region

#Region "Right Click Event"

        Private Sub objapplication_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles objapplication.RightClickEvent
            Try
                Select Case objaddon.objapplication.Forms.ActiveForm.TypeEx
                    Case "MBAPSI", "PAYINIT", "PAYM", "170", "426", "FOITR"
                        objrightclickevent.RightClickEvent(eventInfo, BubbleEvent)

                End Select
            Catch ex As Exception

            End Try
        End Sub

#End Region


    End Class

End Namespace
