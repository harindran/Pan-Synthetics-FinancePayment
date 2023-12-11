Imports System.Windows.Forms
Imports SAPbobsCOM
Imports SAPbouiCOM.Framework

Namespace Finance_Payment

    Public Class clsMenuEvent
        Dim objform As SAPbouiCOM.Form
        Dim objglobalmethods As New clsGlobalMethods
        Dim AccJENo, StrSQL As String
        Public Sub MenuEvent_For_StandardMenu(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Try
                Select Case objaddon.objapplication.Forms.ActiveForm.TypeEx
                    Case "MBAPSI"
                        Mul_Branch_AP_Service_Invoice_MenuEvent(pVal, BubbleEvent)
                    Case "141", "170", "-170", "426", "-426", "392", "-392"
                        Default_Sample_MenuEvent(pVal, BubbleEvent)
                    Case "PAYINIT"
                        PaymentInit_MenuEvent(pVal, BubbleEvent)
                    Case "PAYM"
                        Payment_Means_MenuEvent(pVal, BubbleEvent)
                    Case "FINPAY"
                        InPayments_MenuEvent(pVal, BubbleEvent)
                    Case "FOUTPAY"
                        OutPayments_MenuEvent(pVal, BubbleEvent)
                    Case "FOITR"
                        InternalReconciliation_MenuEvent(pVal, BubbleEvent)

                End Select
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Default_Sample_MenuEvent(ByVal pval As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                Dim oUDFForm As SAPbouiCOM.Form
                If pval.BeforeAction = True Then
                    Select Case pval.MenuUID
                        Case "6005"
                            ' objform = objaddon.objapplication.Forms.ActiveForm
                            If objform.Items.Item("58").Specific.Selected = False Then Exit Sub
                            If objaddon.objapplication.Forms.ActiveForm.Items.Item("chkactive").Specific.Checked = True And objaddon.objapplication.Forms.ActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                BubbleEvent = False
                            End If
                        Case "6913"
                            'If objform.TypeEx = "392" Then
                            '    oUDFForm = objaddon.objapplication.Forms.Item(objform.UDFFormUID)
                            '    oUDFForm.Items.Item("U_TransId").Enabled = False
                            '    oUDFForm.Items.Item("U_IEntry").Enabled = False
                            '    oUDFForm.Items.Item("U_OEntry").Enabled = False
                            '    oUDFForm.Items.Item("U_IntRecNo").Enabled = False
                            'End If
                        Case "1284" 'Cancel
                            If objform.TypeEx = "170" Or objform.TypeEx = "426" Then
                                AccJENo = objform.Items.Item("tjeno").Specific.String
                                Try
                                    oUDFForm = objaddon.objapplication.Forms.Item(objform.UDFFormUID)
                                    If AccJENo = "" Then AccJENo = oUDFForm.Items.Item("U_JENo").Specific.String
                                Catch ex As Exception
                                    AccJENo = objform.Items.Item("tjeno").Specific.String
                                End Try
                                If AccJENo <> "" Then
                                    StrSQL = objglobalmethods.getSingleValue("select ""StornoToTr"" from OJDT where ""StornoToTr""='" & AccJENo & "' ")
                                    If StrSQL = "0" Or StrSQL = "" Then
                                        objaddon.objapplication.StatusBar.SetText("Auto-Generated JE is not Cancelled. Kindly Cancel the Journal Entry: " & AccJENo, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        objaddon.objapplication.MessageBox("Auto-Generated JE is not Cancelled. Kindly Cancel the Journal Entry: " & AccJENo, 1, "OK")
                                        BubbleEvent = False
                                    End If
                                End If
                                'If TempForm Then
                                '    AccJENo = objform.Items.Item("tjeno").Specific.String
                                'End If
                            End If
                    End Select
                Else
                    oUDFForm = objaddon.objapplication.Forms.Item(objform.UDFFormUID)
                    Select Case pval.MenuUID
                        Case "1284" 'Cancel
                            'If pval.InnerEvent = True Then Exit Sub
                            If objform.TypeEx = "170" Or objform.TypeEx = "426" Then
                                'If TempForm Then
                                '    'Dim ii As String = objform.Items.Item("tjeno").Specific.String
                                '    If AccJENo = "" Then Exit Sub
                                '    If objaddon.objIncPayment.Cancelling_JournalEntry(objform.UniqueID, AccJENo) Then 'objform.Items.Item("tjeno").Specific.String
                                '        AccJENo = ""
                                '        ''objaddon.objapplication.SetStatusBarMessage("JE Reversed... ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                '        TempForm = False
                                '    End If
                                'End If

                            End If
                        Case "1281" 'Find
                            If objform.TypeEx = "141" Then
                                oUDFForm.Items.Item("U_MBAPNo").Enabled = True
                            ElseIf objform.TypeEx = "170" Or objform.TypeEx = "426" Then
                                oUDFForm.Items.Item("U_JENo").Enabled = True
                                oUDFForm.Items.Item("U_Select").Enabled = True
                                objform.Items.Item("tjeno").Visible = False
                                objform.Items.Item("ljeno").Visible = False
                            ElseIf objform.TypeEx = "392" Then
                                oUDFForm.Items.Item("U_TransId").Enabled = True
                                oUDFForm.Items.Item("U_IEntry").Enabled = True
                                oUDFForm.Items.Item("U_OEntry").Enabled = True
                                oUDFForm.Items.Item("U_IntRecNo").Enabled = True
                                oUDFForm.Items.Item("U_ReconNum").Enabled = True
                                oUDFForm.Items.Item("U_IntRecEntry").Enabled = True
                            End If
                        Case "1287" 'Duplicate
                            If objform.TypeEx = "392" Then
                                If oUDFForm.Items.Item("U_TransId").Enabled = False Then oUDFForm.Items.Item("U_TransId").Enabled = True : oUDFForm.Items.Item("U_TransId").Specific.String = "" Else oUDFForm.Items.Item("U_TransId").Specific.String = ""
                                If oUDFForm.Items.Item("U_IEntry").Enabled = False Then oUDFForm.Items.Item("U_IEntry").Enabled = True : oUDFForm.Items.Item("U_IEntry").Specific.String = "" Else oUDFForm.Items.Item("U_IEntry").Specific.String = ""
                                If oUDFForm.Items.Item("U_OEntry").Enabled = False Then oUDFForm.Items.Item("U_OEntry").Enabled = True : oUDFForm.Items.Item("U_OEntry").Specific.String = "" Else oUDFForm.Items.Item("U_OEntry").Specific.String = ""
                                If oUDFForm.Items.Item("U_IntRecNo").Enabled = False Then oUDFForm.Items.Item("U_IntRecNo").Enabled = True : oUDFForm.Items.Item("U_IntRecNo").Specific.String = "" Else oUDFForm.Items.Item("U_IntRecNo").Specific.String = ""
                                If oUDFForm.Items.Item("U_ReconNum").Enabled = False Then oUDFForm.Items.Item("U_ReconNum").Enabled = True : oUDFForm.Items.Item("U_ReconNum").Specific.String = "" Else oUDFForm.Items.Item("U_ReconNum").Specific.String = ""
                            ElseIf objform.TypeEx = "141" Then
                                If oUDFForm.Items.Item("U_MBAPNo").Enabled = False Then oUDFForm.Items.Item("U_MBAPNo").Enabled = True : oUDFForm.Items.Item("U_MBAPNo").Specific.String = "" Else oUDFForm.Items.Item("U_MBAPNo").Specific.String = ""
                            End If


                        Case "1282"
                            If objform.TypeEx = "141" Then
                                oUDFForm.Items.Item("U_MBAPNo").Enabled = False
                            ElseIf objform.TypeEx = "170" Or objform.TypeEx = "426" Then
                                oUDFForm.Items.Item("U_JENo").Enabled = False
                                objform.Items.Item("tjeno").Visible = False
                                objform.Items.Item("ljeno").Visible = False
                                objform.Items.Item("chkactive").Visible = False
                                'ElseIf objform.TypeEx = "392" Then
                                '    oUDFForm.Items.Item("U_TransId").Enabled = False
                                '    oUDFForm.Items.Item("U_IEntry").Enabled = False
                                '    oUDFForm.Items.Item("U_OEntry").Enabled = False
                                '    oUDFForm.Items.Item("U_IntRecNo").Enabled = False
                            End If

                        Case Else
                            If objform.TypeEx = "141" Then
                                oUDFForm.Items.Item("U_MBAPNo").Enabled = False
                            ElseIf objform.TypeEx = "170" Or objform.TypeEx = "426" Then
                                oUDFForm.Items.Item("U_JENo").Enabled = False
                                oUDFForm.Items.Item("U_Select").Enabled = False
                                'ElseIf objform.TypeEx = "392" Then
                                '    oUDFForm.Items.Item("U_TransId").Enabled = False
                                '    oUDFForm.Items.Item("U_IEntry").Enabled = False
                                '    oUDFForm.Items.Item("U_OEntry").Enabled = False
                                '    oUDFForm.Items.Item("U_IntRecNo").Enabled = False
                            End If
                    End Select
                End If
            Catch ex As Exception
                'objaddon.objapplication.SetStatusBarMessage("Error in Standart Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub

#Region "Mul_Branch_AP_Service_Invoice"

        Private Sub Mul_Branch_AP_Service_Invoice_MenuEvent(ByRef pval As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Dim Matrix0 As SAPbouiCOM.Matrix
            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                Matrix0 = objform.Items.Item("mtxcont").Specific
                If pval.BeforeAction = True Then
                    Select Case pval.MenuUID
                        Case "1283", "1284" 'Remove & Cancel
                            objaddon.objapplication.SetStatusBarMessage("Remove or Cancel is not allowed ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        Case "1293"  'Delete Row
                    End Select
                Else
                    Dim DBSource As SAPbouiCOM.DBDataSource
                    DBSource = objform.DataSources.DBDataSources.Item("@MIPL_OAPI")
                    Select Case pval.MenuUID
                        Case "1281" 'Find Mode
                            objform.Items.Item("t_docnum").Enabled = True
                            objform.Items.Item("tposdate").Enabled = True
                            objform.Items.Item("tdocdate").Enabled = True
                            objform.Items.Item("tduedate").Enabled = True
                            objform.ActiveItem = "t_docnum"
                            objform.Items.Item("t_docnum").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            Matrix0.Item.Enabled = False
                        Case "1282" ' Add Mode
                            objform.Items.Item("tposdate").Specific.string = Now.Date.ToString("dd/MM/yy")
                            objform.Items.Item("tdocdate").Specific.string = Now.Date.ToString("dd/MM/yy")
                            objform.Items.Item("tremark").Specific.String = "Created By " & objaddon.objcompany.UserName & " on " & Now.ToString
                            objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "vcode", "#")
                            objaddon.objglobalmethods.LoadSeries(objform, DBSource, "MIAPSI")

                        Case "1288", "1289", "1290", "1291"

                        Case "1293"
                            DeleteRow(Matrix0, "@MIPL_API1")
                        Case "1292"
                            objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "vcode", "#")
                        Case "1304" 'Refresh
                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
                ' objaddon.objapplication.SetStatusBarMessage("Error in Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try
        End Sub

#End Region

#Region "Payment"

        Private Sub PaymentInit_MenuEvent(ByRef pval As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Dim Matrix0 As SAPbouiCOM.Matrix
            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                Matrix0 = objform.Items.Item("mtxdata").Specific
                If pval.BeforeAction = True Then
                    Select Case pval.MenuUID
                        Case "1283", "1284" 'Remove & Cancel
                            'objaddon.objapplication.SetStatusBarMessage("Remove or Cancel is not allowed ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            'BubbleEvent = False
                        Case "1293"  'Delete Row
                    End Select
                Else
                    Select Case pval.MenuUID
                        Case "1293"
                            Try
                                Dim USERSource As SAPbouiCOM.UserDataSource
                                USERSource = objform.DataSources.UserDataSources.Item("UD_3")
                                objform.Freeze(True)
                                For i As Integer = 1 To Matrix0.VisualRowCount
                                    Matrix0.GetLineData(i)
                                    USERSource.Value = i
                                    Matrix0.SetLineData(i)
                                Next
                                objform.Freeze(False)
                            Catch ex As Exception
                                objform.Freeze(False)
                                objaddon.objapplication.StatusBar.SetText("DeleteRow  Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
                            Finally
                            End Try

                        Case "1304" 'Refresh
                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
                ' objaddon.objapplication.SetStatusBarMessage("Error in Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try
        End Sub

        Private Sub InPayments_MenuEvent(ByRef pval As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Dim Matrix0 As SAPbouiCOM.Matrix
            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                Matrix0 = objform.Items.Item("mtxcont").Specific
                If pval.BeforeAction = True Then
                    Select Case pval.MenuUID
                        Case "1283", "1284" 'Remove & Cancel
                            objaddon.objapplication.SetStatusBarMessage("Remove or Cancel is not allowed ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        Case "1293"  'Delete Row
                    End Select
                Else
                    Select Case pval.MenuUID
                        Case "1281" 'Find Mode
                            objform.Items.Item("tdocdate").Enabled = True
                            objform.Items.Item("t_docnum").Enabled = True
                            objform.Items.Item("ttranno").Enabled = True
                            objform.ActiveItem = "t_docnum"
                            objform.Items.Item("t_docnum").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            Matrix0.Item.Enabled = False
                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
                ' objaddon.objapplication.SetStatusBarMessage("Error in Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try
        End Sub

        Private Sub OutPayments_MenuEvent(ByRef pval As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Dim Matrix0 As SAPbouiCOM.Matrix
            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                Matrix0 = objform.Items.Item("mtxcont").Specific
                If pval.BeforeAction = True Then
                    Select Case pval.MenuUID
                        Case "1283", "1284" 'Remove & Cancel
                            objaddon.objapplication.SetStatusBarMessage("Remove or Cancel is not allowed ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        Case "1293"  'Delete Row
                    End Select
                Else
                    Select Case pval.MenuUID
                        Case "1281" 'Find Mode
                            objform.Items.Item("tdocdate").Enabled = True
                            objform.Items.Item("t_docnum").Enabled = True
                            objform.Items.Item("ttranno").Enabled = True
                            objform.ActiveItem = "t_docnum"
                            objform.Items.Item("t_docnum").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            Matrix0.Item.Enabled = False
                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
                ' objaddon.objapplication.SetStatusBarMessage("Error in Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try
        End Sub

        Private Sub InternalReconciliation_MenuEvent(ByRef pval As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Dim Matrix0 As SAPbouiCOM.Matrix
            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                Matrix0 = objform.Items.Item("mtxcont").Specific
                If pval.BeforeAction = True Then
                    Select Case pval.MenuUID
                        Case "6913"
                            If objaddon.objapplication.Forms.ActiveForm.TypeEx = "FOITR" Then
                                BubbleEvent = False
                            End If
                        Case "1284" ' Cancel
                            If objaddon.objapplication.MessageBox("Cancelling of an entry cannot be reversed. Do you want to continue?", 2, "Yes", "No") <> 1 Then BubbleEvent = False : Exit Sub
                            'objaddon.objapplication.SetStatusBarMessage("Remove or Cancel is not allowed ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            'BubbleEvent = False
                            If objaddon.HANA Then
                                BranchSplitReconciliation = objaddon.objglobalmethods.getSingleValue("select ""U_BranchSplit"" from OADM")
                            Else
                                BranchSplitReconciliation = objaddon.objglobalmethods.getSingleValue("select U_BranchSplit from OADM")
                            End If
                            If BranchSplitReconciliation = "Y" Then 'With splitup JE
                                If objform.Items.Item("ttransid").Specific.String <> "" Then
                                    If Cancelling_IntBranch_RecoJournalEntry(objform.UniqueID, objform.Items.Item("ttransid").Specific.String) Then
                                        objaddon.objapplication.StatusBar.SetText("Cancelled Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                    End If
                                Else
                                    If Cancelling_IntBranch_RecoJournalEntry(objform.UniqueID, objform.Items.Item("ttransid").Specific.String, "N") Then
                                        objaddon.objapplication.StatusBar.SetText("Cancelled Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                    End If
                                End If
                            Else
                                For iRow As Integer = 1 To Matrix0.VisualRowCount
                                    If Matrix0.Columns.Item("jeno").Cells.Item(iRow).Specific.Value <> "" Then
                                        If Cancelling_IntBranch_RecoJournalEntry(objform.UniqueID, Matrix0.Columns.Item("jeno").Cells.Item(iRow).Specific.Value, "Y", iRow) Then
                                            objaddon.objapplication.StatusBar.SetText("Cancelled Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                        End If
                                    Else
                                        If Matrix0.Columns.Item("recono").Cells.Item(iRow).Specific.Value <> "" Then
                                            If Cancelling_IntBranch_RecoJournalEntry(objform.UniqueID, "", "N", iRow) Then
                                                objaddon.objapplication.StatusBar.SetText("Cancelled Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                            End If
                                        End If
                                    End If
                                Next
                            End If

                        Case "1293"  'Delete Row
                    End Select
                Else
                    Select Case pval.MenuUID
                        Case "1284"


                        Case "1281" 'Find Mode
                            objform.Items.Item("tdocdate").Enabled = True
                            objform.Items.Item("t_docnum").Enabled = True
                            objform.Items.Item("tpaydate").Enabled = True
                            objform.Items.Item("ttransid").Enabled = True
                            objform.Items.Item("ttotdue").Enabled = True
                            objform.Items.Item("trevje").Enabled = True
                            objform.Items.Item("txtentry").Enabled = True
                            objform.ActiveItem = "t_docnum"
                            objform.Items.Item("t_docnum").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            Matrix0.Item.Enabled = False
                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
                ' objaddon.objapplication.SetStatusBarMessage("Error in Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try
        End Sub

        Private Function Cancelling_IntBranch_RecoJournalEntry(ByVal FormUID As String, ByVal JETransId As String, Optional ByVal WithJE As String = "Y", Optional ByVal LineNo As Integer = 1) As Boolean
            Try
                Dim TransId As String
                Dim objmatrix As SAPbouiCOM.Matrix
                Dim objjournalentry As SAPbobsCOM.JournalEntries
                If JETransId = "" And WithJE = "Y" Then Return True
                Dim ErrorFlag As Boolean
                Dim objRs As SAPbobsCOM.Recordset
                Dim strSQL As String
                Try
                    objform = objaddon.objapplication.Forms.Item(FormUID)
                    If BranchSplitReconciliation = "Y" Then 'With splitup JE
                        objmatrix = objform.Items.Item("mtxreco").Specific
                    Else
                        objmatrix = objform.Items.Item("mtxcont").Specific
                    End If

                    objRs = objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                    If WithJE = "Y" Then
                        Dim GetStatus As String = objglobalmethods.getSingleValue("select distinct 1 as ""Status"" from OJDT where ""StornoToTr""=" & JETransId & "")
                        If GetStatus = "1" Then
                            TransId = objglobalmethods.getSingleValue("select ""TransId"" from OJDT where ""StornoToTr""=" & JETransId & "")
                            GoTo Reco
                        End If
                        strSQL = "Select T0.""Series"",T0.""TaxDate"",T0.""DueDate"",T0.""RefDate"",T0.""Ref1"",T0.""Ref2"",T0.""Memo"",T1.""Account"",T1.""Credit"",T1.""Debit"",T1.""BPLId"",T1.""U_TBranch"","
                        strSQL += vbCrLf + "(Select ""CardCode"" from OCRD where ""CardCode""=T1.""ShortName"") as ""BPCode"""
                        strSQL += vbCrLf + "from OJDT T0 join JDT1 T1 ON T0.""TransId""=T1.""TransId"" where  T1.""TransId""='" & JETransId & "' order by T1.""Line_ID"""
                        objRs.DoQuery(strSQL)
                        If objRs.RecordCount = 0 Then Return True
                        If Not objaddon.objcompany.InTransaction Then objaddon.objcompany.StartTransaction()
                        objjournalentry = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
                        objaddon.objapplication.StatusBar.SetText("Journal Entry Reversing Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        objjournalentry.TaxDate = objRs.Fields.Item("TaxDate").Value.ToString 'objJEHeader.GetValue("TaxDate", 0)
                        objjournalentry.DueDate = objRs.Fields.Item("DueDate").Value.ToString 'objJEHeader.GetValue("DueDate", 0)
                        objjournalentry.ReferenceDate = objRs.Fields.Item("RefDate").Value.ToString 'objJEHeader.GetValue("RefDate", 0)
                        objjournalentry.Reference = objRs.Fields.Item("Ref1").Value.ToString 'objJEHeader.GetValue("Ref1", 0)
                        objjournalentry.Reference2 = objRs.Fields.Item("Ref2").Value.ToString 'objJEHeader.GetValue("Ref2", 0)
                        objjournalentry.Reference3 = Now.ToString
                        objjournalentry.Memo = objRs.Fields.Item("Memo").Value.ToString & "(Reversal) - " & Trim(JETransId) 'objJEHeader.GetValue("Memo", 0) & " (Reversal) - " & Trim(JETransId)
                        objjournalentry.Series = objRs.Fields.Item("Series").Value.ToString 'objJEHeader.GetValue("Series", 0)
                        objjournalentry.UserFields.Fields.Item("U_IntRecNo").Value = objRs.Fields.Item("Ref2").Value.ToString
                        For AccRow As Integer = 0 To objRs.RecordCount - 1
                            If objRs.Fields.Item("BPCode").Value.ToString <> "" Then objjournalentry.Lines.ShortName = objRs.Fields.Item("BPCode").Value.ToString Else objjournalentry.Lines.AccountCode = objRs.Fields.Item("Account").Value.ToString
                            'objjournalentry.Lines.AccountCode = objRs.Fields.Item("Account").Value.ToString
                            If CDbl(objRs.Fields.Item("Credit").Value.ToString) <> 0 Then objjournalentry.Lines.Debit = CDbl(objRs.Fields.Item("Credit").Value.ToString) Else objjournalentry.Lines.Credit = CDbl(objRs.Fields.Item("Debit").Value.ToString)
                            objjournalentry.Lines.BPLID = objRs.Fields.Item("BPLId").Value.ToString
                            objjournalentry.Lines.UserFields.Fields.Item("U_TBranch").Value = objRs.Fields.Item("U_TBranch").Value.ToString 'Branch
                            objjournalentry.Lines.Add()
                            objRs.MoveNext()
                        Next

                        If objjournalentry.Add <> 0 Then
                            ErrorFlag = True
                            If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            objaddon.objapplication.SetStatusBarMessage("Journal Reverse: " & objaddon.objcompany.GetLastErrorDescription & "-" & objaddon.objcompany.GetLastErrorCode, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(objjournalentry)
                            'Return False
                        Else
                            If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                            TransId = objaddon.objcompany.GetNewObjectKey()
Reco:
                            If BranchSplitReconciliation = "Y" Then 'With splitup JE
                                objform.Items.Item("trevje").Specific.String = TransId
                            Else
                                objmatrix.Columns.Item("revjeno").Cells.Item(LineNo).Specific.String = TransId
                                strSQL = "Update ""@MI_ITR1"" Set ""U_RevJENo""=" & TransId & " Where ""DocEntry""='" & objform.Items.Item("txtentry").Specific.String & "' and ""LineId""='" & LineNo & "' "
                                objRs.DoQuery(strSQL)
                            End If

                            objRs.DoQuery("Update OJDT set ""StornoToTr""=" & JETransId & " where ""TransId""=" & TransId & "")
                            For i As Integer = 1 To objmatrix.VisualRowCount
                                If objmatrix.Columns.Item("recono").Cells.Item(i).Specific.String = "" Then Continue For
                                Dim RecoNum As Integer = CInt(objmatrix.Columns.Item("recono").Cells.Item(i).Specific.String) ' IIf(CInt(objmatrix.Columns.Item("recono").Cells.Item(i).Specific.String) = 0, 0, CInt(objmatrix.Columns.Item("recono").Cells.Item(i).Specific.String))
                                If RecoNum <> 0 Then
                                    If Cancelling_IntBranch_ManualReconciliation(RecoNum) = False Then
                                        ErrorFlag = True
                                    End If
                                    objRs.DoQuery("Update OITR Set ""U_RevTransId""='" & TransId & "' where ""ReconNum""='" & RecoNum & "'")
                                End If
                            Next
                            'Return True
                        End If
                    Else
                        Try
                            For i As Integer = 1 To objmatrix.VisualRowCount
                                If objmatrix.Columns.Item("recono").Cells.Item(i).Specific.String = "" Then Continue For
                                Dim RecoNum As Integer = CInt(objmatrix.Columns.Item("recono").Cells.Item(i).Specific.String) ' IIf(CInt(objmatrix.Columns.Item("recono").Cells.Item(i).Specific.String) = 0, 0, CInt(objmatrix.Columns.Item("recono").Cells.Item(i).Specific.String))
                                If RecoNum <> 0 Then
                                    If Cancelling_IntBranch_ManualReconciliation(RecoNum) = False Then
                                        ErrorFlag = True
                                    End If
                                End If
                            Next
                        Catch ex As Exception
                        End Try

                        'Return True
                    End If
                    If ErrorFlag Then
                        'If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        If BranchSplitReconciliation = "Y" Then 'With splitup JE 
                            objform.Items.Item("trevje").Specific.String = ""
                        Else
                            For i As Integer = 1 To objmatrix.VisualRowCount
                                objmatrix.Columns.Item("revjeno").Cells.Item(i).Specific.String = ""
                            Next
                        End If

                    Else
                        'If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        objform.Items.Item("cmbstat").Specific.Select("C", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        'objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        If BranchSplitReconciliation = "Y" Then 'With splitup JE 
                            objform.Items.Item("1").Click()
                            objform.Items.Item("trevje").Visible = True
                            objform.Items.Item("lrevje").Visible = True
                            objform.Items.Item("lnkrevje").Visible = True
                        Else
                            objmatrix.Columns.Item("revjeno").Visible = True
                        End If

                        objaddon.objapplication.StatusBar.SetText("Transactions Cancelled Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        Return True
                    End If
                Catch ex As Exception
                    If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    objaddon.objapplication.SetStatusBarMessage("Transaction Cancelling Error " & objaddon.objcompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                    Return False
                End Try
                objRs = Nothing
            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage("Transaction Cancelling Error: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                Return False
            Finally

                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End Function

        Private Function Cancelling_IntBranch_ManualReconciliation(ByVal RecoNum As Integer) As Boolean
            Try
                Dim service As IInternalReconciliationsService = objaddon.objcompany.GetCompanyService().GetBusinessService(ServiceTypes.InternalReconciliationsService)
                Dim reconParams As IInternalReconciliationParams = service.GetDataInterface(InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationParams)
                reconParams.ReconNum = RecoNum
                Dim Status As String = ""
                If objaddon.HANA Then
                    Status = objglobalmethods.getSingleValue("select 1 as ""Status"" from OITR where ""ReconNum""=" & RecoNum & " and ""IsSystem""='N' and ""Canceled""='Y'")
                Else
                    Status = objglobalmethods.getSingleValue("select 1 as Status from OITR where ReconNum=" & RecoNum & " and IsSystem='N' and Canceled='Y'")
                End If
                If Status = "1" Then Return True
                Try
                    service.Cancel(reconParams)
                Catch ex As Exception
                    'Return False
                End Try
                Return True
            Catch ex As Exception
                GC.Collect()
            End Try
        End Function

        Private Sub Payment_Means_MenuEvent(ByRef pval As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Dim Matrix0 As SAPbouiCOM.Matrix
            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                Matrix0 = objform.Items.Item("mtxcheq").Specific
                If pval.BeforeAction = True Then
                    Select Case pval.MenuUID
                        Case "CPYD"  'Copy Due
                            If objform.Items.Item("tcurr").Specific.selected.Value = MainCurr Then
                                If objform.ActiveItem = "tctot" Then
                                    objform.Items.Item("tctot").Specific.String = objform.Items.Item("tbaldue").Specific.String
                                ElseIf objform.ActiveItem = "tbtot" Then
                                    objform.Items.Item("tbtot").Specific.String = objform.Items.Item("tbaldue").Specific.String
                                Else
                                    Dim ColID As Integer = Matrix0.GetCellFocus().ColumnIndex
                                    Dim RowID As Integer = Matrix0.GetCellFocus().rowIndex
                                    If ColID = 2 Then 'chamt
                                        Matrix0.Columns.Item("chamt").Cells.Item(RowID).Specific.String = CDbl(objform.Items.Item("tbaldue").Specific.String) '+ Val(Matrix0.Columns.Item("chamt").ColumnSetting.SumValue)
                                    End If
                                End If
                            Else
                                If objform.ActiveItem = "tctot" Then
                                    objform.Items.Item("tctot").Specific.String = objform.Items.Item("tbalduec").Specific.String
                                ElseIf objform.ActiveItem = "tbtot" Then
                                    objform.Items.Item("tbtot").Specific.String = objform.Items.Item("tbalduec").Specific.String
                                Else
                                    Dim ColID As Integer = Matrix0.GetCellFocus().ColumnIndex
                                    Dim RowID As Integer = Matrix0.GetCellFocus().rowIndex
                                    If ColID = 2 Then 'chamt
                                        Matrix0.Columns.Item("chamt").Cells.Item(RowID).Specific.String = CDbl(objform.Items.Item("tbalduec").Specific.String) '+ Val(Matrix0.Columns.Item("chamt").ColumnSetting.SumValue)
                                    End If
                                End If
                            End If
                    End Select
                Else
                    Select Case pval.MenuUID
                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
                ' objaddon.objapplication.SetStatusBarMessage("Error in Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try
        End Sub

#End Region

        Sub DeleteRow(ByVal objMatrix As SAPbouiCOM.Matrix, ByVal TableName As String)
            Try
                Dim DBSource As SAPbouiCOM.DBDataSource
                'objMatrix = objform.Items.Item("20").Specific
                objMatrix.FlushToDataSource()
                DBSource = objform.DataSources.DBDataSources.Item(TableName) '"@MIREJDET1"
                For i As Integer = 1 To objMatrix.VisualRowCount
                    objMatrix.GetLineData(i)
                    DBSource.Offset = i - 1
                    DBSource.SetValue("LineId", DBSource.Offset, i)
                    objMatrix.SetLineData(i)
                    objMatrix.FlushToDataSource()
                Next
                DBSource.RemoveRecord(DBSource.Size - 1)
                objMatrix.LoadFromDataSource()

            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("DeleteRow  Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
            Finally
            End Try
        End Sub
    End Class
End Namespace