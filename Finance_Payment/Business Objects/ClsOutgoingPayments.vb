Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports SAPbobsCOM
Imports SAPbouiCOM.Framework
Namespace Finance_Payment
    Public Class ClsOutgoingPayments
        Public Const Formtype = "426"
        Dim objform As SAPbouiCOM.Form
        Dim objmatrix As SAPbouiCOM.Matrix
        Dim strSQL As String
        Dim objRs As SAPbobsCOM.Recordset
        Dim cmbbranch As SAPbouiCOM.Column

        Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.Item(FormUID)
                objmatrix = objform.Items.Item("71").Specific
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If pVal.BeforeAction Then
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT, SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                            If pVal.ItemUID = "71" And (pVal.ColUID = "U_Branch" Or pVal.ColUID = "U_AcctCode" Or pVal.ColUID = "U_AcctName" Or pVal.ColUID = "U_Source") Then
                                'If objform.Items.Item("chkactive").Specific.checked = False Then Exit Sub
                                If objaddon.HANA Then
                                    objRs.DoQuery("select distinct ""PmtClrAct"" from OBPL where ""Disabled"" ='N' ")
                                Else
                                    objRs.DoQuery("select distinct PmtClrAct from OBPL where Disabled ='N'")
                                End If
                                Dim SuccessFlag As Boolean = False
                                If objRs.RecordCount > 0 Then
                                    For i As Integer = 0 To objRs.RecordCount - 1
                                        If objmatrix.Columns.Item("8").Cells.Item(pVal.Row).Specific.String = objRs.Fields.Item(0).Value.ToString Then
                                            SuccessFlag = True
                                        End If
                                        objRs.MoveNext()
                                    Next
                                End If
                                If SuccessFlag = False Then BubbleEvent = False
                                Dim ColItem As SAPbouiCOM.Column = objmatrix.Columns.Item("U_AcctCode")
                                ColItem.ChooseFromListUID = "cflacctcode"
                                ColItem.ChooseFromListAlias = "AcctCode"
                                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                                    If objmatrix.Columns.Item("U_Source").Cells.Item(pVal.Row).Specific.Selected.Value = "1" Then 'Business Partner
                                        ColItem.ChooseFromListUID.Remove(pVal.Row)
                                        ColItem.ChooseFromListUID = "cflcardcode"
                                        ColItem.ChooseFromListAlias = "CardCode"
                                        AddCFLCondition(FormUID, "cflcardcode", "validFor", "Y", "", "")
                                    Else 'G/L Account
                                        ColItem.ChooseFromListUID.Remove(pVal.Row)
                                        ColItem.ChooseFromListUID = "cflacctcode"
                                        ColItem.ChooseFromListAlias = "AcctCode"
                                        AddCFLCondition(FormUID, "cflacctcode", "Postable", "Y", "LocManTran", "N")
                                    End If
                                End If
                            ElseIf pVal.ItemUID = "71" And (pVal.ColUID = "8") Then
                                If objform.Items.Item("chkactive").Specific.checked = False Then Exit Sub
                                If pVal.Row = 1 Then Exit Sub
                                If objaddon.HANA Then
                                    objRs.DoQuery("select distinct ""PmtClrAct"" from OBPL where ""Disabled"" ='N' ")
                                Else
                                    objRs.DoQuery("select distinct PmtClrAct from OBPL where Disabled ='N'")
                                End If
                                If objRs.RecordCount > 0 Then
                                    For i As Integer = 0 To objRs.RecordCount - 1
                                        If Validate_Row(FormUID, pVal.Row - 1, objRs.Fields.Item(0).Value.ToString) Then
                                            BubbleEvent = False
                                        End If
                                        objRs.MoveNext()
                                    Next
                                End If
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                            If pVal.ItemUID = "71" And (pVal.ColUID = "U_AcctName") Then
                                BubbleEvent = False
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                            If pVal.InnerEvent = True Then Exit Sub
                            'If objform.Items.Item("chkactive").Specific.checked = False Then Exit Sub
                            If (pVal.ItemUID = "71" And pVal.ColUID = "U_AcctCode") Then
                                If objmatrix.Columns.Item("U_Source").Cells.Item(pVal.Row).Specific.Selected.Value = "1" Then
                                    If objaddon.HANA Then
                                        objRs.DoQuery("Select  ""CardName"" from OCRD where ""CardCode""='" & objmatrix.Columns.Item("U_AcctCode").Cells.Item(pVal.Row).Specific.String & "' ")
                                    Else
                                        objRs.DoQuery("Select CardName from OCRD where CardCode='" & objmatrix.Columns.Item("U_AcctCode").Cells.Item(pVal.Row).Specific.String & "'")
                                    End If
                                Else
                                    If objaddon.HANA Then
                                        objRs.DoQuery("Select  ""AcctName"" from OACT where ""AcctCode""='" & objmatrix.Columns.Item("U_AcctCode").Cells.Item(pVal.Row).Specific.String & "' ")
                                    Else
                                        objRs.DoQuery("Select AcctName from OACT where AcctCode='" & objmatrix.Columns.Item("U_AcctCode").Cells.Item(pVal.Row).Specific.String & "'")
                                    End If
                                End If
                                If objRs.RecordCount > 0 Then objmatrix.Columns.Item("U_AcctName").Cells.Item(pVal.Row).Specific.String = objRs.Fields.Item(0).Value.ToString : objmatrix.AutoResizeColumns()
                                'ElseIf pVal.ItemUID = "71" And (pVal.ColUID = "8" Or pVal.ColUID = "U_AcctCode") Then
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_CLICK
                            If pVal.ItemUID = "1" And objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                If objform.Items.Item("58").Specific.Selected = False Then Exit Sub
                                'If objform.Items.Item("chkactive").Specific.checked = False Then Exit Sub
                                If objform.Items.Item("1320002037").Specific.Selected Is Nothing Then objaddon.objapplication.StatusBar.SetText("Please update header level branch...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
                                If objaddon.HANA Then
                                    objRs.DoQuery("select distinct ""PmtClrAct"" from OBPL where ""Disabled"" ='N' ")
                                Else
                                    objRs.DoQuery("select distinct PmtClrAct from OBPL where Disabled ='N'")
                                End If
                                If objRs.RecordCount > 0 Then
                                    For i As Integer = 0 To objRs.RecordCount - 1
                                        Dim GLC As String = objRs.Fields.Item(0).Value.ToString
                                        If Validate_Row(FormUID, -1, GLC) Then
                                            BubbleEvent = False
                                        End If
                                        If ValidTransaction(FormUID, GLC) Then
                                            If objform.Items.Item("chkactive").Specific.checked = False Then
                                                objaddon.objapplication.MessageBox("Select the Multi Branch Selection...", , "OK")
                                                objaddon.objapplication.StatusBar.SetText("Select the Multi Branch Selection...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                BubbleEvent = False : Exit Sub
                                            End If
                                        End If
                                        objRs.MoveNext()
                                    Next
                                End If
                                If ValidateAccountLines(FormUID) = True Then
                                    objaddon.objapplication.MessageBox("Header & line level branch should be different for branch control account...", , "OK")
                                    objaddon.objapplication.StatusBar.SetText("Header & line level branch should be different for branch control account...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                End If

                            End If
                        Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                            If (pVal.ItemUID = "56" Or pVal.ItemUID = "57") Then
                                If objform.Items.Item("chkactive").Specific.checked = True Then If objform.Items.Item("tjeno").Specific.String = "" Then BubbleEvent = False
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            'Dim ii As Integer = objaddon.objapplication.FontHeight
                            'Dim aa As Integer = objform.Items.Item("145").Height
                            'MsgBox(aa)
                            'objform.Items.Item("10002011").Enabled = False
                    End Select
                Else
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                            If pVal.ItemUID = "71" And pVal.ColUID = "U_AcctCode" Then
                                Dim objCFLEvent As SAPbouiCOM.ChooseFromListEvent
                                Dim objDataTable As SAPbouiCOM.DataTable
                                objCFLEvent = pVal
                                objDataTable = objCFLEvent.SelectedObjects
                                If Not objDataTable Is Nothing Then
                                    'objmatrix.Columns.Item("U_AcctCode").Cells.Item(pVal.Row).Specific.String = objDataTable.GetValue("AcctCode", 0)
                                    If objmatrix.Columns.Item("U_Source").Cells.Item(pVal.Row).Specific.Selected.Value = "1" Then
                                        objmatrix.SetCellWithoutValidation(pVal.Row, "U_AcctCode", objDataTable.GetValue("CardCode", 0))
                                    Else
                                        objmatrix.SetCellWithoutValidation(pVal.Row, "U_AcctCode", objDataTable.GetValue("AcctCode", 0))
                                    End If
                                End If
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                            If pVal.ColUID = "U_Source" Then
                                Dim ColItem As SAPbouiCOM.Column = objmatrix.Columns.Item("U_AcctCode")
                                If pVal.ItemChanged = True Then objmatrix.SetCellWithoutValidation(pVal.Row, "U_AcctCode", "") : objmatrix.SetCellWithoutValidation(pVal.Row, "U_AcctName", "")
                                If objmatrix.Columns.Item("U_Source").Cells.Item(pVal.Row).Specific.Selected.Value = "1" Then 'Business Partner
                                    ColItem.ChooseFromListUID.Remove(pVal.Row)
                                    ColItem.ChooseFromListUID = "cflcardcode"
                                    ColItem.ChooseFromListAlias = "CardCode"
                                Else 'G/L Account
                                    ColItem.ChooseFromListUID.Remove(pVal.Row)
                                    ColItem.ChooseFromListUID = "cflacctcode"
                                    ColItem.ChooseFromListAlias = "AcctCode"
                                End If
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                            Try
                                objform.Items.Item("ljeno").Left = objform.Items.Item("53").Left
                                objform.Items.Item("tjeno").Left = objform.Items.Item("52").Left
                                objform.Items.Item("lnkjeno").Left = objform.Items.Item("54").Left
                            Catch ex As Exception
                            End Try
                        Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                            If pVal.ItemUID = "71" And (pVal.ColUID = "U_AcctCode" Or pVal.ColUID = "U_AcctName" Or pVal.ColUID = "U_Branch") Then
                                objform.Freeze(True) : objmatrix.AutoResizeColumns() : objform.Freeze(False)
                            End If

                        Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                            If pVal.ItemUID = "58" And pVal.ActionSuccess Then
                                CreateButton(FormUID)
                                AddChooseFromList("2", "cflcardcode")
                                AddChooseFromList("1", "cflacctcode")
                            ElseIf (pVal.ItemUID = "56" Or pVal.ItemUID = "57" Or pVal.ItemUID = "1" Or pVal.ItemUID = "10002011") And pVal.ActionSuccess Then
                                ItemHide(FormUID, "ljeno")
                                ItemHide(FormUID, "tjeno")
                                ItemHide(FormUID, "chkactive")
                            End If
                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
            End Try
        End Sub

        Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.Item(BusinessObjectInfo.FormUID)
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                If BusinessObjectInfo.BeforeAction = True Then
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                            If objform.Items.Item("58").Specific.Selected = False Then Exit Sub
                            If objform.Items.Item("chkactive").Specific.checked = False Then Exit Sub
                            If objform.Items.Item("1320002037").Specific.Selected Is Nothing Then objaddon.objapplication.StatusBar.SetText("Select the header level branch...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
                            If objaddon.HANA Then
                                objRs.DoQuery("select distinct ""PmtClrAct"" from OBPL where ""Disabled"" ='N' ")
                            Else
                                objRs.DoQuery("select distinct PmtClrAct from OBPL where Disabled ='N'")
                            End If
                            Dim TranFlag As Boolean = False
                            If objRs.RecordCount > 0 Then
                                For i As Integer = 0 To objRs.RecordCount - 1
                                    If ValidTransaction(objform.UniqueID, objRs.Fields.Item(0).Value.ToString) Then
                                        TranFlag = True
                                    End If
                                    objRs.MoveNext()
                                Next
                            End If
                            If TranFlag = False Then Exit Sub
                            If JournalEntry(objform.UniqueID, objform.Items.Item("1320002037").Specific.Selected.Value) = False Then
                                objaddon.objapplication.MessageBox("Error occurred while Creating the Journal Transaction...", , "OK")
                                objaddon.objapplication.StatusBar.SetText("Error occurred while Creating the Journal Transaction...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                    End Select
                Else
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                            Dim DocEntry As String
                            If objform.Items.Item("58").Specific.Selected = False Then Exit Sub
                            If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And BusinessObjectInfo.ActionSuccess = True And BusinessObjectInfo.BeforeAction = False Then
                                DocEntry = objform.DataSources.DBDataSources.Item("OVPM").GetValue("DocEntry", 0)
                                If objform.Items.Item("tjeno").Specific.String <> "" Then
                                    objRs.DoQuery("update OJDT set ""U_TransId""='" & objform.Items.Item("tjeno").Specific.String & "' where ""TransId""=(select ""TransId"" from OVPM where ""DocEntry""='" & DocEntry & "')")
                                    objRs.DoQuery("update OJDT set ""U_OEntry""='" & DocEntry & "' where ""TransId""='" & objform.Items.Item("tjeno").Specific.String & "'")
                                    DocEntry = objform.DataSources.DBDataSources.Item("OVPM").GetValue("TransId", 0)
                                    If DocEntry <> "" Then Update_Standard_JE_TranBranch(objform.UniqueID, DocEntry)
                                End If
                                'objform.Items.Item("52").Specific.String
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                            'Dim TEntry As String
                            'TEntry = objform.DataSources.DBDataSources.Item("OVPM").GetValue("TransId", 0)
                            'If TEntry <> "" Then Update_Standard_JE_TranBranch(objform.UniqueID, TEntry)

                            If objform.Items.Item("58").Specific.Selected = True Then
                                CreateButton(objform.UniqueID)
                                If ItemExists(objform.UniqueID, "tjeno") = True Then objform.Items.Item("tjeno").Enabled = False : objform.Items.Item("tjeno").Visible = True
                                If ItemExists(objform.UniqueID, "ljeno") = True Then objform.Items.Item("ljeno").Visible = True
                                If ItemExists(objform.UniqueID, "chkactive") = True Then objform.Items.Item("chkactive").Enabled = False : objform.Items.Item("chkactive").Visible = True
                            Else
                                ItemHide(objform.UniqueID, "ljeno")
                                ItemHide(objform.UniqueID, "tjeno")
                                ItemHide(objform.UniqueID, "chkactive")
                            End If
                            cmbbranch = objmatrix.Columns.Item("U_Branch")
                            cmbbranch.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                            cmbbranch = objmatrix.Columns.Item("U_Source")
                            cmbbranch.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                    End Select
                End If

            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        End Sub

        Public Sub CreateButton(ByVal FormUID As String)
            Try
                Dim objItem As SAPbouiCOM.Item
                Dim objLabel As SAPbouiCOM.StaticText
                Dim objCheck As SAPbouiCOM.CheckBox
                Dim objedit As SAPbouiCOM.EditText
                Dim objlink As SAPbouiCOM.LinkedButton
                objform = objaddon.objapplication.Forms.Item(FormUID)
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                If ItemExists(FormUID, "ljeno") = False Then
                    objItem = objform.Items.Add("ljeno", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                    objItem.Left = objform.Items.Item("53").Left
                    objItem.Width = objform.Items.Item("53").Width '80
                    objItem.Top = objform.Items.Item("53").Top + objform.Items.Item("53").Height + 2
                    objItem.Height = objform.Items.Item("53").Height ' 14
                    objLabel = objItem.Specific
                    objLabel.Caption = "JE No."
                End If
                If ItemExists(FormUID, "tjeno") = False Then
                    objItem = objform.Items.Add("tjeno", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                    objItem.Left = objform.Items.Item("52").Left '+ objform.Items.Item("52").Width + 10
                    objItem.Width = objform.Items.Item("52").Width
                    objItem.Top = objform.Items.Item("ljeno").Top
                    objItem.Height = objform.Items.Item("52").Height 'ljeno
                    objItem.LinkTo = "ljeno"
                    objedit = objItem.Specific
                    objedit.Item.Enabled = False
                    objedit.DataBind.SetBound(True, "OVPM", "U_JENo")
                    objItem.Enabled = False
                End If
                If ItemExists(FormUID, "lnkjeno") = False Then
                    objItem = objform.Items.Add("lnkjeno", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
                    objItem.Left = objform.Items.Item("53").Left + objform.Items.Item("53").Width + 4
                    objItem.Width = 12
                    objItem.Top = objform.Items.Item("53").Top + objform.Items.Item("53").Height + 2
                    objItem.Height = objform.Items.Item("54").Height '10
                    objlink = objItem.Specific
                    objlink.LinkedObjectType = "30"
                    objlink.Item.LinkTo = "tjeno"
                End If
                If ItemExists(FormUID, "chkactive") = False Then
                    objItem = objform.Items.Add("chkactive", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
                    objItem.Left = objform.Items.Item("73").Left
                    Dim Fieldsize As Size = TextRenderer.MeasureText("Multi Branch Selection", New Font("Arial", 12.0F))
                    objItem.Width = Fieldsize.Width '140
                    objItem.Top = objform.Items.Item("73").Top + objform.Items.Item("73").Height + 2
                    objItem.Height = objaddon.objapplication.FontHeight + 6 ' 16
                    objCheck = objItem.Specific
                    objCheck.Caption = "Multi Branch Selection"
                    objCheck.DataBind.SetBound(True, "OVPM", "U_Select")
                End If

                Dim objRs As SAPbobsCOM.Recordset
                cmbbranch = objmatrix.Columns.Item("U_Branch")
                If Not cmbbranch.ValidValues.Count <= 1 Then Exit Sub

                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If objaddon.HANA Then
                    objRs.DoQuery("select  ""BPLId"",""BPLName"" from OBPL where ""Disabled""='N' ")
                Else
                    objRs.DoQuery("select  BPLId,BPLName from OBPL where Disabled='N' ")
                End If
                If objRs.RecordCount > 0 Then
                    For i As Integer = 0 To objRs.RecordCount - 1
                        cmbbranch.ValidValues.Add(objRs.Fields.Item(0).Value.ToString, objRs.Fields.Item(1).Value.ToString)
                        objRs.MoveNext()
                    Next
                    cmbbranch.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                End If
                'objaddon.objapplication.SetStatusBarMessage("Button Created", SAPbouiCOM.BoMessageTime.bmt_Short, False)
            Catch ex As Exception
            End Try

        End Sub

        Private Function ItemExists(ByVal FormUID As String, ByVal ItemUID As String) As Boolean
            Try
                Dim Flag As Boolean = False
                objform = objaddon.objapplication.Forms.Item(FormUID)
                If objform.Items.Item(ItemUID).UniqueID = ItemUID Then
                    If objform.Items.Item(ItemUID).Visible = False Then
                        objform.Items.Item(ItemUID).Visible = True
                        If ItemUID = "tjeno" Then objform.Items.Item(ItemUID).Enabled = False
                        If ItemUID = "chkactive" Then If objform.Items.Item(ItemUID).Specific.Checked = True Then objform.Items.Item(ItemUID).Specific.Checked = True
                    End If
                    Flag = True
                End If
                Return Flag
            Catch ex As Exception
                Return False
            End Try
        End Function

        Private Function ItemHide(ByVal FormUID As String, ByVal ItemUID As String) As Boolean
            Try
                Dim Flag As Boolean = False
                objform = objaddon.objapplication.Forms.Item(FormUID)
                If objform.Items.Item(ItemUID).UniqueID = ItemUID Then
                    If objform.Items.Item(ItemUID).Visible = True Then
                        objform.Items.Item(ItemUID).Visible = False
                    End If
                    Flag = True
                End If
                Return Flag
            Catch ex As Exception
                Return False
            End Try
        End Function

        Private Sub AddChooseFromList(ByVal ObjectType As String, ByVal CFLID As String)
            Try

                Dim oCFLs As SAPbouiCOM.ChooseFromListCollection

                oCFLs = objform.ChooseFromLists
                'If oCFLs.Item(CFLID).UniqueID = CFLID Then Exit Sub
                Dim oCFL As SAPbouiCOM.ChooseFromList
                Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
                oCFLCreationParams = objaddon.objapplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

                ' Adding 2 CFL, one for the button and one for the edit text.
                oCFLCreationParams.MultiSelection = False
                oCFLCreationParams.ObjectType = ObjectType ' "2"
                oCFLCreationParams.UniqueID = CFLID ' "CFL1"

                oCFL = oCFLs.Add(oCFLCreationParams)

            Catch
                'MsgBox(Err.Description)
            End Try
        End Sub

        Private Sub AddCFLCondition(ByVal FormUID As String, ByVal CFLUID As String, ByVal AliasCol1 As String, ByVal AliasVal1 As String, ByVal AliasCo2 As String, ByVal AliasVal2 As String)
            Try
                objform = objaddon.objapplication.Forms.Item(FormUID)
                Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item(CFLUID) '"EMP_1"
                Dim oConds As SAPbouiCOM.Conditions
                Dim oCond As SAPbouiCOM.Condition
                Dim oEmptyConds As New SAPbouiCOM.Conditions
                Dim rsetCFL As SAPbobsCOM.Recordset = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oCFL.SetConditions(oEmptyConds)
                oConds = oCFL.GetConditions()

                oCond = oConds.Add()
                oCond.Alias = AliasCol1 ' "Postable"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = AliasVal1 '"Y"

                If AliasCo2 <> "" Then
                    oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                    oCond = oConds.Add()
                    oCond.Alias = AliasCo2 ' "LocManTran"
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCond.CondVal = AliasVal2 ' "N"
                End If

                oCFL.SetConditions(oConds)
            Catch ex As Exception

            End Try
        End Sub

        Private Function JournalEntry(ByVal FormUID As String, ByVal Branch As String) As Boolean
            Try
                Dim TransId, Series As String
                Dim objjournalentry As SAPbobsCOM.JournalEntries
                Dim JEAmount As String
                Dim oEdit As SAPbouiCOM.EditText
                Dim DocDate As Date
                Try
                    objform = objaddon.objapplication.Forms.Item(FormUID)
                    objmatrix = objform.Items.Item("71").Specific
                    If objform.Items.Item("tjeno").Specific.String = "" Then
                        objjournalentry = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
                        objaddon.objapplication.StatusBar.SetText("Journal Entry Creating Please wait...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        If Not objaddon.objcompany.InTransaction Then objaddon.objcompany.StartTransaction()

                        oEdit = objform.Items.Item("10").Specific 'Posting Date
                        DocDate = Date.ParseExact(oEdit.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                        objjournalentry.ReferenceDate = DocDate ' Posting Date
                        oEdit = objform.Items.Item("121").Specific 'Due Date
                        DocDate = Date.ParseExact(oEdit.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                        objjournalentry.DueDate = DocDate   'Due Date
                        oEdit = objform.Items.Item("90").Specific 'Tax Date
                        DocDate = Date.ParseExact(oEdit.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                        objjournalentry.TaxDate = DocDate   'Document Date

                        objjournalentry.Reference = "Out AccountPay JE"
                        objjournalentry.Reference2 = "Out AccPay On: " & Now.ToString
                        objjournalentry.Memo = objform.Items.Item("59").Specific.String
                        If Localization = "IN" Then
                            If objaddon.HANA Then
                                Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 ""Series"" FROM NNM1 WHERE ""ObjectCode""='30' and ""Indicator""=(Select ""Indicator"" from OFPR where '" & DocDate.ToString("yyyyMMdd") & "' Between ""F_RefDate"" and ""T_RefDate"") " &
                                                                                  " and Ifnull(""Locked"",'')='N' and ""BPLId""='" & Branch & "'")
                            Else
                                Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 Series FROM NNM1 WHERE ObjectCode='30' and Indicator=(Select Indicator from OFPR where '" & DocDate.ToString("yyyyMMdd") & "' Between F_RefDate and T_RefDate) " &
                                                                                  " and Isnull(Locked,'')='N' and BPLId='" & Branch & "'")
                            End If
                        Else
                            objjournalentry.AutoVAT = BoYesNoEnum.tNO
                            objjournalentry.AutomaticWT = BoYesNoEnum.tNO
                            If objaddon.HANA Then
                                Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 ""Series"" FROM NNM1 WHERE ""ObjectCode""='30' and ""Indicator""=(Select ""Indicator"" from OFPR where '" & DocDate.ToString("yyyyMMdd") & "' Between ""F_RefDate"" and ""T_RefDate"") " &
                                                                                  " and Ifnull(""Locked"",'')='N' and ""BPLId""='" & Branch & "'")
                            Else
                                Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 Series FROM NNM1 WHERE ObjectCode='30' and Indicator=(Select Indicator from OFPR where '" & DocDate.ToString("yyyyMMdd") & "' Between F_RefDate and T_RefDate) " &
                                                                                  " and Isnull(Locked,'')='N' and BPLId='" & Branch & "'")
                            End If
                        End If
                        If Series <> "" Then objjournalentry.Series = Series
                        'objjournalentry.UserFields.Fields.Item("U_ITranId").Value = ""
                        For AccRow As Integer = 1 To objmatrix.VisualRowCount
                            If objmatrix.Columns.Item("U_AcctCode").Cells.Item(AccRow).Specific.String <> "" And Not objmatrix.Columns.Item("U_Branch").Cells.Item(AccRow).Specific.Selected Is Nothing Then
                                If Branch <> objmatrix.Columns.Item("U_Branch").Cells.Item(AccRow).Specific.Selected.Value Then
                                    JEAmount = objmatrix.Columns.Item("5").Cells.Item(AccRow).Specific.String.ToString.Remove(0, 4)
                                    objjournalentry.Lines.AccountCode = objmatrix.Columns.Item("8").Cells.Item(AccRow).Specific.String
                                    objjournalentry.Lines.Credit = JEAmount
                                    objjournalentry.Lines.BPLID = objmatrix.Columns.Item("U_Branch").Cells.Item(AccRow).Specific.Selected.Value
                                    objjournalentry.Lines.UserFields.Fields.Item("U_TBranch").Value = objform.Items.Item("1320002037").Specific.Selected.Description 'Branch
                                    'If InDT.Rows(Row)("cc1").ToString <> "" Then objjournalentry.Lines.CostingCode = InDT.Rows(Row)("cc1").ToString
                                    'If InDT.Rows(Row)("cc2").ToString <> "" Then objjournalentry.Lines.CostingCode2 = InDT.Rows(Row)("cc2").ToString
                                    'If InDT.Rows(Row)("cc3").ToString <> "" Then objjournalentry.Lines.CostingCode3 = InDT.Rows(Row)("cc3").ToString
                                    'If InDT.Rows(Row)("cc4").ToString <> "" Then objjournalentry.Lines.CostingCode4 = InDT.Rows(Row)("cc4").ToString
                                    'If InDT.Rows(Row)("cc5").ToString <> "" Then objjournalentry.Lines.CostingCode5 = InDT.Rows(Row)("cc5").ToString
                                    'objjournalentry.Lines.ProjectCode = ""
                                    objjournalentry.Lines.Add()
                                    'objjournalentry.Lines.AccountCode = objmatrix.Columns.Item("U_AcctCode").Cells.Item(AccRow).Specific.String
                                    If objmatrix.Columns.Item("U_Source").Cells.Item(AccRow).Specific.Selected.Value = "1" Then 'Business Partner 
                                        objjournalentry.Lines.ShortName = objmatrix.Columns.Item("U_AcctCode").Cells.Item(AccRow).Specific.String
                                    Else
                                        objjournalentry.Lines.AccountCode = objmatrix.Columns.Item("U_AcctCode").Cells.Item(AccRow).Specific.String
                                    End If
                                    objjournalentry.Lines.Debit = JEAmount
                                    objjournalentry.Lines.BPLID = objmatrix.Columns.Item("U_Branch").Cells.Item(AccRow).Specific.Selected.Value
                                    objjournalentry.Lines.UserFields.Fields.Item("U_TBranch").Value = objform.Items.Item("1320002037").Specific.Selected.Description ' Branch
                                    'If InDT.Rows(Row)("cc1").ToString <> "" Then objjournalentry.Lines.CostingCode = InDT.Rows(Row)("cc1").ToString
                                    'If InDT.Rows(Row)("cc2").ToString <> "" Then objjournalentry.Lines.CostingCode2 = InDT.Rows(Row)("cc2").ToString
                                    'If InDT.Rows(Row)("cc3").ToString <> "" Then objjournalentry.Lines.CostingCode3 = InDT.Rows(Row)("cc3").ToString
                                    'If InDT.Rows(Row)("cc4").ToString <> "" Then objjournalentry.Lines.CostingCode4 = InDT.Rows(Row)("cc4").ToString
                                    'If InDT.Rows(Row)("cc5").ToString <> "" Then objjournalentry.Lines.CostingCode5 = InDT.Rows(Row)("cc5").ToString
                                    'objjournalentry.Lines.ProjectCode = ""
                                    objjournalentry.Lines.Add()
                                End If

                            End If
                        Next

                        If objjournalentry.Add <> 0 Then
                            If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            objaddon.objapplication.SetStatusBarMessage("Journal: " & objaddon.objcompany.GetLastErrorDescription & "-" & objaddon.objcompany.GetLastErrorCode, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(objjournalentry)
                            Return False
                        Else
                            If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                            TransId = objaddon.objcompany.GetNewObjectKey()
                            objform.Items.Item("tjeno").Specific.String = TransId
                            objaddon.objapplication.SetStatusBarMessage("Journal Entry Created Successfully...", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                            Return True
                        End If
                    Else
                        Return True
                    End If

                Catch ex As Exception
                    If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    objaddon.objapplication.SetStatusBarMessage("JE Posting Error" & objaddon.objcompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                    Return False
                End Try

            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage("JE " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                Return False
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End Function

        Private Function Validate_Row(ByVal FormUID As String, ByVal Row As Integer, ByVal GLCode As String) As Boolean
            Try
                Dim flag As Boolean = False
                objform = objaddon.objapplication.Forms.Item(FormUID)
                objmatrix = objform.Items.Item("71").Specific
                If Row = 0 Then
                    For IRow As Integer = objmatrix.VisualRowCount To objmatrix.VisualRowCount - 1 Step -1
                        If objmatrix.Columns.Item("8").Cells.Item(IRow).Specific.String <> "" And objmatrix.Columns.Item("8").Cells.Item(IRow).Specific.String = GLCode Then
                            If objmatrix.Columns.Item("5").Cells.Item(IRow).Specific.String = "" Or objmatrix.Columns.Item("U_AcctCode").Cells.Item(IRow).Specific.String = "" Or objmatrix.Columns.Item("U_Branch").Cells.Item(IRow).Specific.Selected Is Nothing Then
                                objaddon.objapplication.MessageBox("Select the Amount/ GL Account/ Branch...on Row: " & IRow, , "OK")
                                objaddon.objapplication.StatusBar.SetText("Select the Amount/ GL Account/ Branch...on Row: " & IRow, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                flag = True ': Return True
                            End If
                        End If
                    Next
                ElseIf Row = -1 Then
                    For IRow As Integer = 1 To objmatrix.VisualRowCount
                        If objmatrix.Columns.Item("8").Cells.Item(IRow).Specific.String <> "" And objmatrix.Columns.Item("8").Cells.Item(IRow).Specific.String = GLCode Then
                            If objmatrix.Columns.Item("5").Cells.Item(IRow).Specific.String = "" Or objmatrix.Columns.Item("U_AcctCode").Cells.Item(IRow).Specific.String = "" Or objmatrix.Columns.Item("U_Branch").Cells.Item(IRow).Specific.Selected Is Nothing Then
                                objaddon.objapplication.MessageBox("Select the Amount/ GL Account/ Branch fields since multi branch selection has been enabled...on Row: " & IRow, , "OK")
                                objaddon.objapplication.StatusBar.SetText("Select the Amount/ GL Account/ Branch fields since multi branch selection has been enabled...on Row: " & IRow, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                flag = True ': Return True
                            End If

                        End If
                    Next

                Else
                    If objmatrix.Columns.Item("8").Cells.Item(Row).Specific.String <> "" And objmatrix.Columns.Item("8").Cells.Item(Row).Specific.String = GLCode Then
                        If objmatrix.Columns.Item("5").Cells.Item(Row).Specific.String = "" Or objmatrix.Columns.Item("U_AcctCode").Cells.Item(Row).Specific.String = "" Or objmatrix.Columns.Item("U_Branch").Cells.Item(Row).Specific.Selected Is Nothing Then
                            objaddon.objapplication.MessageBox("Select the Amount/ GL Account/ Branch...on Row: " & Row, , "OK")
                            objaddon.objapplication.StatusBar.SetText("Select Amount/ GL Account/ Branch...on Row: " & Row, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            flag = True ': Return True
                        End If
                    End If
                End If


                Return flag
            Catch ex As Exception
            End Try
        End Function

        Private Function ValidateAccountLines(ByVal FormUID As String) As Boolean
            Try
                Dim SameBranchFlag As Boolean
                Dim GLAccount As String = ""
                objform = objaddon.objapplication.Forms.Item(FormUID)
                objmatrix = objform.Items.Item("71").Specific
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If objaddon.HANA Then
                    objRs.DoQuery("select distinct ""PmtClrAct"" from OBPL where ""Disabled"" ='N'")
                Else
                    objRs.DoQuery("select distinct PmtClrAct from OBPL where Disabled ='N'")
                End If
                If objRs.RecordCount > 0 Then
                    For i As Integer = 0 To objRs.RecordCount - 1
                        GLAccount = objRs.Fields.Item(0).Value.ToString
                        For IRow As Integer = 1 To objmatrix.VisualRowCount
                            If objmatrix.Columns.Item("8").Cells.Item(IRow).Specific.String = "" Then Continue For
                            If objmatrix.Columns.Item("8").Cells.Item(IRow).Specific.String = GLAccount And Not objmatrix.Columns.Item("U_Branch").Cells.Item(IRow).Specific.Selected Is Nothing Then
                                If objform.Items.Item("1320002037").Specific.Selected.Value = objmatrix.Columns.Item("U_Branch").Cells.Item(IRow).Specific.Selected.Value Then
                                    SameBranchFlag = True
                                End If
                            End If
                        Next
                        objRs.MoveNext()
                    Next
                End If
                objRs = Nothing
                Return SameBranchFlag

            Catch ex As Exception
                Return False
            End Try
        End Function

        Private Function ValidTransaction(ByVal FormUID As String, ByVal GLCode As String) As Boolean
            Try
                Dim SameBranchFlag As Boolean
                objform = objaddon.objapplication.Forms.Item(FormUID)
                objmatrix = objform.Items.Item("71").Specific
                For IRow As Integer = 1 To objmatrix.VisualRowCount
                    If objmatrix.Columns.Item("8").Cells.Item(IRow).Specific.String = "" Then Continue For
                    If objmatrix.Columns.Item("8").Cells.Item(IRow).Specific.String = GLCode And Not objmatrix.Columns.Item("U_Branch").Cells.Item(IRow).Specific.Selected Is Nothing Then
                        SameBranchFlag = True
                    End If
                Next
                Return SameBranchFlag

            Catch ex As Exception
                Return False
            End Try
        End Function

        Private Function Update_Standard_JE_TranBranch(ByVal FormUID As String, ByVal TransEntry As String) As Boolean
            Try
                Dim objjournalentry As SAPbobsCOM.JournalEntries
                Dim Row As Integer = 1
                Try
                    objform = objaddon.objapplication.Forms.Item(FormUID)
                    objmatrix = objform.Items.Item("71").Specific
                    If TransEntry <> "" Then
                        objjournalentry = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
                        If objjournalentry.GetByKey(Trim(TransEntry)) Then
                            objaddon.objapplication.StatusBar.SetText("Standard JE Updating Please wait...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            For ww As Integer = 1 To objjournalentry.Lines.Count - 1
                                If objmatrix.Columns.Item("U_AcctCode").Cells.Item(Row).Specific.String <> "" And Not objmatrix.Columns.Item("U_Branch").Cells.Item(Row).Specific.Selected Is Nothing Then
                                    objjournalentry.Lines.SetCurrentLine(Row)
                                    objjournalentry.Lines.UserFields.Fields.Item("U_TBranch").Value = objmatrix.Columns.Item("U_Branch").Cells.Item(Row).Specific.Selected.Description 'Branch
                                End If
                                Row += 1
                            Next
                            'For AccRow As Integer = 1 To objmatrix.VisualRowCount
                            '    If objmatrix.Columns.Item("U_AcctCode").Cells.Item(AccRow).Specific.String <> "" And Not objmatrix.Columns.Item("U_Branch").Cells.Item(AccRow).Specific.Selected Is Nothing Then
                            '        objjournalentry.Lines.SetCurrentLine(AccRow)
                            '        objjournalentry.Lines.UserFields.Fields.Item("U_TBranch").Value = objmatrix.Columns.Item("U_Branch").Cells.Item(AccRow).Specific.Selected.Description 'Branch
                            '    End If
                            'Next
                            If objjournalentry.Update() <> 0 Then
                                objaddon.objapplication.SetStatusBarMessage("JE: " & objaddon.objcompany.GetLastErrorDescription & "-" & objaddon.objcompany.GetLastErrorCode, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(objjournalentry)
                                Return False
                            Else
                                objaddon.objapplication.SetStatusBarMessage("Standard JE Updated Successfully...", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                Return True
                            End If
                        End If
                    Else
                        Return True
                    End If

                Catch ex As Exception
                    objaddon.objapplication.SetStatusBarMessage("JE" & objaddon.objcompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                    Return False
                End Try

            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage("JE " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                Return False
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End Function

    End Class
End Namespace
