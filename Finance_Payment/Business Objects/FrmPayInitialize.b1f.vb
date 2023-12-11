Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace Finance_Payment
    <FormAttribute("PAYINIT", "Business Objects/FrmPayInitialize.b1f")>
    Friend Class FrmPayInitialize
        Inherits UserFormBase
        Public WithEvents objform As SAPbouiCOM.Form
        Private WithEvents pCFL As SAPbouiCOM.ISBOChooseFromListEventArg
        Dim FormCount As Integer = 0
        Dim strSQL As String
        Dim objRs As SAPbobsCOM.Recordset

        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("101").Specific, SAPbouiCOM.Button)
            Me.Button1 = CType(Me.GetItem("2").Specific, SAPbouiCOM.Button)
            Me.OptionBtn0 = CType(Me.GetItem("opincpay").Specific, SAPbouiCOM.OptionBtn)
            Me.OptionBtn1 = CType(Me.GetItem("opoutpay").Specific, SAPbouiCOM.OptionBtn)
            Me.StaticText0 = CType(Me.GetItem("lpaytype").Specific, SAPbouiCOM.StaticText)
            Me.StaticText1 = CType(Me.GetItem("lpaydate").Specific, SAPbouiCOM.StaticText)
            Me.EditText0 = CType(Me.GetItem("tpaydate").Specific, SAPbouiCOM.EditText)
            Me.StaticText2 = CType(Me.GetItem("lbpcode").Specific, SAPbouiCOM.StaticText)
            Me.EditText1 = CType(Me.GetItem("txtbpcode").Specific, SAPbouiCOM.EditText)
            Me.CheckBox0 = CType(Me.GetItem("chkmulbp").Specific, SAPbouiCOM.CheckBox)
            Me.Matrix0 = CType(Me.GetItem("mtxdata").Specific, SAPbouiCOM.Matrix)
            Me.OptionBtn2 = CType(Me.GetItem("opinrec").Specific, SAPbouiCOM.OptionBtn)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler ResizeAfter, AddressOf Me.Form_ResizeAfter

        End Sub

        Private Sub OnCustomInitialize()
            Try
                objform = objaddon.objapplication.Forms.GetForm("PAYINIT", Me.FormCount)
                objform = objaddon.objapplication.Forms.ActiveForm
                objform.Freeze(True)
                Try
                    objform.EnableMenu("1281", False)
                    objform.EnableMenu("1282", False)
                Catch ex As Exception
                End Try
                Matrix0.AutoResizeColumns()
                OptionBtn1.GroupWith("opincpay")
                OptionBtn2.GroupWith("opincpay")
                objform.Items.Item("tpaydate").Specific.string = Now.Date.ToString("yyyyMMdd")
                objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "bpcode", "#")
                CheckBox0.Checked = True
                CheckBox0.Item.Enabled = False
                'OptionBtn0.Selected = True
                OptionBtn2.Selected = True
                OptionBtn0.Item.Visible = False
                OptionBtn1.Item.Visible = False
                OptionBtn2.Item.Left = OptionBtn0.Item.Left
                objform.Left = (objaddon.objapplication.Desktop.Width - objform.Width) / 2
                objform.Top = (objaddon.objapplication.Desktop.Height - objform.MaxHeight) / 3
                objform.Freeze(False)
            Catch ex As Exception
                objform.Freeze(False)
            End Try
        End Sub

#Region "Fields"
        Private WithEvents Button0 As SAPbouiCOM.Button
        Private WithEvents Button1 As SAPbouiCOM.Button
        Private WithEvents OptionBtn0 As SAPbouiCOM.OptionBtn
        Private WithEvents OptionBtn1 As SAPbouiCOM.OptionBtn
        Private WithEvents StaticText0 As SAPbouiCOM.StaticText
        Private WithEvents StaticText1 As SAPbouiCOM.StaticText
        Private WithEvents EditText0 As SAPbouiCOM.EditText
        Private WithEvents StaticText2 As SAPbouiCOM.StaticText
        Private WithEvents EditText1 As SAPbouiCOM.EditText
        Private WithEvents CheckBox0 As SAPbouiCOM.CheckBox
        Private WithEvents Matrix0 As SAPbouiCOM.Matrix
        Private WithEvents OptionBtn2 As SAPbouiCOM.OptionBtn
#End Region

        Private Sub EditText1_ChooseFromListBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles EditText1.ChooseFromListBefore
            Try
                If pVal.ActionSuccess = True Then Exit Sub
                Try
                    Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFL_BP")
                    Dim oConds As SAPbouiCOM.Conditions
                    Dim oCond As SAPbouiCOM.Condition
                    Dim oEmptyConds As New SAPbouiCOM.Conditions
                    oCFL.SetConditions(oEmptyConds)
                    oConds = oCFL.GetConditions()

                    oCond = oConds.Add()
                    oCond.Alias = "validFor"
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCond.CondVal = "Y"
                    oCFL.SetConditions(oConds)
                Catch ex As Exception
                    SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                End Try
            Catch ex As Exception

            End Try

        End Sub

        Private Sub EditText1_ChooseFromListAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText1.ChooseFromListAfter
            Try
                If pVal.ActionSuccess = False Then Exit Sub
                pCFL = pVal
                If Not pCFL.SelectedObjects Is Nothing Then
                    Try
                        EditText1.Value = pCFL.SelectedObjects.Columns.Item("CardCode").Cells.Item(0).Value
                    Catch ex As Exception
                        EditText1.Value = pCFL.SelectedObjects.Columns.Item("CardCode").Cells.Item(0).Value
                    End Try
                End If
            Catch ex As Exception
            End Try
        End Sub

        Private Sub Matrix0_ChooseFromListBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Matrix0.ChooseFromListBefore
            Try
                If pVal.ActionSuccess = True Then Exit Sub
                Try
                    If Matrix0.Columns.Item("bpcode").Cells.Item(pVal.Row).Specific.String = "" Then
                        If Matrix0.Columns.Item("bpname").Cells.Item(pVal.Row).Specific.String <> "" Then Matrix0.ClearRowData(pVal.Row)
                    End If
                    Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFL_BP1")
                    Dim oConds As SAPbouiCOM.Conditions
                    Dim oCond As SAPbouiCOM.Condition
                    Dim oEmptyConds As New SAPbouiCOM.Conditions
                    oCFL.SetConditions(oEmptyConds)
                    oConds = oCFL.GetConditions()

                    oCond = oConds.Add()
                    oCond.Alias = "validFor"
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCond.CondVal = "Y"
                    'oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                    'oCond = oConds.Add()
                    'oCond.Alias = "Balance"
                    'oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL
                    'oCond.CondVal = "0"
                    If PaymentWithReco = "N" Then
                        If OptionBtn0.Selected = True Then
                            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                            oCond = oConds.Add()
                            oCond.Alias = "CardType"
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                            oCond.CondVal = "C"
                        ElseIf OptionBtn1.Selected = True Then
                            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                            oCond = oConds.Add()
                            oCond.Alias = "CardType"
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                            oCond.CondVal = "S"
                        End If
                    End If


                    oCFL.SetConditions(oConds)
                Catch ex As Exception
                    SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                End Try
            Catch ex As Exception
            End Try

        End Sub

        Private Sub Matrix0_ChooseFromListAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.ChooseFromListAfter
            Try
                If pVal.ActionSuccess = False Then Exit Sub
                pCFL = pVal
                If pCFL.SelectedObjects Is Nothing Then Exit Sub
                Try
                    Dim CurRow As Integer = pVal.Row
                    For Row As Integer = 0 To pCFL.SelectedObjects.Rows.Count - 1
                        For NRow As Integer = Row + 1 To pCFL.SelectedObjects.Rows.Count - 1
                            If pCFL.SelectedObjects.Columns.Item("Currency").Cells.Item(Row).Value <> pCFL.SelectedObjects.Columns.Item("Currency").Cells.Item(NRow).Value And pCFL.SelectedObjects.Columns.Item("Currency").Cells.Item(Row).Value <> "##" And pCFL.SelectedObjects.Columns.Item("Currency").Cells.Item(NRow).Value <> "##" Then
                                objaddon.objapplication.SetStatusBarMessage("Inconsistency in BP currency. Select all business partners with the same currency.", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                Exit Sub
                            End If
                        Next

                        For i As Integer = Matrix0.VisualRowCount - 1 To 1 Step -1
                            If Matrix0.Columns.Item("bpcode").Cells.Item(i).Specific.String <> "" And Matrix0.Columns.Item("bpcur").Cells.Item(i).Specific.String <> "" Then
                                If pCFL.SelectedObjects.Columns.Item("Currency").Cells.Item(Row).Value <> Matrix0.Columns.Item("bpcur").Cells.Item(i).Specific.String And pCFL.SelectedObjects.Columns.Item("Currency").Cells.Item(Row).Value <> "##" And Matrix0.Columns.Item("bpcur").Cells.Item(i).Specific.String <> "##" Then
                                    objaddon.objapplication.SetStatusBarMessage("Inconsistency in BP currency.. Select all business partners with the same currency.", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                    Exit Sub
                                End If
                                'strSQL = pCFL.SelectedObjects.Columns.Item("CardCode").Cells.Item(Row).Value
                                If Matrix0.Columns.Item("bpcode").Cells.Item(i).Specific.String = pCFL.SelectedObjects.Columns.Item("CardCode").Cells.Item(Row).Value Then
                                    objaddon.objapplication.SetStatusBarMessage("Duplicate value...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                    Exit Sub
                                End If

                            End If
                        Next
                    Next

                    For Row As Integer = 0 To pCFL.SelectedObjects.Rows.Count - 1
                        objform.DataSources.UserDataSources.Item("UD_3").Value = CurRow
                        objform.DataSources.DBDataSources.Item("OCRD").SetValue("CardCode", 0, pCFL.SelectedObjects.Columns.Item("CardCode").Cells.Item(Row).Value)
                        objform.DataSources.DBDataSources.Item("OCRD").SetValue("CardName", 0, pCFL.SelectedObjects.Columns.Item("CardName").Cells.Item(Row).Value)
                        objform.DataSources.DBDataSources.Item("OCRD").SetValue("Currency", 0, pCFL.SelectedObjects.Columns.Item("Currency").Cells.Item(Row).Value)
                        objform.DataSources.UserDataSources.Item("UD_4").ValueEx = pCFL.SelectedObjects.Columns.Item("Balance").Cells.Item(Row).Value
                        Matrix0.SetLineData(CurRow)
                        CurRow += 1
                        objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "bpcode", "#")
                    Next

                Catch ex As Exception
                    'Matrix0.Columns.Item("bpcode").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("CardCode").Cells.Item(0).Value
                End Try
                'objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "bpcode", "#")
                Matrix0.AutoResizeColumns()

            Catch ex As Exception
            End Try

        End Sub

        Private Sub Button0_ClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.ClickAfter
            Try
                Dim CardCode As String = ""
                PayInitDate = Date.ParseExact(EditText0.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                Dim mulbp As String = "", PayInout As String = ""
                If CheckBox0.Checked Then
                    mulbp = "Y"
                End If
                If OptionBtn0.Selected = True Then
                    PayInout = "Y"
                ElseIf OptionBtn1.Selected = True Then
                    PayInout = "N"
                ElseIf OptionBtn2.Selected = True Then
                    PayInout = "IR"
                End If
                For i As Integer = 1 To Matrix0.VisualRowCount
                    If Matrix0.Columns.Item("bpcode").Cells.Item(i).Specific.String <> "" Then
                        If i = 1 Then
                            CardCode = "'" + Matrix0.Columns.Item("bpcode").Cells.Item(i).Specific.String + "'"
                        Else
                            CardCode += ",'" + Matrix0.Columns.Item("bpcode").Cells.Item(i).Specific.String + "'"
                        End If
                    End If
                Next

                If PayInout = "Y" Then
                    Query = GetPaymentQuery("IN", PayInitDate.ToString("yyyyMMdd"), CardCode)
                    If ActivateExchangeRateWindow() Then
                        Exit Sub
                    End If
                    If Not objaddon.FormExist("FINPAY") Then
                        Dim activeform As New FrmInPayments
                        activeform.Show()
                    End If
                ElseIf PayInout = "N" Then
                    Query = GetPaymentQuery("OUT", PayInitDate.ToString("yyyyMMdd"), CardCode)
                    If ActivateExchangeRateWindow() Then
                        Exit Sub
                    End If
                    If Not objaddon.FormExist("FOUTPAY") Then
                        Dim activeform As New FrmOutPayments
                        activeform.Show()
                    End If
                ElseIf PayInout = "IR" Then
                    Query = Get_InternalReconciliation_Query(PayInitDate.ToString("yyyyMMdd"), CardCode)
                    If Not objaddon.FormExist("FOITR") Then
                        Dim activeform As New FrmInternalReconciliation
                        activeform.Show()
                    End If

                End If

            Catch ex As Exception

            End Try

        End Sub

        Private Function ActivateExchangeRateWindow() As Boolean
            Try
                Dim objRs1 As SAPbobsCOM.Recordset
                Dim oCombo As SAPbouiCOM.ComboBox
                Dim DocDate As Date = Date.ParseExact(EditText0.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                Dim GetMonth As Integer = Month(DocDate)
                Dim GetYear As Integer = Year(DocDate)
                Dim GetDate As String = DocDate.ToString("dd")
                strSQL = Query.Remove(Query.Length - 35, 35)
                strSQL = strSQL.Remove(7, 67)
                strSQL = strSQL.Insert(7, "Distinct B.""DocCur"" ")
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRs1 = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRs.DoQuery(strSQL)
                For Rec As Integer = 0 To objRs.RecordCount - 1
                    If objRs.Fields.Item(0).Value.ToString <> MainCurr Then
                        strSQL = "Select ""Rate"",""Currency"" from ORTT where ""RateDate""= '" & DocDate.ToString("yyyyMMdd") & "' and ""Currency""='" & objRs.Fields.Item(0).Value.ToString & "' "
                        objRs1.DoQuery(strSQL)
                        If objRs1.RecordCount = 0 Then
                            objaddon.objapplication.Menus.Item("3333").Activate()
                            Dim oForm As SAPbouiCOM.Form
                            Dim oMatrix As SAPbouiCOM.Matrix
                            oForm = objaddon.objapplication.Forms.ActiveForm

                            oCombo = oForm.Items.Item("12").Specific
                            If oCombo.Selected.Value <> GetYear Then oCombo.Select(GetYear, SAPbouiCOM.BoSearchKey.psk_ByValue)
                            oCombo = oForm.Items.Item("13").Specific
                            If oCombo.Selected.Value <> GetMonth Then oCombo.Select(GetMonth, SAPbouiCOM.BoSearchKey.psk_ByValue)
                            oMatrix = oForm.Items.Item("4").Specific
                            Dim ColId As String = ""
                            For i As Integer = 0 To oMatrix.Columns.Count - 1
                                If oMatrix.Columns.Item(i).TitleObject.Caption = objRs.Fields.Item(0).Value.ToString Then
                                    ColId = oMatrix.Columns.Item(i).UniqueID
                                End If
                            Next
                            oMatrix.Columns.Item(0).Cells.Item(CInt(GetDate)).Click()
                            oMatrix.Columns.Item(ColId).Cells.Item(CInt(GetDate)).Click()
                            Return True
                        End If
                    End If
                    objRs.MoveNext()
                Next
                Return False
            Catch ex As Exception
                Return False
            End Try
        End Function

        Private Function GetPaymentQuery(ByVal PayType As String, ByVal DocDate As String, ByVal CardCode As String) As String
            Try
                If objaddon.HANA Then
                    strSQL = "Select ROW_NUMBER() OVER (ORDER BY B.""CardCode"",B.""DocDate"") AS ""LineId"",* from (SELECT "
                    strSQL += vbCrLf + "Case when A.""DocCur""=A.""MainCur"" Then A.""DocTotalLC"" Else A.""DocTotalFC"" End ""DTotal"","
                    strSQL += vbCrLf + "Case when A.""DocCur""=A.""MainCur"" Then A.""BalanceLC"" Else A.""BalanceFC"" End ""BalDue"","
                    strSQL += vbCrLf + "* FROM "
                    strSQL += vbCrLf + "(SELECT 'N' AS ""Selected"",T1.""TransId"",T1.""Line_ID"",T1.""DebCred"",CASE WHEN ""FCCurrency"" IS NULL THEN (SELECT ""MainCurncy"" FROM OADM) ELSE ""FCCurrency"" END AS ""DocCur"",(SELECT ""MainCurncy"" FROM OADM) as ""MainCur"","
                    strSQL += vbCrLf + "T1.""TransType"" AS ""ObjType"",T1.""Ref2"" as ""NumAtCard"",CASE WHEN T1.""SourceID"" IS NULL THEN T1.""TransId"" ELSE T1.""BaseRef"" END AS ""DocNum"",CASE WHEN T1.""SourceID"" IS NULL THEN T1.""TransId"" ELSE T1.""SourceID"" END ""DocEntry"","
                    strSQL += vbCrLf + "CASE WHEN T1.""TransType""='13' THEN 'IN' WHEN T1.""TransType""='14' THEN 'CN' WHEN T1.""TransType""='203' or T1.""TransType""='204' THEN 'DT' WHEN T1.""TransType""='18' THEN 'PU' WHEN T1.""TransType""='19' THEN 'PC' Else 'JE' END AS ""Origin"","
                    strSQL += vbCrLf + "CASE WHEN T1.""TransType""='13' THEN 'A/R Invoice' WHEN T1.""TransType""='14' THEN 'A/R Credit Memo'  WHEN T1.""TransType""='203' or T1.""TransType""='204' THEN 'A/R Down Payment' WHEN T1.""TransType""='18' THEN 'A/P Invoice'"
                    strSQL += vbCrLf + "WHEN T1.""TransType""='19' THEN 'A/R Credit Memo' ELSE 'Journal Entry' END AS ""DocType"","
                    strSQL += vbCrLf + "T1.""ShortName"" AS ""CardCode"",(SELECT ""CardName"" FROM OCRD where ""CardCode""=T1.""ShortName"") as ""CardName"","
                    If PayType = "IN" Then
                        strSQL += vbCrLf + "CASE WHEN T1.""Credit""<>0  THEN  -T1.""Credit"" ELSE T1.""Debit"" END AS ""DocTotalLC"",CASE WHEN T1.""FCCredit""<>0  THEN  -T1.""FCCredit"" Else T1.""FCDebit"" END AS ""DocTotalFC"","
                        strSQL += vbCrLf + "CASE WHEN T1.""BalDueCred""<>0  THEN  -T1.""BalDueCred"" ELSE T1.""BalDueDeb"" END  AS ""BalanceLC"",CASE WHEN T1.""BalFcCred""<>0  THEN  -T1.""BalFcCred"" ELSE T1.""BalFcDeb"" END AS ""BalanceFC"","
                    Else
                        strSQL += vbCrLf + "CASE WHEN T1.""Credit""<>0  THEN  T1.""Credit"" ELSE -T1.""Debit"" END AS ""DocTotalLC"",CASE WHEN T1.""FCCredit""<>0  THEN  T1.""FCCredit"" Else -T1.""FCDebit"" END AS ""DocTotalFC"","
                        strSQL += vbCrLf + "CASE WHEN T1.""BalDueCred""<>0  THEN  T1.""BalDueCred"" ELSE -T1.""BalDueDeb"" END AS ""BalanceLC"",CASE WHEN T1.""BalFcCred""<>0  THEN  T1.""BalFcCred"" ELSE -T1.""BalFcDeb"" END AS ""BalanceFC"","
                    End If
                    strSQL += vbCrLf + "To_Varchar(T0.""RefDate"",'yyyyMMdd') AS ""DocDate"",DAYS_BETWEEN(T0.""DueDate"",CURRENT_DATE) AS ""OverDueDays"","
                    strSQL += vbCrLf + "T1.""BPLId"",(SELECT ""BPLName"" FROM OBPL WHERE ""BPLId""=T1.""BPLId"") AS ""BPLName"""
                    strSQL += vbCrLf + "FROM OJDT T0 join JDT1 T1 ON T0.""TransId""=T1.""TransId"" where T1.""DprId"" is null"
                    strSQL += vbCrLf + ") A "  'WHERE T0.""TransType"" IN (13,14,24,46,30,18,19)
                    strSQL += vbCrLf + "WHERE A.""DocDate""<='" & DocDate & "' and A.""BPLId"" in (Select T0.""BPLId"" from OBPL T0 join USR6 T1 on T0.""BPLId""=T1.""BPLId"" where T1.""UserCode""='" & objaddon.objcompany.UserName & "' and T0.""Disabled""<>'Y')  and A.""CardCode"" In (" & CardCode & ") and A.""BalanceLC""<>0"
                    strSQL += vbCrLf + ") B where B.""BalDue""<>0"
                    strSQL += vbCrLf + "ORDER BY B.""CardCode"",B.""DocDate"""
                Else

                End If
                'If objaddon.HANA Then
                '    strSQL = "Select ROW_NUMBER() OVER (Order BY A.""DocDate"") AS ""LineId"", * from "
                '    strSQL += vbCrLf + "(SELECT ROW_NUMBER() OVER () as ""#"",'N' as ""Selected"",T1.""TransId"",T1.""Line_ID"",T0.""DocCur"",T0.""ObjType"",T0.""NumAtCard"",T0.""DocNum"",T0.""DocEntry"",'IN' as ""Origin"",'A/R Invoice' as ""DocType"",T0.""CardCode"", T0.""CardName"","
                '    If PayType = "IN" Then  'AR Invoice
                '        strSQL += vbCrLf + "T0.""DocTotal"" as ""DocTotal"", T0.""DocTotalFC"" as ""DocTotalFC"",SUM(T0.""DocTotal""-T0.""PaidToDate"") ""Balance"",SUM(T0.""DocTotalFC""-T0.""PaidFC"") ""BalanceFC"","
                '    Else
                '        strSQL += vbCrLf + "-T0.""DocTotal"" as ""DocTotal"",-T0.""DocTotalFC"" as ""DocTotalFC"",-SUM(T0.""DocTotal""-T0.""PaidToDate"") ""Balance"",-SUM(T0.""DocTotalFC""-T0.""PaidFC"") ""BalanceFC"","
                '    End If
                '    strSQL += vbCrLf + "To_Varchar(T0.""DocDate"",'yyyyMMdd') as ""DocDate"",DAYS_BETWEEN(T0.""DocDueDate"",CURRENT_DATE) as ""OverDueDays"",T0.""BPLId"",(Select ""BPLName"" from OBPL where ""BPLId""=T0.""BPLId"") as ""BPLName"""
                '    strSQL += vbCrLf + "FROM OINV T0 join JDT1 T1 on T1.""SourceID""=T0.""DocEntry"" and T1.""BaseRef""=T0.""DocNum"" Group by T0.""CardCode"", T0.""CardName"",T0.""DocCur"",T0.""DocDate"",T0.""DocTotal"",T0.""DocTotalFC"",T0.""DocDueDate"",T0.""DocNum"",T0.""DocEntry"",T0.""BPLId"",T0.""ObjType"",T0.""NumAtCard"",T1.""TransId"",T1.""Line_ID"""
                '    strSQL += vbCrLf + "Union all"
                '    strSQL += vbCrLf + "SELECT ROW_NUMBER() OVER () as ""#"",'N' as ""Selected"",T1.""TransId"",T1.""Line_ID"",T0.""DocCur"",T0.""ObjType"",T0.""NumAtCard"",T0.""DocNum"",T0.""DocEntry"",'CN' as ""Origin"",'A/R Credit Memo' as ""DocType"",T0.""CardCode"", T0.""CardName"","
                '    If PayType = "IN" Then 'AR Credit Memo  
                '        strSQL += vbCrLf + "-T0.""DocTotal"",-T0.""DocTotalFC"",-SUM(T0.""DocTotal""-T0.""PaidToDate"") ""Balance"",-SUM(T0.""DocTotalFC""-T0.""PaidFC"") ""BalanceFC"","
                '    Else
                '        strSQL += vbCrLf + "T0.""DocTotal"",T0.""DocTotalFC"",SUM(T0.""DocTotal""-T0.""PaidToDate"") ""Balance"",SUM(T0.""DocTotalFC""-T0.""PaidFC"") ""BalanceFC"","
                '    End If
                '    strSQL += vbCrLf + "To_Varchar(T0.""DocDate"",'yyyyMMdd') as ""DocDate"",DAYS_BETWEEN(T0.""DocDueDate"",CURRENT_DATE) as ""OverDueDays"",T0.""BPLId"",(Select ""BPLName"" from OBPL where ""BPLId""=T0.""BPLId"") as ""BPLName"""
                '    strSQL += vbCrLf + "FROM ORIN T0 join JDT1 T1 on T1.""SourceID""=T0.""DocEntry"" and T1.""BaseRef""=T0.""DocNum"" Group by T0.""CardCode"", T0.""CardName"",T0.""DocCur"",T0.""DocDate"",T0.""DocTotal"",T0.""DocTotalFC"",T0.""DocDueDate"",T0.""DocNum"",T0.""DocEntry"",T0.""BPLId"",T0.""ObjType"",T0.""NumAtCard"",T1.""TransId"",T1.""Line_ID"""
                '    strSQL += vbCrLf + "Union all"
                '    strSQL += vbCrLf + "SELECT ROW_NUMBER() OVER () as ""#"",'N' as ""Selected"",T1.""TransId"",T1.""Line_ID"",T0.""DocCur"",T0.""ObjType"",T0.""NumAtCard"",T0.""DocNum"",T0.""DocEntry"",'PU' as ""Origin"",'A/P Invoice' as ""DocType"",T0.""CardCode"", T0.""CardName"","
                '    If PayType = "IN" Then 'AP Invoice
                '        strSQL += vbCrLf + "-T0.""DocTotal"",-T0.""DocTotalFC"",-SUM(T0.""DocTotal""-T0.""PaidToDate"") ""Balance"",-SUM(T0.""DocTotalFC""-T0.""PaidFC"") ""BalanceFC"","
                '    Else
                '        strSQL += vbCrLf + "T0.""DocTotal"",T0.""DocTotalFC"",SUM(T0.""DocTotal""-T0.""PaidToDate"") ""Balance"",-SUM(T0.""DocTotalFC""-T0.""PaidFC"") ""BalanceFC"","
                '    End If
                '    strSQL += vbCrLf + "To_Varchar(T0.""DocDate"",'yyyyMMdd') as ""DocDate"",DAYS_BETWEEN(T0.""DocDueDate"",CURRENT_DATE) as ""OverDueDays"",T0.""BPLId"",(Select ""BPLName"" from OBPL where ""BPLId""=T0.""BPLId"") as ""BPLName"""
                '    strSQL += vbCrLf + "FROM OPCH T0 join JDT1 T1 on T1.""SourceID""=T0.""DocEntry"" and T1.""BaseRef""=T0.""DocNum"" Group by T0.""CardCode"", T0.""CardName"",T0.""DocCur"",T0.""DocDate"",T0.""DocTotal"",T0.""DocTotalFC"",T0.""DocDueDate"",T0.""DocNum"",T0.""DocEntry"",T0.""BPLId"",T0.""ObjType"",T0.""NumAtCard"",T1.""TransId"",T1.""Line_ID"""
                '    strSQL += vbCrLf + "Union all"
                '    strSQL += vbCrLf + "SELECT ROW_NUMBER() OVER () as ""#"",'N' as ""Selected"",T1.""TransId"",T1.""Line_ID"",T0.""DocCur"",T0.""ObjType"",T0.""NumAtCard"",T0.""DocNum"",T0.""DocEntry"",'PC' as ""Origin"",'A/P Credit Memo' as ""DocType"",T0.""CardCode"", T0.""CardName"","
                '    If PayType = "IN" Then 'AP Credit Memo
                '        strSQL += vbCrLf + "T0.""DocTotal"",T0.""DocTotalFC"",SUM(T0.""DocTotal""-T0.""PaidToDate"") ""Balance"",SUM(T0.""DocTotalFC""-T0.""PaidFC"") ""BalanceFC"","
                '    Else
                '        strSQL += vbCrLf + "-T0.""DocTotal"",-T0.""DocTotalFC"",-SUM(T0.""DocTotal""-T0.""PaidToDate"") ""Balance"",-SUM(T0.""DocTotalFC""-T0.""PaidFC"") ""BalanceFC"","
                '    End If
                '    strSQL += vbCrLf + "To_Varchar(T0.""DocDate"",'yyyyMMdd') as ""DocDate"",DAYS_BETWEEN(T0.""DocDueDate"",CURRENT_DATE) as ""OverDueDays"",T0.""BPLId"",(Select ""BPLName"" from OBPL where ""BPLId""=T0.""BPLId"") as ""BPLName"""
                '    strSQL += vbCrLf + "FROM ORPC T0 join JDT1 T1 on T1.""SourceID""=T0.""DocEntry"" and T1.""BaseRef""=T0.""DocNum"" Group by T0.""CardCode"", T0.""CardName"",T0.""DocCur"",T0.""DocDate"",T0.""DocTotal"",T0.""DocTotalFC"",T0.""DocDueDate"",T0.""DocNum"",T0.""DocEntry"",T0.""BPLId"",T0.""ObjType"",T0.""NumAtCard"",T1.""TransId"",T1.""Line_ID"""
                '    strSQL += vbCrLf + "Union all"
                '    strSQL += vbCrLf + "SELECT ROW_NUMBER() OVER () as ""#"",'N' as ""Selected"",T1.""TransId"",T1.""Line_ID"",Case when ""FCCurrency"" <>'' then ""FCCurrency"" Else (select ""MainCurncy"" from OADM) End,T1.""ObjType"",T1.""Ref1"",T1.""TransId"",T1.""BaseRef"",'JE' as ""Origin"",'Journal Entry' as ""DocType"",T1.""ShortName"" as ""CardCode"","
                '    strSQL += vbCrLf + "Case when (Select 1 from OCRD where ""CardCode""=T1.""ShortName"")='1' Then (Select ""CardName"" from OCRD where ""CardCode""=T1.""ShortName"") Else (Select ""AcctName"" from OACT where ""AcctCode""=T1.""ShortName"") End as ""CardName"","
                '    If PayType = "IN" Then 'Journal Entry
                '        strSQL += vbCrLf + "case when T1.""Credit""<>0  Then  -T1.""Credit"" Else T1.""Debit"" End as ""DocTotal"",case when T1.""FCCredit""<>0  Then  -T1.""FCCredit"" Else T1.""FCDebit"" End as ""DocTotalFC"",case when T1.""BalDueCred""<>0  Then  -T1.""BalDueCred"" Else T1.""BalDueDeb"" End as ""Balance"","
                '        strSQL += vbCrLf + "Case when T1.""BalFcCred""<>0  Then  -T1.""BalFcCred"" Else T1.""BalFcDeb"" End as ""BalanceFC"","
                '    Else
                '        strSQL += vbCrLf + "case when T1.""Credit""<>0  Then  T1.""Credit"" Else -T1.""Debit"" End as ""DocTotal"",case when T1.""FCCredit""<>0  Then  T1.""FCCredit"" Else -T1.""FCDebit"" End as ""DocTotalFC"",case when T1.""BalDueCred""<>0  Then  T1.""BalDueCred"" Else -T1.""BalDueDeb"" End as ""Balance"","
                '        strSQL += vbCrLf + "Case when T1.""BalFcCred""<>0  Then  T1.""BalFcCred"" Else -T1.""BalFcDeb"" End as ""BalanceFC"","
                '    End If
                '    strSQL += vbCrLf + "To_Varchar(T1.""RefDate"",'yyyyMMdd') as ""DocDate"",DAYS_BETWEEN(T1.""RefDate"",CURRENT_DATE) as ""OverDueDays"",T1.""BPLId"",(Select ""BPLName"" from OBPL where ""BPLId""=T1.""BPLId"") as ""BPLName"" "
                '    strSQL += vbCrLf + "FROM JDT1 T1 where T1.""TransType""in (30,24,46) and T1.""DprId"" is null"
                '    strSQL += vbCrLf + ") as A WHERE A.""DocDate""<='" & DocDate & "'  and A.""CardCode"" In (" & CardCode & ") and A.""Balance""<>0 "
                'Else

                'End If
                Return strSQL
            Catch ex As Exception
                Return ""
            End Try
        End Function

        Private Function Get_InternalReconciliation_Query(ByVal DocDate As String, ByVal CardCode As String) As String
            Try
                If objaddon.HANA Then
                    strSQL = "SELECT ROW_NUMBER() OVER (ORDER BY A.""CardCode"",A.""DocDate"") AS ""LineId"",* FROM  "
                    strSQL += vbCrLf + "(SELECT 'N' AS ""Selected"",T1.""TransId"",T1.""Line_ID"",T1.""DebCred"",CASE WHEN ""FCCurrency"" IS NULL THEN (SELECT ""MainCurncy"" FROM OADM) ELSE ""FCCurrency"" END AS ""DocCur"","
                    strSQL += vbCrLf + "T1.""TransType"" AS ""ObjType"",T1.""LineMemo"" as ""LineMemo"",T1.""BaseRef"" AS ""DocNum"",T1.""CreatedBy"" as ""DocEntry"","
                    strSQL += vbCrLf + "CASE WHEN T1.""TransType""='13' THEN 'IN' WHEN T1.""TransType""='14' THEN 'CN' WHEN T1.""TransType""='203' or T1.""TransType""='204' THEN 'DT' WHEN T1.""TransType""='18' THEN 'PU' WHEN T1.""TransType""='19' THEN 'PC'"
                    strSQL += vbCrLf + "WHEN T1.""TransType""='24' THEN 'RC' WHEN T1.""TransType""='46' THEN 'PS' Else 'JE' END AS ""Origin"","
                    strSQL += vbCrLf + "CASE WHEN T1.""TransType""='13' THEN 'A/R Invoice' WHEN T1.""TransType""='14' THEN 'A/R Credit Memo' WHEN T1.""TransType""='203' THEN 'A/R DownPayment' WHEN T1.""TransType""='18' THEN 'A/P Invoice'"
                    strSQL += vbCrLf + "WHEN T1.""TransType""='19' THEN 'A/R Credit Memo' WHEN T1.""TransType""='24' THEN 'Incoming Payment' WHEN T1.""TransType""='46' THEN 'Outgoing Payment' ELSE 'Journal Entry' END AS ""DocType"","
                    strSQL += vbCrLf + "T1.""ShortName"" AS ""CardCode"",(SELECT ""CardName"" FROM OCRD where ""CardCode""=T1.""ShortName"") as ""CardName"", (Select ""CardType"" from OCRD where ""CardCode""=T1.""ShortName"") as ""CardType"","
                    strSQL += vbCrLf + "CASE WHEN T1.""Credit""<>0  THEN  -T1.""Credit"" ELSE T1.""Debit"" END AS ""DocTotal"",CASE WHEN T1.""BalDueCred""<>0  THEN -T1.""BalDueCred"" ELSE T1.""BalDueDeb"" END AS ""Balance"","
                    strSQL += vbCrLf + "To_Varchar(T0.""RefDate"",'yyyyMMdd') AS ""DocDate"",T1.""BPLId"",(SELECT ""BPLName"" FROM OBPL WHERE ""BPLId""=T1.""BPLId"") AS ""BPLName"",T0.""Ref1"",T0.""Ref2"",T0.""Ref3"""
                    strSQL += vbCrLf + "FROM OJDT T0 join JDT1 T1 ON T0.""TransId""=T1.""TransId"" where T1.""DprId"" is null "
                    strSQL += vbCrLf + ") A "
                    strSQL += vbCrLf + "WHERE A.""DocDate""<='" & DocDate & "' and A.""BPLId"" in (Select T0.""BPLId"" from OBPL T0 join USR6 T1 on T0.""BPLId""=T1.""BPLId"" where T1.""UserCode""='" & objaddon.objcompany.UserName & "' and T0.""Disabled""<>'Y') and A.""CardCode"" In (" & CardCode & ") and A.""Balance""<>0"
                    strSQL += vbCrLf + "ORDER BY A.""CardCode"",A.""DocDate"""
                Else

                End If

                Return strSQL
            Catch ex As Exception
                Return ""
            End Try
        End Function

        Private Sub CheckBox0_PressedAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles CheckBox0.PressedAfter
            Try
                If CheckBox0.Checked = True Then
                    objform.ActiveItem = "tpaydate"
                    StaticText2.Item.Visible = False
                    EditText1.Item.Visible = False
                    Matrix0.Item.Visible = True
                Else
                    Matrix0.Item.Visible = False
                    StaticText2.Item.Visible = True
                    EditText1.Item.Visible = True
                End If
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Matrix0_LostFocusAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.LostFocusAfter
            Try

                objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "bpcode", "#")
            Catch ex As Exception

            End Try

        End Sub

        Private Sub RemoveLastrow(ByVal omatrix As SAPbouiCOM.Matrix, ByVal Columname_check As String)
            Try
                If omatrix.VisualRowCount = 0 Then Exit Sub
                If Columname_check.ToString = "" Then Exit Sub
                If omatrix.Columns.Item(Columname_check).Cells.Item(omatrix.VisualRowCount).Specific.string = "" Then
                    omatrix.DeleteRow(omatrix.VisualRowCount)
                End If
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Button0_ClickBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Button0.ClickBefore
            Try
                If OptionBtn0.Selected = False And OptionBtn1.Selected = False And OptionBtn2.Selected = False Then
                    objaddon.objapplication.SetStatusBarMessage("Payment Type is Missing.", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                    BubbleEvent = False : Exit Sub
                End If
                If EditText0.Value = "" Then
                    objaddon.objapplication.SetStatusBarMessage("Payment Date is Missing.", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                    BubbleEvent = False : Exit Sub
                End If
                If Matrix0.VisualRowCount = 0 Or Matrix0.Columns.Item("bpcode").Cells.Item(1).Specific.String = "" Then
                    objaddon.objapplication.SetStatusBarMessage("Line level is Missing.", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                    BubbleEvent = False : Exit Sub
                End If
                'RemoveLastrow(Matrix0, "bpcode")
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Form_ResizeAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                Matrix0.AutoResizeColumns()
            Catch ex As Exception

            End Try

        End Sub
    End Class
End Namespace
