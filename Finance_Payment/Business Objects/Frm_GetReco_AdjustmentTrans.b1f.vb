Option Strict Off
Option Explicit On

Imports System.Drawing
Imports SAPbouiCOM.Framework

Namespace Finance_Payment
    <FormAttribute("AdjTran", "Business Objects/Frm_GetReco_AdjustmentTrans.b1f")>
    Friend Class Frm_GetReco_AdjustmentTrans
        Inherits UserFormBase
        Public WithEvents objform As SAPbouiCOM.Form
        Private WithEvents pCFL As SAPbouiCOM.ISBOChooseFromListEventArg
        Dim FormCount As Integer = 0
        Dim strSQL As String
        Private WithEvents objDTable As SAPbouiCOM.DataTable
        Public Shared oSelectDT As New DataTable
        Public Shared objActDT As New DataTable
        Public WithEvents objRecoDR As DataRow
        Public Shared objRecoDT As New DataTable
        Private WithEvents objCheck As SAPbouiCOM.CheckBox

        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("1").Specific, SAPbouiCOM.Button)
            Me.Button1 = CType(Me.GetItem("2").Specific, SAPbouiCOM.Button)
            Me.StaticText0 = CType(Me.GetItem("lamtrec").Specific, SAPbouiCOM.StaticText)
            Me.EditText0 = CType(Me.GetItem("tamtrec").Specific, SAPbouiCOM.EditText)
            Me.StaticText1 = CType(Me.GetItem("ltotamt").Specific, SAPbouiCOM.StaticText)
            Me.EditText1 = CType(Me.GetItem("ttotamt").Specific, SAPbouiCOM.EditText)
            Me.StaticText2 = CType(Me.GetItem("lbranch").Specific, SAPbouiCOM.StaticText)
            Me.EditText2 = CType(Me.GetItem("tbcode").Specific, SAPbouiCOM.EditText)
            Me.Matrix0 = CType(Me.GetItem("mtxdata").Specific, SAPbouiCOM.Matrix)
            Me.EditText3 = CType(Me.GetItem("tbname").Specific, SAPbouiCOM.EditText)
            Me.EditText4 = CType(Me.GetItem("trow").Specific, SAPbouiCOM.EditText)
            Me.StaticText3 = CType(Me.GetItem("lrow").Specific, SAPbouiCOM.StaticText)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler ResizeAfter, AddressOf Me.Form_ResizeAfter

        End Sub

#Region "Fields"

        Private WithEvents Button0 As SAPbouiCOM.Button
        Private WithEvents Button1 As SAPbouiCOM.Button
        Private WithEvents StaticText0 As SAPbouiCOM.StaticText
        Private WithEvents EditText0 As SAPbouiCOM.EditText
        Private WithEvents StaticText1 As SAPbouiCOM.StaticText
        Private WithEvents EditText1 As SAPbouiCOM.EditText
        Private WithEvents StaticText2 As SAPbouiCOM.StaticText
        Private WithEvents EditText2 As SAPbouiCOM.EditText
        Private WithEvents Matrix0 As SAPbouiCOM.Matrix
        Private WithEvents EditText3 As SAPbouiCOM.EditText
        Private WithEvents EditText4 As SAPbouiCOM.EditText
        Private WithEvents StaticText3 As SAPbouiCOM.StaticText

#End Region

        Private Sub OnCustomInitialize()
            Try
                objform = objaddon.objapplication.Forms.GetForm("AdjTran", Me.FormCount)
                'objform = objaddon.objapplication.Forms.ActiveForm
                objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                bModal = True
                oSelectDT.Clear()
                If oSelectDT.Columns.Count = 0 Then
                    oSelectDT.Columns.Add("paytot", GetType(Double))
                    oSelectDT.Columns.Add("#", GetType(String))
                End If
                'If objRecoDT.Rows.Count > 0 Then objRecoDT.Clear()
                objActDT.Clear()
                If objActDT.Columns.Count = 0 Then
                    For iCol As Integer = 0 To Matrix0.Columns.Count - 1
                        If iCol <> 1 Then
                            If Matrix0.Columns.Item(iCol).UniqueID = "paytot" Then
                                objActDT.Columns.Add(Matrix0.Columns.Item(iCol).UniqueID, GetType(Double))
                            Else
                                objActDT.Columns.Add(Matrix0.Columns.Item(iCol).UniqueID)
                            End If

                        End If
                    Next
                    objActDT.Columns.Add("Row")
                    objActDT.Columns.Add("TBranchC")
                    objActDT.Columns.Add("TBranchN")
                End If
                'Matrix0.Columns.Item("tranline").Visible = False
                'Matrix0.Columns.Item("object").Visible = False
                'Matrix0.Columns.Item("debcred").Visible = False
            Catch ex As Exception
                objform.Freeze(False)
            End Try
        End Sub

        Public Function Get_InternalReconciliation_Adjustment_Query(ByVal IntRecoDT As DataRow, ByVal DocDate As String, ByVal CardCode As String, ByVal TransId As String, ByVal RecAmt As Double, ByVal BranchCode As String, ByVal BranchName As String, ByVal TranCard As String) As String
            Try
                'Dim paytotsum As String = IntRecoDT.Item("paytot").ToString ' objActualDT.Compute("SUM(paytot)", "").ToString
                Dim DebCred As String = IIf(IntRecoDT.Item("paytot") < 0, "C", "D")
                If objaddon.HANA Then
                    strSQL = "SELECT ROW_NUMBER() OVER (ORDER BY A.""BP Code"",A.""Doc Date"") AS ""LineId"",* FROM  "
                    strSQL += vbCrLf + "(SELECT 'N' AS ""Selected"",T1.""TransId"","
                    strSQL += vbCrLf + "CASE WHEN T1.""TransType""='13' THEN 'IN' WHEN T1.""TransType""='14' THEN 'CN' WHEN T1.""TransType""='203' or T1.""TransType""='204' THEN 'DT' WHEN T1.""TransType""='18' THEN 'PU' WHEN T1.""TransType""='19' THEN 'PC'"
                    strSQL += vbCrLf + "WHEN T1.""TransType""='24' THEN 'RC' WHEN T1.""TransType""='46' THEN 'PS' Else 'JE' END AS ""Origin"",T1.""BaseRef"" AS ""DocNum"","
                    strSQL += vbCrLf + "T1.""Line_ID"",T1.""DebCred""," 'CASE WHEN ""FCCurrency"" IS NULL THEN (SELECT ""MainCurncy"" FROM OADM) ELSE ""FCCurrency"" END AS ""Doc Currency"","
                    strSQL += vbCrLf + "T1.""TransType"" AS ""ObjType"",T1.""ShortName"" AS ""BP Code"",(SELECT ""CardName"" FROM OCRD where ""CardCode""=T1.""ShortName"") as ""BP Name"","
                    strSQL += vbCrLf + "T0.""RefDate"" AS ""Doc Date"",CASE WHEN T1.""Credit""<>0  THEN  -T1.""Credit"" ELSE T1.""Debit"" END AS ""Document Total"","
                    strSQL += vbCrLf + "CASE WHEN T1.""BalDueCred""<>0  THEN -T1.""BalDueCred"" ELSE T1.""BalDueDeb"" END AS ""Balance"","
                    strSQL += vbCrLf + "CASE WHEN T1.""BalDueCred""<>0  THEN -T1.""BalDueCred"" ELSE T1.""BalDueDeb"" END AS ""Amount to Reconcile"","
                    strSQL += vbCrLf + "T1.""BPLId"" ""Branch Id"",(SELECT ""BPLName"" FROM OBPL WHERE ""BPLId""=T1.""BPLId"") AS ""Branch Name"","
                    strSQL += vbCrLf + "T1.""LineMemo"" as ""LineMemo"",T1.""CreatedBy"" as ""Document Entry"","
                    strSQL += vbCrLf + "CASE WHEN T1.""TransType""='13' THEN 'A/R Invoice' WHEN T1.""TransType""='14' THEN 'A/R Credit Memo' WHEN T1.""TransType""='203' THEN 'A/R DownPayment' WHEN T1.""TransType""='18' THEN 'A/P Invoice'"
                    strSQL += vbCrLf + "WHEN T1.""TransType""='19' THEN 'A/R Credit Memo' WHEN T1.""TransType""='24' THEN 'Incoming Payment' WHEN T1.""TransType""='46' THEN 'Outgoing Payment' ELSE 'Journal Entry' END AS ""Doc Type"","
                    strSQL += vbCrLf + "T0.""Ref1"",T0.""Ref2"",T0.""Ref3"",(Select ""CardType"" from OCRD where ""CardCode""=T1.""ShortName"") as ""Card Type"""
                    strSQL += vbCrLf + "FROM OJDT T0 join JDT1 T1 ON T0.""TransId""=T1.""TransId"" where T1.""DprId"" is null"
                    strSQL += vbCrLf + ") A "
                    strSQL += vbCrLf + "WHERE A.""Doc Date""<='" & DocDate & "' and A.""Branch Id"" in (Select T0.""BPLId"" from OBPL T0 join USR6 T1 on T0.""BPLId""=T1.""BPLId"" where T1.""UserCode""='" & objaddon.objcompany.UserName & "' and T0.""Disabled""<>'Y') and A.""BP Code"" In (" & CardCode & ") and A.""Balance""<>0"
                    strSQL += vbCrLf + "and A.""TransId""<>" & TransId & ""
                    strSQL += vbCrLf + "and ""DebCred""<>'" & DebCred & "'" '(select distinct B.""DebCred"" from JDT1 B where B.""TransId""=" & TransId & " and B.""ShortName"" in ('" & TranCard & "'))
                    strSQL += vbCrLf + "ORDER BY A.""BP Code"",A.""Doc Date"""
                Else

                End If
                'objRecoDR = IntRecoDT
                'objRecoDT.Clear()
                objRecoDT = IntRecoDT.Table
                objform.DataSources.DataTables.Item("DT_0").Clear()
                objDTable = objform.DataSources.DataTables.Item("DT_0")
                objDTable.Clear()
                objDTable.ExecuteQuery(strSQL)
                Matrix0.Clear()
                Matrix0.LoadFromDataSourceEx()
                Matrix0.AutoResizeColumns()
                EditText0.Value = RecAmt
                EditText2.Value = BranchCode
                EditText3.Value = BranchName
                EditText4.Value = IntRecoDT.Item("Row")
                'Grid0.Columns.Item("Selected").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
                'Dim col As SAPbouiCOM.EditTextColumn
                'col = Grid0.Columns.Item("Amount to Reconcile")
                'col.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'For GridCols As Integer = 0 To Grid0.Columns.Count - 1
                '    If Grid0.Columns.Item(GridCols).UniqueID = "Selected" Then
                '        Grid0.Columns.Item("Selected").Editable = True
                '    Else
                '        Grid0.Columns.Item(GridCols).Editable = False
                '    End If
                'Next
                'Grid0.Columns.Item("AmounttoReconcile").Editable = True
                ''Grid0.RowHeaders.SetText(1, "1")
                'Grid0.RowHeaders.TitleObject.Caption = "#"
                'Dim oGC As SAPbouiCOM.GridColumn
                'Dim oEditGC As SAPbouiCOM.EditTextColumn
                'oGC = Grid0.Columns.Item(2)
                'oGC.Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                'oEditGC = oGC
                'Dim oST As SAPbouiCOM.BoColumnSumType = oEditGC.ColumnSetting.SumType
                'oEditGC.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                Return strSQL
            Catch ex As Exception
                Return ""
            End Try
        End Function

        Private Sub Matrix0_LinkPressedBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Matrix0.LinkPressedBefore
            Dim ColItem As SAPbouiCOM.Column = Matrix0.Columns.Item("origin")
            Dim objlink As SAPbouiCOM.LinkedButton = ColItem.ExtendedObject
            Dim oForm As SAPbouiCOM.Form
            Try
                Select Case pVal.ColUID
                    Case "origin"
                        If Matrix0.Columns.Item("object").Cells.Item(pVal.Row).Specific.String = "13" Then
                            objaddon.objapplication.Menus.Item("2053").Activate()  'AR Invoice
                            oForm = objaddon.objapplication.Forms.ActiveForm
                            oForm.Freeze(True)
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            oForm.Items.Item("8").Specific.String = Matrix0.Columns.Item("originno").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("10").Specific.String = Matrix0.Columns.Item("date").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oForm.Freeze(False)
                            BubbleEvent = False
                        ElseIf Matrix0.Columns.Item("object").Cells.Item(pVal.Row).Specific.String = "14" Then
                            objaddon.objapplication.Menus.Item("2055").Activate()  'AR Credit Memo
                            oForm = objaddon.objapplication.Forms.ActiveForm
                            oForm.Freeze(True)
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            oForm.Items.Item("8").Specific.String = Matrix0.Columns.Item("originno").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("10").Specific.String = Matrix0.Columns.Item("date").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oForm.Freeze(False)
                            BubbleEvent = False
                        ElseIf Matrix0.Columns.Item("object").Cells.Item(pVal.Row).Specific.String = "18" Then
                            objaddon.objapplication.Menus.Item("2308").Activate()  'AP Invoice
                            oForm = objaddon.objapplication.Forms.ActiveForm
                            oForm.Freeze(True)
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            oForm.Items.Item("8").Specific.String = Matrix0.Columns.Item("originno").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("10").Specific.String = Matrix0.Columns.Item("date").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oForm.Freeze(False)
                            BubbleEvent = False
                        ElseIf Matrix0.Columns.Item("object").Cells.Item(pVal.Row).Specific.String = "19" Then
                            objaddon.objapplication.Menus.Item("2309").Activate()  'AP Credit Memo
                            oForm = objaddon.objapplication.Forms.ActiveForm
                            oForm.Freeze(True)
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            oForm.Items.Item("8").Specific.String = Matrix0.Columns.Item("originno").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("10").Specific.String = Matrix0.Columns.Item("date").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oForm.Freeze(False)
                            BubbleEvent = False
                        ElseIf Matrix0.Columns.Item("object").Cells.Item(pVal.Row).Specific.String = "203" Then
                            objaddon.objapplication.Menus.Item("2071").Activate()  'AR DownPayment
                            oForm = objaddon.objapplication.Forms.ActiveForm
                            oForm.Freeze(True)
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            oForm.Items.Item("8").Specific.String = Matrix0.Columns.Item("originno").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("10").Specific.String = Matrix0.Columns.Item("date").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oForm.Freeze(False)
                            BubbleEvent = False
                        ElseIf Matrix0.Columns.Item("object").Cells.Item(pVal.Row).Specific.String = "204" Then
                            objaddon.objapplication.Menus.Item("2317").Activate()  'AP DownPayment
                            oForm = objaddon.objapplication.Forms.ActiveForm
                            oForm.Freeze(True)
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            oForm.Items.Item("8").Specific.String = Matrix0.Columns.Item("originno").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("10").Specific.String = Matrix0.Columns.Item("date").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oForm.Freeze(False)
                            BubbleEvent = False
                        ElseIf Matrix0.Columns.Item("object").Cells.Item(pVal.Row).Specific.String = "24" Then
                            objaddon.objapplication.Menus.Item("2817").Activate()  'Incoming Payment
                            oForm = objaddon.objapplication.Forms.ActiveForm
                            oForm.Freeze(True)
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            oForm.Items.Item("3").Specific.String = Matrix0.Columns.Item("originno").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("10").Specific.String = Matrix0.Columns.Item("date").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oForm.Freeze(False)
                            BubbleEvent = False
                        ElseIf Matrix0.Columns.Item("object").Cells.Item(pVal.Row).Specific.String = "46" Then
                            objaddon.objapplication.Menus.Item("2818").Activate()  'Outgoing Payment
                            oForm = objaddon.objapplication.Forms.ActiveForm
                            oForm.Freeze(True)
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            oForm.Items.Item("3").Specific.String = Matrix0.Columns.Item("originno").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("10").Specific.String = Matrix0.Columns.Item("date").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oForm.Freeze(False)
                            BubbleEvent = False
                        Else 'If Matrix0.Columns.Item("object").Cells.Item(pVal.Row).Specific.String = "30" Then
                            'objlink.LinkedObjectType = "30" ' Matrix0.Columns.Item("object").Cells.Item(pVal.Row).Specific.String
                            'objlink.Item.LinkTo = "trannum"
                            objaddon.objapplication.Menus.Item("1540").Activate()  'Outgoing Payment
                            oForm = objaddon.objapplication.Forms.ActiveForm
                            oForm.Freeze(True)
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            oForm.Items.Item("5").Specific.String = Matrix0.Columns.Item("trannum").Cells.Item(pVal.Row).Specific.String
                            'oForm.Items.Item("10").Specific.String = Matrix0.Columns.Item("date").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oForm.Freeze(False)
                            BubbleEvent = False
                        End If
                End Select

            Catch ex As Exception
                oForm.Freeze(False)
                oForm = Nothing
            End Try
        End Sub

        Private Sub Form_ResizeAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                Matrix0.AutoResizeColumns()
            Catch ex As Exception
            End Try

        End Sub

        Private Function build_Matrix_DataTable(ByVal sKeyFieldID As String) As DataTable
            Dim objcheckbox As SAPbouiCOM.CheckBox
            Try
                Dim oDT As New DataTable
                'Add all of the columns by unique ID to the DataTable
                For iCol As Integer = 0 To Matrix0.Columns.Count - 1
                    'Skip invisible columns
                    'If oMatrix.Columns.Item(iCol).Visible = False Then Continue For
                    If iCol <> 1 Then
                        oDT.Columns.Add(Matrix0.Columns.Item(iCol).UniqueID)
                    End If
                Next
                'Now, add all of the data into the DataTable
                For iRow As Integer = 1 To Matrix0.VisualRowCount
                    objcheckbox = Matrix0.Columns.Item("select").Cells.Item(iRow).Specific
                    If objcheckbox.Checked = True Then
                        Dim oRow As DataRow = oDT.NewRow
                        For iCol As Integer = 0 To Matrix0.Columns.Count - 1
                            'If oMatrix.Columns.Item(iCol).Visible = False Then Continue For
                            If iCol <> 1 Then
                                oRow.Item(Matrix0.Columns.Item(iCol).UniqueID) = Matrix0.Columns.Item(iCol).Cells.Item(iRow).Specific.Value
                            End If
                        Next
                        'If the Key field has no value, then the row is empty, skip adding it.
                        If oRow(sKeyFieldID).ToString.Trim = 0 Then Continue For
                        oDT.Rows.Add(oRow)
                    End If
                Next

                Return oDT
            Catch ex As Exception
                Return Nothing
            End Try

        End Function

        Private Function Matrix_DataTable(ByVal Row As Integer, ByVal ColName As String) As DataTable
            Try
                Dim objcheckbox As SAPbouiCOM.CheckBox
                Dim DataFlag As Boolean
                objcheckbox = Matrix0.Columns.Item("select").Cells.Item(Row).Specific
                Dim oRow As DataRow = objActDT.NewRow
                If objcheckbox.Checked = True Then
                    If objActDT.Rows.Count > 0 Then
                        For DTRow As Integer = 0 To objActDT.Rows.Count - 1
                            If objActDT.Rows(DTRow)("#").ToString = Matrix0.Columns.Item("#").Cells.Item(Row).Specific.Value Then
                                If ColName <> "" Then
                                    objActDT.Rows(DTRow)(Matrix0.Columns.Item(ColName).UniqueID) = Matrix0.Columns.Item(ColName).Cells.Item(Row).Specific.Value
                                    DataFlag = True
                                    Exit For
                                End If
                            End If
                        Next
                        If DataFlag = False Then
                            For iCol As Integer = 0 To Matrix0.Columns.Count - 1
                                If iCol <> 1 Then
                                    oRow.Item(Matrix0.Columns.Item(iCol).UniqueID) = Matrix0.Columns.Item(iCol).Cells.Item(Row).Specific.Value
                                End If
                            Next
                            oRow.Item("TBranchC") = EditText2.Value
                            oRow.Item("TBranchN") = EditText3.Value
                            oRow.Item("Row") = EditText4.Value
                            objActDT.Rows.Add(oRow)
                        End If
                    Else
                        For iCol As Integer = 0 To Matrix0.Columns.Count - 1
                            If iCol <> 1 Then
                                oRow.Item(Matrix0.Columns.Item(iCol).UniqueID) = Matrix0.Columns.Item(iCol).Cells.Item(Row).Specific.Value
                            End If
                        Next
                        oRow.Item("TBranchC") = EditText2.Value
                        oRow.Item("TBranchN") = EditText3.Value
                        oRow.Item("Row") = EditText4.Value
                        objActDT.Rows.Add(oRow)
                    End If

                Else
                    For DTRow As Integer = 0 To objActDT.Rows.Count - 1
                        If objActDT.Rows(DTRow)("#").ToString = Matrix0.Columns.Item("#").Cells.Item(Row).Specific.Value Then
                            objActDT.Rows(DTRow).Delete()
                            Exit For
                        End If
                    Next
                End If

                Return objActDT
            Catch ex As Exception
                Return Nothing
            End Try
        End Function

        Private Sub Calc_Total(ByVal Row As Integer)
            Try
                Dim value As Double
                Dim DataFlag As Boolean
                Dim objcheckbox As SAPbouiCOM.CheckBox
                Dim oRow As DataRow = oSelectDT.NewRow
                objcheckbox = Matrix0.Columns.Item("select").Cells.Item(Row).Specific

                If objcheckbox.Checked = True Then
                    If oSelectDT.Rows.Count > 0 Then
                        For DTRow As Integer = 0 To oSelectDT.Rows.Count - 1
                            If oSelectDT.Rows(DTRow)("#").ToString = Matrix0.Columns.Item("#").Cells.Item(Row).Specific.Value Then
                                oSelectDT.Rows(DTRow)("paytot") = Matrix0.Columns.Item("paytot").Cells.Item(Row).Specific.Value
                                DataFlag = True
                                Exit For
                            End If
                        Next
                        If DataFlag = False Then
                            oRow.Item(Matrix0.Columns.Item("paytot").UniqueID) = Matrix0.Columns.Item("paytot").Cells.Item(Row).Specific.Value
                            oRow.Item(Matrix0.Columns.Item("#").UniqueID) = Matrix0.Columns.Item("#").Cells.Item(Row).Specific.Value
                            oSelectDT.Rows.Add(oRow)
                        End If
                    Else
                        oRow.Item(Matrix0.Columns.Item("paytot").UniqueID) = Matrix0.Columns.Item("paytot").Cells.Item(Row).Specific.Value
                        oRow.Item(Matrix0.Columns.Item("#").UniqueID) = Matrix0.Columns.Item("#").Cells.Item(Row).Specific.Value
                        oSelectDT.Rows.Add(oRow)
                    End If
                Else
                    For DTRow As Integer = 0 To oSelectDT.Rows.Count - 1
                        If oSelectDT.Rows(DTRow)("#").ToString = Matrix0.Columns.Item("#").Cells.Item(Row).Specific.Value Then
                            oSelectDT.Rows(DTRow).Delete()
                            Exit For
                        End If
                    Next
                End If

                For i As Integer = 0 To oSelectDT.Rows.Count - 1
                    If Val(oSelectDT.Rows(i)("paytot").ToString) <> 0 Then
                        value = Math.Round(value + CDbl(oSelectDT.Rows(i)("paytot")), SumRound)
                    End If
                Next
                value = EditText0.Value + value
                EditText1.Value = value
                objform.Update()
                'odbdsHeader.SetValue("U_Total", 0, value) 'Total
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Matrix0_PressedAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.PressedAfter
            Try
                objCheck = Matrix0.Columns.Item("select").Cells.Item(pVal.Row).Specific
                objform.Freeze(True)
                If pVal.ColUID = "select" Then
                    If objCheck.Checked = True Then
                        'Matrix1.SelectRow(pVal.Row, True, True)
                        Matrix0.CommonSetting.SetRowBackColor(pVal.Row, Color.PeachPuff.ToArgb)
                    Else
                        'Matrix1.SelectRow(pVal.Row, False, True)
                        Matrix0.CommonSetting.SetRowBackColor(pVal.Row, Matrix0.Item.BackColor)
                        Matrix0.Columns.Item("baldue").Cells.Item(pVal.Row).Specific.String = CDbl(Matrix0.Columns.Item("baldue").Cells.Item(pVal.Row).Specific.String)
                        'Matrix0.SetCellWithoutValidation(pVal.Row, "paytot", CDbl(Matrix0.Columns.Item("baldue").Cells.Item(pVal.Row).Specific.String))
                    End If

                    'Calculate_Total()
                    Calc_Total(pVal.Row)
                    Matrix_DataTable(pVal.Row, pVal.ColUID)
                End If
            Catch ex As Exception
            Finally
                objform.Freeze(False)
            End Try

        End Sub

        Private Sub Matrix0_ValidateBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Matrix0.ValidateBefore
            Try
                'If pVal.ItemChanged = False Then Exit Sub
                If pVal.InnerEvent = True Then Exit Sub
                Dim Balance, PayTotal As Double
                objform.Freeze(True)
                'Dim ActTotal As Double
                If Val(Matrix0.Columns.Item("baldue").Cells.Item(pVal.Row).Specific.String) <> 0 Then Balance = CDbl(Matrix0.Columns.Item("baldue").Cells.Item(pVal.Row).Specific.String) Else Balance = 0
                If Val(Matrix0.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String) <> 0 Then PayTotal = CDbl(Matrix0.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String) Else PayTotal = 0
                'If Val(Matrix0.Columns.Item("pay").Cells.Item(pVal.Row).Specific.String) > 0 Then ActTotal = CDbl(Matrix0.Columns.Item("pay").Cells.Item(pVal.Row).Specific.String) Else ActTotal = 0
                Select Case pVal.ColUID
                    Case "paytot"
                        If pVal.InnerEvent = False Then
                            If PayTotal <= 0 Then
                                PayTotal = -CDbl(Matrix0.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String)
                                Balance = -CDbl(Matrix0.Columns.Item("baldue").Cells.Item(pVal.Row).Specific.String)
                                If PayTotal > Balance Or PayTotal = 0 Then
                                    'Matrix0.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String = CDbl(Matrix0.Columns.Item("baldue").Cells.Item(pVal.Row).Specific.String)
                                    'Matrix0.SetCellWithoutValidation(pVal.Row, "paytot", CDbl(Matrix0.Columns.Item("baldue").Cells.Item(pVal.Row).Specific.String))
                                    Matrix0.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String = CDbl(Matrix0.Columns.Item("baldue").Cells.Item(pVal.Row).Specific.String)
                                End If
                            ElseIf PayTotal > 0 Then
                                If PayTotal > Balance Then
                                    'Matrix0.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String = CDbl(Matrix0.Columns.Item("baldue").Cells.Item(pVal.Row).Specific.String)
                                    'Matrix0.SetCellWithoutValidation(pVal.Row, "paytot", CDbl(Matrix0.Columns.Item("baldue").Cells.Item(pVal.Row).Specific.String))
                                    Matrix0.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String = CDbl(Matrix0.Columns.Item("baldue").Cells.Item(pVal.Row).Specific.String)
                                End If
                            Else
                                'Matrix0.SetCellWithoutValidation(pVal.Row, "paytot", CDbl(Matrix0.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String))
                                Matrix0.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String = CDbl(Matrix0.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String)
                            End If
                            If pVal.ItemChanged = True Then
                                objCheck = Matrix0.Columns.Item("select").Cells.Item(pVal.Row).Specific
                                objCheck.Checked = True
                                Matrix0.CommonSetting.SetRowBackColor(pVal.Row, Color.PeachPuff.ToArgb)
                                'Matrix0.SetCellWithoutValidation(pVal.Row, "paytot", CDbl(Matrix0.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String))
                                Matrix0.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String = CDbl(Matrix0.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String)
                            End If
                            'Matrix0.Columns.Item("pay").Cells.Item(pVal.Row).Specific.String = Matrix0.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String
                            Calc_Total(pVal.Row)
                        End If
                End Select

                If pVal.ItemChanged = True Then
                    Matrix_DataTable(pVal.Row, pVal.ColUID)
                End If
            Catch ex As Exception
            Finally
                objform.Freeze(False)
            End Try

        End Sub

        Private Sub Button0_ClickBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Button0.ClickBefore
            Try
                If CDbl(EditText1.Value) <> 0 Then 'Reco Amount - Total Amount
                    objaddon.objapplication.SetStatusBarMessage("Selected Transactions is not matching from Reconciliation amount. Please check... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    BubbleEvent = False : Exit Sub
                End If
                objRecoDT.Merge(objActDT, True)
                Dim objRecoform As SAPbouiCOM.Form
                Dim objRecMatrix As SAPbouiCOM.Matrix
                objRecoform = objaddon.objapplication.Forms.GetForm("FOITR", 1)
                objRecMatrix = objRecoform.Items.Item("mtxreco").Specific
                Dim odbdsDetails As SAPbouiCOM.DBDataSource
                odbdsDetails = objRecoform.DataSources.DBDataSources.Item("@MI_ITR2") '(CType(2, Object)
                LoadDT_RecoMatrix(objRecoDT, odbdsDetails, objRecMatrix)
                objRecMatrix = objRecoform.Items.Item("mtxcont").Specific
                'Dim vv As Integer = CInt(objRecoDT.Rows(0)("#").ToString)
                objRecMatrix.Columns.Item("jeadj").Cells.Item(CInt(objRecoDT.Rows(0)("#").ToString)).Specific.String = "Y"
                objPayIntRecoDT = objRecoDT
                objform.Close()
            Catch ex As Exception

            End Try
        End Sub

        Private Sub LoadDT_RecoMatrix(ByVal FinalDT As DataTable, ByVal odbdsDetails As SAPbouiCOM.DBDataSource, ByVal Matrix1 As SAPbouiCOM.Matrix)
            Try
                Dim IRow As Integer = 0
                If Matrix1.VisualRowCount = 0 Then IRow = 1 Else IRow = Matrix1.VisualRowCount

                For DTRow As Integer = 0 To FinalDT.Rows.Count - 1
                    If FinalDT.Rows(DTRow)("trannum").ToString <> "" Then
                        Matrix1.AddRow()
                        Matrix1.GetLineData(Matrix1.VisualRowCount)
                        odbdsDetails.SetValue("LineId", 0, IRow) '
                        odbdsDetails.SetValue("U_Row", 0, FinalDT.Rows(DTRow)("Row").ToString)
                        odbdsDetails.SetValue("U_TransId", 0, FinalDT.Rows(DTRow)("trannum").ToString)
                        odbdsDetails.SetValue("U_TLine", 0, FinalDT.Rows(DTRow)("tranline").ToString)
                        odbdsDetails.SetValue("U_DebCred", 0, FinalDT.Rows(DTRow)("debcred").ToString)
                        odbdsDetails.SetValue("U_CardType", 0, FinalDT.Rows(DTRow)("cardtype").ToString)
                        odbdsDetails.SetValue("U_Origin", 0, FinalDT.Rows(DTRow)("origin").ToString)
                        odbdsDetails.SetValue("U_OriginNo", 0, FinalDT.Rows(DTRow)("originno").ToString)
                        odbdsDetails.SetValue("U_Object", 0, FinalDT.Rows(DTRow)("object").ToString)
                        odbdsDetails.SetValue("U_CardCode", 0, FinalDT.Rows(DTRow)("cardc").ToString)
                        odbdsDetails.SetValue("U_CardName", 0, FinalDT.Rows(DTRow)("cardn").ToString)
                        odbdsDetails.SetValue("U_DocDate", 0, FinalDT.Rows(DTRow)("date").ToString)
                        odbdsDetails.SetValue("U_Total", 0, FinalDT.Rows(DTRow)("amount").ToString)
                        odbdsDetails.SetValue("U_BalDue", 0, FinalDT.Rows(DTRow)("baldue").ToString)
                        odbdsDetails.SetValue("U_PayTotal", 0, FinalDT.Rows(DTRow)("paytot").ToString)
                        odbdsDetails.SetValue("U_BranchId", 0, FinalDT.Rows(DTRow)("branchc").ToString)
                        odbdsDetails.SetValue("U_BranchNam", 0, FinalDT.Rows(DTRow)("branchn").ToString)
                        odbdsDetails.SetValue("U_TBranchId", 0, FinalDT.Rows(DTRow)("TBranchC").ToString)
                        odbdsDetails.SetValue("U_TBranchNam", 0, FinalDT.Rows(DTRow)("TBranchN").ToString)
                        Matrix1.SetLineData(Matrix1.VisualRowCount)
                        IRow += 1
                    End If
                Next
                Matrix1.AutoResizeColumns()
            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage("Loading Issue: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub


    End Class
End Namespace
