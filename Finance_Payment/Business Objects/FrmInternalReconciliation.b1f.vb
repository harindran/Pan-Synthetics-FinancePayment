Option Strict Off
Option Explicit On

Imports System.Drawing
Imports System.Windows.Forms
Imports SAPbobsCOM
Imports SAPbouiCOM.Framework

Namespace Finance_Payment
    <FormAttribute("FOITR", "Business Objects/FrmInternalReconciliation.b1f")>
    Friend Class FrmInternalReconciliation
        Inherits UserFormBase
        Public WithEvents objform As SAPbouiCOM.Form
        Private WithEvents pCFL As SAPbouiCOM.ISBOChooseFromListEventArg
        Public WithEvents odbdsDetails As SAPbouiCOM.DBDataSource
        Public WithEvents odbdsHeader As SAPbouiCOM.DBDataSource
        Public WithEvents odbdsDetails1 As SAPbouiCOM.DBDataSource
        Dim FormCount As Integer = 0
        Dim objRs As SAPbobsCOM.Recordset
        Dim strSQL As String
        Public Shared objFDT As New DataTable
        Public Shared oSelectedDT As New DataTable
        Private WithEvents objCheck As SAPbouiCOM.CheckBox
        Private Shared objActualDT As New DataTable

        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("1").Specific, SAPbouiCOM.Button)
            Me.Button1 = CType(Me.GetItem("2").Specific, SAPbouiCOM.Button)
            Me.Matrix1 = CType(Me.GetItem("mtxcont").Specific, SAPbouiCOM.Matrix)
            Me.StaticText0 = CType(Me.GetItem("ldocdate").Specific, SAPbouiCOM.StaticText)
            Me.EditText0 = CType(Me.GetItem("tdocdate").Specific, SAPbouiCOM.EditText)
            Me.StaticText1 = CType(Me.GetItem("ldocnum").Specific, SAPbouiCOM.StaticText)
            Me.ComboBox0 = CType(Me.GetItem("Series").Specific, SAPbouiCOM.ComboBox)
            Me.EditText1 = CType(Me.GetItem("t_docnum").Specific, SAPbouiCOM.EditText)
            Me.Folder0 = CType(Me.GetItem("fldrcont").Specific, SAPbouiCOM.Folder)
            Me.Folder1 = CType(Me.GetItem("fldrreco").Specific, SAPbouiCOM.Folder)
            Me.StaticText3 = CType(Me.GetItem("lremark").Specific, SAPbouiCOM.StaticText)
            Me.EditText3 = CType(Me.GetItem("tremark").Specific, SAPbouiCOM.EditText)
            Me.StaticText4 = CType(Me.GetItem("ltotdue").Specific, SAPbouiCOM.StaticText)
            Me.EditText4 = CType(Me.GetItem("ttotdue").Specific, SAPbouiCOM.EditText)
            Me.StaticText2 = CType(Me.GetItem("lpaydate").Specific, SAPbouiCOM.StaticText)
            Me.EditText2 = CType(Me.GetItem("tpaydate").Specific, SAPbouiCOM.EditText)
            Me.Matrix0 = CType(Me.GetItem("mtxreco").Specific, SAPbouiCOM.Matrix)
            Me.StaticText5 = CType(Me.GetItem("ltransid").Specific, SAPbouiCOM.StaticText)
            Me.EditText5 = CType(Me.GetItem("ttransid").Specific, SAPbouiCOM.EditText)
            Me.LinkedButton0 = CType(Me.GetItem("lnkje").Specific, SAPbouiCOM.LinkedButton)
            Me.StaticText6 = CType(Me.GetItem("lstat").Specific, SAPbouiCOM.StaticText)
            Me.ComboBox1 = CType(Me.GetItem("cmbstat").Specific, SAPbouiCOM.ComboBox)
            Me.StaticText7 = CType(Me.GetItem("lrevje").Specific, SAPbouiCOM.StaticText)
            Me.EditText6 = CType(Me.GetItem("trevje").Specific, SAPbouiCOM.EditText)
            Me.LinkedButton1 = CType(Me.GetItem("lnkrevje").Specific, SAPbouiCOM.LinkedButton)
            Me.EditText7 = CType(Me.GetItem("txtentry").Specific, SAPbouiCOM.EditText)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler DataAddAfter, AddressOf Me.Form_DataAddAfter
            AddHandler LoadAfter, AddressOf Me.Form_LoadAfter
            AddHandler ResizeAfter, AddressOf Me.Form_ResizeAfter
            AddHandler DataLoadAfter, AddressOf Me.Form_DataLoadAfter
            AddHandler DataAddBefore, AddressOf Me.Form_DataAddBefore

        End Sub

        Private Sub OnCustomInitialize()
            Try
                'objform = objaddon.objapplication.Forms.GetForm("FOITR", Me.FormCount)
                objform.Freeze(True)
                Field_Setup()
                If objaddon.HANA Then
                    BranchSplitReconciliation = objaddon.objglobalmethods.getSingleValue("select ""U_BranchSplit"" from OADM")
                Else
                    BranchSplitReconciliation = objaddon.objglobalmethods.getSingleValue("select U_BranchSplit from OADM")
                End If
                If BranchSplitReconciliation = "Y" Then 'With splitup JE
                    Matrix1.Columns.Item("jeno").Visible = False
                    Matrix1.Columns.Item("recono").Visible = False
                    Matrix1.Columns.Item("tranadj").Visible = True
                    Matrix1.Columns.Item("jeadj").Visible = True
                    Matrix1.Columns.Item("revjeno").Visible = False
                    Folder1.Item.Visible = True
                    If ComboBox1.Selected.Value = "O" Then
                        EditText6.Item.Visible = False
                        StaticText7.Item.Visible = False
                        LinkedButton1.Item.Visible = False
                    Else
                        EditText6.Item.Visible = True
                        StaticText7.Item.Visible = True
                        LinkedButton1.Item.Visible = True
                    End If
                Else
                    Matrix1.Columns.Item("jeno").Visible = True
                    Matrix1.Columns.Item("recono").Visible = True
                    Matrix1.Columns.Item("tranadj").Visible = False
                    Matrix1.Columns.Item("jeadj").Visible = False
                    Matrix1.Columns.Item("revjeno").Visible = False
                    Folder1.Item.Visible = False
                    EditText5.Item.Visible = False
                    StaticText5.Item.Visible = False
                    LinkedButton0.Item.Visible = False
                    EditText6.Item.Visible = False
                    StaticText7.Item.Visible = False
                    LinkedButton1.Item.Visible = False
                End If
                Matrix1.Columns.Item("#").Visible = False
                Matrix0.Columns.Item("cardtype").Visible = False
                Matrix0.Columns.Item("row").Visible = False
                Matrix0.Columns.Item("trannum").Visible = False
                Matrix0.Columns.Item("tranline").Visible = False
                Matrix0.Columns.Item("object").Visible = False
                Matrix0.Columns.Item("debcred").Visible = False
                Matrix1.CommonSetting.FixedColumnsCount = 2
                Matrix0.AutoResizeColumns()
                objform.Freeze(False)
                If Link_Value <> "" And Link_objtype = "MIOITR" Then
                    objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                    EditText7.Item.Enabled = True
                    EditText7.Value = Link_Value
                    objform.ActiveItem = "tremark"
                    EditText7.Item.Enabled = False
                    objform.Items.Item("1").Click()
                    Link_Value = "" : Link_objtype = ""
                    Exit Sub
                End If
                'objform = objaddon.objapplication.Forms.ActiveForm
                objform.Freeze(True)
                odbdsHeader = objform.DataSources.DBDataSources.Item("@MI_OITR") 'CType(0, Object)
                odbdsDetails = objform.DataSources.DBDataSources.Item("@MI_ITR1") 'CType(1, Object)
                odbdsDetails1 = objform.DataSources.DBDataSources.Item("@MI_ITR2") 'CType(2, Object)
                If objaddon.objapplication.Menus.Item("6913").Checked = True Then
                    objaddon.objapplication.SendKeys("^+U")
                End If
                objaddon.objglobalmethods.LoadSeries(objform, odbdsHeader, "MIOITR")
                Dim FSize As Size = TextRenderer.MeasureText(Folder0.Caption, New Font("Arial", 12.0F))
                Folder0.Item.Width = FSize.Width + 20
                FSize = TextRenderer.MeasureText(Folder1.Caption, New Font("Arial", 12.0F))
                Folder1.Item.Width = FSize.Width + 20
                'CheckBox0.Item.Height = CheckBox0.Item.Height + 4
                'CheckBox0.Item.Width = CheckBox0.Item.Width + 30
                If oSelectedDT.Rows.Count > 0 Then oSelectedDT.Clear()
                If oSelectedDT.Columns.Count = 0 Then
                    oSelectedDT.Columns.Add("paytot", GetType(Double))
                    oSelectedDT.Columns.Add("#", GetType(String))
                End If
                If objActualDT.Rows.Count > 0 Then objActualDT.Clear()
                If objActualDT.Columns.Count = 0 Then
                    For iCol As Integer = 0 To Matrix1.Columns.Count - 1
                        If iCol <> 1 And iCol <> 18 Then
                            If Matrix1.Columns.Item(iCol).UniqueID = "paytot" Then
                                objActualDT.Columns.Add(Matrix1.Columns.Item(iCol).UniqueID, GetType(Double))
                            Else
                                objActualDT.Columns.Add(Matrix1.Columns.Item(iCol).UniqueID)
                            End If
                        End If
                    Next
                    objActualDT.Columns.Add("Row")
                End If
                objform.Items.Item("tdocdate").Specific.string = Now.Date.ToString("yyyyMMdd")
                objform.Items.Item("tpaydate").Specific.string = PayInitDate.ToString("yyyyMMdd")
                objform.Items.Item("tremark").Specific.String = "Created By " & objaddon.objcompany.UserName & " on " & Now.ToString
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRs.DoQuery("select distinct T0.""BnkChgAct"" as ""BCGAcct"",T0.""LinkAct_3"" as ""CahAcct"",""LinkAct_24"" as ""Rounding"",""GLGainXdif"",""GLLossXdif"",""ExDiffAct"" " &
                              ",(Select ""SumDec"" from OADM) as ""SumDec"",(Select ""RateDec"" from OADM) as ""RateDec""" &
                              "from OACP T0 left join OFPR T1 on T1.""Category""=T0.""PeriodCat"" where T0.""PeriodCat""=(Select ""Category"" from OFPR where CURRENT_DATE Between ""F_RefDate"" and ""T_RefDate"")")
                If objRs.RecordCount > 0 Then
                    If objRs.Fields.Item(6).Value.ToString <> "" Then SumRound = objRs.Fields.Item(6).Value.ToString
                    If objRs.Fields.Item(7).Value.ToString <> "" Then RateRound = objRs.Fields.Item(7).Value.ToString
                End If
                Folder0.Item.Click()
                If Not LoadData(Query) Then
                    objform.Close()
                End If
                Query = ""
                objform.Freeze(False)
            Catch ex As Exception
                objform.Freeze(False)
            End Try
        End Sub

#Region "Fields"

        Private WithEvents Button0 As SAPbouiCOM.Button
        Private WithEvents Button1 As SAPbouiCOM.Button
        Private WithEvents Matrix1 As SAPbouiCOM.Matrix
        Private WithEvents StaticText0 As SAPbouiCOM.StaticText
        Private WithEvents EditText0 As SAPbouiCOM.EditText
        Private WithEvents StaticText1 As SAPbouiCOM.StaticText
        Private WithEvents ComboBox0 As SAPbouiCOM.ComboBox
        Private WithEvents EditText1 As SAPbouiCOM.EditText
        Private WithEvents Folder0 As SAPbouiCOM.Folder
        Private WithEvents Folder1 As SAPbouiCOM.Folder
        Private WithEvents StaticText3 As SAPbouiCOM.StaticText
        Private WithEvents EditText3 As SAPbouiCOM.EditText
        Private WithEvents StaticText4 As SAPbouiCOM.StaticText
        Private WithEvents EditText4 As SAPbouiCOM.EditText
        Private WithEvents StaticText6 As SAPbouiCOM.StaticText
        Private WithEvents ComboBox1 As SAPbouiCOM.ComboBox
        Private WithEvents StaticText7 As SAPbouiCOM.StaticText
        Private WithEvents EditText6 As SAPbouiCOM.EditText
        Private WithEvents LinkedButton1 As SAPbouiCOM.LinkedButton
        Private WithEvents StaticText2 As SAPbouiCOM.StaticText
        Private WithEvents EditText2 As SAPbouiCOM.EditText
        Private WithEvents Matrix0 As SAPbouiCOM.Matrix
        Private WithEvents StaticText5 As SAPbouiCOM.StaticText
        Private WithEvents EditText5 As SAPbouiCOM.EditText
        Private WithEvents LinkedButton0 As SAPbouiCOM.LinkedButton
        Private WithEvents EditText7 As SAPbouiCOM.EditText

#End Region

#Region "Form Events"

        Private Sub Button0_ClickBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Button0.ClickBefore
            'Try
            '    If objform.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then Exit Sub
            '    Dim Amt As Double
            '    Dim Line As Integer = 0
            '    Dim ErrorFlag As Boolean = False
            '    If BranchSplitReconciliation = "Y" Then ' with splitup JE
            '        Try
            '            Dim lineflag As Boolean = False
            '            'If EditText5.Value <> "" Then
            '            If Matrix0.VisualRowCount > 0 Then
            '                For i As Integer = 1 To Matrix0.VisualRowCount
            '                    If Matrix0.Columns.Item("recono").Cells.Item(i).Specific.String <> "" Then
            '                        lineflag = True
            '                        Exit For
            '                    End If
            '                Next
            '            End If
            '            If lineflag = True Then Exit Sub
            '            'End If
            '            If objPayIntRecoDT.Rows.Count > 0 Then objFDT = objPayIntRecoDT 'Else objaddon.objapplication.StatusBar.SetText("Rows required for reconciling the transactions...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
            '            If objFDT.Rows.Count = 0 Then objaddon.objapplication.StatusBar.SetText("Rows required for reconciling the transactions...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
            '            Dim paytotsum As String = objFDT.Compute("SUM(paytot)", "").ToString
            '            If paytotsum <> "0" Then objaddon.objapplication.StatusBar.SetText("Reconciliation difference must be zero before reconciling...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
            '            Dim BranchDT = From dr In objFDT.AsEnumerable()
            '                           Group dr By Ph = New With {Key .DTLine = dr.Field(Of String)("Row"), Key .branch = dr.Field(Of String)("branchc")} Into drg = Group
            '                           Where drg.Sum(Function(dr) dr.Field(Of Double)("paytot")) = 0
            '                           Select New With {                        'Where drg.Sum(Function(dr) dr.Field(Of String)("paytot")) = 0
            '        .branch = Ph.branch,
            '        .line = Ph.DTLine,
            '        .LengthSum = drg.Sum(Function(dr) dr.Field(Of Double)("paytot"))
            '        }
            '            objaddon.objapplication.StatusBar.SetText("Creating transactions.Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            '            If objaddon.objcompany.InTransaction = False Then objaddon.objcompany.StartTransaction()

            '            For Each RowID In BranchDT
            '                Amt = Math.Round(CDbl(RowID.LengthSum), SumRound)
            '                If CDbl(Amt) = 0 Then
            '                    If BranchReconciliation_Consolidated(objFDT, RowID.branch.ToString, RowID.line.ToString) = False Then
            '                        ErrorFlag = True
            '                        objaddon.objapplication.StatusBar.SetText("Error occurred while reconciling the transactions...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) ': BubbleEvent = False : Exit Sub
            '                    End If
            '                End If
            '            Next
            '            Dim JEBranchDT = From dr In objFDT.AsEnumerable()
            '                             Group dr By Ph = New With {Key .DTLine = dr.Field(Of String)("Row"), Key .branch = dr.Field(Of String)("branchc")} Into drg = Group
            '                             Where drg.Sum(Function(dr) dr.Field(Of Double)("paytot")) <> 0
            '                             Select New With {
            '        .branch = Ph.branch,
            '        .line = Ph.DTLine,
            '        .LengthSum = drg.Sum(Function(dr) dr.Field(Of Double)("paytot"))
            '        }
            '            If JEBranchDT.Count > 0 Then
            '                If JournalEntry_Consolidated(objFDT) = False Then
            '                    ErrorFlag = True
            '                    objaddon.objapplication.StatusBar.SetText("Error occurred while reconciling the transactions...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) ': BubbleEvent = False : Exit Sub
            '                End If
            '            End If

            '            If ErrorFlag = True Then
            '                If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            '                Try
            '                    EditText5.Value = ""
            '                    objform.Freeze(True)
            '                    Matrix0.FlushToDataSource()
            '                    For rowNum As Integer = 0 To odbdsDetails1.Size - 1
            '                        odbdsDetails1.SetValue("U_RecoNo", rowNum, "")
            '                    Next
            '                    Matrix0.LoadFromDataSource()
            '                Catch ex As Exception
            '                Finally
            '                    objform.Freeze(False)
            '                End Try
            '                objform.Refresh()
            '                objaddon.objapplication.StatusBar.SetText("Error while reconciling the multi-branch transactions...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
            '            Else
            '                If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            '                objaddon.objapplication.StatusBar.SetText("Multi-branch Internal Reconciliations Created Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            '                objPayIntRecoDT.Clear()
            '            End If
            '        Catch ex As Exception
            '        End Try
            '    Else 'Consolidated JE
            '        Try
            '            If Not CDbl(EditText4.Value) = 0 Then objaddon.objapplication.StatusBar.SetText("Reconciliation difference must be zero before reconciling...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub

            '            If objActualDT.Rows.Count > 0 Then objFDT = objActualDT Else objFDT = build_Matrix_DataTable("paytot", Matrix1)
            '            If objFDT.Rows.Count = 0 Then objaddon.objapplication.StatusBar.SetText("Rows required for reconciling the transactions...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub

            '            Dim otherBranchDT = From dr In objFDT.AsEnumerable()
            '                                Group dr By Ph = dr.Field(Of String)("branchc") Into drg = Group
            '                                Select New With {                   'Where drg.Sum(Function(dr) dr.Field(Of String)("paytot")) = 0  'Ph <> Branch And
            '    .branch = Ph,
            '    .LengthSum = drg.Sum(Function(dr) dr.Field(Of Double)("paytot"))
            '    }
            '            '    Dim otherBranchDT = From dr In objFDT.AsEnumerable()
            '            '                        Group dr By Ph = New With {Key .branch = dr.Field(Of String)("branchc"), Key .DTLine = dr.Field(Of String)("#")} Into drg = Group
            '            '                        Select New With {                        'Where drg.Sum(Function(dr) dr.Field(Of String)("paytot")) = 0
            '            '.branch = Ph.branch,
            '            '.line = Ph.DTLine,
            '            '.LengthSum = drg.Sum(Function(dr) dr.Field(Of String)("paytot"))
            '            '}
            '            objaddon.objapplication.StatusBar.SetText("Creating transactions.Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            '            If objaddon.objcompany.InTransaction = False Then objaddon.objcompany.StartTransaction()
            '            For Each RowID In otherBranchDT
            '                Amt = Math.Round(CDbl(RowID.LengthSum), SumRound)
            '                If CDbl(Amt) = 0 Then
            '                    If BranchReconciliation(objFDT, RowID.branch.ToString()) = False Then
            '                        ErrorFlag = True
            '                    End If
            '                Else
            '                    If JournalEntry(objFDT, RowID.branch.ToString(), CDbl(Amt)) = False Then
            '                        ErrorFlag = True
            '                        objaddon.objapplication.StatusBar.SetText("Error occurred while reconciling the transactions...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) ': BubbleEvent = False : Exit Sub
            '                    End If
            '                End If
            '            Next

            '            If ErrorFlag = True Then
            '                If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            '                Try
            '                    EditText5.Value = ""
            '                    objform.Freeze(True)
            '                    Matrix1.FlushToDataSource()
            '                    For rowNum As Integer = 0 To odbdsDetails.Size - 1
            '                        odbdsDetails.SetValue("U_JENo", rowNum, "")
            '                        odbdsDetails.SetValue("U_RecoNo", rowNum, "")
            '                    Next
            '                    Matrix1.LoadFromDataSource()
            '                Catch ex As Exception
            '                Finally
            '                    objform.Freeze(False)
            '                End Try
            '                objform.Refresh()
            '                objaddon.objapplication.StatusBar.SetText("Error while reconciling the transactions...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
            '            Else
            '                If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            '                objaddon.objapplication.StatusBar.SetText("Internal Reconciliations Created Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            '            End If
            '        Catch ex As Exception
            '        End Try
            '    End If
            'Catch ex As Exception
            'End Try


        End Sub

        Private Sub Form_DataAddBefore(ByRef pVal As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
            Try
                If objform.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then Exit Sub
                Dim Amt As Double
                Dim Line As Integer = 0
                Dim ErrorFlag As Boolean = False

                If BranchSplitReconciliation = "Y" Then ' with splitup JE
                    Try
                        Dim lineflag As Boolean = False
                        'If EditText5.Value <> "" Then
                        If Matrix0.VisualRowCount > 0 Then
                            For i As Integer = 1 To Matrix0.VisualRowCount
                                If Matrix0.Columns.Item("recono").Cells.Item(i).Specific.String <> "" Then
                                    lineflag = True
                                    Exit For
                                End If
                            Next
                        End If
                        If lineflag = True Then Exit Sub
                        'End If
                        If objPayIntRecoDT.Rows.Count > 0 Then objFDT = objPayIntRecoDT 'Else objaddon.objapplication.StatusBar.SetText("Rows required for reconciling the transactions...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
                        If objFDT.Rows.Count = 0 Then
                            objaddon.objapplication.MessageBox("Rows required for reconciling the transactions...", 0, "OK")
                            objaddon.objapplication.StatusBar.SetText("Rows required for reconciling the transactions...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                        Dim paytotsum As String = objFDT.Compute("SUM(paytot)", "").ToString
                        If paytotsum <> "0" Then
                            objaddon.objapplication.MessageBox("Reconciliation difference must be zero before reconciling...", 0, "OK")
                            objaddon.objapplication.StatusBar.SetText("Reconciliation difference must be zero before reconciling...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                        If (objaddon.objapplication.MessageBox("You cannot change this document after you have added it. Continue?", 2, "Yes", "No") <> 1) Then BubbleEvent = False : Return
                        Dim BranchDT = From dr In objFDT.AsEnumerable()
                                       Group dr By Ph = New With {Key .DTLine = dr.Field(Of String)("Row"), Key .branch = dr.Field(Of String)("branchc")} Into drg = Group
                                       Where drg.Sum(Function(dr) dr.Field(Of Double)("paytot")) = 0
                                       Select New With {                        'Where drg.Sum(Function(dr) dr.Field(Of String)("paytot")) = 0
                    .branch = Ph.branch,
                    .line = Ph.DTLine,
                    .LengthSum = drg.Sum(Function(dr) dr.Field(Of Double)("paytot"))
                    }
                        objaddon.objapplication.StatusBar.SetText("Creating transactions.Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        If objaddon.objcompany.InTransaction = False Then objaddon.objcompany.StartTransaction()

                        For Each RowID In BranchDT
                            Amt = Math.Round(CDbl(RowID.LengthSum), SumRound)
                            If CDbl(Amt) = 0 Then
                                If BranchReconciliation_Consolidated(objFDT, RowID.branch.ToString, RowID.line.ToString) = False Then
                                    ErrorFlag = True
                                    objaddon.objapplication.StatusBar.SetText("Error occurred while reconciling the transactions...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) ': BubbleEvent = False : Exit Sub
                                End If
                            End If
                        Next
                        Dim JEBranchDT = From dr In objFDT.AsEnumerable()
                                         Group dr By Ph = New With {Key .DTLine = dr.Field(Of String)("Row"), Key .branch = dr.Field(Of String)("branchc")} Into drg = Group
                                         Where drg.Sum(Function(dr) dr.Field(Of Double)("paytot")) <> 0
                                         Select New With {
                    .branch = Ph.branch,
                    .line = Ph.DTLine,
                    .LengthSum = drg.Sum(Function(dr) dr.Field(Of Double)("paytot"))
                    }
                        If JEBranchDT.Count > 0 Then
                            If JournalEntry_Consolidated(objFDT) = False Then
                                ErrorFlag = True
                                objaddon.objapplication.StatusBar.SetText("Error occurred while reconciling the transactions...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) ': BubbleEvent = False : Exit Sub
                            End If
                        End If

                        If ErrorFlag = True Then
                            If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            Try
                                EditText5.Value = ""
                                objform.Freeze(True)
                                Matrix0.FlushToDataSource()
                                For rowNum As Integer = 0 To odbdsDetails1.Size - 1
                                    odbdsDetails1.SetValue("U_RecoNo", rowNum, "")
                                Next
                                Matrix0.LoadFromDataSource()
                            Catch ex As Exception
                                BubbleEvent = False
                            Finally
                                objform.Freeze(False)
                            End Try
                            objform.Update()
                            objaddon.objapplication.MessageBox("Error while reconciling the transactions... " + clsModule.objaddon.objcompany.GetLastErrorDescription() + "-" + clsModule.objaddon.objcompany.GetLastErrorCode(), 0, "OK")
                            objaddon.objapplication.StatusBar.SetText("Error while reconciling the multi-branch transactions...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
                        Else
                            If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                            Matrix0.FlushToDataSource()
                            objaddon.objapplication.StatusBar.SetText("Multi-branch Internal Reconciliations Created Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            objPayIntRecoDT.Clear()
                        End If
                    Catch ex As Exception
                        BubbleEvent = False
                        objaddon.objapplication.MessageBox("Exception:  " + ex.Message, 0, "OK")
                        objaddon.objapplication.StatusBar.SetText("Exception:  " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End Try
                Else 'Consolidated JE
                    Try
                        If Not CDbl(EditText4.Value) = 0 Then
                            objaddon.objapplication.MessageBox("Reconciliation difference must be zero before reconciling...", 0, "OK")
                            objaddon.objapplication.StatusBar.SetText("Reconciliation difference must be zero before reconciling...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If

                        If objActualDT.Rows.Count > 0 Then objFDT = objActualDT Else objFDT = build_Matrix_DataTable("paytot", Matrix1)
                        If objFDT.Rows.Count = 0 Then
                            objaddon.objapplication.MessageBox("Rows required for reconciling the transactions...", 0, "OK")
                            objaddon.objapplication.StatusBar.SetText("Rows required for reconciling the transactions...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Sub
                        End If
                        If (objaddon.objapplication.MessageBox("You cannot change this document after you have added it. Continue?", 2, "Yes", "No") <> 1) Then BubbleEvent = False : Return
                        Dim otherBranchDT = From dr In objFDT.AsEnumerable()
                                            Group dr By Ph = dr.Field(Of String)("branchc") Into drg = Group
                                            Select New With {                   'Where drg.Sum(Function(dr) dr.Field(Of String)("paytot")) = 0  'Ph <> Branch And
                .branch = Ph,
                .LengthSum = drg.Sum(Function(dr) dr.Field(Of Double)("paytot"))
                }
                        '    Dim otherBranchDT = From dr In objFDT.AsEnumerable()
                        '                        Group dr By Ph = New With {Key .branch = dr.Field(Of String)("branchc"), Key .DTLine = dr.Field(Of String)("#")} Into drg = Group
                        '                        Select New With {                        'Where drg.Sum(Function(dr) dr.Field(Of String)("paytot")) = 0
                        '.branch = Ph.branch,
                        '.line = Ph.DTLine,
                        '.LengthSum = drg.Sum(Function(dr) dr.Field(Of String)("paytot"))
                        '}
                        objaddon.objapplication.StatusBar.SetText("Creating transactions.Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        If objaddon.objcompany.InTransaction = False Then objaddon.objcompany.StartTransaction()
                        For Each RowID In otherBranchDT
                            Amt = Math.Round(CDbl(RowID.LengthSum), SumRound)
                            If CDbl(Amt) = 0 Then
                                If BranchReconciliation(objFDT, RowID.branch.ToString()) = False Then
                                    ErrorFlag = True
                                    Exit For
                                End If
                            Else
                                If JournalEntry(objFDT, RowID.branch.ToString(), CDbl(Amt)) = False Then
                                    ErrorFlag = True
                                    objaddon.objapplication.StatusBar.SetText("Error occurred while reconciling the transactions...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) ': BubbleEvent = False : Exit Sub
                                    Exit For
                                End If
                            End If
                        Next

                        If ErrorFlag = True Then
                            If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            Try
                                EditText5.Value = ""
                                objform.Freeze(True)
                                Matrix1.FlushToDataSource()
                                For rowNum As Integer = 0 To odbdsDetails.Size - 1
                                    If odbdsDetails.GetValue("U_JENo", rowNum) <> "" Or odbdsDetails.GetValue("U_RecoNo", rowNum) <> "" Then
                                        odbdsDetails.SetValue("U_JENo", rowNum, "")
                                        odbdsDetails.SetValue("U_RecoNo", rowNum, "")
                                    End If

                                Next

                                For DTRow As Integer = 0 To objActualDT.Rows.Count - 1
                                    If Not objActualDT.Rows(DTRow)("jeno").ToString = String.Empty Or Not objActualDT.Rows(DTRow)("recono").ToString = String.Empty Then
                                        objActualDT.Rows(DTRow)("jeno") = ""
                                        objActualDT.Rows(DTRow)("recono") = ""
                                    End If
                                Next

                                Matrix1.LoadFromDataSource()
                            Catch ex As Exception
                                BubbleEvent = False
                            Finally
                                objform.Freeze(False)
                            End Try
                            objform.Update()
                            objaddon.objapplication.MessageBox("Error while reconciling the transactions... " + clsModule.objaddon.objcompany.GetLastErrorDescription() + "-" + clsModule.objaddon.objcompany.GetLastErrorCode(), 0, "OK")
                            objaddon.objapplication.StatusBar.SetText("Error while reconciling the transactions...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
                        Else
                            If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                            Matrix1.FlushToDataSource()
                            'For rowNum As Integer = 0 To odbdsDetails.Size - 1
                            '    strSQL = odbdsDetails.GetValue("U_JENo", rowNum)
                            '    strSQL = odbdsDetails.GetValue("U_RecoNo", rowNum)
                            'Next
                            objaddon.objapplication.StatusBar.SetText("Branch Internal Reconciliation Created Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        End If
                    Catch ex As Exception
                        BubbleEvent = False
                        objaddon.objapplication.MessageBox("Exception:  " + ex.Message, 0, "OK")
                        objaddon.objapplication.StatusBar.SetText("Exception:  " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End Try
                End If
            Catch ex As Exception
                BubbleEvent = False
                objaddon.objapplication.MessageBox("Form_DataAdd Exception:  " + ex.Message, 0, "OK")
                objaddon.objapplication.StatusBar.SetText("Form_DataAdd Exception:  " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        End Sub

        Private Sub Form_ResizeAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                'objform.Freeze(True)
                Matrix1.AutoResizeColumns()
                'objform.Freeze(False)
            Catch ex As Exception
                'objform.Freeze(False)
            End Try

        End Sub

        Private Sub Form_LoadAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                objform = objaddon.objapplication.Forms.GetForm("FOITR", pVal.FormTypeCount)
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Form_DataAddAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo)
            Try
                If pVal.ActionSuccess = False Then Exit Sub
                For DTRow As Integer = 0 To objActualDT.Rows.Count - 1
                    If Not objActualDT.Rows(DTRow)("jeno").ToString = String.Empty Then
                        objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        strSQL = "Update OJDT Set ""U_IntRecEntry""='" & objform.DataSources.DBDataSources.Item("@MI_OITR").GetValue("DocEntry", 0) & "' Where ""TransId""='" & objActualDT.Rows(DTRow)("jeno").ToString & "' "
                        objRs.DoQuery(strSQL)
                    End If
                Next
            Catch ex As Exception

            End Try
        End Sub


#End Region

#Region "Functions"

        Private Function LoadData(ByVal Query As String) As Boolean
            Try
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRs.DoQuery(Query)
                Matrix1.Clear()
                odbdsDetails.Clear()
                If objRs.RecordCount > 0 Then
                    objaddon.objapplication.StatusBar.SetText("Loading data Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    objform.Freeze(True)
                    While Not objRs.EoF
                        Matrix1.AddRow()
                        Matrix1.GetLineData(Matrix1.VisualRowCount)
                        odbdsDetails.SetValue("LineId", 0, objRs.Fields.Item("LineId").Value.ToString)
                        odbdsDetails.SetValue("U_Select", 0, objRs.Fields.Item("Selected").Value.ToString)
                        odbdsDetails.SetValue("U_TransId", 0, objRs.Fields.Item("TransId").Value.ToString)
                        odbdsDetails.SetValue("U_TLine", 0, objRs.Fields.Item("Line_ID").Value.ToString)
                        odbdsDetails.SetValue("U_DebCred", 0, objRs.Fields.Item("DebCred").Value.ToString)
                        odbdsDetails.SetValue("U_CardType", 0, objRs.Fields.Item("CardType").Value.ToString)
                        odbdsDetails.SetValue("U_Origin", 0, objRs.Fields.Item("Origin").Value.ToString)
                        odbdsDetails.SetValue("U_OriginNo", 0, objRs.Fields.Item("DocNum").Value.ToString)
                        odbdsDetails.SetValue("U_DocEntry", 0, objRs.Fields.Item("DocEntry").Value.ToString)
                        odbdsDetails.SetValue("U_CardCode", 0, objRs.Fields.Item("CardCode").Value.ToString)
                        odbdsDetails.SetValue("U_CardName", 0, objRs.Fields.Item("CardName").Value.ToString)
                        odbdsDetails.SetValue("U_DocDate", 0, objRs.Fields.Item("DocDate").Value)
                        odbdsDetails.SetValue("U_Total", 0, CDbl(objRs.Fields.Item("DocTotal").Value.ToString))
                        odbdsDetails.SetValue("U_BalDue", 0, CDbl(objRs.Fields.Item("Balance").Value.ToString))
                        odbdsDetails.SetValue("U_PayTotal", 0, objRs.Fields.Item("Balance").Value.ToString)
                        odbdsDetails.SetValue("U_Memo", 0, objRs.Fields.Item("LineMemo").Value.ToString)
                        odbdsDetails.SetValue("U_BranchId", 0, objRs.Fields.Item("BPLId").Value.ToString)
                        odbdsDetails.SetValue("U_BranchNam", 0, objRs.Fields.Item("BPLName").Value.ToString)
                        odbdsDetails.SetValue("U_Object", 0, objRs.Fields.Item("ObjType").Value.ToString)
                        odbdsDetails.SetValue("U_Pay", 0, objRs.Fields.Item("Balance").Value.ToString)
                        odbdsDetails.SetValue("U_Ref1", 0, objRs.Fields.Item("Ref1").Value.ToString)
                        odbdsDetails.SetValue("U_Ref2", 0, objRs.Fields.Item("Ref2").Value.ToString)
                        odbdsDetails.SetValue("U_Ref3", 0, objRs.Fields.Item("Ref3").Value.ToString)
                        'objform.DataSources.UserDataSources.Item("UD_0").Value = objRs.Fields.Item("Balance").Value.ToString
                        Matrix1.SetLineData(Matrix1.VisualRowCount)
                        objRs.MoveNext()
                    End While
                    Matrix1.AutoResizeColumns()
                    objaddon.objapplication.StatusBar.SetText("Data Loaded Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    objform.Freeze(False)
                    Return True
                Else
                    objaddon.objapplication.StatusBar.SetText("No records found...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Return False
                End If
            Catch ex As Exception
                objform.Freeze(False)
                Return False
            End Try
        End Function

        Private Function build_Matrix_DataTable(ByVal sKeyFieldID As String, ByVal MatrixID As SAPbouiCOM.Matrix) As DataTable
            Dim objcheckbox As SAPbouiCOM.CheckBox
            Try
                Dim oDT As New DataTable
                'Add all of the columns by unique ID to the DataTable
                For iCol As Integer = 0 To MatrixID.Columns.Count - 1
                    'Skip invisible columns
                    'If oMatrix.Columns.Item(iCol).Visible = False Then Continue For
                    If iCol <> 1 Then
                        oDT.Columns.Add(MatrixID.Columns.Item(iCol).UniqueID)
                    End If
                Next
                'Now, add all of the data into the DataTable
                For iRow As Integer = 1 To MatrixID.VisualRowCount
                    objcheckbox = MatrixID.Columns.Item("select").Cells.Item(iRow).Specific
                    If objcheckbox.Checked = True Then
                        Dim oRow As DataRow = oDT.NewRow
                        For iCol As Integer = 0 To MatrixID.Columns.Count - 1
                            'If oMatrix.Columns.Item(iCol).Visible = False Then Continue For
                            If iCol <> 1 Then
                                oRow.Item(MatrixID.Columns.Item(iCol).UniqueID) = MatrixID.Columns.Item(iCol).Cells.Item(iRow).Specific.Value
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
                objcheckbox = Matrix1.Columns.Item("select").Cells.Item(Row).Specific
                Dim oRow As DataRow = objActualDT.NewRow
                If objcheckbox.Checked = True Then
                    If objActualDT.Rows.Count > 0 Then
                        For DTRow As Integer = 0 To objActualDT.Rows.Count - 1
                            If objActualDT.Rows(DTRow)("#").ToString = Matrix1.Columns.Item("#").Cells.Item(Row).Specific.Value Then
                                If ColName <> "" Then
                                    objActualDT.Rows(DTRow)(Matrix1.Columns.Item(ColName).UniqueID) = Matrix1.Columns.Item(ColName).Cells.Item(Row).Specific.Value
                                    DataFlag = True
                                    Exit For
                                End If
                            End If
                        Next
                        If DataFlag = False Then
                            For iCol As Integer = 0 To Matrix1.Columns.Count - 1
                                If iCol <> 1 And iCol <> 18 Then
                                    oRow.Item(Matrix1.Columns.Item(iCol).UniqueID) = Matrix1.Columns.Item(iCol).Cells.Item(Row).Specific.Value
                                End If
                            Next
                            If objActualDT.Rows.Count = 0 Then oRow.Item("Row") = 0 Else oRow.Item("Row") = objActualDT.Rows.Count - 1
                            objActualDT.Rows.Add(oRow)
                        End If
                    Else
                        For iCol As Integer = 0 To Matrix1.Columns.Count - 1
                            If iCol <> 1 And iCol <> 18 Then
                                oRow.Item(Matrix1.Columns.Item(iCol).UniqueID) = Matrix1.Columns.Item(iCol).Cells.Item(Row).Specific.Value
                            End If
                        Next
                        If objActualDT.Rows.Count = 0 Then oRow.Item("Row") = 0 Else oRow.Item("Row") = objActualDT.Rows.Count - 1
                        objActualDT.Rows.Add(oRow)
                    End If
                Else
                    For DTRow As Integer = 0 To objActualDT.Rows.Count - 1
                        If objActualDT.Rows(DTRow)("#").ToString = Matrix1.Columns.Item("#").Cells.Item(Row).Specific.Value Then
                            objActualDT.Rows(DTRow).Delete()
                            Exit For
                        End If
                    Next
                    If Matrix1.Columns.Item("jeadj").Cells.Item(Row).Specific.String = "Y" Then Matrix1.Columns.Item("jeadj").Cells.Item(Row).Specific.String = ""
                    If Matrix1.Columns.Item("tranadj").Cells.Item(Row).Specific.Checked = True Then Matrix1.Columns.Item("tranadj").Cells.Item(Row).Specific.Checked = False
                End If
                Return objActualDT
            Catch ex As Exception
                Return Nothing
            End Try
        End Function

        Private Sub Calculate_Total()
            Try
                Dim objDT As New DataTable
                Dim value As Double
                objDT.Columns.Add("paytot", GetType(Double))

                Dim objcheckbox As SAPbouiCOM.CheckBox
                For iRow As Integer = 1 To Matrix1.VisualRowCount
                    objcheckbox = Matrix1.Columns.Item("select").Cells.Item(iRow).Specific
                    If objcheckbox.Checked = True Then
                        Dim oRow As DataRow = objDT.NewRow
                        oRow.Item(Matrix1.Columns.Item("paytot").UniqueID) = Matrix1.Columns.Item("paytot").Cells.Item(iRow).Specific.Value
                        objDT.Rows.Add(oRow)
                    End If
                Next
                For i As Integer = 0 To objDT.Rows.Count - 1
                    If objDT.Rows(i)("paytot").ToString <> "" Then
                        value += CDbl(objDT.Rows(i)("paytot").ToString)
                    End If
                Next
                odbdsHeader.SetValue("U_Total", 0, value) 'Total
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Calc_Total_14092021(ByVal Row As Integer)
            Try
                Dim value As Double
                Dim DataFlag As Boolean
                Dim objcheckbox As SAPbouiCOM.CheckBox
                Dim oRow As DataRow = oSelectedDT.NewRow
                objcheckbox = Matrix1.Columns.Item("select").Cells.Item(Row).Specific

                If objcheckbox.Checked = True Then
                    If oSelectedDT.Rows.Count > 0 Then
                        For DTRow As Integer = 0 To oSelectedDT.Rows.Count - 1
                            If oSelectedDT.Rows(DTRow)("#").ToString = Matrix1.Columns.Item("#").Cells.Item(Row).Specific.Value Then
                                oSelectedDT.Rows(DTRow)("paytot") = Matrix1.Columns.Item("paytot").Cells.Item(Row).Specific.Value
                                DataFlag = True
                                Exit For
                            End If
                        Next
                        If DataFlag = False Then
                            oRow.Item(Matrix1.Columns.Item("paytot").UniqueID) = Matrix1.Columns.Item("paytot").Cells.Item(Row).Specific.Value
                            oRow.Item(Matrix1.Columns.Item("#").UniqueID) = Matrix1.Columns.Item("#").Cells.Item(Row).Specific.Value
                            oSelectedDT.Rows.Add(oRow)
                        End If
                    Else
                        oRow.Item(Matrix1.Columns.Item("paytot").UniqueID) = Matrix1.Columns.Item("paytot").Cells.Item(Row).Specific.Value
                        oRow.Item(Matrix1.Columns.Item("#").UniqueID) = Matrix1.Columns.Item("#").Cells.Item(Row).Specific.Value
                        oSelectedDT.Rows.Add(oRow)
                    End If
                Else
                    For DTRow As Integer = 0 To oSelectedDT.Rows.Count - 1
                        If oSelectedDT.Rows(DTRow)("#").ToString = Matrix1.Columns.Item("#").Cells.Item(Row).Specific.Value Then
                            oSelectedDT.Rows(DTRow).Delete()
                            Exit For
                        End If
                    Next
                End If

                For i As Integer = 0 To oSelectedDT.Rows.Count - 1
                    If Val(oSelectedDT.Rows(i)("paytot").ToString) <> 0 Then
                        value = Math.Round(value + CDbl(oSelectedDT.Rows(i)("paytot")), SumRound)
                    End If
                Next
                odbdsHeader.SetValue("U_Total", 0, value) 'Total
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Field_Setup()
            Try
                Matrix1.Columns.Item("branchc").Visible = False
                Matrix1.Columns.Item("pay").Visible = False
                Matrix1.Columns.Item("tranline").Visible = False
                Matrix1.Columns.Item("debcred").Visible = False
                Matrix1.Columns.Item("cardtype").Visible = False
                Matrix1.Columns.Item("object").Visible = False
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "tdocdate", False, True, False)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "cmbstat", False, True, False)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "tremark", True, True, False)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "Series", True, True, False)
            Catch ex As Exception

            End Try
        End Sub

        Private Function JournalEntry(ByVal InDT As DataTable, ByVal Branch As String, ByVal JEAmount As Double) As Boolean
            Try
                Dim TransId As String = "", GLCode As String, Series As String, CardCode As String = "", MatLine As String = "", CardType As String = ""
                Dim objrecset As SAPbobsCOM.Recordset
                Dim objjournalentry As SAPbobsCOM.JournalEntries
                Dim Amount As Double
                Dim DTLine As Integer = 0
                Try
                    For DTRow As Integer = 0 To InDT.Rows.Count - 1
                        If InDT.Rows(DTRow)("branchc").ToString = Branch Then
                            TransId = InDT.Rows(DTRow)("jeno").ToString
                            MatLine = CInt(InDT.Rows(DTRow)("#").ToString)
                            DTLine = DTRow
                            CardCode = InDT.Rows(DTRow)("cardc").ToString()
                            CardType = InDT.Rows(DTRow)("cardtype").ToString()
                            Exit For
                        End If
                    Next
                    If TransId = "" Then
                        Try
                            objjournalentry = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
                            objaddon.objapplication.StatusBar.SetText("Journal Entry Creating Please wait..." & GetBranchName(Branch), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            'If Not objaddon.objcompany.InTransaction Then objaddon.objcompany.StartTransaction()
                            Dim oEdit As SAPbouiCOM.EditText
                            oEdit = objform.Items.Item("tdocdate").Specific
                            Dim DocDate As Date = Date.ParseExact(oEdit.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                            objjournalentry.ReferenceDate = DocDate  'Now.Date.ToString("yyyyMMdd") 
                            'objjournalentry.DueDate = Now.Date.ToString("yyyyMMdd") 'DocDate
                            objjournalentry.TaxDate = DocDate  ' ConvertDate.ToString("dd/MM/yy") 
                            objjournalentry.Reference = "Int Rec Payment JE"
                            objjournalentry.Memo = "Posted thro' recon On: " & Now.ToString
                            objjournalentry.UserFields.Fields.Item("U_IntRecNo").Value = EditText1.Value
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
                            Amount = IIf(JEAmount < 0, -JEAmount, JEAmount)

                            If JEAmount < 0 Then objjournalentry.Lines.Credit = Amount Else objjournalentry.Lines.Debit = Amount
                            GLCode = objaddon.objglobalmethods.getSingleValue("select ""PmtClrAct"" ""ControlAccount"" from OBPL where ""BPLId""='" & Branch & "'")
                            objjournalentry.Lines.AccountCode = GLCode
                            objjournalentry.Lines.BPLID = Branch
                            objjournalentry.Lines.Add()

                            objjournalentry.Lines.ShortName = CardCode
                            If JEAmount < 0 Then objjournalentry.Lines.Debit = Amount Else objjournalentry.Lines.Credit = Amount
                            'If CardType = "C" Then
                            '    If JEAmount < 0 Then objjournalentry.Lines.Debit = Amount Else objjournalentry.Lines.Credit = Amount
                            'Else
                            '    If JEAmount < 0 Then objjournalentry.Lines.Credit = Amount Else objjournalentry.Lines.Debit = Amount
                            'End If
                            objjournalentry.Lines.BPLID = Branch
                            objjournalentry.Lines.Add()

                            'For Row As Integer = 0 To InDT.Rows.Count - 1
                            '    If InDT.Rows(Row)("branchc").ToString = Branch And InDT.Rows(Row)("jeno").ToString = String.Empty Then
                            '        JEAmount = IIf(CDbl(InDT.Rows(Row)("paytot").ToString) < 0, -CDbl(InDT.Rows(Row)("paytot").ToString), CDbl(InDT.Rows(Row)("paytot").ToString))
                            '        GLCode = objaddon.objglobalmethods.getSingleValue("select ""PmtClrAct"" ""ControlAccount"" from OBPL where ""BPLId""='" & Branch & "'")
                            '        objjournalentry.Lines.AccountCode = GLCode
                            '        If CDbl(InDT.Rows(Row)("paytot").ToString) < 0 And InDT.Rows(Row)("cardtype").ToString = "C" Then objjournalentry.Lines.Credit = JEAmount Else objjournalentry.Lines.Debit = JEAmount
                            '        objjournalentry.Lines.BPLID = Branch
                            '        objjournalentry.Lines.Add()
                            '        If CardCode = "" Then CardCode = InDT.Rows(Row)("cardc").ToString()
                            '        If MatLine = "" Then MatLine = CInt(InDT.Rows(Row)("#").ToString())
                            '        objjournalentry.Lines.ShortName = InDT.Rows(Row)("cardc").ToString
                            '        If CDbl(InDT.Rows(Row)("paytot").ToString) < 0 And InDT.Rows(Row)("cardtype").ToString = "C" Then objjournalentry.Lines.Debit = JEAmount Else objjournalentry.Lines.Credit = JEAmount
                            '        objjournalentry.Lines.BPLID = Branch
                            '        objjournalentry.Lines.Add()
                            '    End If
                            'Next
                            If objjournalentry.Add <> 0 Then
                                'If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                objaddon.objapplication.SetStatusBarMessage("Journal: " & GetBranchName(Branch) & "-" & objaddon.objcompany.GetLastErrorDescription & "-" & objaddon.objcompany.GetLastErrorCode, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                Return False
                            Else
                                'If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                TransId = objaddon.objcompany.GetNewObjectKey()
                                Matrix1.Columns.Item("jeno").Cells.Item(CInt(MatLine)).Specific.String = TransId
                                objActualDT.Rows(DTLine)("jeno") = TransId
                                If Matrix1.Columns.Item("jeno").Cells.Item(CInt(MatLine)).Specific.String <> "" And Matrix1.Columns.Item("recono").Cells.Item(CInt(MatLine)).Specific.String = "" Then
                                    If MultiBranch_InternalReconciliation(InDT, MatLine, DTLine, TransId, CardCode, Branch) = False Then
                                        objaddon.objapplication.StatusBar.SetText("MultiBranch_InternalReco" & GetBranchName(Branch), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        Return False
                                    End If
                                End If
                                objaddon.objapplication.SetStatusBarMessage("Journal Entry Created Successfully..." & GetBranchName(Branch), SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                            End If
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(objjournalentry)
                        Catch ex As Exception
                            'If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            objaddon.objapplication.SetStatusBarMessage("JE Posting Error" & GetBranchName(Branch) & "-" & objaddon.objcompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                        End Try
                    Else
                        Try
                            If Not InDT.Rows(DTLine)("jeno").ToString = String.Empty And (InDT.Rows(DTLine)("recono").ToString = String.Empty Or InDT.Rows(DTLine)("recono").ToString = "0") Then
                                If MultiBranch_InternalReconciliation(InDT, MatLine, DTLine, TransId, InDT.Rows(DTLine)("cardc").ToString, Branch) = True Then
                                    objaddon.objapplication.StatusBar.SetText("MultiBranch_InternalReco" & GetBranchName(Branch), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    Return False
                                End If
                            End If
                        Catch ex As Exception
                            objaddon.objapplication.SetStatusBarMessage("JE Rec " & GetBranchName(Branch) & "-" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                        End Try
                    End If
                    objrecset = Nothing
                    Return True
                Catch ex As Exception
                    'If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                    Return False
                End Try
            Catch ex As Exception
                'If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                objaddon.objapplication.SetStatusBarMessage("JE " & GetBranchName(Branch) & "-" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                Return False
            Finally
                'GC.Collect()
                'GC.WaitForPendingFinalizers()
            End Try
        End Function

        Private Function MultiBranch_InternalReconciliation(ByVal InDT As DataTable, ByVal MatLine As Integer, ByVal DTLine As Integer, ByVal transid As String, ByVal BPCode As String, ByVal Branch As String) As Boolean
            Try
                Dim service As IInternalReconciliationsService = objaddon.objcompany.GetCompanyService().GetBusinessService(ServiceTypes.InternalReconciliationsService)
                Dim openTrans As InternalReconciliationOpenTrans = service.GetDataInterface(InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationOpenTrans)
                Dim reconParams As IInternalReconciliationParams = service.GetDataInterface(InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationParams)
                openTrans.CardOrAccount = CardOrAccountEnum.coaCard
                'openTrans.ReconDate = DocumentDate

                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim RecAmount As Double
                Dim Row As Integer = 0
                strSQL = "select CASE WHEN T1.""BalDueCred""<>0  THEN  T1.""BalDueCred"" ELSE T1.""BalDueDeb"" END AS ""Balance"",T1.""Line_ID"" from OJDT T0 join JDT1 T1 ON T0.""TransId""=T1.""TransId"" where  T1.""TransId""='" & transid & "' and T1.""ShortName""='" & BPCode & "'"
                objRs.DoQuery(strSQL)
                If objRs.RecordCount > 0 Then
                    Dim DocDate As Date = Date.ParseExact(EditText0.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                    openTrans.ReconDate = DocDate
                    For Rec As Integer = 0 To objRs.RecordCount - 1
                        If Val(objRs.Fields.Item(0).Value.ToString) <> 0 Then
                            RecAmount = Math.Round(CDbl(objRs.Fields.Item(0).Value.ToString), SumRound)
                            openTrans.InternalReconciliationOpenTransRows.Add()
                            openTrans.InternalReconciliationOpenTransRows.Item(Row).Selected = BoYesNoEnum.tYES
                            openTrans.InternalReconciliationOpenTransRows.Item(Row).TransId = transid
                            openTrans.InternalReconciliationOpenTransRows.Item(Row).TransRowId = CInt(objRs.Fields.Item(1).Value.ToString) 'InDT.Rows(Row)("tranline").ToString
                            openTrans.InternalReconciliationOpenTransRows.Item(Row).ReconcileAmount = RecAmount 'CDbl(objRs.Fields.Item(0).Value.ToString)
                            Row += 1
                        End If
                        objRs.MoveNext()
                    Next

                End If
                For DTRow As Integer = 0 To InDT.Rows.Count - 1
                    If InDT.Rows(DTRow)("branchc").ToString = Branch Then
                        If InDT.Rows(DTRow)("debcred").ToString = "C" Then RecAmount = Math.Round(-CDbl(InDT.Rows(DTRow)("paytot").ToString), SumRound) Else RecAmount = Math.Round(CDbl(InDT.Rows(DTRow)("paytot").ToString), SumRound)
                        'RecAmount = CDbl(InDT.Rows(DTRow)("paytot").ToString)
                        openTrans.InternalReconciliationOpenTransRows.Add()
                        openTrans.InternalReconciliationOpenTransRows.Item(Row).Selected = BoYesNoEnum.tYES
                        openTrans.InternalReconciliationOpenTransRows.Item(Row).TransId = InDT.Rows(DTRow)("trannum").ToString
                        openTrans.InternalReconciliationOpenTransRows.Item(Row).TransRowId = InDT.Rows(DTRow)("tranline").ToString
                        openTrans.InternalReconciliationOpenTransRows.Item(Row).ReconcileAmount = RecAmount ' CDbl(InDT.Rows(DTRow)("paytot").ToString) '
                        Row += 1
                    End If
                Next
                Dim Reconum As Integer = 0
                Try
                    reconParams = service.Add(openTrans)
                Catch ex As Exception
                    If Reconum = 0 Then objaddon.objapplication.StatusBar.SetText("Reconciled Error : " & GetBranchName(Branch) & "-" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : Return False
                End Try

                Reconum = reconParams.ReconNum
                Matrix1.Columns.Item("recono").Cells.Item(MatLine).Specific.String = Reconum
                objActualDT.Rows(DTLine)("recono") = Reconum
                objRs.DoQuery("Update OJDT Set ""U_ReconNum""='" & Reconum & "' where ""TransId""='" & transid & "'")
                objaddon.objapplication.StatusBar.SetText("Reconciled successfully..." & GetBranchName(Branch), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(openTrans)
                'GC.Collect()
                Return True
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("Recon: " & GetBranchName(Branch) & "-" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End Try

        End Function

        Private Function BranchReconciliation(ByVal InDT As DataTable, ByVal Branch As String) As Boolean
            Try
                Dim service As IInternalReconciliationsService = objaddon.objcompany.GetCompanyService().GetBusinessService(ServiceTypes.InternalReconciliationsService)
                Dim openTrans As InternalReconciliationOpenTrans = service.GetDataInterface(InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationOpenTrans)
                Dim reconParams As IInternalReconciliationParams = service.GetDataInterface(InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationParams)
                openTrans.CardOrAccount = CardOrAccountEnum.coaCard
                Dim RecAmount As Double
                Dim Row As Integer = 0
                Dim Line As String = ""
                Dim DTLine As Integer = 0
                objRs = objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset)

                For DTRow As Integer = 0 To InDT.Rows.Count - 1
                    If InDT.Rows(DTRow)("branchc").ToString = Branch And Not InDT.Rows(DTRow)("recono").ToString = String.Empty Then
                        Line = CInt(InDT.Rows(DTRow)("#").ToString)
                        DTLine = DTRow
                        Exit For
                    End If
                Next
                If Line = "" Then
                    Dim DocDate As Date = Date.ParseExact(EditText0.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                    openTrans.ReconDate = DocDate
                    For DTRow As Integer = 0 To InDT.Rows.Count - 1
                        If InDT.Rows(DTRow)("branchc").ToString = Branch Then
                            If Line = "" Then Line = InDT.Rows(DTRow)("#").ToString
                            If InDT.Rows(DTRow)("debcred").ToString = "C" Then RecAmount = Math.Round(-CDbl(InDT.Rows(DTRow)("paytot").ToString), SumRound) Else RecAmount = Math.Round(CDbl(InDT.Rows(DTRow)("paytot").ToString), SumRound)
                            openTrans.InternalReconciliationOpenTransRows.Add()
                            openTrans.InternalReconciliationOpenTransRows.Item(Row).Selected = BoYesNoEnum.tYES
                            openTrans.InternalReconciliationOpenTransRows.Item(Row).TransId = InDT.Rows(DTRow)("trannum").ToString
                            openTrans.InternalReconciliationOpenTransRows.Item(Row).TransRowId = InDT.Rows(DTRow)("tranline").ToString
                            openTrans.InternalReconciliationOpenTransRows.Item(Row).ReconcileAmount = RecAmount
                            Row += 1
                        End If
                    Next
                    Dim Reconum As Integer = 0
                    Try
                        reconParams = service.Add(openTrans)
                    Catch ex As Exception
                        objaddon.objapplication.StatusBar.SetText("Reconciled Error..." & GetBranchName(Branch) & "-" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : Return False
                    End Try
                    Reconum = reconParams.ReconNum
                    Matrix1.Columns.Item("recono").Cells.Item(CInt(Line)).Specific.String = Reconum
                    objActualDT.Rows(DTLine)("recono") = Reconum
                    objaddon.objapplication.StatusBar.SetText("Reconciled successfully..." & GetBranchName(Branch), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(openTrans)
                    GC.Collect()
                End If

                Return True
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("Recon: " & GetBranchName(Branch) & "-" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End Try

        End Function

        Private Function GetBranchName(ByVal BCode As String) As String
            Try
                Dim BName As String
                BName = objaddon.objglobalmethods.getSingleValue("Select ""BPLName"" from OBPL where ""BPLId""='" & BCode & "' ")

                Return BName
            Catch ex As Exception
                Return 0
            End Try
        End Function

        Private Function JournalEntry_Consolidated(ByVal InDT As DataTable) As Boolean
            Try
                Dim TransId As String = "", GLCode As String, Series As String ', CardCode As String = "", MatLine As String = "", CardType As String = ""
                Dim objrecset As SAPbobsCOM.Recordset
                Dim objjournalentry As SAPbobsCOM.JournalEntries
                Dim Amount, DTAmt As Double
                Dim DTLine As Integer = 0
                Try
                    '        Dim otherBranDT = From dr In InDT.AsEnumerable()
                    '                          Group dr By Ph = dr.Field(Of String)("branchc") Into drg = Group
                    '                          Select New With {                   'Where drg.Sum(Function(dr) dr.Field(Of String)("paytot")) = 0  'Ph <> Branch And
                    '.branch = Ph,
                    '.LengthSum = drg.Sum(Function(dr) dr.Field(Of Double)("paytot"))
                    '}

                    Dim otherBranchDT = From dr In InDT.AsEnumerable()
                                        Group dr By Ph = New With {Key .DTLine = dr.Field(Of String)("Row"), Key .branch = dr.Field(Of String)("branchc")} Into drg = Group
                                        Where drg.Sum(Function(dr) dr.Field(Of Double)("paytot")) <> 0
                                        Select New With {                        'Where drg.Sum(Function(dr) dr.Field(Of String)("paytot")) = 0
                .branch = Ph.branch,
                .line = Ph.DTLine,
                .LengthSum = drg.Sum(Function(dr) dr.Field(Of Double)("paytot"))
                }

                    If otherBranchDT.Count = 0 Then Exit Function
                    objjournalentry = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
                    objaddon.objapplication.StatusBar.SetText("Journal Entry Creating Please wait...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    'If Not objaddon.objcompany.InTransaction Then objaddon.objcompany.StartTransaction()
                    Dim oEdit As SAPbouiCOM.EditText
                    oEdit = objform.Items.Item("tdocdate").Specific
                    Dim DocDate As Date = Date.ParseExact(oEdit.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                    objjournalentry.ReferenceDate = DocDate  'Now.Date.ToString("yyyyMMdd") 
                    'objjournalentry.DueDate = Now.Date.ToString("yyyyMMdd") 'DocDate
                    objjournalentry.TaxDate = DocDate  ' ConvertDate.ToString("dd/MM/yy") 
                    objjournalentry.Reference = "Int Rec Payment JE"
                    objjournalentry.Reference2 = EditText1.Value
                    objjournalentry.Reference3 = Now.ToString
                    objjournalentry.Memo = "Posted thro' recon add-on" '& Now.ToString
                    objjournalentry.UserFields.Fields.Item("U_IntRecNo").Value = EditText1.Value
                    If Localization = "IN" Then
                        If objaddon.HANA Then
                            Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 ""Series"" FROM NNM1 WHERE ""ObjectCode""='30' and ""Indicator""=(Select ""Indicator"" from OFPR where '" & DocDate.ToString("yyyyMMdd") & "' Between ""F_RefDate"" and ""T_RefDate"") " &
                                                                                                  " and Ifnull(""Locked"",'')='N' and ""BPLId""=(select ""BPLId"" from OBPL where Ifnull(""MainBPL"",'')='Y')")
                        Else
                            Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 Series FROM NNM1 WHERE ObjectCode='30' and Indicator=(Select Indicator from OFPR where '" & DocDate.ToString("yyyyMMdd") & "' Between F_RefDate and T_RefDate) " &
                                                                                                  " and Isnull(Locked,'')='N' and BPLId=(select BPLId from OBPL where Isnull(MainBPL,'')='Y')")
                        End If
                    Else
                        objjournalentry.AutoVAT = BoYesNoEnum.tNO
                        objjournalentry.AutomaticWT = BoYesNoEnum.tNO
                        If objaddon.HANA Then
                            Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 ""Series"" FROM NNM1 WHERE ""ObjectCode""='30' and ""Indicator""=(Select ""Indicator"" from OFPR where '" & DocDate.ToString("yyyyMMdd") & "' Between ""F_RefDate"" and ""T_RefDate"") " &
                                                                                                  " and Ifnull(""Locked"",'')='N' and ""BPLId""=(select ""BPLId"" from OBPL where Ifnull(""MainBPL"",'')='Y')")
                        Else
                            Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 Series FROM NNM1 WHERE ObjectCode='30' and Indicator=(Select Indicator from OFPR where '" & DocDate.ToString("yyyyMMdd") & "' Between F_RefDate and T_RefDate) " &
                                                                                                  " and Isnull(Locked,'')='N' and BPLId=(select BPLId from OBPL where Isnull(MainBPL,'')='Y')")
                        End If
                    End If
                    If Series <> "" Then objjournalentry.Series = Series

                    For DTRow As Integer = 0 To InDT.Rows.Count - 1
                        If InDT.Rows(DTRow)("Row").ToString <> "" Then
                            Amount = IIf(CDbl(InDT.Rows(DTRow)("paytot").ToString) < 0, -CDbl(InDT.Rows(DTRow)("paytot").ToString), CDbl(InDT.Rows(DTRow)("paytot").ToString))
                            objjournalentry.Lines.ShortName = InDT.Rows(DTRow)("cardc").ToString
                            If CDbl(InDT.Rows(DTRow)("paytot").ToString) < 0 Then objjournalentry.Lines.Debit = Amount Else objjournalentry.Lines.Credit = Amount
                            objjournalentry.Lines.BPLID = InDT.Rows(DTRow)("branchc").ToString ' Branch
                            objjournalentry.Lines.Add()
                            If InDT.Rows(DTRow)("tbranchc").ToString <> "" And InDT.Rows(DTRow)("branchc").ToString <> InDT.Rows(DTRow)("tbranchc").ToString Then
                                Amount = IIf(CDbl(InDT.Rows(DTRow)("paytot").ToString) < 0, -CDbl(InDT.Rows(DTRow)("paytot").ToString), CDbl(InDT.Rows(DTRow)("paytot").ToString))
                                If CDbl(InDT.Rows(DTRow)("paytot").ToString) < 0 Then objjournalentry.Lines.Credit = Amount Else objjournalentry.Lines.Debit = Amount
                                GLCode = objaddon.objglobalmethods.getSingleValue("select ""PmtClrAct"" ""ControlAccount"" from OBPL where ""BPLId""='" & InDT.Rows(DTRow)("branchc").ToString & "'")
                                objjournalentry.Lines.AccountCode = GLCode
                                objjournalentry.Lines.BPLID = InDT.Rows(DTRow)("branchc").ToString ' Branch
                                If InDT.Rows(DTRow)("tbranchn").ToString <> "" Then objjournalentry.Lines.UserFields.Fields.Item("U_TBranch").Value = InDT.Rows(DTRow)("tbranchn").ToString
                                objjournalentry.Lines.Add()
                                Amount = IIf(CDbl(InDT.Rows(DTRow)("paytot").ToString) < 0, -CDbl(InDT.Rows(DTRow)("paytot").ToString), CDbl(InDT.Rows(DTRow)("paytot").ToString))
                                If CDbl(InDT.Rows(DTRow)("paytot").ToString) < 0 Then objjournalentry.Lines.Debit = Amount Else objjournalentry.Lines.Credit = Amount
                                objjournalentry.Lines.AccountCode = GLCode
                                objjournalentry.Lines.BPLID = InDT.Rows(DTRow)("tbranchc").ToString ' Branch
                                If InDT.Rows(DTRow)("tbranchn").ToString <> "" Then objjournalentry.Lines.UserFields.Fields.Item("U_TBranch").Value = InDT.Rows(DTRow)("branchn").ToString
                                objjournalentry.Lines.Add()
                            End If

                        End If
                    Next

                    If objjournalentry.Add <> 0 Then
                        'If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        objaddon.objapplication.SetStatusBarMessage("Journal: " & "-" & objaddon.objcompany.GetLastErrorDescription & "-" & objaddon.objcompany.GetLastErrorCode, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                        Return False
                    Else
                        'If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        TransId = objaddon.objcompany.GetNewObjectKey()
                        EditText5.Value = TransId
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objjournalentry)
                        objaddon.objapplication.SetStatusBarMessage("Journal Entry Created Successfully...", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                        If MultiBranch_InternalReconciliation_Consolidated_JE(InDT, TransId) = False Then
                            objaddon.objapplication.SetStatusBarMessage("Journal Reconciliation: " & "-" & objaddon.objcompany.GetLastErrorDescription & "-" & objaddon.objcompany.GetLastErrorCode, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                            Return False
                        End If
                        'For IRow As Integer = 0 To objActualDT.Rows.Count - 1
                        '    Matrix1.Columns.Item("jeno").Cells.Item(CInt(InDT.Rows(IRow)("#").ToString)).Specific.String = TransId
                        'Next
                    End If

                Catch ex As Exception
                    'If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    objaddon.objapplication.SetStatusBarMessage("JE Posting Error" & "-" & objaddon.objcompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                End Try
                objrecset = Nothing
                Return True

            Catch ex As Exception
                'If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                objaddon.objapplication.SetStatusBarMessage("JE " & "-" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                Return False
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End Function

        Private Function MultiBranch_InternalReconciliation_Consolidated_JE(ByVal InDT As DataTable, ByVal transid As String) As Boolean
            Try



                Dim BranchDT = From dr In InDT.AsEnumerable()
                               Group dr By Ph = New With {Key .branch = dr.Field(Of String)("branchc")} Into drg = Group
                               Select New With {                        'Where drg.Sum(Function(dr) dr.Field(Of String)("paytot")) = 0
                .branch = Ph.branch,
                .LengthSum = drg.Sum(Function(dr) dr.Field(Of Double)("paytot"))
                }

                For Each RecRow In BranchDT

                    Dim service As IInternalReconciliationsService = objaddon.objcompany.GetCompanyService().GetBusinessService(ServiceTypes.InternalReconciliationsService)
                    Dim openTrans As InternalReconciliationOpenTrans = service.GetDataInterface(InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationOpenTrans)
                    Dim reconParams As IInternalReconciliationParams = service.GetDataInterface(InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationParams)
                    openTrans.CardOrAccount = CardOrAccountEnum.coaCard
                    'openTrans.ReconDate = DocumentDate
                    objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim RecAmount As Double
                    Dim Row As Integer = 0
                    objRs.DoQuery("select CASE WHEN T1.""BalDueCred""<>0  THEN  T1.""BalDueCred"" ELSE T1.""BalDueDeb"" END AS ""Balance"",T1.""Line_ID"" from OJDT T0 join JDT1 T1 ON T0.""TransId""=T1.""TransId"" where  T1.""TransId""='" & transid & "' and T1.""ShortName"" in (Select ""CardCode"" from OCRD) and T1.""BPLId""='" & RecRow.branch.ToString & "'")
                    If objRs.RecordCount > 0 Then
                        Dim DocDate As Date = Date.ParseExact(EditText0.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                        openTrans.ReconDate = DocDate
                        For Rec As Integer = 0 To objRs.RecordCount - 1
                            If Val(objRs.Fields.Item(0).Value.ToString) <> 0 Then
                                RecAmount = Math.Round(CDbl(objRs.Fields.Item(0).Value.ToString), SumRound)
                                openTrans.InternalReconciliationOpenTransRows.Add()
                                openTrans.InternalReconciliationOpenTransRows.Item(Row).Selected = BoYesNoEnum.tYES
                                openTrans.InternalReconciliationOpenTransRows.Item(Row).TransId = transid
                                openTrans.InternalReconciliationOpenTransRows.Item(Row).TransRowId = CInt(objRs.Fields.Item(1).Value.ToString) 'InDT.Rows(Row)("tranline").ToString
                                openTrans.InternalReconciliationOpenTransRows.Item(Row).ReconcileAmount = RecAmount 'CDbl(objRs.Fields.Item(0).Value.ToString)
                                Row += 1
                            End If
                            objRs.MoveNext()
                        Next
                    End If

                    For DTRow As Integer = 0 To InDT.Rows.Count - 1
                        If InDT.Rows(DTRow)("branchc").ToString = RecRow.branch.ToString Then
                            If InDT.Rows(DTRow)("debcred").ToString = "C" Then RecAmount = Math.Round(-CDbl(InDT.Rows(DTRow)("paytot").ToString), SumRound) Else RecAmount = Math.Round(CDbl(InDT.Rows(DTRow)("paytot").ToString), SumRound)
                            'RecAmount = CDbl(InDT.Rows(DTRow)("paytot").ToString)
                            openTrans.InternalReconciliationOpenTransRows.Add()
                            openTrans.InternalReconciliationOpenTransRows.Item(Row).Selected = BoYesNoEnum.tYES
                            openTrans.InternalReconciliationOpenTransRows.Item(Row).TransId = InDT.Rows(DTRow)("trannum").ToString
                            openTrans.InternalReconciliationOpenTransRows.Item(Row).TransRowId = InDT.Rows(DTRow)("tranline").ToString
                            openTrans.InternalReconciliationOpenTransRows.Item(Row).ReconcileAmount = RecAmount ' CDbl(InDT.Rows(DTRow)("paytot").ToString) '
                            Row += 1
                        End If
                    Next


                    Dim Reconum As Integer = 0
                    Try
                        reconParams = service.Add(openTrans)
                    Catch ex As Exception
                        If Reconum = 0 Then objaddon.objapplication.StatusBar.SetText("Reconciled Error : " & GetBranchName(RecRow.branch.ToString) & "-" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : Return False
                    End Try

                    Reconum = reconParams.ReconNum

                    For DTRow As Integer = 0 To InDT.Rows.Count - 1
                        If InDT.Rows(DTRow)("branchc").ToString = RecRow.branch.ToString Then
                            Matrix0.Columns.Item("recono").Cells.Item(CInt(DTRow + 1)).Specific.String = Reconum 'CInt(InDT.Rows(DTRow)("#").ToString)
                            objRs.DoQuery("Update OITR Set ""U_TransId""='" & transid & "' where ""ReconNum""='" & Reconum & "'")
                            objaddon.objapplication.StatusBar.SetText("Reconciled successfully..." & GetBranchName(RecRow.branch.ToString), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                        End If
                    Next
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(openTrans)
                    GC.Collect()
                Next

                Return True
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("Recon: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End Try

        End Function

        Private Function BranchReconciliation_Consolidated(ByVal InDT As DataTable, ByVal Branch As String, ByVal DLine As String) As Boolean
            Try
                Dim service As IInternalReconciliationsService = objaddon.objcompany.GetCompanyService().GetBusinessService(ServiceTypes.InternalReconciliationsService)
                Dim openTrans As InternalReconciliationOpenTrans = service.GetDataInterface(InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationOpenTrans)
                Dim reconParams As IInternalReconciliationParams = service.GetDataInterface(InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationParams)
                openTrans.CardOrAccount = CardOrAccountEnum.coaCard
                Dim RecAmount As Double
                Dim Row As Integer = 0
                'Dim Line As String = ""
                objRs = objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset)

                'For DTRow As Integer = 0 To InDT.Rows.Count - 1
                '    If InDT.Rows(DTRow)("branchc").ToString = Branch And Not InDT.Rows(DTRow)("recono").ToString = String.Empty Then
                '        Line = CInt(InDT.Rows(DTRow)("#").ToString)
                '        Exit For
                '    End If
                'Next
                'If Line = "" Then
                Dim DocDate As Date = Date.ParseExact(EditText0.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                openTrans.ReconDate = DocDate
                For DTRow As Integer = 0 To InDT.Rows.Count - 1
                    If InDT.Rows(DTRow)("branchc").ToString = Branch And InDT.Rows(DTRow)("Row").ToString = DLine Then
                        'If Line = "" Then Line = InDT.Rows(DTRow)("#").ToString
                        If InDT.Rows(DTRow)("debcred").ToString = "C" Then RecAmount = Math.Round(-CDbl(InDT.Rows(DTRow)("paytot").ToString), SumRound) Else RecAmount = Math.Round(CDbl(InDT.Rows(DTRow)("paytot").ToString), SumRound)
                        openTrans.InternalReconciliationOpenTransRows.Add()
                        openTrans.InternalReconciliationOpenTransRows.Item(Row).Selected = BoYesNoEnum.tYES
                        openTrans.InternalReconciliationOpenTransRows.Item(Row).TransId = InDT.Rows(DTRow)("trannum").ToString
                        openTrans.InternalReconciliationOpenTransRows.Item(Row).TransRowId = InDT.Rows(DTRow)("tranline").ToString
                        openTrans.InternalReconciliationOpenTransRows.Item(Row).ReconcileAmount = RecAmount
                        Row += 1
                    End If
                Next
                Dim Reconum As Integer = 0
                Try
                    reconParams = service.Add(openTrans)
                Catch ex As Exception
                    objaddon.objapplication.StatusBar.SetText("Reconciled Error..." & GetBranchName(Branch) & "-" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : Return False
                End Try
                Reconum = reconParams.ReconNum
                For DTRow As Integer = 0 To InDT.Rows.Count - 1
                    'Dim MRow As Integer = CInt(InDT.Rows(DTRow)("#").ToString)
                    If InDT.Rows(DTRow)("branchc").ToString = Branch Then
                        Matrix0.Columns.Item("recono").Cells.Item(CInt(DTRow + 1)).Specific.String = Reconum 'CInt(InDT.Rows(DTRow)("#").ToString)
                    End If
                Next


                objaddon.objapplication.StatusBar.SetText("Reconciled successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                'Matrix1.Columns.Item("recono").Cells.Item(CInt(Line)).Specific.String = Reconum
                'objaddon.objapplication.StatusBar.SetText("Reconciled successfully..." & GetBranchName(Branch), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(openTrans)
                GC.Collect()
                'End If

                Return True
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("Recon: " & GetBranchName(Branch) & "-" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End Try

        End Function

#End Region

#Region "Matrix Events"

        Private Sub Matrix1_PressedAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix1.PressedAfter
            Try
                objCheck = Matrix1.Columns.Item("select").Cells.Item(pVal.Row).Specific
                If pVal.ColUID = "select" Then
                    If objCheck.Checked = True Then
                        'Matrix1.SelectRow(pVal.Row, True, True)
                        Matrix1.CommonSetting.SetRowBackColor(pVal.Row, Color.PeachPuff.ToArgb)
                    Else
                        'Matrix1.SelectRow(pVal.Row, False, True)
                        Matrix1.CommonSetting.SetRowBackColor(pVal.Row, Matrix1.Item.BackColor)
                        Matrix1.SetCellWithoutValidation(pVal.Row, "paytot", CDbl(Matrix1.Columns.Item("baldue").Cells.Item(pVal.Row).Specific.String))
                    End If
                    'Calculate_Total()
                    Calc_Total_14092021(pVal.Row)
                    Matrix_DataTable(pVal.Row, pVal.ColUID)
                    If BranchSplitReconciliation = "N" Then Exit Sub
                    Dim paytotsum As String = objActualDT.Compute("SUM(paytot)", "").ToString
                    If paytotsum = "0" Then Exit Sub
                    If objActualDT.Rows.Count = 1 Then
                        If Matrix0.VisualRowCount > 0 Then Matrix0.Clear()
                    End If

                    If objActualDT.Rows.Count <= 1 Then Exit Sub
                    Dim MatRow As Integer = CInt(objActualDT.Rows(objActualDT.Rows.Count - 1)("#").ToString)
                    If Matrix1.Columns.Item("jeadj").Cells.Item(CInt(objActualDT.Rows(objActualDT.Rows.Count - 1)("#").ToString)).Specific.String = "" Then
                        objCheck = Matrix1.Columns.Item("select").Cells.Item(pVal.Row).Specific
                        objCheck.Checked = False
                        Matrix1.CommonSetting.SetRowBackColor(pVal.Row, Matrix1.Item.BackColor)
                        'objaddon.objapplication.StatusBar.SetText("Please adjust the previously selected transaction. ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        For DTRow As Integer = 0 To objActualDT.Rows.Count - 1
                            If objActualDT.Rows(DTRow)("#").ToString = Matrix1.Columns.Item("#").Cells.Item(pVal.Row).Specific.Value Then
                                objActualDT.Rows(DTRow).Delete()
                                Exit For
                            End If
                        Next
                        For DTRow As Integer = 0 To oSelectedDT.Rows.Count - 1
                            If oSelectedDT.Rows(DTRow)("#").ToString = Matrix1.Columns.Item("#").Cells.Item(pVal.Row).Specific.Value Then
                                oSelectedDT.Rows(DTRow).Delete()
                                Exit For
                            End If
                        Next
                        Exit Sub

                    End If
                ElseIf pVal.ColUID = "tranadj" Then
                    Dim objtranadj As SAPbouiCOM.CheckBox
                    objtranadj = Matrix1.Columns.Item("tranadj").Cells.Item(pVal.Row).Specific
                    If objCheck.Checked = True And objtranadj.Checked = True Then
                        If BranchSplitReconciliation = "Y" Then
                            Dim tranid As String = Matrix1.Columns.Item("trannum").Cells.Item(pVal.Row).Specific.String
                            Dim objPayform As SAPbouiCOM.Form
                            Dim objMatrix As SAPbouiCOM.Matrix
                            Dim CardCode As String = ""
                            objPayform = objaddon.objapplication.Forms.GetForm("PAYINIT", 1)
                            objMatrix = objPayform.Items.Item("mtxdata").Specific
                            For i As Integer = 1 To objMatrix.VisualRowCount
                                If objMatrix.Columns.Item("bpcode").Cells.Item(i).Specific.String <> "" Then
                                    If i = 1 Then
                                        CardCode = "'" + objMatrix.Columns.Item("bpcode").Cells.Item(i).Specific.String + "'"
                                    Else
                                        CardCode += ",'" + objMatrix.Columns.Item("bpcode").Cells.Item(i).Specific.String + "'"
                                    End If
                                End If
                            Next
                            If Matrix1.Columns.Item("jeadj").Cells.Item(pVal.Row).Specific.String = "" Then
                                If Not objaddon.FormExist("AdjTran") Then
                                    Dim activeform As New Frm_GetReco_AdjustmentTrans
                                    activeform.Show()
                                    activeform.Get_InternalReconciliation_Adjustment_Query(objActualDT.Rows(objActualDT.Rows.Count - 1), PayInitDate.ToString("yyyyMMdd"), CardCode, tranid, CDbl(Matrix1.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String), Matrix1.Columns.Item("branchc").Cells.Item(pVal.Row).Specific.String, Matrix1.Columns.Item("branchn").Cells.Item(pVal.Row).Specific.String, Matrix1.Columns.Item("cardc").Cells.Item(pVal.Row).Specific.String)
                                End If
                            End If
                        End If
                    Else
                        'objaddon.objapplication.StatusBar.SetText("Select the transaction... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        objtranadj.Checked = False
                        Matrix1.Columns.Item("jeadj").Cells.Item(pVal.Row).Specific.String = ""
                        If objActualDT.Columns.Contains("TBranchC") Then
                            For DTRow = objActualDT.Rows.Count - 1 To 0 Step -1
                                If objActualDT.Rows(DTRow)("TBranchC").ToString <> "" Then
                                    objActualDT.Rows(DTRow).Delete()
                                End If
                            Next
                            'For DTRow As Integer = 0 To oSelectedDT.Rows.Count - 1
                            '    If oSelectedDT.Rows(DTRow)("TBranchC").ToString <> "" Then
                            '        oSelectedDT.Rows(DTRow).Delete()
                            '    End If
                            'Next
                        End If
                        If Matrix0.VisualRowCount > 0 Then
                            Folder1.Item.Click()
                            Matrix0.Clear()
                            Folder0.Item.Click()
                        End If

                    End If
                End If
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Matrix1_LinkPressedBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Matrix1.LinkPressedBefore
            Dim ColItem As SAPbouiCOM.Column = Matrix1.Columns.Item("origin")
            Dim objlink As SAPbouiCOM.LinkedButton = ColItem.ExtendedObject
            Dim oForm As SAPbouiCOM.Form
            Try
                Select Case pVal.ColUID
                    Case "origin"
                        If Matrix1.Columns.Item("object").Cells.Item(pVal.Row).Specific.String = "13" Then
                            objaddon.objapplication.Menus.Item("2053").Activate()  'AR Invoice
                            oForm = objaddon.objapplication.Forms.ActiveForm
                            oForm.Freeze(True)
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            oForm.Items.Item("8").Specific.String = Matrix1.Columns.Item("originno").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("10").Specific.String = Matrix1.Columns.Item("date").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oForm.Freeze(False)
                            BubbleEvent = False
                        ElseIf Matrix1.Columns.Item("object").Cells.Item(pVal.Row).Specific.String = "14" Then
                            objaddon.objapplication.Menus.Item("2055").Activate()  'AR Credit Memo
                            oForm = objaddon.objapplication.Forms.ActiveForm
                            oForm.Freeze(True)
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            oForm.Items.Item("8").Specific.String = Matrix1.Columns.Item("originno").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("10").Specific.String = Matrix1.Columns.Item("date").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oForm.Freeze(False)
                            BubbleEvent = False
                        ElseIf Matrix1.Columns.Item("object").Cells.Item(pVal.Row).Specific.String = "18" Then
                            objaddon.objapplication.Menus.Item("2308").Activate()  'AP Invoice
                            oForm = objaddon.objapplication.Forms.ActiveForm
                            oForm.Freeze(True)
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            oForm.Items.Item("8").Specific.String = Matrix1.Columns.Item("originno").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("10").Specific.String = Matrix1.Columns.Item("date").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oForm.Freeze(False)
                            BubbleEvent = False
                        ElseIf Matrix1.Columns.Item("object").Cells.Item(pVal.Row).Specific.String = "19" Then
                            objaddon.objapplication.Menus.Item("2309").Activate()  'AP Credit Memo
                            oForm = objaddon.objapplication.Forms.ActiveForm
                            oForm.Freeze(True)
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            oForm.Items.Item("8").Specific.String = Matrix1.Columns.Item("originno").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("10").Specific.String = Matrix1.Columns.Item("date").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oForm.Freeze(False)
                            BubbleEvent = False
                        ElseIf Matrix1.Columns.Item("object").Cells.Item(pVal.Row).Specific.String = "203" Then
                            objaddon.objapplication.Menus.Item("2071").Activate()  'AR DownPayment
                            oForm = objaddon.objapplication.Forms.ActiveForm
                            oForm.Freeze(True)
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            oForm.Items.Item("8").Specific.String = Matrix1.Columns.Item("originno").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("10").Specific.String = Matrix1.Columns.Item("date").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oForm.Freeze(False)
                            BubbleEvent = False
                        ElseIf Matrix1.Columns.Item("object").Cells.Item(pVal.Row).Specific.String = "204" Then
                            objaddon.objapplication.Menus.Item("2317").Activate()  'AP DownPayment
                            oForm = objaddon.objapplication.Forms.ActiveForm
                            oForm.Freeze(True)
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            oForm.Items.Item("8").Specific.String = Matrix1.Columns.Item("originno").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("10").Specific.String = Matrix1.Columns.Item("date").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oForm.Freeze(False)
                            BubbleEvent = False
                        ElseIf Matrix1.Columns.Item("object").Cells.Item(pVal.Row).Specific.String = "24" Then
                            objaddon.objapplication.Menus.Item("2817").Activate()  'Incoming Payment
                            oForm = objaddon.objapplication.Forms.ActiveForm
                            oForm.Freeze(True)
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            oForm.Items.Item("3").Specific.String = Matrix1.Columns.Item("originno").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("10").Specific.String = Matrix1.Columns.Item("date").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oForm.Freeze(False)
                            BubbleEvent = False
                        ElseIf Matrix1.Columns.Item("object").Cells.Item(pVal.Row).Specific.String = "46" Then
                            objaddon.objapplication.Menus.Item("2818").Activate()  'Outgoing Payment
                            oForm = objaddon.objapplication.Forms.ActiveForm
                            oForm.Freeze(True)
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            oForm.Items.Item("3").Specific.String = Matrix1.Columns.Item("originno").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("10").Specific.String = Matrix1.Columns.Item("date").Cells.Item(pVal.Row).Specific.String
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
                            oForm.Items.Item("5").Specific.String = Matrix1.Columns.Item("trannum").Cells.Item(pVal.Row).Specific.String
                            'oForm.Items.Item("10").Specific.String = Matrix1.Columns.Item("date").Cells.Item(pVal.Row).Specific.String
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

        Private Sub Matrix1_ValidateBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Matrix1.ValidateBefore
            Try
                'If pVal.ItemChanged = False Then Exit Sub
                If pVal.InnerEvent = True Then Exit Sub
                Dim Balance, PayTotal As Double
                'Dim ActTotal As Double
                If Val(Matrix1.Columns.Item("baldue").Cells.Item(pVal.Row).Specific.String) <> 0 Then Balance = CDbl(Matrix1.Columns.Item("baldue").Cells.Item(pVal.Row).Specific.String) Else Balance = 0
                If Val(Matrix1.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String) <> 0 Then PayTotal = CDbl(Matrix1.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String) Else PayTotal = 0
                'If Val(Matrix1.Columns.Item("pay").Cells.Item(pVal.Row).Specific.String) > 0 Then ActTotal = CDbl(Matrix1.Columns.Item("pay").Cells.Item(pVal.Row).Specific.String) Else ActTotal = 0
                Select Case pVal.ColUID
                    Case "paytot"
                        If pVal.InnerEvent = False Then
                            If PayTotal <= 0 Then
                                PayTotal = -CDbl(Matrix1.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String)
                                Balance = -CDbl(Matrix1.Columns.Item("baldue").Cells.Item(pVal.Row).Specific.String)
                                If PayTotal > Balance Or PayTotal = 0 Then
                                    'Matrix1.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String = CDbl(Matrix1.Columns.Item("baldue").Cells.Item(pVal.Row).Specific.String)
                                    Matrix1.SetCellWithoutValidation(pVal.Row, "paytot", CDbl(Matrix1.Columns.Item("baldue").Cells.Item(pVal.Row).Specific.String))
                                End If
                            ElseIf PayTotal > 0 Then
                                If PayTotal > Balance Then
                                    'Matrix1.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String = CDbl(Matrix1.Columns.Item("baldue").Cells.Item(pVal.Row).Specific.String)
                                    Matrix1.SetCellWithoutValidation(pVal.Row, "paytot", CDbl(Matrix1.Columns.Item("baldue").Cells.Item(pVal.Row).Specific.String))
                                End If
                            Else
                                Matrix1.SetCellWithoutValidation(pVal.Row, "paytot", CDbl(Matrix1.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String))
                            End If
                            If pVal.ItemChanged = True Then
                                objCheck = Matrix1.Columns.Item("select").Cells.Item(pVal.Row).Specific
                                objCheck.Checked = True
                                Matrix1.CommonSetting.SetRowBackColor(pVal.Row, Color.PeachPuff.ToArgb)
                                'Matrix1.SelectRow(pVal.Row, True, True)
                                Matrix1.SetCellWithoutValidation(pVal.Row, "paytot", CDbl(Matrix1.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String))
                            End If
                            Matrix1.Columns.Item("pay").Cells.Item(pVal.Row).Specific.String = Matrix1.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String
                            'Calculate_Total()
                            Calc_Total_14092021(pVal.Row)
                        End If
                        Matrix1.SetCellWithoutValidation(pVal.Row, "details", Matrix1.Columns.Item("details").Cells.Item(pVal.Row).Specific.String)


                End Select
                If pVal.ItemChanged = True Then
                    Matrix_DataTable(pVal.Row, pVal.ColUID)
                End If
                If BranchSplitReconciliation = "N" Then Exit Sub
                Dim paytotsum As String = objActualDT.Compute("SUM(paytot)", "").ToString
                If paytotsum = "0" Then Exit Sub
                If objActualDT.Rows.Count = 1 Then
                    If Matrix0.VisualRowCount > 0 Then Matrix0.Clear()
                End If
                If objActualDT.Rows.Count <= 1 Then Exit Sub
                Dim MatRow As Integer = CInt(objActualDT.Rows(objActualDT.Rows.Count - 1)("#").ToString)
                If Matrix1.Columns.Item("jeadj").Cells.Item(CInt(objActualDT.Rows(objActualDT.Rows.Count - 1)("#").ToString)).Specific.String = "" Then
                    objCheck = Matrix1.Columns.Item("select").Cells.Item(pVal.Row).Specific
                    objCheck.Checked = False
                    Matrix1.CommonSetting.SetRowBackColor(pVal.Row, Matrix1.Item.BackColor)
                    'objaddon.objapplication.StatusBar.SetText("Please adjust the previously selected transaction. ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    For DTRow As Integer = 0 To objActualDT.Rows.Count - 1
                        If objActualDT.Rows(DTRow)("#").ToString = Matrix1.Columns.Item("#").Cells.Item(pVal.Row).Specific.Value Then
                            objActualDT.Rows(DTRow).Delete()
                            'Exit For
                        End If
                    Next
                    For DTRow As Integer = 0 To oSelectedDT.Rows.Count - 1
                        If oSelectedDT.Rows(DTRow)("#").ToString = Matrix1.Columns.Item("#").Cells.Item(pVal.Row).Specific.Value Then
                            oSelectedDT.Rows(DTRow).Delete()
                            'Exit For
                        End If
                    Next
                    'Matrix1.Columns.Item("paytot").Cells.Item(CInt(objActualDT.Rows(objActualDT.Rows.Count - 1)("#").ToString)).Click()
                    Exit Sub
                End If
                'If CheckBox0.Checked = True Then
                '    objCheck = Matrix1.Columns.Item("tranadj").Cells.Item(pVal.Row).Specific
                '    If objCheck.Checked = True Then
                '        'If Matrix1.Columns.Item("jeadj").Cells.Item(CInt(objActualDT.Rows(objActualDT.Rows.Count - 1)("#").ToString)).Specific.String = "" Then
                '        '    objaddon.objapplication.StatusBar.SetText("Please adjust the previous transactions: " & CStr(objActualDT.Rows(objActualDT.Rows.Count - 1)("#").ToString), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '        '    Exit Sub
                '        'End If
                '        Dim tranid As String = Matrix1.Columns.Item("trannum").Cells.Item(pVal.Row).Specific.String
                '        Dim objPayform As SAPbouiCOM.Form
                '        Dim objMatrix As SAPbouiCOM.Matrix
                '        Dim CardCode As String = ""
                '        objPayform = objaddon.objapplication.Forms.GetForm("PAYINIT", 0)
                '        objMatrix = objPayform.Items.Item("mtxdata").Specific
                '        For i As Integer = 1 To objMatrix.VisualRowCount
                '            If objMatrix.Columns.Item("bpcode").Cells.Item(i).Specific.String <> "" Then
                '                If i = 1 Then
                '                    CardCode = "'" + objMatrix.Columns.Item("bpcode").Cells.Item(i).Specific.String + "'"
                '                Else
                '                    CardCode += ",'" + objMatrix.Columns.Item("bpcode").Cells.Item(i).Specific.String + "'"
                '                End If
                '            End If
                '        Next
                '        If Matrix1.Columns.Item("jeadj").Cells.Item(pVal.Row).Specific.String = "" Then
                '            If Not objaddon.FormExist("AdjTran") Then
                '                Dim activeform As New Frm_GetReco_AdjustmentTrans
                '                activeform.Show()
                '                activeform.Get_InternalReconciliation_Adjustment_Query(objActualDT.Rows(objActualDT.Rows.Count - 1), PayInitDate.ToString("yyyyMMdd"), CardCode, tranid, CDbl(Matrix1.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String), Matrix1.Columns.Item("branchc").Cells.Item(pVal.Row).Specific.String, Matrix1.Columns.Item("branchn").Cells.Item(pVal.Row).Specific.String, Matrix1.Columns.Item("cardc").Cells.Item(pVal.Row).Specific.String)

                '            End If
                '        End If
                '    End If
                'End If

            Catch ex As Exception

            End Try

        End Sub

        Private Sub DeleteRow()
            Try
                Dim Flag As Boolean = False
                'Dim objSelect As SAPbouiCOM.CheckBox
                'Matrix1.Columns.Item("select").TitleObject.Click(SAPbouiCOM.BoCellClickType.ct_Double)
                'For i As Integer = Matrix1.VisualRowCount To 1 Step -1
                '    objSelect = Matrix1.Columns.Item("select").Cells.Item(i).Specific
                '    If objSelect.Checked = False Then
                '        Matrix1.DeleteRow(i)
                '        odbdsDetails.RemoveRecord(i - 1)
                '        Flag = True
                '    End If
                'Next
                Matrix1.FlushToDataSource()
                For index = odbdsDetails.Size - 1 To 0 Step -1
                    If odbdsDetails.GetValue("U_Select", index) = "N" Then
                        odbdsDetails.RemoveRecord(index) : Flag = True
                    End If
                Next
                Matrix1.LoadFromDataSourceEx()
                If Flag = True Then
                    'For i As Integer = 1 To Matrix1.VisualRowCount
                    '    objSelect = Matrix1.Columns.Item("select").Cells.Item(i).Specific
                    '    If objSelect.Checked = True Then
                    '        Matrix1.Columns.Item("#").Cells.Item(i).Specific.String = i
                    '    End If
                    'Next

                    For index = 0 To odbdsDetails.Size - 1
                        odbdsDetails.SetValue("LineId", index, index + 1)
                    Next
                    Matrix1.LoadFromDataSourceEx()
                    objform.Freeze(False)
                    If objform.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    objform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    'objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                End If
            Catch ex As Exception
                objform.Freeze(False)
                'objAddOn.objApplication.StatusBar.SetText("DeleteRow  Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Finally
            End Try
        End Sub

        Private Sub Form_DataLoadAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo)
            Try
                objform.EnableMenu("1282", False)
                DeleteRow()
                For i As Integer = 1 To Matrix1.VisualRowCount
                    objCheck = Matrix1.Columns.Item("select").Cells.Item(i).Specific
                    If objCheck.Checked = True Then
                        If objCheck.Checked = True Then
                            'Matrix1.SelectRow(i, True, True)
                            Matrix1.CommonSetting.SetRowBackColor(i, Color.PeachPuff.ToArgb)
                        End If
                    End If
                Next
                Matrix1.AutoResizeColumns()
                Matrix0.AutoResizeColumns()
                Matrix0.Item.Enabled = False
                Matrix1.Item.Enabled = False
                'objform.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                If BranchSplitReconciliation = "Y" Then ' with splitup JE
                    If ComboBox1.Selected.Value = "O" Then
                        EditText6.Item.Visible = False
                        StaticText7.Item.Visible = False
                        LinkedButton1.Item.Visible = False
                    Else
                        EditText6.Item.Visible = True
                        StaticText7.Item.Visible = True
                        LinkedButton1.Item.Visible = True
                    End If
                Else
                    Matrix1.Columns.Item("revjeno").Visible = True
                    EditText6.Item.Visible = False
                    StaticText7.Item.Visible = False
                    LinkedButton1.Item.Visible = False
                End If
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Folder1_PressedAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Folder1.PressedAfter
            Try
                If Matrix0.VisualRowCount = 0 Then Exit Sub
                objform.Freeze(True)
                Matrix0.AutoResizeColumns()
                objform.Freeze(False)
            Catch ex As Exception
                objform.Freeze(False)
            End Try

        End Sub

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
                            objaddon.objapplication.Menus.Item("1540").Activate()  'JE
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

#End Region

    End Class
End Namespace
