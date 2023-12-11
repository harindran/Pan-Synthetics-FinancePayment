Option Strict Off
Option Explicit On

Imports System.Drawing
Imports SAPbobsCOM
Imports System.Linq
Imports SAPbouiCOM.Framework
Imports System.Runtime.CompilerServices
Namespace Finance_Payment
    <FormAttribute("FINPAY", "Business Objects/FrmInPayments.b1f")>
    Friend Class FrmInPayments
        Inherits UserFormBase
        Public WithEvents objform As SAPbouiCOM.Form
        Private WithEvents pCFL As SAPbouiCOM.ISBOChooseFromListEventArg
        Public WithEvents odbdsDetails As SAPbouiCOM.DBDataSource
        Public WithEvents odbdsHeader As SAPbouiCOM.DBDataSource
        Public Shared CurBranch, CurBPCode As String
        Dim FormCount As Integer = 0
        Dim strSQL As String
        Dim objRs As SAPbobsCOM.Recordset
        Dim Total As Double
        Public Shared objFinalDT As New DataTable
        Public Shared objActualDT As New DataTable
        Public Shared oSelectedDT As New DataTable
        Private WithEvents objCheck As SAPbouiCOM.CheckBox
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("1").Specific, SAPbouiCOM.Button)
            Me.Button1 = CType(Me.GetItem("2").Specific, SAPbouiCOM.Button)
            Me.Matrix0 = CType(Me.GetItem("mtxcont").Specific, SAPbouiCOM.Matrix)
            Me.StaticText0 = CType(Me.GetItem("ldocdate").Specific, SAPbouiCOM.StaticText)
            Me.EditText0 = CType(Me.GetItem("tdocdate").Specific, SAPbouiCOM.EditText)
            Me.StaticText1 = CType(Me.GetItem("lremark").Specific, SAPbouiCOM.StaticText)
            Me.EditText1 = CType(Me.GetItem("tremark").Specific, SAPbouiCOM.EditText)
            Me.StaticText2 = CType(Me.GetItem("ldocnum").Specific, SAPbouiCOM.StaticText)
            Me.ComboBox0 = CType(Me.GetItem("Series").Specific, SAPbouiCOM.ComboBox)
            Me.EditText2 = CType(Me.GetItem("t_docnum").Specific, SAPbouiCOM.EditText)
            Me.StaticText3 = CType(Me.GetItem("ltotdue").Specific, SAPbouiCOM.StaticText)
            Me.EditText3 = CType(Me.GetItem("ttotdue").Specific, SAPbouiCOM.EditText)
            Me.Folder0 = CType(Me.GetItem("fldrcont").Specific, SAPbouiCOM.Folder)
            Me.Folder1 = CType(Me.GetItem("fldr2").Specific, SAPbouiCOM.Folder)
            Me.Button4 = CType(Me.GetItem("paymeans").Specific, SAPbouiCOM.Button)
            Me.StaticText4 = CType(Me.GetItem("lbiref").Specific, SAPbouiCOM.StaticText)
            Me.EditText4 = CType(Me.GetItem("tbidate").Specific, SAPbouiCOM.EditText)
            Me.StaticText5 = CType(Me.GetItem("lbidate").Specific, SAPbouiCOM.StaticText)
            Me.EditText5 = CType(Me.GetItem("tbitot").Specific, SAPbouiCOM.EditText)
            Me.StaticText6 = CType(Me.GetItem("lbtot").Specific, SAPbouiCOM.StaticText)
            Me.StaticText7 = CType(Me.GetItem("lbigl").Specific, SAPbouiCOM.StaticText)
            Me.EditText6 = CType(Me.GetItem("tbigl").Specific, SAPbouiCOM.EditText)
            Me.EditText7 = CType(Me.GetItem("tbiref").Specific, SAPbouiCOM.EditText)
            Me.StaticText9 = CType(Me.GetItem("lcigl").Specific, SAPbouiCOM.StaticText)
            Me.EditText8 = CType(Me.GetItem("tcigl").Specific, SAPbouiCOM.EditText)
            Me.EditText9 = CType(Me.GetItem("tcitot").Specific, SAPbouiCOM.EditText)
            Me.StaticText10 = CType(Me.GetItem("lcitot").Specific, SAPbouiCOM.StaticText)
            Me.StaticText11 = CType(Me.GetItem("ltran").Specific, SAPbouiCOM.StaticText)
            Me.StaticText12 = CType(Me.GetItem("lcash").Specific, SAPbouiCOM.StaticText)
            Me.StaticText8 = CType(Me.GetItem("lchigl").Specific, SAPbouiCOM.StaticText)
            Me.EditText10 = CType(Me.GetItem("tchigl").Specific, SAPbouiCOM.EditText)
            Me.StaticText13 = CType(Me.GetItem("lchek").Specific, SAPbouiCOM.StaticText)
            Me.Matrix1 = CType(Me.GetItem("mtxcheq").Specific, SAPbouiCOM.Matrix)
            Me.EditText11 = CType(Me.GetItem("tchtot").Specific, SAPbouiCOM.EditText)
            Me.StaticText14 = CType(Me.GetItem("lchtot").Specific, SAPbouiCOM.StaticText)
            Me.Matrix2 = CType(Me.GetItem("mtxcr").Specific, SAPbouiCOM.Matrix)
            Me.EditText12 = CType(Me.GetItem("tcrtot").Specific, SAPbouiCOM.EditText)
            Me.StaticText15 = CType(Me.GetItem("lcrtot").Specific, SAPbouiCOM.StaticText)
            Me.StaticText16 = CType(Me.GetItem("lcred").Specific, SAPbouiCOM.StaticText)
            Me.StaticText17 = CType(Me.GetItem("ltranno").Specific, SAPbouiCOM.StaticText)
            Me.EditText13 = CType(Me.GetItem("ttranno").Specific, SAPbouiCOM.EditText)
            Me.LinkedButton0 = CType(Me.GetItem("lktran").Specific, SAPbouiCOM.LinkedButton)
            Me.StaticText18 = CType(Me.GetItem("ljeno").Specific, SAPbouiCOM.StaticText)
            Me.EditText14 = CType(Me.GetItem("tjeno").Specific, SAPbouiCOM.EditText)
            Me.LinkedButton1 = CType(Me.GetItem("lnkje").Specific, SAPbouiCOM.LinkedButton)
            Me.StaticText19 = CType(Me.GetItem("lrecno").Specific, SAPbouiCOM.StaticText)
            Me.EditText15 = CType(Me.GetItem("trecno").Specific, SAPbouiCOM.EditText)
            Me.LinkedButton2 = CType(Me.GetItem("lnkrec").Specific, SAPbouiCOM.LinkedButton)
            Me.StaticText20 = CType(Me.GetItem("lbcggl").Specific, SAPbouiCOM.StaticText)
            Me.EditText16 = CType(Me.GetItem("tbcggl").Specific, SAPbouiCOM.EditText)
            Me.EditText17 = CType(Me.GetItem("tbcgtot").Specific, SAPbouiCOM.EditText)
            Me.StaticText21 = CType(Me.GetItem("lbcgtot").Specific, SAPbouiCOM.StaticText)
            Me.StaticText22 = CType(Me.GetItem("lbcg").Specific, SAPbouiCOM.StaticText)
            Me.StaticText23 = CType(Me.GetItem("lcurr").Specific, SAPbouiCOM.StaticText)
            Me.EditText18 = CType(Me.GetItem("tcurr").Specific, SAPbouiCOM.EditText)
            Me.EditText19 = CType(Me.GetItem("tcurtot").Specific, SAPbouiCOM.EditText)
            Me.StaticText25 = CType(Me.GetItem("lcurtot").Specific, SAPbouiCOM.StaticText)
            Me.StaticText26 = CType(Me.GetItem("ltotfc").Specific, SAPbouiCOM.StaticText)
            Me.EditText20 = CType(Me.GetItem("ttotfc").Specific, SAPbouiCOM.EditText)
            Me.EditText21 = CType(Me.GetItem("tboeamt").Specific, SAPbouiCOM.EditText)
            Me.StaticText27 = CType(Me.GetItem("lboeamt").Specific, SAPbouiCOM.StaticText)
            Me.EditText24 = CType(Me.GetItem("tfc").Specific, SAPbouiCOM.EditText)
            Me.StaticText30 = CType(Me.GetItem("lfc").Specific, SAPbouiCOM.StaticText)
            Me.StaticText28 = CType(Me.GetItem("lforex").Specific, SAPbouiCOM.StaticText)
            Me.EditText22 = CType(Me.GetItem("tforex").Specific, SAPbouiCOM.EditText)
            Me.LinkedButton3 = CType(Me.GetItem("lkforex").Specific, SAPbouiCOM.LinkedButton)
            Me.StaticText31 = CType(Me.GetItem("lcfxje").Specific, SAPbouiCOM.StaticText)
            Me.EditText25 = CType(Me.GetItem("tcfxje").Specific, SAPbouiCOM.EditText)
            Me.LinkedButton5 = CType(Me.GetItem("lkcfxje").Specific, SAPbouiCOM.LinkedButton)
            Me.StaticText24 = CType(Me.GetItem("lacttot").Specific, SAPbouiCOM.StaticText)
            Me.EditText23 = CType(Me.GetItem("tacttot").Specific, SAPbouiCOM.EditText)
            Me.EditText26 = CType(Me.GetItem("tblexrate").Specific, SAPbouiCOM.EditText)
            Me.StaticText29 = CType(Me.GetItem("lblexrate").Specific, SAPbouiCOM.StaticText)
            Me.EditText27 = CType(Me.GetItem("tacttotal").Specific, SAPbouiCOM.EditText)
            Me.StaticText32 = CType(Me.GetItem("lblacttot").Specific, SAPbouiCOM.StaticText)
            Me.StaticText33 = CType(Me.GetItem("lpaydate").Specific, SAPbouiCOM.StaticText)
            Me.EditText28 = CType(Me.GetItem("tpaydate").Specific, SAPbouiCOM.EditText)
            Me.StaticText34 = CType(Me.GetItem("linseries").Specific, SAPbouiCOM.StaticText)
            Me.EditText29 = CType(Me.GetItem("tinseries").Specific, SAPbouiCOM.EditText)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler CloseAfter, AddressOf Me.Form_CloseAfter
            AddHandler ResizeAfter, AddressOf Me.Form_ResizeAfter
            AddHandler DataLoadAfter, AddressOf Me.Form_DataLoadAfter

        End Sub

        Private Sub OnCustomInitialize()
            Try
                objform = objaddon.objapplication.Forms.GetForm("FINPAY", Me.FormCount)
                objform = objaddon.objapplication.Forms.ActiveForm
                objform.Freeze(True)
                odbdsHeader = objform.DataSources.DBDataSources.Item(CType(0, Object)) 'Item("@MI_ORCT")
                odbdsDetails = objform.DataSources.DBDataSources.Item(CType(1, Object)) 'Item("@MI_RCT1")
                objaddon.objglobalmethods.LoadSeries(objform, odbdsHeader, "MIORCT")
                oSelectedDT.Clear()
                If oSelectedDT.Columns.Count = 0 Then
                    oSelectedDT.Columns.Add("paytot", GetType(Double))
                    oSelectedDT.Columns.Add("doccur", GetType(String))
                    oSelectedDT.Columns.Add("#", GetType(String))
                End If
                objActualDT.Clear()
                If objActualDT.Columns.Count = 0 Then
                    For iCol As Integer = 0 To Matrix0.Columns.Count - 1
                        If Matrix0.Columns.Item(iCol).UniqueID = "paytot" Then
                            objActualDT.Columns.Add(Matrix0.Columns.Item(iCol).UniqueID, GetType(Double))
                        Else
                            objActualDT.Columns.Add(Matrix0.Columns.Item(iCol).UniqueID)
                        End If
                    Next
                End If
                objform.Items.Item("tdocdate").Specific.string = Now.Date.ToString("yyyyMMdd")
                objform.Items.Item("tpaydate").Specific.string = PayInitDate.ToString("yyyyMMdd")
                objform.Items.Item("tremark").Specific.String = "Created By " & objaddon.objcompany.UserName & " on " & Now.ToString
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If objaddon.objglobalmethods.Get_Branch_Assigned_Series("24", PayInitDate.ToString("yyyyMMdd")) Then
                    StaticText34.Item.Visible = False
                    EditText29.Item.Visible = False
                Else
                    StaticText34.Item.Visible = True
                    EditText29.Item.Visible = True
                End If
                objRs.DoQuery("select distinct T0.""BnkChgAct"" as ""BCGAcct"",T0.""LinkAct_3"" as ""CahAcct"",""LinkAct_24"" as ""Rounding"",""GLGainXdif"",""GLLossXdif"",""ExDiffAct"" " &
                              ",(Select ""SumDec"" from OADM) as ""SumDec"",(Select ""RateDec"" from OADM) as ""RateDec""" &
                              "from OACP T0 left join OFPR T1 on T1.""Category""=T0.""PeriodCat"" where T0.""PeriodCat""=(Select ""Category"" from OFPR where CURRENT_DATE Between ""F_RefDate"" and ""T_RefDate"")")
                If objRs.RecordCount > 0 Then
                    If objRs.Fields.Item(0).Value.ToString <> "" Then BCGAcct = objRs.Fields.Item(0).Value.ToString
                    If objRs.Fields.Item(1).Value.ToString <> "" Then CashAcct = objRs.Fields.Item(1).Value.ToString
                    If objRs.Fields.Item(2).Value.ToString <> "" Then RoundAcct = objRs.Fields.Item(2).Value.ToString
                    If objRs.Fields.Item(3).Value.ToString <> "" Then Forexgain = objRs.Fields.Item(3).Value.ToString
                    If objRs.Fields.Item(4).Value.ToString <> "" Then Forexloss = objRs.Fields.Item(4).Value.ToString
                    If objRs.Fields.Item(5).Value.ToString <> "" Then ForexDiff = objRs.Fields.Item(5).Value.ToString
                    If objRs.Fields.Item(6).Value.ToString <> "" Then SumRound = objRs.Fields.Item(6).Value.ToString
                    If objRs.Fields.Item(7).Value.ToString <> "" Then RateRound = objRs.Fields.Item(7).Value.ToString
                End If
                StaticText25.Caption = MainCurr + " Total"
                Field_Setup()
                If CostCenter = "S" Then
                    Matrix0.Columns.Item("distrule").Visible = False
                    If objaddon.HANA Then
                        objRs.DoQuery("select 'cc'||""DimCode"" as ""Code"",* from ODIM where ""DimActive""='Y'")
                    Else
                        objRs.DoQuery("select 'cc'& DimCode as Code,* from ODIM where DimActive='Y'")
                    End If
                    If objRs.RecordCount > 0 Then
                        For i As Integer = 0 To objRs.RecordCount - 1
                            Matrix0.Columns.Item(objRs.Fields.Item("Code").Value.ToString).Visible = True
                            Matrix0.Columns.Item(objRs.Fields.Item("Code").Value.ToString).TitleObject.Caption = objRs.Fields.Item("DimDesc").Value.ToString
                            objRs.MoveNext()
                        Next
                    End If
                End If
                Folder0.Item.Click()
                If Not LoadData(Query) Then
                    objform.Close()
                End If
                objform.Freeze(False)
            Catch ex As Exception
                objform.Freeze(False)
            End Try
        End Sub

#Region "Fields"
        Private WithEvents Button0 As SAPbouiCOM.Button
        Private WithEvents Button1 As SAPbouiCOM.Button
        Private WithEvents Matrix0 As SAPbouiCOM.Matrix
        Private WithEvents StaticText0 As SAPbouiCOM.StaticText
        Private WithEvents EditText0 As SAPbouiCOM.EditText
        Private WithEvents StaticText1 As SAPbouiCOM.StaticText
        Private WithEvents EditText1 As SAPbouiCOM.EditText
        Private WithEvents StaticText2 As SAPbouiCOM.StaticText
        Private WithEvents ComboBox0 As SAPbouiCOM.ComboBox
        Private WithEvents EditText2 As SAPbouiCOM.EditText
        Private WithEvents StaticText3 As SAPbouiCOM.StaticText
        Private WithEvents EditText3 As SAPbouiCOM.EditText
        Private WithEvents Folder0 As SAPbouiCOM.Folder
        Private WithEvents Folder1 As SAPbouiCOM.Folder
        Private WithEvents StaticText4 As SAPbouiCOM.StaticText
        Private WithEvents EditText4 As SAPbouiCOM.EditText
        Private WithEvents StaticText5 As SAPbouiCOM.StaticText
        Private WithEvents EditText5 As SAPbouiCOM.EditText
        Private WithEvents StaticText6 As SAPbouiCOM.StaticText
        Private WithEvents StaticText7 As SAPbouiCOM.StaticText
        Private WithEvents EditText6 As SAPbouiCOM.EditText
        Private WithEvents EditText7 As SAPbouiCOM.EditText
        Private WithEvents StaticText9 As SAPbouiCOM.StaticText
        Private WithEvents EditText8 As SAPbouiCOM.EditText
        Private WithEvents EditText9 As SAPbouiCOM.EditText
        Private WithEvents StaticText10 As SAPbouiCOM.StaticText
        Private WithEvents StaticText11 As SAPbouiCOM.StaticText
        Private WithEvents StaticText12 As SAPbouiCOM.StaticText
        Private WithEvents Button4 As SAPbouiCOM.Button
        Private WithEvents StaticText8 As SAPbouiCOM.StaticText
        Private WithEvents EditText10 As SAPbouiCOM.EditText
        Private WithEvents StaticText13 As SAPbouiCOM.StaticText
        Private WithEvents Matrix1 As SAPbouiCOM.Matrix
        Private WithEvents EditText11 As SAPbouiCOM.EditText
        Private WithEvents StaticText14 As SAPbouiCOM.StaticText
        Private WithEvents Matrix2 As SAPbouiCOM.Matrix
        Private WithEvents EditText12 As SAPbouiCOM.EditText
        Private WithEvents StaticText15 As SAPbouiCOM.StaticText
        Private WithEvents StaticText16 As SAPbouiCOM.StaticText
        Private WithEvents StaticText17 As SAPbouiCOM.StaticText
        Private WithEvents EditText13 As SAPbouiCOM.EditText
        Private WithEvents LinkedButton0 As SAPbouiCOM.LinkedButton
        Private WithEvents StaticText18 As SAPbouiCOM.StaticText
        Private WithEvents EditText14 As SAPbouiCOM.EditText
        Private WithEvents LinkedButton1 As SAPbouiCOM.LinkedButton
        Private WithEvents StaticText19 As SAPbouiCOM.StaticText
        Private WithEvents EditText15 As SAPbouiCOM.EditText
        Private WithEvents LinkedButton2 As SAPbouiCOM.LinkedButton
        Private WithEvents StaticText20 As SAPbouiCOM.StaticText
        Private WithEvents EditText16 As SAPbouiCOM.EditText
        Private WithEvents EditText17 As SAPbouiCOM.EditText
        Private WithEvents StaticText21 As SAPbouiCOM.StaticText
        Private WithEvents StaticText22 As SAPbouiCOM.StaticText
        Private WithEvents StaticText23 As SAPbouiCOM.StaticText
        Private WithEvents EditText18 As SAPbouiCOM.EditText
        Private WithEvents EditText19 As SAPbouiCOM.EditText
        Private WithEvents StaticText25 As SAPbouiCOM.StaticText
        Private WithEvents StaticText26 As SAPbouiCOM.StaticText
        Private WithEvents EditText20 As SAPbouiCOM.EditText
        Private WithEvents EditText21 As SAPbouiCOM.EditText
        Private WithEvents StaticText27 As SAPbouiCOM.StaticText
        Private WithEvents EditText24 As SAPbouiCOM.EditText
        Private WithEvents StaticText30 As SAPbouiCOM.StaticText
        Private WithEvents StaticText28 As SAPbouiCOM.StaticText
        Private WithEvents EditText22 As SAPbouiCOM.EditText
        Private WithEvents LinkedButton3 As SAPbouiCOM.LinkedButton
        Private WithEvents StaticText31 As SAPbouiCOM.StaticText
        Private WithEvents EditText25 As SAPbouiCOM.EditText
        Private WithEvents LinkedButton5 As SAPbouiCOM.LinkedButton
        Private WithEvents StaticText24 As SAPbouiCOM.StaticText
        Private WithEvents EditText23 As SAPbouiCOM.EditText
        Private WithEvents EditText26 As SAPbouiCOM.EditText
        Private WithEvents StaticText29 As SAPbouiCOM.StaticText
        Private WithEvents EditText27 As SAPbouiCOM.EditText
        Private WithEvents StaticText32 As SAPbouiCOM.StaticText
        Private WithEvents StaticText33 As SAPbouiCOM.StaticText
        Private WithEvents EditText28 As SAPbouiCOM.EditText

#End Region

#Region "Matrix Events"

        Private Sub Matrix0_DoubleClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.DoubleClickAfter
            Try
                'Dim totamt As Double
                'If pVal.ColUID = "select" Then
                '    Dim tick As String = "N"
                '    Try
                '        objform.Freeze(True)
                '        Matrix0.FlushToDataSource()
                '        If odbdsDetails.GetValue("U_Select", 0).Trim().ToUpper() = "N" Then
                '            tick = "Y"
                '        End If
                '        For rowNum As Integer = 0 To odbdsDetails.Size - 1
                '            odbdsDetails.SetValue("U_Select", rowNum, tick)
                '            If tick = "Y" Then
                '                totamt += odbdsDetails.GetValue("U_PayTotal", rowNum).Trim()
                '            Else
                '                totamt = 0
                '            End If
                '        Next
                '        odbdsHeader.SetValue("U_Total", 0, totamt)
                '        Matrix0.LoadFromDataSource()
                '        Total = odbdsHeader.GetValue("U_Total", 0)
                '    Catch ex As Exception
                '    Finally
                '        objform.Freeze(False)
                '    End Try
                'End If
            Catch ex As Exception
                objform.Freeze(False)
            End Try

        End Sub

        Private Sub Matrix0_LostFocusAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.LostFocusAfter
            Try
                If pVal.Row = 0 Then Exit Sub
                If pVal.ActionSuccess = False Then Exit Sub
                Select Case pVal.ColUID
                    Case "distrule"
                        If Matrix0.Columns.Item("distrule").Cells.Item(pVal.Row).Specific.String <> "" Then
                            Dim code As String
                            code = Matrix0.Columns.Item("distrule").Cells.Item(pVal.Row).Specific.String
                            Dim CostCodes() As String = code.Split(";")
                            If CostCodes.Count > 0 Then
                                If CostCodes.ElementAtOrDefault(0) <> "" Then Matrix0.Columns.Item("cc1").Cells.Item(pVal.Row).Specific.String = CostCodes(0)
                                If CostCodes.ElementAtOrDefault(1) <> "" Then Matrix0.Columns.Item("cc2").Cells.Item(pVal.Row).Specific.String = CostCodes(1)
                                If CostCodes.ElementAtOrDefault(2) <> "" Then Matrix0.Columns.Item("cc3").Cells.Item(pVal.Row).Specific.String = CostCodes(2)
                                If CostCodes.ElementAtOrDefault(3) <> "" Then Matrix0.Columns.Item("cc4").Cells.Item(pVal.Row).Specific.String = CostCodes(3)
                                If CostCodes.ElementAtOrDefault(4) <> "" Then Matrix0.Columns.Item("cc5").Cells.Item(pVal.Row).Specific.String = CostCodes(4)
                            End If
                        End If
                        '        If Matrix0.Columns.Item("distrule").Cells.Item(pVal.Row).Specific.String = "" Then
                        '            If CostCenter = "U" Then
                        '                Dim oform As New FrmDistRule
                        '                oform.Show()
                        '            End If
                        '        End If
                End Select
                'objform.Freeze(True)
                'Matrix0.AutoResizeColumns()
                'objform.Freeze(False)
            Catch ex As Exception
                objform.Freeze(False)
            End Try

        End Sub

        Private Sub Matrix0_LinkPressedBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Matrix0.LinkPressedBefore
            Dim ColItem As SAPbouiCOM.Column = Matrix0.Columns.Item("docnum")
            Dim objlink As SAPbouiCOM.LinkedButton = ColItem.ExtendedObject
            Dim oForm As SAPbouiCOM.Form
            Try
                Select Case pVal.ColUID
                    Case "docnum"
                        If Matrix0.Columns.Item("object").Cells.Item(pVal.Row).Specific.String = "13" Then
                            objaddon.objapplication.Menus.Item("2053").Activate()  'AR Invoice
                            oForm = objaddon.objapplication.Forms.ActiveForm
                            oForm.Freeze(True)
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            oForm.Items.Item("8").Specific.String = Matrix0.Columns.Item("docnum").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("10").Specific.String = Matrix0.Columns.Item("date").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oForm.Freeze(False)
                            BubbleEvent = False
                        ElseIf Matrix0.Columns.Item("object").Cells.Item(pVal.Row).Specific.String = "14" Then
                            objaddon.objapplication.Menus.Item("2055").Activate()  'AR Credit Memo
                            oForm = objaddon.objapplication.Forms.ActiveForm
                            oForm.Freeze(True)
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            oForm.Items.Item("8").Specific.String = Matrix0.Columns.Item("docnum").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("10").Specific.String = Matrix0.Columns.Item("date").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oForm.Freeze(False)
                            BubbleEvent = False
                        ElseIf Matrix0.Columns.Item("object").Cells.Item(pVal.Row).Specific.String = "18" Then
                            objaddon.objapplication.Menus.Item("2308").Activate()  'AP Invoice
                            oForm = objaddon.objapplication.Forms.ActiveForm
                            oForm.Freeze(True)
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            oForm.Items.Item("8").Specific.String = Matrix0.Columns.Item("docnum").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("10").Specific.String = Matrix0.Columns.Item("date").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oForm.Freeze(False)
                            BubbleEvent = False
                        ElseIf Matrix0.Columns.Item("object").Cells.Item(pVal.Row).Specific.String = "19" Then
                            objaddon.objapplication.Menus.Item("2309").Activate()  'AP Credit Memo
                            oForm = objaddon.objapplication.Forms.ActiveForm
                            oForm.Freeze(True)
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            oForm.Items.Item("8").Specific.String = Matrix0.Columns.Item("docnum").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("10").Specific.String = Matrix0.Columns.Item("date").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oForm.Freeze(False)
                            BubbleEvent = False
                        ElseIf Matrix0.Columns.Item("object").Cells.Item(pVal.Row).Specific.String = "203" Then
                            objaddon.objapplication.Menus.Item("2071").Activate()  'AR DownPayment
                            oForm = objaddon.objapplication.Forms.ActiveForm
                            oForm.Freeze(True)
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            oForm.Items.Item("8").Specific.String = Matrix0.Columns.Item("docnum").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("10").Specific.String = Matrix0.Columns.Item("date").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oForm.Freeze(False)
                            BubbleEvent = False
                        ElseIf Matrix0.Columns.Item("object").Cells.Item(pVal.Row).Specific.String = "204" Then
                            objaddon.objapplication.Menus.Item("2317").Activate()  'AP DownPayment
                            oForm = objaddon.objapplication.Forms.ActiveForm
                            oForm.Freeze(True)
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            oForm.Items.Item("8").Specific.String = Matrix0.Columns.Item("docnum").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("10").Specific.String = Matrix0.Columns.Item("date").Cells.Item(pVal.Row).Specific.String
                            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oForm.Freeze(False)
                            BubbleEvent = False
                        Else 'If Matrix0.Columns.Item("object").Cells.Item(pVal.Row).Specific.String = "30" Then
                            objlink.LinkedObjectType = "30" ' Matrix0.Columns.Item("object").Cells.Item(pVal.Row).Specific.String
                            objlink.Item.LinkTo = "docnum"
                        End If
                    Case "distrule"
                        Try
                            If Matrix0.Columns.Item("distrule").Cells.Item(pVal.Row).Specific.String <> "" Then
                                Link_Value = Matrix0.Columns.Item("distrule").Cells.Item(pVal.Row).Specific.String
                                oForm = New FrmDistRule
                                oForm.Show()
                            End If
                        Catch ex As Exception
                        End Try
                End Select

            Catch ex As Exception
                oForm.Freeze(False)
                oForm = Nothing
            End Try
        End Sub

        Private Sub Matrix0_PressedAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.PressedAfter
            Try
                If pVal.Row <= 0 Or objform.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then Exit Sub
                objCheck = Matrix0.Columns.Item("select").Cells.Item(pVal.Row).Specific
                If pVal.ColUID = "select" Then
                    If objCheck.Checked = True Then
                        'Matrix0.SelectRow(pVal.Row, True, True)
                        Matrix0.CommonSetting.SetRowBackColor(pVal.Row, Color.PeachPuff.ToArgb) 'BlanchedAlmond Wheat Tan SandyBrown PaleGoldenrod
                    Else
                        Matrix0.CommonSetting.SetRowBackColor(pVal.Row, Matrix0.Item.BackColor)
                        'Matrix0.SelectRow(pVal.Row, False, True)
                        Matrix0.SetCellWithoutValidation(pVal.Row, "paytot", CDbl(Matrix0.Columns.Item("baldue").Cells.Item(pVal.Row).Specific.String))
                    End If
                    'Calculate_Total()
                    Calc_Total(pVal.Row)
                    Matrix_DataTable(pVal.Row, "")
                    Clear_Payments()
                    'If Matrix0.Columns.Item("select").Cells.Item(pVal.Row).Specific.Checked = True Then
                    '    Total += CDbl(Matrix0.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String)
                    'Else
                    '    Total -= CDbl(Matrix0.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String)
                    'End If
                    'odbdsHeader.SetValue("U_Total", 0, value) 'Total
                End If
            Catch ex As Exception
            End Try

        End Sub

        Private Sub Matrix0_ValidateBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Matrix0.ValidateBefore
            Try
                'If pVal.ItemChanged = False Then Exit Sub
                If pVal.InnerEvent = True Then Exit Sub
                Dim Balance, PayTotal As Double
                Dim Disc, ActTotal As Double
                If Val(Matrix0.Columns.Item("baldue").Cells.Item(pVal.Row).Specific.String) <> 0 Then Balance = CDbl(Matrix0.Columns.Item("baldue").Cells.Item(pVal.Row).Specific.String) Else Balance = 0
                If Val(Matrix0.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String) <> 0 Then PayTotal = CDbl(Matrix0.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String) Else PayTotal = 0
                If Val(Matrix0.Columns.Item("pay").Cells.Item(pVal.Row).Specific.String) <> 0 Then ActTotal = CDbl(Matrix0.Columns.Item("pay").Cells.Item(pVal.Row).Specific.String) Else ActTotal = 0
                Select Case pVal.ColUID
                    Case "paytot"
                        If pVal.InnerEvent = False Then
                            If PayTotal <= 0 Then
                                PayTotal = -CDbl(Matrix0.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String)
                                Balance = -CDbl(Matrix0.Columns.Item("baldue").Cells.Item(pVal.Row).Specific.String)
                                If PayTotal > Balance Or PayTotal = 0 Then
                                    'Matrix0.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String = CDbl(Matrix0.Columns.Item("baldue").Cells.Item(pVal.Row).Specific.String)
                                    Matrix0.SetCellWithoutValidation(pVal.Row, "paytot", CDbl(Matrix0.Columns.Item("baldue").Cells.Item(pVal.Row).Specific.String))
                                End If
                            ElseIf PayTotal > 0 Then
                                If PayTotal > Balance Then
                                    'Matrix0.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String = CDbl(Matrix0.Columns.Item("baldue").Cells.Item(pVal.Row).Specific.String)
                                    Matrix0.SetCellWithoutValidation(pVal.Row, "paytot", CDbl(Matrix0.Columns.Item("baldue").Cells.Item(pVal.Row).Specific.String))
                                    'objaddon.objapplication.StatusBar.SetText("Total Amount is greater than the balance amount...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    'Matrix0.Columns.Item("paytot").Cells.Item(pVal.Row).Click(): BubbleEvent = False
                                End If
                            Else
                                'Matrix0.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String = CDbl(Matrix0.Columns.Item("baldue").Cells.Item(pVal.Row).Specific.String)
                                Matrix0.SetCellWithoutValidation(pVal.Row, "paytot", CDbl(Matrix0.Columns.Item("baldue").Cells.Item(pVal.Row).Specific.String))
                            End If
                            objCheck = Matrix0.Columns.Item("select").Cells.Item(pVal.Row).Specific
                            If pVal.ItemChanged = True And objCheck.Checked = False Then
                                objCheck.Checked = True
                                objform.Update()
                                'Matrix0.SelectRow(pVal.Row, True, True)
                                Matrix0.CommonSetting.SetRowBackColor(pVal.Row, Color.PeachPuff.ToArgb)
                                Matrix0.SetCellWithoutValidation(pVal.Row, "paytot", CDbl(Matrix0.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String))
                            End If
                            Matrix0.Columns.Item("pay").Cells.Item(pVal.Row).Specific.String = Matrix0.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String
                            If pVal.ItemChanged = True Then
                                Calc_Total(pVal.Row)
                            End If
                            'Calculate_Total()
                            'Clear_Payments()
                        End If
                    Case "cashdisc"
                        If Val(Matrix0.Columns.Item("cashdisc").Cells.Item(pVal.Row).Specific.String) <> 0 Then
                            PayTotal = CDbl(Matrix0.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String)
                            Disc = CDbl(Matrix0.Columns.Item("cashdisc").Cells.Item(pVal.Row).Specific.String)
                            If PayTotal < 0 Then
                                If Disc > 0 Then
                                    objaddon.objapplication.StatusBar.SetText("In Discount field, enter a valid number...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Matrix0.Columns.Item("cashdisc").Cells.Item(pVal.Row).Click() : BubbleEvent = False : Exit Sub
                                End If
                            ElseIf PayTotal > 0 Then
                                If Disc < 0 Then
                                    objaddon.objapplication.StatusBar.SetText("In Discount Field, enter a valid number...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Matrix0.Columns.Item("cashdisc").Cells.Item(pVal.Row).Click() : BubbleEvent = False : Exit Sub
                                End If
                            End If
                            Disc = CDbl(Matrix0.Columns.Item("pay").Cells.Item(pVal.Row).Specific.String) - CDbl(Matrix0.Columns.Item("cashdisc").Cells.Item(pVal.Row).Specific.String) '(ActTotal * (CDbl(Matrix0.Columns.Item("cashdisc").Cells.Item(pVal.Row).Specific.String) / 100))
                            Matrix0.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String = CStr(Disc)
                        Else
                            Matrix0.Columns.Item("paytot").Cells.Item(pVal.Row).Specific.String = Matrix0.Columns.Item("pay").Cells.Item(pVal.Row).Specific.String
                        End If
                        'Calculate_Total()
                        Calc_Total(pVal.Row)
                End Select
                If pVal.ItemChanged = True Then
                    Matrix_DataTable(pVal.Row, pVal.ColUID)
                End If
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Matrix0_ChooseFromListBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Matrix0.ChooseFromListBefore
            Try
                If pVal.ActionSuccess = True Then Exit Sub
                If pVal.ColUID = "cc1" Then
                    Try
                        Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFL_0")
                        Dim oConds As SAPbouiCOM.Conditions
                        Dim oCond As SAPbouiCOM.Condition
                        Dim oEmptyConds As New SAPbouiCOM.Conditions
                        oCFL.SetConditions(oEmptyConds)
                        oConds = oCFL.GetConditions()

                        oCond = oConds.Add()
                        oCond.Alias = "DimCode"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = "1"
                        oCFL.SetConditions(oConds)
                    Catch ex As Exception
                        SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try
                ElseIf pVal.ColUID = "cc2" Then
                    Try
                        Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFL_1")
                        Dim oConds As SAPbouiCOM.Conditions
                        Dim oCond As SAPbouiCOM.Condition
                        Dim oEmptyConds As New SAPbouiCOM.Conditions
                        oCFL.SetConditions(oEmptyConds)
                        oConds = oCFL.GetConditions()

                        oCond = oConds.Add()
                        oCond.Alias = "DimCode"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = "2"
                        oCFL.SetConditions(oConds)
                    Catch ex As Exception
                        SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try
                ElseIf pVal.ColUID = "cc3" Then
                    Try
                        Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFL_2")
                        Dim oConds As SAPbouiCOM.Conditions
                        Dim oCond As SAPbouiCOM.Condition
                        Dim oEmptyConds As New SAPbouiCOM.Conditions
                        oCFL.SetConditions(oEmptyConds)
                        oConds = oCFL.GetConditions()

                        oCond = oConds.Add()
                        oCond.Alias = "DimCode"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = "3"
                        oCFL.SetConditions(oConds)
                    Catch ex As Exception
                        SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try
                ElseIf pVal.ColUID = "cc4" Then
                    Try
                        Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFL_3")
                        Dim oConds As SAPbouiCOM.Conditions
                        Dim oCond As SAPbouiCOM.Condition
                        Dim oEmptyConds As New SAPbouiCOM.Conditions
                        oCFL.SetConditions(oEmptyConds)
                        oConds = oCFL.GetConditions()

                        oCond = oConds.Add()
                        oCond.Alias = "DimCode"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = "4"
                        oCFL.SetConditions(oConds)
                    Catch ex As Exception
                        SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try
                ElseIf pVal.ColUID = "cc5" Then
                    Try
                        Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFL_4")
                        Dim oConds As SAPbouiCOM.Conditions
                        Dim oCond As SAPbouiCOM.Condition
                        Dim oEmptyConds As New SAPbouiCOM.Conditions
                        oCFL.SetConditions(oEmptyConds)
                        oConds = oCFL.GetConditions()

                        oCond = oConds.Add()
                        oCond.Alias = "DimCode"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = "5"
                        oCFL.SetConditions(oConds)
                    Catch ex As Exception
                        SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try
                    'ElseIf pVal.ColUID = "distrule" Then
                    '    If CostCenter = "U" Then
                    '        If Not objaddon.FormExist("DistRule") Then
                    '            Dim oform As New FrmDistRule
                    '            oform.Show()
                    '        End If
                    '    End If
                    '    BubbleEvent = False
                End If

            Catch ex As Exception
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End Sub

        Private Sub Matrix0_ChooseFromListAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.ChooseFromListAfter
            Try
                If pVal.ActionSuccess = False Then Exit Sub
                If pVal.ColUID = "cc1" And pVal.ActionSuccess = True Then
                    Try
                        pCFL = pVal
                        If Not pCFL.SelectedObjects Is Nothing Then
                            Try
                                Matrix0.Columns.Item("cc1").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("OcrCode").Cells.Item(0).Value
                            Catch ex As Exception
                                Matrix0.Columns.Item("cc1").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("OcrCode").Cells.Item(0).Value
                            End Try
                        End If
                    Catch ex As Exception
                    End Try
                ElseIf pVal.ColUID = "cc2" And pVal.ActionSuccess = True Then
                    Try
                        pCFL = pVal
                        If Not pCFL.SelectedObjects Is Nothing Then
                            Try
                                Matrix0.Columns.Item("cc2").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("OcrCode").Cells.Item(0).Value
                            Catch ex As Exception
                                Matrix0.Columns.Item("cc2").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("OcrCode").Cells.Item(0).Value
                            End Try
                        End If
                    Catch ex As Exception
                    End Try
                ElseIf pVal.ColUID = "cc3" And pVal.ActionSuccess = True Then
                    Try
                        pCFL = pVal
                        If Not pCFL.SelectedObjects Is Nothing Then
                            Try
                                Matrix0.Columns.Item("cc3").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("OcrCode").Cells.Item(0).Value
                            Catch ex As Exception
                                Matrix0.Columns.Item("cc3").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("OcrCode").Cells.Item(0).Value
                            End Try
                        End If
                    Catch ex As Exception
                    End Try
                ElseIf pVal.ColUID = "cc4" And pVal.ActionSuccess = True Then
                    Try
                        pCFL = pVal
                        If Not pCFL.SelectedObjects Is Nothing Then
                            Try
                                Matrix0.Columns.Item("cc4").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("OcrCode").Cells.Item(0).Value
                            Catch ex As Exception
                                Matrix0.Columns.Item("cc4").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("OcrCode").Cells.Item(0).Value
                            End Try
                        End If
                    Catch ex As Exception
                    End Try
                ElseIf pVal.ColUID = "cc5" And pVal.ActionSuccess = True Then
                    Try
                        pCFL = pVal
                        If Not pCFL.SelectedObjects Is Nothing Then
                            Try
                                Matrix0.Columns.Item("cc5").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("OcrCode").Cells.Item(0).Value
                            Catch ex As Exception
                                Matrix0.Columns.Item("cc5").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("OcrCode").Cells.Item(0).Value
                            End Try
                        End If
                    Catch ex As Exception
                    End Try
                End If
                Matrix0.AutoResizeColumns()
                Matrix_DataTable(pVal.Row, pVal.ColUID)
                'objaddon.objapplication.Menus.Item("1300").Activate()
                GC.Collect()
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End Try

        End Sub

        Private Sub Matrix0_KeyDownAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.KeyDownAfter
            Try
                Dim ColID As Integer = Matrix0.GetCellFocus().ColumnIndex
                If pVal.CharPressed = 38 And pVal.Modifiers = SAPbouiCOM.BoModifiersEnum.mt_None Then  'up
                    Matrix0.SetCellFocus(pVal.Row - 1, ColID)
                    Matrix0.CommonSetting.SetRowBackColor(pVal.Row, Color.PeachPuff.ToArgb)
                    'Matrix0.SelectRow(pVal.Row - 1, True, False)
                ElseIf pVal.CharPressed = 40 And pVal.Modifiers = SAPbouiCOM.BoModifiersEnum.mt_None Then 'down
                    Matrix0.SetCellFocus(pVal.Row + 1, ColID)
                    Matrix0.CommonSetting.SetRowBackColor(pVal.Row, Color.PeachPuff.ToArgb)
                    'Matrix0.SelectRow(pVal.Row + 1, True, False)
                End If
                Select Case pVal.ColUID
                    Case "distrule"
                        If pVal.CharPressed = 13 Then
                            If Matrix0.Columns.Item("distrule").Cells.Item(pVal.Row).Specific.String = "" Then
                                If CostCenter = "U" Then
                                    OEForm = objaddon.objapplication.Forms.ActiveForm
                                    Dim oform As New FrmDistRule
                                    oform.Show()
                                End If
                            End If
                        End If

                End Select
            Catch ex As Exception

            End Try

        End Sub

#End Region

#Region "Form Events"

        Private Sub Button0_ClickBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Button0.ClickBefore
            Try
                If Button0.Item.Enabled = False Then Exit Sub
                If objform.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then Exit Sub
                If pVal.InnerEvent = True Then BubbleEvent = False : Exit Sub
                If CDbl(EditText3.Value) = 0 Then objaddon.objapplication.StatusBar.SetText("Confirmation amount must be greater than 0...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
                Dim Line As Integer = 0
                Dim Amt, GetTot As Double
                Dim ErrorFlag As Boolean = False
                RemoveLastrow(Matrix2, "cardgl")
                RemoveLastrow(Matrix1, "chnum")
                'objFinalDT.Clear()
                If objActualDT.Rows.Count > 0 Then objFinalDT = objActualDT Else objFinalDT = build_Matrix_DataTable("paytot")
                'If objFinalDT.Rows.Count > 50 Then objaddon.objapplication.StatusBar.SetText("Maximum Rows Selected...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub

                Try
                    'Order By drg.Max(Function(dr) dr.Field(Of Double)("paytot")) Descending
                    Dim GetPayDT = From dr In objFinalDT.AsEnumerable()
                                   Where dr.Field(Of String)("object") = "13" And dr.Field(Of Double)("paytot") > 0
                                   Group dr By Ph = New With {Key .branch = dr.Field(Of String)("branchc"), Key .DTLine = dr.Field(Of String)("#"), Key .BPCode = dr.Field(Of String)("cardc")} Into drg = Group
                                   Order By drg.Max(Function(dr) dr.Field(Of Double)("paytot")) Descending
                                   Select New With {
                    .branch = Ph.branch,
                    .line = Ph.DTLine,
                    .bpcode = Ph.BPCode,
                    .LengthSum = drg.Max(Function(dr) dr.Field(Of Double)("paytot"))
                    }
                    For Each RowID In GetPayDT
                        Line = CInt(RowID.line.ToString())
                        CurBranch = RowID.branch.ToString()
                        CurBPCode = RowID.bpcode.ToString()
                        If Line <> 0 Then Exit For
                    Next
                Catch ex As Exception
                End Try
                If Line = "0" Then objaddon.objapplication.StatusBar.SetText("Seems No Invoice transactions selected...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
                If Val(EditText5.Value) = 0 And Val(EditText9.Value) = 0 And Val(EditText11.Value) = 0 And Val(EditText12.Value) = 0 Then objaddon.objapplication.StatusBar.SetText("Please update the payment means...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
                ' GetTot = Val(EditText5.Value) + Val(EditText9.Value) + Val(EditText11.Value) + Val(EditText12.Value)
                'If CDbl(EditText3.Value) <> GetTot Then objaddon.objapplication.StatusBar.SetText("Found due amount. Please update payment means...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
                If EditText18.Value = MainCurr Then
                    GetTot = CDbl(EditText5.Value) + CDbl(EditText9.Value) + CDbl(EditText11.Value) + CDbl(EditText12.Value)
                    If CDbl(EditText3.Value) <> GetTot Then objaddon.objapplication.StatusBar.SetText("Found due amount. Please update payment means...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
                End If
                DocumentDate = Date.ParseExact(EditText0.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                If objaddon.objglobalmethods.Get_Branch_Assigned_Series("24", PayInitDate.ToString("yyyyMMdd")) = False Then
                    If EditText29.Value = "" Then objaddon.objapplication.StatusBar.SetText("Please Select the Series for posting the Incoming Payment...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub

                End If

                objaddon.objapplication.StatusBar.SetText("Creating payment.Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                If objaddon.objcompany.InTransaction = False Then objaddon.objcompany.StartTransaction()
                If IncomingPayment(objFinalDT, Line, CurBranch, CurBPCode) = False Then
                    ErrorFlag = True
                    If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    objaddon.objapplication.StatusBar.SetText("Error while creating payment...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
                End If

                Try
                    Dim Branch As String
                    '        Dim curbranchDT = From dr In objFinalDT.AsEnumerable()
                    '                          Where dr.Field(Of String)("branchc") = CurBranch And dr.Field(Of String)("cardc") = CurBPCode
                    '                          Group dr By Ph = dr.Field(Of String)("branchc") Into drg = Group
                    '                          Select New With {
                    '.branch = Ph,
                    '.LengthSum = drg.Sum(Function(dr) dr.Field(Of String)("paytot"))
                    '}
                    'For Each RowID In curbranchDT
                    '    Amt = RowID.LengthSum.ToString()
                    '    If CDbl(EditText3.Value) - CDbl(Amt) = 0 Then
                    '    Else
                    '        SameBranchReconciliation(objFinalDT)
                    '    End If
                    'Next
                    Dim BranchTotal, CurTotal, ExcRate, DiffTotal, Forex As Double
                    'If EditText18.Value <> MainCurr Then
                    '    objRs.DoQuery("select T1.""DebCred"",T1.""Line_ID"",CASE WHEN T1.""BalDueCred""<>0  THEN  T1.""BalDueCred"" ELSE T1.""BalDueDeb"" END AS ""Balance"" from JDT1 T1 where T1.""ShortName"" in (Select ""CardCode"" from OCRD) and T1.""TransId""=(Select ""TransId"" from ORCT where ""DocEntry""=" & EditText13.Value & ") ")
                    '    If CDbl(objRs.Fields.Item(2).Value.ToString) <> 0 Then
                    '        Forex = IIf(Val(EditText26.Value) = 0, 0, CDbl(EditText26.Value)) ' CDbl(EditText24.Value) * CDbl(EditText21.Value)
                    '        For i As Integer = 0 To objFinalDT.Rows.Count - 1
                    '            If objFinalDT.Rows(i)("branchc").ToString = CurBranch And objFinalDT.Rows(i)("cardc").ToString <> CurBPCode Then ' And InDT.Rows(i)("cardc").ToString <> CardCode
                    '                If objFinalDT.Rows(i)("doccur").ToString = MainCurr Then
                    '                    BranchTotal = Math.Round(BranchTotal + CDbl(objFinalDT.Rows(i)("paytot").ToString), SumRound)
                    '                    CurTotal = Math.Round(CurTotal + CDbl(objFinalDT.Rows(i)("paytot").ToString), SumRound)
                    '                Else
                    '                    ExcRate = Math.Round(GetTransaction_ExchangeRate(objFinalDT.Rows(i)("object").ToString, objFinalDT.Rows(i)("transid").ToString), RateRound)
                    '                    BranchTotal = Math.Round(BranchTotal + (CDbl(objFinalDT.Rows(i)("paytot").ToString) * ExcRate), SumRound)
                    '                    CurTotal = Math.Round(CurTotal + (CDbl(objFinalDT.Rows(i)("paytot").ToString) * CDbl(EditText21.Value)), SumRound)
                    '                End If
                    '            End If
                    '        Next
                    '        DiffTotal = Math.Round((BranchTotal - CurTotal), SumRound) 'Math.Round(IIf(BranchTotal > CurTotal, (BranchTotal - CurTotal), (CurTotal - BranchTotal)), SumRound)
                    '        Forex = Forex + DiffTotal '(BranchTotal - ActTotal)
                    '        EditText26.Value = Forex
                    '        If Forex <> 0 Then '- CDbl(EditText3.Value)
                    '            If EditText22.Value = "" Then
                    '                If Forex_JournalEntry(objFinalDT, Line, Forex, True) = False Then 'Forex - CDbl(EditText3.Value)
                    '                    ErrorFlag = True
                    '                End If
                    '            End If
                    '        End If
                    '    End If
                    'End If

                    For i As Integer = 0 To objFinalDT.Rows.Count - 1
                        If objFinalDT.Rows(i)("branchc").ToString = CurBranch And objFinalDT.Rows(i)("cardc").ToString = CurBPCode Then
                            If objFinalDT.Rows(i)("doccur").ToString = MainCurr Then
                                BranchTotal = Math.Round(BranchTotal + CDbl(objFinalDT.Rows(i)("paytot").ToString), SumRound)
                                CurTotal = Math.Round(CurTotal + CDbl(objFinalDT.Rows(i)("paytot").ToString), SumRound)
                            Else
                                strSQL = "Select ""Rate"",""Currency"" from ORTT where ""RateDate""= '" & DocumentDate.ToString("yyyyMMdd") & "' and ""Currency""='" & objFinalDT.Rows(i)("doccur").ToString & "' "
                                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                objRs.DoQuery(strSQL)
                                If EditText18.Value <> MainCurr And EditText18.Value = objFinalDT.Rows(i)("doccur").ToString Then
                                    CurTotal = Math.Round(CurTotal + (CDbl(objFinalDT.Rows(i)("paytot").ToString) * Math.Round(CDbl(EditText21.Value), RateRound)), SumRound)
                                Else
                                    CurTotal = Math.Round(CurTotal + (CDbl(objFinalDT.Rows(i)("paytot").ToString) * Math.Round(CDbl(objRs.Fields.Item(0).Value.ToString), RateRound)), SumRound)
                                End If
                                ExcRate = Math.Round(GetTransaction_ExchangeRate(objFinalDT.Rows(i)("object").ToString, objFinalDT.Rows(i)("transid").ToString), RateRound)
                                BranchTotal = Math.Round(BranchTotal + (CDbl(objFinalDT.Rows(i)("paytot").ToString) * ExcRate), SumRound)
                            End If
                        End If
                    Next
                    Forex = IIf(Val(EditText26.Value) = 0, 0, CDbl(EditText26.Value))
                    DiffTotal = Math.Round(CurTotal - BranchTotal, SumRound)
                    Forex = Math.Round(Forex + DiffTotal, SumRound)

                    If Forex <> 0 Then
                        If EditText22.Value = "" Then
                            If Forex_JournalEntry(objFinalDT, Line, Forex, True) = False Then 'Forex - CDbl(EditText3.Value)
                                ErrorFlag = True
                            End If
                        End If
                    End If

                    If CDbl(EditText3.Value) - (DiffTotal) = 0 Then
                    Else
                        If SameBranchReconciliation(objFinalDT) = False Then
                            ErrorFlag = True
                        End If
                    End If

                    'Dim CurBranchTotal As Double
                    'For i As Integer = 0 To objFinalDT.Rows.Count - 1
                    '    If objFinalDT.Rows(i)("branchc").ToString = CurBranch And objFinalDT.Rows(i)("cardc").ToString = CurBPCode Then 'And objFinalDT.Rows(i)("cardc").ToString = CurBPCode
                    '        If objFinalDT.Rows(i)("doccur").ToString = MainCurr Then
                    '            CurBranchTotal += Math.Round(CDbl(objFinalDT.Rows(i)("paytot").ToString), 6)
                    '        Else
                    '            strSQL = "Select ""Rate"",""Currency"" from ORTT where ""RateDate""= '" & Now.Date.ToString("yyyyMMdd") & "' and ""Currency""='" & objFinalDT.Rows(i)("doccur").ToString & "' "
                    '            objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '            objRs.DoQuery(strSQL)
                    '            CurBranchTotal += Math.Round(CDbl(objFinalDT.Rows(i)("paytot").ToString) * CDbl(objRs.Fields.Item(0).Value.ToString), 6)
                    '        End If
                    '    End If
                    'Next

                    'If CDbl(EditText3.Value) - CurBranchTotal = 0 Then
                    'Else
                    '    SameBranchReconciliation(objFinalDT)
                    'End If

                    '------------------------------------------------------------------------------
                    'Dim objExcRateDT As New DataTable
                    'If objExcRateDT.Columns.Count = 0 Then
                    '    objExcRateDT.Columns.Add("Rate", GetType(Double))
                    '    objExcRateDT.Columns.Add("doccur", GetType(String))
                    '    objExcRateDT.Columns.Add("date", GetType(String))
                    'End If
                    ''Dim DocDate As Date = Date.ParseExact(EditText0.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                    'strSQL = "Select ""Rate"",""Currency"",""RateDate"" from ORTT where ""RateDate""= '" & DocumentDate.ToString("yyyyMMdd") & "'"
                    'objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    'objRs.DoQuery(strSQL)
                    'If objRs.RecordCount = 0 Then
                    '    If objExcRateDT.Rows.Count = 0 Then
                    '        objExcRateDT.Rows.Add(1, MainCurr)
                    '    End If
                    'Else
                    '    For j = 0 To objRs.RecordCount - 1
                    '        If objExcRateDT.Rows.Count = 0 Then
                    '            objExcRateDT.Rows.Add(1, MainCurr)
                    '        End If
                    '        objExcRateDT.Rows.Add(objRs.Fields.Item(0).Value, objRs.Fields.Item(1).Value, objRs.Fields.Item(2).Value)
                    '        objRs.MoveNext()
                    '    Next
                    'End If

                    'Dim ERate = From dr In objExcRateDT.AsEnumerable()
                    '            Select New With {Key .Curr = dr.Field(Of String)("doccur"),
                    '                   Key .Rate = CDbl(dr.Field(Of Double)("Rate")), Key .date = dr.Field(Of String)("date")}

                    'Dim DT = From dr In objFinalDT.AsEnumerable()
                    '         Select New With {Key .Curr = dr.Field(Of String)("doccur"), Key .date = dr.Field(Of String)("date"), Key .Branch = dr.Field(Of String)("branchc"), Key .BPCode = dr.Field(Of String)("cardc"), Key .Tot = dr.Field(Of Double)("paytot")
                    '                   }

                    '                    Dim CurrentBranchDT = From f In DT Join m In ERate On f.Curr Equals m.Curr
                    '                                          Where f.Branch = CurBranch 'And f.BPCode = CurBPCode
                    '                                          Group New With {f, m} By ph = f.Branch Into drg = Group
                    '                                          Select New With {Key .branch = ph,
                    '    Key .Amount = drg.Sum(Function(x) If(x.f.Curr = MainCurr, CDbl(x.f.Tot), CDbl(x.f.Tot * x.m.Rate)))
                    '}

                    '                    For Each RowID In CurrentBranchDT
                    '                        Amt = CDbl(RowID.Amount)
                    '                        If CDbl(EditText3.Value) - CDbl(Amt) = 0 Then
                    '                        Else
                    '                            If SameBranchReconciliation(objFinalDT) = False Then
                    '                                ErrorFlag = True
                    '                            End If
                    '                        End If
                    '                    Next

                    '                    Dim otherBranchDT = From f In DT Join m In ERate On f.Curr Equals m.Curr
                    '                                        Where f.Branch <> CurBranch
                    '                                        Group New With {f, m} By ph = f.Branch Into drg = Group
                    '                                        Select New With {Key .branch = ph,
                    '    Key .Amount = drg.Sum(Function(x) If(x.f.Curr = MainCurr, CDbl(x.f.Tot), CDbl(x.f.Tot * x.m.Rate)))
                    '}

                    Dim otherBranchDT = From dr In objFinalDT.AsEnumerable()
                                        Where dr.Field(Of String)("branchc") <> CurBranch
                                        Group dr By Ph = dr.Field(Of String)("branchc") Into drg = Group
                                        Select New With {
                .branch = Ph,
                .LengthSum = drg.Sum(Function(dr) dr.Field(Of Double)("paytot"))
                }

                    For Each RowID In otherBranchDT
                        Branch = RowID.branch.ToString()
                        Amt = CDbl(RowID.LengthSum)
                        If JournalEntry_BranchWise(objFinalDT, Branch, CDbl(Amt)) = False Then
                            objaddon.objapplication.StatusBar.SetText("JournalEntry_BranchWise", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            ErrorFlag = True
                        End If
                    Next

                    '------------------------------------------------------------------------------------------------------

                    '        Dim otherBranchDT = From dr In objFinalDT.AsEnumerable()
                    '                            Where dr.Field(Of String)("branchc") <> CurBranch
                    '                            Group dr By Ph = dr.Field(Of String)("branchc") Into drg = Group
                    '                            Select New With {
                    '.branch = Ph,                                                           '.cardcode = Function(dr) dr.Field(Of String)("cardc"),
                    '.LengthSum = drg.Sum(Function(dr) dr.Field(Of Double)("paytot"))
                    '}
                    '        For Each RowID In otherBranchDT
                    '            Branch = RowID.branch.ToString()
                    '            Amt = CDbl(RowID.LengthSum)
                    '            JournalEntry(objFinalDT, Branch)
                    '            'JournalEntry_BranchWise(objFinalDT, Branch, CDbl(Amt))
                    '        Next

                    If Val(EditText19.Value) > 0 Then
                        If CDbl(EditText23.Value) - CDbl(EditText19.Value) > 0 Then
                            If EditText25.Value = "" Then
                                If Forex_JournalEntry(objFinalDT, Line, CDbl(EditText23.Value) - CDbl(EditText19.Value)) = False Then
                                    objaddon.objapplication.StatusBar.SetText("Forex_JournalEntry", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    ErrorFlag = True
                                End If
                            End If
                        End If
                    End If
                    If Disc_JournalEntry(objFinalDT) = False Then
                        objaddon.objapplication.StatusBar.SetText("Disc_JournalEntry", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        ErrorFlag = True
                    End If
                    If ErrorFlag = True Then
                        If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        EditText13.Value = ""
                        EditText14.Value = ""
                        EditText15.Value = ""
                        EditText13.Value = ""
                        EditText14.Value = ""
                        EditText15.Value = ""
                        EditText22.Value = ""
                        EditText25.Value = ""
                        Try
                            objform.Freeze(True)
                            Matrix0.FlushToDataSource()
                            For rowNum As Integer = 0 To odbdsDetails.Size - 1
                                odbdsDetails.SetValue("U_JENo", rowNum, "")
                                odbdsDetails.SetValue("U_RecoNo", rowNum, "")
                                odbdsDetails.SetValue("U_DiscJE", rowNum, "")
                                odbdsDetails.SetValue("U_DiscRecoNo", rowNum, "")
                            Next
                            Matrix0.LoadFromDataSource()
                        Catch ex As Exception
                        Finally
                            objform.Freeze(False)
                        End Try
                        objform.Refresh()
                        objaddon.objapplication.StatusBar.SetText("Error while creating payment transactions.Please check...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
                    Else
                        If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        objaddon.objapplication.StatusBar.SetText("Payment Created Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    End If
                Catch ex As Exception
                End Try

            Catch ex As Exception
            End Try

        End Sub

        Private Sub Button4_ClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button4.ClickAfter
            Try
                If CDbl(EditText3.Value) = 0 Then objaddon.objapplication.StatusBar.SetText("Please select the transaction...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : Exit Sub
                If CDbl(EditText3.Value) < 0 Then objaddon.objapplication.StatusBar.SetText("Please select the transaction greater than 0...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : Exit Sub
                If Not objaddon.FormExist("PAYM") Then
                    objform = objaddon.objapplication.Forms.ActiveForm
                    If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then objPayDT = objActualDT Else objPayDT = GetPaymentDT("paytot")
                    If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then Clear_Payments()
                    Dim activeform As New FrmPaymentMeans
                    activeform.Show()
                End If
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End Try
        End Sub

        Private Function GetPaymentDT(ByVal sKeyFieldID As String) As DataTable
            'Dim oForm As SAPbouiCOM.Form = Nothing
            Dim objcheckbox As SAPbouiCOM.CheckBox
            Try
                Dim oDT As New DataTable
                oDT.Columns.Add(Matrix0.Columns.Item("paytot").UniqueID)
                oDT.Columns.Add(Matrix0.Columns.Item("doccur").UniqueID)

                For iRow As Integer = 1 To Matrix0.VisualRowCount
                    objcheckbox = Matrix0.Columns.Item("select").Cells.Item(iRow).Specific
                    If objcheckbox.Checked = True Then
                        Dim oRow As DataRow = oDT.NewRow
                        oRow.Item(Matrix0.Columns.Item("paytot").UniqueID) = Matrix0.Columns.Item("paytot").Cells.Item(iRow).Specific.Value
                        oRow.Item(Matrix0.Columns.Item("doccur").UniqueID) = Matrix0.Columns.Item("doccur").Cells.Item(iRow).Specific.Value

                        If oRow(sKeyFieldID).ToString.Trim = 0 Then Continue For
                        oDT.Rows.Add(oRow)
                    End If
                Next


                Return oDT
            Catch ex As Exception
                Return Nothing
            End Try

        End Function

        Private Sub Form_CloseAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                'objform = objaddon.objapplication.Forms.GetForm("PAYINIT", Me.FormCount)
                'objform.Close()
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Form_ResizeAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                Matrix0.AutoResizeColumns()
            Catch ex As Exception
            End Try
        End Sub

        'Private Sub Button2_ClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button2.ClickAfter
        '    Try
        '        If InternReco(objFinalDT) Then 'PaymentReconciliation()
        '            For i As Integer = 1 To Matrix0.VisualRowCount
        '                Matrix0.DeleteRow(Matrix0.GetNextSelectedRow())
        '                objaddon.objmenuevent.DeleteRow(Matrix0, "@MI_RCT1")
        '            Next
        '        End If
        '    Catch ex As Exception
        '    End Try
        'End Sub

        Private Sub Form_DataLoadAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo)
            Try
                objform.EnableMenu("1282", False)
                If EditText18.Value = MainCurr Then
                Else
                    EditText20.Value = EditText24.Value
                End If
                DeleteRow()
                For i As Integer = 1 To Matrix0.VisualRowCount
                    objCheck = Matrix0.Columns.Item("select").Cells.Item(i).Specific
                    If objCheck.Checked = True Then
                        If objCheck.Checked = True Then
                            Matrix0.CommonSetting.SetRowBackColor(i, Color.PeachPuff.ToArgb)
                            'Matrix0.SelectRow(i, True, True)
                            'Else
                            '    Matrix0.SelectRow(i, False, True)
                        End If
                    End If
                Next
                objform.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                Matrix0.AutoResizeColumns()
            Catch ex As Exception

            End Try

        End Sub

#End Region

#Region "Functions"

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

        Private Sub DeleteRow()
            Try
                Dim Flag As Boolean = False
                Dim objSelect As SAPbouiCOM.CheckBox
                Matrix0.Columns.Item("select").TitleObject.Sortable = True
                Matrix0.Columns.Item("select").TitleObject.Click(SAPbouiCOM.BoCellClickType.ct_Double)
                For i As Integer = Matrix0.VisualRowCount To 1 Step -1
                    objSelect = Matrix0.Columns.Item("select").Cells.Item(i).Specific
                    If objSelect.Checked = False Then
                        Matrix0.DeleteRow(i)
                        odbdsDetails.RemoveRecord(i)
                        Flag = True
                    End If
                Next
                Matrix0.Columns.Item("select").TitleObject.Sortable = False
                If Flag = True Then
                    For i As Integer = 1 To Matrix0.VisualRowCount
                        objSelect = Matrix0.Columns.Item("select").Cells.Item(i).Specific
                        If objSelect.Checked = True Then
                            Matrix0.Columns.Item("#").Cells.Item(i).Specific.String = i
                        End If
                    Next
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

        Private Sub Field_Setup()
            Try
                StaticText24.Item.Visible = False
                EditText23.Item.Visible = False
                EditText27.Item.Visible = False
                StaticText32.Item.Visible = False
                Matrix0.Columns.Item("cc1").Visible = False
                Matrix0.Columns.Item("cc2").Visible = False
                Matrix0.Columns.Item("cc3").Visible = False
                Matrix0.Columns.Item("cc4").Visible = False
                Matrix0.Columns.Item("cc5").Visible = False
                Matrix0.Columns.Item("#").Visible = False
                Matrix0.Columns.Item("object").Visible = False
                Matrix0.Columns.Item("doctype").Visible = False
                Matrix0.Columns.Item("branchc").Visible = False
                Matrix0.Columns.Item("round").Visible = False
                Matrix0.Columns.Item("pay").Visible = False
                Matrix0.Columns.Item("docentry").Visible = False
                'Matrix0.Columns.Item("transid").Visible = False
                Matrix0.Columns.Item("tranline").Visible = False
                Matrix0.Columns.Item("debcred").Visible = False
                Matrix0.Columns.Item("totlc").Visible = False
                Matrix0.Columns.Item("balduelc").Visible = False
                Matrix0.CommonSetting.FixedColumnsCount = 2
                Matrix1.Columns.Item("chbranch").Visible = False
                Matrix1.Columns.Item("chact").Visible = False
                Matrix1.Columns.Item("chendor").Visible = False
                Matrix1.Columns.Item("chissue").Visible = False
                Matrix1.Columns.Item("chfiscal").Visible = False
            Catch ex As Exception

            End Try
        End Sub

        Private Function LoadData(ByVal Query As String) As Boolean
            Try
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRs.DoQuery(Query)
                Matrix0.Clear()
                odbdsDetails.Clear()
                If objRs.RecordCount > 0 Then
                    objaddon.objapplication.StatusBar.SetText("Loading data Please wait...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    objform.Freeze(True)
                    While Not objRs.EoF
                        Matrix0.AddRow()
                        Matrix0.GetLineData(Matrix0.VisualRowCount)
                        odbdsDetails.SetValue("LineId", 0, objRs.Fields.Item("LineId").Value.ToString)
                        odbdsDetails.SetValue("U_Select", 0, objRs.Fields.Item("Selected").Value.ToString)
                        odbdsDetails.SetValue("U_DebCred", 0, objRs.Fields.Item("DebCred").Value.ToString)
                        odbdsDetails.SetValue("U_DocNum", 0, objRs.Fields.Item("DocNum").Value.ToString)
                        odbdsDetails.SetValue("U_DocEntry", 0, objRs.Fields.Item("DocEntry").Value.ToString)
                        odbdsDetails.SetValue("U_ObjType", 0, objRs.Fields.Item("Origin").Value.ToString)
                        odbdsDetails.SetValue("U_DocType", 0, objRs.Fields.Item("DocType").Value.ToString)
                        odbdsDetails.SetValue("U_DocCur", 0, objRs.Fields.Item("DocCur").Value.ToString)
                        odbdsDetails.SetValue("U_CardCode", 0, objRs.Fields.Item("CardCode").Value.ToString)
                        odbdsDetails.SetValue("U_RefNo", 0, objRs.Fields.Item("NumAtCard").Value.ToString)
                        odbdsDetails.SetValue("U_CardName", 0, objRs.Fields.Item("CardName").Value.ToString)
                        odbdsDetails.SetValue("U_DocDate", 0, objRs.Fields.Item("DocDate").Value)
                        odbdsDetails.SetValue("U_TransId", 0, objRs.Fields.Item("TransId").Value.ToString)
                        odbdsDetails.SetValue("U_TLine", 0, objRs.Fields.Item("Line_ID").Value.ToString)
                        odbdsDetails.SetValue("U_DueDays", 0, objRs.Fields.Item("OverDueDays").Value.ToString)

                        odbdsDetails.SetValue("U_Total", 0, objRs.Fields.Item("DTotal").Value.ToString)
                        odbdsDetails.SetValue("U_BalDue", 0, objRs.Fields.Item("BalDue").Value.ToString)
                        odbdsDetails.SetValue("U_PayTotal", 0, objRs.Fields.Item("BalDue").Value.ToString)

                        'odbdsDetails.SetValue("U_TotalLC", 0, objRs.Fields.Item("TotalLC").Value.ToString)
                        'odbdsDetails.SetValue("U_BalDueLC", 0, objRs.Fields.Item("BalDueLC").Value.ToString)
                        'If objRs.Fields.Item("DocCur").Value.ToString <> MainCurr Then
                        '    strSQL = objaddon.objglobalmethods.getSingleValue("Select ""Rate"" from ORTT where ""RateDate""= '" & Now.Date.ToString("yyyyMMdd") & "' and ""Currency""='" & objRs.Fields.Item("DocCur").Value.ToString & "' ")
                        '    strSQL = IIf(strSQL = "", 1, strSQL)
                        'End If

                        'If CDbl(objRs.Fields.Item("DocTotalFC").Value.ToString) <> 0 Then
                        '    odbdsDetails.SetValue("U_Total", 0, CDbl(objRs.Fields.Item("DocTotalFC").Value.ToString) * CDbl(strSQL))
                        '    odbdsDetails.SetValue("U_BalDue", 0, CDbl(objRs.Fields.Item("BalanceFC").Value.ToString) * CDbl(strSQL))
                        '    odbdsDetails.SetValue("U_TotalFC", 0, objRs.Fields.Item("DocTotalFC").Value.ToString)
                        '    odbdsDetails.SetValue("U_BalDueFC", 0, objRs.Fields.Item("BalanceFC").Value.ToString)
                        'Else
                        '    odbdsDetails.SetValue("U_Total", 0, objRs.Fields.Item("DocTotal").Value.ToString)
                        '    odbdsDetails.SetValue("U_BalDue", 0, objRs.Fields.Item("Balance").Value.ToString)
                        '    odbdsDetails.SetValue("U_TotalFC", 0, vbEmpty)
                        '    odbdsDetails.SetValue("U_BalDueFC", 0, vbEmpty)
                        'End If

                        'If CDbl(objRs.Fields.Item("BalanceFC").Value.ToString) <> 0 Then
                        '    odbdsDetails.SetValue("U_PayTotal", 0, CDbl(objRs.Fields.Item("BalanceFC").Value.ToString) * CDbl(strSQL))
                        'Else
                        '    odbdsDetails.SetValue("U_PayTotal", 0, objRs.Fields.Item("Balance").Value.ToString)
                        'End If

                        odbdsDetails.SetValue("U_CashDisc", 0, 0)
                        odbdsDetails.SetValue("U_Round", 0, 0)

                        odbdsDetails.SetValue("U_BranchId", 0, objRs.Fields.Item("BPLId").Value.ToString)
                        odbdsDetails.SetValue("U_BranchNam", 0, objRs.Fields.Item("BPLName").Value.ToString)
                        odbdsDetails.SetValue("U_Object", 0, objRs.Fields.Item("ObjType").Value.ToString)
                        odbdsDetails.SetValue("U_Pay", 0, objRs.Fields.Item("BalanceLC").Value.ToString)
                        'objform.DataSources.UserDataSources.Item("UD_1").Value = Math.Round(Convert.ToDouble(objRs.Fields.Item("BalanceLC").Value), 6)
                        Matrix0.SetLineData(Matrix0.VisualRowCount)
                        objRs.MoveNext()
                    End While
                    Matrix0.AutoResizeColumns()
                    For i As Integer = 1 To Matrix0.VisualRowCount
                        If Matrix0.Columns.Item("objtype").Cells.Item(i).Specific.String = "JE" Then
                            Matrix0.CommonSetting.SetRowFontColor(i, Color.DarkRed.ToArgb) ' Color.OrangeRed.ToArgb
                            Matrix0.CommonSetting.SetCellEditable(i, 19, False)
                        End If
                        If CInt(Matrix0.Columns.Item("duedays").Cells.Item(i).Specific.String) > 0 Then
                            Matrix0.CommonSetting.SetCellFontColor(i, 14, Color.Blue.B)
                        End If
                    Next
                    objform.Freeze(False)
                    objaddon.objapplication.StatusBar.SetText("Data Loaded Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    'objaddon.objapplication.Menus.Item("1300").Activate()
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

        Private Function GetTransaction_ExchangeRate(ByVal ObjType As String, ByVal TransId As String) As Double
            Try
                Dim Rate As Double

                If ObjType = "13" Then
                    strSQL = objaddon.objglobalmethods.getSingleValue("Select T0.""DocRate"" from OINV T0 where T0.""TransId""=" & TransId & "")
                    Rate = CDbl(strSQL)
                ElseIf ObjType = "14" Then
                    strSQL = objaddon.objglobalmethods.getSingleValue("Select T0.""DocRate"" from ORIN T0 where T0.""TransId""=" & TransId & "")
                    Rate = CDbl(strSQL)
                ElseIf ObjType = "24" Then
                    strSQL = objaddon.objglobalmethods.getSingleValue("Select T0.""DocRate"" from ORCT T0 where T0.""TransId""=" & TransId & "")
                    Rate = CDbl(strSQL)
                ElseIf ObjType = "18" Then
                    strSQL = objaddon.objglobalmethods.getSingleValue("Select T0.""DocRate"" from OPCH T0 where T0.""TransId""=" & TransId & "")
                    Rate = CDbl(strSQL)
                ElseIf ObjType = "19" Then
                    strSQL = objaddon.objglobalmethods.getSingleValue("Select T0.""DocRate"" from ORPC T0 where T0.""TransId""=" & TransId & "")
                    Rate = CDbl(strSQL)
                ElseIf ObjType = "46" Then
                    strSQL = objaddon.objglobalmethods.getSingleValue("Select T0.""DocRate"" from OVPM T0 where T0.""TransId""=" & TransId & "")
                    Rate = CDbl(strSQL)
                ElseIf ObjType = "30" Then
                    'strSQL = objaddon.objglobalmethods.getSingleValue("Select To_Varchar(T0.""RefDate"",'yyyyMMdd') from OJDT T0 where T0.""TransId""=" & TransId & "")
                    strSQL = objaddon.objglobalmethods.getSingleValue("Select case when T0.""Credit""<>0 Then T0.""Credit""/T0.""FCCredit"" Else T0.""Debit""/T0.""FCDebit"" End as ""ExcRate"" from JDT1 T0 where T0.""TransId""=" & TransId & " and ""ShortName"" in (Select ""CardCode"" from OCRD)")
                    'strSQL = "Select ""Rate"",""Currency"" from ORTT where ""RateDate""= '" & strSQL & "' and ""Currency""='" & Currency & "' "
                    'objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    'objRs.DoQuery(strSQL)
                    'If objRs.RecordCount > 0 Then
                    '    Rate = CDbl(objRs.Fields.Item(0).Value) 'CDbl(strSQL)
                    'Else
                    '    Rate = 1
                    'End If
                    Rate = CDbl(strSQL)
                Else
                    Rate = 1
                End If

                Return Rate
            Catch ex As Exception
                Return 0
            End Try
        End Function

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
                Dim oRow As DataRow = objActualDT.NewRow
                If objcheckbox.Checked = True Then
                    If objActualDT.Rows.Count > 0 Then
                        For DTRow As Integer = 0 To objActualDT.Rows.Count - 1
                            If objActualDT.Rows(DTRow)("#").ToString = Matrix0.Columns.Item("#").Cells.Item(Row).Specific.Value Then
                                If ColName <> "" Then
                                    objActualDT.Rows(DTRow)(Matrix0.Columns.Item(ColName).UniqueID) = Matrix0.Columns.Item(ColName).Cells.Item(Row).Specific.Value
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
                            objActualDT.Rows.Add(oRow)
                        End If
                    Else
                        For iCol As Integer = 0 To Matrix0.Columns.Count - 1
                            If iCol <> 1 Then
                                oRow.Item(Matrix0.Columns.Item(iCol).UniqueID) = Matrix0.Columns.Item(iCol).Cells.Item(Row).Specific.Value
                            End If
                        Next
                        objActualDT.Rows.Add(oRow)
                    End If
                Else
                    For DTRow As Integer = 0 To objActualDT.Rows.Count - 1
                        If objActualDT.Rows(DTRow)("#").ToString = Matrix0.Columns.Item("#").Cells.Item(Row).Specific.Value Then
                            objActualDT.Rows(DTRow).Delete()
                            Exit For
                        End If
                    Next
                End If

                Return objActualDT
            Catch ex As Exception
                Return Nothing
            End Try
        End Function

        'Private Sub Calculate_Total_Old()
        '    Try
        '        Dim objDT As New DataTable
        '        Dim valueLC, valueFC As Double
        '        Dim LCFlag, FCFlag As Boolean
        '        objDT.Columns.Add("paytot", GetType(Double))
        '        objDT.Columns.Add("doccur", GetType(String))
        '        'objDT.Columns.Add("cashdisc", GetType(Double))
        '        Dim objcheckbox As SAPbouiCOM.CheckBox
        '        For iRow As Integer = 1 To Matrix0.VisualRowCount
        '            objcheckbox = Matrix0.Columns.Item("select").Cells.Item(iRow).Specific
        '            If objcheckbox.Checked = True Then
        '                Dim oRow As DataRow = objDT.NewRow
        '                oRow.Item(Matrix0.Columns.Item("paytot").UniqueID) = Matrix0.Columns.Item("paytot").Cells.Item(iRow).Specific.Value
        '                oRow.Item(Matrix0.Columns.Item("doccur").UniqueID) = Matrix0.Columns.Item("doccur").Cells.Item(iRow).Specific.Value
        '                objDT.Rows.Add(oRow)
        '            End If
        '        Next
        '        For i As Integer = 0 To objDT.Rows.Count - 1
        '            If objDT.Rows(i)("paytot").ToString <> "" Then
        '                If objDT.Rows(i)("doccur").ToString = MainCurr Then
        '                    valueLC += CDbl(objDT.Rows(i)("paytot").ToString)
        '                    LCFlag = True
        '                Else
        '                    strSQL = objaddon.objglobalmethods.getSingleValue("Select ""Rate"" from ORTT where ""RateDate""= '" & DocumentDate.ToString("yyyyMMdd") & "' and ""Currency""='" & objDT.Rows(i)("doccur").ToString & "' ")
        '                    strSQL = IIf(strSQL = "", 1, strSQL)
        '                    valueLC += CDbl(objDT.Rows(i)("paytot").ToString) * CDbl(strSQL)
        '                    valueFC += CDbl(objDT.Rows(i)("paytot").ToString)
        '                    FCFlag = True
        '                End If
        '            End If
        '        Next
        '        odbdsHeader.SetValue("U_Total", 0, valueLC) 'Total
        '        odbdsHeader.SetValue("U_ActTotal", 0, valueLC) 'Actual Total
        '        If LCFlag And FCFlag Then
        '            EditText20.Value = "****"
        '            LCFlag = False : FCFlag = False
        '        Else
        '            EditText20.Value = valueFC
        '        End If
        '    Catch ex As Exception

        '    End Try
        'End Sub

        Private Sub Calc_Total(ByVal Row As Integer)
            Try
                Dim valueLC, valueFC, ActualValue As Double
                Dim LCFlag, FCFlag, DataFlag As Boolean
                Dim objcheckbox As SAPbouiCOM.CheckBox
                Dim chkFCdiffcurr As String = ""
                Dim oRow As DataRow = oSelectedDT.NewRow
                If Row = -1 Then GoTo ExCalc
                objcheckbox = Matrix0.Columns.Item("select").Cells.Item(Row).Specific
                If objcheckbox.Checked = True Then
                    If oSelectedDT.Rows.Count > 0 Then
                        For DTRow As Integer = 0 To oSelectedDT.Rows.Count - 1
                            If oSelectedDT.Rows(DTRow)("#").ToString = Matrix0.Columns.Item("#").Cells.Item(Row).Specific.Value Then
                                oSelectedDT.Rows(DTRow)("paytot") = Matrix0.Columns.Item("paytot").Cells.Item(Row).Specific.Value
                                DataFlag = True
                                Exit For
                            End If
                        Next
                        If DataFlag = False Then
                            oRow.Item(Matrix0.Columns.Item("paytot").UniqueID) = Matrix0.Columns.Item("paytot").Cells.Item(Row).Specific.Value
                            oRow.Item(Matrix0.Columns.Item("doccur").UniqueID) = Matrix0.Columns.Item("doccur").Cells.Item(Row).Specific.Value
                            oRow.Item(Matrix0.Columns.Item("#").UniqueID) = Matrix0.Columns.Item("#").Cells.Item(Row).Specific.Value
                            oSelectedDT.Rows.Add(oRow)
                        End If
                    Else
                        oRow.Item(Matrix0.Columns.Item("paytot").UniqueID) = Matrix0.Columns.Item("paytot").Cells.Item(Row).Specific.Value
                        oRow.Item(Matrix0.Columns.Item("doccur").UniqueID) = Matrix0.Columns.Item("doccur").Cells.Item(Row).Specific.Value
                        oRow.Item(Matrix0.Columns.Item("#").UniqueID) = Matrix0.Columns.Item("#").Cells.Item(Row).Specific.Value
                        oSelectedDT.Rows.Add(oRow)
                    End If
                Else
                    For DTRow As Integer = 0 To oSelectedDT.Rows.Count - 1
                        If oSelectedDT.Rows(DTRow)("#").ToString = Matrix0.Columns.Item("#").Cells.Item(Row).Specific.Value Then
                            oSelectedDT.Rows(DTRow).Delete()
                            Exit For
                        End If
                    Next
                End If
ExCalc:
                DocumentDate = Date.ParseExact(EditText0.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                For i As Integer = 0 To oSelectedDT.Rows.Count - 1
                    If Val(oSelectedDT.Rows(i)("paytot").ToString) <> 0 Then
                        If oSelectedDT.Rows(i)("doccur").ToString = MainCurr Then
                            valueLC = Math.Round((valueLC + oSelectedDT.Rows(i)("paytot")), SumRound)
                            ActualValue = Math.Round((ActualValue + oSelectedDT.Rows(i)("paytot")), SumRound)
                            LCFlag = True
                        Else
                            strSQL = objaddon.objglobalmethods.getSingleValue("Select ""Rate"" from ORTT where ""RateDate""= '" & DocumentDate.ToString("yyyyMMdd") & "' and ""Currency""='" & oSelectedDT.Rows(i)("doccur").ToString & "' ")
                            strSQL = IIf(strSQL = "", 1, strSQL)
                            If chkFCdiffcurr = "" Then chkFCdiffcurr = oSelectedDT.Rows(i)("doccur").ToString
                            If oSelectedDT.Rows(i)("doccur").ToString <> chkFCdiffcurr Then
                                LCFlag = True
                            End If
                            valueLC = Math.Round(valueLC + CDbl(oSelectedDT.Rows(i)("paytot") * Math.Round(CDbl(strSQL), RateRound)), SumRound) ' Math.Round(((valueLC + oSelectedDT.Rows(i)("paytot")) * Math.Round(CDbl(strSQL), 6)), 6)
                            valueFC = Math.Round((valueFC + oSelectedDT.Rows(i)("paytot")), SumRound)
                            ActualValue = Math.Round((ActualValue + oSelectedDT.Rows(i)("paytot")), SumRound)
                            FCFlag = True
                        End If
                    End If
                Next
                odbdsHeader.SetValue("U_Total", 0, valueLC) 'Total
                odbdsHeader.SetValue("U_ActTotal", 0, valueLC) 'Actual Total
                odbdsHeader.SetValue("U_InTotal", 0, valueLC) 'ActualValue
                If LCFlag And FCFlag Then
                    EditText20.Value = "****"
                    LCFlag = False : FCFlag = False
                Else
                    EditText20.Value = valueFC
                End If
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Clear_Payments()
            Try
                If Val(EditText5.Value) > 0 Or Val(EditText9.Value) > 0 Or Val(EditText11.Value) > 0 Or Val(EditText12.Value) > 0 Then
                    EditText5.Value = "" 'Bank Transfer total
                    EditText9.Value = "" ' Cash Total
                    EditText12.Value = "" ' Credit Total
                    'EditText11.Value = "" ' Cheque Total
                    'EditText21.Value = ""
                    EditText24.Value = ""   ' Overall FC
                    EditText26.Value = ""   'Exchange Rate
                    'Matrix1.Clear()
                    'Matrix2.Clear()
                End If
            Catch ex As Exception

            End Try
        End Sub

        Private Function IncomingPayment(ByVal InDT As DataTable, ByVal LineNum As Integer, ByVal Branch As String, ByVal CardCode As String) As Boolean
            Try
                Dim objIncom As SAPbobsCOM.Payments
                Dim DocEntry, Series As String
                'Dim InvTotal, Payment_on_Acct, PaidTotal As Double
                If EditText13.Value = "" Then
                    objIncom = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
                    objIncom.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_IncomingPayments
                    objIncom.DocType = SAPbobsCOM.BoRcptTypes.rCustomer
                    objIncom.CardCode = CardCode ' Matrix0.Columns.Item("cardc").Cells.Item(LineNum).Specific.String
                    objIncom.BPLID = Branch ' Matrix0.Columns.Item("branchc").Cells.Item(LineNum).Specific.String
                    objIncom.Remarks = "In Payment add-on"
                    objIncom.UserFields.Fields.Item("U_PayInNo").Value = EditText2.Value
                    Dim DocDate As Date = Date.ParseExact(EditText0.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                    objIncom.DocDate = DocDate 'Matrix0.Columns.Item("date").Cells.Item(LineNum).Specific.String
                    'InvTotal = CDbl(Matrix0.Columns.Item("paytot").Cells.Item(LineNum).Specific.String)
                    'PaidTotal = CDbl(EditText3.Value)
                    'Payment_on_Acct = PaidTotal - InvTotal
                    objIncom.DocCurrency = EditText18.Value 'Matrix0.Columns.Item("doccur").Cells.Item(LineNum).Specific.String
                    If Val(EditText21.Value) > 0 Then objIncom.DocRate = CDbl(EditText21.Value)
                    'objIncom.LocalCurrency = BoYesNoEnum.tYES
                    If Localization = "IN" Then
                        If objaddon.HANA Then
                            Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 ""Series"" FROM NNM1 WHERE ""ObjectCode""='24' and ""Indicator""=(Select ""Indicator"" from OFPR where '" & DocDate.ToString("yyyyMMdd") & "' Between ""F_RefDate"" and ""T_RefDate"") " &
                                                                                      " and Ifnull(""Locked"",'')='N' and ""BPLId""='" & Branch & "' Order by ""Series"" desc")
                        Else
                            Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 Series FROM NNM1 WHERE ObjectCode='24' and Indicator=(Select Indicator from OFPR where '" & DocDate.ToString("yyyyMMdd") & "' Between F_RefDate and T_RefDate) " &
                                                                                      " and Isnull(Locked,'')='N' and BPLId='" & Branch & "' Order by Series desc")
                        End If
                    Else
                        If objaddon.HANA Then
                            Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 ""Series"" FROM NNM1 WHERE ""ObjectCode""='24' and ""Indicator""=(Select ""Indicator"" from OFPR where '" & DocDate.ToString("yyyyMMdd") & "' Between ""F_RefDate"" and ""T_RefDate"") " &
                                                                                      " and Ifnull(""Locked"",'')='N' and ""BPLId""='" & Branch & "' Order by ""Series"" desc")
                        Else
                            Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 Series FROM NNM1 WHERE ObjectCode='24' and Indicator=(Select Indicator from OFPR where '" & DocDate.ToString("yyyyMMdd") & "' Between F_RefDate and T_RefDate) " &
                                                                                      " and Isnull(Locked,'')='N' and BPLId='" & Branch & "' Order by Series desc")
                        End If
                    End If
                    If Series <> "" Then objIncom.Series = Series Else objIncom.Series = EditText29.Value

                    'If Payment_on_Acct <> 0 Then
                    '    objIncom.AccountPayments.AccountCode = objaddon.objglobalmethods.getSingleValue("Select ""DebPayAcct"" from OCRD where ""CardCode""='" & Matrix0.Columns.Item("cardc").Cells.Item(LineNum).Specific.String & "'")
                    '    objIncom.AccountPayments.SumPaid = Payment_on_Acct
                    '    objIncom.AccountPayments.Add()
                    'End If

                    If Matrix1.VisualRowCount > 0 Then
                        If Val(Matrix1.Columns.Item("chamt").Cells.Item(1).Specific.String) > 0 Then 'Cheque
                            objIncom.CheckAccount = EditText10.Value
                            For Row As Integer = 1 To Matrix1.VisualRowCount
                                If Val(Matrix1.Columns.Item("chamt").Cells.Item(Row).Specific.String) > 0 Then
                                    If Matrix1.Columns.Item("chnum").Cells.Item(Row).Specific.String <> "" Then objIncom.Checks.CheckNumber = Trim(Matrix1.Columns.Item("chnum").Cells.Item(Row).Specific.String)
                                    objIncom.Checks.CheckAccount = EditText10.Value ' Matrix1.Columns.Item("chgl").Cells.Item(Row).Specific.String
                                    objIncom.Checks.CheckSum = Math.Round(CDbl(Matrix1.Columns.Item("chamt").Cells.Item(Row).Specific.String), SumRound)
                                    Dim oedit As SAPbouiCOM.EditText
                                    oedit = Matrix1.Columns.Item("chdate").Cells.Item(Row).Specific
                                    Dim chDate As Date = Date.ParseExact(oedit.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                                    objIncom.Checks.DueDate = chDate ' Matrix1.Columns.Item("chdate").Cells.Item(Row).Specific.String
                                    If Matrix1.Columns.Item("chcty").Cells.Item(Row).Specific.String <> "" Then objIncom.Checks.CountryCode = Matrix1.Columns.Item("chcty").Cells.Item(Row).Specific.String
                                    If Matrix1.Columns.Item("chbank").Cells.Item(Row).Specific.String <> "" Then objIncom.Checks.BankCode = Matrix1.Columns.Item("chbank").Cells.Item(Row).Specific.String
                                    If Matrix1.Columns.Item("chbranch").Cells.Item(Row).Specific.String <> "" Then objIncom.Checks.Branch = Trim(Matrix1.Columns.Item("chbranch").Cells.Item(Row).Specific.String)
                                    If Matrix1.Columns.Item("chact").Cells.Item(Row).Specific.String <> "" Then objIncom.Checks.AccounttNum = Trim(Matrix1.Columns.Item("chact").Cells.Item(Row).Specific.String)
                                    If Matrix1.Columns.Item("chissue").Cells.Item(Row).Specific.String <> "" Then objIncom.Checks.OriginallyIssuedBy = Matrix1.Columns.Item("chissue").Cells.Item(Row).Specific.String
                                    If Matrix1.Columns.Item("chfiscal").Cells.Item(Row).Specific.String <> "" Then objIncom.Checks.FiscalID = Trim(Matrix1.Columns.Item("chfiscal").Cells.Item(Row).Specific.String)
                                    If Matrix1.Columns.Item("chendor").Cells.Item(Row).Specific.String = "Y" Then
                                        objIncom.Checks.Trnsfrable = BoYesNoEnum.tYES
                                    Else
                                        objIncom.Checks.Trnsfrable = BoYesNoEnum.tNO
                                    End If
                                    objIncom.Checks.Add()
                                End If
                            Next
                        End If
                    End If
                    If Val(EditText17.Value) > 0 Then 'BCG Amount
                        objIncom.BankAccount = EditText16.Value
                        objIncom.BankChargeAmount = Math.Round(CDbl(EditText17.Value), SumRound)
                    End If


                    If Val(EditText5.Value) > 0 Then 'Transfer
                        objIncom.TransferAccount = EditText6.Value
                        objIncom.TransferSum = Math.Round(CDbl(EditText5.Value), SumRound)
                        objIncom.TransferDate = Date.ParseExact(EditText4.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                        objIncom.TransferReference = EditText7.Value
                    End If
                    If Matrix2.VisualRowCount > 0 Then
                        'Dim DocDate As Date = Date.ParseExact(objEdit.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                        If Val(Matrix2.Columns.Item("amtdue").Cells.Item(1).Specific.String) > 0 Then  'Card
                            objIncom.CreditCards.CreditCard = Matrix2.Columns.Item("cardname").Cells.Item(1).Specific.String
                            objIncom.CreditCards.CreditCardNumber = Matrix2.Columns.Item("cardno").Cells.Item(1).Specific.String
                            objIncom.CreditCards.CreditAcct = Matrix2.Columns.Item("cardgl").Cells.Item(1).Specific.String
                            Dim GetDate As String = objaddon.objglobalmethods.getSingleValue("SELECT LAST_DAY (TO_DATE('" & Matrix2.Columns.Item("valid").Cells.Item(1).Specific.String & "', 'MM/YY')) ""last day"" FROM DUMMY")
                            objIncom.CreditCards.CardValidUntil = CDate(GetDate)
                            If Matrix2.Columns.Item("idno").Cells.Item(1).Specific.String <> "" Then objIncom.CreditCards.OwnerIdNum = Matrix2.Columns.Item("idno").Cells.Item(1).Specific.String
                            If Matrix2.Columns.Item("telno").Cells.Item(1).Specific.String <> "" Then objIncom.CreditCards.OwnerPhone = Matrix2.Columns.Item("telno").Cells.Item(1).Specific.String
                            objIncom.CreditCards.CreditSum = Math.Round(CDbl(Matrix2.Columns.Item("amtdue").Cells.Item(1).Specific.String), SumRound)
                            If Matrix2.Columns.Item("appcode").Cells.Item(1).Specific.String <> "" Then objIncom.CreditCards.VoucherNum = Matrix2.Columns.Item("appcode").Cells.Item(1).Specific.String
                            If Matrix2.Columns.Item("trantype").Cells.Item(1).Specific.String = "I" Then
                                objIncom.CreditCards.CreditType = BoRcptCredTypes.cr_InternetTransaction
                            ElseIf Matrix2.Columns.Item("trantype").Cells.Item(1).Specific.String = "S" Then
                                objIncom.CreditCards.CreditType = BoRcptCredTypes.cr_Regular
                            Else
                                objIncom.CreditCards.CreditType = BoRcptCredTypes.cr_Telephone
                            End If
                            objIncom.CreditCards.Add()
                        End If
                    End If
                    If Val(EditText9.Value) > 0 Then 'Cash
                        objIncom.CashAccount = EditText8.Value
                        objIncom.CashSum = Math.Round(CDbl(EditText9.Value), SumRound)
                    End If
                    For i As Integer = 0 To InDT.Rows.Count - 1
                        If CDbl(InDT.Rows(i)("paytot").ToString) <> 0 And InDT.Rows(i)("cardc").ToString = CurBPCode And InDT.Rows(i)("branchc").ToString = CurBranch Then
                            objIncom.Invoices.TotalDiscount = CDbl(InDT.Rows(i)("cashdisc").ToString)
                            If InDT.Rows(i)("doccur").ToString = MainCurr Then
                                objIncom.Invoices.SumApplied = Math.Round(CDbl(InDT.Rows(i)("paytot").ToString), SumRound) ' InvTotal
                            Else
                                'strSQL = "Select ""Rate"",""Currency"" from ORTT where ""RateDate""= '" & DocDate.ToString("yyyyMMdd") & "' and ""Currency""='" & EditText18.Value & "' "
                                'objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                'objRs.DoQuery(strSQL)
                                objIncom.Invoices.AppliedFC = Math.Round(CDbl(InDT.Rows(i)("paytot").ToString), SumRound) '/ CDbl(objRs.Fields.Item(0).Value.ToString)
                            End If
                            objIncom.Invoices.DocEntry = CInt(InDT.Rows(i)("docentry").ToString)
                            If InDT.Rows(i)("objtype").ToString = "IN" Then
                                objIncom.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_Invoice
                            ElseIf InDT.Rows(i)("objtype").ToString = "CN" Then
                                objIncom.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_CredItnote
                            ElseIf InDT.Rows(i)("objtype").ToString = "JE" Then
                                objIncom.Invoices.DocLine = CInt(InDT.Rows(i)("tranline").ToString)
                                objIncom.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_JournalEntry
                            End If

                            If InDT.Rows(i)("cc1").ToString <> "" Then objIncom.Invoices.DistributionRule = InDT.Rows(i)("cc1").ToString
                            If InDT.Rows(i)("cc2").ToString <> "" Then objIncom.Invoices.DistributionRule2 = InDT.Rows(i)("cc2").ToString
                            If InDT.Rows(i)("cc3").ToString <> "" Then objIncom.Invoices.DistributionRule3 = InDT.Rows(i)("cc3").ToString
                            If InDT.Rows(i)("cc4").ToString <> "" Then objIncom.Invoices.DistributionRule4 = InDT.Rows(i)("cc4").ToString
                            If InDT.Rows(i)("cc5").ToString <> "" Then objIncom.Invoices.DistributionRule5 = InDT.Rows(i)("cc5").ToString
                            objIncom.Invoices.Add()
                        End If
                    Next

#Region "Getfrom Matrix values"
                    'For i As Integer = 1 To Matrix0.VisualRowCount
                    '    If Matrix0.Columns.Item("select").Cells.Item(i).Specific.Checked = True And Matrix0.Columns.Item("cardc").Cells.Item(i).Specific.String = CurBPCode And Matrix0.Columns.Item("branchc").Cells.Item(i).Specific.String = CurBranch Then
                    '        objIncom.Invoices.TotalDiscount = CDbl(Matrix0.Columns.Item("cashdisc").Cells.Item(i).Specific.String)
                    '        objIncom.Invoices.SumApplied = CDbl(Matrix0.Columns.Item("paytot").Cells.Item(i).Specific.String) ' InvTotal
                    '        objIncom.Invoices.DocEntry = CInt(Matrix0.Columns.Item("docentry").Cells.Item(i).Specific.String)
                    '        If Matrix0.Columns.Item("objtype").Cells.Item(i).Specific.String = "IN" Then
                    '            objIncom.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_Invoice
                    '        ElseIf Matrix0.Columns.Item("objtype").Cells.Item(i).Specific.String = "CN" Then
                    '            objIncom.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_CredItnote
                    '        ElseIf Matrix0.Columns.Item("objtype").Cells.Item(i).Specific.String = "JE" Then
                    '            objIncom.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_JournalEntry
                    '        End If
                    '        If Matrix0.Columns.Item("cc1").Cells.Item(i).Specific.String <> "" Then objIncom.Invoices.DistributionRule = Matrix0.Columns.Item("cc1").Cells.Item(LineNum).Specific.String
                    '        If Matrix0.Columns.Item("cc2").Cells.Item(i).Specific.String <> "" Then objIncom.Invoices.DistributionRule2 = Matrix0.Columns.Item("cc2").Cells.Item(LineNum).Specific.String
                    '        If Matrix0.Columns.Item("cc3").Cells.Item(i).Specific.String <> "" Then objIncom.Invoices.DistributionRule3 = Matrix0.Columns.Item("cc3").Cells.Item(LineNum).Specific.String
                    '        If Matrix0.Columns.Item("cc4").Cells.Item(i).Specific.String <> "" Then objIncom.Invoices.DistributionRule4 = Matrix0.Columns.Item("cc4").Cells.Item(LineNum).Specific.String
                    '        If Matrix0.Columns.Item("cc5").Cells.Item(i).Specific.String <> "" Then objIncom.Invoices.DistributionRule5 = Matrix0.Columns.Item("cc5").Cells.Item(LineNum).Specific.String
                    '        objIncom.Invoices.Add()
                    '    End If
                    'Next
#End Region


                    Dim ret As Long
                    ret = objIncom.Add()

                    If ret <> 0 Then
                        'If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        objaddon.objapplication.StatusBar.SetText("Incoming Payment: Branch: " & GetBranchName(Branch) & " BPCode: " & CardCode & "-" & objaddon.objcompany.GetLastErrorDescription & "-" & objaddon.objcompany.GetLastErrorCode & " on Line: " & LineNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objIncom)
                        Return False
                    Else
                        'If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        DocEntry = objaddon.objcompany.GetNewObjectKey()
                        EditText13.Value = DocEntry
                        objaddon.objapplication.StatusBar.SetText("Incoming Payment successfully created..." & GetBranchName(Branch) & " BPCode: " & CardCode, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objIncom)
                        GC.Collect()
                        If JournalEntry_BranchTransfer(objFinalDT, CurBranch, CurBPCode) = False Then
                            objaddon.objapplication.StatusBar.SetText("JE_BranchTransfer", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            Return False
                        End If
                        Return True
                    End If
                Else
                    If JournalEntry_BranchTransfer(objFinalDT, CurBranch, CurBPCode) = False Then
                        objaddon.objapplication.StatusBar.SetText("JE_BranchTransfer", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        Return False
                    End If
                    Return True
                End If
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("Incoming Payment: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End Try
        End Function

        Private Function JournalEntry_BranchTransfer(ByVal InDT As DataTable, ByVal Branch As String, ByVal CardCode As String) As Boolean
            Try
                Dim TransId, GLCode, Series As String
                'Dim objrecset As SAPbobsCOM.Recordset
                Dim objjournalentry As SAPbobsCOM.JournalEntries
                Dim JEAmount As Double = 0
                Dim BranchTotal, CurTotal, ExcRate As Double
                If EditText14.Value <> "" Then Return True
                'Dim DocDate As Date = Date.ParseExact(EditText0.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                For i As Integer = 0 To InDT.Rows.Count - 1
                    If InDT.Rows(i)("branchc").ToString = Branch Then ' And InDT.Rows(i)("cardc").ToString <> CardCode
                        If InDT.Rows(i)("doccur").ToString = MainCurr Then
                            BranchTotal = Math.Round(BranchTotal + CDbl(InDT.Rows(i)("paytot").ToString), SumRound)
                            CurTotal = Math.Round(CurTotal + CDbl(InDT.Rows(i)("paytot").ToString), SumRound)
                        Else
                            strSQL = "Select ""Rate"",""Currency"" from ORTT where ""RateDate""= '" & DocumentDate.ToString("yyyyMMdd") & "' and ""Currency""='" & InDT.Rows(i)("doccur").ToString & "' " ' InDT.Rows(i)("date").ToString
                            objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            objRs.DoQuery(strSQL)
                            CurTotal = Math.Round(CurTotal + (CDbl(InDT.Rows(i)("paytot").ToString) * Math.Round(CDbl(objRs.Fields.Item(0).Value.ToString), RateRound)), SumRound)
                            'ExcRate = Math.Round(CDbl(objRs.Fields.Item(0).Value.ToString), RateRound)
                            ExcRate = Math.Round(GetTransaction_ExchangeRate(InDT.Rows(i)("object").ToString, InDT.Rows(i)("transid").ToString), RateRound)
                            BranchTotal = Math.Round(BranchTotal + (CDbl(InDT.Rows(i)("paytot").ToString) * ExcRate), SumRound)
                        End If
                    End If
                Next

                'ActTotal = Math.Round(IIf(CDbl(EditText3.Value) > CDbl(EditText27.Value), (CDbl(EditText3.Value) - CDbl(EditText27.Value)), (CDbl(EditText27.Value) - CDbl(EditText3.Value))), SumRound)
                'BranchTotal = Math.Round(BranchTotal - (CDbl(EditText3.Value) + ActTotal), SumRound)
                BranchTotal = Math.Round(BranchTotal - CDbl(EditText3.Value), SumRound)
                'BranchTotal = Math.Round(BranchTotal - (CDbl(EditText3.Value) + CDbl(EditText26.Value)), SumRound)
                'Dim Amount As String = objaddon.objglobalmethods.getSingleValue("select CASE WHEN T1.""BalDueCred""<>0  THEN  -T1.""BalDueCred"" ELSE T1.""BalDueDeb"" END AS ""Balance"" from JDT1 T1 where T1.""ShortName"" ='" & CurBPCode & "' and T1.""TransId""=(Select ""TransId"" from ORCT where ""DocEntry""=" & EditText13.Value & ") ")
                If BranchTotal = 0 Then Return True
                JEAmount = Math.Round(IIf(BranchTotal < 0, -BranchTotal, BranchTotal), SumRound)
                objjournalentry = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
                objaddon.objapplication.StatusBar.SetText("Journal Entry Creating Please wait..." & GetBranchName(Branch), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                'If Not objaddon.objcompany.InTransaction Then objaddon.objcompany.StartTransaction()
                Dim oEdit As SAPbouiCOM.EditText
                oEdit = objform.Items.Item("tdocdate").Specific
                Dim DocDate As Date = Date.ParseExact(oEdit.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                objjournalentry.ReferenceDate = DocDate
                'objjournalentry.DueDate = Now.Date.ToString("yyyyMMdd") 'DocDate
                objjournalentry.TaxDate = DocDate
                objjournalentry.Reference = "In Payment JE"
                objjournalentry.Memo = "Posted thro' Inpay On:" & Now.ToString
                objjournalentry.UserFields.Fields.Item("U_PayInNo").Value = EditText2.Value
                If Localization = "IN" Then
                    If objaddon.HANA Then
                        Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 ""Series"" FROM NNM1 WHERE ""ObjectCode""='30' and ""Indicator""=(Select ""Indicator"" from OFPR where '" & DocDate.ToString("yyyyMMdd") & "' Between ""F_RefDate"" and ""T_RefDate"") " &
                                                                                      " and Ifnull(""Locked"",'')='N' and ""BPLId""='" & Branch & "'")
                    Else
                        Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 Series FROM NNM1 WHERE ObjectCode='30' and Indicator=(Select Indicator from OFPR where '" & DocDate.ToString("yyyyMMdd") & "' Between F_RefDate and T_RefDate) " &
                                                                                      " and Isnull(Locked,'')='N' and BPLId='" & Branch & "'")
                    End If
                Else
                    If objaddon.HANA Then
                        Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 ""Series"" FROM NNM1 WHERE ""ObjectCode""='30' and ""Indicator""=(Select ""Indicator"" from OFPR where '" & DocDate.ToString("yyyyMMdd") & "' Between ""F_RefDate"" and ""T_RefDate"") " &
                                                                                      " and Ifnull(""Locked"",'')='N' and ""BPLId""='" & Branch & "'")
                    Else
                        Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 Series FROM NNM1 WHERE ObjectCode='30' and Indicator=(Select Indicator from OFPR where '" & DocDate.ToString("yyyyMMdd") & "' Between F_RefDate and T_RefDate) " &
                                                                                      " and Isnull(Locked,'')='N' and BPLId='" & Branch & "'")
                    End If
                End If
                If Series <> "" Then objjournalentry.Series = Series
                GLCode = objaddon.objglobalmethods.getSingleValue("select ""PmtClrAct"" ""ControlAccount"" from OBPL where ""BPLId""='" & Branch & "'")
                objjournalentry.Lines.AccountCode = GLCode
                If BranchTotal < 0 Then objjournalentry.Lines.Credit = JEAmount Else objjournalentry.Lines.Debit = JEAmount

                'If EditText18.Value = MainCurr Then
                '    If BranchTotal < 0 Then objjournalentry.Lines.Credit = JEAmount Else objjournalentry.Lines.Debit = JEAmount
                'Else
                '    objjournalentry.Lines.FCCurrency = EditText18.Value
                '    If BranchTotal < 0 Then objjournalentry.Lines.FCCredit = JEAmount Else objjournalentry.Lines.FCDebit = JEAmount
                'End If
                'objjournalentry.Lines.Debit = JEAmount
                objjournalentry.Lines.BPLID = Branch
                'If InDT.Rows(Row)("cc1").ToString <> "" Then objjournalentry.Lines.CostingCode = InDT.Rows(Row)("cc1").ToString
                'If InDT.Rows(Row)("cc2").ToString <> "" Then objjournalentry.Lines.CostingCode2 = InDT.Rows(Row)("cc2").ToString
                'If InDT.Rows(Row)("cc3").ToString <> "" Then objjournalentry.Lines.CostingCode3 = InDT.Rows(Row)("cc3").ToString
                'If InDT.Rows(Row)("cc4").ToString <> "" Then objjournalentry.Lines.CostingCode4 = InDT.Rows(Row)("cc4").ToString
                'If InDT.Rows(Row)("cc5").ToString <> "" Then objjournalentry.Lines.CostingCode5 = InDT.Rows(Row)("cc5").ToString
                objjournalentry.Lines.Add()
                objjournalentry.Lines.ShortName = CardCode
                If BranchTotal < 0 Then objjournalentry.Lines.Debit = JEAmount Else objjournalentry.Lines.Credit = JEAmount
                'If EditText18.Value = MainCurr Then
                '    If BranchTotal < 0 Then objjournalentry.Lines.Debit = JEAmount Else objjournalentry.Lines.Credit = JEAmount
                'Else
                '    objjournalentry.Lines.FCCurrency = EditText18.Value
                '    If BranchTotal < 0 Then objjournalentry.Lines.FCDebit = JEAmount Else objjournalentry.Lines.FCCredit = JEAmount
                'End If
                'objjournalentry.Lines.Credit = JEAmount
                objjournalentry.Lines.BPLID = Branch
                'If InDT.Rows(Row)("cc1").ToString <> "" Then objjournalentry.Lines.CostingCode = InDT.Rows(Row)("cc1").ToString
                'If InDT.Rows(Row)("cc2").ToString <> "" Then objjournalentry.Lines.CostingCode2 = InDT.Rows(Row)("cc2").ToString
                'If InDT.Rows(Row)("cc3").ToString <> "" Then objjournalentry.Lines.CostingCode3 = InDT.Rows(Row)("cc3").ToString
                'If InDT.Rows(Row)("cc4").ToString <> "" Then objjournalentry.Lines.CostingCode4 = InDT.Rows(Row)("cc4").ToString
                'If InDT.Rows(Row)("cc5").ToString <> "" Then objjournalentry.Lines.CostingCode5 = InDT.Rows(Row)("cc5").ToString
                objjournalentry.Lines.Add()

                If objjournalentry.Add <> 0 Then
                    'If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    objaddon.objapplication.SetStatusBarMessage("JE BranchTransfer: " & GetBranchName(Branch) & "-" & objaddon.objcompany.GetLastErrorDescription & "-" & objaddon.objcompany.GetLastErrorCode, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objjournalentry)
                    Return False
                Else
                    'If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    TransId = objaddon.objcompany.GetNewObjectKey()
                    EditText14.Value = TransId
                    objaddon.objapplication.SetStatusBarMessage("JE BranchTransfer Created Successfully..." & GetBranchName(Branch), SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objjournalentry)
                    Return True
                End If

            Catch ex As Exception
                'If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                objaddon.objapplication.SetStatusBarMessage("JE BranchTransfer Posting Error" & GetBranchName(Branch) & "-" & objaddon.objcompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                Return False
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End Function

        Private Function Forex_JournalEntry(ByVal InDT As DataTable, ByVal Row As Integer, ByVal Amount As Double, Optional ByVal RecFlag As Boolean = False) As Boolean
            Try
                Dim TransId, Series As String
                'Dim objrecset As SAPbobsCOM.Recordset
                Dim objjournalentry As SAPbobsCOM.JournalEntries
                Dim JEAmount As Double
                Try
                    If EditText25.Value = "" Or EditText22.Value = "" Then
                        For i As Integer = 0 To InDT.Rows.Count - 1
                            If InDT.Rows(i)("#").ToString = Row Then
                                Row = i
                                Exit For
                            End If
                        Next
                        objjournalentry = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
                        objaddon.objapplication.StatusBar.SetText("Forex Journal Entry Creating Please wait...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        'If Not objaddon.objcompany.InTransaction Then objaddon.objcompany.StartTransaction()
                        Dim oEdit As SAPbouiCOM.EditText
                        oEdit = objform.Items.Item("tdocdate").Specific
                        Dim DocDate As Date = Date.ParseExact(oEdit.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                        objjournalentry.ReferenceDate = DocDate
                        objjournalentry.TaxDate = DocDate
                        objjournalentry.Reference = "Forex Payment JE"
                        objjournalentry.Memo = "Forexpay On: " & Now.ToString
                        If Localization = "IN" Then
                            If objaddon.HANA Then
                                Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 ""Series"" FROM NNM1 WHERE ""ObjectCode""='30' and ""Indicator""=(Select ""Indicator"" from OFPR where '" & DocDate.ToString("yyyyMMdd") & "' Between ""F_RefDate"" and ""T_RefDate"") " &
                                                                                  " and Ifnull(""Locked"",'')='N' and ""BPLId""='" & InDT.Rows(Row)("branchc").ToString & "'")
                            Else
                                Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 Series FROM NNM1 WHERE ObjectCode='30' and Indicator=(Select Indicator from OFPR where '" & DocDate.ToString("yyyyMMdd") & "' Between F_RefDate and T_RefDate) " &
                                                                                  " and Isnull(Locked,'')='N' and BPLId='" & InDT.Rows(Row)("branchc").ToString & "'")
                            End If
                        Else
                            objjournalentry.AutoVAT = BoYesNoEnum.tNO
                            objjournalentry.AutomaticWT = BoYesNoEnum.tNO
                            If objaddon.HANA Then
                                Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 ""Series"" FROM NNM1 WHERE ""ObjectCode""='30' and ""Indicator""=(Select ""Indicator"" from OFPR where '" & DocDate.ToString("yyyyMMdd") & "' Between ""F_RefDate"" and ""T_RefDate"") " &
                                                                                  " and Ifnull(""Locked"",'')='N' and ""BPLId""='" & InDT.Rows(Row)("branchc").ToString & "'")
                            Else
                                Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 Series FROM NNM1 WHERE ObjectCode='30' and Indicator=(Select Indicator from OFPR where '" & DocDate.ToString("yyyyMMdd") & "' Between F_RefDate and T_RefDate) " &
                                                                                  " and Isnull(Locked,'')='N' and BPLId='" & InDT.Rows(Row)("branchc").ToString & "'")
                            End If
                        End If
                        If Series <> "" Then objjournalentry.Series = Series
                        JEAmount = IIf(Amount < 0, -Amount, Amount)
                        If RecFlag Then
                            If Amount < 0 Then objjournalentry.Lines.AccountCode = Forexgain Else objjournalentry.Lines.AccountCode = Forexloss
                        Else
                            objjournalentry.Lines.AccountCode = ForexDiff
                        End If
                        If Amount < 0 Then objjournalentry.Lines.Credit = JEAmount Else objjournalentry.Lines.Debit = JEAmount
                        objjournalentry.Lines.BPLID = InDT.Rows(Row)("branchc").ToString
                        If InDT.Rows(Row)("cc1").ToString <> "" Then objjournalentry.Lines.CostingCode = InDT.Rows(Row)("cc1").ToString
                        If InDT.Rows(Row)("cc2").ToString <> "" Then objjournalentry.Lines.CostingCode2 = InDT.Rows(Row)("cc2").ToString
                        If InDT.Rows(Row)("cc3").ToString <> "" Then objjournalentry.Lines.CostingCode3 = InDT.Rows(Row)("cc3").ToString
                        If InDT.Rows(Row)("cc4").ToString <> "" Then objjournalentry.Lines.CostingCode4 = InDT.Rows(Row)("cc4").ToString
                        If InDT.Rows(Row)("cc5").ToString <> "" Then objjournalentry.Lines.CostingCode5 = InDT.Rows(Row)("cc5").ToString
                        objjournalentry.Lines.Add()
                        objjournalentry.Lines.ShortName = InDT.Rows(Row)("cardc").ToString
                        If Amount < 0 Then objjournalentry.Lines.Debit = JEAmount Else objjournalentry.Lines.Credit = JEAmount

                        objjournalentry.Lines.BPLID = InDT.Rows(Row)("branchc").ToString
                        If InDT.Rows(Row)("cc1").ToString <> "" Then objjournalentry.Lines.CostingCode = InDT.Rows(Row)("cc1").ToString
                        If InDT.Rows(Row)("cc2").ToString <> "" Then objjournalentry.Lines.CostingCode2 = InDT.Rows(Row)("cc2").ToString
                        If InDT.Rows(Row)("cc3").ToString <> "" Then objjournalentry.Lines.CostingCode3 = InDT.Rows(Row)("cc3").ToString
                        If InDT.Rows(Row)("cc4").ToString <> "" Then objjournalentry.Lines.CostingCode4 = InDT.Rows(Row)("cc4").ToString
                        If InDT.Rows(Row)("cc5").ToString <> "" Then objjournalentry.Lines.CostingCode5 = InDT.Rows(Row)("cc5").ToString
                        objjournalentry.Lines.Add()
                        If objjournalentry.Add <> 0 Then
                            'If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            objaddon.objapplication.SetStatusBarMessage("Forex Journal: " & objaddon.objcompany.GetLastErrorDescription & "-" & objaddon.objcompany.GetLastErrorCode, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(objjournalentry)
                            Return False
                        Else
                            'If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                            TransId = objaddon.objcompany.GetNewObjectKey()
                            If RecFlag Then
                                EditText22.Value = TransId
                                'If EditText23.Value = "" Or Val(EditText23.Value) = 0 Then
                                '    Forex_JEInternReco(EditText14.Value, Amount, TransId)
                                'End If
                            Else
                                EditText25.Value = TransId
                            End If
                            objaddon.objapplication.SetStatusBarMessage("Forex Journal Entry Created Successfully...", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                            Return True
                        End If
                    Else
                        Return True
                        'Try
                        '    If RecFlag Then
                        '        If EditText23.Value = "" Or Val(EditText23.Value) = 0 Then
                        '            Forex_JEInternReco(EditText14.Value, Amount, EditText22.Value)
                        '        End If
                        '    End If
                        'Catch ex As Exception
                        '    objaddon.objapplication.SetStatusBarMessage("JE Rec " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                        'End Try
                    End If

                Catch ex As Exception
                    'If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    objaddon.objapplication.SetStatusBarMessage("JE Posting Error" & objaddon.objcompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                End Try

            Catch ex As Exception
                'If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                objaddon.objapplication.SetStatusBarMessage("JE " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                Return False
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End Function

        'Private Function JournalEntry(ByVal InDT As DataTable, ByVal Branch As String) As Boolean
        '    Try
        '        Dim TransId, GLCode, Series As String
        '        Dim objrecset As SAPbobsCOM.Recordset
        '        Dim objjournalentry As SAPbobsCOM.JournalEntries
        '        Dim JEAmount, ExRate As Double
        '        '    Dim branchDT = From dr In objFinalDT.AsEnumerable()
        '        '                   Where dr.Field(Of String)("branchc") <> CurBranch
        '        '                   Group dr By Ph = New With {Key .branch = dr.Field(Of String)("branchc"), Key .BPCode = dr.Field(Of String)("cardc")} Into drg = Group
        '        '                   Select New With {
        '        '.branch = Ph.branch,
        '        '.code = Ph.BPCode,
        '        '.LengthSum = drg.Sum(Function(dr) dr.Field(Of String)("paytot"))
        '        '} Distinct
        '        Try
        '            For Row As Integer = 0 To InDT.Rows.Count - 1
        '                If CDbl(InDT.Rows(Row)("paytot").ToString) <> 0 And InDT.Rows(Row)("branchc").ToString = Branch And InDT.Rows(Row)("jeno").ToString = String.Empty Then
        '                    objjournalentry = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
        '                    objaddon.objapplication.StatusBar.SetText("Journal Entry Creating Please wait...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        '                    If Not objaddon.objcompany.InTransaction Then objaddon.objcompany.StartTransaction()
        '                    Dim oEdit As SAPbouiCOM.EditText
        '                    oEdit = objform.Items.Item("tdocdate").Specific
        '                    Dim DocDate As Date = Date.ParseExact(oEdit.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
        '                    objjournalentry.ReferenceDate = DocDate 'ConvertDate.ToString("dd/MM/yy") 'DocDate 'Now.Date.ToString("yyyyMMdd") 
        '                    'objjournalentry.DueDate = Now.Date.ToString("yyyyMMdd") 'DocDate
        '                    objjournalentry.TaxDate = DocDate  ' ConvertDate.ToString("dd/MM/yy") 'DocDate 'Now.Date.ToString("yyyyMMdd") 
        '                    objjournalentry.Reference = "In Payment JE"
        '                    objjournalentry.Memo = "Posted thro' Inpay On: " & Now.ToString
        '                    objjournalentry.UserFields.Fields.Item("U_PayInNo").Value = EditText2.Value
        '                    If Localization = "IN" Then
        '                        If objaddon.HANA Then
        '                            Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 ""Series"" FROM NNM1 WHERE ""ObjectCode""='30' and ""Indicator""=(Select ""Indicator"" from OFPR where '" & DocDate.ToString("yyyyMMdd") & "' Between ""F_RefDate"" and ""T_RefDate"") " &
        '                                                                              " and Ifnull(""Locked"",'')='N' and ""BPLId""='" & InDT.Rows(Row)("branchc").ToString & "'")
        '                        Else
        '                            Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 Series FROM NNM1 WHERE ObjectCode='30' and Indicator=(Select Indicator from OFPR where '" & DocDate.ToString("yyyyMMdd") & "' Between F_RefDate and T_RefDate) " &
        '                                                                              " and Isnull(Locked,'')='N' and BPLId='" & InDT.Rows(Row)("branchc").ToString & "'")
        '                        End If
        '                    Else
        '                        objjournalentry.AutoVAT = BoYesNoEnum.tNO
        '                        objjournalentry.AutomaticWT = BoYesNoEnum.tNO
        '                        If objaddon.HANA Then
        '                            Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 ""Series"" FROM NNM1 WHERE ""ObjectCode""='30' and ""Indicator""=(Select ""Indicator"" from OFPR where '" & DocDate.ToString("yyyyMMdd") & "' Between ""F_RefDate"" and ""T_RefDate"") " &
        '                                                                              " and Ifnull(""Locked"",'')='N' and ""BPLId""='" & InDT.Rows(Row)("branchc").ToString & "'")
        '                        Else
        '                            Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 Series FROM NNM1 WHERE ObjectCode='30' and Indicator=(Select Indicator from OFPR where '" & DocDate.ToString("yyyyMMdd") & "' Between F_RefDate and T_RefDate) " &
        '                                                                              " and Isnull(Locked,'')='N' and BPLId='" & InDT.Rows(Row)("branchc").ToString & "'")
        '                        End If
        '                    End If
        '                    If Series <> "" Then objjournalentry.Series = Series
        '                    If InDT.Rows(Row)("doccur").ToString = MainCurr Then
        '                        JEAmount = IIf(CDbl(InDT.Rows(Row)("paytot").ToString) < 0, -CDbl(InDT.Rows(Row)("paytot").ToString), CDbl(InDT.Rows(Row)("paytot").ToString))
        '                    Else
        '                        ExRate = GetTransaction_ExchangeRate(InDT.Rows(Row)("doccur").ToString, InDT.Rows(Row)("object").ToString, InDT.Rows(Row)("transid").ToString)
        '                        JEAmount = IIf(CDbl(InDT.Rows(Row)("paytot").ToString) < 0, Math.Round(ExRate * -CDbl(InDT.Rows(Row)("paytot").ToString), 6), Math.Round(ExRate * CDbl(InDT.Rows(Row)("paytot").ToString), 6))
        '                    End If

        '                    GLCode = objaddon.objglobalmethods.getSingleValue("select ""PmtClrAct"" ""ControlAccount"" from OBPL where ""BPLId""='" & InDT.Rows(Row)("branchc").ToString & "'")
        '                    objjournalentry.Lines.AccountCode = GLCode
        '                    If CDbl(InDT.Rows(Row)("paytot").ToString) < 0 Then objjournalentry.Lines.Credit = JEAmount Else objjournalentry.Lines.Debit = JEAmount
        '                    objjournalentry.Lines.BPLID = InDT.Rows(Row)("branchc").ToString
        '                    If InDT.Rows(Row)("cc1").ToString <> "" Then objjournalentry.Lines.CostingCode = InDT.Rows(Row)("cc1").ToString
        '                    If InDT.Rows(Row)("cc2").ToString <> "" Then objjournalentry.Lines.CostingCode2 = InDT.Rows(Row)("cc2").ToString
        '                    If InDT.Rows(Row)("cc3").ToString <> "" Then objjournalentry.Lines.CostingCode3 = InDT.Rows(Row)("cc3").ToString
        '                    If InDT.Rows(Row)("cc4").ToString <> "" Then objjournalentry.Lines.CostingCode4 = InDT.Rows(Row)("cc4").ToString
        '                    If InDT.Rows(Row)("cc5").ToString <> "" Then objjournalentry.Lines.CostingCode5 = InDT.Rows(Row)("cc5").ToString
        '                    objjournalentry.Lines.Add()
        '                    objjournalentry.Lines.ShortName = InDT.Rows(Row)("cardc").ToString
        '                    If CDbl(InDT.Rows(Row)("paytot").ToString) < 0 Then objjournalentry.Lines.Debit = JEAmount Else objjournalentry.Lines.Credit = JEAmount
        '                    objjournalentry.Lines.BPLID = InDT.Rows(Row)("branchc").ToString
        '                    If InDT.Rows(Row)("cc1").ToString <> "" Then objjournalentry.Lines.CostingCode = InDT.Rows(Row)("cc1").ToString
        '                    If InDT.Rows(Row)("cc2").ToString <> "" Then objjournalentry.Lines.CostingCode2 = InDT.Rows(Row)("cc2").ToString
        '                    If InDT.Rows(Row)("cc3").ToString <> "" Then objjournalentry.Lines.CostingCode3 = InDT.Rows(Row)("cc3").ToString
        '                    If InDT.Rows(Row)("cc4").ToString <> "" Then objjournalentry.Lines.CostingCode4 = InDT.Rows(Row)("cc4").ToString
        '                    If InDT.Rows(Row)("cc5").ToString <> "" Then objjournalentry.Lines.CostingCode5 = InDT.Rows(Row)("cc5").ToString
        '                    objjournalentry.Lines.Add()
        '                    If objjournalentry.Add <> 0 Then
        '                        If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
        '                        objaddon.objapplication.SetStatusBarMessage("Journal: " & objaddon.objcompany.GetLastErrorDescription & "-" & objaddon.objcompany.GetLastErrorCode, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        '                        Return False
        '                    Else
        '                        If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
        '                        TransId = objaddon.objcompany.GetNewObjectKey()

        '                        Matrix0.Columns.Item("jeno").Cells.Item(CInt(InDT.Rows(Row)("#").ToString)).Specific.String = TransId
        '                        If Matrix0.Columns.Item("jeno").Cells.Item(CInt(InDT.Rows(Row)("#").ToString)).Specific.String <> "" And Matrix0.Columns.Item("recono").Cells.Item(CInt(InDT.Rows(Row)("#").ToString)).Specific.String = "" Then
        '                            'JEInternReco(InDT.Rows(Row)("transid").ToString, Row, TransId, CDbl(InDT.Rows(Row)("paytot").ToString))
        '                            'JEInternReco(InDT.Rows(Row)("transid").ToString, InDT.Rows(Row)("tranline").ToString, CInt(InDT.Rows(Row)("#").ToString), TransId, CDbl(InDT.Rows(Row)("paytot").ToString))
        '                            JEInternReco(InDT, Row, InDT.Rows(Row)("cardc").ToString, TransId)
        '                        End If
        '                        objaddon.objapplication.SetStatusBarMessage("Journal Entry Created Successfully...", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
        '                    End If
        '                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objjournalentry)
        '                Else
        '                    Try
        '                        If Not InDT.Rows(Row)("jeno").ToString = String.Empty And InDT.Rows(Row)("branchc").ToString = Branch And (InDT.Rows(Row)("recono").ToString = String.Empty Or InDT.Rows(Row)("recono").ToString = "0") Then
        '                            'JEInternReco(InDT.Rows(Row)("transid").ToString, InDT.Rows(Row)("tranline").ToString, CInt(InDT.Rows(Row)("#").ToString), InDT.Rows(Row)("jeno").ToString, CDbl(InDT.Rows(Row)("paytot").ToString))
        '                            JEInternReco(InDT, Row, InDT.Rows(Row)("cardc").ToString, InDT.Rows(Row)("jeno").ToString)
        '                        End If
        '                    Catch ex As Exception
        '                        objaddon.objapplication.SetStatusBarMessage("JE Rec " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        '                    End Try
        '                End If
        '            Next
        '        Catch ex As Exception
        '            If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
        '            objaddon.objapplication.SetStatusBarMessage("JE Posting Error" & objaddon.objcompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        '        End Try


        '        'For j = 1 To Matrix0.VisualRowCount
        '        '    If CDbl(Matrix0.Columns.Item("paytot").Cells.Item(j).Specific.String) <> 0 Then
        '        '        GLCode = objaddon.objglobalmethods.getSingleValue("select ""PmtClrAct"" ""ControlAccount"" from OBPL where ""BPLId""='" & Matrix0.Columns.Item("branchc").Cells.Item(j).Specific.String & "'")
        '        '        'objjournalentry.Lines.ShortName = Matrix0.Columns.Item("cardc").Cells.Item(j).Specific.String
        '        '        If CDbl(Matrix0.Columns.Item("paytot").Cells.Item(j).Specific.String) < 0 Then objjournalentry.Lines.ShortName = Matrix0.Columns.Item("cardc").Cells.Item(j).Specific.String Else objjournalentry.Lines.AccountCode = GLCode
        '        '        If CDbl(Matrix0.Columns.Item("paytot").Cells.Item(j).Specific.String) < 0 Then objjournalentry.Lines.Debit = CDbl(Matrix0.Columns.Item("paytot").Cells.Item(j).Specific.String) Else objjournalentry.Lines.Credit = CDbl(Matrix0.Columns.Item("paytot").Cells.Item(j).Specific.String)
        '        '        objjournalentry.Lines.Debit = 0
        '        '        objjournalentry.Lines.BPLID = Matrix0.Columns.Item("branchc").Cells.Item(j).Specific.String
        '        '        objjournalentry.Lines.Add()
        '        '        'objjournalentry.Lines.AccountCode = ""
        '        '        'objjournalentry.Lines.ShortName = Matrix0.Columns.Item("cardc").Cells.Item(j).Specific.String
        '        '        'If CDbl(Matrix0.Columns.Item("paytot").Cells.Item(j).Specific.String) < 0 Then objjournalentry.Lines.Debit = CDbl(Matrix0.Columns.Item("paytot").Cells.Item(j).Specific.String) Else objjournalentry.Lines.Credit = CDbl(Matrix0.Columns.Item("paytot").Cells.Item(j).Specific.String)
        '        '        'objjournalentry.Lines.Credit = 0
        '        '        'objjournalentry.Lines.BPLID = Matrix0.Columns.Item("branchc").Cells.Item(j).Specific.String
        '        '        'objjournalentry.Lines.Add()
        '        '    End If
        '        '    If Matrix0.Columns.Item("cc1").Cells.Item(j).Specific.String <> "" Then objjournalentry.Lines.CostingCode = Matrix0.Columns.Item("cc1").Cells.Item(j).Specific.String
        '        '    If Matrix0.Columns.Item("cc2").Cells.Item(j).Specific.String <> "" Then objjournalentry.Lines.CostingCode2 = Matrix0.Columns.Item("cc2").Cells.Item(j).Specific.String
        '        '    If Matrix0.Columns.Item("cc3").Cells.Item(j).Specific.String <> "" Then objjournalentry.Lines.CostingCode3 = Matrix0.Columns.Item("cc3").Cells.Item(j).Specific.String
        '        '    If Matrix0.Columns.Item("cc4").Cells.Item(j).Specific.String <> "" Then objjournalentry.Lines.CostingCode4 = Matrix0.Columns.Item("cc4").Cells.Item(j).Specific.String
        '        '    If Matrix0.Columns.Item("cc5").Cells.Item(j).Specific.String <> "" Then objjournalentry.Lines.CostingCode5 = Matrix0.Columns.Item("cc5").Cells.Item(j).Specific.String
        '        '    'If oMatrix1.Columns.Item("Remarks").Cells.Item(j).Specific.String <> "" Then objjournalentry.Lines.Reference1 = oMatrix1.Columns.Item("Remarks").Cells.Item(j).Specific.String
        '        '    'objjournalentry.Lines.LocationCode = ""

        '        'Next

        '        objrecset = Nothing
        '        Return True
        '    Catch ex As Exception
        '        'If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
        '        objaddon.objapplication.SetStatusBarMessage("JE " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        '        Return False
        '    Finally
        '        GC.Collect()
        '        GC.WaitForPendingFinalizers()
        '    End Try
        'End Function

        Private Function Disc_JournalEntry(ByVal InDT As DataTable) As Boolean
            Try
                Dim TransId, Series As String
                'Dim objrecset As SAPbobsCOM.Recordset
                Dim objjournalentry As SAPbobsCOM.JournalEntries
                Dim JEAmount As Double
                Try
                    For Row As Integer = 0 To InDT.Rows.Count - 1
                        Dim Calculate As Double = (CDbl(InDT.Rows(Row)("cashdisc").ToString) / CDbl(InDT.Rows(Row)("paytot").ToString)) * 100

                        If Calculate > 0.01 And CDbl(InDT.Rows(Row)("cashdisc").ToString) <> 0 And InDT.Rows(Row)("cdjeno").ToString = String.Empty Then
                            objjournalentry = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
                            objaddon.objapplication.StatusBar.SetText("Journal Entry Creating Please wait...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            'If Not objaddon.objcompany.InTransaction Then objaddon.objcompany.StartTransaction()
                            Dim oEdit As SAPbouiCOM.EditText
                            oEdit = objform.Items.Item("tdocdate").Specific
                            Dim DocDate As Date = Date.ParseExact(oEdit.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                            objjournalentry.ReferenceDate = DocDate 'ConvertDate.ToString("dd/MM/yy") 'DocDate 'Now.Date.ToString("yyyyMMdd") 
                            'objjournalentry.DueDate = Now.Date.ToString("yyyyMMdd") 'DocDate
                            objjournalentry.TaxDate = DocDate  ' ConvertDate.ToString("dd/MM/yy") 'DocDate 'Now.Date.ToString("yyyyMMdd") 

                            objjournalentry.Reference = "In Payment JE"
                            objjournalentry.Memo = "Auto Posted On: " & Now.ToString
                            objjournalentry.UserFields.Fields.Item("U_PayInNo").Value = EditText2.Value
                            If Localization = "IN" Then
                                If objaddon.HANA Then
                                    Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 ""Series"" FROM NNM1 WHERE ""ObjectCode""='30' and ""Indicator""=(Select ""Indicator"" from OFPR where '" & DocDate.ToString("yyyyMMdd") & "' Between ""F_RefDate"" and ""T_RefDate"") " &
                                                                                      " and Ifnull(""Locked"",'')='N' and ""BPLId""='" & InDT.Rows(Row)("branchc").ToString & "'")
                                Else
                                    Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 Series FROM NNM1 WHERE ObjectCode='30' and Indicator=(Select Indicator from OFPR where '" & DocDate.ToString("yyyyMMdd") & "' Between F_RefDate and T_RefDate) " &
                                                                                      " and Isnull(Locked,'')='N' and BPLId='" & InDT.Rows(Row)("branchc").ToString & "'")
                                End If
                            Else
                                If objaddon.HANA Then
                                    Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 ""Series"" FROM NNM1 WHERE ""ObjectCode""='30' and ""Indicator""=(Select ""Indicator"" from OFPR where '" & DocDate.ToString("yyyyMMdd") & "' Between ""F_RefDate"" and ""T_RefDate"") " &
                                                                                      " and Ifnull(""Locked"",'')='N' and ""BPLId""='" & InDT.Rows(Row)("branchc").ToString & "'")
                                Else
                                    Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 Series FROM NNM1 WHERE ObjectCode='30' and Indicator=(Select Indicator from OFPR where '" & DocDate.ToString("yyyyMMdd") & "' Between F_RefDate and T_RefDate) " &
                                                                                      " and Isnull(Locked,'')='N' and BPLId='" & InDT.Rows(Row)("branchc").ToString & "'")
                                End If
                            End If
                            If Series <> "" Then objjournalentry.Series = Series
                            JEAmount = IIf(CDbl(InDT.Rows(Row)("cashdisc").ToString) < 0, -CDbl(InDT.Rows(Row)("cashdisc").ToString), CDbl(InDT.Rows(Row)("cashdisc").ToString))
                            'GLCode = objaddon.objglobalmethods.getSingleValue("select ""PmtClrAct"" ""ControlAccount"" from OBPL where ""BPLId""='" & InDT.Rows(Row)("branchc").ToString & "'")
                            objjournalentry.Lines.AccountCode = RoundAcct
                            If CDbl(InDT.Rows(Row)("cashdisc").ToString) < 0 Then objjournalentry.Lines.Credit = JEAmount Else objjournalentry.Lines.Debit = JEAmount
                            objjournalentry.Lines.BPLID = InDT.Rows(Row)("branchc").ToString
                            If InDT.Rows(Row)("cc1").ToString <> "" Then objjournalentry.Lines.CostingCode = InDT.Rows(Row)("cc1").ToString
                            If InDT.Rows(Row)("cc2").ToString <> "" Then objjournalentry.Lines.CostingCode2 = InDT.Rows(Row)("cc2").ToString
                            If InDT.Rows(Row)("cc3").ToString <> "" Then objjournalentry.Lines.CostingCode3 = InDT.Rows(Row)("cc3").ToString
                            If InDT.Rows(Row)("cc4").ToString <> "" Then objjournalentry.Lines.CostingCode4 = InDT.Rows(Row)("cc4").ToString
                            If InDT.Rows(Row)("cc5").ToString <> "" Then objjournalentry.Lines.CostingCode5 = InDT.Rows(Row)("cc5").ToString
                            objjournalentry.Lines.Add()
                            objjournalentry.Lines.ShortName = InDT.Rows(Row)("cardc").ToString
                            If CDbl(InDT.Rows(Row)("cashdisc").ToString) < 0 Then objjournalentry.Lines.Debit = JEAmount Else objjournalentry.Lines.Credit = JEAmount
                            objjournalentry.Lines.BPLID = InDT.Rows(Row)("branchc").ToString
                            If InDT.Rows(Row)("cc1").ToString <> "" Then objjournalentry.Lines.CostingCode = InDT.Rows(Row)("cc1").ToString
                            If InDT.Rows(Row)("cc2").ToString <> "" Then objjournalentry.Lines.CostingCode2 = InDT.Rows(Row)("cc2").ToString
                            If InDT.Rows(Row)("cc3").ToString <> "" Then objjournalentry.Lines.CostingCode3 = InDT.Rows(Row)("cc3").ToString
                            If InDT.Rows(Row)("cc4").ToString <> "" Then objjournalentry.Lines.CostingCode4 = InDT.Rows(Row)("cc4").ToString
                            If InDT.Rows(Row)("cc5").ToString <> "" Then objjournalentry.Lines.CostingCode5 = InDT.Rows(Row)("cc5").ToString
                            objjournalentry.Lines.Add()
                            If objjournalentry.Add <> 0 Then
                                'If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                objaddon.objapplication.SetStatusBarMessage("Journal: " & objaddon.objcompany.GetLastErrorDescription & "-" & objaddon.objcompany.GetLastErrorCode, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(objjournalentry)
                                Return False
                            Else
                                'If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                TransId = objaddon.objcompany.GetNewObjectKey()
                                Matrix0.Columns.Item("cdjeno").Cells.Item(CInt(InDT.Rows(Row)("#").ToString)).Specific.String = TransId
                                If Matrix0.Columns.Item("cdjeno").Cells.Item(CInt(InDT.Rows(Row)("#").ToString)).Specific.String <> "" And Matrix0.Columns.Item("recono").Cells.Item(CInt(InDT.Rows(Row)("#").ToString)).Specific.String = "" Then
                                    'JEInternReco(InDT.Rows(Row)("transid").ToString, InDT.Rows(Row)("tranline").ToString, CInt(InDT.Rows(Row)("#").ToString), TransId, CDbl(InDT.Rows(Row)("cashdisc").ToString))
                                    JEInternReco(InDT, Row, InDT.Rows(Row)("cardc").ToString, TransId, "Y")
                                End If
                                objaddon.objapplication.SetStatusBarMessage("Journal Entry Created Successfully...", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(objjournalentry)
                                Return True
                            End If
                        Else
                            If Not InDT.Rows(Row)("cdjeno").ToString = String.Empty And InDT.Rows(Row)("recono").ToString = String.Empty Then
                                'JEInternReco(InDT.Rows(Row)("transid").ToString, InDT.Rows(Row)("tranline").ToString, CInt(InDT.Rows(Row)("#").ToString), InDT.Rows(Row)("cdjeno").ToString, CDbl(InDT.Rows(Row)("cashdisc").ToString))
                                JEInternReco(InDT, Row, InDT.Rows(Row)("cardc").ToString, InDT.Rows(Row)("cdjeno").ToString, "Y")
                            End If
                            Return True
                        End If
                    Next
                Catch ex As Exception
                    'If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    objaddon.objapplication.SetStatusBarMessage("JE Error" & objaddon.objcompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                End Try
            Catch ex As Exception
                'If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                'objaddon.objapplication.SetStatusBarMessage("JE Posting Error" & objaddon.objcompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                objaddon.objapplication.SetStatusBarMessage("JE " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                Return False
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End Function

        Private Function SameBranchReconciliation(ByVal InDT As DataTable) As Boolean
            Try
                Dim service As IInternalReconciliationsService = objaddon.objcompany.GetCompanyService().GetBusinessService(ServiceTypes.InternalReconciliationsService)
                Dim openTrans As InternalReconciliationOpenTrans = service.GetDataInterface(InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationOpenTrans)
                Dim reconParams As IInternalReconciliationParams = service.GetDataInterface(InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationParams)
                openTrans.CardOrAccount = CardOrAccountEnum.coaCard
                'If EditText15.Value <> "" Then Return False
                Dim RecAmount, ExRate As Double
                Dim GetStat As String
                Dim Row As Integer = 0

                objRs = objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                GetStat = objaddon.objglobalmethods.getSingleValue("SELECT CASE WHEN T0.""DebCred"" = 'C' AND T0.""BalDueCred"" = 0 THEN 'Reconciled' WHEN T0.""DebCred"" = 'C' AND T0.""Credit"" = T0.""BalDueCred"" THEN 'Unreconciled' " &
                                                                   " WHEN T0.""DebCred"" = 'C' AND T0.""Credit"" <> T0.""BalDueCred"" THEN 'Partial' WHEN T0.""DebCred"" = 'D' AND T0.""BalDueDeb"" = 0 THEN 'Reconciled' " &
                                                                   " WHEN T0.""DebCred"" = 'D' AND T0.""Debit"" = T0.""BalDueDeb"" THEN 'Unreconciled' WHEN T0.""DebCred"" = 'D' AND T0.""Debit"" <> T0.""BalDueDeb"" THEN 'Partial' " &
                                                                   " END ""Status"" FROM JDT1 T0 INNER JOIN OJDT T1 ON T0.""TransId"" = T1.""TransId"" WHERE T0.""TransId"" = (Select ""TransId"" from ORCT where ""DocEntry""=" & EditText13.Value & ") " &
                                                                   "AND T0.""ShortName"" = '" & CurBPCode & "'")
                If GetStat = "Reconciled" Then Return True
                Dim DocDate As Date = Date.ParseExact(EditText0.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                openTrans.ReconDate = DocDate
                For DTRow As Integer = 0 To InDT.Rows.Count - 1
                    If InDT.Rows(DTRow)("branchc").ToString = CurBranch Then
                        If Row = 0 Then
                            Dim Inpaytranid As String = objaddon.objglobalmethods.getSingleValue("select ""TransId"" from ORCT where ""DocEntry"" =" & EditText13.Value & "")
                            'Dim Amount As String = objaddon.objglobalmethods.getSingleValue("select ""NoDocSum"" from ORCT where ""DocEntry"" =" & EditText13.Value & "")
                            objRs.DoQuery("select CASE WHEN T1.""BalDueCred""<>0  THEN  T1.""BalDueCred"" ELSE T1.""BalDueDeb"" END AS ""Balance"",T1.""Line_ID"",T1.""DebCred"" from JDT1 T1 where T1.""ShortName"" ='" & CurBPCode & "' and T1.""TransId""=(Select ""TransId"" from ORCT where ""DocEntry""=" & EditText13.Value & ") ")
                            'Amount = CDbl(Amount) ' IIf(CDbl(Amount) < 0, -CDbl(Amount), CDbl(Amount)) '-CDbl(Amount) 
                            If objRs.RecordCount > 0 And CDbl(objRs.Fields.Item(0).Value.ToString) <> 0 Then
                                'If objRs.Fields.Item(2).Value.ToString = "C" Then RecAmount = CDbl(objRs.Fields.Item(0).Value.ToString) Else RecAmount = CDbl(objRs.Fields.Item(0).Value.ToString)
                                RecAmount = Math.Round(CDbl(objRs.Fields.Item(0).Value.ToString), SumRound)
                                openTrans.InternalReconciliationOpenTransRows.Add()
                                openTrans.InternalReconciliationOpenTransRows.Item(Row).Selected = BoYesNoEnum.tYES
                                openTrans.InternalReconciliationOpenTransRows.Item(Row).TransId = Inpaytranid 'InDT.Rows(Row)("transid").ToString
                                openTrans.InternalReconciliationOpenTransRows.Item(Row).TransRowId = CInt(objRs.Fields.Item(1).Value.ToString) '1 'InDT.Rows(Row)("tranline").ToString
                                openTrans.InternalReconciliationOpenTransRows.Item(Row).ReconcileAmount = RecAmount ' CDbl(objRs.Fields.Item(0).Value.ToString)
                                Row += 1
                            End If
                            If EditText22.Value <> "" Then 'Forex
                                objRs.DoQuery("select CASE WHEN T1.""BalDueCred""<>0  THEN  T1.""BalDueCred"" ELSE T1.""BalDueDeb"" END AS ""Balance"",T1.""Line_ID"",T1.""DebCred"" from JDT1 T1 where T1.""ShortName"" ='" & CurBPCode & "' and T1.""TransId""=" & EditText22.Value & "")
                                ' If objRs.Fields.Item(2).Value.ToString = "C" Then RecAmount = CDbl(objRs.Fields.Item(0).Value.ToString) Else RecAmount = CDbl(objRs.Fields.Item(0).Value.ToString)
                                RecAmount = Math.Round(CDbl(objRs.Fields.Item(0).Value.ToString), SumRound)
                                openTrans.InternalReconciliationOpenTransRows.Add()
                                openTrans.InternalReconciliationOpenTransRows.Item(Row).Selected = BoYesNoEnum.tYES
                                openTrans.InternalReconciliationOpenTransRows.Item(Row).TransId = EditText22.Value
                                openTrans.InternalReconciliationOpenTransRows.Item(Row).TransRowId = CInt(objRs.Fields.Item(1).Value.ToString)
                                openTrans.InternalReconciliationOpenTransRows.Item(Row).ReconcileAmount = RecAmount
                                Row += 1
                            End If

                            If EditText14.Value <> "" Then 'je
                                objRs.DoQuery("select CASE WHEN T1.""BalDueCred""<>0  THEN  T1.""BalDueCred"" ELSE T1.""BalDueDeb"" END AS ""Balance"",T1.""Line_ID"",T1.""DebCred"" from OJDT T0 join JDT1 T1 ON T0.""TransId""=T1.""TransId"" where T1.""BPLId""='" & CurBranch & "' and T1.""TransId""='" & EditText14.Value & "' and T1.""ShortName""='" & CurBPCode & "'")
                                'If Val(RecAmount) <> 0 Then RecAmount = CDbl(RecAmount) ' IIf(CDbl(Amount) < 0, -CDbl(Amount), CDbl(Amount)) ' -CDbl(Amount)
                                If objRs.RecordCount > 0 Then
                                    If Val(objRs.Fields.Item(0).Value.ToString) <> 0 Then
                                        'If objRs.Fields.Item(2).Value.ToString = "C" Then RecAmount = CDbl(objRs.Fields.Item(0).Value.ToString) Else RecAmount = CDbl(objRs.Fields.Item(0).Value.ToString)
                                        RecAmount = Math.Round(CDbl(objRs.Fields.Item(0).Value.ToString), SumRound)
                                        'RecAmount = CDbl(RecAmount) ' IIf(CDbl(Amount) < 0, -CDbl(Amount), CDbl(Amount))  '-CDbl(Amount) 
                                        openTrans.InternalReconciliationOpenTransRows.Add()
                                        openTrans.InternalReconciliationOpenTransRows.Item(Row).Selected = BoYesNoEnum.tYES
                                        openTrans.InternalReconciliationOpenTransRows.Item(Row).TransId = EditText14.Value 'InDT.Rows(Row)("transid").ToString
                                        openTrans.InternalReconciliationOpenTransRows.Item(Row).TransRowId = CInt(objRs.Fields.Item(1).Value.ToString) 'InDT.Rows(Row)("tranline").ToString
                                        openTrans.InternalReconciliationOpenTransRows.Item(Row).ReconcileAmount = RecAmount ' CDbl(objRs.Fields.Item(0).Value.ToString)
                                        Row += 1
                                    End If
                                End If
                                If InDT.Rows(DTRow)("cardc").ToString <> CurBPCode Then
                                    'RecAmount = IIf(CDbl(InDT.Rows(DTRow)("paytot").ToString) < 0, -CDbl(InDT.Rows(DTRow)("paytot").ToString), CDbl(InDT.Rows(DTRow)("paytot").ToString))
                                    'RecAmount = -CDbl(InDT.Rows(DTRow)("paytot").ToString)
                                    If InDT.Rows(DTRow)("doccur").ToString = MainCurr Then
                                        If InDT.Rows(DTRow)("debcred").ToString = "C" Then RecAmount = Math.Round(-CDbl(InDT.Rows(DTRow)("paytot").ToString), SumRound) Else RecAmount = Math.Round(CDbl(InDT.Rows(DTRow)("paytot").ToString), SumRound)
                                    Else
                                        ExRate = Math.Round(GetTransaction_ExchangeRate(InDT.Rows(DTRow)("object").ToString, InDT.Rows(DTRow)("transid").ToString), RateRound) ' DocumentDate.ToString("yyyyMMdd")
                                        If InDT.Rows(DTRow)("debcred").ToString = "C" Then RecAmount = Math.Round(ExRate * -CDbl(InDT.Rows(DTRow)("paytot").ToString), SumRound) Else RecAmount = Math.Round(ExRate * CDbl(InDT.Rows(DTRow)("paytot").ToString), SumRound)
                                    End If

                                    openTrans.InternalReconciliationOpenTransRows.Add()
                                    openTrans.InternalReconciliationOpenTransRows.Item(Row).Selected = BoYesNoEnum.tYES
                                    openTrans.InternalReconciliationOpenTransRows.Item(Row).TransId = InDT.Rows(DTRow)("transid").ToString
                                    openTrans.InternalReconciliationOpenTransRows.Item(Row).TransRowId = InDT.Rows(DTRow)("tranline").ToString
                                    openTrans.InternalReconciliationOpenTransRows.Item(Row).ReconcileAmount = RecAmount
                                    Row += 1
                                End If
                            Else
                                If InDT.Rows(DTRow)("cardc").ToString <> CurBPCode Then
                                    'RecAmount = IIf(CDbl(InDT.Rows(DTRow)("paytot").ToString) < 0, -CDbl(InDT.Rows(DTRow)("paytot").ToString), CDbl(InDT.Rows(DTRow)("paytot").ToString))
                                    If InDT.Rows(DTRow)("doccur").ToString = MainCurr Then
                                        If InDT.Rows(DTRow)("debcred").ToString = "C" Then RecAmount = Math.Round(-CDbl(InDT.Rows(DTRow)("paytot").ToString), SumRound) Else RecAmount = Math.Round(CDbl(InDT.Rows(DTRow)("paytot").ToString), SumRound)
                                    Else
                                        ExRate = Math.Round(GetTransaction_ExchangeRate(InDT.Rows(DTRow)("object").ToString, InDT.Rows(DTRow)("transid").ToString), RateRound)
                                        If InDT.Rows(DTRow)("debcred").ToString = "C" Then RecAmount = Math.Round(ExRate * -CDbl(InDT.Rows(DTRow)("paytot").ToString), SumRound) Else RecAmount = Math.Round(ExRate * CDbl(InDT.Rows(DTRow)("paytot").ToString), SumRound)
                                    End If
                                    openTrans.InternalReconciliationOpenTransRows.Add()
                                    openTrans.InternalReconciliationOpenTransRows.Item(Row).Selected = BoYesNoEnum.tYES
                                    openTrans.InternalReconciliationOpenTransRows.Item(Row).TransId = InDT.Rows(DTRow)("transid").ToString
                                    openTrans.InternalReconciliationOpenTransRows.Item(Row).TransRowId = InDT.Rows(DTRow)("tranline").ToString
                                    openTrans.InternalReconciliationOpenTransRows.Item(Row).ReconcileAmount = RecAmount
                                    Row += 1
                                End If
                            End If
                        Else
                            'RecAmount = -CDbl(InDT.Rows(DTRow)("paytot").ToString)
                            If InDT.Rows(DTRow)("cardc").ToString <> CurBPCode Then
                                If InDT.Rows(DTRow)("doccur").ToString = MainCurr Then
                                    If InDT.Rows(DTRow)("debcred").ToString = "C" Then RecAmount = Math.Round(-CDbl(InDT.Rows(DTRow)("paytot").ToString), SumRound) Else RecAmount = Math.Round(CDbl(InDT.Rows(DTRow)("paytot").ToString), SumRound)
                                Else
                                    ExRate = Math.Round(GetTransaction_ExchangeRate(InDT.Rows(DTRow)("object").ToString, InDT.Rows(DTRow)("transid").ToString), RateRound)
                                    If InDT.Rows(DTRow)("debcred").ToString = "C" Then RecAmount = Math.Round(ExRate * -CDbl(InDT.Rows(DTRow)("paytot").ToString), SumRound) Else RecAmount = Math.Round(ExRate * CDbl(InDT.Rows(DTRow)("paytot").ToString), SumRound)
                                End If
                                openTrans.InternalReconciliationOpenTransRows.Add()
                                openTrans.InternalReconciliationOpenTransRows.Item(Row).Selected = BoYesNoEnum.tYES
                                openTrans.InternalReconciliationOpenTransRows.Item(Row).TransId = InDT.Rows(DTRow)("transid").ToString
                                openTrans.InternalReconciliationOpenTransRows.Item(Row).TransRowId = InDT.Rows(DTRow)("tranline").ToString
                                openTrans.InternalReconciliationOpenTransRows.Item(Row).ReconcileAmount = RecAmount
                                Row += 1
                            End If
                        End If
                    End If
                Next
                Try
                    reconParams = service.Add(openTrans)
                Catch ex As Exception
                    objaddon.objapplication.StatusBar.SetText("SameBranch Reco: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End Try
                Dim Reconum As Integer = reconParams.ReconNum
                If Reconum = 0 Then objaddon.objapplication.StatusBar.SetText("SameBranch_Reco Error...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : Return False
                If EditText15.Value = "" Then
                    EditText15.Value = Reconum
                Else
                    EditText15.Value = EditText15.Value & "," & Reconum
                End If
                objaddon.objapplication.StatusBar.SetText("SameBranch Reconciled successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(openTrans)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(reconParams)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(service)
                GC.Collect()
                Return True
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("SameBranch Recon: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End Try

        End Function

        Private Function Forex_JEInternReco(ByVal Intransid As String, ByVal ForexAmount As Double, ByVal Forextransid As String) As Boolean
            Try
                Dim objRs1 As SAPbobsCOM.Recordset
                Dim service As IInternalReconciliationsService = objaddon.objcompany.GetCompanyService().GetBusinessService(ServiceTypes.InternalReconciliationsService)
                Dim openTrans As InternalReconciliationOpenTrans = service.GetDataInterface(InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationOpenTrans)
                Dim reconParams As IInternalReconciliationParams = service.GetDataInterface(InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationParams)
                openTrans.CardOrAccount = CardOrAccountEnum.coaCard
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRs1 = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim RecAmount As Double

                objRs1.DoQuery("Select T1.""DebCred"",T1.""Line_ID"",CASE WHEN T1.""BalDueCred""<>0  THEN  T1.""BalDueCred"" ELSE T1.""BalDueDeb"" END AS ""Balance"" from JDT1 T1 where T1.""ShortName"" in (Select ""CardCode"" from OCRD) and T1.""TransId"" = " & Forextransid & "")
                If objRs1.Fields.Item(0).Value.ToString = "C" Then RecAmount = ForexAmount Else RecAmount = ForexAmount
                RecAmount = ForexAmount
                RecAmount = CDbl(objRs1.Fields.Item(2).Value.ToString)
                'objRs.DoQuery("Select T1.""DebCred"",T1.""Line_ID"",CASE WHEN T1.""BalDueCred""<>0  THEN  T1.""BalDueCred"" ELSE T1.""BalDueDeb"" END AS ""Balance"" from JDT1 T1 where T1.""TransId"" = " & Intransid & "")
                objRs.DoQuery("select T1.""DebCred"",T1.""Line_ID"",CASE WHEN T1.""BalDueCred""<>0  THEN  T1.""BalDueCred"" ELSE T1.""BalDueDeb"" END AS ""Balance"" from JDT1 T1 where T1.""ShortName"" in (Select ""CardCode"" from OCRD) and T1.""TransId""=(Select ""TransId"" from ORCT where ""DocEntry""=" & EditText13.Value & ") ")
                RecAmount = CDbl(objRs.Fields.Item(2).Value.ToString)
                If ForexAmount <> CDbl(objRs.Fields.Item(2).Value.ToString) Then Return True

                openTrans.InternalReconciliationOpenTransRows.Add()
                openTrans.InternalReconciliationOpenTransRows.Item(0).Selected = BoYesNoEnum.tYES
                openTrans.InternalReconciliationOpenTransRows.Item(0).TransId = Forextransid
                openTrans.InternalReconciliationOpenTransRows.Item(0).TransRowId = objRs1.Fields.Item(1).Value.ToString '0
                openTrans.InternalReconciliationOpenTransRows.Item(0).ReconcileAmount = CDbl(objRs1.Fields.Item(2).Value.ToString) 'RecAmount
                openTrans.InternalReconciliationOpenTransRows.Add()
                openTrans.InternalReconciliationOpenTransRows.Item(1).Selected = BoYesNoEnum.tYES
                openTrans.InternalReconciliationOpenTransRows.Item(1).TransId = Intransid
                openTrans.InternalReconciliationOpenTransRows.Item(1).TransRowId = objRs.Fields.Item(1).Value.ToString '1
                openTrans.InternalReconciliationOpenTransRows.Item(1).ReconcileAmount = CDbl(objRs.Fields.Item(2).Value.ToString) ' Amount
                Try
                    reconParams = service.Add(openTrans)
                Catch ex As Exception
                End Try

                Dim Reconum As Integer = reconParams.ReconNum
                'EditText23.Value = Reconum
                objaddon.objapplication.StatusBar.SetText("Reconciled successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(openTrans)
                GC.Collect()
                Return True
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("Recon: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End Try

        End Function

        Private Function JEInternReco(ByVal InDT As DataTable, ByVal Row As Integer, ByVal cardcode As String, ByVal transid As String, Optional ByVal Disc As String = "") As Boolean
            Try
                Dim service As IInternalReconciliationsService = objaddon.objcompany.GetCompanyService().GetBusinessService(ServiceTypes.InternalReconciliationsService)
                Dim openTrans As InternalReconciliationOpenTrans = service.GetDataInterface(InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationOpenTrans)
                Dim reconParams As IInternalReconciliationParams = service.GetDataInterface(InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationParams)
                openTrans.CardOrAccount = CardOrAccountEnum.coaCard
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim RecAmount, Amount, ExRate As Double '= IIf(CDbl(InDT.Rows(Row)("paytot").ToString) < 0, -CDbl(InDT.Rows(Row)("paytot").ToString), CDbl(InDT.Rows(Row)("paytot").ToString))

                If InDT.Rows(Row)("doccur").ToString = MainCurr Then
                    If InDT.Rows(Row)("debcred").ToString = "C" Then RecAmount = Math.Round(-CDbl(InDT.Rows(Row)("paytot").ToString), SumRound) Else RecAmount = Math.Round(CDbl(InDT.Rows(Row)("paytot").ToString), SumRound)
                Else
                    ExRate = Math.Round(GetTransaction_ExchangeRate(InDT.Rows(Row)("object").ToString, InDT.Rows(Row)("transid").ToString), RateRound)
                    If InDT.Rows(Row)("debcred").ToString = "C" Then RecAmount = Math.Round(ExRate * -CDbl(InDT.Rows(Row)("paytot").ToString), SumRound) Else RecAmount = Math.Round(ExRate * CDbl(InDT.Rows(Row)("paytot").ToString), SumRound)
                End If

                objRs.DoQuery("Select T1.""DebCred"",T1.""Line_ID"",CASE WHEN T1.""BalDueCred""<>0  THEN  T1.""BalDueCred"" ELSE T1.""BalDueDeb"" END AS ""Balance"" from JDT1 T1 where T1.""ShortName"" ='" & cardcode & "' and T1.""TransId"" = " & transid & "")
                If objRs.RecordCount > 0 Then
                    'If objRs.Fields.Item(0).Value.ToString = "C" Then Amount = CDbl(objRs.Fields.Item(2).Value.ToString) Else Amount = CDbl(objRs.Fields.Item(2).Value.ToString)
                    Amount = CDbl(objRs.Fields.Item(2).Value.ToString)
                End If
                Dim DocDate As Date = Date.ParseExact(EditText0.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                openTrans.ReconDate = DocDate
                openTrans.InternalReconciliationOpenTransRows.Add()
                openTrans.InternalReconciliationOpenTransRows.Item(0).Selected = BoYesNoEnum.tYES
                openTrans.InternalReconciliationOpenTransRows.Item(0).TransId = InDT.Rows(Row)("transid").ToString
                openTrans.InternalReconciliationOpenTransRows.Item(0).TransRowId = InDT.Rows(Row)("tranline").ToString  '0
                openTrans.InternalReconciliationOpenTransRows.Item(0).ReconcileAmount = RecAmount 'RecAmount
                openTrans.InternalReconciliationOpenTransRows.Add()
                openTrans.InternalReconciliationOpenTransRows.Item(1).Selected = BoYesNoEnum.tYES
                openTrans.InternalReconciliationOpenTransRows.Item(1).TransId = transid
                openTrans.InternalReconciliationOpenTransRows.Item(1).TransRowId = objRs.Fields.Item(1).Value.ToString '1
                openTrans.InternalReconciliationOpenTransRows.Item(1).ReconcileAmount = Amount
                Try
                    reconParams = service.Add(openTrans)
                Catch ex As Exception
                    objaddon.objapplication.StatusBar.SetText("Disc Reco: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End Try

                Dim Reconum As Integer = reconParams.ReconNum
                If Reconum = 0 Then objaddon.objapplication.StatusBar.SetText("Reconciled Error...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : Return False
                'If EditText15.Value = "" Then
                '    EditText15.Value = Reconum
                'Else
                '    EditText15.Value = EditText15.Value & "," & Reconum
                'End If
                If Disc = "Y" Then
                    Matrix0.Columns.Item("cdrecono").Cells.Item(CInt(InDT.Rows(Row)("#").ToString)).Specific.String = Reconum
                Else
                    Matrix0.Columns.Item("recono").Cells.Item(CInt(InDT.Rows(Row)("#").ToString)).Specific.String = Reconum
                End If
                objaddon.objapplication.StatusBar.SetText("Reconciled successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(openTrans)
                GC.Collect()
                Return True
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("Recon: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End Try

        End Function

        'Private Function InternReco(ByVal InDT As DataTable) As Boolean
        '    Try
        '        Dim service As IInternalReconciliationsService = objaddon.objcompany.GetCompanyService().GetBusinessService(ServiceTypes.InternalReconciliationsService)
        '        Dim openTrans As InternalReconciliationOpenTrans = service.GetDataInterface(InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationOpenTrans)
        '        Dim reconParams As IInternalReconciliationParams = service.GetDataInterface(InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationParams)
        '        openTrans.CardOrAccount = CardOrAccountEnum.coaCard
        '        'If EditText3.Value <> 0 Then Return False
        '        'Dim RecAmount As Double = IIf(CDbl(InDT.Rows(RowID)("paytot").ToString) < 0, -CDbl(InDT.Rows(RowID)("paytot").ToString), CDbl(InDT.Rows(RowID)("paytot").ToString))
        '        'openTrans.InternalReconciliationOpenTransRows.Add()
        '        'openTrans.InternalReconciliationOpenTransRows.Item(0).Selected = BoYesNoEnum.tYES
        '        'openTrans.InternalReconciliationOpenTransRows.Item(0).TransId = InDT.Rows(Row)("transid").ToString
        '        'openTrans.InternalReconciliationOpenTransRows.Item(0).TransRowId = 0
        '        'openTrans.InternalReconciliationOpenTransRows.Item(0).ReconcileAmount = RecAmount
        '        'openTrans.InternalReconciliationOpenTransRows.Add()
        '        'openTrans.InternalReconciliationOpenTransRows.Item(1).Selected = BoYesNoEnum.tYES
        '        'openTrans.InternalReconciliationOpenTransRows.Item(1).TransId = InDT.Rows(Row)("transid").ToString
        '        'openTrans.InternalReconciliationOpenTransRows.Item(1).TransRowId = 1
        '        'openTrans.InternalReconciliationOpenTransRows.Item(1).ReconcileAmount = RecAmount

        '        Dim Row As Integer = 0
        '        'For DTRow As Integer = 0 To InDT.Rows.Count - 1
        '        '    Dim RecAmount As Double = IIf(CDbl(InDT.Rows(Row)("paytot").ToString) < 0, -CDbl(InDT.Rows(Row)("paytot").ToString), CDbl(InDT.Rows(Row)("paytot").ToString))
        '        '    openTrans.InternalReconciliationOpenTransRows.Add()
        '        '    openTrans.InternalReconciliationOpenTransRows.Item(Row).Selected = BoYesNoEnum.tYES
        '        '    openTrans.InternalReconciliationOpenTransRows.Item(Row).TransId = InDT.Rows(Row)("transid").ToString
        '        '    openTrans.InternalReconciliationOpenTransRows.Item(Row).TransRowId = InDT.Rows(Row)("tranline").ToString
        '        '    openTrans.InternalReconciliationOpenTransRows.Item(Row).ReconcileAmount = RecAmount
        '        '    Row += 1
        '        'Next
        '        For i As Integer = 1 To Matrix0.VisualRowCount
        '            If Matrix0.Columns.Item("select").Cells.Item(i).Specific.Checked = True Then
        '                Dim RecAmount As Double = IIf(CDbl(Matrix0.Columns.Item("paytot").Cells.Item(i).Specific.String) < 0, -CDbl(Matrix0.Columns.Item("paytot").Cells.Item(i).Specific.String), CDbl(Matrix0.Columns.Item("paytot").Cells.Item(i).Specific.String))
        '                RecAmount = -CDbl(Matrix0.Columns.Item("paytot").Cells.Item(i).Specific.String)
        '                openTrans.InternalReconciliationOpenTransRows.Add()
        '                openTrans.InternalReconciliationOpenTransRows.Item(Row).Selected = BoYesNoEnum.tYES
        '                openTrans.InternalReconciliationOpenTransRows.Item(Row).TransId = Matrix0.Columns.Item("transid").Cells.Item(i).Specific.String
        '                openTrans.InternalReconciliationOpenTransRows.Item(Row).TransRowId = Matrix0.Columns.Item("tranline").Cells.Item(i).Specific.String
        '                openTrans.InternalReconciliationOpenTransRows.Item(Row).ReconcileAmount = RecAmount
        '                Row += 1
        '            End If
        '        Next
        '        reconParams = service.Add(openTrans)
        '        Dim Reconum As Integer = reconParams.ReconNum
        '        If EditText15.Value = "" Then
        '            EditText15.Value = Reconum
        '        Else
        '            EditText15.Value = EditText15.Value & "," & Reconum
        '        End If
        '        objaddon.objapplication.StatusBar.SetText("Reconciled successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        '        System.Runtime.InteropServices.Marshal.ReleaseComObject(openTrans)
        '        GC.Collect()
        '        Return True
        '    Catch ex As Exception
        '        objaddon.objapplication.StatusBar.SetText("Recon: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        '        Return False
        '    End Try

        'End Function

        'Private Function PaymentReconciliation() As Boolean
        '    Try
        '        Dim service As InternalReconciliationsService = objaddon.objcompany.GetCompanyService.GetBusinessService(ServiceTypes.InternalReconciliationsService)
        '        Dim transParams As InternalReconciliationOpenTransParams = service.GetDataInterface(InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationOpenTransParams)
        '        'Dim openTrans As InternalReconciliationOpenTrans = service.GetDataInterface(InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationOpenTrans)
        '        Dim reconParams As IInternalReconciliationParams = service.GetDataInterface(InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationParams)
        '        'Dim transId As Integer = 190, transRowId As Integer = 1
        '        transParams.ReconDate = DateTime.Today
        '        transParams.DateType = ReconSelectDateTypeEnum.rsdtDocDate
        '        transParams.FromDate = New DateTime(2021, 4, 2) ' Now.Date ' 
        '        transParams.ToDate = New DateTime(2021, 8, 5) ' Now.Date '

        '        transParams.CardOrAccount = CardOrAccountEnum.coaCard
        '        transParams.InternalReconciliationBPs.Add()
        '        transParams.InternalReconciliationBPs.Item(0).BPCode = "C000003"
        '        Dim openTrans As InternalReconciliationOpenTrans = service.GetOpenTransactions(transParams)
        '        'Dim row As InternalReconciliationOpenTransRows = openTrans.InternalReconciliationOpenTransRows
        '        'openTrans.CardOrAccount = CardOrAccountEnum.coaCard
        '        For i As Integer = 1 To Matrix0.VisualRowCount
        '            If Matrix0.Columns.Item("select").Cells.Item(i).Specific.Checked = True Then
        '                For Each row In openTrans.InternalReconciliationOpenTransRows
        '                    If Matrix0.Columns.Item("cardc").Cells.Item(i).Specific.String = openTrans.InternalReconciliationOpenTransRows.Item(row).ShortName And (Matrix0.Columns.Item("transid").Cells.Item(i).Specific.String = openTrans.InternalReconciliationOpenTransRows.Item(row).TransId) Then
        '                        row.Selected = BoYesNoEnum.tYES
        '                        row.ReconcileAmount = CDbl(Matrix0.Columns.Item("paytot").Cells.Item(i).Specific.String)
        '                    End If
        '                    If (Matrix0.Columns.Item("transid").Cells.Item(i).Specific.String = row.TransId And Matrix0.Columns.Item("tranline").Cells.Item(i).Specific.String = row.TransRowId) Then
        '                        row.Selected = BoYesNoEnum.tYES
        '                        row.ReconcileAmount = CDbl(Matrix0.Columns.Item("paytot").Cells.Item(i).Specific.String)
        '                    End If
        '                Next
        '            End If
        '        Next
        '        If EditText3.Value <> 0 Then Return False
        '        'Dim reconParams As InternalReconciliationParams = service.GetDataInterface(InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationParams)
        '        reconParams = service.Add(openTrans)
        '        Dim ii As Integer = reconParams.ReconNum
        '        objaddon.objapplication.StatusBar.SetText("Reconciled successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        '        Return True
        '    Catch ex As Exception
        '        Return False
        '    End Try

        'End Function

        Private Sub EditText19_ValidateAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText19.ValidateAfter
            Try
                If pVal.ItemChanged = False Or pVal.InnerEvent = True Then Exit Sub
                Dim PayFlag As Boolean
                If Val(EditText5.Value) > 0 Or Val(EditText9.Value) > 0 Or Val(EditText11.Value) > 0 Or Val(EditText12.Value) > 0 Then
                    PayFlag = True
                End If
                If PayFlag = False Then
                    If Val(EditText19.Value) > 0 And Val(EditText3.Value) > 0 Then
                        EditText3.Value = EditText19.Value
                    Else
                        EditText3.Value = EditText23.Value
                    End If
                End If
            Catch ex As Exception
            End Try

        End Sub

        Private Function JournalEntry_BranchWise(ByVal InDT As DataTable, ByVal Branch As String, ByVal JEAmount As Double) As Boolean
            Try
                Dim TransId As String = "", GLCode As String, Series As String, CardCode As String = "", MatLine As String = ""
                Dim objrecset As SAPbobsCOM.Recordset
                Dim objjournalentry As SAPbobsCOM.JournalEntries
                Dim Amount, BranchTotal, ExRate, CurTot, forex, forexAmt As Double
                Dim DTLine As Integer = 0
                Try
                    objrecset = objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                    For DTRow As Integer = 0 To InDT.Rows.Count - 1
                        If InDT.Rows(DTRow)("branchc").ToString = Branch Then
                            TransId = InDT.Rows(DTRow)("jeno").ToString
                            MatLine = CInt(InDT.Rows(DTRow)("#").ToString)
                            DTLine = DTRow
                            CardCode = InDT.Rows(DTRow)("cardc").ToString()
                            Exit For
                        End If
                    Next

                    'Dim DocDate As Date = Date.ParseExact(EditText0.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                    For i As Integer = 0 To InDT.Rows.Count - 1
                        If InDT.Rows(i)("branchc").ToString = Branch Then
                            If InDT.Rows(i)("doccur").ToString = MainCurr Then
                                BranchTotal = Math.Round(BranchTotal + CDbl(InDT.Rows(i)("paytot").ToString), SumRound)
                                CurTot = Math.Round(CurTot + CDbl(InDT.Rows(i)("paytot").ToString), SumRound)
                            Else
                                strSQL = "Select ""Rate"",""Currency"" from ORTT where ""RateDate""= '" & DocumentDate.ToString("yyyyMMdd") & "' and ""Currency""='" & InDT.Rows(i)("doccur").ToString & "' "
                                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                objRs.DoQuery(strSQL)
                                If CDbl(InDT.Rows(i)("paytot").ToString) = CDbl(InDT.Rows(i)("baldue").ToString) Then
                                    CurTot = Math.Round(CurTot + (CDbl(InDT.Rows(i)("paytot").ToString) * Math.Round(CDbl(objRs.Fields.Item(0).Value.ToString), RateRound)), SumRound)
                                    'ExRate = Math.Round(GetTransaction_ExchangeRate(InDT.Rows(i)("object").ToString, InDT.Rows(i)("transid").ToString), RateRound)
                                    'BranchTotal = Math.Round(BranchTotal + (CDbl(InDT.Rows(i)("paytot").ToString) * Math.Round(ExRate, RateRound)), SumRound)
                                    strSQL = "Select Case when T1.""BalDueCred"" <>0 Then -T1.""BalDueCred"" Else T1.""BalDueDeb"" End as ""DocTotal"",T1.* from OJDT T0 join JDT1 T1 on T1.""TransId"" =T0.""TransId"" where T1.""TransId""='" & InDT.Rows(i)("transid").ToString & "' "
                                    strSQL += vbCrLf + "And T1.""ShortName"" in (select ""CardCode"" from OCRD)"
                                    objrecset.DoQuery(strSQL)
                                    If objrecset.RecordCount > 0 Then
                                        BranchTotal = Math.Round(BranchTotal + CDbl(objrecset.Fields.Item(0).Value.ToString), SumRound)
                                    End If
                                Else
                                    CurTot = Math.Round(CurTot + (CDbl(InDT.Rows(i)("paytot").ToString) * Math.Round(CDbl(objRs.Fields.Item(0).Value.ToString), RateRound)), SumRound)
                                    ExRate = Math.Round(GetTransaction_ExchangeRate(InDT.Rows(i)("object").ToString, InDT.Rows(i)("transid").ToString), RateRound)
                                    BranchTotal = Math.Round(BranchTotal + (CDbl(InDT.Rows(i)("paytot").ToString) * Math.Round(ExRate, RateRound)), SumRound) 'CDbl(objRs.Fields.Item(0).Value.ToString)
                                End If
                                'CurTot = Math.Round(CurTot + (CDbl(InDT.Rows(i)("paytot").ToString) * Math.Round(CDbl(objRs.Fields.Item(0).Value.ToString), RateRound)), SumRound)
                                'ExRate = Math.Round(GetTransaction_ExchangeRate(InDT.Rows(i)("object").ToString, InDT.Rows(i)("transid").ToString), RateRound)
                                'BranchTotal = Math.Round(BranchTotal + (CDbl(InDT.Rows(i)("paytot").ToString) * Math.Round(ExRate, RateRound)), SumRound)  'CDbl(objRs.Fields.Item(0).Value.ToString)
                            End If
                        End If
                    Next
                    forex = Math.Round(BranchTotal - CurTot, SumRound)
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
                            objjournalentry.Reference = "In Payment JE"
                            objjournalentry.Memo = "Posted thro' recon On: " & Now.ToString
                            objjournalentry.UserFields.Fields.Item("U_PayInNo").Value = EditText2.Value

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
                            'Amount = IIf(JEAmount < 0, -JEAmount, JEAmount)
                            Amount = IIf(BranchTotal < 0, -BranchTotal, BranchTotal)
                            GLCode = objaddon.objglobalmethods.getSingleValue("select ""PmtClrAct"" ""ControlAccount"" from OBPL where ""BPLId""='" & Branch & "'")
                            objjournalentry.Lines.AccountCode = GLCode
                            If forex <> 0 Then
                                If BranchTotal < 0 Then objjournalentry.Lines.Credit = -CurTot Else objjournalentry.Lines.Debit = CurTot
                            Else
                                If BranchTotal < 0 Then objjournalentry.Lines.Credit = Amount Else objjournalentry.Lines.Debit = Amount
                            End If
                            objjournalentry.Lines.BPLID = Branch
                            objjournalentry.Lines.Add()
                            objjournalentry.Lines.ShortName = CardCode
                            If BranchTotal < 0 Then objjournalentry.Lines.Debit = Amount Else objjournalentry.Lines.Credit = Amount

                            objjournalentry.Lines.BPLID = Branch
                            objjournalentry.Lines.Add()
                            'If forex <> 0 Then
                            '    If forex < 0 Then objjournalentry.Lines.AccountCode = Forexgain Else objjournalentry.Lines.AccountCode = Forexloss
                            '    If forex < 0 Then objjournalentry.Lines.Credit = -forex Else objjournalentry.Lines.Debit = forex
                            '    objjournalentry.Lines.BPLID = Branch
                            '    objjournalentry.Lines.Add()
                            'End If
                            If forex <> 0 Then
                                forexAmt = IIf(forex < 0, -forex, forex)
                                If forex < 0 Then objjournalentry.Lines.AccountCode = Forexgain Else objjournalentry.Lines.AccountCode = Forexloss
                                If forex < 0 Then objjournalentry.Lines.Credit = forexAmt Else objjournalentry.Lines.Debit = forexAmt
                                objjournalentry.Lines.BPLID = Branch
                                objjournalentry.Lines.Add()
                            End If
                            If objjournalentry.Add <> 0 Then
                                'If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                objaddon.objapplication.SetStatusBarMessage("Journal: " & GetBranchName(Branch) & "-" & objaddon.objcompany.GetLastErrorDescription & "-" & objaddon.objcompany.GetLastErrorCode, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(objjournalentry)
                                Return False
                            Else
                                'If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                TransId = objaddon.objcompany.GetNewObjectKey()
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(objjournalentry)
                                Matrix0.Columns.Item("jeno").Cells.Item(CInt(MatLine)).Specific.String = TransId
                                If Matrix0.Columns.Item("jeno").Cells.Item(CInt(MatLine)).Specific.String <> "" And Matrix0.Columns.Item("recono").Cells.Item(CInt(MatLine)).Specific.String = "" Then
                                    If Branchwise_InternalReconciliation(InDT, MatLine, TransId, Branch) = False Then
                                        objaddon.objapplication.StatusBar.SetText("Branchwise_InternalReco " & GetBranchName(Branch), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        Return False
                                    End If
                                End If
                                objform.Update()
                                objaddon.objapplication.SetStatusBarMessage("Branchwise JE Created Successfully... " & GetBranchName(Branch), SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                                Return True
                            End If

                        Catch ex As Exception
                            'If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            objaddon.objapplication.SetStatusBarMessage("Branchwise JE Posting Error " & GetBranchName(Branch) & "-" & objaddon.objcompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                        End Try
                    Else
                        Try
                            If Not InDT.Rows(DTLine)("jeno").ToString = String.Empty And (InDT.Rows(DTLine)("recono").ToString = String.Empty Or InDT.Rows(DTLine)("recono").ToString = "0") Then
                                If Branchwise_InternalReconciliation(InDT, MatLine, TransId, Branch) = False Then
                                    objaddon.objapplication.StatusBar.SetText("Branchwise_InternalReco" & GetBranchName(Branch), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    Return False
                                End If
                            End If
                            Return True
                        Catch ex As Exception
                            objaddon.objapplication.SetStatusBarMessage("JE Rec " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                        End Try
                    End If
                Catch ex As Exception
                    'If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                    Return False
                End Try
            Catch ex As Exception
                'If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                objaddon.objapplication.SetStatusBarMessage("JE " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                Return False
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End Function

        Private Function Branchwise_InternalReconciliation(ByVal InDT As DataTable, ByVal Line As Integer, ByVal transid As String, ByVal Branch As String) As Boolean
            Try
                Dim service As IInternalReconciliationsService = objaddon.objcompany.GetCompanyService().GetBusinessService(ServiceTypes.InternalReconciliationsService)
                Dim openTrans As InternalReconciliationOpenTrans = service.GetDataInterface(InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationOpenTrans)
                Dim reconParams As IInternalReconciliationParams = service.GetDataInterface(InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationParams)
                openTrans.CardOrAccount = CardOrAccountEnum.coaCard
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim RecAmount, ExRate As Double
                Dim Row As Integer = 0
                Dim GetStat As String
                Dim objrecset As SAPbobsCOM.Recordset
                objrecset = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                GetStat = objaddon.objglobalmethods.getSingleValue("SELECT CASE WHEN T0.""DebCred"" = 'C' AND T0.""BalDueCred"" = 0 THEN 'Reconciled' WHEN T0.""DebCred"" = 'C' AND T0.""Credit"" = T0.""BalDueCred"" THEN 'Unreconciled' " &
                                                                   " WHEN T0.""DebCred"" = 'C' AND T0.""Credit"" <> T0.""BalDueCred"" THEN 'Partial' WHEN T0.""DebCred"" = 'D' AND T0.""BalDueDeb"" = 0 THEN 'Reconciled' " &
                                                                   " WHEN T0.""DebCred"" = 'D' AND T0.""Debit"" = T0.""BalDueDeb"" THEN 'Unreconciled' WHEN T0.""DebCred"" = 'D' AND T0.""Debit"" <> T0.""BalDueDeb"" THEN 'Partial' " &
                                                                   " END ""Status"" FROM JDT1 T0 INNER JOIN OJDT T1 ON T0.""TransId"" = T1.""TransId"" WHERE T0.""TransId"" = " & transid & " " &
                                                                   "AND T0.""ShortName"" In (Select ""CardCode"" from OCRD)")
                If GetStat = "Reconciled" Then Return True
                objRs.DoQuery("select CASE WHEN T1.""BalDueCred""<>0  THEN  T1.""BalDueCred"" ELSE T1.""BalDueDeb"" END AS ""Balance"",T1.""Line_ID"" from OJDT T0 join JDT1 T1 ON T0.""TransId""=T1.""TransId"" where  T1.""TransId""='" & transid & "' and T1.""ShortName"" In (Select ""CardCode"" from OCRD)")
                Dim DocDate As Date = Date.ParseExact(EditText0.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                openTrans.ReconDate = DocDate
                If objRs.RecordCount > 0 Then
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
                        If InDT.Rows(DTRow)("doccur").ToString = MainCurr Then
                            If InDT.Rows(DTRow)("debcred").ToString = "C" Then RecAmount = Math.Round(-CDbl(InDT.Rows(DTRow)("paytot").ToString), SumRound) Else RecAmount = Math.Round(CDbl(InDT.Rows(DTRow)("paytot").ToString), SumRound)
                        Else
                            If CDbl(InDT.Rows(DTRow)("paytot").ToString) = CDbl(InDT.Rows(DTRow)("baldue").ToString) Then
                                strSQL = "Select Case when T1.""BalDueCred"" <>0 Then T1.""BalDueCred"" Else T1.""BalDueDeb"" End as ""DocTotal"",T1.""DebCred"" from OJDT T0 join JDT1 T1 on T1.""TransId"" =T0.""TransId"" where T1.""TransId""='" & InDT.Rows(DTRow)("transid").ToString & "' "
                                strSQL += vbCrLf + "And T1.""ShortName"" in (select ""CardCode"" from OCRD)"
                                objrecset.DoQuery(strSQL)
                                If objrecset.RecordCount > 0 Then
                                    RecAmount = Math.Round(CDbl(objrecset.Fields.Item(0).Value.ToString), SumRound)
                                End If
                            Else
                                ExRate = Math.Round(GetTransaction_ExchangeRate(InDT.Rows(DTRow)("object").ToString, InDT.Rows(DTRow)("transid").ToString), RateRound)
                                If InDT.Rows(DTRow)("debcred").ToString = "C" Then RecAmount = Math.Round(ExRate * -CDbl(InDT.Rows(DTRow)("paytot").ToString), SumRound) Else RecAmount = Math.Round(ExRate * CDbl(InDT.Rows(DTRow)("paytot").ToString), SumRound)
                                'CurTot = Math.Round(CurTot + (CDbl(InDT.Rows(DTRow)("paytot").ToString) * Math.Round(CDbl(objRs.Fields.Item(0).Value.ToString), RateRound)), SumRound)
                                'ExRate = Math.Round(GetTransaction_ExchangeRate(InDT.Rows(DTRow)("object").ToString, InDT.Rows(DTRow)("transid").ToString), RateRound)
                                'BranchTotal = Math.Round(BranchTotal + (CDbl(InDT.Rows(DTRow)("paytot").ToString) * Math.Round(ExRate, RateRound)), SumRound) 'CDbl(objRs.Fields.Item(0).Value.ToString)
                            End If
                            'ExRate = Math.Round(GetTransaction_ExchangeRate(InDT.Rows(DTRow)("object").ToString, InDT.Rows(DTRow)("transid").ToString), RateRound)
                            'If InDT.Rows(DTRow)("debcred").ToString = "C" Then RecAmount = Math.Round(ExRate * -CDbl(InDT.Rows(DTRow)("paytot").ToString), SumRound) Else RecAmount = Math.Round(ExRate * CDbl(InDT.Rows(DTRow)("paytot").ToString), SumRound)
                        End If
                        'If InDT.Rows(DTRow)("debcred").ToString = "C" Then RecAmount = -CDbl(InDT.Rows(DTRow)("paytot").ToString) Else RecAmount = CDbl(InDT.Rows(DTRow)("paytot").ToString)
                        openTrans.InternalReconciliationOpenTransRows.Add()
                        openTrans.InternalReconciliationOpenTransRows.Item(Row).Selected = BoYesNoEnum.tYES
                        openTrans.InternalReconciliationOpenTransRows.Item(Row).TransId = InDT.Rows(DTRow)("transid").ToString
                        openTrans.InternalReconciliationOpenTransRows.Item(Row).TransRowId = InDT.Rows(DTRow)("tranline").ToString
                        openTrans.InternalReconciliationOpenTransRows.Item(Row).ReconcileAmount = RecAmount ' CDbl(InDT.Rows(DTRow)("paytot").ToString) '
                        Row += 1
                    End If
                Next
                Dim Reconum As Integer = 0
                Try
                    reconParams = service.Add(openTrans)
                Catch ex As Exception
                    objaddon.objapplication.StatusBar.SetText("Branchwise_InternReco Error : " & GetBranchName(Branch) & "-" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : Return False
                End Try
                Reconum = reconParams.ReconNum
                If Reconum = 0 Then objaddon.objapplication.StatusBar.SetText("Branchwise Reco Error..." & GetBranchName(Branch), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : Return False

                Matrix0.Columns.Item("recono").Cells.Item(Line).Specific.String = Reconum
                objaddon.objapplication.StatusBar.SetText("Branchwise Reconciled successfully..." & GetBranchName(Branch), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(openTrans)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(reconParams)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(service)
                GC.Collect()
                Return True
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("Branchwise_InternReco Reco: " & GetBranchName(Branch) & "-" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End Try

        End Function

        Private Sub EditText0_ValidateAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText0.ValidateAfter
            Try
                If pVal.ItemChanged = False Then Exit Sub
                Calc_Total(-1)
                DocumentDate = Date.ParseExact(EditText0.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                Clear_Payments()
            Catch ex As Exception

            End Try

        End Sub

        Private Function GetBranchName(ByVal BCode As String) As String
            Try
                Dim BName As String
                BName = objaddon.objglobalmethods.getSingleValue("Select ""BPLName"" from OBPL where ""BPLId""='" & BCode & "' ")

                Return BName
            Catch ex As Exception
                Return 0
            End Try
        End Function

        Private Function IncomingPayment_Consolidated(ByVal InDT As DataTable, ByVal LineNum As Integer, ByVal Branch As String, ByVal CardCode As String, ByVal Amount As Double) As Boolean
            Try
                Dim objIncom As SAPbobsCOM.Payments
                Dim DocEntry, Series As String
                Dim Total As Double
                For i As Integer = 0 To InDT.Rows.Count - 1
                    If InDT.Rows(i)("branchc").ToString = CardCode Then
                        If InDT.Rows(i)("doccur").ToString = MainCurr Then
                            Total = Math.Round(Total + CDbl(InDT.Rows(i)("paytot").ToString), SumRound)
                        Else
                            strSQL = "Select ""Rate"",""Currency"" from ORTT where ""RateDate""= '" & DocumentDate.ToString("yyyyMMdd") & "' and ""Currency""='" & InDT.Rows(i)("doccur").ToString & "' " ' InDT.Rows(i)("date").ToString
                            objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            objRs.DoQuery(strSQL)
                            Total = Math.Round(Total + (CDbl(InDT.Rows(i)("paytot").ToString) * Math.Round(CDbl(objRs.Fields.Item(0).Value.ToString), RateRound)), SumRound)
                            'ExcRate = Math.Round(GetTransaction_ExchangeRate(InDT.Rows(i)("object").ToString, InDT.Rows(i)("transid").ToString), RateRound)
                            'BranchTotal = Math.Round(BranchTotal + (CDbl(InDT.Rows(i)("paytot").ToString) * ExcRate), SumRound)
                        End If
                    End If
                Next

                objIncom = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
                objIncom.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_IncomingPayments
                objIncom.DocType = SAPbobsCOM.BoRcptTypes.rCustomer
                objIncom.CardCode = CardCode
                objIncom.BPLID = Branch
                objIncom.Remarks = "In Payment add-on"
                objIncom.UserFields.Fields.Item("U_PayInNo").Value = EditText2.Value
                Dim DocDate As Date = Date.ParseExact(EditText0.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                objIncom.DocDate = DocDate
                objIncom.DocCurrency = EditText18.Value 'Matrix0.Columns.Item("doccur").Cells.Item(LineNum).Specific.String
                If Val(EditText21.Value) > 0 Then objIncom.DocRate = CDbl(EditText21.Value)
                'objIncom.LocalCurrency = BoYesNoEnum.tYES
                If Localization = "IN" Then
                    If objaddon.HANA Then
                        Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 ""Series"" FROM NNM1 WHERE ""ObjectCode""='24' and ""Indicator""=(Select ""Indicator"" from OFPR where '" & DocDate.ToString("yyyyMMdd") & "' Between ""F_RefDate"" and ""T_RefDate"") " &
                                                                                      " and Ifnull(""Locked"",'')='N' and ""BPLId""='" & Branch & "'")
                    Else
                        Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 Series FROM NNM1 WHERE ObjectCode='24' and Indicator=(Select Indicator from OFPR where '" & DocDate.ToString("yyyyMMdd") & "' Between F_RefDate and T_RefDate) " &
                                                                                      " and Isnull(Locked,'')='N' and BPLId='" & Branch & "'")
                    End If
                Else
                    If objaddon.HANA Then
                        Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 ""Series"" FROM NNM1 WHERE ""ObjectCode""='24' and ""Indicator""=(Select ""Indicator"" from OFPR where '" & DocDate.ToString("yyyyMMdd") & "' Between ""F_RefDate"" and ""T_RefDate"") " &
                                                                                      " and Ifnull(""Locked"",'')='N' and ""BPLId""='" & Branch & "'")
                    Else
                        Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 Series FROM NNM1 WHERE ObjectCode='24' and Indicator=(Select Indicator from OFPR where '" & DocDate.ToString("yyyyMMdd") & "' Between F_RefDate and T_RefDate) " &
                                                                                      " and Isnull(Locked,'')='N' and BPLId='" & Branch & "'")
                    End If
                End If
                If Series <> "" Then objIncom.Series = Series
                If Matrix1.VisualRowCount > 0 Then
                    If Val(Matrix1.Columns.Item("chamt").Cells.Item(1).Specific.String) > 0 Then 'Cheque
                        objIncom.CheckAccount = EditText10.Value
                        For Row As Integer = 1 To Matrix1.VisualRowCount
                            If Val(Matrix1.Columns.Item("chamt").Cells.Item(Row).Specific.String) > 0 Then
                                If Matrix1.Columns.Item("chnum").Cells.Item(Row).Specific.String <> "" Then objIncom.Checks.CheckNumber = Trim(Matrix1.Columns.Item("chnum").Cells.Item(Row).Specific.String)
                                objIncom.Checks.CheckAccount = EditText10.Value ' Matrix1.Columns.Item("chgl").Cells.Item(Row).Specific.String
                                objIncom.Checks.CheckSum = Math.Round(CDbl(Matrix1.Columns.Item("chamt").Cells.Item(Row).Specific.String), SumRound)
                                Dim oedit As SAPbouiCOM.EditText
                                oedit = Matrix1.Columns.Item("chdate").Cells.Item(Row).Specific
                                Dim chDate As Date = Date.ParseExact(oedit.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                                objIncom.Checks.DueDate = chDate ' Matrix1.Columns.Item("chdate").Cells.Item(Row).Specific.String
                                If Matrix1.Columns.Item("chcty").Cells.Item(Row).Specific.String <> "" Then objIncom.Checks.CountryCode = Matrix1.Columns.Item("chcty").Cells.Item(Row).Specific.String
                                If Matrix1.Columns.Item("chbank").Cells.Item(Row).Specific.String <> "" Then objIncom.Checks.BankCode = Matrix1.Columns.Item("chbank").Cells.Item(Row).Specific.String
                                If Matrix1.Columns.Item("chbranch").Cells.Item(Row).Specific.String <> "" Then objIncom.Checks.Branch = Trim(Matrix1.Columns.Item("chbranch").Cells.Item(Row).Specific.String)
                                If Matrix1.Columns.Item("chact").Cells.Item(Row).Specific.String <> "" Then objIncom.Checks.AccounttNum = Trim(Matrix1.Columns.Item("chact").Cells.Item(Row).Specific.String)
                                If Matrix1.Columns.Item("chissue").Cells.Item(Row).Specific.String <> "" Then objIncom.Checks.OriginallyIssuedBy = Matrix1.Columns.Item("chissue").Cells.Item(Row).Specific.String
                                If Matrix1.Columns.Item("chfiscal").Cells.Item(Row).Specific.String <> "" Then objIncom.Checks.FiscalID = Trim(Matrix1.Columns.Item("chfiscal").Cells.Item(Row).Specific.String)
                                If Matrix1.Columns.Item("chendor").Cells.Item(Row).Specific.String = "Y" Then
                                    objIncom.Checks.Trnsfrable = BoYesNoEnum.tYES
                                Else
                                    objIncom.Checks.Trnsfrable = BoYesNoEnum.tNO
                                End If
                                objIncom.Checks.Add()
                            End If
                        Next
                    End If
                End If
                If Val(EditText17.Value) > 0 Then 'BCG Amount
                    objIncom.BankAccount = EditText16.Value
                    objIncom.BankChargeAmount = Math.Round(CDbl(EditText17.Value), SumRound)
                End If
                If Val(EditText5.Value) > 0 Then 'Transfer
                    objIncom.TransferAccount = EditText6.Value
                    objIncom.TransferSum = Math.Round(CDbl(EditText5.Value), SumRound)
                    objIncom.TransferDate = Date.ParseExact(EditText4.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                    objIncom.TransferReference = EditText7.Value
                End If
                If Matrix2.VisualRowCount > 0 Then
                    'Dim DocDate As Date = Date.ParseExact(objEdit.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                    If Val(Matrix2.Columns.Item("amtdue").Cells.Item(1).Specific.String) > 0 Then  'Card
                        objIncom.CreditCards.CreditCard = Matrix2.Columns.Item("cardname").Cells.Item(1).Specific.String
                        objIncom.CreditCards.CreditCardNumber = Matrix2.Columns.Item("cardno").Cells.Item(1).Specific.String
                        objIncom.CreditCards.CreditAcct = Matrix2.Columns.Item("cardgl").Cells.Item(1).Specific.String
                        Dim GetDate As String = objaddon.objglobalmethods.getSingleValue("SELECT LAST_DAY (TO_DATE('" & Matrix2.Columns.Item("valid").Cells.Item(1).Specific.String & "', 'MM/YY')) ""last day"" FROM DUMMY")
                        objIncom.CreditCards.CardValidUntil = CDate(GetDate)
                        If Matrix2.Columns.Item("idno").Cells.Item(1).Specific.String <> "" Then objIncom.CreditCards.OwnerIdNum = Matrix2.Columns.Item("idno").Cells.Item(1).Specific.String
                        If Matrix2.Columns.Item("telno").Cells.Item(1).Specific.String <> "" Then objIncom.CreditCards.OwnerPhone = Matrix2.Columns.Item("telno").Cells.Item(1).Specific.String
                        objIncom.CreditCards.CreditSum = Math.Round(CDbl(Matrix2.Columns.Item("amtdue").Cells.Item(1).Specific.String), SumRound)
                        If Matrix2.Columns.Item("appcode").Cells.Item(1).Specific.String <> "" Then objIncom.CreditCards.VoucherNum = Matrix2.Columns.Item("appcode").Cells.Item(1).Specific.String
                        If Matrix2.Columns.Item("trantype").Cells.Item(1).Specific.String = "I" Then
                            objIncom.CreditCards.CreditType = BoRcptCredTypes.cr_InternetTransaction
                        ElseIf Matrix2.Columns.Item("trantype").Cells.Item(1).Specific.String = "S" Then
                            objIncom.CreditCards.CreditType = BoRcptCredTypes.cr_Regular
                        Else
                            objIncom.CreditCards.CreditType = BoRcptCredTypes.cr_Telephone
                        End If
                        objIncom.CreditCards.Add()
                    End If
                End If
                If Val(EditText9.Value) > 0 Then 'Cash
                    objIncom.CashAccount = EditText8.Value
                    objIncom.CashSum = Math.Round(CDbl(EditText9.Value), SumRound)
                End If

                For i As Integer = 0 To InDT.Rows.Count - 1
                    If CDbl(InDT.Rows(i)("paytot").ToString) <> 0 And InDT.Rows(i)("cardc").ToString = CurBPCode And InDT.Rows(i)("branchc").ToString = CurBranch Then
                        objIncom.Invoices.TotalDiscount = CDbl(InDT.Rows(i)("cashdisc").ToString)
                        If InDT.Rows(i)("doccur").ToString = MainCurr Then
                            objIncom.Invoices.SumApplied = Math.Round(CDbl(InDT.Rows(i)("paytot").ToString), SumRound) ' InvTotal
                        Else
                            objIncom.Invoices.AppliedFC = Math.Round(CDbl(InDT.Rows(i)("paytot").ToString), SumRound) '/ CDbl(objRs.Fields.Item(0).Value.ToString)
                        End If
                        objIncom.Invoices.DocEntry = CInt(InDT.Rows(i)("docentry").ToString)
                        If InDT.Rows(i)("objtype").ToString = "IN" Then
                            objIncom.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_Invoice
                        ElseIf InDT.Rows(i)("objtype").ToString = "CN" Then
                            objIncom.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_CredItnote
                        ElseIf InDT.Rows(i)("objtype").ToString = "JE" Then
                            objIncom.Invoices.DocLine = CInt(InDT.Rows(i)("tranline").ToString)
                            objIncom.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_JournalEntry
                        End If

                        If InDT.Rows(i)("cc1").ToString <> "" Then objIncom.Invoices.DistributionRule = InDT.Rows(i)("cc1").ToString
                        If InDT.Rows(i)("cc2").ToString <> "" Then objIncom.Invoices.DistributionRule2 = InDT.Rows(i)("cc2").ToString
                        If InDT.Rows(i)("cc3").ToString <> "" Then objIncom.Invoices.DistributionRule3 = InDT.Rows(i)("cc3").ToString
                        If InDT.Rows(i)("cc4").ToString <> "" Then objIncom.Invoices.DistributionRule4 = InDT.Rows(i)("cc4").ToString
                        If InDT.Rows(i)("cc5").ToString <> "" Then objIncom.Invoices.DistributionRule5 = InDT.Rows(i)("cc5").ToString
                        objIncom.Invoices.Add()
                    End If
                Next
                Dim ret As Long
                ret = objIncom.Add()
                If ret <> 0 Then
                    'If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    objaddon.objapplication.StatusBar.SetText("Incoming Payment: Branch: " & GetBranchName(Branch) & " BPCode: " & CardCode & "-" & objaddon.objcompany.GetLastErrorDescription & "-" & objaddon.objcompany.GetLastErrorCode & " on Line: " & LineNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objIncom)
                    Return False
                Else
                    'If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    DocEntry = objaddon.objcompany.GetNewObjectKey()
                    EditText13.Value = DocEntry
                    objaddon.objapplication.StatusBar.SetText("Incoming Payment successfully created..." & GetBranchName(Branch) & " BPCode: " & CardCode, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objIncom)
                    GC.Collect()
                    If JournalEntry_BranchTransfer(objFinalDT, CurBranch, CurBPCode) = False Then
                        objaddon.objapplication.StatusBar.SetText("JE_BranchTransfer", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        Return False
                    End If
                    Return True
                End If

            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("Incoming Payment: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End Try
        End Function

        Private WithEvents StaticText34 As SAPbouiCOM.StaticText
        Private WithEvents EditText29 As SAPbouiCOM.EditText

#End Region

    End Class
End Namespace
