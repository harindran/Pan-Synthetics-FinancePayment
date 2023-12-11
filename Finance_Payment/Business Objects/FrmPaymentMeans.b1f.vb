Option Strict Off
Option Explicit On

Imports System.Drawing
Imports System.Text.RegularExpressions
Imports SAPbobsCOM
Imports SAPbouiCOM.Framework
Imports System.Linq
Imports System.Runtime.CompilerServices

Namespace Finance_Payment
    <FormAttribute("PAYM", "Business Objects/FrmPaymentMeans.b1f")>
    Friend Class FrmPaymentMeans
        Inherits UserFormBase
        Public WithEvents objform As SAPbouiCOM.Form
        Public WithEvents objPayform As SAPbouiCOM.Form
        Private WithEvents pCFL As SAPbouiCOM.ISBOChooseFromListEventArg
        Dim FormCount As Integer = 0
        Dim objRs As SAPbobsCOM.Recordset
        Dim strSQL As String
        Private WithEvents objMatrix As SAPbouiCOM.Matrix
        Private WithEvents objCombo As SAPbouiCOM.ComboBox
        Dim objDT As New DataTable
        Private WithEvents cmbcol As SAPbouiCOM.Column
        Private WithEvents oEdit As SAPbouiCOM.EditText
        Dim RowID As Integer = 0
        Dim _BT, _CH, _CR, _CA As Boolean 'Bank,Cheque,Credit, Cash
        Public FindPayment As String

        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("101").Specific, SAPbouiCOM.Button)
            Me.Button1 = CType(Me.GetItem("2").Specific, SAPbouiCOM.Button)
            Me.StaticText0 = CType(Me.GetItem("lcurr").Specific, SAPbouiCOM.StaticText)
            Me.ComboBox0 = CType(Me.GetItem("tcurr").Specific, SAPbouiCOM.ComboBox)
            Me.Folder0 = CType(Me.GetItem("3").Specific, SAPbouiCOM.Folder)
            Me.Folder1 = CType(Me.GetItem("4").Specific, SAPbouiCOM.Folder)
            Me.Folder2 = CType(Me.GetItem("5").Specific, SAPbouiCOM.Folder)
            Me.Folder3 = CType(Me.GetItem("6").Specific, SAPbouiCOM.Folder)
            Me.StaticText1 = CType(Me.GetItem("l_oamt").Specific, SAPbouiCOM.StaticText)
            Me.StaticText2 = CType(Me.GetItem("l_baldue").Specific, SAPbouiCOM.StaticText)
            Me.EditText0 = CType(Me.GetItem("toamt").Specific, SAPbouiCOM.EditText)
            Me.EditText1 = CType(Me.GetItem("tbaldue").Specific, SAPbouiCOM.EditText)
            Me.StaticText3 = CType(Me.GetItem("lpaid").Specific, SAPbouiCOM.StaticText)
            Me.EditText2 = CType(Me.GetItem("tpaid").Specific, SAPbouiCOM.EditText)
            Me.StaticText4 = CType(Me.GetItem("l_bankchar").Specific, SAPbouiCOM.StaticText)
            Me.EditText3 = CType(Me.GetItem("tbankchar").Specific, SAPbouiCOM.EditText)
            Me.StaticText15 = CType(Me.GetItem("lctot").Specific, SAPbouiCOM.StaticText)
            Me.EditText11 = CType(Me.GetItem("tctot").Specific, SAPbouiCOM.EditText)
            Me.StaticText17 = CType(Me.GetItem("lbtot").Specific, SAPbouiCOM.StaticText)
            Me.EditText13 = CType(Me.GetItem("tbtot").Specific, SAPbouiCOM.EditText)
            Me.StaticText18 = CType(Me.GetItem("lbgl").Specific, SAPbouiCOM.StaticText)
            Me.EditText14 = CType(Me.GetItem("tbgl").Specific, SAPbouiCOM.EditText)
            Me.LinkedButton2 = CType(Me.GetItem("lkbgl").Specific, SAPbouiCOM.LinkedButton)
            Me.StaticText19 = CType(Me.GetItem("lbdate").Specific, SAPbouiCOM.StaticText)
            Me.EditText15 = CType(Me.GetItem("tbdate").Specific, SAPbouiCOM.EditText)
            Me.StaticText20 = CType(Me.GetItem("lbref").Specific, SAPbouiCOM.StaticText)
            Me.EditText16 = CType(Me.GetItem("tbref").Specific, SAPbouiCOM.EditText)
            Me.StaticText21 = CType(Me.GetItem("lbgln").Specific, SAPbouiCOM.StaticText)
            Me.StaticText22 = CType(Me.GetItem("lchgl").Specific, SAPbouiCOM.StaticText)
            Me.EditText17 = CType(Me.GetItem("tchgl").Specific, SAPbouiCOM.EditText)
            Me.LinkedButton3 = CType(Me.GetItem("lkchgl").Specific, SAPbouiCOM.LinkedButton)
            Me.StaticText23 = CType(Me.GetItem("tchgln").Specific, SAPbouiCOM.StaticText)
            Me.Matrix0 = CType(Me.GetItem("mtxcheq").Specific, SAPbouiCOM.Matrix)
            Me.StaticText24 = CType(Me.GetItem("lcrtot").Specific, SAPbouiCOM.StaticText)
            Me.EditText18 = CType(Me.GetItem("tcrtot").Specific, SAPbouiCOM.EditText)
            Me.StaticText25 = CType(Me.GetItem("lcgl").Specific, SAPbouiCOM.StaticText)
            Me.EditText19 = CType(Me.GetItem("tcgl").Specific, SAPbouiCOM.EditText)
            Me.LinkedButton4 = CType(Me.GetItem("lkcgl").Specific, SAPbouiCOM.LinkedButton)
            Me.StaticText26 = CType(Me.GetItem("lcgln").Specific, SAPbouiCOM.StaticText)
            Me.Matrix1 = CType(Me.GetItem("mtxcr").Specific, SAPbouiCOM.Matrix)
            Me.StaticText5 = CType(Me.GetItem("lrate").Specific, SAPbouiCOM.StaticText)
            Me.EditText4 = CType(Me.GetItem("trate").Specific, SAPbouiCOM.EditText)
            Me.EditText5 = CType(Me.GetItem("toamtc").Specific, SAPbouiCOM.EditText)
            Me.EditText6 = CType(Me.GetItem("tbalduec").Specific, SAPbouiCOM.EditText)
            Me.StaticText6 = CType(Me.GetItem("ltrantot").Specific, SAPbouiCOM.StaticText)
            Me.EditText7 = CType(Me.GetItem("ttrantot").Specific, SAPbouiCOM.EditText)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()

        End Sub

        Private Sub OnCustomInitialize()
            Try
                Dim Flag As Boolean = False
                objform = objaddon.objapplication.Forms.GetForm("PAYM", Me.FormCount)
                'objform = objaddon.objapplication.Forms.ActiveForm
                'objform.Freeze(True)
                Try
                    objform.EnableMenu("1281", False)
                    objform.EnableMenu("1282", False)
                Catch ex As Exception
                End Try
                pModal = True
                StaticText6.Item.Visible = False
                EditText7.Item.Visible = False
                Matrix1.Layout = SAPbouiCOM.BoMatrixLayoutType.mlt_VerticalStaticTitle
                objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "chamt", "#")
                objaddon.objglobalmethods.Matrix_Addrow(Matrix1, "amtdue", "#")
                Matrix1.Columns.Item("trantype").ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                Matrix1.Columns.Item("trantype").Cells.Item(1).Specific.Select("S", SAPbouiCOM.BoSearchKey.psk_ByValue)
                'Matrix0.Columns.Item("chbank").ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                Matrix0.Columns.Item("chendor").ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                Matrix0.Columns.Item("chamt").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                Matrix1.AutoResizeColumns()

                LoadCombo("select distinct T0.""CreditCard"",T0.""CardName"" from OCRC T0 order by T0.""CardName""", "cardname")
                'LoadCombo("select distinct T0.""CountryCod"",T1.""Name"" from ODSC T0 join OCRY T1 on T0.""CountryCod""=T1.""Code""", "chcty")

                objPayform = objaddon.objapplication.Forms.GetForm("PAYINIT", 1)
                objMatrix = objPayform.Items.Item("mtxdata").Specific

                For i As Integer = 1 To objMatrix.VisualRowCount
                    If objMatrix.Columns.Item("bpcur").Cells.Item(i).Specific.String <> "##" And objMatrix.Columns.Item("bpcode").Cells.Item(i).Specific.String <> "" Then
                        Flag = True
                        Exit For
                    End If
                Next
                Folder0.Item.Click()
                If objPayform.Items.Item("opincpay").Specific.Selected = True Then
                    objPayform = objaddon.objapplication.Forms.GetForm("FINPAY", 1)
                    Matrix0.Columns.Item("glacc").Visible = False
                    FindPayment = "IN"
                    Matrix0.Columns.Item("chbranch").Visible = False
                    Matrix0.Columns.Item("chact").Visible = False
                    Matrix0.Columns.Item("chendor").Visible = False
                    Matrix0.Columns.Item("chissue").Visible = False
                    Matrix0.Columns.Item("chfiscal").Visible = False
                    'Matrix2.Item.Visible = False
                    'Matrix0.Item.Left = 7
                    'Matrix0.Item.Top = 82
                    'Matrix0.Item.Height = 181
                    'Matrix0.Item.Width = 488
                    'GetData_Payment("FINPAY")
                Else 'If objPayform.Items.Item("opoutpay").Specific.Selected = True Then
                    objPayform = objaddon.objapplication.Forms.GetForm("FOUTPAY", 1)
                    StaticText22.Item.Visible = False
                    LinkedButton3.Item.Visible = False
                    EditText17.Item.Visible = False
                    StaticText23.Item.Visible = False
                    FindPayment = "OUT"
                    Matrix0.Columns.Item("chissue").Visible = False
                    Matrix0.Columns.Item("chfiscal").Visible = False
                    'Matrix0.Item.Visible = False
                    'Matrix2.Item.Left = 7
                    'Matrix2.Item.Top = 82
                    'Matrix2.Item.Height = 181
                    'Matrix2.Item.Width = 488
                    'GetData_Payment("FOUTPAY")
                End If

                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRs.DoQuery("select distinct T0.""BnkChgAct"" as ""BCGAcct"",T0.""LinkAct_3"" as ""CahAcct"",""LinkAct_24"" as ""Rounding"",""GLGainXdif"",""GLLossXdif"",""ExDiffAct"" " &
                              ",(Select ""SumDec"" from OADM) as ""SumDec"",(Select ""RateDec"" from OADM) as ""RateDec""" &
                              "from OACP T0 left join OFPR T1 on T1.""Category""=T0.""PeriodCat"" where T0.""PeriodCat""=(Select ""Category"" from OFPR where CURRENT_DATE Between ""F_RefDate"" and ""T_RefDate"")")
                If objRs.RecordCount > 0 Then
                    If objRs.Fields.Item(6).Value.ToString <> "" Then SumRound = objRs.Fields.Item(6).Value.ToString
                    If objRs.Fields.Item(7).Value.ToString <> "" Then RateRound = objRs.Fields.Item(7).Value.ToString
                End If
                objRs.DoQuery("select T0.""CurrCode"",T0.""CurrName"" from OCRN T0 order by T0.""CurrName""")
                If objRs.RecordCount > 0 Then
                    For i As Integer = 0 To objRs.RecordCount - 1
                        Try
                            ComboBox0.ValidValues.Add(objRs.Fields.Item(0).Value.ToString, objRs.Fields.Item(1).Value.ToString)
                            objRs.MoveNext()
                        Catch ex As Exception
                            objRs.MoveNext()
                        End Try
                    Next
                    If objPayform.Items.Item("tcurr").Specific.String = "" Then
                        strSQL = objaddon.objglobalmethods.getSingleValue("select ""MainCurncy"" from OADM")
                        If strSQL <> "" Then
                            ComboBox0.Select(strSQL, SAPbouiCOM.BoSearchKey.psk_ByValue)
                        End If
                    Else
                        Dim cmbval As String = Trim(objPayform.Items.Item("tcurr").Specific.String)
                        ComboBox0.Select(cmbval, SAPbouiCOM.BoSearchKey.psk_ByValue)
                    End If
                    If Flag Then Folder3.Item.Click() : ComboBox0.Item.Enabled = False ': objform.ActiveItem = "tctot"
                End If

                objPayform = objaddon.objapplication.Forms.GetForm("PAYINIT", 1)
                If objPayform.Items.Item("opincpay").Specific.Selected = True Then
                    GetData_Payment("FINPAY")
                Else 'If objPayform.Items.Item("opoutpay").Specific.Selected = True Then
                    GetData_Payment("FOUTPAY")
                End If

                Calc_Total(objPayDT)
                Folder2.Item.Enabled = False
                objform.Settings.Enabled = True
                'objform.Freeze(False)
            Catch ex As Exception
                'objaddon.objapplication.StatusBar.SetText("F-" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                'objform.Freeze(False)
            End Try
        End Sub

#Region "Fields"
        Private WithEvents Button0 As SAPbouiCOM.Button
        Private WithEvents Button1 As SAPbouiCOM.Button
        Private WithEvents StaticText0 As SAPbouiCOM.StaticText
        Private WithEvents ComboBox0 As SAPbouiCOM.ComboBox
        Private WithEvents Folder0 As SAPbouiCOM.Folder
        Private WithEvents Folder1 As SAPbouiCOM.Folder
        Private WithEvents Folder2 As SAPbouiCOM.Folder
        Private WithEvents Folder3 As SAPbouiCOM.Folder
        Private WithEvents StaticText1 As SAPbouiCOM.StaticText
        Private WithEvents StaticText2 As SAPbouiCOM.StaticText
        Private WithEvents EditText0 As SAPbouiCOM.EditText
        Private WithEvents EditText1 As SAPbouiCOM.EditText
        Private WithEvents StaticText3 As SAPbouiCOM.StaticText
        Private WithEvents EditText2 As SAPbouiCOM.EditText
        Private WithEvents StaticText4 As SAPbouiCOM.StaticText
        Private WithEvents EditText3 As SAPbouiCOM.EditText
        Private WithEvents StaticText15 As SAPbouiCOM.StaticText
        Private WithEvents EditText11 As SAPbouiCOM.EditText
        Private WithEvents StaticText17 As SAPbouiCOM.StaticText
        Private WithEvents EditText13 As SAPbouiCOM.EditText
        Private WithEvents StaticText18 As SAPbouiCOM.StaticText
        Private WithEvents EditText14 As SAPbouiCOM.EditText
        Private WithEvents LinkedButton2 As SAPbouiCOM.LinkedButton
        Private WithEvents StaticText19 As SAPbouiCOM.StaticText
        Private WithEvents EditText15 As SAPbouiCOM.EditText
        Private WithEvents StaticText20 As SAPbouiCOM.StaticText
        Private WithEvents EditText16 As SAPbouiCOM.EditText
        Private WithEvents StaticText21 As SAPbouiCOM.StaticText
        Private WithEvents StaticText22 As SAPbouiCOM.StaticText
        Private WithEvents EditText17 As SAPbouiCOM.EditText
        Private WithEvents LinkedButton3 As SAPbouiCOM.LinkedButton
        Private WithEvents StaticText23 As SAPbouiCOM.StaticText
        Private WithEvents Matrix0 As SAPbouiCOM.Matrix
        Private WithEvents StaticText24 As SAPbouiCOM.StaticText
        Private WithEvents EditText18 As SAPbouiCOM.EditText
        Private WithEvents StaticText25 As SAPbouiCOM.StaticText
        Private WithEvents EditText19 As SAPbouiCOM.EditText
        Private WithEvents LinkedButton4 As SAPbouiCOM.LinkedButton
        Private WithEvents StaticText26 As SAPbouiCOM.StaticText
        Private WithEvents Matrix1 As SAPbouiCOM.Matrix
        Private WithEvents EditText4 As SAPbouiCOM.EditText
        Private WithEvents EditText5 As SAPbouiCOM.EditText
        Private WithEvents EditText6 As SAPbouiCOM.EditText
        Private WithEvents StaticText5 As SAPbouiCOM.StaticText
        Private WithEvents StaticText6 As SAPbouiCOM.StaticText
        Private WithEvents EditText7 As SAPbouiCOM.EditText
#End Region

        Private Sub Button0_ClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.ClickAfter
            Try
                objform.Close()
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        End Sub

        Private Sub LoadCombo(ByVal Query As String, ByVal ColName As String)
            Try
                If ColName = "chcty" Then
                    cmbcol = Matrix0.Columns.Item(ColName)
                Else
                    cmbcol = Matrix1.Columns.Item(ColName)
                End If

                cmbcol.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRs.DoQuery(Query)
                If objRs.RecordCount > 0 Then
                    cmbcol.ValidValues.Add("-1", "")
                    For i As Integer = 0 To objRs.RecordCount - 1
                        Try
                            cmbcol.ValidValues.Add(objRs.Fields.Item(0).Value, objRs.Fields.Item(1).Value)
                            objRs.MoveNext()
                        Catch ex As Exception
                            objRs.MoveNext()
                        End Try
                    Next
                End If
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub

        Private Sub GetData_Payment(ByVal FormType As String)
            Try
                objPayform = objaddon.objapplication.Forms.GetForm(FormType, 1)

                'If objPayform.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                EditText0.Value = objPayform.Items.Item("ttotdue").Specific.String
                EditText1.Value = objPayform.Items.Item("ttotdue").Specific.String
                If Val(objPayform.Items.Item("tfc").Specific.String) <> 0 Then EditText5.Value = objPayform.Items.Item("tfc").Specific.String
                If objPayform.Items.Item("tcurr").Specific.String = MainCurr Then
                    EditText2.Value = objPayform.Items.Item("ttotdue").Specific.String
                Else
                    EditText2.Value = objPayform.Items.Item("tfc").Specific.String
                End If
                'Else
                '    EditText0.Value = objPayform.Items.Item("ttotdue").Specific.String
                '    EditText1.Value = objPayform.Items.Item("ttotdue").Specific.String
                '    EditText2.Value = objPayform.Items.Item("tfc").Specific.String
                'End If

                If objPayform.Items.Item("tcurr").Specific.String <> "" Then
                    If Val(objPayform.Items.Item("tboeamt").Specific.String) > 0 Then EditText4.Value = objPayform.Items.Item("tboeamt").Specific.String
                End If
                'Currency_FieldSetup()

                If Val(objPayform.Items.Item("tbcgtot").Specific.String) > 0 Then EditText3.Value = objPayform.Items.Item("tbcgtot").Specific.String
                'Transfer
                oEdit = objPayform.Items.Item("tbidate").Specific
                If objPayform.Items.Item("tbigl").Specific.String <> "" Then EditText14.Value = objPayform.Items.Item("tbigl").Specific.String
                If objPayform.Items.Item("tbidate").Specific.String <> "" Then EditText15.Value = oEdit.Value
                If objPayform.Items.Item("tbiref").Specific.String <> "" Then EditText16.Value = objPayform.Items.Item("tbiref").Specific.String
                If Val(objPayform.Items.Item("tbitot").Specific.String) > 0 Then EditText13.Value = objPayform.Items.Item("tbitot").Specific.String : _BT = True
                'Cash
                If objPayform.Items.Item("tcigl").Specific.String <> "" Then EditText19.Value = objPayform.Items.Item("tcigl").Specific.String Else EditText19.Value = CashAcct
                If CashAcct <> "" Then
                    Dim ActName As String = objaddon.objglobalmethods.getSingleValue("Select ""AcctName"" from OACT where ""AcctCode""='" & CashAcct & "'")
                    If ActName <> "" Then
                        StaticText26.Caption = ActName
                    End If
                End If
                If Val(objPayform.Items.Item("tcitot").Specific.String) > 0 Then EditText11.Value = objPayform.Items.Item("tcitot").Specific.String : _CA = True
                'Cheque
                If objPayform.Items.Item("tchigl").Specific.String <> "" Then EditText17.Value = objPayform.Items.Item("tchigl").Specific.String
                If Val(objPayform.Items.Item("tcrtot").Specific.String) > 0 Then EditText18.Value = objPayform.Items.Item("tcrtot").Specific.String

                If Val(objPayform.Items.Item("tchtot").Specific.String) > 0 Then
                    If Assign_DataTable_To_Matrix(objPayform, Matrix0, "mtxcheq", "chamt", False) Then
                        _CH = True
                    End If
                End If

                'Credit Card
                'If Val(objPayform.Items.Item("tcrtot").Specific.String) > 0 Then
                '    If Assign_DataTable_To_Matrix(objPayform, Matrix1, "mtxcr", "amtdue", False) Then
                '        _CR = True
                '    End If
                'End If


                If objPayform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    Folder3.Item.Click()
                    If _BT Then
                        Folder1.Item.Click()
                    End If
                    If _CH Then
                        Folder0.Item.Click()
                    End If
                    'If _CR Then
                    '    Folder2.Item.Click()
                    'End If
                    If _CA Then
                        Folder3.Item.Click()
                    End If
                Else
                    'EditText1.Value = 0
                    'EditText2.Value = objPayform.Items.Item("ttotdue").Specific.String
                    If _BT = False Then
                        Folder1.Item.Enabled = False
                    Else
                        Folder1.Item.Click()
                    End If
                    If _CH = False Then
                        Folder0.Item.Enabled = False
                    Else
                        Folder0.Item.Click()
                    End If
                    'If _CR = False Then
                    '    Folder2.Item.Enabled = False
                    'Else
                    '    Folder2.Item.Click()
                    'End If
                    If _CA = False Then
                        Folder3.Item.Enabled = False
                    Else
                        Folder3.Item.Click()
                    End If
                    'StaticText5.Item.Visible = True
                    'EditText4.Item.Visible = True
                    'EditText5.Item.Visible = True
                    'EditText6.Item.Visible = True
                    objform = objaddon.objapplication.Forms.GetForm("PAYM", 1)
                    objform.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                End If
                'objform.Update()
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub

        'Private Sub CalculateTotal()
        '    Try
        '        Dim Paid, Overall, balancedue, bankcharge, OverallFC, balancedueFC, ExcRate As Double
        '        Dim BTot, CheckTot, CashTot, CardTot As Double
        '        objform.Freeze(True)
        '        If ComboBox0.Selected.Value = MainCurr Then
        '            If Val(EditText0.Value) = 0 Then Overall = 0 Else Overall = CDbl(EditText0.Value)
        '            If Val(EditText1.Value) = 0 Then balancedue = 0 Else balancedue = CDbl(EditText1.Value)
        '            If Val(EditText3.Value) = 0 Then bankcharge = 0 Else bankcharge = CDbl(EditText3.Value)
        '        Else
        '            If Val(EditText4.Value) = 0 Then ExcRate = 1 Else ExcRate = CDbl(EditText4.Value)
        '            If Val(EditText0.Value) = 0 Then Overall = 0 Else Overall = CDbl(EditText0.Value)
        '            If Val(EditText1.Value) = 0 Then balancedue = 0 Else balancedue = CDbl(EditText1.Value)
        '            If Val(EditText5.Value) = 0 Then OverallFC = 0 Else OverallFC = CDbl(EditText5.Value)
        '            If Val(EditText6.Value) = 0 Then balancedueFC = 0 Else balancedueFC = CDbl(EditText6.Value)
        '            If Val(EditText3.Value) = 0 Then bankcharge = 0 Else bankcharge = CDbl(EditText3.Value)
        '        End If


        '        If Val(EditText13.Value) = 0 Then BTot = 0 Else BTot = CDbl(EditText13.Value)
        '        If Val(EditText11.Value) = 0 Then CashTot = 0 Else CashTot = CDbl(EditText11.Value)
        '        If Val(Matrix1.Columns.Item("amtdue").Cells.Item(1).Specific.String) = 0 Then CardTot = 0 Else CardTot = CDbl(Matrix1.Columns.Item("amtdue").Cells.Item(1).Specific.String)
        '        If Val(Matrix0.Columns.Item("chamt").ColumnSetting.SumValue) = 0 Then CheckTot = 0 Else CheckTot = CDbl(Matrix0.Columns.Item("chamt").ColumnSetting.SumValue)

        '        Paid = BTot + CashTot + CardTot + CheckTot
        '        If ComboBox0.Selected.Value = MainCurr Then
        '            EditText1.Value = Overall - (Paid + bankcharge)  'balance due
        '            EditText2.Value = bankcharge + Paid 'Paid
        '        Else
        '            EditText1.Value = Math.Round(Overall - ((Paid + bankcharge) * ExcRate), 6)  'balance due
        '            EditText2.Value = bankcharge + Paid 'Paid
        '            EditText6.Value = OverallFC - (Paid + bankcharge)
        '        End If

        '        objform.Freeze(False)
        '    Catch ex As Exception
        '        objform.Freeze(False)
        '    End Try
        'End Sub

        Private Sub CFLcondition(ByVal pVal As SAPbouiCOM.SBOItemEventArg, ByVal CFLID As String)
            If pVal.ActionSuccess = True Then Exit Sub
            Try
                Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item(CFLID)
                Dim oConds As SAPbouiCOM.Conditions
                Dim oCond As SAPbouiCOM.Condition
                Dim oEmptyConds As New SAPbouiCOM.Conditions
                oCFL.SetConditions(oEmptyConds)
                oConds = oCFL.GetConditions()

                oCond = oConds.Add()
                oCond.Alias = "Postable"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = "Y"
                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                oCond = oConds.Add()
                oCond.Alias = "LocManTran"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = "N"
                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                oCond = oConds.Add()
                oCond.Alias = "FrozenFor"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = "N"
                oCFL.SetConditions(oConds)
            Catch ex As Exception
                'SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End Try
        End Sub

        Private Sub ChooseFromList_AfterAction_AccountSelection_Matrix(ByVal pVal As SAPbouiCOM.SBOItemEventArg, ByVal MatrixName As SAPbouiCOM.Matrix, ByVal colname_acctcode As String, ByVal colname_acctname As String)
            Try
                If pVal.ActionSuccess = False Then Exit Sub
                pCFL = pVal
                If Not pCFL.SelectedObjects Is Nothing Then
                    Try
                        MatrixName.Columns.Item(colname_acctcode).Cells.Item(pVal.Row).Specific.string = pCFL.SelectedObjects.Columns.Item("AcctCode").Cells.Item(0).Value
                    Catch ex As Exception
                    End Try
                    Try
                        MatrixName.Columns.Item(colname_acctname).Cells.Item(pVal.Row).Specific.string = pCFL.SelectedObjects.Columns.Item("AcctName").Cells.Item(0).Value
                    Catch ex As Exception
                    End Try
                End If
            Catch ex As Exception
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End Sub

        Private Sub ChooseFromList_AfterAction_AccountSelection(ByVal pVal As SAPbouiCOM.SBOItemEventArg, ByVal editext_acctcode As SAPbouiCOM.EditText, ByVal acctname As SAPbouiCOM.StaticText)
            Try
                If pVal.ActionSuccess = False Then Exit Sub
                pCFL = pVal
                If Not pCFL.SelectedObjects Is Nothing Then
                    Try
                        editext_acctcode.Value = pCFL.SelectedObjects.Columns.Item("AcctCode").Cells.Item(0).Value
                    Catch ex As Exception
                    End Try
                    Try
                        acctname.Caption = pCFL.SelectedObjects.Columns.Item("AcctName").Cells.Item(0).Value
                    Catch ex As Exception
                    End Try
                End If
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            Catch ex As Exception
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End Sub

        Private Sub EditText17_ChooseFromListBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles EditText17.ChooseFromListBefore
            Try 'Cheque GL
                CFLcondition(pVal, "CFL_0")
            Catch ex As Exception
            End Try

        End Sub

        Private Sub EditText17_ChooseFromListAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText17.ChooseFromListAfter
            Try 'Cheque GL
                ChooseFromList_AfterAction_AccountSelection(pVal, EditText17, StaticText23)
            Catch ex As Exception
            End Try

        End Sub

        Private Sub Button0_ClickBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Button0.ClickBefore
            Try
                'If pVal.InnerEvent = True Then BubbleEvent = False : Exit Sub
                If Button0.Item.Enabled = False Then Exit Sub
                If Validate() Then
                    objaddon.objapplication.StatusBar.SetText("Please update the mandatory fields...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    BubbleEvent = False : Exit Sub
                End If
                objPayform = objaddon.objapplication.Forms.GetForm("PAYINIT", FormCount)
                If objPayform.Items.Item("opincpay").Specific.Selected = True Then
                    objPayform = objaddon.objapplication.Forms.GetForm("FINPAY", FormCount)
                ElseIf objPayform.Items.Item("opoutpay").Specific.Selected = True Then
                    objPayform = objaddon.objapplication.Forms.GetForm("FOUTPAY", FormCount)
                End If

                If objPayform.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then Exit Sub
                If EditText2.Value < 0 Then objaddon.objapplication.StatusBar.SetText("Negative amount...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
                If ComboBox0.Selected.Value = MainCurr Then
                    If CDbl(EditText0.Value) - CDbl(EditText2.Value) <> 0 Then objaddon.objapplication.StatusBar.SetText("You should not allow to post the transaction with due amount...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
                Else
                    'objaddon.objapplication.StatusBar.SetText("Please use system currency, other currencies will be soon...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    'BubbleEvent = False : Exit Sub
                    If CDbl(EditText5.Value) - CDbl(EditText2.Value) <> 0 Then objaddon.objapplication.StatusBar.SetText("You should not allow to post the transaction with due amount...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
                    objPayform.Items.Item("tfc").Specific.String = EditText5.Value
                    objPayform.Items.Item("ttotdue").Specific.String = EditText0.Value ' EditText0.Value
                    objPayform.Items.Item("tacttotal").Specific.String = EditText7.Value
                End If

                If Val(EditText11.Value) > 0 Then 'Cash
                    If EditText19.Value = "" Then objaddon.objapplication.StatusBar.SetText("Please update G/L account for cash...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
                End If
                If Val(EditText13.Value) > 0 Then  'Bank Transfer
                    If EditText14.Value = "" Then objaddon.objapplication.StatusBar.SetText("Please update G/L account for bank transfer...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
                End If
                If Matrix0.VisualRowCount > 0 Then  ' Cheque
                    If Val(Matrix0.Columns.Item("chamt").Cells.Item(1).Specific.String) > 0 Then
                        If FindPayment = "Y" Then
                            If EditText17.Value = "" Then objaddon.objapplication.StatusBar.SetText("Please update G/L account for Cheque...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
                        End If

                    End If
                End If

                'If Matrix1.VisualRowCount > 0 Then  'Card
                '    If Val(Matrix1.Columns.Item("amtdue").Cells.Item(1).Specific.String) > 0 Then
                '        If Matrix1.Columns.Item("cardgl").Cells.Item(1).Specific.String = "" Then objaddon.objapplication.StatusBar.SetText("Please update G/L account for Card...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
                '    End If
                'End If
                If Val(EditText3.Value) > 0 Then  'BCG Amt
                    If BCGAcct = "" Then objaddon.objapplication.StatusBar.SetText("Please define G/L account for Bank...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
                End If
                objPayform.Items.Item("tcurr").Specific.String = ComboBox0.Selected.Value
                If ComboBox0.Selected.Value <> MainCurr Then
                    Dim forex As Double = Math.Round(CDbl(EditText0.Value) - CDbl(EditText5.Value) * CDbl(EditText4.Value), SumRound) 'Math.Round(((CDbl(EditText5.Value) * CDbl(EditText4.Value)) - CDbl(EditText0.Value)), SumRound) 'Math.Round((EditText0.Value - (EditText5.Value * EditText4.Value)) - (EditText0.Value - EditText7.Value), SumRound)
                    objPayform.Items.Item("tblexrate").Specific.String = forex ' EditText1.Value 
                End If

                'Transfer
                objPayform.Items.Item("tbigl").Specific.String = EditText14.Value
                objPayform.Items.Item("tbidate").Specific.String = EditText15.Value
                objPayform.Items.Item("tbiref").Specific.String = EditText16.Value
                objPayform.Items.Item("tbitot").Specific.String = EditText13.Value
                'Cash
                objPayform.Items.Item("tcigl").Specific.String = EditText19.Value
                objPayform.Items.Item("tcitot").Specific.String = EditText11.Value
                'Cheque
                objPayform.Items.Item("tchigl").Specific.String = EditText17.Value
                Dim ChTot As Double
                For i As Integer = 1 To Matrix0.VisualRowCount
                    If Val(Matrix0.Columns.Item("chamt").Cells.Item(i).Specific.String) > 0 Then
                        ChTot += CDbl(Matrix0.Columns.Item("chamt").Cells.Item(i).Specific.String)
                    End If
                Next
                objPayform.Items.Item("tchtot").Specific.String = ChTot ' Matrix0.Columns.Item("chamt").ColumnSetting.SumValue
                objPayform.Items.Item("tcrtot").Specific.String = EditText18.Value

                objPayform.Items.Item("tbcgtot").Specific.String = EditText3.Value
                objPayform.Items.Item("tbcggl").Specific.String = BCGAcct

                objMatrix = objPayform.Items.Item("mtxcheq").Specific
                If Assign_DataTable_To_Matrix(objform, objMatrix, "mtxcheq", "chamt", True) Then

                End If
                'Credit
                'objMatrix = objPayform.Items.Item("mtxcr").Specific
                'If Assign_DataTable_To_Matrix(objform, objMatrix, "mtxcr", "amtdue", True) Then

                'End If
                If ComboBox0.Selected.Value <> MainCurr Then
                    objPayform.Items.Item("tboeamt").Specific.String = EditText4.Value
                End If

            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        End Sub

        Private Function buildMatrixTable(ByVal oForm As SAPbouiCOM.Form, ByVal sMatrixUID As String, ByVal sKeyFieldID As String) As DataTable
            'Dim oForm As SAPbouiCOM.Form = Nothing
            Dim oMatrix As SAPbouiCOM.Matrix = Nothing
            Try
                Dim oDT As New DataTable
                'oForm = objaddon.objapplication.Forms.GetForm(pVal.FormType, pVal.FormTypeCount)
                oMatrix = oForm.Items.Item(sMatrixUID).Specific
                'Add all of the columns by unique ID to the DataTable
                For iCol As Integer = 0 To oMatrix.Columns.Count - 1
                    'Skip invisible columns
                    If oMatrix.Columns.Item(iCol).Visible = False Then Continue For
                    oDT.Columns.Add(oMatrix.Columns.Item(iCol).UniqueID)
                Next
                'Now, add all of the data into the DataTable
                For iRow As Integer = 1 To oMatrix.RowCount
                    Dim oRow As DataRow = oDT.NewRow
                    For iCol As Integer = 0 To oMatrix.Columns.Count - 1
                        If oMatrix.Columns.Item(iCol).Visible = False Then Continue For
                        oRow.Item(oMatrix.Columns.Item(iCol).UniqueID) = oMatrix.Columns.Item(iCol).Cells.Item(iRow).Specific.Value
                    Next
                    'If the Key field has no value, then the row is empty, skip adding it.
                    If oRow(sKeyFieldID).ToString.Trim = 0 Then Continue For
                    oDT.Rows.Add(oRow)
                Next

                Return oDT
            Catch ex As Exception
                Return Nothing
            End Try

        End Function

        Private Function GetPaymentDT(ByVal oForm As SAPbouiCOM.Form, ByVal sMatrixUID As String, ByVal sKeyFieldID As String) As DataTable
            'Dim oForm As SAPbouiCOM.Form = Nothing
            Dim oMatrix As SAPbouiCOM.Matrix = Nothing
            Dim objcheckbox As SAPbouiCOM.CheckBox
            Try
                Dim oDT As New DataTable
                'oForm = objaddon.objapplication.Forms.GetForm(pVal.FormType, pVal.FormTypeCount)
                oMatrix = oForm.Items.Item(sMatrixUID).Specific
                'Add all of the columns by unique ID to the DataTable
                For iCol As Integer = 0 To oMatrix.Columns.Count - 1
                    'Skip invisible columns
                    If oMatrix.Columns.Item(iCol).Visible = False Then Continue For
                    If oMatrix.Columns.Item(iCol).UniqueID = "total" Or oMatrix.Columns.Item(iCol).UniqueID = "baldue" Or oMatrix.Columns.Item(iCol).UniqueID = "paytot" Or oMatrix.Columns.Item(iCol).UniqueID = "doccur" Then
                        oDT.Columns.Add(oMatrix.Columns.Item(iCol).UniqueID)
                    End If
                Next
                For iRow As Integer = 1 To oMatrix.VisualRowCount
                    objcheckbox = oMatrix.Columns.Item("select").Cells.Item(iRow).Specific
                    If objcheckbox.Checked = True Then
                        Dim oRow As DataRow = oDT.NewRow
                        For iCol As Integer = 0 To oMatrix.Columns.Count - 1
                            If oMatrix.Columns.Item(iCol).Visible = False Then Continue For
                            If oMatrix.Columns.Item(iCol).UniqueID = "total" Or oMatrix.Columns.Item(iCol).UniqueID = "baldue" Or oMatrix.Columns.Item(iCol).UniqueID = "paytot" Or oMatrix.Columns.Item(iCol).UniqueID = "doccur" Then
                                oRow.Item(oMatrix.Columns.Item(iCol).UniqueID) = oMatrix.Columns.Item(iCol).Cells.Item(iRow).Specific.Value
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

        Private Function Assign_DataTable_To_Matrix(ByVal oForm As SAPbouiCOM.Form, ByVal DataMatrix As SAPbouiCOM.Matrix, ByVal FromMatrixID As String, ByVal ColName As String, ByVal Button_click As Boolean) As Boolean
            Try
                Dim Flag As Boolean = False
                Dim cmbcolumn As SAPbouiCOM.ComboBox

                objDT = buildMatrixTable(oForm, FromMatrixID, ColName) '(objPayform, "mtxcheq", "chamt")
                For i As Integer = 0 To objDT.Rows.Count - 1
                    If DataMatrix.VisualRowCount = 0 Then
                        DataMatrix.AddRow()
                    Else
                        If CDbl(DataMatrix.Columns.Item(ColName).Cells.Item(DataMatrix.VisualRowCount).Specific.String) > 0 Then
                            DataMatrix.AddRow()
                        End If
                    End If
                    'If FromMatrixID = "mtxcheq" Then
                    '    RowID += 1
                    'End If
                    RowID += 1
                    'Dim MatRow As Integer
                    For j As Integer = 0 To objDT.Columns.Count - 1
                        If objDT.Rows(i)(j).ToString <> "" Then
                            Flag = True
                            'MatRow += j
                            If Button_click = True Then
                                'For MatRow As Integer = 1 To DataMatrix.VisualRowCount
                                '    If DataMatrix.Columns.Item(MatRow).UniqueID <> objDT.Columns(j).ToString Then Continue For
                                '    DataMatrix.Columns.Item(j).Cells.Item(RowID).Specific.String = objDT.Rows(i)(j).ToString
                                'Next
                                If FindPayment = "IN" And (j = 5) Then
                                    DataMatrix.Columns.Item(7).Cells.Item(RowID).Specific.String = objDT.Rows(i)(j).ToString
                                Else
                                    DataMatrix.Columns.Item(j).Cells.Item(RowID).Specific.String = objDT.Rows(i)(j).ToString
                                End If
                            Else
                                'Dim va As String = objDT.Rows(i)(j).ToString
                                If (FromMatrixID = "mtxcheq" And (j = 8)) Or (FromMatrixID = "mtxcr" And (j = 0 Or j = 6 Or j = 11)) Then 'j = 3 Or j = 4 Or
                                    cmbcolumn = DataMatrix.Columns.Item(j).Cells.Item(RowID).Specific
                                    cmbcolumn.Select(objDT.Rows(i)(j).ToString, SAPbouiCOM.BoSearchKey.psk_ByValue)
                                Else
                                    If FindPayment = "IN" And (j = 9) Then
                                        DataMatrix.Columns.Item(10).Cells.Item(RowID).Specific.String = objDT.Rows(i)(j).ToString
                                    ElseIf FindPayment = "IN" And (j = 10) Then
                                        DataMatrix.Columns.Item(11).Cells.Item(RowID).Specific.String = objDT.Rows(i)(j).ToString
                                    ElseIf FindPayment = "IN" And (j = 5) Then
                                        DataMatrix.Columns.Item(7).Cells.Item(RowID).Specific.String = objDT.Rows(i)(j).ToString
                                    Else
                                        DataMatrix.Columns.Item(j).Cells.Item(RowID).Specific.String = objDT.Rows(i)(j).ToString
                                    End If
                                End If
                            End If
                        End If
                    Next
                Next

                DataMatrix.AutoResizeColumns()
                Return Flag
            Catch ex As Exception

            End Try
        End Function

        Private Sub EditText14_ChooseFromListBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles EditText14.ChooseFromListBefore
            Try 'Bank Transfer GL
                CFLcondition(pVal, "CFL_1")
            Catch ex As Exception
            End Try

        End Sub

        Private Sub EditText14_ChooseFromListAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText14.ChooseFromListAfter
            Try 'Bank Transfer GL
                ChooseFromList_AfterAction_AccountSelection(pVal, EditText14, StaticText21)
            Catch ex As Exception
            End Try
        End Sub

        Private Sub EditText19_ChooseFromListBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles EditText19.ChooseFromListBefore
            Try 'Cash GL
                CFLcondition(pVal, "CFL_3")
            Catch ex As Exception
            End Try


        End Sub

        Private Sub EditText19_ChooseFromListAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText19.ChooseFromListAfter
            Try 'Cash GL
                ChooseFromList_AfterAction_AccountSelection(pVal, EditText19, StaticText26)
            Catch ex As Exception
            End Try

        End Sub

        Private Sub Matrix0_LostFocusAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.LostFocusAfter
            Try
                If pVal.Row = 0 Then Exit Sub
                If pVal.ActionSuccess = False Then Exit Sub
                Select Case pVal.ColUID
                    Case "chamt"
                        'If Val(Matrix0.Columns.Item("chamt").Cells.Item(pVal.Row).Specific.String) > 0 Then
                        If Val(Matrix0.Columns.Item("chamt").Cells.Item(Matrix0.VisualRowCount).Specific.string) > 0 Then
                            objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "chamt", "#")
                        End If

                        'If Val(Matrix0.Columns.Item("chamt").Cells.Item(Matrix0.VisualRowCount).Specific.string) > 0 Then
                        '    Matrix0.AddRow(1)
                        '    Matrix0.ClearRowData(Matrix0.VisualRowCount)
                        '    Matrix0.Columns.Item("#").Cells.Item(Matrix0.VisualRowCount).Specific.string = Matrix0.VisualRowCount
                        'End If
                        'End If
                End Select
                'objform.Freeze(True)
                Matrix0.AutoResizeColumns()
                'objform.Freeze(False)
            Catch ex As Exception
                'objform.Freeze(False)
            End Try

        End Sub

        Private Sub Matrix0_ComboSelectAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.ComboSelectAfter
            Try
                If pVal.InnerEvent = True Then Exit Sub
                Select Case pVal.ColUID
                    Case "chcty"
                        Dim cmbbank As SAPbouiCOM.ComboBox
                        cmbbank = Matrix0.Columns.Item("chbank").Cells.Item(pVal.Row).Specific
                        objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        objRs.DoQuery("select T0.""BankCode"",T0.""BankName"" from ODSC T0 where T0.""CountryCod""='" & Matrix0.Columns.Item("chcty").Cells.Item(pVal.Row).Specific.Selected.Value & "' Order by T0.""BankName""")
                        If objRs.RecordCount > 0 Then
                            If cmbbank.ValidValues.Count > 0 Then
                                For i = cmbbank.ValidValues.Count - 1 To 0 Step -1
                                    cmbbank.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                                Next
                            End If
                            cmbbank.ValidValues.Add("-1", "")
                            For i As Integer = 0 To objRs.RecordCount - 1
                                Try
                                    cmbbank.ValidValues.Add(objRs.Fields.Item("BankCode").Value, objRs.Fields.Item("BankName").Value)
                                    objRs.MoveNext()
                                Catch ex As Exception
                                    objRs.MoveNext()
                                End Try
                            Next
                        End If
                    Case "chbank"

                End Select

            Catch ex As Exception

            End Try

        End Sub

        Private Sub Matrix0_ChooseFromListBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Matrix0.ChooseFromListBefore
            Try
                If pVal.ActionSuccess = True Then Exit Sub
                'If pVal.InnerEvent = True Then Exit Sub
                If pVal.ColUID = "chbranch" Or pVal.ColUID = "chact" Then
                    Try
                        Dim oCFL As SAPbouiCOM.ChooseFromList
                        If pVal.ColUID = "chbranch" Then
                            If FindPayment = "IN" Then
                                oCFL = objform.ChooseFromLists.Item("CFL_4")
                            ElseIf FindPayment = "OUT" Then
                                'Dim ColItem As SAPbouiCOM.Column = Matrix0.Columns.Item("chbranch")
                                'ColItem.ChooseFromListUID.Remove(pVal.Row)
                                'objform.Update()
                                'objform.Refresh()
                                'ColItem.ChooseFromListUID = "C_Branch"
                                'ColItem.ChooseFromListAlias = "DfltBranch"
                                'oCFL = objform.ChooseFromLists.Item("C_Branch")
                                Exit Sub
                            Else
                                Exit Sub
                            End If
                        Else
                            oCFL = objform.ChooseFromLists.Item("CFL_5")
                        End If

                        Dim oConds As SAPbouiCOM.Conditions
                        Dim oCond As SAPbouiCOM.Condition
                        Dim oEmptyConds As New SAPbouiCOM.Conditions
                        oCFL.SetConditions(oEmptyConds)
                        oConds = oCFL.GetConditions()

                        If FindPayment = "IN" Then
                            oCond = oConds.Add()
                            oCond.Alias = "Country"
                            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_NONE
                            oCond.CondVal = "-1"
                        End If

                        'oCond = oConds.Add()
                        'oCond.Alias = "BankCode"
                        'oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_NONE
                        'oCond.CondVal = "-1"

                        'oCond = oConds.Add()
                        'oCond.Alias = "Country"
                        'oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        'oCond.CondVal = Matrix0.Columns.Item("chcty").Cells.Item(pVal.Row).Specific.String
                        'oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                        'oCond = oConds.Add()
                        'oCond.Alias = "BankCode"
                        'oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        'oCond.CondVal = Matrix0.Columns.Item("chbank").Cells.Item(pVal.Row).Specific.String
                        oCFL.SetConditions(oConds)
                    Catch ex As Exception
                        'SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try
                ElseIf pVal.ColUID = "chissue" Then
                    Try
                        Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFL_6")
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
                        'SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try
                ElseIf pVal.ColUID = "chcty" Then
                    Try
                        Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("C_Country")
                        Dim oConds As SAPbouiCOM.Conditions
                        Dim oCond As SAPbouiCOM.Condition
                        Dim oEmptyConds As New SAPbouiCOM.Conditions
                        oCFL.SetConditions(oEmptyConds)
                        oConds = oCFL.GetConditions()

                        Dim rsetCFL As SAPbobsCOM.Recordset = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oCFL.SetConditions(oEmptyConds)
                        oConds = oCFL.GetConditions()

                        rsetCFL.DoQuery("select distinct T0.""CountryCod"",T1.""Name"" from ODSC T0 join OCRY T1 on T0.""CountryCod""=T1.""Code""")
                        rsetCFL.MoveFirst()
                        For i As Integer = 1 To rsetCFL.RecordCount
                            If i = (rsetCFL.RecordCount) Then
                                oCond = oConds.Add()
                                oCond.Alias = "Code"
                                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                                oCond.CondVal = Trim(rsetCFL.Fields.Item(0).Value)
                            Else
                                oCond = oConds.Add()
                                oCond.Alias = "Code"
                                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                                oCond.CondVal = Trim(rsetCFL.Fields.Item(0).Value)
                                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                            End If
                            rsetCFL.MoveNext()
                        Next
                        If rsetCFL.RecordCount = 0 Then
                            oCond = oConds.Add()
                            oCond.Alias = "Code"
                            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_NONE
                            oCond.CondVal = "-1"
                        End If
                        oCFL.SetConditions(oConds)
                    Catch ex As Exception
                        'SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try
                ElseIf pVal.ColUID = "chbank" Then
                    Try
                        Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("C_BankCode")
                        Dim oConds As SAPbouiCOM.Conditions
                        Dim oCond As SAPbouiCOM.Condition
                        Dim oEmptyConds As New SAPbouiCOM.Conditions
                        oCFL.SetConditions(oEmptyConds)
                        oConds = oCFL.GetConditions()

                        oCond = oConds.Add()
                        oCond.Alias = "CountryCod"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = Matrix0.Columns.Item("chcty").Cells.Item(pVal.Row).Specific.String
                        oCFL.SetConditions(oConds)
                    Catch ex As Exception
                        'SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try
                ElseIf pVal.ColUID = "glacc" Then
                    Try
                        Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("C_glacc")
                        Dim oConds As SAPbouiCOM.Conditions
                        Dim oCond As SAPbouiCOM.Condition
                        Dim oEmptyConds As New SAPbouiCOM.Conditions
                        oCFL.SetConditions(oEmptyConds)
                        oConds = oCFL.GetConditions()

                        oCond = oConds.Add()
                        oCond.Alias = "Postable"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = "Y"
                        oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                        oCond = oConds.Add()
                        oCond.Alias = "LocManTran"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = "N"
                        oCFL.SetConditions(oConds)
                    Catch ex As Exception
                        'SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try
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
                If pVal.ColUID = "chbranch" And pVal.ActionSuccess = True Then
                    Try
                        pCFL = pVal
                        If pCFL.SelectedObjects Is Nothing Then Exit Sub

                        Try
                            Matrix0.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("Branch").Cells.Item(0).Value
                        Catch ex As Exception
                        End Try
                    Catch ex As Exception
                    End Try
                ElseIf pVal.ColUID = "chact" And pVal.ActionSuccess = True Then
                    Try
                        pCFL = pVal
                        If pCFL.SelectedObjects Is Nothing Then Exit Sub

                        Try
                            Matrix0.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("Account").Cells.Item(0).Value
                        Catch ex As Exception
                        End Try
                    Catch ex As Exception
                    End Try
                ElseIf pVal.ColUID = "chissue" And pVal.ActionSuccess = True Then
                    Try
                        pCFL = pVal
                        If pCFL.SelectedObjects Is Nothing Then Exit Sub
                        Try
                            Matrix0.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("CardCode").Cells.Item(0).Value
                        Catch ex As Exception
                        End Try
                    Catch ex As Exception
                    End Try
                ElseIf pVal.ColUID = "chcty" And pVal.ActionSuccess = True Then
                    Try
                        pCFL = pVal
                        If pCFL.SelectedObjects Is Nothing Then Exit Sub
                        Try
                            Matrix0.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("Code").Cells.Item(0).Value
                        Catch ex As Exception
                        End Try
                    Catch ex As Exception
                    End Try
                ElseIf pVal.ColUID = "chbank" And pVal.ActionSuccess = True Then
                    Try
                        pCFL = pVal
                        If pCFL.SelectedObjects Is Nothing Then Exit Sub
                        If FindPayment = "IN" Then
                            Try
                                Matrix0.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("BankCode").Cells.Item(0).Value
                            Catch ex As Exception
                                Matrix0.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("BankCode").Cells.Item(0).Value
                            End Try

                        ElseIf FindPayment = "OUT" Then
                            Try
                                Matrix0.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("BankCode").Cells.Item(0).Value
                                strSQL = "Select T0.""DfltBranch"",T0.""DfltAcct"",T1.""GLAccount"" from ODSC T0 left join DSC1 T1"
                                strSQL += vbCrLf + "on T0.""BankCode""=T1.""BankCode"" and T0.""AbsEntry""=T1.""BankKey"" and T0.""DfltBranch""=T1.""Branch"""
                                strSQL += vbCrLf + "where T0.""BankCode""='" & pCFL.SelectedObjects.Columns.Item("BankCode").Cells.Item(0).Value & "'"
                                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                objRs.DoQuery(strSQL)
                                If objRs.RecordCount > 0 Then
                                    Matrix0.Columns.Item("chbranch").Cells.Item(pVal.Row).Specific.String = objRs.Fields.Item(0).Value.ToString
                                    Matrix0.Columns.Item("chact").Cells.Item(pVal.Row).Specific.String = objRs.Fields.Item(1).Value.ToString
                                    Matrix0.Columns.Item("glacc").Cells.Item(pVal.Row).Specific.String = objRs.Fields.Item(2).Value.ToString
                                    Matrix0.Columns.Item("chnum").Cells.Item(pVal.Row).Click()
                                End If
                            Catch ex As Exception
                                Matrix0.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("BankCode").Cells.Item(0).Value
                                strSQL = "Select T0.""DfltBranch"",T0.""DfltAcct"",T1.""GLAccount"" from ODSC T0 left join DSC1 T1"
                                strSQL += vbCrLf + "on T0.""BankCode""=T1.""BankCode"" and T0.""AbsEntry""=T1.""BankKey"" and T0.""DfltBranch""=T1.""Branch"""
                                strSQL += vbCrLf + "where T0.""BankCode""='" & pCFL.SelectedObjects.Columns.Item("BankCode").Cells.Item(0).Value & "'"
                                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                objRs.DoQuery(strSQL)
                                If objRs.RecordCount > 0 Then
                                    Matrix0.Columns.Item("chbranch").Cells.Item(pVal.Row).Specific.String = objRs.Fields.Item(0).Value.ToString
                                    Matrix0.Columns.Item("chact").Cells.Item(pVal.Row).Specific.String = objRs.Fields.Item(1).Value.ToString
                                    Matrix0.Columns.Item("glacc").Cells.Item(pVal.Row).Specific.String = objRs.Fields.Item(2).Value.ToString
                                    Matrix0.Columns.Item("chnum").Cells.Item(pVal.Row).Click()
                                End If
                            End Try

                        End If

                    Catch ex As Exception
                    End Try
                ElseIf pVal.ColUID = "glacc" And pVal.ActionSuccess = True Then
                    Try
                        pCFL = pVal
                        If pCFL.SelectedObjects Is Nothing Then Exit Sub
                        Try
                            Matrix0.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("AcctCode").Cells.Item(0).Value
                        Catch ex As Exception
                        End Try
                    Catch ex As Exception
                    End Try
                End If
                Matrix0.AutoResizeColumns()
                'objaddon.objapplication.Menus.Item("1300").Activate()
                GC.Collect()
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End Try

        End Sub

        Private Sub Matrix0_ValidateAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.ValidateAfter
            Try 'Cheque
                If pVal.ItemChanged = False Then Exit Sub
                Select Case pVal.ColUID
                    Case "chamt"
                        'CalculateTotal()
                        Calc_Total(objPayDT)
                End Select
            Catch ex As Exception
            End Try

        End Sub

        Private Sub Matrix1_ValidateAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix1.ValidateAfter
            'Try 'Card
            '    If pVal.ItemChanged = False Then Exit Sub
            '    Select Case pVal.ColUID
            '        Case "amtdue"
            '            CalculateTotal()
            '            EditText18.Value = CDbl(Matrix1.Columns.Item("amtdue").Cells.Item(1).Specific.String)
            '        Case "valid"
            '            If Len(Matrix1.Columns.Item("valid").Cells.Item(1).Specific.String) = 4 Then
            '                Matrix1.Columns.Item("valid").Cells.Item(1).Specific.String = Left(Matrix1.Columns.Item("valid").Cells.Item(1).Specific.String, 2) + "/" + Right(Matrix1.Columns.Item("valid").Cells.Item(1).Specific.String, 2)
            '            Else
            '                Matrix1.Columns.Item("valid").Cells.Item(pVal.Row).Click()
            '            End If
            '    End Select
            'Catch ex As Exception
            'End Try

        End Sub

        Private Sub EditText13_ValidateAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText13.ValidateAfter
            Try 'Bank Transfer
                If pVal.ItemChanged = False Then Exit Sub
                'CalculateTotal()
                Calc_Total(objPayDT)
            Catch ex As Exception
            End Try
        End Sub

        Private Sub EditText11_ValidateAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText11.ValidateAfter
            Try 'Cash
                If pVal.ItemChanged = False Then Exit Sub 'Or pVal.InnerEvent = True
                'CalculateTotal()
                Calc_Total(objPayDT)
            Catch ex As Exception
            End Try

        End Sub

        Private Sub EditText3_ValidateAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText3.ValidateAfter
            Try 'Bank Charge
                If pVal.ItemChanged = False Then Exit Sub
                'CalculateTotal()
                Calc_Total(objPayDT)
            Catch ex As Exception
            End Try
        End Sub

        Private Sub Matrix1_ValidateBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Matrix1.ValidateBefore
            Try 'Card
                'If pVal.ItemChanged = False Or pVal.InnerEvent = True Then Exit Sub
                Select Case pVal.ColUID
                    Case "amtdue"
                        'CalculateTotal()
                        Calc_Total(objPayDT)
                        EditText18.Value = CDbl(Matrix1.Columns.Item("amtdue").Cells.Item(1).Specific.String)
                    Case "valid"
                        If pVal.InnerEvent = True Then Exit Sub
                        If Matrix1.Columns.Item("valid").Cells.Item(1).Specific.String = "" Then Exit Sub
                        If Not Regex.IsMatch(Matrix1.Columns.Item("valid").Cells.Item(1).Specific.String, "^[0-9 ]+$") Then objaddon.objapplication.StatusBar.SetText("Date value not valid...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
                        If Left(Matrix1.Columns.Item("valid").Cells.Item(1).Specific.String, 2) > 12 Then objaddon.objapplication.StatusBar.SetText("Date value not valid...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
                        If Matrix1.Columns.Item("valid").Cells.Item(1).Specific.String.ToString Like "##/##" Then Exit Sub
                        Matrix1.Columns.Item("valid").Cells.Item(1).Specific.String = Left(Matrix1.Columns.Item("valid").Cells.Item(1).Specific.String, 2) + "/" + Right(Matrix1.Columns.Item("valid").Cells.Item(1).Specific.String, 2)
                End Select
            Catch ex As Exception
            End Try
        End Sub

        Private Sub Matrix1_ChooseFromListBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Matrix1.ChooseFromListBefore
            'CreditCard
            Try
                If pVal.ActionSuccess = True Then Exit Sub
                If pVal.ColUID = "cardgl" Then
                    Try
                        Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFL_7")
                        Dim oConds As SAPbouiCOM.Conditions
                        Dim oCond As SAPbouiCOM.Condition
                        Dim oEmptyConds As New SAPbouiCOM.Conditions
                        oCFL.SetConditions(oEmptyConds)
                        oConds = oCFL.GetConditions()

                        oCond = oConds.Add()
                        oCond.Alias = "Postable"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = "Y"
                        oCFL.SetConditions(oConds)
                    Catch ex As Exception
                        SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try
                End If

            Catch ex As Exception
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End Sub

        Private Sub Matrix1_ChooseFromListAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix1.ChooseFromListAfter
            'CreditCard
            Try
                If pVal.ActionSuccess = False Then Exit Sub
                If pVal.ColUID = "cardgl" And pVal.ActionSuccess = True Then
                    Try
                        pCFL = pVal
                        If pCFL.SelectedObjects Is Nothing Then Exit Sub
                        Try
                            Matrix1.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("AcctCode").Cells.Item(0).Value
                        Catch ex As Exception
                        End Try
                    Catch ex As Exception
                    End Try
                End If
                'Matrix1.AutoResizeColumns()
                'objaddon.objapplication.Menus.Item("1300").Activate()
                GC.Collect()
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End Try
        End Sub

        Private Sub Matrix1_ComboSelectAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix1.ComboSelectAfter
            Try
                If pVal.ColUID = "cardname" Then
                    objform.Freeze(True)
                    strSQL = objaddon.objglobalmethods.getSingleValue("select T0.""AcctCode"" from OCRC T0 where T0.""CreditCard""='" & Matrix1.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Selected.Value & "'")
                    If strSQL <> "" Then
                        Matrix1.Columns.Item("cardgl").Cells.Item(pVal.Row).Specific.String = Trim(strSQL)
                    End If
                    objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    objRs.DoQuery("select T1.""CrTypeCode"",T1.""CrTypeName"" from OCRC T0 join OCRP T1 on T0.""CreditCard""=T1.""CreditCard"" where T1.""CreditCard""='" & Matrix1.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Selected.Value & "'")
                    If objRs.RecordCount > 0 Then
                        cmbcol = Matrix1.Columns.Item("paymet")
                        cmbcol.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                        If cmbcol.ValidValues.Count > 0 Then
                            For i = cmbcol.ValidValues.Count - 1 To 0 Step -1
                                cmbcol.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                            Next
                        End If
                        For i As Integer = 0 To objRs.RecordCount - 1
                            Try
                                cmbcol.ValidValues.Add(objRs.Fields.Item("CrTypeCode").Value, objRs.Fields.Item("CrTypeName").Value)
                                objRs.MoveNext()
                            Catch ex As Exception
                                objRs.MoveNext()
                            End Try
                        Next
                        Matrix1.Columns.Item("paymet").Cells.Item(1).Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                        cmbcol.ValidValues.Add("-1", "")
                    End If
                    objform.Freeze(False)
                End If

            Catch ex As Exception
                objform.Freeze(False)
            End Try

        End Sub

        Private Sub EditText4_ValidateAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText4.ValidateAfter
            Try ' Exchange Rate
                If pVal.ItemChanged = False Or pVal.InnerEvent = True Then Exit Sub
                Calc_Total(objPayDT)
                'Dim LCTot, LCOverdue, FCTot, FCOverdue, ExcRate As Double
                'If Val(EditText4.Value) = 0 Then ExcRate = 1 Else ExcRate = CDbl(EditText4.Value)
                ''If Val(EditText0.Value) = 0 Then LCTot = 0 Else LCTot = CDbl(EditText0.Value)
                ''If Val(EditText1.Value) = 0 Then LCOverdue = 0 Else LCOverdue = CDbl(EditText1.Value)
                ''If Val(EditText5.Value) = 0 Then FCTot = 0 Else FCTot = CDbl(EditText5.Value)
                ''If Val(EditText6.Value) = 0 Then FCOverdue = 0 Else FCOverdue = CDbl(EditText6.Value)

                'EditText0.Value = LCTot 'CDbl(EditText5.Value) * CDbl(EditText4.Value)
                'EditText1.Value = LCOverdue 'CDbl(EditText6.Value) * CDbl(EditText4.Value)

                'EditText5.Value = FCTot 'CDbl(EditText0.Value) / CDbl(EditText4.Value)
                'EditText6.Value = FCOverdue 'CDbl(EditText1.Value) / CDbl(EditText4.Value)

            Catch ex As Exception
            End Try

        End Sub

        Private Sub Calc_Total(ByVal PayDT As DataTable)
            Try
                Dim LCTot, FCTot, PayTotal, ExcRate, bankcharge, Paid, OExcRate, OldTranValue As Double
                'Dim Overall, balancedue, OverallFC, balancedueFC As Double
                Dim BTot, CheckTot, CashTot, CardTot As Double
                Dim objExcRateDT As New DataTable
                objform = objaddon.objapplication.Forms.ActiveForm
                objform.Freeze(True)
                If ComboBox0.Selected.Value = MainCurr Then
                    If Val(EditText0.Value) = 0 Then LCTot = 0 Else LCTot = Math.Round(CDbl(EditText0.Value), SumRound)
                    'If Val(EditText1.Value) = 0 Then LCOverdue = 0 Else LCOverdue = Math.Round(CDbl(EditText1.Value), 6)
                    If Val(EditText3.Value) = 0 Then bankcharge = 0 Else bankcharge = Math.Round(CDbl(EditText3.Value), SumRound)
                Else
                    If Val(EditText4.Value) = 0 Then ExcRate = 1 Else ExcRate = Math.Round(CDbl(EditText4.Value), RateRound)
                    If Val(EditText3.Value) = 0 Then bankcharge = 0 Else bankcharge = Math.Round(CDbl(EditText3.Value), SumRound)
                    objPayform = objaddon.objapplication.Forms.GetForm("PAYINIT", 1)
                    If objPayform.Items.Item("opincpay").Specific.Selected = True Then
                        objPayform = objaddon.objapplication.Forms.GetForm("FINPAY", 1)
                    ElseIf objPayform.Items.Item("opoutpay").Specific.Selected = True Then
                        objPayform = objaddon.objapplication.Forms.GetForm("FOUTPAY", 1)
                    End If

                    For i As Integer = 0 To PayDT.Rows.Count - 1
                        PayTotal = CDbl(PayDT.Rows(i)("paytot").ToString)
                        If PayTotal <> 0 Then
                            If PayDT.Rows(i)("doccur").ToString = MainCurr Then
                                LCTot = Math.Round(LCTot + PayTotal, SumRound)
                                FCTot = Math.Round(FCTot + (PayTotal / ExcRate), SumRound)
                                OldTranValue = Math.Round(OldTranValue + PayTotal, SumRound)
                            ElseIf PayDT.Rows(i)("doccur").ToString = ComboBox0.Selected.Value Then
                                LCTot = Math.Round(LCTot + (PayTotal * ExcRate), SumRound)
                                FCTot = Math.Round(FCTot + PayTotal, SumRound)
                                OExcRate = Math.Round(GetTransaction_ExchangeRate(PayDT.Rows(i)("object").ToString, PayDT.Rows(i)("transid").ToString), RateRound)
                                OldTranValue = Math.Round(OldTranValue + (PayTotal * OExcRate), SumRound)
                            Else
                                strSQL = objaddon.objglobalmethods.getSingleValue("Select ""Rate"" from ORTT where ""RateDate""= '" & DocumentDate.ToString("yyyyMMdd") & "' and ""Currency""='" & PayDT.Rows(i)("doccur").ToString & "' ") 'PayDT.Rows(i)("date").ToString
                                strSQL = IIf(Val(strSQL) = 0, 1, strSQL)
                                LCTot = Math.Round(LCTot + (PayTotal * CDbl(strSQL)), SumRound) 'OExcRate
                                FCTot = Math.Round(FCTot + ((PayTotal * CDbl(strSQL)) / ExcRate), SumRound) 'OExcRate
                                OExcRate = Math.Round(GetTransaction_ExchangeRate(PayDT.Rows(i)("object").ToString, PayDT.Rows(i)("transid").ToString), RateRound)
                                OldTranValue = Math.Round(OldTranValue + (PayTotal * OExcRate), SumRound) 'OExcRate
                            End If
                        End If
                    Next
                End If

                If Val(EditText13.Value) = 0 Then BTot = 0 Else BTot = Math.Round(CDbl(EditText13.Value), SumRound)
                If Val(EditText11.Value) = 0 Then CashTot = 0 Else CashTot = Math.Round(CDbl(EditText11.Value), SumRound)
                'If Val(Matrix1.Columns.Item("amtdue").Cells.Item(1).Specific.String) = 0 Then CardTot = 0 Else CardTot = CDbl(Matrix1.Columns.Item("amtdue").Cells.Item(1).Specific.String)
                'If Val(Matrix0.Columns.Item("chamt").ColumnSetting.SumValue) = 0 Then CheckTot = 0 Else CheckTot = CDbl(Matrix0.Columns.Item("chamt").ColumnSetting.SumValue)
                For i As Integer = 1 To Matrix0.VisualRowCount
                    If Val(Matrix0.Columns.Item("chamt").Cells.Item(i).Specific.String) > 0 Then
                        CheckTot = Math.Round(CheckTot + CDbl(Matrix0.Columns.Item("chamt").Cells.Item(i).Specific.String), SumRound)
                    End If
                Next
                Paid = Math.Round(BTot + CashTot + CardTot + CheckTot, SumRound)
                If ComboBox0.Selected.Value = MainCurr Then
                    'EditText0.Value = Math.Round((LCTot), SumRound)
                    EditText1.Value = Math.Round(LCTot - (Paid + bankcharge), SumRound)  'balance LC
                    EditText2.Value = Math.Round(bankcharge + Paid, SumRound) 'Paid
                Else
                    EditText0.Value = Math.Round((LCTot), SumRound) '(LCTot / ExcRate) * ExcRate
                    EditText1.Value = Math.Round((Math.Round((LCTot / ExcRate), SumRound) * ExcRate) - ((Paid + bankcharge) * ExcRate), SumRound)  'Balance LC
                    EditText2.Value = Math.Round(bankcharge + Paid, SumRound) 'Paid
                    EditText5.Value = Math.Round((LCTot / ExcRate), SumRound) 'OverallFc
                    EditText6.Value = Math.Round((LCTot / ExcRate) - (Paid + bankcharge), SumRound) 'Balance FC
                    EditText7.Value = Math.Round(OldTranValue, SumRound)
                End If
                objform.Freeze(False)
            Catch ex As Exception
                objform.Freeze(False)
            End Try
        End Sub

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
                    strSQL = objaddon.objglobalmethods.getSingleValue("Select case when T0.""Credit""<>0 Then T0.""Credit""/T0.""FCCredit"" Else T0.""Debit""/T0.""FCDebit"" End as ""ExcRate"" from JDT1 T0 where T0.""TransId""=" & TransId & " and ""ShortName"" in (Select ""CardCode"" from OCRD)")
                    Rate = CDbl(strSQL)
                Else
                    Rate = 1
                End If

                Return Rate
            Catch ex As Exception
                Return 0
            End Try
        End Function

        Private Sub Clear_Payments()
            Try ' bank Charge , cash,bank,credit,paid,cheque
                If Val(EditText3.Value) > 0 Or Val(EditText11.Value) > 0 Or Val(EditText13.Value) > 0 Or Val(EditText18.Value) > 0 Or Val(EditText2.Value) > 0 Or Val(Matrix0.Columns.Item("chamt").ColumnSetting.SumValue) > 0 Then
                    EditText3.Value = ""
                    EditText11.Value = ""
                    EditText13.Value = ""
                    EditText18.Value = ""
                    EditText2.Value = ""
                    Matrix1.Clear()
                    Matrix0.Clear()
                    objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "chamt", "#")
                    objaddon.objglobalmethods.Matrix_Addrow(Matrix1, "amtdue", "#")
                End If
            Catch ex As Exception

            End Try
        End Sub

        Private Sub ComboBox0_ComboSelectAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles ComboBox0.ComboSelectAfter
            Try
                'If ComboBox0.Selected.Value = MainCurr Then
                '    StaticText5.Item.Visible = False
                '    EditText4.Item.Visible = False
                '    EditText5.Item.Visible = False
                '    EditText6.Item.Visible = False
                'Else
                '    If ActivateExchangeRateWindow() Then
                '        Exit Sub
                '    End If
                '    StaticText5.Item.Visible = True
                '    EditText4.Item.Visible = True
                '    EditText5.Item.Visible = True
                '    EditText6.Item.Visible = True
                'End If
                If pVal.ItemChanged = False Then Exit Sub
                Clear_Payments()
                Currency_FieldSetup()
                Calc_Total(objPayDT)

                'objPayform = objaddon.objapplication.Forms.GetForm("PAYINIT", 1)
                'If objPayform.Items.Item("opincpay").Specific.Selected = True Then
                '    GetData_Payment("FINPAY")
                'ElseIf objPayform.Items.Item("opoutpay").Specific.Selected = True Then
                '    GetData_Payment("FOUTPAY")
                'End If
            Catch ex As Exception
            End Try
        End Sub

        Private Sub EditText13_KeyDownBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles EditText13.KeyDownBefore
            Try 'Copy Balance Due
                If pVal.Modifiers = SAPbouiCOM.BoModifiersEnum.mt_SHIFT And (pVal.CharPressed = 66 Or pVal.CharPressed = 98) Then
                    If ComboBox0.Selected.Value = MainCurr Then
                        EditText13.Value = EditText1.Value
                    Else
                        EditText13.Value = EditText6.Value
                    End If
                    BubbleEvent = False
                End If
            Catch ex As Exception

            End Try

        End Sub

        Private Sub EditText11_KeyDownBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles EditText11.KeyDownBefore
            Try 'Copy Balance Due
                If pVal.Modifiers = SAPbouiCOM.BoModifiersEnum.mt_SHIFT And (pVal.CharPressed = 66 Or pVal.CharPressed = 98) Then
                    If ComboBox0.Selected.Value = MainCurr Then
                        EditText11.Value = EditText1.Value
                    Else
                        EditText11.Value = EditText6.Value
                    End If
                    BubbleEvent = False
                End If
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Matrix0_KeyDownBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Matrix0.KeyDownBefore
            Try 'Copy Balance Due
                Select Case pVal.ColUID
                    Case "chamt"
                        If pVal.Modifiers = SAPbouiCOM.BoModifiersEnum.mt_SHIFT And (pVal.CharPressed = 66 Or pVal.CharPressed = 98) Then
                            If ComboBox0.Selected.Value = MainCurr Then
                                Matrix0.Columns.Item("chamt").Cells.Item(pVal.Row).Specific.String = CDbl(EditText1.Value) + Val(Matrix0.Columns.Item("chamt").ColumnSetting.SumValue)
                            Else
                                Matrix0.Columns.Item("chamt").Cells.Item(pVal.Row).Specific.String = CDbl(EditText6.Value) + Val(Matrix0.Columns.Item("chamt").ColumnSetting.SumValue)
                            End If
                            BubbleEvent = False
                        End If
                End Select

            Catch ex As Exception

            End Try

        End Sub

        Private Sub Folder0_PressedAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Folder0.PressedAfter
            Try
                Matrix0.AutoResizeColumns()
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Currency_FieldSetup()
            Try
                If ComboBox0.Selected.Value = MainCurr Then
                    StaticText5.Item.Visible = False
                    EditText4.Item.Visible = False
                    EditText5.Item.Visible = False
                    EditText6.Item.Visible = False
                Else
                    objPayform = objaddon.objapplication.Forms.GetForm("PAYINIT", 1)
                    If objPayform.Items.Item("opincpay").Specific.Selected = True Then
                        objPayform = objaddon.objapplication.Forms.GetForm("FINPAY", 1)
                    ElseIf objPayform.Items.Item("opoutpay").Specific.Selected = True Then
                        objPayform = objaddon.objapplication.Forms.GetForm("FOUTPAY", 1)
                    End If
                    If objPayform.Items.Item("tcurr").Specific.String = "" Then
                        If ActivateExchangeRateWindow() Then
                            Exit Sub
                        End If
                    End If
                    StaticText5.Item.Visible = True
                    EditText4.Item.Visible = True
                    EditText5.Item.Visible = True
                    EditText6.Item.Visible = True
                End If
            Catch ex As Exception

            End Try
        End Sub

        Private Function ActivateExchangeRateWindow()
            Try
                If ComboBox0.Selected.Value <> MainCurr Then
                    strSQL = objaddon.objglobalmethods.getSingleValue("Select ""Rate"" from ORTT where ""RateDate""= '" & DocumentDate.ToString("yyyyMMdd") & "' and ""Currency""='" & ComboBox0.Selected.Value & "' ")
                    If strSQL = "" Then
                        objaddon.objapplication.Menus.Item("3333").Activate()
                        Dim oForm As SAPbouiCOM.Form
                        Dim oMatrix As SAPbouiCOM.Matrix
                        Dim oCombo As SAPbouiCOM.ComboBox
                        oForm = objaddon.objapplication.Forms.ActiveForm
                        oCombo = oForm.Items.Item("12").Specific
                        If oCombo.Selected.Value <> Year(Now.Date) Then oCombo.Select(Year(Now.Date), SAPbouiCOM.BoSearchKey.psk_ByValue)
                        oCombo = oForm.Items.Item("13").Specific
                        If oCombo.Selected.Value <> Month(Now.Date) Then oCombo.Select(Month(Now.Date), SAPbouiCOM.BoSearchKey.psk_ByValue)
                        oMatrix = oForm.Items.Item("4").Specific
                        Dim ColId As String = ""
                        For i As Integer = 0 To oMatrix.Columns.Count - 1
                            If oMatrix.Columns.Item(i).TitleObject.Caption = ComboBox0.Selected.Value Then
                                ColId = oMatrix.Columns.Item(i).UniqueID
                            End If
                        Next
                        oMatrix.Columns.Item(0).Cells.Item(CInt(Now.Date.ToString("dd"))).Click()
                        oMatrix.Columns.Item(ColId).Cells.Item(CInt(Now.Date.ToString("dd"))).Click()
                        Return True
                    Else
                        EditText4.Value = Math.Round(CDbl(strSQL), RateRound)
                        EditText5.Value = Math.Round(CDbl(EditText0.Value) / CDbl(strSQL), SumRound)
                        EditText6.Value = Math.Round(CDbl(EditText1.Value) / CDbl(strSQL), SumRound)
                        Folder3.Item.Click()
                        objform.ActiveItem = "tctot"
                        Return False
                    End If

                End If
            Catch ex As Exception
                Return False
            End Try
        End Function

        Private Function Validate() As Boolean
            Try
                Dim Flag As Boolean = False
                If Val(EditText11.Value) > 0 Then  'cash
                    If EditText19.Value = "" Then Flag = True
                End If
                If Val(EditText13.Value) > 0 Then 'bank
                    If EditText14.Value = "" Then Flag = True
                    If EditText15.Value = "" Then Flag = True
                    'If EditText16.Value = "" Then Flag = True
                End If
                'If Val(EditText18.Value) > 0 Then 'Credit

                'End If
                If Val(Matrix0.Columns.Item("chamt").ColumnSetting.SumValue) > 0 Then 'Cheque
                    If FindPayment = "IN" Then
                        If EditText17.Value = "" Then Flag = True
                    End If
                    For i As Integer = 1 To Matrix0.VisualRowCount
                        If Val(Matrix0.Columns.Item("chamt").Cells.Item(i).Specific.String) > 0 Then
                            If Matrix0.Columns.Item("chdate").Cells.Item(i).Specific.String = "" Then Flag = True
                            'If Matrix0.Columns.Item("chbranch").Cells.Item(i).Specific.String = "" Then Flag = True
                            'If Matrix0.Columns.Item("chact").Cells.Item(i).Specific.String = "" Then Flag = True
                            If Matrix0.Columns.Item("chnum").Cells.Item(i).Specific.String = "" Then Flag = True
                        End If
                    Next
                End If
                Return Flag
            Catch ex As Exception

            End Try
        End Function


    End Class
End Namespace


