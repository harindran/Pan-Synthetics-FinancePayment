Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace Finance_Payment
    <FormAttribute("MBAPSI", "Business Objects/FrmAPService_Invoice.b1f")>
    Friend Class FrmAPService_Invoice
        Inherits UserFormBase
        Public WithEvents objform As SAPbouiCOM.Form
        Private WithEvents pCFL As SAPbouiCOM.ISBOChooseFromListEventArg
        Private WithEvents odbdsDetails As SAPbouiCOM.DBDataSource
        Public WithEvents odbdsHeader As SAPbouiCOM.DBDataSource
        Dim objRs As SAPbobsCOM.Recordset
        Dim FormCount As Integer = 0
        Dim strSQL As String
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("1").Specific, SAPbouiCOM.Button)
            Me.Button1 = CType(Me.GetItem("2").Specific, SAPbouiCOM.Button)
            Me.StaticText0 = CType(Me.GetItem("l_docnum").Specific, SAPbouiCOM.StaticText)
            Me.ComboBox0 = CType(Me.GetItem("Series").Specific, SAPbouiCOM.ComboBox)
            Me.EditText0 = CType(Me.GetItem("t_docnum").Specific, SAPbouiCOM.EditText)
            Me.StaticText1 = CType(Me.GetItem("lposdate").Specific, SAPbouiCOM.StaticText)
            Me.EditText1 = CType(Me.GetItem("tposdate").Specific, SAPbouiCOM.EditText)
            Me.StaticText2 = CType(Me.GetItem("lduedate").Specific, SAPbouiCOM.StaticText)
            Me.EditText2 = CType(Me.GetItem("tduedate").Specific, SAPbouiCOM.EditText)
            Me.StaticText3 = CType(Me.GetItem("ldocdate").Specific, SAPbouiCOM.StaticText)
            Me.EditText3 = CType(Me.GetItem("tdocdate").Specific, SAPbouiCOM.EditText)
            Me.Matrix0 = CType(Me.GetItem("mtxcont").Specific, SAPbouiCOM.Matrix)
            Me.StaticText4 = CType(Me.GetItem("lremark").Specific, SAPbouiCOM.StaticText)
            Me.EditText4 = CType(Me.GetItem("tremark").Specific, SAPbouiCOM.EditText)
            Me.Folder0 = CType(Me.GetItem("fldrcont").Specific, SAPbouiCOM.Folder)
            Me.Folder1 = CType(Me.GetItem("Item_0").Specific, SAPbouiCOM.Folder)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler DataAddBefore, AddressOf Me.Form_DataAddBefore
            AddHandler DataLoadAfter, AddressOf Me.Form_DataLoadAfter

        End Sub

        Private Sub OnCustomInitialize()
            Try
                Dim cmbbranch, cmbtaxcode As SAPbouiCOM.Column
                objform = objaddon.objapplication.Forms.GetForm("MBAPSI", Me.FormCount)
                objform = objaddon.objapplication.Forms.ActiveForm
                objform.Freeze(True)
                odbdsHeader = objform.DataSources.DBDataSources.Item("@MIPL_OAPI")
                odbdsDetails = objform.DataSources.DBDataSources.Item("@MIPL_API1")
                objaddon.objglobalmethods.LoadSeries(objform, odbdsHeader, "MIAPSI")
                objform.Items.Item("tposdate").Specific.string = Now.Date.ToString("yyyyMMdd")
                objform.Items.Item("tdocdate").Specific.string = Now.Date.ToString("yyyyMMdd")
                objform.Items.Item("tremark").Specific.String = "Created By " & objaddon.objcompany.UserName & " on " & Now.ToString
                objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "vcode", "#")
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If Not Localization = "IN" Then
                    Matrix0.Columns.Item("saccode").Visible = False
                    Matrix0.Columns.Item("taxcode").Visible = False
                    cmbtaxcode = Matrix0.Columns.Item("otaxcode")
                    If objaddon.HANA Then
                        objRs.DoQuery("select  ""Code"",""Name"" from OVTG where ""Inactive""='N' and ""Category""='I'")
                    Else
                        objRs.DoQuery("select  Code,Name from OVTG where Inactive='N' and Category='I'")
                    End If
                    If objRs.RecordCount > 0 Then
                        For i As Integer = 0 To objRs.RecordCount - 1
                            cmbtaxcode.ValidValues.Add(objRs.Fields.Item(0).Value.ToString, objRs.Fields.Item(1).Value.ToString)
                            objRs.MoveNext()
                        Next
                    End If
                Else
                    Matrix0.Columns.Item("otaxcode").Visible = False
                End If
                Matrix0.Columns.Item("cc1").Visible = False
                Matrix0.Columns.Item("cc2").Visible = False
                Matrix0.Columns.Item("cc3").Visible = False
                Matrix0.Columns.Item("cc4").Visible = False
                Matrix0.Columns.Item("cc5").Visible = False
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
                cmbbranch = Matrix0.Columns.Item("branch")
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If objaddon.HANA Then
                    'objRs.DoQuery("select ""BPLId"",""BPLName"" from OBPL where ifnull(""Disabled"",'N') <>'Y'")
                    objRs.DoQuery("Select T0.""BPLId"",T0.""BPLName"" from OBPL T0 join USR6 T1 on T0.""BPLId""=T1.""BPLId"" where T1.""UserCode""='" & objaddon.objcompany.UserName & "' and T0.""Disabled""<>'Y';")
                Else
                    'objRs.DoQuery("select BPLId,BPLName from OBPL where isnull(Disabled,'N') <>'Y'")
                    objRs.DoQuery("Select T0.BPLId,T0.BPLName from OBPL T0 join USR6 T1 on T0.BPLId=T1.BPLId where T1.UserCode='" & objaddon.objcompany.UserName & "' and T0.Disabled<>'Y';")
                End If
                If objRs.RecordCount > 0 Then
                    For i As Integer = 0 To objRs.RecordCount - 1
                        cmbbranch.ValidValues.Add(objRs.Fields.Item(0).Value.ToString, objRs.Fields.Item(1).Value.ToString)
                        objRs.MoveNext()
                    Next
                End If

                'Matrix0.Columns.Item("total").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                Matrix0.Columns.Item("gtotal").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                objform.Settings.Enabled = True
                objRs = Nothing
                Matrix0.AutoResizeColumns()
                objform.Freeze(False)
            Catch ex As Exception
                objform.Freeze(False)
                objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub

#Region "Fields"
        Private WithEvents Button0 As SAPbouiCOM.Button
        Private WithEvents Button1 As SAPbouiCOM.Button
        Private WithEvents StaticText0 As SAPbouiCOM.StaticText
        Private WithEvents ComboBox0 As SAPbouiCOM.ComboBox
        Private WithEvents EditText0 As SAPbouiCOM.EditText
        Private WithEvents StaticText1 As SAPbouiCOM.StaticText
        Private WithEvents EditText1 As SAPbouiCOM.EditText
        Private WithEvents StaticText2 As SAPbouiCOM.StaticText
        Private WithEvents EditText2 As SAPbouiCOM.EditText
        Private WithEvents StaticText3 As SAPbouiCOM.StaticText
        Private WithEvents EditText3 As SAPbouiCOM.EditText
        Private WithEvents Matrix0 As SAPbouiCOM.Matrix
        Private WithEvents StaticText4 As SAPbouiCOM.StaticText
        Private WithEvents EditText4 As SAPbouiCOM.EditText
        Private WithEvents Folder0 As SAPbouiCOM.Folder
#End Region

        Private Sub Button0_ClickBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Button0.ClickBefore
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                If pVal.InnerEvent = True Then BubbleEvent = False : Exit Sub
                If EditText2.Value = "" Then
                    objaddon.objapplication.SetStatusBarMessage("Due Date is Missing", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                    BubbleEvent = False : Exit Sub
                End If
                If Matrix0.VisualRowCount <= 1 And Matrix0.Columns.Item("vcode").Cells.Item(1).Specific.string = "" Then objaddon.objapplication.SetStatusBarMessage("Row is Empty...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : BubbleEvent = False : Exit Sub
                For Row As Integer = 1 To Matrix0.VisualRowCount
                    If Matrix0.Columns.Item("vcode").Cells.Item(Row).Specific.string <> "" Then
                        If Localization = "IN" Then
                            If Matrix0.Columns.Item("glacc").Cells.Item(Row).Specific.string = "" Or Matrix0.Columns.Item("desc").Cells.Item(Row).Specific.string = "" Or Matrix0.Columns.Item("saccode").Cells.Item(Row).Specific.string = "" Or Matrix0.Columns.Item("taxcode").Cells.Item(Row).Specific.string = "" Or CDbl(Matrix0.Columns.Item("total").Cells.Item(Row).Specific.string) <= 0 Or Matrix0.Columns.Item("branch").Cells.Item(Row).Specific.Selected Is Nothing Or Matrix0.Columns.Item("docdate").Cells.Item(Row).Specific.string = "" Then
                                objaddon.objapplication.SetStatusBarMessage("Please update the line level values on row: " & Row, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False : Exit Sub
                            End If
                        Else
                            If Matrix0.Columns.Item("glacc").Cells.Item(Row).Specific.string = "" Or Matrix0.Columns.Item("desc").Cells.Item(Row).Specific.string = "" Or Matrix0.Columns.Item("otaxcode").Cells.Item(Row).Specific.Selected Is Nothing Or CDbl(Matrix0.Columns.Item("total").Cells.Item(Row).Specific.string) <= 0 Or Matrix0.Columns.Item("branch").Cells.Item(Row).Specific.Selected Is Nothing Or Matrix0.Columns.Item("docdate").Cells.Item(Row).Specific.string = "" Then
                                objaddon.objapplication.SetStatusBarMessage("Please update the line level values on row: " & Row, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False : Exit Sub
                            End If
                        End If

                    End If
                Next
                RemoveLastrow(Matrix0, "vcode")
                Dim ValidCode As String
                For i As Integer = 1 To Matrix0.VisualRowCount
                    If Not Matrix0.Columns.Item("branch").Cells.Item(i).Specific.Selected Is Nothing And Matrix0.Columns.Item("vcode").Cells.Item(i).Specific.string <> "" Then
                        If objaddon.HANA Then
                            ValidCode = objaddon.objglobalmethods.getSingleValue("select 1 from OCRD T0 inner join CRD8 T1 on T0.""CardCode""=T1.""CardCode"" where T1.""BPLId""='" & Matrix0.Columns.Item("branch").Cells.Item(i).Specific.Selected.Value & "' and T0.""CardCode""='" & Matrix0.Columns.Item("vcode").Cells.Item(i).Specific.string & "'")
                        Else
                            ValidCode = objaddon.objglobalmethods.getSingleValue("select 1 from OCRD T0 inner join CRD8 T1 on T0.CardCode=T1.CardCode where T1.BPLId='" & Matrix0.Columns.Item("branch").Cells.Item(i).Specific.Selected.Value & "' and T0.CardCode='" & Matrix0.Columns.Item("vcode").Cells.Item(i).Specific.string & "'")
                        End If
                        If ValidCode <> "1" Then
                            objaddon.objapplication.StatusBar.SetText("Selected Branch is not assigned for the vendor " & Matrix0.Columns.Item("vcode").Cells.Item(i).Specific.string & " on Line " & i & " .Please check...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Matrix0.Columns.Item("vcode").Cells.Item(i).Click()
                            BubbleEvent = False : Exit Sub
                        End If
                    End If
                Next
                If Create_AP_Service_Invoice() = False Then
                    objaddon.objapplication.StatusBar.SetText("Error occurred in A/P Service Invoice...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    BubbleEvent = False : Exit Sub
                Else
                    objaddon.objapplication.StatusBar.SetText("A/P Service Invoice Created Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                End If
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

        Private Sub Matrix0_ChooseFromListBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Matrix0.ChooseFromListBefore
            Try
                If pVal.ActionSuccess = True Then Exit Sub
                If pVal.ColUID = "vcode" Then
                    Try
                        Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFL_VC")
                        Dim oConds As SAPbouiCOM.Conditions
                        Dim oCond As SAPbouiCOM.Condition
                        Dim oEmptyConds As New SAPbouiCOM.Conditions
                        oCFL.SetConditions(oEmptyConds)
                        oConds = oCFL.GetConditions()

                        oCond = oConds.Add()
                        oCond.Alias = "CardType"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = "S"
                        oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                        oCond = oConds.Add()
                        oCond.Alias = "validFor"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = "Y"
                        oCFL.SetConditions(oConds)
                    Catch ex As Exception
                        SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try
                ElseIf pVal.ColUID = "glacc" Then
                    Try
                        Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFL_GL")
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
                        SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try
                ElseIf pVal.ColUID = "taxcode" Then
                    Try
                        Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFL_Tax")
                        Dim oConds As SAPbouiCOM.Conditions
                        Dim oCond As SAPbouiCOM.Condition
                        Dim oEmptyConds As New SAPbouiCOM.Conditions
                        oCFL.SetConditions(oEmptyConds)
                        oConds = oCFL.GetConditions()

                        'oCond = oConds.Add()
                        'oCond.Alias = "Name"
                        'oCond.Operation = SAPbouiCOM.BoConditionOperation.co_CONTAIN
                        'oCond.CondVal = "GST"

                        oCFL.SetConditions(oConds)
                    Catch ex As Exception
                        SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try
                ElseIf pVal.ColUID = "cc1" Then
                    Try
                        Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFL_C1")
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
                        Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFL_C2")
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
                        Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFL_C3")
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
                        Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFL_C4")
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
                        Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFL_C5")
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
                If pVal.ColUID = "vcode" And pVal.ActionSuccess = True Then
                    Try
                        pCFL = pVal
                        If pCFL.SelectedObjects Is Nothing Then Exit Sub

                        Try
                            Matrix0.Columns.Item("vcode").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("CardCode").Cells.Item(0).Value
                        Catch ex As Exception
                            'Matrix0.Columns.Item("vcode").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("CardCode").Cells.Item(0).Value
                        End Try
                        Try
                            Matrix0.Columns.Item("vname").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("CardName").Cells.Item(0).Value
                        Catch ex As Exception
                            'Matrix0.Columns.Item("vname").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("CardName").Cells.Item(0).Value
                        End Try
                        'Dim cmbbranch As SAPbouiCOM.ComboBox
                        'cmbbranch = Matrix0.Columns.Item("branch").Cells.Item(pVal.Row).Specific
                        'objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        'If objaddon.HANA Then
                        '    objRs.DoQuery("Select T0.""BPLId"",T0.""BPLName"" from OBPL T0 join CRD8 T1 on T0.""BPLId""=T1.""BPLId"" where ifnull(T0.""Disabled"",'N') <>'Y' and T1.""CardCode""='" & pCFL.SelectedObjects.Columns.Item("CardCode").Cells.Item(0).Value & "'")
                        'Else
                        '    objRs.DoQuery("Select T0.BPLId,T0.BPLName from OBPL T0 join CRD8 T1 on T0.BPLId=T1.BPLId where isnull(T0.Disabled,'N') <>'Y' and T1.CardCode='" & pCFL.SelectedObjects.Columns.Item("CardCode").Cells.Item(0).Value & "'")
                        'End If
                        'If objRs.RecordCount > 0 Then
                        '    Dim j As Integer = 0
                        '    If cmbbranch.ValidValues.Count > 0 Then
                        '        While j <= cmbbranch.ValidValues.Count - 1
                        '            cmbbranch.ValidValues.Remove(j, SAPbouiCOM.BoSearchKey.psk_Index)
                        '        End While
                        '    End If
                        '    For i As Integer = 0 To objRs.RecordCount - 1
                        '        cmbbranch.ValidValues.Add(objRs.Fields.Item(0).Value.ToString, objRs.Fields.Item(1).Value.ToString)
                        '        objRs.MoveNext()
                        '    Next
                        'End If
                        'objform.Update()
                    Catch ex As Exception
                    End Try
                ElseIf pVal.ColUID = "glacc" And pVal.ActionSuccess = True Then
                    Try
                        pCFL = pVal
                        If pCFL.SelectedObjects Is Nothing Then Exit Sub
                        Try
                            Matrix0.Columns.Item("glacc").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("AcctCode").Cells.Item(0).Value
                        Catch ex As Exception
                            'Matrix0.Columns.Item("glacc").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("AcctCode").Cells.Item(0).Value
                        End Try
                        Try
                            Matrix0.Columns.Item("glaccn").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("AcctName").Cells.Item(0).Value
                        Catch ex As Exception
                            'Matrix0.Columns.Item("glaccn").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("AcctName").Cells.Item(0).Value
                        End Try
                        'objform.Update()
                    Catch ex As Exception
                    End Try
                ElseIf pVal.ColUID = "taxcode" And pVal.ActionSuccess = True Then
                    Try
                        pCFL = pVal
                        If Not pCFL.SelectedObjects Is Nothing Then
                            Try
                                Matrix0.Columns.Item("taxcode").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("Code").Cells.Item(0).Value
                            Catch ex As Exception
                                Matrix0.Columns.Item("taxcode").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("Code").Cells.Item(0).Value
                            End Try
                        End If
                    Catch ex As Exception
                    End Try
                ElseIf pVal.ColUID = "saccode" And pVal.ActionSuccess = True Then
                    Try
                        pCFL = pVal
                        If Not pCFL.SelectedObjects Is Nothing Then
                            Try
                                Matrix0.Columns.Item("saccode").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("ServCode").Cells.Item(0).Value
                            Catch ex As Exception
                                Matrix0.Columns.Item("saccode").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("ServCode").Cells.Item(0).Value
                            End Try
                        End If
                    Catch ex As Exception
                    End Try
                ElseIf pVal.ColUID = "cc1" And pVal.ActionSuccess = True Then
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
                'objaddon.objapplication.Menus.Item("1300").Activate()
                GC.Collect()
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End Try
        End Sub

        Private Sub Matrix0_LostFocusAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.LostFocusAfter
            Try
                If pVal.Row = 0 Then Exit Sub
                If pVal.ActionSuccess = False Then Exit Sub
                Select Case pVal.ColUID
                    Case "vcode"
                        If Matrix0.Columns.Item("vcode").Cells.Item(pVal.Row).Specific.String <> "" Then
                            Try
                                If objaddon.HANA Then
                                    Matrix0.Columns.Item("vname").Cells.Item(pVal.Row).Specific.String = objaddon.objglobalmethods.getSingleValue("Select ""CardName"" from OCRD where ""CardCode""='" & Matrix0.Columns.Item("vcode").Cells.Item(pVal.Row).Specific.String & "' ")
                                Else
                                    Matrix0.Columns.Item("vname").Cells.Item(pVal.Row).Specific.String = objaddon.objglobalmethods.getSingleValue("Select CardName from OCRD where CardCode='" & Matrix0.Columns.Item("vcode").Cells.Item(pVal.Row).Specific.String & "' ")
                                End If

                            Catch ex As Exception
                            End Try
                            objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "vcode", "#")
                        End If
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
                    '    If Matrix0.Columns.Item("distrule").Cells.Item(pVal.Row).Specific.String = "" Then
                    '        If CostCenter = "U" Then
                    '            Dim oform As New FrmDistRule
                    '            oform.Show()
                    '        End If
                    '    End If
                    Case "glacc"
                        Try
                            If Matrix0.Columns.Item("glacc").Cells.Item(pVal.Row).Specific.String <> "" Then
                                If objaddon.HANA Then
                                    Matrix0.Columns.Item("glaccn").Cells.Item(pVal.Row).Specific.String = objaddon.objglobalmethods.getSingleValue("Select ""AcctName"" from OACT where ""AcctCode""='" & Matrix0.Columns.Item("glacc").Cells.Item(pVal.Row).Specific.String & "' ")
                                Else
                                    Matrix0.Columns.Item("glaccn").Cells.Item(pVal.Row).Specific.String = objaddon.objglobalmethods.getSingleValue("Select AcctName from OACT where AcctCode='" & Matrix0.Columns.Item("glacc").Cells.Item(pVal.Row).Specific.String & "' ")
                                End If
                            End If
                        Catch ex As Exception
                        End Try
                    Case "taxcode", "otaxcode", "total"
                        Dim GTotal As Double = 0
                        If Localization = "IN" Then
                            If Matrix0.Columns.Item("taxcode").Cells.Item(pVal.Row).Specific.String <> "" And Matrix0.Columns.Item("total").Cells.Item(pVal.Row).Specific.String > 0 Then
                                'Matrix0.Columns.Item("gtotal").Cells.Item(pVal.Row).Specific.String = CDbl(CDbl(Matrix0.Columns.Item("total").Cells.Item(pVal.Row).Specific.String) + ((CDbl(Matrix0.Columns.Item("total").Cells.Item(pVal.Row).Specific.String) * GetTaxAmount(Matrix0.Columns.Item("taxcode").Cells.Item(pVal.Row).Specific.String)) / 100))
                                GTotal = CDbl(CDbl(Matrix0.Columns.Item("total").Cells.Item(pVal.Row).Specific.String) + ((CDbl(Matrix0.Columns.Item("total").Cells.Item(pVal.Row).Specific.String) * GetTaxAmount(Matrix0.Columns.Item("taxcode").Cells.Item(pVal.Row).Specific.String)) / 100))
                                Matrix0.Columns.Item("gtotal").Cells.Item(pVal.Row).Specific.String = CStr(GTotal)
                            End If
                        Else
                            If Not Matrix0.Columns.Item("otaxcode").Cells.Item(pVal.Row).Specific.Selected Is Nothing And Matrix0.Columns.Item("total").Cells.Item(pVal.Row).Specific.String > 0 Then
                                GTotal = CDbl(CDbl(Matrix0.Columns.Item("total").Cells.Item(pVal.Row).Specific.String) + ((CDbl(Matrix0.Columns.Item("total").Cells.Item(pVal.Row).Specific.String) * GetTaxAmount(Matrix0.Columns.Item("otaxcode").Cells.Item(pVal.Row).Specific.Selected.Value)) / 100))
                                Matrix0.Columns.Item("gtotal").Cells.Item(pVal.Row).Specific.String = CStr(GTotal)
                                'Matrix0.Columns.Item("gtotal").Cells.Item(pVal.Row).Specific.String = CDbl(CDbl(Matrix0.Columns.Item("total").Cells.Item(pVal.Row).Specific.String) + ((CDbl(Matrix0.Columns.Item("total").Cells.Item(pVal.Row).Specific.String) * GetTaxAmount(Matrix0.Columns.Item("otaxcode").Cells.Item(pVal.Row).Specific.Selected.Value)) / 100))
                            End If
                        End If

                End Select
                'Matrix0.Columns.Item("total").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                'Matrix0.Columns.Item("gtotal").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                objform.Freeze(True)
                Matrix0.AutoResizeColumns()
                objform.Freeze(False)
            Catch ex As Exception
                objform.Freeze(False)
            End Try

        End Sub

        Private Sub ComboBox0_ComboSelectAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles ComboBox0.ComboSelectAfter
            Try
                odbdsHeader.SetValue("DocNum", 0, objaddon.objglobalmethods.GetDocNum("MIAPSI", CInt(ComboBox0.Selected.Value)))
            Catch ex As Exception

            End Try

        End Sub

        Private Function Create_AP_Service_Invoice() As Boolean
            Try
                Dim DocEntry As String, BranchEnabled, saccode, Series As String
                Dim objPurchaseInvoice As SAPbobsCOM.Documents
                Dim objEdit As SAPbouiCOM.EditText
                Dim MBAPDocNum As Long
                Dim TFlag As Boolean = False
                If objaddon.HANA Then
                    BranchEnabled = objaddon.objglobalmethods.getSingleValue("select ""MltpBrnchs"" from OADM")
                Else
                    BranchEnabled = objaddon.objglobalmethods.getSingleValue("select MltpBrnchs from OADM")
                End If
                If Not BranchEnabled = "Y" Then objaddon.objapplication.StatusBar.SetText("Branch not enabled...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning) : Exit Function
                objaddon.objapplication.StatusBar.SetText("A/P Service Invoice Creating Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                If Not objaddon.objcompany.InTransaction Then objaddon.objcompany.StartTransaction()
                For Row As Integer = 1 To Matrix0.VisualRowCount
                    If Matrix0.Columns.Item("vcode").Cells.Item(Row).Specific.string <> "" And Matrix0.Columns.Item("tentry").Cells.Item(Row).Specific.string = "" Then
                        objPurchaseInvoice = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                        Dim DocDate As Date = Date.ParseExact(EditText1.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                        objEdit = Matrix0.Columns.Item("docdate").Cells.Item(Row).Specific
                        Dim TaxDate As Date = Date.ParseExact(objEdit.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                        If objaddon.HANA Then
                            strSQL = objaddon.objglobalmethods.getSingleValue("SELECT To_Varchar(ADD_DAYS('" & DocDate.ToString("yyyyMMdd") & "', ""ExtraDays""),'yyyyMMdd') as ""DueDate"" FROM OCRD T0 join OCTG T1 on T0.""GroupNum""=T1.""GroupNum"" where T0.""CardCode""='" & Matrix0.Columns.Item("vcode").Cells.Item(Row).Specific.string & "'")
                        Else
                            strSQL = objaddon.objglobalmethods.getSingleValue("SELECT Format(DATEADD(dd,ExtraDays,'" & DocDate.ToString("yyyyMMdd") & "'),'yyyyMMdd') as DueDate FROM OCRD T0 join OCTG T1 on T0.GroupNum=T1.GroupNum where T0.CardCode='" & Matrix0.Columns.Item("vcode").Cells.Item(Row).Specific.string & "'")
                        End If
                        'Dim DueDate As Date = Date.ParseExact(EditText2.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                        Dim DueDate As Date = Date.ParseExact(strSQL, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                        MBAPDocNum = objform.BusinessObject.GetNextSerialNumber(objform.Items.Item("Series").Specific.Selected.value)
                        objPurchaseInvoice.CardCode = Matrix0.Columns.Item("vcode").Cells.Item(Row).Specific.string
                        objPurchaseInvoice.DocDate = DocDate
                        objPurchaseInvoice.DocDueDate = DueDate
                        objPurchaseInvoice.TaxDate = TaxDate 'DocDate
                        objPurchaseInvoice.JournalMemo = "Auto-Generated->  " & Now.ToString
                        objPurchaseInvoice.Comments = Matrix0.Columns.Item("remarks").Cells.Item(Row).Specific.string '& " Multi-Branch Service Invoice DocNum-> " & CStr(MBAPDocNum)
                        If Matrix0.Columns.Item("refno").Cells.Item(Row).Specific.string <> "" Then objPurchaseInvoice.UserFields.Fields.Item("U_ymhbpref").Value = Matrix0.Columns.Item("refno").Cells.Item(Row).Specific.string
                        If Localization = "IN" Then
                            If objaddon.HANA Then
                                Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 ""Series"" FROM NNM1 WHERE ""ObjectCode""='18' and ""Indicator""=(Select ""Indicator"" from OFPR where '" & DocDate.ToString("yyyyMMdd") & "' Between ""F_RefDate"" and ""T_RefDate"") " & 'CURRENT_DATE
                                                                                  "  and ifnull(""U_APSeries"",'Y')='Y' and ""BPLId""='" & Matrix0.Columns.Item("branch").Cells.Item(Row).Specific.Selected.Value & "'") 'and ""DocSubType""='GA'
                            Else
                                Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 Series FROM NNM1 WHERE ObjectCode='18' and Indicator=(Select Indicator from OFPR where '" & DocDate.ToString("yyyyMMdd") & "' Between F_RefDate and T_RefDate) " &
                                                                                  " and DocSubType='GA' and isnull(U_APSeries,'Y')='Y' and BPLId='" & Matrix0.Columns.Item("branch").Cells.Item(Row).Specific.Selected.Value & "'")
                            End If
                        Else
                            If objaddon.HANA Then
                                Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 ""Series"" FROM NNM1 WHERE ""ObjectCode""='18' and ""Indicator""=(Select ""Indicator"" from OFPR where '" & DocDate.ToString("yyyyMMdd") & "' Between ""F_RefDate"" and ""T_RefDate"") " &
                                                                                  " and ifnull(""U_APSeries"",'Y')='Y' and ""BPLId""='" & Matrix0.Columns.Item("branch").Cells.Item(Row).Specific.Selected.Value & "'")
                            Else
                                Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 Series FROM NNM1 WHERE ObjectCode='18' and Indicator=(Select Indicator from OFPR where '" & DocDate.ToString("yyyyMMdd") & "' Between F_RefDate and T_RefDate) " &
                                                                                  " and isnull(U_APSeries,'Y')='Y' and BPLId='" & Matrix0.Columns.Item("branch").Cells.Item(Row).Specific.Selected.Value & "'")
                            End If
                        End If

                        If Series = "" Then objaddon.objapplication.StatusBar.SetText("Numbering Series not found.Please check...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning) : Exit Function
                        objPurchaseInvoice.Series = Series
                        'objPurchaseInvoice.Series = 110
                        objPurchaseInvoice.UserFields.Fields.Item("U_MBAPNo").Value = CStr(MBAPDocNum)
                        If BranchEnabled = "Y" Then
                            objPurchaseInvoice.BPL_IDAssignedToInvoice = Matrix0.Columns.Item("branch").Cells.Item(Row).Specific.Selected.Value
                        End If
                        objPurchaseInvoice.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service
                        objPurchaseInvoice.Lines.AccountCode = Matrix0.Columns.Item("glacc").Cells.Item(Row).Specific.string
                        objPurchaseInvoice.Lines.ItemDescription = Matrix0.Columns.Item("desc").Cells.Item(Row).Specific.string
                        If Matrix0.Columns.Item("cc1").Cells.Item(Row).Specific.string <> "" Then objPurchaseInvoice.Lines.CostingCode = Matrix0.Columns.Item("cc1").Cells.Item(Row).Specific.string
                        If Matrix0.Columns.Item("cc2").Cells.Item(Row).Specific.string <> "" Then objPurchaseInvoice.Lines.CostingCode2 = Matrix0.Columns.Item("cc2").Cells.Item(Row).Specific.string
                        If Matrix0.Columns.Item("cc3").Cells.Item(Row).Specific.string <> "" Then objPurchaseInvoice.Lines.CostingCode3 = Matrix0.Columns.Item("cc3").Cells.Item(Row).Specific.string
                        If Matrix0.Columns.Item("cc4").Cells.Item(Row).Specific.string <> "" Then objPurchaseInvoice.Lines.CostingCode4 = Matrix0.Columns.Item("cc4").Cells.Item(Row).Specific.string
                        If Matrix0.Columns.Item("cc5").Cells.Item(Row).Specific.string <> "" Then objPurchaseInvoice.Lines.CostingCode5 = Matrix0.Columns.Item("cc5").Cells.Item(Row).Specific.string
                        If Matrix0.Columns.Item("lrefno").Cells.Item(Row).Specific.string <> "" Then objPurchaseInvoice.Lines.UserFields.Fields.Item("U_ymhbpref").Value = Matrix0.Columns.Item("lrefno").Cells.Item(Row).Specific.string
                        If Localization = "IN" Then
                            objPurchaseInvoice.Lines.TaxCode = Matrix0.Columns.Item("taxcode").Cells.Item(Row).Specific.string
                            If objaddon.HANA Then
                                saccode = objaddon.objglobalmethods.getSingleValue("select ""AbsEntry"" from OSAC where ""ServCode""='" & Matrix0.Columns.Item("saccode").Cells.Item(Row).Specific.string & "'")
                            Else
                                saccode = objaddon.objglobalmethods.getSingleValue("select AbsEntry from OSAC where ServCode='" & Matrix0.Columns.Item("saccode").Cells.Item(Row).Specific.string & "'")
                            End If
                            objPurchaseInvoice.Lines.SACEntry = CInt(saccode)
                            If objaddon.HANA Then
                                objPurchaseInvoice.Lines.LocationCode = objaddon.objglobalmethods.getSingleValue("Select T1.""Location"" from OBPL T0 join OWHS T1 On T1.""BPLid""=T0.""BPLId"" And T0.""DflWhs""=T1.""WhsCode"" where T0.""BPLId""='" & Matrix0.Columns.Item("branch").Cells.Item(Row).Specific.Selected.Value & "'")
                            Else
                                objPurchaseInvoice.Lines.LocationCode = objaddon.objglobalmethods.getSingleValue("Select T1.Location from OBPL T0 join OWHS T1 On T1.BPLid=T0.BPLId And T0.DflWhs=T1.WhsCode where T0.BPLId='" & Matrix0.Columns.Item("branch").Cells.Item(Row).Specific.Selected.Value & "'")
                            End If
                        Else
                            objPurchaseInvoice.Lines.VatGroup = Matrix0.Columns.Item("otaxcode").Cells.Item(Row).Specific.Selected.Value
                        End If
                        objPurchaseInvoice.Lines.LineTotal = Matrix0.Columns.Item("total").Cells.Item(Row).Specific.string
                        'objPurchaseInvoice.Lines.GrossTotal = Matrix0.Columns.Item("gtotal").Cells.Item(Row).Specific.string
                        objPurchaseInvoice.Lines.Add()

                        If objPurchaseInvoice.Add() <> 0 Then
                            TFlag = True
                            objaddon.objapplication.SetStatusBarMessage("A/P Service Invoice: " & objaddon.objcompany.GetLastErrorDescription & "-" & objaddon.objcompany.GetLastErrorCode & " on Line: " & Row, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                            objaddon.objapplication.MessageBox("A/P Service Invoice: " & objaddon.objcompany.GetLastErrorDescription & "-" & objaddon.objcompany.GetLastErrorCode & " on Line: " & Row,, "OK")
                            Exit For
                        Else
                            'Dim sNewObjCode As String = ""
                            'objaddon.objcompany.GetNewObjectCode(sNewObjCode)
                            'Dim str = CLng(sNewObjCode)
                            DocEntry = objaddon.objcompany.GetNewObjectKey()
                            Matrix0.Columns.Item("tentry").Cells.Item(Row).Specific.String = DocEntry

                        End If
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objPurchaseInvoice)
                        GC.Collect()
                    End If
                Next
                If TFlag = True Then
                    If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    For i As Integer = 1 To Matrix0.VisualRowCount
                        If Matrix0.Columns.Item("tentry").Cells.Item(i).Specific.String <> "" Then
                            Matrix0.Columns.Item("tentry").Cells.Item(i).Specific.String = ""
                        End If
                    Next
                    'objaddon.objapplication.StatusBar.SetText("Error in A/P Service Invoice...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                Else
                    If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    'objaddon.objapplication.StatusBar.SetText("A/P Service Invoice Created Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    Return True
                End If

            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                'objaddon.objapplication.MessageBox(ex.Message, , "OK")
                Return False
                'If objAddOn.objCompany.InTransaction Then objAddOn.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End Try

        End Function

        Private Function GetTaxAmount(ByVal Tax As String) As String
            Try
                If objaddon.HANA Then
                    If Localization = "IN" Then
                        Tax = objaddon.objglobalmethods.getSingleValue("select ""Rate"" from OSTC where ""Code""='" & Tax & "' ")
                    Else
                        Tax = objaddon.objglobalmethods.getSingleValue("select ""Rate"" from OVTG where ""Code""='" & Tax & "' ")
                    End If
                Else
                    If Localization = "IN" Then
                        Tax = objaddon.objglobalmethods.getSingleValue("select Rate from OSTC where Code='" & Tax & "' ")
                    Else
                        Tax = objaddon.objglobalmethods.getSingleValue("select Rate from OVTG where Code='" & Tax & "' ")
                    End If

                End If
                Return Tax
            Catch ex As Exception
                Return ""
            End Try
        End Function

        Private Sub Matrix0_ValidateAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.ValidateAfter
            Try
                If pVal.ItemChanged = False Then Exit Sub
                Dim GTotal As Double = 0
                Select Case pVal.ColUID
                    Case "taxcode", "otaxcode", "total"
                        If Localization = "IN" Then
                            If Matrix0.Columns.Item("taxcode").Cells.Item(pVal.Row).Specific.String <> "" And Matrix0.Columns.Item("total").Cells.Item(pVal.Row).Specific.String > 0 Then
                                'Matrix0.Columns.Item("gtotal").Cells.Item(pVal.Row).Specific.String = CDbl(CDbl(Matrix0.Columns.Item("total").Cells.Item(pVal.Row).Specific.String) + ((CDbl(Matrix0.Columns.Item("total").Cells.Item(pVal.Row).Specific.String) * GetTaxAmount(Matrix0.Columns.Item("taxcode").Cells.Item(pVal.Row).Specific.String)) / 100))
                                GTotal = CDbl(CDbl(Matrix0.Columns.Item("total").Cells.Item(pVal.Row).Specific.String) + ((CDbl(Matrix0.Columns.Item("total").Cells.Item(pVal.Row).Specific.String) * GetTaxAmount(Matrix0.Columns.Item("taxcode").Cells.Item(pVal.Row).Specific.String)) / 100))
                            End If
                        Else
                            If Not Matrix0.Columns.Item("otaxcode").Cells.Item(pVal.Row).Specific.Selected Is Nothing And Matrix0.Columns.Item("total").Cells.Item(pVal.Row).Specific.String > 0 Then
                                GTotal = CDbl(CDbl(Matrix0.Columns.Item("total").Cells.Item(pVal.Row).Specific.String) + ((CDbl(Matrix0.Columns.Item("total").Cells.Item(pVal.Row).Specific.String) * GetTaxAmount(Matrix0.Columns.Item("otaxcode").Cells.Item(pVal.Row).Specific.Selected.Value)) / 100))
                                'Matrix0.Columns.Item("gtotal").Cells.Item(pVal.Row).Specific.String = CDbl(CDbl(Matrix0.Columns.Item("total").Cells.Item(pVal.Row).Specific.String) + ((CDbl(Matrix0.Columns.Item("total").Cells.Item(pVal.Row).Specific.String) * GetTaxAmount(Matrix0.Columns.Item("otaxcode").Cells.Item(pVal.Row).Specific.Selected.Value)) / 100))
                            End If
                        End If
                        Matrix0.Columns.Item("gtotal").Cells.Item(pVal.Row).Specific.String = CStr(GTotal)
                End Select

            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try

        End Sub

        Private Sub Form_DataAddBefore(ByRef pVal As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
            'Try
            '    If Create_AP_Service_Invoice() = False Then
            '        objaddon.objapplication.StatusBar.SetText("Error occurred in A/P Service Invoice...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '        BubbleEvent = False : Exit Sub
            '    Else
            '        objaddon.objapplication.StatusBar.SetText("A/P Service Invoice Created Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            '    End If
            'Catch ex As Exception

            'End Try

        End Sub

        Private Sub Form_DataLoadAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo)
            Try
                Matrix0.Columns.Item("branch").DisplayDesc = True
                Matrix0.AutoResizeColumns()
                objform.EnableMenu("1282", True)
                objform.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
            Catch ex As Exception

            End Try

        End Sub

        Private Sub EditText2_LostFocusAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText2.LostFocusAfter
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    Dim DocDate As Date = Date.ParseExact(EditText1.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                    Dim DelDate As Date = Date.ParseExact(EditText2.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                    If DateTime.Compare(DocDate, DelDate) > 0 Then
                        objaddon.objapplication.StatusBar.SetText("In ""Due Date"" field, enter date that is equal to or later than posting date.  Field: Delivery Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        EditText2.Item.Click()
                    End If
                    objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "vcode", "#")
                End If
            Catch ex As Exception
            End Try

        End Sub

        Private WithEvents Folder1 As SAPbouiCOM.Folder

        Private Sub Button0_PressedAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.PressedAfter
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                If pVal.ActionSuccess = True And objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    objaddon.objglobalmethods.LoadSeries(objform, odbdsHeader, "MIAPSI")
                    objform.Items.Item("tposdate").Specific.string = Now.Date.ToString("yyyyMMdd")
                    objform.Items.Item("tdocdate").Specific.string = Now.Date.ToString("yyyyMMdd")
                    objform.Items.Item("tremark").Specific.String = "Created By " & objaddon.objcompany.UserName & " on " & Now.ToString
                    objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "vcode", "#")
                End If
            Catch ex As Exception
            End Try

        End Sub

        Private Sub Matrix0_KeyDownAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.KeyDownAfter
            Try
                'Select Case pVal.ColUID
                '    Case "vcode"
                '        If pVal.CharPressed = 40 And pVal.Modifiers = SAPbouiCOM.BoModifiersEnum.mt_CTRL Then
                '            Matrix0.Columns.Item("vcode").Cells.Item(pVal.Row).Specific.String = Matrix0.Columns.Item("vcode").Cells.Item(pVal.Row - 1).Specific.String
                '            Matrix0.Columns.Item("vname").Cells.Item(pVal.Row).Specific.String = Matrix0.Columns.Item("vname").Cells.Item(pVal.Row - 1).Specific.String
                '        End If
                '    Case "glacc"
                '        If pVal.CharPressed = 40 And pVal.Modifiers = SAPbouiCOM.BoModifiersEnum.mt_CTRL Then
                '            Matrix0.Columns.Item("glacc").Cells.Item(pVal.Row).Specific.String = Matrix0.Columns.Item("glacc").Cells.Item(pVal.Row - 1).Specific.String
                '            Matrix0.Columns.Item("glaccn").Cells.Item(pVal.Row).Specific.String = Matrix0.Columns.Item("glaccn").Cells.Item(pVal.Row - 1).Specific.String
                '        End If
                'End Select

                Dim ColID As Integer = Matrix0.GetCellFocus().ColumnIndex
                If pVal.CharPressed = 38 And pVal.Modifiers = SAPbouiCOM.BoModifiersEnum.mt_None Then  'up
                    Matrix0.SetCellFocus(pVal.Row - 1, ColID)
                    Matrix0.SelectRow(pVal.Row - 1, True, False)
                ElseIf pVal.CharPressed = 40 And pVal.Modifiers = SAPbouiCOM.BoModifiersEnum.mt_None Then 'down
                    Matrix0.SetCellFocus(pVal.Row + 1, ColID)
                    Matrix0.SelectRow(pVal.Row + 1, True, False)
                    'ElseIf pVal.CharPressed = 37 And pVal.Modifiers = SAPbouiCOM.BoModifiersEnum.mt_None Then 'Left
                    '    Matrix0.SetCellFocus(pVal.Row, ColID - 1)
                    'ElseIf pVal.CharPressed = 39 And pVal.Modifiers = SAPbouiCOM.BoModifiersEnum.mt_None Then 'Right
                    '    Matrix0.SetCellFocus(pVal.Row, ColID + 1)

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

        Private Sub Matrix0_LinkPressedBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Matrix0.LinkPressedBefore
            Try
                Select Case pVal.ColUID
                    Case "distrule"
                        If Matrix0.Columns.Item("distrule").Cells.Item(pVal.Row).Specific.String <> "" Then
                            Link_Value = Matrix0.Columns.Item("distrule").Cells.Item(pVal.Row).Specific.String
                            Dim oform As New FrmDistRule
                            oform.Show()
                        End If
                End Select
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Folder0_PressedAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Folder0.PressedAfter
            Try
                objform.Settings.MatrixUID = "mtxcont"
            Catch ex As Exception
            End Try

        End Sub

        Private Sub EditText1_LostFocusAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText1.LostFocusAfter
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    'If EditText1.Value <> "" Then
                    EditText3.Value = EditText1.Value
                    'End If
                End If
            Catch ex As Exception

            End Try

        End Sub
    End Class

End Namespace
