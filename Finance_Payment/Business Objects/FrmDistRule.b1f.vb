Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace Finance_Payment
    <FormAttribute("DistRule", "Business Objects/FrmDistRule.b1f")>
    Friend Class FrmDistRule
        Inherits UserFormBase
        Public WithEvents objform As SAPbouiCOM.Form
        Private WithEvents pCFL As SAPbouiCOM.ISBOChooseFromListEventArg
        Dim FormCount As Integer = 0
        Dim strSQL As String
        Private WithEvents objDTable As SAPbouiCOM.DataTable
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("1").Specific, SAPbouiCOM.Button)
            Me.Button1 = CType(Me.GetItem("2").Specific, SAPbouiCOM.Button)
            Me.Matrix0 = CType(Me.GetItem("mtxdata").Specific, SAPbouiCOM.Matrix)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler CloseBefore, AddressOf Me.Form_CloseBefore
            AddHandler DataLoadAfter, AddressOf Me.Form_DataLoadAfter

        End Sub

        Private WithEvents Button0 As SAPbouiCOM.Button

        Private Sub OnCustomInitialize()
            Try
                objform = objaddon.objapplication.Forms.GetForm("DistRule", Me.FormCount)
                'objform = objaddon.objapplication.Forms.ActiveForm
                objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                If objaddon.HANA Then
                    strSQL = "Select ""DimDesc"" from ODIM where Ifnull(""DimActive"",'')='Y' order by ""DimCode"""
                Else
                    strSQL = "Select DimDesc from ODIM where Isnull(DimActive,'')='Y' order by DimCode"
                End If
                bModal = True
                'If objform.DataSources.DataTables.Count.Equals(0) Then
                '    objform.DataSources.DataTables.Add("DT_List")
                'Else
                '    objform.DataSources.DataTables.Item("DT_List").Clear()
                'End If
                objform.DataSources.DataTables.Item("DT_0").Clear()
                objDTable = objform.DataSources.DataTables.Item("DT_0")
                objDTable.Clear()
                objDTable.ExecuteQuery(strSQL)
                Matrix0.Clear()
                Matrix0.LoadFromDataSourceEx()
                Matrix0.AutoResizeColumns()
                If Link_Value <> "-1" Then
                    objform.Freeze(True)
                    Dim CostCodes() As String = Link_Value.Split(";")
                    For j As Integer = 0 To CostCodes.Length - 1
                        Matrix0.Columns.Item("drcode").Cells.Item(j + 1).Specific.String = CostCodes(j)
                    Next
                    objform.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                    objform.Freeze(False)
                    Link_Value = "-1" : Exit Sub
                End If
            Catch ex As Exception
                objform.Freeze(False)
            End Try
        End Sub

        Private WithEvents Button1 As SAPbouiCOM.Button
        Private WithEvents Matrix0 As SAPbouiCOM.Matrix

        Private Sub Button0_ClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.ClickAfter
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objform.Close() : Exit Sub
                'Dim objSIform As SAPbouiCOM.Form
                Dim objMatrix As SAPbouiCOM.Matrix
                'objSIform = objaddon.objapplication.Forms.GetForm("MBAPSI", FormCount)
                objMatrix = OEForm.Items.Item("mtxcont").Specific
                Dim Row As Integer = objMatrix.GetCellFocus.rowIndex()
                Dim code As String = ""
                For i As Integer = 1 To Matrix0.VisualRowCount
                    If Matrix0.Columns.Item("drcode").Cells.Item(i).Specific.String <> "" Then
                        If i = 1 Then
                            code = Matrix0.Columns.Item("drcode").Cells.Item(i).Specific.String
                        Else
                            code += ";" + Matrix0.Columns.Item("drcode").Cells.Item(i).Specific.String
                        End If
                    End If
                Next
                objMatrix.Columns.Item("distrule").Cells.Item(Row).Specific.String = code
                Dim CostCodes() As String = code.Split(";")
                If CostCodes.Count > 0 Then
                    If CostCodes.ElementAtOrDefault(0) <> "" Then objMatrix.Columns.Item("cc1").Cells.Item(Row).Specific.String = CostCodes(0)
                    If CostCodes.ElementAtOrDefault(1) <> "" Then objMatrix.Columns.Item("cc2").Cells.Item(Row).Specific.String = CostCodes(1)
                    If CostCodes.ElementAtOrDefault(2) <> "" Then objMatrix.Columns.Item("cc3").Cells.Item(Row).Specific.String = CostCodes(2)
                    If CostCodes.ElementAtOrDefault(3) <> "" Then objMatrix.Columns.Item("cc4").Cells.Item(Row).Specific.String = CostCodes(3)
                    If CostCodes.ElementAtOrDefault(4) <> "" Then objMatrix.Columns.Item("cc5").Cells.Item(Row).Specific.String = CostCodes(4)
                End If
                objMatrix.AutoResizeColumns()
                objform.Close()

            Catch ex As Exception
                objform.Freeze(False)
            End Try

        End Sub

        Private Sub Form_CloseBefore(pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean)
            Try
                If pVal.InnerEvent = False Then Exit Sub
                BubbleEvent = False
            Catch ex As Exception

            End Try


        End Sub

        Private Sub Button1_ClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button1.ClickAfter
            Try
                objform.Close()
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Matrix0_ChooseFromListBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Matrix0.ChooseFromListBefore
            Try
                If pVal.ActionSuccess = True Then Exit Sub
                Try
                    Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFL_0")
                    Dim oConds As SAPbouiCOM.Conditions
                    Dim oCond As SAPbouiCOM.Condition
                    Dim oEmptyConds As New SAPbouiCOM.Conditions
                    oCFL.SetConditions(oEmptyConds)
                    oConds = oCFL.GetConditions()
                    Dim Code As String
                    If objaddon.HANA Then
                        Code = objaddon.objglobalmethods.getSingleValue("Select ""DimCode"" from ODIM where ""DimDesc""='" & Matrix0.Columns.Item("dim").Cells.Item(pVal.Row).Specific.String & "'")
                    Else
                        Code = objaddon.objglobalmethods.getSingleValue("Select DimCode from ODIM where DimDesc='" & Matrix0.Columns.Item("dim").Cells.Item(pVal.Row).Specific.String & "'")
                    End If

                    oCond = oConds.Add()
                    oCond.Alias = "DimCode"
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCond.CondVal = Code
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
                If Not pCFL.SelectedObjects Is Nothing Then
                    Try
                        Matrix0.Columns.Item("drcode").Cells.Item(pVal.Row).Specific.String = Trim(pCFL.SelectedObjects.Columns.Item("OcrCode").Cells.Item(0).Value.ToString)
                        Matrix0.Columns.Item("drname").Cells.Item(pVal.Row).Specific.String = Trim(pCFL.SelectedObjects.Columns.Item("OcrName").Cells.Item(0).Value.ToString)
                    Catch ex As Exception
                        Matrix0.Columns.Item("drcode").Cells.Item(pVal.Row).Specific.String = Trim(pCFL.SelectedObjects.Columns.Item("OcrCode").Cells.Item(0).Value.ToString)
                        Matrix0.Columns.Item("drname").Cells.Item(pVal.Row).Specific.String = Trim(pCFL.SelectedObjects.Columns.Item("OcrName").Cells.Item(0).Value.ToString)
                        'objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End Try
                    If objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    Matrix0.AutoResizeColumns()
                End If
            Catch ex As Exception
                ' objaddon.objapplication.StatusBar.SetText("CFL_After: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End Try

        End Sub

        Private Sub Form_DataLoadAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo)
            'Try
            '    objform.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
            'Catch ex As Exception

            'End Try

        End Sub

    End Class
End Namespace
