Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports SAPbobsCOM
Imports SAPbouiCOM.Framework
Namespace Finance_Payment
    Public Class ClsJE
        Public Const Formtype = "392"
        Dim objform, objUDFForm As SAPbouiCOM.Form
        Dim objmatrix As SAPbouiCOM.Matrix
        Dim strSQL As String
        Dim objRs As SAPbobsCOM.Recordset
        Public WithEvents odbdsHeader As SAPbouiCOM.DBDataSource

        Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.Item(FormUID)
                objmatrix = objform.Items.Item("76").Specific
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                If pVal.BeforeAction Then
                        Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_CLICK

                        Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            'If objaddon.objapplication.Menus.Item("6913").Checked = True Then
                            '    objaddon.objapplication.SendKeys("^+U")
                            'Else
                            '    MsgBox("1")
                            'End If
                    End Select
                Else
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                            objUDFForm = objaddon.objapplication.Forms.Item(objform.UDFFormUID)
                            objUDFForm.Items.Item("lnkirnum").Top = objUDFForm.Items.Item("U_IntRecEntry").Top + 2
                            objUDFForm.Items.Item("lnkreco").Top = objUDFForm.Items.Item("U_ReconNum").Top + 2

                        Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD


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
                odbdsHeader = objform.DataSources.DBDataSources.Item("OJDT")
                If BusinessObjectInfo.BeforeAction = True Then
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                    End Select
                Else
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                            If BusinessObjectInfo.ActionSuccess Then
                                If odbdsHeader.GetValue("U_IEntry", 0) <> "" Then
                                    strSQL = objaddon.objglobalmethods.getSingleValue("Select ""Canceled"" from ORCT Where ""DocEntry""=" & odbdsHeader.GetValue("U_IEntry", 0) & "")
                                    If strSQL = "Y" Then Exit Sub
                                    If Cancel_Payment(BusinessObjectInfo.FormUID, "24", odbdsHeader.GetValue("U_IEntry", 0), odbdsHeader.GetValue("TransId", 0)) = False Then
                                        objaddon.objapplication.StatusBar.SetText("Payment Cancellation Failed. Kindly Cancel the Incoming Payment Manually!!!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        objaddon.objapplication.MessageBox("Payment Cancellation Failed. Kindly Cancel the Incoming Payment Manually!!!", 1, "OK")
                                        objUDFForm = objaddon.objapplication.Forms.Item(objform.UDFFormUID)
                                        objUDFForm.Items.Item("1000012").Click()
                                    End If
                                ElseIf odbdsHeader.GetValue("U_OEntry", 0) <> "" Then
                                    strSQL = objaddon.objglobalmethods.getSingleValue("Select ""Canceled"" from OVPM Where ""DocEntry""=" & odbdsHeader.GetValue("U_OEntry", 0) & "")
                                    If strSQL = "Y" Then Exit Sub
                                    If Cancel_Payment(BusinessObjectInfo.FormUID, "46", odbdsHeader.GetValue("U_OEntry", 0), odbdsHeader.GetValue("TransId", 0)) = False Then
                                        objaddon.objapplication.StatusBar.SetText("Payment Cancellation Failed. Kindly Cancel the Outgoing Payment Manually!!!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        objaddon.objapplication.MessageBox("Payment Cancellation Failed. Kindly Cancel the Outgoing Payment Manually", 1, "OK")
                                        objUDFForm = objaddon.objapplication.Forms.Item(objform.UDFFormUID)
                                        objUDFForm.Items.Item("1000010").Click()
                                    End If
                                End If

                            End If
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                            If objmatrix.Columns.Item("U_TBranch").Editable = True Then objmatrix.Columns.Item("U_TBranch").Editable = False
                            If objaddon.objapplication.Menus.Item("6913").Checked = False Then
                                objaddon.objapplication.SendKeys("^+U")
                            End If
                            'objaddon.objapplication.Menus.Item("1304").Activate()
                    End Select
                End If

            Catch ex As Exception
                'objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        End Sub

        Private Function Cancel_Payment(ByVal FormUID As String, ByVal ObjType As String, ByVal PaymentEntry As String, ByVal JE As String)
            Try
                objform = objaddon.objapplication.Forms.Item(FormUID)
                Dim objPayment As SAPbobsCOM.Payments = Nothing
                Dim Header As String = ""
                If ObjType = "24" Then
                    objPayment = objaddon.objcompany.GetBusinessObject(BoObjectTypes.oIncomingPayments)
                    Header = "ORCT"
                ElseIf ObjType = "46" Then
                    objPayment = objaddon.objcompany.GetBusinessObject(BoObjectTypes.oVendorPayments)
                    Header = "OVPM"
                End If
                If objPayment.GetByKey(PaymentEntry) Then
                    If objPayment.Cancel() <> 0 Then
                        objaddon.objapplication.SetStatusBarMessage("Payment Entry: " & objaddon.objcompany.GetLastErrorDescription & "-" & objaddon.objcompany.GetLastErrorCode, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                        Return False
                    Else
                        objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        strSQL = "Update " & Header & " Set ""U_RecRem""='Auto Cancelled from JE-'|| " & JE & " Where ""DocEntry""='" & PaymentEntry & "' "
                        objRs.DoQuery(strSQL)
                        objaddon.objapplication.SetStatusBarMessage("Payment Entry Cancelled Successfully...", SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                        Return True
                    End If
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objPayment)
                End If

            Catch ex As Exception

            End Try
        End Function

    End Class
End Namespace
