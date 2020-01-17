Public Class UnlockPay

    Private cSBOApplication As SAPbouiCOM.Application '//OBJETO DE APLICACION
    Private cSBOCompany As SAPbobsCOM.Company     '//OBJETO DE CONEXION
    Private coForm As SAPbouiCOM.Form           '//FORMA
    Private csFormUID As String
    Private stDocNum As String
    Friend Monto As Double


    '//----- METODO DE CREACION DE LA CLASE
    Public Sub New()
        MyBase.New()
        cSBOApplication = oCatchingEvents.SBOApplication
        cSBOCompany = oCatchingEvents.SBOCompany
        Me.stDocNum = stDocNum
    End Sub

    Public Function Valid(ByVal DocNum As String)

        Dim stQueryH As String
        Dim oRecSetH As SAPbobsCOM.Recordset
        Dim CardCode, Banco, FDeposito, FormaPago, Pago, Banco1, FDeposito1, FormaPago1, Pago1, Banco2, FDeposito2, FormaPago2, Pago2 As String

        oRecSetH = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try
            stQueryH = "Select ""CardCode"",""U_CSM_BANCO"",""U_CSM_FDEPOSITO"",""U_CSM_FORMAPAGO"",""U_CSM_IMPORTEPAGADO"", ""U_CSM_BANCO1"",""U_CSM_FDEPOSITO1"",""U_CSM_FORMAPAGO1"",""U_CSM_IMPORTEPAGADO1"", ""U_CSM_BANCO2"",""U_CSM_FDEPOSITO2"",""U_CSM_FORMAPAGO2"",""U_CSM_IMPORTEPAGADO2"" from OINV where (""DocNum""=" & DocNum & ") and (""U_CSM_FORMAPAGO""<>'SS' or ""U_CSM_FORMAPAGO1""<>'SS' or ""U_CSM_FORMAPAGO2""<>'SS')"
            oRecSetH.DoQuery(stQueryH)

            If oRecSetH.RecordCount > 0 Then

                oRecSetH.MoveFirst()

                CardCode = oRecSetH.Fields.Item("CardCode").Value
                Banco = oRecSetH.Fields.Item("U_CSM_BANCO").Value
                FDeposito = oRecSetH.Fields.Item("U_CSM_FDEPOSITO").Value
                FormaPago = oRecSetH.Fields.Item("U_CSM_FORMAPAGO").Value
                Pago = oRecSetH.Fields.Item("U_CSM_IMPORTEPAGADO").Value
                Banco1 = oRecSetH.Fields.Item("U_CSM_BANCO1").Value
                FDeposito1 = oRecSetH.Fields.Item("U_CSM_FDEPOSITO1").Value
                FormaPago1 = oRecSetH.Fields.Item("U_CSM_FORMAPAGO1").Value
                Pago1 = oRecSetH.Fields.Item("U_CSM_IMPORTEPAGADO1").Value
                Banco2 = oRecSetH.Fields.Item("U_CSM_BANCO2").Value
                FDeposito2 = oRecSetH.Fields.Item("U_CSM_FDEPOSITO2").Value
                FormaPago2 = oRecSetH.Fields.Item("U_CSM_FORMAPAGO2").Value
                Pago2 = oRecSetH.Fields.Item("U_CSM_IMPORTEPAGADO2").Value

                Payment(CardCode, FDeposito, Pago, FDeposito1, Pago1, FDeposito2, Pago2)
                ORINA(DocNum)

            Else

                cSBOApplication.MessageBox("¡¡¡Esta Factura no es valida para cancelar sus pagos!!!")

            End If

        Catch ex As Exception

            cSBOApplication.MessageBox("Error al validar Factura: " & ex.Message)

        End Try

    End Function

    Public Function Payment(ByVal CardCode As String, ByVal FDeposito As String, ByVal Pago As String, ByVal FDeposito1 As String, ByVal Pago1 As String, ByVal FDeposito2 As String, ByVal Pago2 As String)

        Dim stQueryH As String
        Dim oRecSetH As SAPbobsCOM.Recordset
        Dim DocEntry As String
        Dim vPay As SAPbobsCOM.Payments

        oRecSetH = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try

            stQueryH = "Select ""DocEntry"" from ORCT where ""CardCode""='" & CardCode & "' and (""DocDate""='" & FDeposito & "' or ""DocDate""='" & FDeposito1 & "' or ""DocDate""='" & FDeposito2 & "') and (""DocTotal""=" & Pago & " or ""DocTotal""=" & Pago1 & " or ""DocTotal""=" & Pago2 & ")"
            oRecSetH.DoQuery(stQueryH)

            If oRecSetH.RecordCount > 0 Then

                oRecSetH.MoveFirst()

                For i = 0 To oRecSetH.RecordCount - 1

                    vPay = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)

                    DocEntry = oRecSetH.Fields.Item("DocEntry").Value
                    vPay.GetByKey(DocEntry)
                    vPay.Cancel()

                    oRecSetH.MoveNext()

                Next

            End If

        Catch ex As Exception

            cSBOApplication.MessageBox("Error al Cancelar el Pago: " & ex.Message)

        End Try

    End Function

    Public Function ORINA(ByVal DocNum As String)

        Dim stQueryH, stQueryH2 As String
        Dim oRecSetH, oRecSetH2 As SAPbobsCOM.Recordset
        Dim oORINA As SAPbobsCOM.Documents
        Dim DocEntry, DocCur As String
        Dim llError As Long
        Dim lsError As String

        oRecSetH = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecSetH2 = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try

            stQueryH = "Select ""DocEntry"",""CardCode"",""SlpCode"",""Project"",""DocTotal"",""DocCur"" from OINV where ""DocNum""=" & DocNum
            oRecSetH.DoQuery(stQueryH)

            If oRecSetH.RecordCount > 0 Then

                oORINA = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
                oRecSetH.MoveFirst()

                DocEntry = oRecSetH.Fields.Item("DocEntry").Value
                DocCur = oRecSetH2.Fields.Item("DocCur").Value

                oORINA.Series = "NC"
                oORINA.CardCode = oRecSetH.Fields.Item("CardCode").Value
                oORINA.DocDate = Now.Date
                oORINA.DocDueDate = Now.Date
                oORINA.DocTotal = oRecSetH.Fields.Item("DocTotal").Value
                oORINA.SalesPersonCode = oRecSetH.Fields.Item("SlpCode").Value
                oORINA.DocumentsOwner = 38
                oORINA.Indicator = oRecSetH.Fields.Item("SlpCode").Value
                oORINA.UserFields.Fields.Item("U_WhsCodeC").Value = oRecSetH.Fields.Item("Project").Value

                stQueryH2 = "Select ""ObjType"",""LineNum"",""ItemCode"",""Price"",""Quantity"",""TaxCode"",""WhsCode"",""Project"",""DiscPrcnt"" from INV1 where ""DocEntry""=" & DocEntry
                oRecSetH2.DoQuery(stQueryH2)

                If oRecSetH2.RecordCount > 0 Then

                    oRecSetH2.MoveFirst()

                    For l = 0 To oRecSetH2.RecordCount - 1

                        oORINA.Lines.ItemCode = oRecSetH2.Fields.Item("ItemCode").Value
                        oORINA.Lines.BaseType = oRecSetH2.Fields.Item("ObjType").Value
                        oORINA.Lines.BaseLine = oRecSetH2.Fields.Item("LineNum").Value
                        oORINA.Lines.BaseEntry = DocEntry
                        oORINA.Lines.Price = oRecSetH2.Fields.Item("Price").Value
                        oORINA.Lines.Quantity = oRecSetH2.Fields.Item("Quantity").Value
                        oORINA.Lines.TaxCode = oRecSetH2.Fields.Item("TaxCode").Value
                        oORINA.Lines.WarehouseCode = oRecSetH2.Fields.Item("WhsCode").Value
                        oORINA.Lines.ProjectCode = oRecSetH2.Fields.Item("Project").Value
                        oORINA.Lines.DiscountPercent = oRecSetH2.Fields.Item("DiscPrcnt").Value
                        oORINA.Lines.Currency = DocCur

                        oORINA.Lines.Add()

                        oRecSetH2.MoveNext()

                    Next

                End If

                If oORINA.Add() <> 0 Then

                    cSBOCompany.GetLastError(llError, lsError)
                    Err.Raise(-1, 1, lsError)

                End If

            End If

        Catch ex As Exception

            cSBOApplication.MessageBox("Error al Crear Nota de Crédito: " & ex.Message)

        End Try

    End Function

End Class
