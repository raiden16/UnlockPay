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
                ORINA()

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

    Public Function ORINA()

    End Function

End Class
