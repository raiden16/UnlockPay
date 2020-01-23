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
        Dim CardCode, Banco, FormaPago, Banco1, FormaPago1, Banco2, FormaPago2, Banco3, FormaPago3, Banco4, FormaPago4, FDepo, FDepo1, FDepo2, FDepo3, FDepo4, DocEntry As String
        Dim FDeposito, FDeposito1, FDeposito2, FDeposito3, FDeposito4 As Date
        Dim Pago, Pago1, Pago2, Pago3, Pago4 As Double
        Dim ContP, ContORIN, ContOBNK As Integer

        oRecSetH = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try
            stQueryH = "Select ""DocEntry"",""CardCode"",
                       ""U_CSM_BANCO"",""U_CSM_FDEPOSITO"",""U_CSM_FORMAPAGO"",""U_CSM_IMPORTEPAGADO"",
                       ""U_CSM_BANCO1"",""U_CSM_FDEPOSITO1"",""U_CSM_FORMAPAGO1"",""U_CSM_IMPORTEPAGADO1"",
                       ""U_CSM_BANCO2"",""U_CSM_FDEPOSITO2"",""U_CSM_FORMAPAGO2"",""U_CSM_IMPORTEPAGADO2"",
                       ""U_CSM_BANCO4"",""U_CSM_FVOUCHER1"",'TPV' as ""TPV1"",""U_CSM_IMPORTEVOUCHER1"",
                       ""U_CSM_BANCO5"",""U_CSM_FVOUCHER2"",'TPV' as ""TPV2"",""U_CSM_IMPORTEVOUCHER2""
                       from OINV where (""DocNum""=" & DocNum & ")"
            oRecSetH.DoQuery(stQueryH)

            DocEntry = oRecSetH.Fields.Item("DocEntry").Value
            ContP = Payment(DocEntry)

            ContORIN = ORINA(DocNum)

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
                Banco3 = oRecSetH.Fields.Item("U_CSM_BANCO4").Value
                FDeposito3 = oRecSetH.Fields.Item("U_CSM_FVOUCHER1").Value
                FormaPago3 = oRecSetH.Fields.Item("TPV1").Value
                Pago3 = oRecSetH.Fields.Item("U_CSM_IMPORTEVOUCHER1").Value
                Banco4 = oRecSetH.Fields.Item("U_CSM_BANCO5").Value
                FDeposito4 = oRecSetH.Fields.Item("U_CSM_FVOUCHER2").Value
                FormaPago4 = oRecSetH.Fields.Item("TPV2").Value
                Pago4 = oRecSetH.Fields.Item("U_CSM_IMPORTEVOUCHER2").Value
                FDepo = Year(FDeposito).ToString + "-" + Month(FDeposito).ToString + "-" + Day(FDeposito).ToString
                FDepo1 = Year(FDeposito1).ToString + "-" + Month(FDeposito1).ToString + "-" + Day(FDeposito1).ToString
                FDepo2 = Year(FDeposito2).ToString + "-" + Month(FDeposito2).ToString + "-" + Day(FDeposito2).ToString
                FDepo3 = Year(FDeposito3).ToString + "-" + Month(FDeposito3).ToString + "-" + Day(FDeposito3).ToString
                FDepo4 = Year(FDeposito4).ToString + "-" + Month(FDeposito4).ToString + "-" + Day(FDeposito4).ToString

                ContOBNK = OBNK(CardCode, Banco, FDepo, FormaPago, Pago, Banco1, FDepo1, FormaPago1, Pago1, Banco2, FDepo2, FormaPago2, Pago2, Banco3, FDepo3, FormaPago3, Pago3, Banco4, FDepo4, FormaPago4, Pago4)

            End If

            If ContP > 0 Then

                cSBOApplication.MessageBox("Cancelación de Pagos exitosa")

            End If

            If ContORIN > 0 Then

                cSBOApplication.MessageBox("Creación de NC exitosa")

            End If

            If ContOBNK > 0 Then

                cSBOApplication.MessageBox("Liberación del Extracto Bancario exitosa")

            End If

        Catch ex As Exception

            cSBOApplication.MessageBox("Error al validar Factura: " & ex.Message)

        End Try

    End Function


    Public Function Payment(ByVal DocEntry As String)

        Dim stQueryH As String
        Dim oRecSetH As SAPbobsCOM.Recordset
        Dim vPay As SAPbobsCOM.Payments
        Dim DocEntryP As String
        Dim contador As Integer

        oRecSetH = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        contador = 0

        Try

            stQueryH = "Select T1.""DocEntry"" from RCT2 T0 inner join ORCT T1 on T1.""DocNum""=T0.""DocNum"" where T0.""DocEntry""=" & DocEntry & " and T0.""InvType""=13"
            oRecSetH.DoQuery(stQueryH)

            If oRecSetH.RecordCount > 0 Then

                oRecSetH.MoveFirst()

                For i = 0 To oRecSetH.RecordCount - 1

                    vPay = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)

                    DocEntryP = oRecSetH.Fields.Item("DocEntry").Value
                    vPay.GetByKey(DocEntryP)

                    If vPay.Cancel() = 0 Then

                        contador = contador + 1

                    End If

                    oRecSetH.MoveNext()

                Next

            End If

            Return contador

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
        Dim contador As Integer

        oRecSetH = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecSetH2 = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        contador = 0

        Try

            stQueryH = "Select ""DocEntry"",""CardCode"",""SlpCode"",""Project"",""DocTotal"",""DocCur"" from OINV where ""DocNum""=" & DocNum
            oRecSetH.DoQuery(stQueryH)

            If oRecSetH.RecordCount > 0 Then

                oORINA = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
                oRecSetH.MoveFirst()

                DocEntry = oRecSetH.Fields.Item("DocEntry").Value
                DocCur = oRecSetH.Fields.Item("DocCur").Value

                oORINA.Series = 6
                oORINA.CardCode = oRecSetH.Fields.Item("CardCode").Value
                oORINA.DocDate = Year(Now.Date).ToString + "-" + Month(Now.Date).ToString + "-" + Day(Now.Date).ToString
                oORINA.DocDueDate = Year(Now.Date).ToString + "-" + Month(Now.Date).ToString + "-" + Day(Now.Date).ToString
                oORINA.DocTotal = oRecSetH.Fields.Item("DocTotal").Value
                oORINA.SalesPersonCode = oRecSetH.Fields.Item("SlpCode").Value
                oORINA.DocumentsOwner = 38
                oORINA.Indicator = oRecSetH.Fields.Item("SlpCode").Value
                oORINA.UserFields.Fields.Item("U_WhsCodeC").Value = oRecSetH.Fields.Item("Project").Value
                oORINA.UserFields.Fields.Item("U_B1SYS_MainUsage").Value = "G02"
                oORINA.EDocGenerationType = 0

                stQueryH2 = "Select T0.""ObjType"",T0.""LineNum"",T0.""ItemCode"",T0.""Price"",T0.""Quantity"",T0.""TaxCode"",T0.""WhsCode"",T0.""Project"",T0.""DiscPrcnt"",T1.""BatchNum"" from INV1 T0 Left Outer Join IBT1 T1 on T1.""BaseType""=T0.""ObjType"" and T1.""BaseEntry""=T0.""DocEntry"" And T1.""BaseLinNum""=T0.""LineNum"" And T1.""ItemCode""=T0.""ItemCode"" where ""DocEntry""=" & DocEntry
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
                        oORINA.Lines.BatchNumbers.BatchNumber = oRecSetH2.Fields.Item("BatchNum").Value
                        oORINA.Lines.BatchNumbers.Quantity = oRecSetH2.Fields.Item("Quantity").Value
                        oORINA.Lines.BatchNumbers.Notes = oRecSetH2.Fields.Item("BatchNum").Value

                        oORINA.Lines.Add()

                        oRecSetH2.MoveNext()

                    Next

                End If

                If oORINA.Add() <> 0 Then

                    cSBOCompany.GetLastError(llError, lsError)
                    Err.Raise(-1, 1, lsError)

                Else

                    contador = contador + 1

                End If

            End If

            Return contador

        Catch ex As Exception

            cSBOApplication.MessageBox("Error al Crear Nota de Crédito: " & ex.Message)

        End Try

    End Function


    Public Function OBNK(ByVal CardCode As String, ByVal Banco As String, ByVal FDeposito As String, ByVal FormaPago As String, ByVal Pago As Double, ByVal Banco1 As String, ByVal FDeposito1 As String, ByVal FormaPago1 As String, ByVal Pago1 As Double, ByVal Banco2 As String, ByVal FDeposito2 As String, ByVal FormaPago2 As String, ByVal Pago2 As Double, ByVal Banco3 As String, ByVal FDeposito3 As String, ByVal FormaPago3 As String, ByVal Pago3 As Double, ByVal Banco4 As String, ByVal FDeposito4 As String, ByVal FormaPago4 As String, ByVal Pago4 As Double)

        Dim stQueryH As String
        Dim oRecSetH As SAPbobsCOM.Recordset
        Dim oCuenta As SAPbobsCOM.BankPages
        Dim Account, Sequence, Ref, RDate As String
        Dim DebAmount, CredAmnt As Double
        Dim oDueDate As Date
        Dim llError As Long
        Dim lsError As String
        Dim contador As Integer

        oRecSetH = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        contador = 0

        Try

            stQueryH = "Select ""AcctCode"",""Sequence"",""Ref"",""DueDate"",""DebAmount"",""CredAmnt"" From OBNK Where (""CardCode"" ='" & CardCode & "') and (""AcctCode""='" & Banco & "' or ""AcctCode""='" & Banco1 & "' or ""AcctCode""='" & Banco2 & "' or ""AcctCode""='" & Banco3 & "' or ""AcctCode""='" & Banco4 & "') and (""DueDate""='" & FDeposito & "' or ""DueDate""='" & FDeposito1 & "' or ""DueDate""='" & FDeposito2 & "' or ""DueDate""='" & FDeposito3 & "' or ""DueDate""='" & FDeposito4 & "') and (""Ref""='" & FormaPago & "' or ""Ref""='" & FormaPago1 & "' or ""Ref""='" & FormaPago2 & "' or ""Ref""='" & FormaPago3 & "' or ""Ref""='" & FormaPago4 & "') and (""CredAmnt""=" & Pago & " or ""CredAmnt""=" & Pago1 & " or ""CredAmnt""=" & Pago2 & " or ""CredAmnt""=" & Pago3 & " or ""CredAmnt""=" & Pago4 & ")"
            oRecSetH.DoQuery(stQueryH)

            If oRecSetH.RecordCount > 0 Then

                oCuenta = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBankPages)
                oRecSetH.MoveFirst()

                For i = 0 To oRecSetH.RecordCount - 1

                    Account = oRecSetH.Fields.Item("AcctCode").Value
                    Sequence = oRecSetH.Fields.Item("Sequence").Value
                    Ref = oRecSetH.Fields.Item("Ref").Value
                    oDueDate = oRecSetH.Fields.Item("DueDate").Value
                    DebAmount = oRecSetH.Fields.Item("DebAmount").Value
                    CredAmnt = oRecSetH.Fields.Item("CredAmnt").Value
                    RDate = Year(oDueDate).ToString + "-" + Month(oDueDate).ToString + "-" + Day(oDueDate).ToString

                    oCuenta.GetByKey(Account, Sequence)
                    oCuenta.CardCode = ""
                    oCuenta.CardName = ""
                    oCuenta.ExternalCode = ""

                    If oCuenta.Update() = 0 Then

                        InsertLog(Account, Ref, RDate, DebAmount, CredAmnt)
                        contador = contador + 1

                    Else

                        cSBOCompany.GetLastError(llError, lsError)
                        Err.Raise(-1, 1, lsError)

                    End If

                    oRecSetH.MoveNext()

                Next

            End If

            Return contador

        Catch ex As Exception

            cSBOApplication.MessageBox("Error al limpiar el extracto bancario: " & ex.Message)

        End Try

    End Function


    Public Function InsertLog(ByVal Account As String, ByVal Ref As String, ByVal oDueDate As String, ByVal DebAmount As Double, ByVal CredAmnt As Double)

        Dim stQueryH, stQueryH2 As String
        Dim oRecSetH, oRecSetH2 As SAPbobsCOM.Recordset
        Dim Code, CurrentDate, RDate As String

        oRecSetH = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecSetH2 = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try

            CurrentDate = Now.Year.ToString + "-" + Now.Month.ToString + "-" + Now.Day.ToString

            stQueryH = "Select case when length(count(""U_AcctCode"")+1)=1 then concat('00',TO_NVARCHAR(count(""U_AcctCode"")+1)) when length(count(""U_AcctCode"")+1)=2 then concat('0',TO_NVARCHAR(count(""U_AcctCode"")+1)) when length(count(""U_AcctCode"")+1)=3 then TO_NVARCHAR(count(""U_AcctCode"")+1) end as ""Codigo"" from ""@LOG_OBNK"" where ""U_AcctCode""=" & Account & " and ""U_DueDate""='" & oDueDate & "'"
            oRecSetH.DoQuery(stQueryH)

            RDate = Year(oDueDate).ToString + Month(oDueDate).ToString + Day(oDueDate).ToString

            Code = RDate + Account.Substring(7, 2).ToString + oRecSetH.Fields.Item("Codigo").Value

            stQueryH2 = "INSERT INTO ""@LOG_OBNK"" VALUES (" & Code & "," & Code & ",'" & oDueDate & "','" & CurrentDate & "','" & Ref & "',null,null,null,'" & Account & "','UnlockPay'," & DebAmount & "," & CredAmnt & ")"
            oRecSetH2.DoQuery(stQueryH2)

        Catch ex As Exception

            cSBOApplication.MessageBox("Error al actualizar LOG_OBNK : " & ex.Message)

        End Try

    End Function


End Class
