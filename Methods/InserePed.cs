using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

namespace IntegratorSales.Methods
{
    class InserePed
    {

        public static void GetMatrixPed()
        {
            try
            {
                DataTable dt = new DataTable();
                dt = MYSQL.MatrizPedido();
                foreach (DataRow row in dt.Rows)
                {
                    foreach (DataColumn col in dt.Columns)
                        InserePedido(Convert.ToInt32(row[col]));
                }
            }
            catch (System.Exception e)
            {
                Console.WriteLine(e.Message.ToString());
                //return false;
            }
            //return true;
        }


        public static bool InsereAdiantamento(int docEntry)
        {
            Documents oDoc = null;
            oDoc = (Documents)Utils.vCompany.GetBusinessObject(BoObjectTypes.oOrders);
            oDoc.GetByKey(docEntry);
            Utils.Log(oDoc.GetAsXML());
            try
            {
                Documents iDoc = null;
                iDoc = (Documents)Utils.vCompany.GetBusinessObject(BoObjectTypes.oDownPayments);


                iDoc.CardCode = oDoc.CardCode;
                iDoc.DocDate = oDoc.DocDate;
                iDoc.DocDueDate = oDoc.DocDueDate;
                iDoc.TaxDate = oDoc.TaxDate;
                iDoc.NumAtCard = oDoc.NumAtCard;
                iDoc.PaymentGroupCode = oDoc.PaymentGroupCode;
                iDoc.PaymentMethod = oDoc.PaymentMethod;
                iDoc.Comments = oDoc.Comments;
                iDoc.TaxExtension.Carrier = oDoc.TaxExtension.Carrier;
                iDoc.TaxExtension.Incoterms = oDoc.TaxExtension.Incoterms;
                iDoc.BPL_IDAssignedToInvoice = oDoc.BPL_IDAssignedToInvoice;
                iDoc.SalesPersonCode = oDoc.SalesPersonCode;
                iDoc.ContactPersonCode = oDoc.ContactPersonCode;
                iDoc.UserFields.Fields.Item("U_IdPortal").Value = oDoc.UserFields.Fields.Item("U_IdPortal").Value;
                iDoc.DownPaymentType = SAPbobsCOM.DownPaymentTypeEnum.dptInvoice;
                for (int i = 0; i <= oDoc.Lines.Count - 1; i++)
                {
                    oDoc.Lines.SetCurrentLine(i);

                    iDoc.Lines.BaseType = 17;
                    iDoc.Lines.BaseEntry = docEntry;
                    iDoc.Lines.BaseLine = oDoc.Lines.LineNum;
                    iDoc.Lines.Add();
                }

                string docEntryInserido = Utils.ReturnValue($"SELECT T0.\"DocEntry\" FROM ODPI T0 WHERE T0.\"U_IdPortal\" = '{oDoc.UserFields.Fields.Item("U_IdPortal").Value}'", "DocEntry");
                Utils.Log(iDoc.GetAsXML());

                // seta os parametros do lines
                if (string.IsNullOrEmpty(docEntryInserido))

                    // tenta adicionar o pedido no sap

                    if (iDoc.Add() != 0)
                    {
                        Utils.Log("****************************************************");
                        Utils.Log("Erro Inserindo Adiantamento no SAP:");
                        Utils.Log($"Pedido: {oDoc.UserFields.Fields.Item("U_IdPortal").Value}");
                        Utils.Log($"Erro: {Utils.vCompany.GetLastErrorDescription()}");
                        return false;

                    }

                Utils.Log("****************************************************");
                Utils.Log("Sucesso Inserindo Adiantamento no SAP:");
                Utils.Log($"Pedido: {oDoc.UserFields.Fields.Item("U_IdPortal").Value}");
                return true;
            }
            catch (Exception ex)
            {
                Utils.Log("****************************************************");
                Utils.Log("Erro Inserindo Adiantamento no SAP:");
                Utils.Log($"Pedido: {oDoc.UserFields.Fields.Item("U_IdPortal").Value}");
                Utils.Log($"Erro: {ex.ToString()}");
                return false;
            }


        }
        public static void InserePedido(int absid)
        {
            Documents oDoc = null;
            string Approval1 = "True";
            string Approval2 = "True";
            string Approval3 = "True";
            string Approval4 = "False";
            int entry = 0;
            string cardcode = "";
            string pedidocliente = "";
            bool isDraft = false;
            try
            {
                DataTable dt = new DataTable();
                dt = MYSQL.Pedido(absid);
                foreach (DataRow row in dt.Rows)
                {
                    ///////////////////// Verificação Aprovação de Marketing
                    ////////// Analise Utilização
                    //if (Convert.ToString(row[22]) == "Bonificação Marketing")
                    //{
                    //  Approval2 = "True";
                    //}
                    //////////// Analise itens do Portal no SAP
                    //DataTable itmkt = new DataTable();
                    //itmkt = MYSQL.ItensMkt(absid);
                    //foreach (DataRow row1 in itmkt.Rows)
                    //{
                    //  foreach (DataColumn col in itmkt.Columns)
                    //    if (SAP.GetStatusItMktv2(Convert.ToString(row1[col])) == "Y") //Alterado 02/12/2018
                    //    {
                    //      Approval2 = "True";
                    //    }

                    //}

                    /////////////////////// Verificação Aprovação Financeira
                    //////////// Analise pelo SQL
                    //Approval4 = SAP.GetStatusItFin(Convert.ToString(row[2]), Convert.ToString(row[23]), Convert.ToString(row[24]));

                    /////////////////////// Setando no SAP as aprovações
                    string app1 = (SAP.GetValue("SELECT T0.\"WtmCode\" FROM OWTM T0 WHERE T0.\"Name\"  = 'Bloqueio Financeiro2'", "WtmCode"));
                    string app2 = (SAP.GetValue("SELECT T0.\"WtmCode\" FROM OWTM T0 WHERE T0.\"Name\"  = 'Obs. do Pedido2'", "WtmCode"));

                    ApprovalRequest(app1, "False");
                    ApprovalRequest(app2, "False");

                    string check1 = SAP.GetValue(string.Format(Utils.getQuery("qCheck1"), Convert.ToString(row["CardCode"])), "CHECK");
                    string check2 = SAP.GetValue(string.Format(Utils.getQuery("qCheck2"), Convert.ToString(row["Remarks"])), "CHECK");
                    if (check1 == "TRUE")
                    {
                        ApprovalRequest(app1, "True");
                        isDraft = true;
                    }
                    if (check2 == "TRUE")
                    {
                        ApprovalRequest(app2, "True");
                        isDraft = true;
                    }

                    //ApprovalRequest("10", Approval1);
                    //ApprovalRequest("11", Approval2);
                    //ApprovalRequest("12", Approval3);
                    //ApprovalRequest("13", Approval4);

                    entry = Convert.ToInt32(row["DocEntry"]);
                    oDoc = (Documents)Utils.vCompany.GetBusinessObject(BoObjectTypes.oOrders);

                    //; ; id,,Status,,,,,,,,,,,NatOper,Freight,DiscountTotal,DocTotal,,itensesboco,NSAP,autor,duplicado
                    // seta os parametros do header
                    oDoc.CardCode = Convert.ToString(row["CardCode"]);
                    cardcode = Convert.ToString(row["CardCode"]);
                    string date = Convert.ToString(row["DocDate"]);
                    oDoc.DocDate = DateTime.ParseExact(row["DocDate"].ToString(), @"yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture);
                    oDoc.DocDueDate = DateTime.ParseExact(row["ShipDate"].ToString(), @"yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture);
                    oDoc.TaxDate = DateTime.ParseExact(row["DocDate"].ToString(), @"yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture);
                    oDoc.NumAtCard = entry.ToString();
                    oDoc.PaymentGroupCode = Convert.ToInt32(row["GroupNum"]);
                    oDoc.PaymentMethod = Convert.ToString(row["PymCode"]);
                    oDoc.Comments = Convert.ToString(row["Remarks"]);
                    oDoc.TaxExtension.Carrier = Convert.ToString(row["Shipper"]);
                    oDoc.TaxExtension.Incoterms = Convert.ToString(row["Incoterms"]);
                    //oDoc.BPL_IDAssignedToInvoice = Convert.ToInt32(row[21]); //Alterado 02/12/2018
                    oDoc.BPL_IDAssignedToInvoice = Convert.ToInt32(row["BPLId"]);
                    oDoc.SalesPersonCode = Convert.ToInt32(row["SlpPerson"]);
                    try
                    {
                        oDoc.ContactPersonCode = Convert.ToInt32(row["Contact"]);
                    }
                    catch { }

                    //oDoc.DiscountPercent = Convert.ToInt32(row["Contact"]);

                    //oDoc.UserFields.Fields.Item("U_BUW_Condsugerida").Value = Convert.ToString(row[10]);
                    //oDoc.UserFields.Fields.Item("U_TipoPedido").Value = SAP.GetTpPedido(Convert.ToString(row[11]));
                    //oDoc.UserFields.Fields.Item("U_autor").Value = Convert.ToString(row[12]);
                    oDoc.UserFields.Fields.Item("U_IdPortal").Value = Convert.ToString(row["id"]);


                    DataTable it = new DataTable();
                    it = MYSQL.ItensPedido(absid);
                    foreach (DataRow itrow in it.Rows)
                    {
                        oDoc.Lines.ItemCode = Convert.ToString(itrow["ItemCode"]);
                        oDoc.Lines.Quantity = Convert.ToDouble(itrow["Quantity"]);
                        oDoc.Lines.UnitPrice = Convert.ToDouble(itrow["Price"]);
                        oDoc.Lines.Usage = SAP.GetdfUsage(Convert.ToString(itrow["Usages"]));
                        oDoc.Lines.DiscountPercent = Convert.ToDouble(itrow["Discount"]);
                        oDoc.Lines.ShipDate = DateTime.ParseExact(itrow["ShipDate"].ToString(), @"yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture);
                        try { oDoc.Lines.CESTCode = int.Parse(SAP.GetValue($"SELECT T0.\"CESTCode\" FROM OITM T0 WHERE T0.\"ItemCode\" = '{oDoc.Lines.ItemCode}'", "CESTCode")); } catch { }
                        try { oDoc.Lines.WarehouseCode = (SAP.GetValue($"SELECT T0.\"DflWhs\" FROM OBPL T0 WHERE T0.\"BPLId\" = {oDoc.BPL_IDAssignedToInvoice}", "DflWhs")); } catch { }
                        //oDoc.Lines.UserFields.Fields.Item("U_BUW_vltabela").Value = Convert.ToString(itrow[3]).Replace('.', ',');
                        //oDoc.Lines.UserFields.Fields.Item("U_BUW_Queima").Value = Convert.ToString(itrow[6]);


                        //If para deposito queima
                        /*if (Convert.ToString(itrow[6]) == "Sim")
                        {
                            if (Convert.ToInt32(row[25]) == 5) //Encomex
                            {
                               oDoc.Lines.WarehouseCode = "MSP02-02";
                            }
                            if (Convert.ToInt32(row[25]) == 1) //SP
                                       { 
                               oDoc.Lines.WarehouseCode = "MSP01-02";
                            }
                        }*/


                        oDoc.Lines.Add();
                    }
                    string docEntryInserido = "";
                    if (isDraft)
                        docEntryInserido = Utils.ReturnValue($"SELECT T0.\"DocEntry\" FROM ODRF T0 WHERE T0.\"U_IdPortal\" = '{Convert.ToString(row["id"])}'", "DocEntry");
                    else
                        docEntryInserido = Utils.ReturnValue($"SELECT T0.\"DocEntry\" FROM ORDR T0 WHERE T0.\"U_IdPortal\" = '{Convert.ToString(row["id"])}'", "DocEntry");

                    // seta os parametros do lines
                    if (string.IsNullOrEmpty(docEntryInserido))

                        // tenta adicionar o pedido no sap
                        if (oDoc.Add() != 0)
                        {
                            Utils.Log("****************************************************");
                            Utils.Log("Erro Inserindo Pedido no SAP:");
                            Utils.Log($"Pedido: {entry}");
                            Utils.Log($"Erro: {Utils.vCompany.GetLastErrorDescription()}");
                            MYSQL.ExecInstruction($"Update Pedido  set Status  = 'Erro', LogInt = '{Utils.vCompany.GetLastErrorDescription()}' where id = {entry}");
                            //throw new Exception(Utils.vCompany.GetLastErrorDescription());
                            return;
                        }

                    Utils.Log("****************************************************");
                    Utils.Log("Sucesso Inserindo Pedido no SAP:");
                    Utils.Log($"Pedido: {entry}");

                    if (isDraft)
                        docEntryInserido = Utils.ReturnValue($"SELECT T0.\"DocEntry\" FROM ODRF T0 WHERE T0.\"U_IdPortal\" = '{Convert.ToString(row["id"])}'", "DocEntry");
                    else
                        docEntryInserido = Utils.ReturnValue($"SELECT T0.\"DocEntry\" FROM ORDR T0 WHERE T0.\"U_IdPortal\" = '{Convert.ToString(row["id"])}'", "DocEntry");

                    if (!isDraft)
                    {
                        string Antecipado = $"SELECT T0.\"U_Antecipado\" FROM OCRD T0 WHERE T0.\"CardCode\" = '{oDoc.CardCode}'";
                        Antecipado = Utils.ReturnValue(Antecipado, "U_Antecipado");
                        Utils.Log($"Adiantamento:{Antecipado}");
                        if (Antecipado == "S")
                        {
                            if (InsereAdiantamento(int.Parse(docEntryInserido)))
                                MYSQL.UpdateStatusPed(entry, docEntryInserido, "Aprovado");
                        }
                        else
                        {
                            MYSQL.UpdateStatusPed(entry, docEntryInserido, "Aprovado");
                        }
                    }
                    else
                        MYSQL.UpdateStatusPed(entry, docEntryInserido, "Em Aprovação");

                    ApprovalRequest(app1, "False");
                    ApprovalRequest(app2, "False");
                    //ApprovalRequest("10", "True");
                    //ApprovalRequest("11", "True");
                    //ApprovalRequest("12", "True");
                    //ApprovalRequest("13", "True");
                }
            }
            catch (System.Exception e)
            {
                Utils.Log("****************************************************");
                Utils.Log("Erro Inserindo Pedido no SAP:");
                Utils.Log($"Pedido: {entry}");
                Utils.Log($"Erro: {e.Message}");
                Console.WriteLine(e.Message.ToString());
                MYSQL.ExecInstruction($"Update Pedido  set Status  = 'Erro', LogInt = '{e.ToString()}' where id = {entry}");
                //return true; //fake true pra não parar
            }
            finally
            {

                Marshal.ReleaseComObject(oDoc);
                oDoc = null;

            }
            //return true;
        }
        public static void insereAprov(int entry)
        {

            Documents oDoc = null;
            oDoc = (Documents)Utils.vCompany.GetBusinessObject(BoObjectTypes.oDrafts);

            oDoc.GetByKey(entry);

            if (oDoc.SaveDraftToDocument() != 0)
            {
                Utils.Log("****************************************************");
                Utils.Log("Erro Inserindo Pedido no SAP:");
                Utils.Log($"Pedido: {entry}");
                Utils.Log($"Erro: {Utils.vCompany.GetLastErrorDescription()}");
                return;

            }

            Utils.Log("****************************************************");
            Utils.Log("Sucesso Inserindo Pedido no SAP:");
            Utils.Log($"Pedido: {entry}");

            string docEntryInserido = Utils.ReturnValue($"SELECT T0.\"DocEntry\" FROM ORDR T0 WHERE T0.\"U_IdPortal\" = '{oDoc.UserFields.Fields.Item("U_IdPortal").Value}'", "DocEntry");

            string Antecipado = $"SELECT T0.\"U_Antecipado\" FROM OCRD T0 WHERE T0.\"CardCode\" = '{oDoc.CardCode}'";
            Antecipado = Utils.ReturnValue(Antecipado, "U_Antecipado");
            Utils.Log($"Adiantamento:{Antecipado}");
            if (Antecipado == "S")
            {
                if (InsereAdiantamento(int.Parse(docEntryInserido)))
                    MYSQL.UpdateStatusPed(entry, docEntryInserido, "Aprovado");
            }

        }
        public static void insereAprov()
        {
            Utils.Log("Iniciando rotina de Insercao Automatica de Pedido ja aprovados.");


            try
            {


                string query = Utils.getQuery("qAprov");

                Recordset oRs2 = (Recordset)Utils.vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                oRs2.DoQuery(query);


                while (!oRs2.EoF)
                {
                    string Entry = oRs2.Fields.Item("DocEntry").Value.ToString();

                    insereAprov(int.Parse(Entry));

                    oRs2.MoveNext();
                }


            }
            catch (Exception e)
            {
                Utils.Log(e.ToString());
            }

        }
        public static string retiraAcentos(string texto)
        {
            string comAcentos = "ÄÅÁÂÀÃäáâàãÉÊËÈéêëèÍÎÏÌíîïìÖÓÔÒÕöóôòõÜÚÛüúûùÇç";
            string semAcentos = "AAAAAAaaaaaEEEEeeeeIIIIiiiiOOOOOoooooUUUuuuuCc";
            for (int i = 0; i < comAcentos.Length; i++)
            {
                texto = texto.Replace(comAcentos[i].ToString(), semAcentos[i].ToString());
            }
            return texto;
        }

        public static bool ApprovalRequest(string code, string TrueorFalse)
        {
            SAPbobsCOM.CompanyService oCmpSrv = Utils.vCompany.GetCompanyService();
            try
            {
                ApprovalTemplatesService oApprovalTemplateService = (ApprovalTemplatesService)oCmpSrv.GetBusinessService(ServiceTypes.ApprovalTemplatesService);
                ApprovalTemplateParams oApprovalTemplateParams = (ApprovalTemplateParams)oApprovalTemplateService.GetDataInterface(SAPbobsCOM.ApprovalTemplatesServiceDataInterfaces.atsdiApprovalTemplateParams);
                oApprovalTemplateParams.Code = Convert.ToInt32(code);
                ApprovalTemplate oApprovalTemplate = oApprovalTemplateService.GetApprovalTemplate(oApprovalTemplateParams);
                oApprovalTemplate.IsActive = BoYesNoEnum.tYES;
                if (TrueorFalse == "False")
                {
                    oApprovalTemplate.IsActive = BoYesNoEnum.tNO;
                }
                oApprovalTemplateService.UpdateApprovalTemplate(oApprovalTemplate);
            }
            catch (System.Exception e)
            {
                Console.WriteLine(e.Message.ToString());
                return false;
            }
            finally
            {

                Marshal.ReleaseComObject(oCmpSrv);
                oCmpSrv = null;

            }
            return true;
        }


    }
}
