using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.Runtime.ConstrainedExecution;

namespace IntegratorSales.Methods
{
    class InsereBP
    {

        public static void GetMatrixBP()
        {
            try
            {
                DataTable dt = new DataTable();
                dt = MYSQL.MatrizLead();
                foreach (DataRow row in dt.Rows)
                {
                    foreach (DataColumn col in dt.Columns)
                        InserePN(Convert.ToInt32(row[col]));
                }
            }
            catch (System.Exception e)
            {
                Console.WriteLine(e.Message.ToString());
                //return false;
            }
            //return true;
        }

        public static void InserePN(int absid)
        {
            BusinessPartners vBP = null;
            int entry = 0;
            int userid = 0;
            string cardcode = "";
            string Lead = "";
            try
            {
                DataTable dt = new DataTable();
                dt = MYSQL.Lead(absid);
                foreach (DataRow row in dt.Rows)
                {



                    //userid = Convert.ToInt32(row[63]);
                    entry = Convert.ToInt32(row["id"]);
                    vBP = (BusinessPartners)Utils.vCompany.GetBusinessObject(BoObjectTypes.oBusinessPartners);
                    //cardcode = "C" + row[2].ToString();
                    // seta os parametros do header
                    vBP.Series = 76;
                    //vBP.UserFields.Fields.Item("U_PTL_Id").Value = entry.ToString();
                    vBP.GroupCode = 100;
                    //vBP.CardCode = cardcode;
                    vBP.CardType = BoCardTypes.cLid;
                    Lead = Regex.Replace(Util.truncatetexto(retiraAcentos(row["CardName"].ToString()), 100), @"[^0-9a-zA-Z]+", " ").ToUpper();
                    vBP.CardName = Regex.Replace(Util.truncatetexto(retiraAcentos(row["CardName"].ToString()), 100), @"[^0-9a-zA-Z]+", " ").ToUpper();
                    vBP.CardForeignName = Regex.Replace(Util.truncatetexto(retiraAcentos(row["CardFName"].ToString()), 100), @"[^0-9a-zA-Z]+", " ").ToUpper();






                    vBP.EmailAddress = Util.truncatetexto(row["Email"].ToString(), 100);
                    //vBP.Phone2 = Util.truncatetexto(row[5].ToString(), 20);
                    vBP.Phone1 = Util.truncatetexto(row["Phone"].ToString(), 20);
                    //vBP.Notes = "Vendedor: " + SAP.GetVendorName(userid.ToString());

                    // endereço de entrega
                    vBP.Addresses.AddressType = BoAddressType.bo_ShipTo;
                    vBP.Addresses.AddressName = string.Format("Endereço Entrega");
                    vBP.Addresses.ZipCode = Util.truncatetexto(row["Cep"].ToString(), 20);
                    vBP.Addresses.Country = "BR";
                    vBP.Addresses.State = Util.truncatetexto(row["Estado"].ToString(), 2);
                    vBP.Addresses.City = Regex.Replace(Util.truncatetexto(retiraAcentos(row["Cidade"].ToString()), 100), @"[^0-9a-zA-Z]+", " ").ToUpper();
                    vBP.Addresses.County = SAP.GetCounty(row["Cidade"].ToString(), row["Estado"].ToString());
                    vBP.Addresses.Block = Regex.Replace(Util.truncatetexto(retiraAcentos(row["Bairro"].ToString()), 100), @"[^0-9a-zA-Z]+", " ").ToUpper();
                    //vBP.Addresses.TypeOfAddress = Regex.Replace(Util.truncatetexto(retiraAcentos(row[7].ToString()), 100), @"[^0-9a-zA-Z]+", " ").ToUpper();
                    vBP.Addresses.Street = Regex.Replace(Util.truncatetexto(retiraAcentos(row["Endereco"].ToString()), 100), @"[^0-9a-zA-Z]+", " ").ToUpper();
                    vBP.Addresses.StreetNo = Regex.Replace(Util.truncatetexto(retiraAcentos(row["Numero"].ToString()), 100), @"[^0-9a-zA-Z]+", " ").ToUpper();
                    vBP.Addresses.BuildingFloorRoom = Regex.Replace(Util.truncatetexto(retiraAcentos(row["Complemento"].ToString()), 100), @"[^0-9a-zA-Z]+", " ").ToUpper();
                    //vBP.Addresses.UserFields.Fields.Item("U_SKILL_indIEDest").Value = SAP.GetParams("PORTAL_CadIndEst");
                    vBP.Addresses.Add();


                    // endereço de cobrança
                    vBP.Addresses.AddressType = BoAddressType.bo_BillTo;
                    vBP.Addresses.AddressName = string.Format("Endereço Cobrança");
                    vBP.Addresses.ZipCode = Util.truncatetexto(row["Cep"].ToString(), 20);
                    vBP.Addresses.Country = "BR";
                    vBP.Addresses.State = Util.truncatetexto(row["Estado"].ToString(), 2);
                    vBP.Addresses.City = Regex.Replace(Util.truncatetexto(retiraAcentos(row["Cidade"].ToString()), 100), @"[^0-9a-zA-Z]+", " ").ToUpper();
                    vBP.Addresses.County = SAP.GetCounty(row["Cidade"].ToString(), row["Estado"].ToString());
                    vBP.Addresses.Block = Regex.Replace(Util.truncatetexto(retiraAcentos(row["Bairro"].ToString()), 100), @"[^0-9a-zA-Z]+", " ").ToUpper();
                    //vBP.Addresses.TypeOfAddress = Regex.Replace(Util.truncatetexto(retiraAcentos(row[7].ToString()), 100), @"[^0-9a-zA-Z]+", " ").ToUpper();
                    vBP.Addresses.Street = Regex.Replace(Util.truncatetexto(retiraAcentos(row["Endereco"].ToString()), 100), @"[^0-9a-zA-Z]+", " ").ToUpper();
                    vBP.Addresses.StreetNo = Regex.Replace(Util.truncatetexto(retiraAcentos(row["Numero"].ToString()), 100), @"[^0-9a-zA-Z]+", " ").ToUpper();
                    vBP.Addresses.BuildingFloorRoom = Regex.Replace(Util.truncatetexto(retiraAcentos(row["Complemento"].ToString()), 100), @"[^0-9a-zA-Z]+", " ").ToUpper();
                    //vBP.Addresses.UserFields.Fields.Item("U_SKILL_indIEDest").Value = SAP.GetParams("PORTAL_CadIndEst");
                    vBP.Addresses.Add();








                    if (row["TaxIdNum2"].ToString() == "Isento")
                    {
                        vBP.FiscalTaxID.TaxId1 = "Isento";
                    }

                    if (row["TaxIdNum2"].ToString() != "Isento")
                    {
                        vBP.FiscalTaxID.TaxId1 = row["TaxIdNum2"].ToString();
                    }

                    vBP.FiscalTaxID.TaxId0 = row["TaxIdNum"].ToString();

                    // Contato Comercial
                    vBP.ContactEmployees.Name = Util.truncatetexto(row["Contato"].ToString(), 50);
                    //vBP.ContactEmployees.FirstName = Util.truncatetexto(row[23].ToString(), 50);
                    //vBP.ContactEmployees.LastName = Util.truncatetexto(row[24].ToString(), 50);
                    //vBP.ContactEmployees.Position = Util.truncatetexto(row[25].ToString(), 90);
                    //vBP.ContactEmployees.Title = Util.truncatetexto(row[25].ToString(), 10);
                    //vBP.ContactEmployees.Phone2 = Util.truncatetexto(row[27].ToString(), 20);
                    vBP.ContactEmployees.Phone1 = Util.truncatetexto(row["Phone"].ToString(), 20);
                    //vBP.ContactEmployees.Fax = Util.truncatetexto(row[30].ToString(), 20);
                    //Console.WriteLine(row[31].ToString());
                    //if (row[31].ToString() != "0000-00-00")
                    //{
                    //    vBP.ContactEmployees.DateOfBirth = DateTime.ParseExact(row[31].ToString(), @"yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture); // Data Nascimento
                    //}
                    ////vBP.ContactEmployees.UserFields.Fields.Item("U_PTL_Obs").Value = row[32].ToString();
                    //vBP.ContactEmployees.MobilePhone = Util.truncatetexto(row[29].ToString(), 50);
                    vBP.ContactEmployees.Add();

                    //// Contato Financeiro
                    //vBP.ContactEmployees.Name = "Financeiro";
                    //vBP.ContactEmployees.FirstName = Util.truncatetexto(row[33].ToString(), 50);
                    //vBP.ContactEmployees.LastName = Util.truncatetexto(row[34].ToString(), 50);
                    //vBP.ContactEmployees.Position = Util.truncatetexto(row[35].ToString(), 90);
                    //vBP.ContactEmployees.Title = Util.truncatetexto(row[35].ToString(), 10);
                    //vBP.ContactEmployees.Phone2 = Util.truncatetexto(row[37].ToString(), 20);
                    //vBP.ContactEmployees.Phone1 = Util.truncatetexto(row[38].ToString(), 20);
                    //vBP.ContactEmployees.Fax = Util.truncatetexto(row[40].ToString(), 20);
                    //Console.WriteLine(row[31].ToString());
                    //if (row[41].ToString() != "0000-00-00")
                    //{
                    //    vBP.ContactEmployees.DateOfBirth = DateTime.ParseExact(row[41].ToString(), @"yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture); // Data Nascimento
                    //}
                    ////vBP.ContactEmployees.UserFields.Fields.Item("U_PTL_Obs").Value = row[42].ToString();
                    //vBP.ContactEmployees.MobilePhone = Util.truncatetexto(row[39].ToString(), 50);
                    //vBP.ContactEmployees.Add();

                    //// Contato Marketing
                    //vBP.ContactEmployees.Name = "Marketing";
                    //vBP.ContactEmployees.FirstName = Util.truncatetexto(row[43].ToString(), 50);
                    //vBP.ContactEmployees.LastName = Util.truncatetexto(row[44].ToString(), 50);
                    //vBP.ContactEmployees.Position = Util.truncatetexto(row[45].ToString(), 90);
                    //vBP.ContactEmployees.Title = Util.truncatetexto(row[45].ToString(), 10);
                    //vBP.ContactEmployees.Phone2 = Util.truncatetexto(row[47].ToString(), 20);
                    //vBP.ContactEmployees.Phone1 = Util.truncatetexto(row[48].ToString(), 20);
                    //vBP.ContactEmployees.Fax = Util.truncatetexto(row[50].ToString(), 20);
                    //Console.WriteLine(row[31].ToString());
                    //if (row[51].ToString() != "0000-00-00")
                    //{
                    //    vBP.ContactEmployees.DateOfBirth = DateTime.ParseExact(row[51].ToString(), @"yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture); // Data Nascimento
                    //}
                    ////vBP.ContactEmployees.UserFields.Fields.Item("U_PTL_Obs").Value = row[52].ToString();
                    //vBP.ContactEmployees.MobilePhone = Util.truncatetexto(row[49].ToString(), 50);
                    //vBP.ContactEmployees.Add();

                    //// Contato Logistica
                    //vBP.ContactEmployees.Name = "Logistica";
                    //vBP.ContactEmployees.FirstName = Util.truncatetexto(row[53].ToString(), 50);
                    //vBP.ContactEmployees.LastName = Util.truncatetexto(row[54].ToString(), 50);
                    //vBP.ContactEmployees.Position = Util.truncatetexto(row[55].ToString(), 90);
                    //vBP.ContactEmployees.Title = Util.truncatetexto(row[55].ToString(), 10);
                    //vBP.ContactEmployees.Phone2 = Util.truncatetexto(row[57].ToString(), 20);
                    //vBP.ContactEmployees.Phone1 = Util.truncatetexto(row[58].ToString(), 20);
                    //vBP.ContactEmployees.Fax = Util.truncatetexto(row[60].ToString(), 20);
                    //Console.WriteLine(row[31].ToString());
                    //if (row[61].ToString() != "0000-00-00")
                    //{
                    //    vBP.ContactEmployees.DateOfBirth = DateTime.ParseExact(row[61].ToString(), @"yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture); // Data Nascimento
                    //}

                    ////vBP.ContactEmployees.UserFields.Fields.Item("U_PTL_Obs").Value = row[62].ToString();
                    //vBP.ContactEmployees.MobilePhone = Util.truncatetexto(row[59].ToString(), 50);
                    //vBP.ContactEmployees.Add();
                    //Console.WriteLine(row[31].ToString());

                    // tenta adicionar o parceiro no sap
                    if (vBP.Add() != 0)
                    {
                        Utils.Log("****************************************************");
                        Utils.Log("Erro Inserindo Lead no SAP:");
                        Utils.Log($"Lead: {Lead}");
                        Utils.Log($"Erro: {Utils.vCompany.GetLastErrorDescription()}");
                        throw new Exception(Utils.vCompany.GetLastErrorDescription());
                    }
                    else
                    {
                        Utils.Log("****************************************************");
                        Utils.Log("Sucesso Inserindo Lead no SAP:");
                        Utils.Log($"Lead: {Lead}");
                        MYSQL.UpdateStatus(entry, Utils.vCompany.GetNewObjectKey());
                    }


                    //SAPbobsCOM.CompanyService oCmpSrv = SAP.vCompany.GetCompanyService();

                    //ActivitiesService oActSrv = (ActivitiesService)oCmpSrv.GetBusinessService(ServiceTypes.ActivitiesService);
                    //Activity oAct = (Activity)oActSrv.GetDataInterface(ActivitiesServiceDataInterfaces.asActivity);
                    //ActivityParams oParams;

                    //oAct.CardCode = cardcode;
                    //oAct.StartDate = DateTime.Now;
                    //oAct.StartTime = DateTime.Now;
                    //oAct.EndDuedate = DateTime.Now;

                    //TimeSpan time = new TimeSpan(0, 30, 0);
                    //DateTime combined = DateTime.Now.Add(time);

                    //oAct.EndTime = combined;
                    //oAct.Activity = BoActivities.cn_Task;
                    //oAct.Notes = "Efetuar Analise do Cadastro.";
                    //oAct.Details = "Cliente adicionado pelo portal de Vendas, favor realizar a conferencia do cadastro.";

                    ////oAct.UserFields.Fields.Item("U_PTL_Id").Value = entry.ToString();
                    //oAct.UserFields.Item("U_PTL_Id").Value = entry.ToString();
                    //oAct.UserFields.Item("U_PTL_Type").Value = "1";
                    //oAct.Reminder = BoYesNoEnum.tYES;
                    //oAct.ReminderType = BoDurations.du_Minuts;
                    //oAct.ReminderPeriod = 1;


                    //oAct.Subject = Convert.ToInt32(SAP.GetParams("PORTAL_AleSub"));
                    //oAct.HandledBy = Convert.ToInt32(SAP.GetParams("PORTAL_AleDftPer"));
                    //oParams = oActSrv.AddActivity(oAct);
                    //long singleActCode = oParams.ActivityCode;


                }
            }
            catch (System.Exception e)
            {
                Console.WriteLine(e.Message.ToString());
                //return false;
            }
            finally
            {

                Marshal.ReleaseComObject(vBP);
                vBP = null;
            }
            //return true;
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

    }
}
