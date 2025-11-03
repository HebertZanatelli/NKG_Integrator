using System.Xml;
using System.IO;
using System.Runtime.InteropServices;
using System;
using SAPbobsCOM;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IntegratorSales
{
    class SAP
    {



        #region Atributos
        public static SAPbobsCOM.Company vCompany = null;
        private bool disposed = false;
        

        public static string sHost;
        public static string sPass;
        public static string sUser;
        public static string sBanco;

        #endregion

        #region Rotinas de ConnectSAP
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        protected virtual void Dispose(bool disposing)
        {
            if (disposed)
                return;

            if (disposing)
            {
                if (vCompany != null && vCompany.Connected)
                {
                    // disconecta o company object
                    vCompany.Disconnect();
                    vCompany = null;
                }
            }

            // Free any unmanaged objects here.
            disposed = true;
        }
        public static bool ConnectSAP()
        {
            try
            {
                

                XmlDocument doc = new XmlDocument();
                string caminho = AppDomain.CurrentDomain.BaseDirectory + "SAP.xml";

                if (vCompany != null && vCompany.Connected)
                    return true;

                if (!File.Exists(caminho))
                    throw new Exception("Sem dados da configuração do banco de dados");

                doc.Load(caminho);
                XmlNode no = doc.SelectSingleNode("./SAP");
                sHost = no.SelectSingleNode("./host").InnerText;
                sPass = no.SelectSingleNode("./pass").InnerText;
                sUser = no.SelectSingleNode("./user").InnerText;
                sBanco = no.SelectSingleNode("./data").InnerText;
                string sVersion = no.SelectSingleNode("./version").InnerText;
                string sUserSAP = no.SelectSingleNode("./userSAP").InnerText;
                string sPassSAP = no.SelectSingleNode("./passSAP").InnerText;

                vCompany = new SAPbobsCOM.Company();

                vCompany.language = BoSuppLangs.ln_Portuguese_Br;
                vCompany.DbUserName = Crypto.Decriptar(sUser);
                vCompany.DbPassword = Crypto.Decriptar(sPass);
                vCompany.CompanyDB = Crypto.Decriptar(sBanco);
                vCompany.UserName = Crypto.Decriptar(sUserSAP);
                vCompany.Password = Crypto.Decriptar(sPassSAP);
                vCompany.Server = "SQL";//Crypto.Decriptar(sHost);
                //vCompany.LicenseServer = "192.168.0.246:30000";
                

                Console.WriteLine(Crypto.Decriptar(sUser));
                Console.WriteLine(Crypto.Decriptar(sPass));
                Console.WriteLine(Crypto.Decriptar(sBanco));
                Console.WriteLine(Crypto.Decriptar(sUserSAP));
                Console.WriteLine(Crypto.Decriptar(sPassSAP));
                Console.WriteLine(Crypto.Decriptar(sHost));


                switch (Crypto.Decriptar(sVersion))
                {
                    case "MSSQL - 2005":
                        vCompany.DbServerType = BoDataServerTypes.dst_MSSQL2005;
                        break;

                    case "MSSQL - 2008":
                        vCompany.DbServerType = BoDataServerTypes.dst_MSSQL2008;
                        break;

                    case "MSSQL - 2012":
                        vCompany.DbServerType = BoDataServerTypes.dst_MSSQL2012;
                        break;



                    default:
                        vCompany.DbServerType = BoDataServerTypes.dst_MSSQL;
                        break;
                }

                if (vCompany.Connect() != 0)
                {
                    string last_err_msg = vCompany.GetLastErrorDescription();
                    
                    throw new Exception(last_err_msg);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Falha na conexão com o SAP.", ex.Message);
                return (false);
            }

            return (true);
        }
        #endregion

        #region Rotinas de GetInfos

        public static DataTable TableUpdBP()
        {
          DataTable oDT = new System.Data.DataTable("TableUpdBP");
          Recordset oRs = (Recordset)vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);


          try
          {
            oDT.Columns.Add("u_ptl_id", typeof(System.Int32));
            oDT.Columns.Add("cardcode", typeof(System.String));
            oDT.Columns.Add("cardtype", typeof(System.String));
            oDT.Columns.Add("frozenfor", typeof(System.String));	
            oDT.Columns.Add("frozencomm", typeof(System.String));


            string query = "";

            query = "SELECT u_ptl_id, cardcode, cardtype, frozenfor, frozencomm FROM ocrd WHERE ( " +
                    "createdate = Cast(Replace(CONVERT(NVARCHAR(max), Getdate(), 102), '.', '-') AS " +
                    "DATETIME) AND Isnull(u_ptl_id, 0) <> 0 ) OR ( updatedate = Cast(Replace( " +
                    "CONVERT(NVARCHAR(max), Getdate(), 112), '.', '-') AS DATETIME) AND Isnull( " +
                    "u_ptl_id, 0) <> 0 )";

            oRs.DoQuery(query);

            System.Data.DataRow oRow = null;

            while (!oRs.EoF)
            {
              oRow = oDT.NewRow();

              oRow["u_ptl_id"] = oRs.Fields.Item("u_ptl_id").Value.ToString().Trim();
              oRow["cardcode"] = oRs.Fields.Item("cardcode").Value.ToString().Trim();
              oRow["cardtype"] = oRs.Fields.Item("cardtype").Value.ToString().Trim();
              oRow["frozenfor"] = oRs.Fields.Item("frozenfor").Value.ToString().Trim();
              oRow["frozencomm"] = oRs.Fields.Item("frozencomm").Value.ToString().Trim();
              
              oDT.Rows.Add(oRow);

              oRs.MoveNext();

            }

            Marshal.ReleaseComObject(oRs);
            oRs = null;
          }
          catch { }
          finally
          {
            if (oRs != null)
              Marshal.ReleaseComObject(oRs);
          }

          return oDT;
        }

        public static DataTable TableTransp()
        {
          DataTable oDT = new System.Data.DataTable("TableTransp");
          Recordset oRs = (Recordset)vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

          try
          {
            oDT.Columns.Add("Code", typeof(System.String));
            oDT.Columns.Add("Nome", typeof(System.String));

            string query = "";
            query = "select CardCode Code, replace(CardCode,'T','') + ' - ' + CardName  as Nome from OCRD where CardCode like 'T%'";
            oRs.DoQuery(query);
            System.Data.DataRow oRow = null;
            while (!oRs.EoF)
            {
              oRow = oDT.NewRow();
              oRow["Code"] = oRs.Fields.Item("Code").Value.ToString().Trim();
              oRow["Nome"] = oRs.Fields.Item("Nome").Value.ToString().Trim();
              oDT.Rows.Add(oRow);
              oRs.MoveNext();
            }
            Marshal.ReleaseComObject(oRs);
            oRs = null;
          }
          catch { }
          finally
          {
            if (oRs != null)
              Marshal.ReleaseComObject(oRs);
          }

          return oDT;
        }

        public static DataTable TablePriceList(string itemcode, string indexador)
        {
            DataTable oDT = new System.Data.DataTable("TableTransp");
            Recordset oRs = (Recordset)vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            try
            {
                //oDT.Columns.Add("id", typeof(System.String));
                oDT.Columns.Add("fk_produto", typeof(System.String));
                oDT.Columns.Add("preco", typeof(System.String));
                oDT.Columns.Add("fk_tabela", typeof(System.String));

                string query = "";
                query = "select  ItemCode fk_produto, cast(round(Price,2)  as nvarchar(max)) preco, PriceList fk_tabela from SBO_BUW_EMCOMEX_PROD_FINAL..ITM1 where ItemCode = '" + itemcode+"'";
                oRs.DoQuery(query);
                System.Data.DataRow oRow = null;
                while (!oRs.EoF)
                {
                    oRow = oDT.NewRow();
                    //oRow["id"] = oRs.Fields.Item("id").Value.ToString().Trim();
                    oRow["fk_produto"] = oRs.Fields.Item("fk_produto").Value.ToString().Trim();
                    oRow["preco"] = oRs.Fields.Item("preco").Value.ToString().Trim();
                    oRow["fk_tabela"] = oRs.Fields.Item("fk_tabela").Value.ToString().Trim();

                    oDT.Rows.Add(oRow);
                    oRs.MoveNext();
                }
                Marshal.ReleaseComObject(oRs);
                oRs = null;
            }
            catch { }
            finally
            {
                if (oRs != null)
                    Marshal.ReleaseComObject(oRs);
            }

            return oDT;
        }
        


        public static DataTable TableCond()
        {
          DataTable oDT = new System.Data.DataTable("TableCond");
          Recordset oRs = (Recordset)vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

          try
          {
            oDT.Columns.Add("Code", typeof(System.String));
            oDT.Columns.Add("Nome", typeof(System.String));

            string query = "";
            query = "select GroupNum Code, PymntGroup Nome from OCTG ";
            oRs.DoQuery(query);
            System.Data.DataRow oRow = null;
            while (!oRs.EoF)
            {
              oRow = oDT.NewRow();
              oRow["Code"] = oRs.Fields.Item("Code").Value.ToString().Trim();
              oRow["Nome"] = oRs.Fields.Item("Nome").Value.ToString().Trim();
              oDT.Rows.Add(oRow);
              oRs.MoveNext();
            }
            Marshal.ReleaseComObject(oRs);
            oRs = null;
          }
          catch { }
          finally
          {
            if (oRs != null)
              Marshal.ReleaseComObject(oRs);
          }

          return oDT;
        }

        public static DataTable TableItem()
        {
          DataTable oDT = new System.Data.DataTable("TableItem");
          Recordset oRs = (Recordset)vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

          try
          {
            oDT.Columns.Add("Code", typeof(System.String));
            oDT.Columns.Add("Nome", typeof(System.String));
            oDT.Columns.Add("ipi", typeof(System.String));
            oDT.Columns.Add("imagem", typeof(System.String));

            string query = "";
            query = "select u_Code Code, u_name Nome, U_IPI ipi, u_IMAGE imagem from [@soni_PTL_LISTAPRECO]";
            oRs.DoQuery(query);
            System.Data.DataRow oRow = null;
            while (!oRs.EoF)
            {
              oRow = oDT.NewRow();
              oRow["Code"] = oRs.Fields.Item("Code").Value.ToString().Trim();
              oRow["Nome"] = oRs.Fields.Item("Nome").Value.ToString().Trim();
              oRow["ipi"] = oRs.Fields.Item("ipi").Value.ToString().Trim();
              oRow["imagem"] = oRs.Fields.Item("imagem").Value.ToString().Trim();
              oDT.Rows.Add(oRow);
              oRs.MoveNext();
            }
            Marshal.ReleaseComObject(oRs);
            oRs = null;
          }
          catch { }
          finally
          {
            if (oRs != null)
              Marshal.ReleaseComObject(oRs);
          }

          return oDT;
        }

        public static DataTable OITWStock(string filial )
        {
            DataTable oDT = new System.Data.DataTable("OITWStock");
            Recordset oRs = (Recordset)vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            try
            {
                oDT.Columns.Add("itemcode", typeof(System.String));
                oDT.Columns.Add("disp", typeof(System.String));

                string query = "";
                query = "SELECT t0.itemcode itemcode, replace(cast(cast(t0.onhand as numeric(19,6)) as nvarchar(max)),'.',',') disp FROM   oitw t0 LEFT OUTER JOIN owhs t1 ON t0.whscode = t1.whscode WHERE  t0.whscode IN ( 'MSP01-01', 'FRJ01-01', 'FMG01-01', 'MSP02-01' ) and t1.bplid = "+filial+"";
                oRs.DoQuery(query);
                System.Data.DataRow oRow = null;
                while (!oRs.EoF)
                {
                    oRow = oDT.NewRow();
                    oRow["itemcode"] = oRs.Fields.Item("itemcode").Value.ToString().Trim();
                    oRow["disp"] = oRs.Fields.Item("disp").Value.ToString().Trim();
                    oDT.Rows.Add(oRow);
                    oRs.MoveNext();
                }
                Marshal.ReleaseComObject(oRs);
                oRs = null;
            }
            catch { }
            finally
            {
                if (oRs != null)
                    Marshal.ReleaseComObject(oRs);
            }

            return oDT;
        }

        public static DataTable LCredSAP()
        {
            DataTable oDT = new System.Data.DataTable("LCredSAP");
            Recordset oRs = (Recordset)vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            try
            {
                oDT.Columns.Add("CardCode", typeof(System.String));
                oDT.Columns.Add("lcred", typeof(System.String));
                oDT.Columns.Add("vlavenc", typeof(System.String));
                oDT.Columns.Add("vlvenc", typeof(System.String));
                oDT.Columns.Add("creddist", typeof(System.String));

                string query = "";
                query = "select cast([Código do Cliente] as nvarchar(max)) CardCode,cast(isnull([Limite de Crédito],0)as nvarchar(max)) lcred,cast(isnull([Valor a Vencer],0) as nvarchar(max))	vlavenc,cast(isnull([Valor Vencido],0) as nvarchar(max))	vlvenc,cast(isnull([Crédio Disponível],0) as nvarchar(max)) creddist from UV_PortalLimiteCliente  ";
                oRs.DoQuery(query);
                System.Data.DataRow oRow = null;
                while (!oRs.EoF)
                {
                    oRow = oDT.NewRow();
                    oRow["CardCode"] = oRs.Fields.Item("CardCode").Value.ToString().Trim();
                    oRow["lcred"] = oRs.Fields.Item("lcred").Value.ToString().Trim();
                    oRow["vlavenc"] = oRs.Fields.Item("vlavenc").Value.ToString().Trim();
                    oRow["vlvenc"] = oRs.Fields.Item("vlvenc").Value.ToString().Trim();
                    oRow["creddist"] = oRs.Fields.Item("creddist").Value.ToString().Trim();
                    oDT.Rows.Add(oRow);
                    oRs.MoveNext();
                }
                Marshal.ReleaseComObject(oRs);
                oRs = null;
            }
            catch { }
            finally
            {
                if (oRs != null)
                    Marshal.ReleaseComObject(oRs);
            }

            return oDT;
        }


        public static DataTable StatusPed()
        {
            DataTable oDT = new System.Data.DataTable("StatusPed");
            Recordset oRs = (Recordset)vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            try
            {
                oDT.Columns.Add("U_IdPortal", typeof(System.String));
                oDT.Columns.Add("WDStatus", typeof(System.String));
                oDT.Columns.Add("Aproval1", typeof(System.String));
                oDT.Columns.Add("Aproval2", typeof(System.String));
                oDT.Columns.Add("Aproval3", typeof(System.String));
                oDT.Columns.Add("Aproval4", typeof(System.String));


                string query = "";
                query = "select isnull(cast(U_IdPortal as nvarchar(max)),0)  U_IdPortal,  max(cast(isnull(WDStatus,'') as nvarchar(max))) WDStatus ,  max(cast(isnull(Aproval1,'') as nvarchar(max))) Aproval1 ,  max(cast(isnull(Aproval2,'') as nvarchar(max))) Aproval2 ,  max(cast(isnull(Aproval3,'') as nvarchar(max))) Aproval3 ,  max(cast(isnull(Aproval4,'') as nvarchar(max))) Aproval4  from [UV_PedStatus] group by isnull(cast(U_IdPortal as nvarchar(max)),0)";
                oRs.DoQuery(query);
                System.Data.DataRow oRow = null;
                while (!oRs.EoF)
                {
                    oRow = oDT.NewRow();
                    oRow["U_IdPortal"] = oRs.Fields.Item("U_IdPortal").Value.ToString().Trim();
                    oRow["WDStatus"] = oRs.Fields.Item("WDStatus").Value.ToString().Trim();
                    oRow["Aproval1"] = oRs.Fields.Item("Aproval1").Value.ToString().Trim();
                    oRow["Aproval2"] = oRs.Fields.Item("Aproval2").Value.ToString().Trim();
                    oRow["Aproval3"] = oRs.Fields.Item("Aproval3").Value.ToString().Trim();
                    oRow["Aproval4"] = oRs.Fields.Item("Aproval4").Value.ToString().Trim();

                    oDT.Rows.Add(oRow);
                    oRs.MoveNext();
                }
                Marshal.ReleaseComObject(oRs);
                oRs = null;
            }
            catch { }
            finally
            {
                if (oRs != null)
                    Marshal.ReleaseComObject(oRs);
            }

            return oDT;
        }

        public static DataTable UltCompra()
        {
            DataTable oDT = new System.Data.DataTable("UltCompra");
            Recordset oRs = (Recordset)vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            try
            {
                oDT.Columns.Add("cardcode", typeof(System.String));
                oDT.Columns.Add("dataa", typeof(System.String));



                string query = "";
                query = "SELECT cardcode, replace(convert(nvarchar(max),Max(docdate),111 ),'/','-')  dataa FROM oinv GROUP  BY cardcode";
                oRs.DoQuery(query);
                System.Data.DataRow oRow = null;
                while (!oRs.EoF)
                {
                    oRow = oDT.NewRow();
                    oRow["cardcode"] = oRs.Fields.Item("cardcode").Value.ToString().Trim();
                    oRow["dataa"] = oRs.Fields.Item("dataa").Value.ToString().Trim();


                    oDT.Rows.Add(oRow);
                    oRs.MoveNext();
                }
                Marshal.ReleaseComObject(oRs);
                oRs = null;
            }
            catch { }
            finally
            {
                if (oRs != null)
                    Marshal.ReleaseComObject(oRs);
            }

            return oDT;
        }

        public static DataTable TableItemSAP( string itemcode)
        {
            DataTable oDT = new System.Data.DataTable("TableItemSAP");
            Recordset oRs = (Recordset)vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            try
            {
                oDT.Columns.Add("ItemCode", typeof(System.String));
                oDT.Columns.Add("1", typeof(System.String));
                oDT.Columns.Add("2", typeof(System.String));
                oDT.Columns.Add("3", typeof(System.String));
                oDT.Columns.Add("4", typeof(System.String));
                oDT.Columns.Add("5", typeof(System.String));
                oDT.Columns.Add("6", typeof(System.String));
                oDT.Columns.Add("7", typeof(System.String));


                string query = "";
                query = "select * from [UF_PortalStock]('"+itemcode+"')";
                oRs.DoQuery(query);
                System.Data.DataRow oRow = null;
                while (!oRs.EoF)
                {
                    oRow = oDT.NewRow();
                    oRow["ItemCode"] = oRs.Fields.Item("ItemCode").Value.ToString().Trim();
                    oRow["1"] = oRs.Fields.Item("1").Value.ToString().Trim();
                    oRow["2"] = oRs.Fields.Item("2").Value.ToString().Trim();
                    oRow["3"] = oRs.Fields.Item("3").Value.ToString().Trim();
                    oRow["4"] = oRs.Fields.Item("4").Value.ToString().Trim();
                    oRow["5"] = oRs.Fields.Item("5").Value.ToString().Trim();
                    oRow["6"] = oRs.Fields.Item("6").Value.ToString().Trim();
                    oRow["7"] = oRs.Fields.Item("7").Value.ToString().Trim();
                    oDT.Rows.Add(oRow);
                    oRs.MoveNext();
                }
                Marshal.ReleaseComObject(oRs);
                oRs = null;
            }
            catch { }
            finally
            {
                if (oRs != null)
                    Marshal.ReleaseComObject(oRs);
            }

            return oDT;
        }



        public static DataTable TablePNSAP(string cardcode)
        {
            DataTable oDT = new System.Data.DataTable("TablePNSAP");
            Recordset oRs = (Recordset)vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            try
            {
                oDT.Columns.Add("codcliente", typeof(System.String));
                oDT.Columns.Add("nome_cliente", typeof(System.String));
                oDT.Columns.Add("cnpj", typeof(System.String));
                oDT.Columns.Add("cpf", typeof(System.String));
                oDT.Columns.Add("email", typeof(System.String));
                oDT.Columns.Add("fk_filial", typeof(System.String));
                oDT.Columns.Add("ultcomp", typeof(System.String));
                oDT.Columns.Add("tel", typeof(System.String));
                oDT.Columns.Add("cel", typeof(System.String));
                oDT.Columns.Add("lcompro", typeof(System.String));
                oDT.Columns.Add("lcred", typeof(System.String));
                oDT.Columns.Add("ldisp", typeof(System.String));
                oDT.Columns.Add("inadi", typeof(System.String));
                oDT.Columns.Add("Cond", typeof(System.String));
                oDT.Columns.Add("fk_tabela", typeof(System.String));
                oDT.Columns.Add("SlpName", typeof(System.String));
                oDT.Columns.Add("Endereco", typeof(System.String));
                oDT.Columns.Add("CEP", typeof(System.String));
                oDT.Columns.Add("building", typeof(System.String));
                oDT.Columns.Add("city", typeof(System.String));
                oDT.Columns.Add("state", typeof(System.String));
                oDT.Columns.Add("block", typeof(System.String));
                oDT.Columns.Add("CardFName", typeof(System.String));
                oDT.Columns.Add("Inativo", typeof(System.String));



                string query = "";
                query = "select * from [UF_PortalPN]('" + cardcode + "')";
                oRs.DoQuery(query);
                System.Data.DataRow oRow = null;
                while (!oRs.EoF)
                {
                    oRow = oDT.NewRow();
                    oRow["codcliente"] = oRs.Fields.Item("codcliente").Value.ToString().Trim();
                    oRow["nome_cliente"] = oRs.Fields.Item("nome_cliente").Value.ToString().Trim();
                    oRow["cnpj"] = oRs.Fields.Item("cnpj").Value.ToString().Trim();
                    oRow["cpf"] = oRs.Fields.Item("cpf").Value.ToString().Trim();
                    oRow["email"] = oRs.Fields.Item("email").Value.ToString().Trim();
                    oRow["fk_filial"] = oRs.Fields.Item("fk_filial").Value.ToString().Trim();
                    oRow["ultcomp"] = oRs.Fields.Item("ultcomp").Value.ToString().Trim();
                    oRow["tel"] = oRs.Fields.Item("tel").Value.ToString().Trim();
                    oRow["cel"] = oRs.Fields.Item("cel").Value.ToString().Trim();
                    oRow["lcompro"] = oRs.Fields.Item("lcompro").Value.ToString().Trim();
                    oRow["lcred"] = oRs.Fields.Item("lcred").Value.ToString().Trim();
                    oRow["ldisp"] = oRs.Fields.Item("ldisp").Value.ToString().Trim();
                    oRow["inadi"] = oRs.Fields.Item("inadi").Value.ToString().Trim();
                    oRow["Cond"] = oRs.Fields.Item("Cond").Value.ToString().Trim();
                    oRow["fk_tabela"] = oRs.Fields.Item("fk_tabela").Value.ToString().Trim();
                    oRow["SlpName"] = oRs.Fields.Item("SlpName").Value.ToString().Trim();
                    oRow["Endereco"] = oRs.Fields.Item("Endereco").Value.ToString().Trim();
                    oRow["CEP"] = oRs.Fields.Item("CEP").Value.ToString().Trim();
                    oRow["building"] = oRs.Fields.Item("building").Value.ToString().Trim();
                    oRow["city"] = oRs.Fields.Item("city").Value.ToString().Trim();
                    oRow["state"] = oRs.Fields.Item("state").Value.ToString().Trim();
                    oRow["block"] = oRs.Fields.Item("block").Value.ToString().Trim();
                    oRow["CardFName"] = oRs.Fields.Item("CardFName").Value.ToString().Trim();
                    oRow["Inativo"] = oRs.Fields.Item("Inativo").Value.ToString().Trim();

                    oDT.Rows.Add(oRow);
                    oRs.MoveNext();
                }
                Marshal.ReleaseComObject(oRs);
                oRs = null;
            }
            catch { }
            finally
            {
                if (oRs != null)
                    Marshal.ReleaseComObject(oRs);
            }

            return oDT;
        }


        public static DataTable TableChefiaSAP(string cardcode)
        {
            DataTable oDT = new System.Data.DataTable("TableChefiaSAP");
            Recordset oRs = (Recordset)vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            try
            {
                oDT.Columns.Add("codcliente", typeof(System.String));
                oDT.Columns.Add("nome_cliente", typeof(System.String));
      

                string query = "";
                query = "select * from [UF_PortalPN]('" + cardcode + "')";
                oRs.DoQuery(query);
                System.Data.DataRow oRow = null;
                while (!oRs.EoF)
                {
                    oRow = oDT.NewRow();
                    oRow["codcliente"] = oRs.Fields.Item("codcliente").Value.ToString().Trim();
                    oRow["nome_cliente"] = oRs.Fields.Item("nome_cliente").Value.ToString().Trim();
      

                    oDT.Rows.Add(oRow);
                    oRs.MoveNext();
                }
                Marshal.ReleaseComObject(oRs);
                oRs = null;
            }
            catch { }
            finally
            {
                if (oRs != null)
                    Marshal.ReleaseComObject(oRs);
            }

            return oDT;
        }

        public static DataTable TablePrice()
        {
          DataTable oDT = new System.Data.DataTable("TablePrice");
          Recordset oRs = (Recordset)vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

          try
          {
            oDT.Columns.Add("Cliente", typeof(System.String));
            oDT.Columns.Add("Item", typeof(System.String));
            oDT.Columns.Add("Valor", typeof(System.String));
            oDT.Columns.Add("Origem", typeof(System.String));

            string query = "";
            query = "select * from ListPricePortal";
            oRs.DoQuery(query);
            System.Data.DataRow oRow = null;
            while (!oRs.EoF)
            {
              oRow = oDT.NewRow();
              oRow["Cliente"] = oRs.Fields.Item("Cliente").Value.ToString().Trim();
              oRow["Item"] = oRs.Fields.Item("Item").Value.ToString().Trim();
              oRow["Valor"] = oRs.Fields.Item("Valor").Value.ToString().Trim();
              oRow["Origem"] = oRs.Fields.Item("Origem").Value.ToString().Trim();
              oDT.Rows.Add(oRow);
              oRs.MoveNext();
            }
            Marshal.ReleaseComObject(oRs);
            oRs = null;
          }
          catch { }
          finally
          {
            if (oRs != null)
              Marshal.ReleaseComObject(oRs);
          }

          return oDT;
        }
        public static DataTable TablePN()
        {
          DataTable oDT = new System.Data.DataTable("TablePN");
          Recordset oRs = (Recordset)vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

          try
          {
            oDT.Columns.Add("Codigo", typeof(System.String));
            oDT.Columns.Add("Razao", typeof(System.String));
            oDT.Columns.Add("Contato", typeof(System.String));
            oDT.Columns.Add("Fone", typeof(System.String));
            oDT.Columns.Add("Email", typeof(System.String));
            oDT.Columns.Add("EndCob", typeof(System.String));
            oDT.Columns.Add("EndEnt", typeof(System.String));
            oDT.Columns.Add("CNPJ", typeof(System.String));
            oDT.Columns.Add("CPF", typeof(System.String));
            oDT.Columns.Add("IE", typeof(System.String));
            oDT.Columns.Add("CondPagto", typeof(System.String));
            oDT.Columns.Add("UserId", typeof(System.String));

            String varname1 = "";
            varname1 = varname1 + "SELECT t0.cardcode Codigo, replace(t0.CardCode,'C','') + ' - ' + t0.cardname Razao, m1.contato Contato, t0.phone2 + ' ' + t0.phone1 " + "\n";
            varname1 = varname1 + "Fone, t0.e_mail Email, Bill.adr EndCob, Ship.adr EndEnt, m0.taxid0 CNPJ, m0.taxid4 CPF, " + "\n";
            varname1 = varname1 + "m0.taxid1 IE, case when isnull(t2.pymntgroup,'') = '' then 'A VISTA' else t2.pymntgroup end CondPagto, t3.u_ptl_vendedor UserId FROM ocrd t0 LEFT OUTER JOIN ( " + "\n";
            varname1 = varname1 + "SELECT cardcode, taxid0, taxid4, taxid1 FROM crd7 WHERE address = '') m0 ON m0.cardcode = t0.cardcode " + "\n";
            varname1 = varname1 + "LEFT OUTER JOIN (SELECT cardcode, Max(NAME) Contato FROM ocpr GROUP BY cardcode) m1 ON m1.cardcode " + "\n";
            varname1 = varname1 + "= t0.cardcode LEFT OUTER JOIN (SELECT p0.cardcode, l2.pymntgroup FROM (SELECT l0.cardcode, Max( " + "\n";
            varname1 = varname1 + "l0.docentry) 'Maxentry' FROM oinv l0 GROUP BY l0.cardcode) P0 LEFT OUTER JOIN oinv l1 ON P0.maxentry " + "\n";
            varname1 = varname1 + "= l1.docentry LEFT OUTER JOIN octg l2 ON l2.groupnum = l1.groupnum) t2 ON t0.cardcode = t2.cardcode " + "\n";
            varname1 = varname1 + "INNER JOIN oslp t3 ON t3.slpcode = t0.slpcode LEFT OUTER JOIN (SELECT cardcode, address, " + "\n";
            varname1 = varname1 + "Isnull(addrtype, '') + ' ' + Isnull(street, '') + ', ' + Isnull(streetno, '') + ' - ' + Isnull( " + "\n";
            varname1 = varname1 + "block, '') + ' - ' + Isnull(city, '') + ' - ' + Isnull(state, '') + ' CEP: ' + Isnull(zipcode, '') " + "\n";
            varname1 = varname1 + "adr FROM crd1 WHERE adrestype = 'B') Bill ON Bill.cardcode = t0.cardcode AND Bill.address = t0.billtodef " + "\n";
            varname1 = varname1 + "LEFT OUTER JOIN (SELECT cardcode, address, Isnull(addrtype, '') + ' ' + Isnull(street, '') + " + "\n";
            varname1 = varname1 + "', ' + Isnull(streetno, '') + ' - ' + Isnull(block, '') + ' - ' + Isnull(city, '') + ' - ' + " + "\n";
            varname1 = varname1 + "Isnull(state, '') + ' CEP: ' + Isnull(zipcode, '') adr FROM crd1 WHERE adrestype = 'S') Ship ON " + "\n";
            varname1 = varname1 + "Ship.cardcode = t0.cardcode AND Ship.address = t0.shiptodef WHERE t0.cardtype = 'C' " + "\n";
            varname1 = varname1 + "/* --and t3.U_PTL_Vendedor is not null*/";

            oRs.DoQuery(varname1);
            System.Data.DataRow oRow = null;
            while (!oRs.EoF)
            {
              oRow = oDT.NewRow();
              oRow["Codigo"] = oRs.Fields.Item("Codigo").Value.ToString().Trim();
              oRow["Razao"] = oRs.Fields.Item("Razao").Value.ToString().Trim();
              oRow["Contato"] = oRs.Fields.Item("Contato").Value.ToString().Trim();
              oRow["Fone"] = oRs.Fields.Item("Fone").Value.ToString().Trim();
              oRow["Email"] = oRs.Fields.Item("Email").Value.ToString().Trim();
              oRow["EndCob"] = oRs.Fields.Item("EndCob").Value.ToString().Trim();
              oRow["EndEnt"] = oRs.Fields.Item("EndEnt").Value.ToString().Trim();
              oRow["CNPJ"] = oRs.Fields.Item("CNPJ").Value.ToString().Trim();
              oRow["CPF"] = oRs.Fields.Item("CPF").Value.ToString().Trim();
              oRow["IE"] = oRs.Fields.Item("IE").Value.ToString().Trim();
              oRow["CondPagto"] = oRs.Fields.Item("CondPagto").Value.ToString().Trim();
              oRow["UserId"] = oRs.Fields.Item("UserId").Value.ToString().Trim();



              oDT.Rows.Add(oRow);
              oRs.MoveNext();
            }
            Marshal.ReleaseComObject(oRs);
            oRs = null;
          }
          catch { }
          finally
          {
            if (oRs != null)
              Marshal.ReleaseComObject(oRs);
          }

          return oDT;
        }

        public static DataTable TableSinglePN(string cardcode)
        {
          DataTable oDT = new System.Data.DataTable("TableSinglePN");
          Recordset oRs = (Recordset)vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

          try
          {
            oDT.Columns.Add("Codigo", typeof(System.String));
            oDT.Columns.Add("Razao", typeof(System.String));
            oDT.Columns.Add("Contato", typeof(System.String));
            oDT.Columns.Add("Fone", typeof(System.String));
            oDT.Columns.Add("Email", typeof(System.String));
            oDT.Columns.Add("EndCob", typeof(System.String));
            oDT.Columns.Add("EndEnt", typeof(System.String));
            oDT.Columns.Add("CNPJ", typeof(System.String));
            oDT.Columns.Add("CPF", typeof(System.String));
            oDT.Columns.Add("IE", typeof(System.String));
            oDT.Columns.Add("CondPagto", typeof(System.String));
            oDT.Columns.Add("UserId", typeof(System.String));

            String varname1 = "";
            varname1 = varname1 + "SELECT t0.cardcode Codigo, replace(t0.CardCode,'C','') + ' - ' + t0.cardname Razao, m1.contato Contato, t0.phone2 + ' ' + t0.phone1 " + "\n";
            varname1 = varname1 + "Fone, t0.e_mail Email, Bill.adr EndCob, Ship.adr EndEnt, m0.taxid0 CNPJ, m0.taxid4 CPF, " + "\n";
            varname1 = varname1 + "m0.taxid1 IE, case when isnull(t2.pymntgroup,'') = '' then 'A VISTA' else t2.pymntgroup end CondPagto, t3.u_ptl_vendedor UserId FROM ocrd t0 LEFT OUTER JOIN ( " + "\n";
            varname1 = varname1 + "SELECT cardcode, taxid0, taxid4, taxid1 FROM crd7 WHERE address = '') m0 ON m0.cardcode = t0.cardcode " + "\n";
            varname1 = varname1 + "LEFT OUTER JOIN (SELECT cardcode, Max(NAME) Contato FROM ocpr GROUP BY cardcode) m1 ON m1.cardcode " + "\n";
            varname1 = varname1 + "= t0.cardcode LEFT OUTER JOIN (SELECT p0.cardcode, l2.pymntgroup FROM (SELECT l0.cardcode, Max( " + "\n";
            varname1 = varname1 + "l0.docentry) 'Maxentry' FROM oinv l0 GROUP BY l0.cardcode) P0 LEFT OUTER JOIN oinv l1 ON P0.maxentry " + "\n";
            varname1 = varname1 + "= l1.docentry LEFT OUTER JOIN octg l2 ON l2.groupnum = l1.groupnum) t2 ON t0.cardcode = t2.cardcode " + "\n";
            varname1 = varname1 + "INNER JOIN oslp t3 ON t3.slpcode = t0.slpcode LEFT OUTER JOIN (SELECT cardcode, address, " + "\n";
            varname1 = varname1 + "Isnull(addrtype, '') + ' ' + Isnull(street, '') + ', ' + Isnull(streetno, '') + ' - ' + Isnull( " + "\n";
            varname1 = varname1 + "block, '') + ' - ' + Isnull(city, '') + ' - ' + Isnull(state, '') + ' CEP: ' + Isnull(zipcode, '') " + "\n";
            varname1 = varname1 + "adr FROM crd1 WHERE adrestype = 'B') Bill ON Bill.cardcode = t0.cardcode AND Bill.address = t0.billtodef " + "\n";
            varname1 = varname1 + "LEFT OUTER JOIN (SELECT cardcode, address, Isnull(addrtype, '') + ' ' + Isnull(street, '') + " + "\n";
            varname1 = varname1 + "', ' + Isnull(streetno, '') + ' - ' + Isnull(block, '') + ' - ' + Isnull(city, '') + ' - ' + " + "\n";
            varname1 = varname1 + "Isnull(state, '') + ' CEP: ' + Isnull(zipcode, '') adr FROM crd1 WHERE adrestype = 'S') Ship ON " + "\n";
            varname1 = varname1 + "Ship.cardcode = t0.cardcode AND Ship.address = t0.shiptodef WHERE t0.cardtype = 'C'  and t0.validFor = 'y' and t0.CardCode = '"+cardcode+"' " + "\n";
            varname1 = varname1 + "/* --and t3.U_PTL_Vendedor is not null*/";

            oRs.DoQuery(varname1);
            System.Data.DataRow oRow = null;
            while (!oRs.EoF)
            {
              oRow = oDT.NewRow();
              oRow["Codigo"] = oRs.Fields.Item("Codigo").Value.ToString().Trim();
              oRow["Razao"] = oRs.Fields.Item("Razao").Value.ToString().Trim();
              oRow["Contato"] = oRs.Fields.Item("Contato").Value.ToString().Trim();
              oRow["Fone"] = oRs.Fields.Item("Fone").Value.ToString().Trim();
              oRow["Email"] = oRs.Fields.Item("Email").Value.ToString().Trim();
              oRow["EndCob"] = oRs.Fields.Item("EndCob").Value.ToString().Trim();
              oRow["EndEnt"] = oRs.Fields.Item("EndEnt").Value.ToString().Trim();
              oRow["CNPJ"] = oRs.Fields.Item("CNPJ").Value.ToString().Trim();
              oRow["CPF"] = oRs.Fields.Item("CPF").Value.ToString().Trim();
              oRow["IE"] = oRs.Fields.Item("IE").Value.ToString().Trim();
              oRow["CondPagto"] = oRs.Fields.Item("CondPagto").Value.ToString().Trim();
              oRow["UserId"] = oRs.Fields.Item("UserId").Value.ToString().Trim();



              oDT.Rows.Add(oRow);
              oRs.MoveNext();
            }
            Marshal.ReleaseComObject(oRs);
            oRs = null;
          }
          catch { }
          finally
          {
            if (oRs != null)
              Marshal.ReleaseComObject(oRs);
          }

          return oDT;
        }
        public static DataTable TableSiglePNPrice(string cardcode)
        {
          DataTable oDT = new System.Data.DataTable("TableSiglePNPrice");
          Recordset oRs = (Recordset)vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

          try
          {
            oDT.Columns.Add("Cliente", typeof(System.String));
            oDT.Columns.Add("Item", typeof(System.String));
            oDT.Columns.Add("Valor", typeof(System.String));
            oDT.Columns.Add("Origem", typeof(System.String));

            string query = "";
            query = "select * from ListPricePortal where Cliente = '" + cardcode + "'";
            oRs.DoQuery(query);
            System.Data.DataRow oRow = null;
            while (!oRs.EoF)
            {
              oRow = oDT.NewRow();
              oRow["Cliente"] = oRs.Fields.Item("Cliente").Value.ToString().Trim();
              oRow["Item"] = oRs.Fields.Item("Item").Value.ToString().Trim();
              oRow["Valor"] = oRs.Fields.Item("Valor").Value.ToString().Trim();
              oRow["Origem"] = oRs.Fields.Item("Origem").Value.ToString().Trim();
              oDT.Rows.Add(oRow);
              oRs.MoveNext();
            }
            Marshal.ReleaseComObject(oRs);
            oRs = null;
          }
          catch { }
          finally
          {
            if (oRs != null)
              Marshal.ReleaseComObject(oRs);
          }

          return oDT;
        }
        public static DataTable TableSingleTransp(string cardcode)
        {
          DataTable oDT = new System.Data.DataTable("TableSingleTransp");
          Recordset oRs = (Recordset)vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

          try
          {
            oDT.Columns.Add("Code", typeof(System.String));
            oDT.Columns.Add("Nome", typeof(System.String));

            string query = "";
            query = "select CardCode Code, replace(CardCode,'T','') + ' - ' + CardName  as Nome from OCRD where CardCode like 'T%' and cardcode = '"+cardcode+"'";
            oRs.DoQuery(query);
            System.Data.DataRow oRow = null;
            while (!oRs.EoF)
            {
              oRow = oDT.NewRow();
              oRow["Code"] = oRs.Fields.Item("Code").Value.ToString().Trim();
              oRow["Nome"] = oRs.Fields.Item("Nome").Value.ToString().Trim();
              oDT.Rows.Add(oRow);
              oRs.MoveNext();
            }
            Marshal.ReleaseComObject(oRs);
            oRs = null;
          }
          catch { }
          finally
          {
            if (oRs != null)
              Marshal.ReleaseComObject(oRs);
          }

          return oDT;
        }
        public static DataTable TableSingleCond(string groupnum)
        {
          DataTable oDT = new System.Data.DataTable("TableSingleCond");
          Recordset oRs = (Recordset)vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

          try
          {
            oDT.Columns.Add("Code", typeof(System.String));
            oDT.Columns.Add("Nome", typeof(System.String));

            string query = "";
            query = "select GroupNum Code, PymntGroup Nome from OCTG where GroupNum = '" + groupnum + "'";
            oRs.DoQuery(query);
            System.Data.DataRow oRow = null;
            while (!oRs.EoF)
            {
              oRow = oDT.NewRow();
              oRow["Code"] = oRs.Fields.Item("Code").Value.ToString().Trim();
              oRow["Nome"] = oRs.Fields.Item("Nome").Value.ToString().Trim();
              oDT.Rows.Add(oRow);
              oRs.MoveNext();
            }
            Marshal.ReleaseComObject(oRs);
            oRs = null;
          }
          catch { }
          finally
          {
            if (oRs != null)
              Marshal.ReleaseComObject(oRs);
          }

          return oDT;
        }

        public static DataTable TableSyncProcessar()
        {
          DataTable oDT = new System.Data.DataTable("TableSyncProcessar");
          Recordset oRs = (Recordset)vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

          try
          {
            oDT.Columns.Add("Code", typeof(System.String));
            oDT.Columns.Add("U_Chave", typeof(System.String));
            oDT.Columns.Add("U_Metodo", typeof(System.String));
            oDT.Columns.Add("U_Status", typeof(System.String));

            string query = "";
            query = "select * from [@SONI_PTL_PROCESSAR] where u_status = 'Processar'";
            oRs.DoQuery(query);
            System.Data.DataRow oRow = null;
            while (!oRs.EoF)
            {
              oRow = oDT.NewRow();
              oRow["Code"] = oRs.Fields.Item("Code").Value.ToString().Trim();
              oRow["U_Chave"] = oRs.Fields.Item("U_Chave").Value.ToString().Trim();
              oRow["U_Metodo"] = oRs.Fields.Item("U_Metodo").Value.ToString().Trim();
              oRow["U_Status"] = oRs.Fields.Item("U_Status").Value.ToString().Trim();
              oDT.Rows.Add(oRow);
              oRs.MoveNext();
            }
            Marshal.ReleaseComObject(oRs);
            oRs = null;
          }
          catch { }
          finally
          {
            if (oRs != null)
              Marshal.ReleaseComObject(oRs);
          }

          return oDT;
        }


        public static DataTable TableSincronization()
        {
            DataTable oDT = new System.Data.DataTable("TableSincronization");
            Recordset oRs = (Recordset)vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            try
            {
                oDT.Columns.Add("Code", typeof(System.String));
                oDT.Columns.Add("Name", typeof(System.String));
                oDT.Columns.Add("U_Chave", typeof(System.String));
                oDT.Columns.Add("U_Objeto", typeof(System.String));
                oDT.Columns.Add("U_Status", typeof(System.String));
                oDT.Columns.Add("U_Acao", typeof(System.String));
                oDT.Columns.Add("U_Data", typeof(System.String));

                

                string query = "";
                query = "select * from [@ZSINCRONIZATION] where u_status = 'Processar'";
                oRs.DoQuery(query);
                System.Data.DataRow oRow = null;
                while (!oRs.EoF)
                {
                    oRow = oDT.NewRow();
                    oRow["Code"] = oRs.Fields.Item("Name").Value.ToString().Trim();
                    oRow["Name"] = oRs.Fields.Item("Code").Value.ToString().Trim();
                    oRow["U_Chave"] = oRs.Fields.Item("U_Chave").Value.ToString().Trim();
                    oRow["U_Objeto"] = oRs.Fields.Item("U_Objeto").Value.ToString().Trim();
                    oRow["U_Status"] = oRs.Fields.Item("U_Status").Value.ToString().Trim();
                    oRow["U_Acao"] = oRs.Fields.Item("U_Acao").Value.ToString().Trim();
                    oRow["U_Data"] = oRs.Fields.Item("U_Data").Value.ToString().Trim();
                    oDT.Rows.Add(oRow);
                    oRs.MoveNext();
                }
                Marshal.ReleaseComObject(oRs);
                oRs = null;
            }
            catch { }
            finally
            {
                if (oRs != null)
                    Marshal.ReleaseComObject(oRs);
            }

            return oDT;
        }

        public static string GetParams(string code)
        {
            string retorno;
            string sql = string.Empty;
            Recordset oRS = null;
                oRS = (SAPbobsCOM.Recordset)vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                sql = "select u_Valor 'Name' from [@ZBUW_PARAMS] where Code = '" + code+"'";
                oRS.DoQuery(sql);
                retorno = Convert.ToString(oRS.Fields.Item("Name").Value);
            Marshal.ReleaseComObject(oRS);
                oRS = null;
                GC.Collect();
                return retorno;
        }

        public static string GetNomeChefia(string code)
        {
            string retorno;
            string sql = string.Empty;
            Recordset oRS = null;
            oRS = (SAPbobsCOM.Recordset)vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            sql = "select isnull(U_Cargo,'') U_Cargo from [@HIERARQUIA_TIPOS] where Code = '"+code+"'";
            oRS.DoQuery(sql);
            retorno = Convert.ToString(oRS.Fields.Item("U_Cargo").Value);
            Marshal.ReleaseComObject(oRS);
            oRS = null;
            GC.Collect();
            return retorno;
        }

        public static string GetNomeCarteira(string code)
        {
            string retorno;
            string sql = string.Empty;
            Recordset oRS = null;
            oRS = (SAPbobsCOM.Recordset)vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            sql = "select isnull(SlpName,'') SlpName from OSLP where SlpCode = '" + code + "'";
            oRS.DoQuery(sql);
            retorno = Convert.ToString(oRS.Fields.Item("SlpName").Value);
            Marshal.ReleaseComObject(oRS);
            oRS = null;
            GC.Collect();
            return retorno;
        }

        public static string Getidpedidoportalsap(string code)
        {
            string retorno;
            string sql = string.Empty;
            Recordset oRS = null;
            oRS = (SAPbobsCOM.Recordset)vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            sql = "select top 1 U_IdPortal from ORDR where DocEntry = "+code+"";
            oRS.DoQuery(sql);
            retorno = Convert.ToString(oRS.Fields.Item("U_IdPortal").Value);
            Marshal.ReleaseComObject(oRS);
            oRS = null;
            GC.Collect();
            return retorno;
        }

        public static string Getidpedidoportalsapinv(string code)
        {
            string retorno;
            string sql = string.Empty;
            Recordset oRS = null;
            oRS = (SAPbobsCOM.Recordset)vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            sql = "select top 1 U_IdPortal from oinv where DocEntry = "+code+"";
            oRS.DoQuery(sql);
            retorno = Convert.ToString(oRS.Fields.Item("U_IdPortal").Value);
            Marshal.ReleaseComObject(oRS);
            oRS = null;
            GC.Collect();
            return retorno;
        }

        public static string GetSerialInvoice(string code)
        {
            string retorno;
            string sql = string.Empty;
            Recordset oRS = null;
            oRS = (SAPbobsCOM.Recordset)vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            sql = "select top 1 serial from oinv where DocEntry = "+code+"";
            oRS.DoQuery(sql);
            retorno = Convert.ToString(oRS.Fields.Item("serial").Value);
            Marshal.ReleaseComObject(oRS);
            oRS = null;
            GC.Collect();
            return retorno;
        }


        public static string GetdtInvoice(string code)
        {
            string retorno;
            string sql = string.Empty;
            Recordset oRS = null;
            oRS = (SAPbobsCOM.Recordset)vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            sql = "select top 1 convert(nvarchar(10),DocDate, 126) dt from oinv where DocEntry = "+code+"";
            oRS.DoQuery(sql);
            retorno = Convert.ToString(oRS.Fields.Item("dt").Value);
            Marshal.ReleaseComObject(oRS);
            oRS = null;
            GC.Collect();
            return retorno;
        }

        public static string Getidpedidoportalsapdlv(string code)
        {
            string retorno;
            string sql = string.Empty;
            Recordset oRS = null;
            oRS = (SAPbobsCOM.Recordset)vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            sql = "select top 1 U_IdPortal from odln where DocEntry = " + code + "";
            oRS.DoQuery(sql);
            retorno = Convert.ToString(oRS.Fields.Item("U_IdPortal").Value);
            Marshal.ReleaseComObject(oRS);
            oRS = null;
            GC.Collect();
            return retorno;
        }

        public static string GetSerialDelivery(string code)
        {
            string retorno;
            string sql = string.Empty;
            Recordset oRS = null;
            oRS = (SAPbobsCOM.Recordset)vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            sql = "select top 1 serial from odln where DocEntry = " + code + "";
            oRS.DoQuery(sql);
            retorno = Convert.ToString(oRS.Fields.Item("serial").Value);
            Marshal.ReleaseComObject(oRS);
            oRS = null;
            GC.Collect();
            return retorno;
        }


        public static string GetdtDelivery(string code)
        {
            string retorno;
            string sql = string.Empty;
            Recordset oRS = null;
            oRS = (SAPbobsCOM.Recordset)vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            sql = "select top 1 convert(nvarchar(10),DocDate, 126) dt from odln where DocEntry = " + code + "";
            oRS.DoQuery(sql);
            retorno = Convert.ToString(oRS.Fields.Item("dt").Value);
            Marshal.ReleaseComObject(oRS);
            oRS = null;
            GC.Collect();
            return retorno;
        }

        public static string Getstatuspedidoportalsap(string code)
        {
            string retorno;
            string sql = string.Empty;
            Recordset oRS = null;
            oRS = (SAPbobsCOM.Recordset)vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            sql = "select top 1 u_motcancel from ORDR where DocEntry = "+code+"";
            oRS.DoQuery(sql);
            retorno = Convert.ToString(oRS.Fields.Item("u_motcancel").Value);
            Marshal.ReleaseComObject(oRS);
            oRS = null;
            GC.Collect();
            return retorno;
        }

        public static string GetStatusItMkt(string itemcode)
        {
            string retorno;
            string sql = string.Empty;
            Recordset oRS = null;
            oRS = (SAPbobsCOM.Recordset)vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            sql = "select QryGroup1 from oitm where itemcode = '"+itemcode+"'";
            oRS.DoQuery(sql);
            retorno = Convert.ToString(oRS.Fields.Item("QryGroup1").Value);
            Marshal.ReleaseComObject(oRS);
            oRS = null;
            GC.Collect();
            return retorno;
        }

        public static string GetStatusItMktv2(string itemcode)
        {
            string retorno;
            string sql = string.Empty;
            Recordset oRS = null;
            oRS = (SAPbobsCOM.Recordset)vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            //sql = "select QryGroup1 from oitm where itemcode = '" + itemcode + "'";
            sql = "SELECT  CASE WHEN '"+itemcode+"' like 'MP%' THEN    'Y'  ELSE   'N' END QryGroup1";  //Alterado 02/12/2018
            oRS.DoQuery(sql);
            retorno = Convert.ToString(oRS.Fields.Item("QryGroup1").Value);
            Marshal.ReleaseComObject(oRS);
            oRS = null;
            GC.Collect();
            return retorno;
        }

        public static string GetStatusItFin(string Cardcode, string cond, string total)
        {
            string retorno;
            string sql = string.Empty;
            Recordset oRS = null;
            oRS = (SAPbobsCOM.Recordset)vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            sql = "select dbo.[ApprovFIN]('"+Cardcode+"',"+cond+",'"+total+"') ret";
            oRS.DoQuery(sql);
            retorno = Convert.ToString(oRS.Fields.Item("ret").Value);
            Marshal.ReleaseComObject(oRS);
            oRS = null;
            GC.Collect();
            return retorno;
        }

        public static string GetTpPedido(string name)
        {
            string retorno;
            string sql = string.Empty;
            Recordset oRS = null;
            oRS = (SAPbobsCOM.Recordset)vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            sql = "select U_Valor valor from [@ZBUW_PARAMS] where Code like 'StatusPed%' and Name = '"+name+"'";
            oRS.DoQuery(sql);
            retorno = Convert.ToString(oRS.Fields.Item("valor").Value);
            Marshal.ReleaseComObject(oRS);
            oRS = null;
            GC.Collect();
            return retorno;
        }
        public static string GetdfUsage(string name)
        {
            string retorno;
            string sql = string.Empty;
            Recordset oRS = null;
            oRS = (SAPbobsCOM.Recordset)Utils.vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            sql = $"SELECT T0.\"ID\" FROM OUSG T0 WHERE T0.\"Usage\"  = '{name}'";
            oRS.DoQuery(sql);
            retorno = Convert.ToString(oRS.Fields.Item("ID").Value);
            Marshal.ReleaseComObject(oRS);
            oRS = null;
            GC.Collect();
            return retorno;
        }
        public static string GetValue(string query, string field)
        {
            string retorno;
            string sql = string.Empty;
            Recordset oRS = null;
            oRS = (SAPbobsCOM.Recordset)Utils.vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            //sql = $"SELECT T0.\"ID\" FROM OUSG T0 WHERE T0.\"Usage\"  = '{name}'";
            oRS.DoQuery(query);
            retorno = Convert.ToString(oRS.Fields.Item(field).Value);
            Marshal.ReleaseComObject(oRS);
            oRS = null;
            GC.Collect();
            return retorno;
        }
        public static string GetVendorName(string Idvendedor)
        {
          string retorno;
          string sql = string.Empty;
          Recordset oRS = null;
          oRS = (SAPbobsCOM.Recordset)vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
          sql = "select name from [@SONI_PTL_VEND] where Code = '" + Idvendedor + "'";
          oRS.DoQuery(sql);
          retorno = Convert.ToString(oRS.Fields.Item("Name").Value);
          Marshal.ReleaseComObject(oRS);
          oRS = null;
          GC.Collect();
          return retorno;
        }
        public static string GetSlpCode(string cardcode)
        {
          string retorno;
          string sql = string.Empty;
          Recordset oRS = null;
          oRS = (SAPbobsCOM.Recordset)vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
          sql = "select slpcode from OCRD where cardcode = '" + cardcode + "'";
          oRS.DoQuery(sql);
          retorno = Convert.ToString(oRS.Fields.Item("slpcode").Value);
          Marshal.ReleaseComObject(oRS);
          oRS = null;
          GC.Collect();
          return retorno;
        }
        public static string GetVendorCode(string cardcode)
        {
          string retorno;
          string sql = string.Empty;
          Recordset oRS = null;
          oRS = (SAPbobsCOM.Recordset)vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
          sql = "select SlpCode from OCRD where CardCode = '"+cardcode+"'";
          oRS.DoQuery(sql);
          retorno = Convert.ToString(oRS.Fields.Item("SlpCode").Value);
          Marshal.ReleaseComObject(oRS);
          oRS = null;
          GC.Collect();
          return retorno;
        }

        public static string GetItemVersion(string item)
        {
          string retorno;
          string sql = string.Empty;
          Recordset oRS = null;
          oRS = (SAPbobsCOM.Recordset)vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
          sql = "select dbo.GetPortalItemCode('"+item+"') ItemCode";
          oRS.DoQuery(sql);
          retorno = Convert.ToString(oRS.Fields.Item("ItemCode").Value);
          Marshal.ReleaseComObject(oRS);
          oRS = null;
          GC.Collect();
          return retorno;
        }

        public static string GetCounty(string cidade, string estado)
        {
          string county;
          string sql = string.Empty;
          Recordset oRS = null;

          oRS = (SAPbobsCOM.Recordset)Utils.vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
          sql = $"select top 1 cast(\"AbsId\" as nvarchar(10)) \"AbsId\" from ocnt where \"Name\" = '{cidade}' and \"State\" = '{estado}'";

          oRS.DoQuery(sql);
          county = oRS.Fields.Item("AbsId").Value.ToString();

          Marshal.ReleaseComObject(oRS);
          oRS = null;
          //GC.Collect();
          return county;
        }

        public static string GetItemName(string itemcode)
        {
            string county;
            string sql = string.Empty;
            Recordset oRS = null;

            oRS = (SAPbobsCOM.Recordset)vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            sql = "select itemname from oitm where itemcode = '"+itemcode+"'";

            oRS.DoQuery(sql);
            county = oRS.Fields.Item("itemname").Value.ToString();

            Marshal.ReleaseComObject(oRS);
            oRS = null;
            //GC.Collect();
            return county;
        }

        public static string GetCondName(string Cond)
        {
            string county;
            string sql = string.Empty;
            Recordset oRS = null;

            oRS = (SAPbobsCOM.Recordset)vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            sql = "select PymntGroup from SBO_BUW_EMCOMEX_PROD_FINAL.dbo.OCTG where GroupNum = "+Cond+"";

            oRS.DoQuery(sql);
            county = oRS.Fields.Item("PymntGroup").Value.ToString();

            Marshal.ReleaseComObject(oRS);
            oRS = null;
            //GC.Collect();
            return county;
        }


        public static string GetTranspName(string transpid)
        {
            string county;
            string sql = string.Empty;
            Recordset oRS = null;

            oRS = (SAPbobsCOM.Recordset)vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            sql = "select Cardname from OCRD where CardCode = '" + transpid + "'";

            oRS.DoQuery(sql);
            county = oRS.Fields.Item("Cardname").Value.ToString();

            Marshal.ReleaseComObject(oRS);
            oRS = null;
            //GC.Collect();
            return county;
        }

        public static bool UpdateStatusProcessar(string code)
        {
          try
          {
            if (vCompany == null || !vCompany.Connected)
              throw new Exception("Sem conexão com o SAP.");

            UserTable table = vCompany.UserTables.Item("ZSINCRONIZATION");

            if (table.GetByKey(code))
            {
            
              table.UserFields.Fields.Item("U_Status").Value = "Processado";
            }


                if (table.Update() != 0)
                    table = null;
              return (false);
          }
          catch (Exception ex)
          {
            Console.WriteLine(ex);
            return (false);
          }

          return (true);
        }


        #endregion

    }
  }