using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;


namespace IntegratorSales
{
    class Utils
    {
        public static SAPbobsCOM.Company vCompany = null;


        public static bool ConnectSAP()
        {
            try
            {
                vCompany = new SAPbobsCOM.Company();

                vCompany.language = BoSuppLangs.ln_Portuguese_Br;

                vCompany.CompanyDB = Properties.Settings.Default.SAP_CompanyDB;
                vCompany.UserName = Properties.Settings.Default.SAP_UserName;
                vCompany.Password = Properties.Settings.Default.SAP_Password;
                vCompany.Server = Properties.Settings.Default.SAP_Server;
                vCompany.LicenseServer = Properties.Settings.Default.SAP_LicenseServer;



                switch (Properties.Settings.Default.SAP_DbServerType)
                {


                    case "MSSQL - 2017":
                        vCompany.DbServerType = BoDataServerTypes.dst_MSSQL2017;
                        break;

                    case "HANA":
                        vCompany.DbServerType = BoDataServerTypes.dst_HANADB;
                        break;

                    default:
                        vCompany.DbServerType = BoDataServerTypes.dst_MSSQL;

                        break;
                }

                Console.WriteLine("tentando conectar a " + vCompany.CompanyDB);
                if (vCompany.Connect() != 0)
                {
                    string last_err_msg = vCompany.GetLastErrorDescription();
                    Console.WriteLine(last_err_msg);
                    throw new Exception(last_err_msg);
                }
                if (vCompany.Connected)
                {
                    Console.WriteLine($"SAP Conectado: {vCompany.CompanyDB}");
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("Falha na conexão com o SAP." + ex.Message.ToString(), ex.Message);
                return (false);
            }

            return (true);
        }


        public static System.Data.DataTable SapDataTableToDotNetDataTable(string xmlcontent)
        {
            var DT = new System.Data.DataTable();
            //var XMLstream = new System.IO.FileStream(pathToXmlFile, FileMode.Open);
            xmlcontent = xmlcontent.Replace((char)(0x1E), ' ');
            var XDoc = System.Xml.Linq.XDocument.Parse(xmlcontent);
            var Columns = XDoc.Element("Matrix").Element("ColumnsInfo").Elements("ColumnInfo");
            foreach (var Column in Columns)
            {
                DT.Columns.Add(Column.Element("UniqueID").Value);
            }
            var Rows = XDoc.Element("Matrix").Element("Rows").Elements("Row");
            var Names = new List<string>();
            foreach (var Row in Rows)
            {
                var DTRow = DT.NewRow();
                var Cells = Row.Element("Columns").Elements("Column");
                foreach (var Cell in Cells)
                {
                    var ColName = Cell.Element("ID").Value;
                    var ColValue = Cell.Element("Value").Value;
                    DTRow[ColName] = ColValue;
                }
                DT.Rows.Add(DTRow);
            }
            return DT;
        }

        public static void SetData()
        {
            CreateTable("SL_SINC", "[SLF] Sinc", BoUTBTableType.bott_NoObject);
            CreateField("@SL_SINC", "Object", "Objeto", BoFieldTypes.db_Alpha, 254);
            CreateField("@SL_SINC", "Key", "Chave", BoFieldTypes.db_Alpha, 254);
            CreateField("@SL_SINC", "Error", "Erro", BoFieldTypes.db_Alpha, 254);
            CreateField("@SL_SINC", "Action", "Acao", BoFieldTypes.db_Alpha, 2);
            CreateField("@SL_SINC", "Status", "Status", BoFieldTypes.db_Alpha, 1, new string[,] { { "O", "Pendente" }, { "C", "Recebido" } });
            CreateField("OITM", "Bloqueado", "Bloqueado Portal", BoFieldTypes.db_Alpha, 1, new string[,] { { "N", "Não" }, { "S", "Sim" } });
            CreateField("OCRD", "Antecipado", "Antecipado Portal", BoFieldTypes.db_Alpha, 1, new string[,] { { "N", "Não" }, { "S", "Sim" } });
            CreateField("ORDR", "IdPortal", "Id Portal", BoFieldTypes.db_Alpha, 254);

            CreateField("OSLP", "Senha", "Senha", BoFieldTypes.db_Alpha, 254);

            CreateField("OITB", "Portal_Integra", "Integração Portal (S/N)", BoFieldTypes.db_Alpha, 1, new string[,] { { "N", "Não" }, { "S", "Sim" } });
            CreateField("OUSG", "Portal_Integra", "Integração Portal (S/N)", BoFieldTypes.db_Alpha, 1, new string[,] { { "Y", "Sim" }, { "N", "Não" } });


        }
        public static void CreateTable(string table, string name, BoUTBTableType type)
        {
            Console.WriteLine("Criando Tabela " + name);
            //Program.oApplication.StatusBar.SetText("Criando Tabela " + name, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
            SAPbobsCOM.UserTablesMD outb;
            outb = (UserTablesMD)vCompany.GetBusinessObject(BoObjectTypes.oUserTables);

            outb.TableName = table;
            outb.TableDescription = name;
            outb.TableType = type;

            outb.Add();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(outb);
        }
        public static void CreateField(string table, string name, string desc, BoFieldTypes Type, int editsize)
        {
            Console.WriteLine("Criando Campo " + name + " :: Tabela " + table);
            //MainClass.oApplication.StatusBar.SetText("Criando Campo " + name + " :: Tabela " + table, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
            SAPbobsCOM.UserFieldsMD oufd;
            oufd = (UserFieldsMD)vCompany.GetBusinessObject(BoObjectTypes.oUserFields);

            oufd.TableName = table;
            oufd.Name = name;
            oufd.Description = desc;
            oufd.Type = Type;
            if (Type == BoFieldTypes.db_Alpha)
                oufd.EditSize = editsize;
            if (Type == BoFieldTypes.db_Float)
                oufd.SubType = BoFldSubTypes.st_Quantity;

            oufd.Add();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oufd);

        }
        public static void CreateField(string table, string name, string desc, BoFieldTypes Type, int editsize, string[,] validvalues)
        {
            Console.WriteLine("Criando Campo " + name + " :: Tabela " + table);
            //MainClass.oApplication.StatusBar.SetText("Criando Campo " + name + " :: Tabela " + table, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
            SAPbobsCOM.UserFieldsMD oufd;
            oufd = (UserFieldsMD)vCompany.GetBusinessObject(BoObjectTypes.oUserFields);

            oufd.TableName = table;
            oufd.Name = name;
            oufd.Description = desc;
            oufd.Type = Type;
            if (Type == BoFieldTypes.db_Alpha)
                oufd.Size = editsize;
            if (Type == BoFieldTypes.db_Float)
                oufd.SubType = BoFldSubTypes.st_Quantity;

            for (int i = 0; i < validvalues.GetLength(0); i++)
            {
                oufd.ValidValues.Value = validvalues[i, 0];
                oufd.ValidValues.Description = validvalues[i, 1];
                oufd.ValidValues.Add();
                string valor = validvalues[i, 0];
                string valor1 = validvalues[i, 1];
            }



            oufd.DefaultValue = validvalues[0, 0];
            oufd.Add();

            string erro = vCompany.GetLastErrorDescription();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oufd);

        }


        public static string getQuery(string qname)
        {
            try
            {
                System.Reflection.Assembly oAssembly = System.Reflection.Assembly.GetEntryAssembly();
                System.IO.StreamReader oStrRdr = null;
                try
                {
                    string xmlStr = string.Empty;
                    oStrRdr = new System.IO.StreamReader(oAssembly.GetManifestResourceStream("IntegratorSales.Queries.Queries.xml"));
                    xmlStr = oStrRdr.ReadToEnd();

                    XmlDocument doc = new XmlDocument();
                    doc.LoadXml(xmlStr);
                    string query = doc.GetElementsByTagName(qname)[0].InnerText;
                    return query;

                }
                catch (Exception er)
                {
                    string msg = er.Message.ToString();
                    return msg;
                }
            }
            catch (Exception er)
            {
                string msg = er.Message.ToString();
                return msg;
            }
        }

        public static void Log(string sMensagem)
        {

            string sCaminhoApp = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase);
            string sCaminhoTexto = System.IO.Path.Combine(sCaminhoApp, "Log.txt");

            System.IO.StreamWriter oStr;
            //  se tiver : significa que tem letra de unidade, caso contrario precisa acrescentar o //
            if (sCaminhoTexto.Replace("file:\\", "").Substring(1, 1) == ":")
                oStr = new System.IO.StreamWriter(sCaminhoTexto.Replace("file:\\", ""), true);
            else
                oStr = new System.IO.StreamWriter(@"\\" + sCaminhoTexto.Replace("file:\\", ""), true);

            oStr.WriteLine(DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + " - " + sMensagem);
            oStr.Close();

            oStr = null;



        }

        public static void AddRecord(string table, string code)
        {
            try
            {

                UserTable oUst = (UserTable)vCompany.UserTables.Item(table);
                oUst.Code = code;
                oUst.Name = code;

                oUst.Add();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUst);
                oUst = null;
                GC.Collect();
            }
            catch { }
        }

        public static void DeleteRecord(string table, string code)
        {
            try
            {
                UserTable oUst = (UserTable)vCompany.UserTables.Item(table);

                if (oUst.GetByKey(code))
                {
                    oUst.Remove();

                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUst);
                oUst = null;
                GC.Collect();
            }
            catch (Exception ex)
            {
            }

        }
        public static string ReturnValue(string query, string field)
        {
            string value = "";
            try
            {

                Recordset oRs2 = (Recordset)Utils.vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                oRs2.DoQuery(query);

                if (oRs2.RecordCount > 0)
                {
                    while (!oRs2.EoF)
                    {
                        value = oRs2.Fields.Item(field).Value.ToString();

                        oRs2.MoveNext();
                    }
                }


                return value;
            }
            catch (Exception ex)
            {

                return "";
            }

        }
        public static Recordset ReturnRow(string query)
        {
            string value = "";
            Recordset oRs2 = (Recordset)Utils.vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            try
            {


                oRs2.DoQuery(query);

                if (oRs2.RecordCount > 0)
                {
                    return oRs2;
                    //while (!oRs2.EoF)
                    //{
                    //  value = oRs2.Fields.Item(field).Value.ToString();

                    //  oRs2.MoveNext();
                    //}
                }


                return oRs2;
            }
            catch (Exception ex)
            {

                return oRs2;
            }

        }
    }
}

