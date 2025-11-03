using System.Data;
using MySql.Data.MySqlClient;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;
using IntegratorSales.Methods;
using System.Collections;
using System.IO;

namespace IntegratorSales
{
    class Program
    {
        public static void CancelDoc(int entry)
        {
            Documents oDoc = null;
            oDoc = (Documents)Utils.vCompany.GetBusinessObject(BoObjectTypes.oPurchaseInvoices);

            oDoc.GetByKey(entry);
            oDoc.CreateCancellationDocument();
        }

        public static void cancelList()
        {
            CancelDoc(725729);
            CancelDoc(725730);
            CancelDoc(725737);
            CancelDoc(725745);
            CancelDoc(725746);
            CancelDoc(725747);
            CancelDoc(725748);
            CancelDoc(725749);
            CancelDoc(725750);
            CancelDoc(725752);
            CancelDoc(725751);
            CancelDoc(725753);
            CancelDoc(725754);
            CancelDoc(725755);
            CancelDoc(725757);
            CancelDoc(725756);
            CancelDoc(725758);
            CancelDoc(725759);
            CancelDoc(725761);
            CancelDoc(725762);
            CancelDoc(725763);
            CancelDoc(725764);
            CancelDoc(725765);
            CancelDoc(725767);
            CancelDoc(725768);
            CancelDoc(725766);
            CancelDoc(725769);
            CancelDoc(725770);
            CancelDoc(725771);
            CancelDoc(725772);
            CancelDoc(725773);
            CancelDoc(725774);
            CancelDoc(725775);
            CancelDoc(725776);
            CancelDoc(725777);
            CancelDoc(725778);
            CancelDoc(725779);
            CancelDoc(725780);
            CancelDoc(725782);
            CancelDoc(725783);
            CancelDoc(725781);
            CancelDoc(725784);
            CancelDoc(725786);
            CancelDoc(725787);
            CancelDoc(725789);
            CancelDoc(725790);
            CancelDoc(725791);
            CancelDoc(725792);
            CancelDoc(725795);
            CancelDoc(725794);
            CancelDoc(725796);
            CancelDoc(725797);
            CancelDoc(725798);
            CancelDoc(725799);
            CancelDoc(725800);
            CancelDoc(725801);
            CancelDoc(725802);
            CancelDoc(725804);
            CancelDoc(725805);
            CancelDoc(725806);
            CancelDoc(725807);
            CancelDoc(725810);
            CancelDoc(725811);
            CancelDoc(725812);
            CancelDoc(725808);
            CancelDoc(725813);
            CancelDoc(725814);
            CancelDoc(725815);
            CancelDoc(725816);
            CancelDoc(725817);
            CancelDoc(725818);
            CancelDoc(725819);
            CancelDoc(725820);
            CancelDoc(725822);
            CancelDoc(725823);
            CancelDoc(725824);
            CancelDoc(725825);
            CancelDoc(725826);
            CancelDoc(725827);
            CancelDoc(725829);
            CancelDoc(725830);
            CancelDoc(725831);
            CancelDoc(725832);
            CancelDoc(725835);
            CancelDoc(725836);
            CancelDoc(725837);
            CancelDoc(725838);
            CancelDoc(725840);
            CancelDoc(725842);
            CancelDoc(725833);
            CancelDoc(725834);
            CancelDoc(725839);
            CancelDoc(725841);
            CancelDoc(725843);
            CancelDoc(725844);
            CancelDoc(725845);
            CancelDoc(725847);
            CancelDoc(725848);
            CancelDoc(725849);
            CancelDoc(725850);
            CancelDoc(725852);
            CancelDoc(725851);
            CancelDoc(725853);
            CancelDoc(725855);
            CancelDoc(725856);
            CancelDoc(725857);
            CancelDoc(725858);
            CancelDoc(725859);
            CancelDoc(725860);
            CancelDoc(725862);
            CancelDoc(725861);
            CancelDoc(725863);
            CancelDoc(725865);
            CancelDoc(725864);
            CancelDoc(725867);
            CancelDoc(725868);
            CancelDoc(725866);
            CancelDoc(725869);
            CancelDoc(725870);
            CancelDoc(725871);
            CancelDoc(725872);
            CancelDoc(725873);
            CancelDoc(725874);
            CancelDoc(725875);
            CancelDoc(725876);
            CancelDoc(725877);
            CancelDoc(725878);
            CancelDoc(725879);
            CancelDoc(725880);
            CancelDoc(725881);
            CancelDoc(725882);
            CancelDoc(725883);
            CancelDoc(725884);
            CancelDoc(725885);
            CancelDoc(725886);
            CancelDoc(725888);
            CancelDoc(725889);
            CancelDoc(725890);
            CancelDoc(725891);
            CancelDoc(725892);
            CancelDoc(725893);
            CancelDoc(725896);
            CancelDoc(725887);
            CancelDoc(725895);
            CancelDoc(725897);
            CancelDoc(725898);
            CancelDoc(725899);
            CancelDoc(725900);
            CancelDoc(725901);
            CancelDoc(725902);
            CancelDoc(725904);
            CancelDoc(725903);
            CancelDoc(725905);
            CancelDoc(725906);
            CancelDoc(725907);
            CancelDoc(725909);
            CancelDoc(725910);
            CancelDoc(725911);
            CancelDoc(725912);
            CancelDoc(725913);
            CancelDoc(725914);
            CancelDoc(725915);
            CancelDoc(725916);
            CancelDoc(725917);
            CancelDoc(725918);
            CancelDoc(725919);
            CancelDoc(725920);
            CancelDoc(725922);
            CancelDoc(725921);
            CancelDoc(725923);
            CancelDoc(725924);
            CancelDoc(725925);
        }

        static void Main(string[] args)
        {

            if (System.Diagnostics.Process.GetProcessesByName(System.IO.Path.GetFileNameWithoutExtension(System.Reflection.Assembly.GetEntryAssembly().Location)).Count() > 1)
            {
                Console.WriteLine("A aplicação já está em execução, impossível continuar!");
                Environment.Exit(0);
            }


            try
            {
                if (Utils.ConnectSAP())
                {
                    cancelList();
                }
            }
            catch (Exception ex)
            { }

        }

        public static void cargaLista()
        {


            string filePath = @"C:\cargaportal.txt";

            try
            {
                // Abre o arquivo para leitura usando StreamReader
                using (StreamReader sr = new StreamReader(filePath))
                {
                    string linha;
                    // Lê cada linha do arquivo e imprime no console
                    while ((linha = sr.ReadLine()) != null)
                    {
                        Console.WriteLine(linha);
                        MYSQL.ExecInstruction(linha);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Ocorreu um erro ao ler o arquivo: " + e.Message);
            }
        }
        public static void RunInsert()
        {
            try
            {
                Recordset oRs = (Recordset)Utils.vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

                string query = "";
                query = Utils.getQuery("qSinc");
                oRs.DoQuery(query);

                while (!oRs.EoF)
                {

                    string Code = oRs.Fields.Item("Name").Value.ToString().Trim();
                    string Name = oRs.Fields.Item("Code").Value.ToString().Trim();
                    string U_Key = oRs.Fields.Item("U_Key").Value.ToString().Trim();
                    string U_Object = oRs.Fields.Item("U_Object").Value.ToString().Trim();
                    string U_Action = oRs.Fields.Item("U_Action").Value.ToString().Trim();
                    Console.WriteLine($"Processando Objeto:{U_Object} Acao:{U_Action} Chave:{U_Key}");
                    if (U_Object == "260")
                        Methods.Generic.InsertRegistryUtil("Utilizacao", U_Key, U_Action, "OUSG", Code);
                    if (U_Object == "40")
                        Methods.Generic.InsertRegistry("CondPagto", U_Key, U_Action, "OCTG", Code);
                    if (U_Object == "147")
                        Methods.Generic.InsertRegistry("FormaPagto", U_Key, U_Action, "OPYM", Code);
                    if (U_Object == "Transp")
                        Methods.Generic.InsertRegistry("Shipper", U_Key, U_Action, "Transp", Code);
                    if (U_Object == "247")
                        Methods.Generic.InsertRegistry("Filial", U_Key, U_Action, "OBPL", Code);
                    if (U_Object == "6")
                        Methods.Generic.InsertRegistryTabelas("tabelas", U_Key, U_Action, "OPLN", Code);

                    if (U_Object == "2")
                        Methods.Generic.InsertRegistryClientes("Clientes", U_Key, U_Action, "Clientes", Code);
                    if (U_Object == "53")
                        Methods.Generic.InsertRegistryVendedor("Vendedor", U_Key, U_Action, "Vendedor", Code);

                    if (U_Object == "4")
                        Methods.Generic.InsertRegistryItems("Itens", U_Key, U_Action, "Itens", Code);

                    if (U_Object == "11")
                        Methods.Generic.InsertRegistryContato("Contato", U_Key, U_Action, "Contato", Code);

                    if (U_Object == "PedidoAprovado")
                        Methods.Generic.UpdateStatusPedido("PedidoAprovado", U_Key, U_Action, "OPLN", Code);

                    if (U_Object == "PedidoCancelado")
                        Methods.Generic.UpdateStatusPedido("PedidoCancelado", U_Key, U_Action, "OPLN", Code);

                    if (U_Object == "PedidoFaturado")
                        Methods.Generic.UpdateStatusPedido("PedidoFaturado", U_Key, U_Action, "OPLN", Code);


                    oRs.MoveNext();
                }

            }
            catch (Exception ex)
            {
                string erro = ex.ToString();
            }

        }
    }
}
