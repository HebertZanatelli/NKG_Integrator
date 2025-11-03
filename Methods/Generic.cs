using SAPbobsCOM;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace IntegratorSales.Methods
{
    class Generic
    {

        public static void InsertRegistryUtil(string table, string key, string action, string searchQuery, string sincCode)
        {

            try
            {
                if (ifExist(table, "Code", key))
                    action = "U";
                else
                    action = "A";

                //MYSQL.DeleteCond(cond);
                string aQuery = Utils.getQuery(searchQuery).Replace("{0}", key);

                string Name = Utils.ReturnValue(aQuery, "Name");
                string Ativo = Utils.ReturnValue(aQuery, "Ativo");

                string query = "";
                if (action == "A")
                {
                    query = $"insert into {table} (Code, Name, Ativo) values ('{key}', '{Name.Replace("'", "")}',  '{Ativo}')";
                }
                if (action == "U")
                {
                    query = $"update {table} set Name = '{Name.Replace("'", "")}', Ativo = '{Ativo.Replace("'", "")}' where Code = '{key}'";
                }
                if (action == "D")
                {
                    query = $"delete from {table} where Code = '{key}'";
                }

                if (MYSQL.ExecInstruction(query))
                    UpdateStatusProcessar(sincCode);
            }
            catch (System.Exception e)
            {
                Console.WriteLine(e.Message.ToString());

            }

            finally
            {

            }


        }

        public static void InsertRegistry(string table, string key, string action, string searchQuery, string sincCode)
        {

            try
            {
                if (ifExist(table, "Code", key))
                    action = "U";
                else
                    action = "A";

                //MYSQL.DeleteCond(cond);
                string aQuery = Utils.getQuery(searchQuery).Replace("{0}", key);

                string Name = Utils.ReturnValue(aQuery, "Name");

                string query = "";
                if (action == "A")
                {
                    query = $"insert into {table} (Code, Name) values ('{key}', '{Name.Replace("'", "")}')";
                }
                if (action == "U")
                {
                    query = $"update {table} set Name = '{Name.Replace("'", "")}' where Code = '{key}'";
                }
                if (action == "D")
                {
                    query = $"delete from {table} where Code = '{key}'";
                }

                if (MYSQL.ExecInstruction(query))
                    UpdateStatusProcessar(sincCode);
            }
            catch (System.Exception e)
            {
                Console.WriteLine(e.Message.ToString());

            }

            finally
            {

            }


        }
        public static void Carga()
        {

            try
            {


                string query = Utils.getQuery("qAll");

                Recordset oRs2 = (Recordset)Utils.vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                oRs2.DoQuery(query);


                while (!oRs2.EoF)
                {
                    string ObjType = oRs2.Fields.Item("ObjType").Value.ToString();
                    string Chave = oRs2.Fields.Item("Chave").Value.ToString();



                    oRs2.MoveNext();
                }


            }
            catch (Exception e)
            {
            }
        }
        public static bool ifExist(string table, string field, string value)
        {
            bool ret = true;
            try
            {


                string query = $"select count(*) Reg from {table} where {field}  = '{value}'";
                int contagem = int.Parse(MYSQL.getvalue(query, "Reg"));
                if (contagem == 0)
                    ret = false;
            }
            catch (Exception e)
            {
                return ret;
            }
            return ret;
        }
        public static void UpdateStatusPedido(string metodo, string key, string action, string searchQuery, string sincCode)
        {

            try
            {
                string query = "";
                if (metodo == "PedidoAprovado")
                {
                    string entry = SAP.GetValue($"SELECT T0.\"U_IdPortal\" FROM ORDR T0 WHERE T0.\"DocEntry\" = {key}", "U_IdPortal");
                    query = $"Update Pedido  set Status  = 'Aprovado', NSAP = {key} where id = {entry}";
                }
                if (metodo == "PedidoCancelado")
                {
                    string entry = SAP.GetValue($"SELECT T0.\"U_IdPortal\" FROM ORDR T0 WHERE T0.\"DocEntry\" = {key}", "U_IdPortal");
                    query = $"Update Pedido  set Status  = 'Cancelado' where id = {entry}";
                }
                if (metodo == "PedidoFaturado")
                {
                    string entry = SAP.GetValue($"SELECT T0.\"U_IdPortal\" FROM OINV T0 WHERE T0.\"DocEntry\" = {key}", "U_IdPortal");
                    query = $"Update Pedido  set Status  = 'Faturado' where id = {entry}";
                }


                if (MYSQL.ExecInstruction(query))
                    UpdateStatusProcessar(sincCode);
            }
            catch (System.Exception e)
            {
                Console.WriteLine(e.Message.ToString());

            }

            finally
            {

            }


        }

        public static void InsertRegistryTabelas(string table, string key, string action, string searchQuery, string sincCode)
        {

            try
            {
                if (ifExist(table, "id", key))
                    action = "U";
                else
                    action = "A";

                //MYSQL.DeleteCond(cond);
                string aQuery = Utils.getQuery(searchQuery).Replace("{0}", key);

                string Name = Utils.ReturnValue(aQuery, "Name");

                string query = "";
                if (action == "A")
                {
                    query = $"insert into {table} (id, nome) values ('{key}', '{Name.Replace("'", "")}')";
                }
                if (action == "U")
                {
                    query = $"update {table} set nome = '{Name.Replace("'", "")}' where id = '{key}'";
                }
                if (action == "D")
                {
                    query = $"delete from {table} where id = '{key}'";
                }

                if (MYSQL.ExecInstruction(query))
                    UpdateStatusProcessar(sincCode);
            }
            catch (System.Exception e)
            {
                Console.WriteLine(e.Message.ToString());

            }

            finally
            {

            }


        }

        public static void InsertRegistryVendedor(string table, string key, string action, string searchQuery, string sincCode)
        {

            try
            {
                //MYSQL.DeleteCond(cond);
                string aQuery = Utils.getQuery(searchQuery).Replace("{0}", key);
                Recordset oRs = Utils.ReturnRow(aQuery);

                if (ifExist("Usuario", "CodeSAP", key))
                    action = "U";
                else
                    action = "A";

                string query = "";
                if (action == "A")
                {
                    query = Utils.getQuery("MySQLVendedorAdd");
                    for (int i = 0; i < oRs.Fields.Count; i++)
                    {
                        string fieldName = oRs.Fields.Item(i).Name.ToString();
                        string fieldValue = oRs.Fields.Item(i).Value.ToString();

                        query = query.Replace($"<{fieldName}:>", $"'{fieldValue.Replace("'", "")}'");
                    }

                }
                if (action == "U")
                {
                    query = Utils.getQuery("MySQLVendedorUpd");
                    for (int i = 0; i < oRs.Fields.Count; i++)
                    {
                        string fieldName = oRs.Fields.Item(i).Name.ToString();
                        string fieldValue = oRs.Fields.Item(i).Value.ToString();

                        query = query.Replace($"<{fieldName}:>", $"'{fieldValue.Replace("'", "")}'");
                    }
                }
                if (action == "D")
                {
                    query = $"delete from Usuario where CodeSAP = '{key}'";
                }

                if (MYSQL.ExecInstruction(query))
                    UpdateStatusProcessar(sincCode);
            }
            catch (System.Exception e)
            {
                Console.WriteLine(e.Message.ToString());

            }

            finally
            {

            }

        }
        public static bool InsereFiliais(string cardcode)
        {
            MYSQL.ExecInstruction($"delete from FilialCliente where fk_cliente = '{cardcode}'");


            Recordset oRs2 = (Recordset)Utils.vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            try
            {
                oRs2.DoQuery($"SELECT *  FROM CRD8 T0 WHERE T0.\"CardCode\" = '{cardcode}'");


                while (!oRs2.EoF)
                {
                    string BPLId = oRs2.Fields.Item("BPLId").Value.ToString();
                    string CardCode = oRs2.Fields.Item("CardCode").Value.ToString();

                    MYSQL.ExecInstruction($"insert into FilialCliente (fk_filial, fk_cliente) values ({BPLId},'{cardcode}')");


                    oRs2.MoveNext();
                }

                return true;
            }
            catch (Exception ex)
            {

                return false;
            }

        }
        public static void InsertRegistryClientes(string table, string key, string action, string searchQuery, string sincCode)
        {

            try
            {
                string cnpj = "";
                //MYSQL.DeleteCond(cond);
                string aQuery = Utils.getQuery(searchQuery).Replace("{0}", key);
                Recordset oRs = Utils.ReturnRow(aQuery);

                if (ifExist("Clientes", "CardCode", key))
                    action = "U";
                else
                    action = "A";

                string query = "";
                if (action == "A")
                {
                    query = Utils.getQuery("MySQLClientesAdd");
                    for (int i = 0; i < oRs.Fields.Count; i++)
                    {
                        string fieldName = oRs.Fields.Item(i).Name.ToString();
                        string fieldValue = oRs.Fields.Item(i).Value.ToString();
                        if (fieldName == "TaxIdNum")
                            cnpj = fieldValue;
                        query = query.Replace($"<{fieldName}:>", $"'{fieldValue.Replace("'", "")}'");
                    }

                }
                if (action == "U")
                {
                    query = Utils.getQuery("MySQLClientesUpd");
                    for (int i = 0; i < oRs.Fields.Count; i++)
                    {
                        string fieldName = oRs.Fields.Item(i).Name.ToString();
                        string fieldValue = oRs.Fields.Item(i).Value.ToString();
                        if (fieldName == "TaxIdNum")
                            cnpj = fieldValue;
                        query = query.Replace($"<{fieldName}:>", $"'{fieldValue.Replace("'", "")}'");
                    }
                }
                if (action == "D")
                {
                    query = $"delete from {table} where CardCode = '{key}'";
                }

                if (MYSQL.ExecInstruction(query))
                    if (InsereFiliais(key))
                        UpdateStatusProcessar(sincCode);

                query = $"delete from Lead where TaxIdNum = '{cnpj}'";
                MYSQL.ExecInstruction(query);
            }
            catch (System.Exception e)
            {
                Console.WriteLine(e.Message.ToString());

            }

            finally
            {

            }

        }
        public static bool callReprovados()
        {
            bool ret = true;
            try
            {
                string query = $"";
             

                Recordset oRs2 = (Recordset)Utils.vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                try
                {
                    oRs2.DoQuery(Utils.getQuery("qReprovados"));
                    while (!oRs2.EoF)
                    {
                        string DocEntry = oRs2.Fields.Item("U_IdPortal").Value.ToString();

                        query = $"Update Pedido  set Status  = 'Reprovado' where id = {DocEntry};";
                        MYSQL.ExecInstruction(query);

                        oRs2.MoveNext();
                    }
                }
                catch (Exception ex)
                {
                    return false;
                }
            }
            catch (Exception e)
            {
                return ret;
            }
            return ret;
        }
        public static bool callInadimplentes()
        {
            bool ret = true;
            try
            {
                string query = $"update Clientes set Inadimplencia = 'Não'";
                MYSQL.ExecInstruction(query);


                Recordset oRs2 = (Recordset)Utils.vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                try
                {
                    oRs2.DoQuery($"SELECT distinct \"CardCode\" FROM INV6 T0  INNER JOIN OINV T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" WHERE T0.\"Status\"  = 'O' and  T0.\"DueDate\" < CURRENT_DATE");
                    while (!oRs2.EoF)
                    {
                        string CardCode = oRs2.Fields.Item("CardCode").Value.ToString();

                        query = $"update Clientes set Inadimplencia = 'Sim' where CardCode = '{CardCode}';";
                        MYSQL.ExecInstruction(query);

                        oRs2.MoveNext();
                    }
                }
                catch (Exception ex)
                {
                    return false;
                }
            }
            catch (Exception e)
            {
                return ret;
            }
            return ret;
        }
        public static bool callPrices(string itemCode)
        {
            bool ret = true;
            try
            {
                string query = $"delete from tabelas_preco where fk_produto  = '{itemCode}'";
                MYSQL.ExecInstruction(query);


                Recordset oRs2 = (Recordset)Utils.vCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                try
                {
                    oRs2.DoQuery($"SELECT *  FROM ITM1 T0 WHERE T0.\"ItemCode\"  = '{itemCode}'");
                    while (!oRs2.EoF)
                    {
                        string PriceList = oRs2.Fields.Item("PriceList").Value.ToString();
                        string price = oRs2.Fields.Item("Price").Value.ToString();

                        query = $"INSERT INTO tabelas_preco (fk_produto, preco, fk_tabela) VALUES('{itemCode}', {price.Replace(",", ".")}, {PriceList});";
                        MYSQL.ExecInstruction(query);

                        oRs2.MoveNext();
                    }
                }
                catch (Exception ex)
                {
                    return false;
                }
            }
            catch (Exception e)
            {
                return ret;
            }
            return ret;
        }
        public static void InsertRegistryItems(string table, string key, string action, string searchQuery, string sincCode)
        {

            try
            {
                //MYSQL.DeleteCond(cond);
                string aQuery = Utils.getQuery(searchQuery).Replace("{0}", key);
                Recordset oRs = Utils.ReturnRow(aQuery);

                if (ifExist("Itens", "ItemCode", key))
                    action = "U";
                else
                    action = "A";

                string query = "";

                if (action == "A")
                {
                    query = Utils.getQuery("MySQLItemAdd");
                    for (int i = 0; i < oRs.Fields.Count; i++)
                    {
                        string fieldName = oRs.Fields.Item(i).Name.ToString();
                        string fieldValue = oRs.Fields.Item(i).Value.ToString();

                        query = query.Replace($"<{fieldName}:>", $"'{fieldValue.Replace("'", "").Replace(",", ".")}'");
                    }

                }
                if (action == "U")
                {
                    query = Utils.getQuery("MySQLItemUpd");
                    for (int i = 0; i < oRs.Fields.Count; i++)
                    {
                        string fieldName = oRs.Fields.Item(i).Name.ToString();
                        string fieldValue = oRs.Fields.Item(i).Value.ToString();
                        if (fieldName == "OnHand")
                            fieldValue = fieldValue.Replace(",", ".");
                        query = query.Replace($"<{fieldName}:>", $"'{fieldValue.Replace("'", "")}'");
                    }
                }
                if (action == "D")
                {
                    query = $"delete from {table} where ItemCode = '{key}'";
                }

                if (MYSQL.ExecInstruction(query))
                    if (callPrices(key))  //Voltar isso depois
                        UpdateStatusProcessar(sincCode);
            }
            catch (System.Exception e)
            {
                Console.WriteLine(e.Message.ToString());

            }

            finally
            {

            }

        }
        public static void InsertRegistryContato(string table, string key, string action, string searchQuery, string sincCode)
        {

            try
            {
                //MYSQL.DeleteCond(cond);
                string aQuery = Utils.getQuery(searchQuery).Replace("{0}", key);
                Recordset oRs = Utils.ReturnRow(aQuery);


                string query = "";
                if (action == "A")
                {
                    query = Utils.getQuery("MySQLContatoAdd");
                    for (int i = 0; i < oRs.Fields.Count; i++)
                    {
                        string fieldName = oRs.Fields.Item(i).Name.ToString();
                        string fieldValue = oRs.Fields.Item(i).Value.ToString();

                        query = query.Replace($"<{fieldName}:>", $"'{fieldValue.Replace("'", "")}'");
                    }

                }
                if (action == "U")
                {
                    query = Utils.getQuery("MySQLContatoUpd");
                    for (int i = 0; i < oRs.Fields.Count; i++)
                    {
                        string fieldName = oRs.Fields.Item(i).Name.ToString();
                        string fieldValue = oRs.Fields.Item(i).Value.ToString();

                        query = query.Replace($"<{fieldName}:>", $"'{fieldValue.Replace("'", "")}'");
                    }
                }
                if (action == "D")
                {
                    query = $"delete from {table} where SAP_Id = '{key}'";
                }

                if (MYSQL.ExecInstruction(query))
                    UpdateStatusProcessar(sincCode);
            }
            catch (System.Exception e)
            {
                Console.WriteLine(e.Message.ToString());

            }

            finally
            {

            }

        }

        public static void UpdateStatusProcessar(string code)
        {
            try
            {
                if (Utils.vCompany == null || !Utils.vCompany.Connected)
                    throw new Exception("Sem conexão com o SAP.");

                UserTable table = Utils.vCompany.UserTables.Item("SL_SINC");

                if (table.GetByKey(code))
                {

                    table.UserFields.Fields.Item("U_Status").Value = "C";
                }


                if (table.Update() != 0)
                    table = null;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);

            }

        }


    }
}
