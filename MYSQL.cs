using System.Data;
using MySql.Data.MySqlClient;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IntegratorSales
{
    class MYSQL
    {
        private static MySqlConnection mConn;
        private static MySqlDataAdapter mAdapter;
        private static DataSet mDataSet;
        private static MySqlCommand command;
        private static MySqlDataReader mdr;

        public static string MYSQL_uid;
        public static string MYSQL_database;
        public static string MYSQL_pwd;
        public static string MYSQL_Server;

        public static bool ConnectDB()
        {
            MYSQL_uid = Properties.Settings.Default.MYSQL_uid;
            MYSQL_database = Properties.Settings.Default.MYSQL_database;
            MYSQL_pwd = Properties.Settings.Default.MYSQL_pwd;
            MYSQL_Server = Properties.Settings.Default.MYSQL_Server;

            string ConnStr = "Persist Security Info=False;server=" + MYSQL_Server + ";database=" + MYSQL_pwd + ";uid=" + MYSQL_uid + ";server=" + MYSQL_Server + ";database=" + MYSQL_database + ";uid=" + MYSQL_uid + ";pwd=" + MYSQL_pwd + "";
            mDataSet = new DataSet();
            //define string de conexao e cria a conexao
            mConn = new MySqlConnection(ConnStr);
            try
            {
                //abre a conexao
                mConn.Open();
            }
            catch (System.Exception e)
            {
                Console.WriteLine(e.Message.ToString());
                return false;
            }
            return true;
        }
        public static bool DisconnectDB()
        {
            try
            {
                if (mConn.State == ConnectionState.Open)
                {
                    mConn.Close();
                }
            }
            catch (System.Exception e)
            {
                Console.WriteLine(e.Message.ToString());
                return false;
            }
            return true;
        }


        public static DataTable ReturnMatrix(string query)
        {
            MYSQL.ConnectDB();
            DataTable mDataTable = new System.Data.DataTable("Matrix");
            //verificva se a conexão esta aberta
            if (mConn.State == ConnectionState.Open)
            {
                //cria um adapter usando a instrução SQL para acessar a tabela Clientes
                mAdapter = new MySqlDataAdapter(query, mConn);
                //preenche o dataset via adapter
                mAdapter.Fill(mDataSet, "Matrix");
                mDataTable = mDataSet.Tables["Matrix"];
            }
            DisconnectDB();
            return mDataTable;
        }

        public static DataTable MatrizCliente()
        {
            MYSQL.ConnectDB();
            DataTable mDataTable = new System.Data.DataTable("MatrizCliente");
            //verificva se a conexão esta aberta
            if (mConn.State == ConnectionState.Open)
            {
                //cria um adapter usando a instrução SQL para acessar a tabela Clientes
                mAdapter = new MySqlDataAdapter("SELECT absid FROM cliente where status = 'Transmitindo'", mConn);
                //preenche o dataset via adapter
                mAdapter.Fill(mDataSet, "MatrizCliente");
                mDataTable = mDataSet.Tables["MatrizCliente"];
            }
            DisconnectDB();
            return mDataTable;
        }

        public static DataTable MatrizLead()
        {
            MYSQL.ConnectDB();
            DataTable mDataTable = new System.Data.DataTable("MatrizCliente");
            //verificva se a conexão esta aberta
            if (mConn.State == ConnectionState.Open)
            {
                //cria um adapter usando a instrução SQL para acessar a tabela Clientes
                mAdapter = new MySqlDataAdapter("SELECT id FROM Lead where ifnull(status,'') = ''", mConn);
                //preenche o dataset via adapter
                mAdapter.Fill(mDataSet, "MatrizCliente");
                mDataTable = mDataSet.Tables["MatrizCliente"];
            }
            DisconnectDB();
            return mDataTable;
        }

        public static DataTable MatrizPedido()
        {
            MYSQL.ConnectDB();
            DataTable mDataTable = new System.Data.DataTable("MatrizPedido");
            //verificva se a conexão esta aberta
            if (mConn.State == ConnectionState.Open)
            {
                //cria um adapter usando a instrução SQL para acessar a tabela Clientes
                mAdapter = new MySqlDataAdapter("select id from brb.Pedido where status in ( 'Transmitindo', 'Erro') ", mConn);
                //preenche o dataset via adapter
                mAdapter.Fill(mDataSet, "MatrizPedido");
                mDataTable = mDataSet.Tables["MatrizPedido"];
            }
            DisconnectDB();
            return mDataTable;
        }

        public static DataTable StockProdutos(string Filial)
        {
            MYSQL.ConnectDB();
            DataTable mDataTable = new System.Data.DataTable("StockProdutos");
            //verificva se a conexão esta aberta
            if (mConn.State == ConnectionState.Open)
            {
                //cria um adapter usando a instrução SQL para acessar a tabela Clientes
                mAdapter = new MySqlDataAdapter("select distinct  id,  CAST(ifnull(stk_" + Filial + ",0) AS char(8)) qtd from portalbuw1.produto", mConn);
                //preenche o dataset via adapter
                mAdapter.Fill(mDataSet, "StockProdutos");
                mDataTable = mDataSet.Tables["StockProdutos"];
            }
            DisconnectDB();
            return mDataTable;
        }

        public static DataTable LCred()
        {
            MYSQL.ConnectDB();
            DataTable mDataTable = new System.Data.DataTable("LCred");
            //verificva se a conexão esta aberta
            if (mConn.State == ConnectionState.Open)
            {
                //cria um adapter usando a instrução SQL para acessar a tabela Clientes
                mAdapter = new MySqlDataAdapter("select codcliente, ifnull(lcompro,0) lcompro, ifnull(lcred,0) lcred, ifnull(ldisp,0) ldisp, ifnull(inadi,0) inadi from cliente", mConn);
                //preenche o dataset via adapter
                mAdapter.Fill(mDataSet, "LCred");
                mDataTable = mDataSet.Tables["LCred"];
            }
            DisconnectDB();
            return mDataTable;
        }

        public static DataTable StatusPed()
        {
            MYSQL.ConnectDB();
            DataTable mDataTable = new System.Data.DataTable("StatusPed");
            //verificva se a conexão esta aberta
            if (mConn.State == ConnectionState.Open)
            {
                //cria um adapter usando a instrução SQL para acessar a tabela Clientes
                mAdapter = new MySqlDataAdapter("select CONVERT(id_pedido, CHAR(50)) id_pedido, ifnull(status,'') status, ifnull(motivo1,'') motivo1, ifnull(motivo2,'') motivo2, ifnull(motivo3,'') motivo3, ifnull(motivo4,'') motivo4  from portalbuw1.pedido", mConn);
                //preenche o dataset via adapter
                mAdapter.Fill(mDataSet, "StatusPed");
                mDataTable = mDataSet.Tables["StatusPed"];
            }
            DisconnectDB();
            return mDataTable;
        }

        public static DataTable UltCompraM()
        {
            MYSQL.ConnectDB();
            DataTable mDataTable = new System.Data.DataTable("UltCompraM");
            //verificva se a conexão esta aberta
            if (mConn.State == ConnectionState.Open)
            {
                //cria um adapter usando a instrução SQL para acessar a tabela Clientes
                mAdapter = new MySqlDataAdapter("select CONVERT(codcliente, CHAR(50)) codcliente, CONVERT(ultcomp, CHAR(50)) ultcomp from portalbuw1.cliente", mConn);
                //preenche o dataset via adapter
                mAdapter.Fill(mDataSet, "UltCompraM");
                mDataTable = mDataSet.Tables["UltCompraM"];
            }
            DisconnectDB();
            return mDataTable;
        }

        public static DataTable StockProdutosPedido(string Filial)
        {
            MYSQL.ConnectDB();
            DataTable mDataTable = new System.Data.DataTable("StockProdutosPedido");
            //verificva se a conexão esta aberta
            if (mConn.State == ConnectionState.Open)
            {
                //cria um adapter usando a instrução SQL para acessar a tabela Clientes
                mAdapter = new MySqlDataAdapter("select  t1.id_item, Cast(sum(t1.qtd) as char(8)) qtd from portalbuw1.pedido t0 inner join portalbuw1.itenspedido t1 on t0.id_pedido = t1.fk_pedido inner join portalbuw1.cliente t2 on t0.fk_cliente = t2.id where t0.status in ('Transmitindo', 'Em Aprovação') and t2.fk_filial = " + Filial + " group by t1.id_item", mConn);
                //preenche o dataset via adapter
                mAdapter.Fill(mDataSet, "StockProdutosPedido");
                mDataTable = mDataSet.Tables["StockProdutosPedido"];
            }
            DisconnectDB();
            return mDataTable;
        }

        public static DataTable Cliente(int absid)
        {
            MYSQL.ConnectDB();
            DataTable mDataTable = new System.Data.DataTable("Clientes");
            //verificva se a conexão esta aberta
            if (mConn.State == ConnectionState.Open)
            {
                string query = "SELECT  absid,  razao, cnpj, ie, emailxml, ddd, telefone, ent_tipo, ent_end, ent_num, ent_comp, ent_bairro, ent_cep, ent_cidade, ent_estado, cob_tipo, cob_end, cob_num, cob_comp, cob_bairro, cob_cep, cob_cidade, cob_estado, com_nome, com_sobrenome, com_setor, com_cargo, com_ddd, com_tel, com_cel, com_fax, ifnull(com_nascimento,'1900-01-01') com_nascimento, com_obs, fin_nome, fin_sobrenome, fin_setor, fin_cargo, fin_ddd, fin_tel, fin_cel, fin_fax, ifnull(fin_nascimento,'1900-01-01') fin_nascimento, fin_obs, mkt_nome, mkt_sobrenome, mkt_setor, mkt_cargo, mkt_ddd, mkt_tel, mkt_cel, mkt_fax, ifnull(mkt_nascimento,'1900-01-01') mkt_nascimento, mkt_obs, log_nome, log_sobrenome, log_setor, log_cargo, log_ddd, log_tel, log_cel, log_fax, ifnull(log_nascimento,'1900-01-01') log_nascimento, log_obs, userid, status, keysap, RetornoSAP FROM cliente where absid = " + absid + "";
                //cria um adapter usando a instrução SQL para acessar a tabela Clientes
                mAdapter = new MySqlDataAdapter(query, mConn);
                //preenche o dataset via adapter
                mAdapter.Fill(mDataSet, "Clientes");
                mDataTable = mDataSet.Tables["Clientes"];
            }
            DisconnectDB();
            return mDataTable;
        }

        public static DataTable Lead(int absid)
        {
            MYSQL.ConnectDB();
            DataTable mDataTable = new System.Data.DataTable("Clientes");
            //verificva se a conexão esta aberta
            if (mConn.State == ConnectionState.Open)
            {
                string query = "SELECT id,TaxIdNum,CardName,CardFName,TaxIdNum2,Contato,Phone,Email,Cep,Endereco,Numero,Complemento,Bairro,Cidade,Estado,Autor,status,CardSAP FROM Lead where id = " + absid + "";
                //cria um adapter usando a instrução SQL para acessar a tabela Clientes
                mAdapter = new MySqlDataAdapter(query, mConn);
                //preenche o dataset via adapter
                mAdapter.Fill(mDataSet, "Clientes");
                mDataTable = mDataSet.Tables["Clientes"];
            }
            DisconnectDB();
            return mDataTable;
        }

        public static DataTable Pedido(int absid)
        {
            MYSQL.ConnectDB();
            DataTable mDataTable = new System.Data.DataTable("Pedido");
            //verificva se a conexão esta aberta
            if (mConn.State == ConnectionState.Open)
            {
                string query = $"select id,DocEntry,Status,DATE_FORMAT(DocDate, '%Y-%m-%d') AS DocDate,DATE_FORMAT(ShipDate, '%Y-%m-%d') AS ShipDate,CardCode,Contact,GroupNum,PymCode,Incoterms,Shipper,(select u.CodeSAP from Usuario u where u.id = SlpPerson) SlpPerson,BPLId,NatOper,Freight,DiscountTotal,DocTotal,Remarks,itensesboco,NSAP,autor,duplicado from brb.Pedido where DocEntry = {absid}";
                //cria um adapter usando a instrução SQL para acessar a tabela Clientes
                mAdapter = new MySqlDataAdapter(query, mConn);
                //preenche o dataset via adapter
                mAdapter.Fill(mDataSet, "Pedido");
                mDataTable = mDataSet.Tables["Pedido"];
            }
            DisconnectDB();
            return mDataTable;
        }

        public static string Getqtdpedido(string code, string filial)
        {
            MYSQL.ConnectDB();

            string retorno = "0";
            string sql = string.Empty;
            sql = "SELECT Sum(qtd) qtd FROM portalbuw1.pedido t0 INNER JOIN portalbuw1.itenspedido t1 ON t1.fk_pedido = t0.id_pedido INNER JOIN portalbuw1.cliente t2 ON t0.fk_cliente = t2.id WHERE t0.status IN( 'Em Aprovação', 'Transmitindo' ) AND fk_filial = " + filial + " AND t1.id_item = '" + code + "' GROUP BY t1.id_item, t2.fk_filial";
            command = new MySqlCommand(sql, mConn);
            mdr = command.ExecuteReader();
            if (mdr.Read())
            {
                retorno = mdr.GetString("qtd");


            }
            DisconnectDB();
            return retorno;
        }
        public static string Getid()
        {
            MYSQL.ConnectDB();

            string retorno = "0";
            string sql = string.Empty;
            sql = "select Convert(Max(id)+1, char(50)) id from cliente";
            command = new MySqlCommand(sql, mConn);
            mdr = command.ExecuteReader();
            if (mdr.Read())
            {
                retorno = mdr.GetString("id");


            }
            DisconnectDB();
            return retorno;
        }

        public static string getvalue(string query, string field)
        {
            MYSQL.ConnectDB();

            string retorno = "0";
            string sql = string.Empty;
            sql = query;
            command = new MySqlCommand(sql, mConn);
            mdr = command.ExecuteReader();
            if (mdr.Read())
            {
                retorno = mdr.GetString(field);


            }
            DisconnectDB();
            return retorno;
        }

        public static string GetidChefia()
        {
            MYSQL.ConnectDB();

            string retorno = "0";
            string sql = string.Empty;
            sql = "select Convert(Max(Id)+1, char(50)) id from chefia";
            command = new MySqlCommand(sql, mConn);
            mdr = command.ExecuteReader();
            if (mdr.Read())
            {
                retorno = mdr.GetString("Id");


            }
            DisconnectDB();
            return retorno;
        }

        public static string GetidCarteira()
        {
            MYSQL.ConnectDB();

            string retorno = "0";
            string sql = string.Empty;
            sql = "select Convert(Max(Id)+1, char(50)) id from carteira";
            command = new MySqlCommand(sql, mConn);
            mdr = command.ExecuteReader();
            if (mdr.Read())
            {
                retorno = mdr.GetString("Id");


            }
            DisconnectDB();
            return retorno;
        }


        public static string Getcondpgto(string code)
        {
            MYSQL.ConnectDB();

            string retorno = "0";
            string sql = string.Empty;
            sql = "select id_cond from portalbuw1.condpgto where Code_SAP = " + code + "";
            command = new MySqlCommand(sql, mConn);
            mdr = command.ExecuteReader();
            if (mdr.Read())
            {
                retorno = mdr.GetString("id_cond");


            }
            DisconnectDB();
            return retorno;
        }


        public static string Getcarteira(string name)
        {
            MYSQL.ConnectDB();

            string retorno = "-1";
            string sql = string.Empty;
            sql = "select id from portalbuw1.carteira where nome = '" + name + "'";
            command = new MySqlCommand(sql, mConn);
            mdr = command.ExecuteReader();
            if (mdr.Read())
            {
                retorno = mdr.GetString("id");


            }
            DisconnectDB();
            return retorno;
        }

        public static string Getmaxidtabelapreco()
        {
            MYSQL.ConnectDB();

            string retorno = "0";
            string sql = string.Empty;
            sql = "select ifnull(max(id),1) id from portalbuw1.tabela_preco";
            command = new MySqlCommand(sql, mConn);
            mdr = command.ExecuteReader();
            if (mdr.Read())
            {
                retorno = mdr.GetString("id");

            }
            DisconnectDB();
            return retorno;
        }




        public static DataTable ItensMkt(int absid)
        {
            MYSQL.ConnectDB();
            DataTable mDataTable = new System.Data.DataTable("ItensMkt");
            //verificva se a conexão esta aberta
            if (mConn.State == ConnectionState.Open)
            {
                string query = "select id_item from portalbuw1.itenspedido where fk_pedido = " + absid + "";
                //cria um adapter usando a instrução SQL para acessar a tabela Clientes
                mAdapter = new MySqlDataAdapter(query, mConn);
                //preenche o dataset via adapter
                mAdapter.Fill(mDataSet, "ItensMkt");
                mDataTable = mDataSet.Tables["ItensMkt"];
            }
            DisconnectDB();
            return mDataTable;
        }

        public static DataTable ItensPedido(int absid)
        {
            MYSQL.ConnectDB();
            DataTable mDataTable = new System.Data.DataTable("ItensPedido");
            //verificva se a conexão esta aberta
            if (mConn.State == ConnectionState.Open)
            {
                string query = $"select id,DocEntry,LineNum,ItemCode,ItemName,LastSalesPrice,OnHand,WhsCode,Discount,Quantity,DATE_FORMAT(ShipDate, '%Y-%m-%d') AS ShipDate,Price,LineTotal,Usages from brb.ItensPedido ip \r\n where DocEntry = {absid}";
                //cria um adapter usando a instrução SQL para acessar a tabela Clientes
                mAdapter = new MySqlDataAdapter(query, mConn);
                //preenche o dataset via adapter
                mAdapter.Fill(mDataSet, "ItensPedido");
                mDataTable = mDataSet.Tables["ItensPedido"];
            }
            DisconnectDB();
            return mDataTable;
        }

        public static bool UpdateStatus(int absid, string cardcode)
        {
            try
            {
                MYSQL.ConnectDB();
                string query = "UPDATE Lead SET status = 'Inserido', CardSAP = '" + cardcode + "' WHERE id = " + absid + " ";
                MySqlCommand command = new MySqlCommand(query, mConn);
                command.ExecuteNonQuery();
            }
            catch
            {
                DisconnectDB();
                return false;
            }
            DisconnectDB();
            return true;
        }


        public static bool UpdateStatusPed(int absid, string docentry, string status)
        {
            try
            {
                MYSQL.ConnectDB();
                string query = $"UPDATE Pedido SET status = '{status}', NSAP = '{docentry}' WHERE id = " + absid + " ";
                MySqlCommand command = new MySqlCommand(query, mConn);
                command.ExecuteNonQuery();
            }
            catch
            {
                DisconnectDB();
                return false;
            }
            DisconnectDB();
            return true;
        }

        public static bool UpdateStatusCanc(int absid, string comments)
        {
            try
            {
                MYSQL.ConnectDB();
                string query = "UPDATE cliente SET status = 'Reprovado', RetornoSAP = '" + comments + "' WHERE absid = " + absid + " ";
                MySqlCommand command = new MySqlCommand(query, mConn);
                command.ExecuteNonQuery();
            }
            catch
            {
                DisconnectDB();
                return false;
            }
            DisconnectDB();
            return true;
        }
        public static bool UpdateStatusAprov(int absid)
        {
            try
            {
                MYSQL.ConnectDB();
                string query = "UPDATE cliente SET status = 'Aprovado', RetornoSAP = ''  WHERE absid = " + absid + " ";
                MySqlCommand command = new MySqlCommand(query, mConn);
                command.ExecuteNonQuery();
            }
            catch
            {
                DisconnectDB();
                return false;
            }
            DisconnectDB();
            return true;
        }

        public static bool TruncateTable(string Table)
        {
            try
            {
                MYSQL.ConnectDB();
                string query = "Truncate table " + Table + " ";
                MySqlCommand command = new MySqlCommand(query, mConn);
                command.ExecuteNonQuery();
            }
            catch
            {
                DisconnectDB();
                return false;
            }
            DisconnectDB();
            return true;
        }
        public static bool DeleteCond(string code)
        {
            try
            {
                MYSQL.ConnectDB();
                string query = "delete from condpgto where Code_SAP = '" + code + "'";
                MySqlCommand command = new MySqlCommand(query, mConn);
                command.ExecuteNonQuery();
            }
            catch
            {
                DisconnectDB();
                return false;
            }
            DisconnectDB();
            return true;
        }


        public static bool DeletePN(string code)
        {
            try
            {
                MYSQL.ConnectDB();
                string query = "delete from codcliente where codcliente = '" + code + "' ";
                MySqlCommand command = new MySqlCommand(query, mConn);
                command.ExecuteNonQuery();
            }
            catch
            {
                DisconnectDB();
                return false;
            }
            DisconnectDB();
            return true;
        }

        public static bool DeleteProduto(string code)
        {
            try
            {
                MYSQL.ConnectDB();
                string query = "delete from produto where id = '" + code + "' ";
                MySqlCommand command = new MySqlCommand(query, mConn);
                command.ExecuteNonQuery();
            }
            catch
            {
                DisconnectDB();
                return false;
            }
            DisconnectDB();
            return true;
        }
        public static bool DeleteProdutoLista(string code)
        {
            try
            {
                MYSQL.ConnectDB();
                string query = "delete from tabela_preco where fk_produto =  '" + code + "' ";
                MySqlCommand command = new MySqlCommand(query, mConn);
                command.ExecuteNonQuery();
            }
            catch
            {
                DisconnectDB();
                return false;
            }
            DisconnectDB();
            return true;
        }

        public static bool ExecInstruction(String ExecQuery)
        {
            try
            {
                MYSQL.ConnectDB();
                string query = ExecQuery;
                MySqlCommand command = new MySqlCommand(query, mConn);
                command.ExecuteNonQuery();

                Utils.Log("****************************************************");
                Utils.Log("Sucesso Inserindo Dados no Portal:");
                Utils.Log($"Query: {ExecQuery}");
            }
            catch (Exception ex)
            {
                Utils.Log("****************************************************");
                Utils.Log("Erro Inserindo Dados no Portal:");
                Utils.Log($"Query: {ExecQuery}");
                Utils.Log($"Erro: {ex.ToString()}");
                DisconnectDB();
                return false;
            }
            DisconnectDB();
            return true;
        }

    }
}
