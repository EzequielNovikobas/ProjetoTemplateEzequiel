using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.IO;
using System.Reflection;
using System.Data.SqlClient;
using System.Data;

namespace AutomacaoGoogleAcademicoDB.DAL
{
    class AccessDB
    {
        private OleDbConnection conn = new System.Data.OleDb.OleDbConnection();
        public void ConectaDB()
        {

            try
            {
                String currentPath = System.Environment.CurrentDirectory + "\\ProjetoTemplate.accdb";
                
                conn.ConnectionString = ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + currentPath);

                conn.Open();
            }
            catch (Exception ex)
            {
                throw new Exception("ERRO:" + ex.Message + " Classe: AccessDB, Método: ConectaDB");
            }
        }

        public void FechaDB()
        {

            try
            {
                conn.Dispose();
                conn.Close();
            }
            catch (Exception ex)
            {
                throw new Exception("ERRO:" + ex.Message + " Classe: AccessDB, Método: FechaDB");
            }
        }

        public void ExecutaQry(string Qry)
        {
            OleDbCommand cmd = new OleDbCommand(Qry, conn);


            try
            {
                //cmd.CommandType = CommandType.Text;

                //cmd.Connection = conn;

                //cmd.CommandText = Qry;

                cmd.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                throw new Exception("ERRO:" + ex.Message + " Classe: AccessDB, Método: ConectaDB");
            }
        }

        public DataSet DS(string Qry, string Tabela)
        {
            OleDbCommand cmd = new OleDbCommand();
            DataSet Ds = new DataSet();
            OleDbDataAdapter Da = new OleDbDataAdapter();

            try
            {
                cmd.CommandType = CommandType.Text;

                cmd.Connection = conn;

                cmd.CommandText = Qry;

                Da.SelectCommand = cmd;

                Da.Fill(Ds, Tabela);

                return Ds;

            }
            catch (Exception ex)
            {
                throw new Exception("ERRO:" + ex.Message + " Classe: AccessDB, Método: ConectaDB");
            }
        }
        public OleDbDataReader DR(string Qry)
        {
            OleDbCommand cmd = new OleDbCommand();


            try
            {
                cmd.CommandType = CommandType.Text;

                cmd.Connection = conn;

                cmd.CommandText = Qry;

                return cmd.ExecuteReader();

            }
            catch (Exception ex)
            {
                throw new Exception("ERRO:" + ex.Message + " Classe: AccessDB, Método: ConectaDB");
            }
        }
    }
}
