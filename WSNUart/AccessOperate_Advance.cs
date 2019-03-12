using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;
using System.IO;

namespace WSNUart
{
    class AccessOperate
    {
        public AccessOperate(string ConnectionString) 
        {
            m_myConnectionString = ConnectionString;
        }
        public AccessOperate()
        {
        }
        string setConnectionString(string strAceVersion, string strStorePath, string strTableName, string strPassword) 
        {
            string myConnectionString = "";
            return myConnectionString;
        }

        string SetstrQueryString(string strTableName)
        {
            string strQueryString = string.Format("select * from {0}", strTableName);
            return strQueryString;
        }

        OleDbConnection MyExecuteConnection(string ConnectionString)
        {
            OleDbConnection Cnn = new OleDbConnection(ConnectionString);
            return Cnn;
        }

        OleDbDataAdapter MyExecuteDataAdapter(string strTableName, string myConnectionString) 
        {
            string strSelectSql = SetstrQueryString(strTableName);
            OleDbConnection objCnn = MyExecuteConnection(myConnectionString);
            objCnn.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(strSelectSql, objCnn);
            objCnn.Close();
            return da;
        }

        public DataSet MyExecuteDataSet(string strTableName, string myConnectionString)
        {
            OleDbDataAdapter objDa = MyExecuteDataAdapter(strTableName, myConnectionString);
            DataSet ds = new DataSet();
            objDa.Fill(ds);
            return ds;
        }

        public DataSet MyExecuteDataSet(string sql)
        {
            OleDbDataAdapter objDa = new OleDbDataAdapter(sql, MyExecuteConnection(m_myConnectionString));
            DataSet ds = new DataSet();
            objDa.Fill(ds);
            return ds;
        }

        void InsertData(string strTableName, string myConnectionString) 
        {
            DataSet objDs = MyExecuteDataSet(strTableName, myConnectionString);
            DataTable dt = objDs.Tables[0];
            OleDbDataAdapter objDa = MyExecuteDataAdapter(strTableName, myConnectionString);
            OleDbCommandBuilder cmdbuilder = new OleDbCommandBuilder(objDa);
        }

        void DataTableTraversal() 
        {
            
        }

        // 可以自定义构造函数 配置数据库
        private string m_myConnectionString;// = string.Format("Provider={0};Data Source={1}", "Microsoft.ACE.OLEDB.12.0", "|DataDirectory|bishe.accdb");
        private string getConnectionString() 
        {
            return m_myConnectionString;
        }
        
        public int myExecuteNoQuery(string sql)
        {
            OleDbConnection cnn = new OleDbConnection(m_myConnectionString);
            cnn.Open();
            OleDbCommand cmd = new OleDbCommand(sql, cnn);
            // OleDbCommand cmd = new OleDbCommand(sql, cnn);
            cmd.CommandType = CommandType.Text;
            int i = cmd.ExecuteNonQuery();
            cnn.Close();
            return i;
        }

        public int myExecuteNoQuery(string sql, string myConnectionString)
        {
            OleDbConnection cnn = new OleDbConnection(myConnectionString);
            cnn.Open();
            OleDbCommand cmd = new OleDbCommand(sql, cnn);
            cmd.CommandType = CommandType.Text;
            int i = cmd.ExecuteNonQuery();
            cnn.Close();
            return i;
        }

        public int myExecuteNoQuery(string sql, string myConnectionString, bool isLast)
        {
            OleDbConnection cnn = new OleDbConnection(myConnectionString);
            if (cnn.State != ConnectionState.Open) 
            {
                cnn.Open();
            }            
            OleDbCommand cmd = new OleDbCommand(sql, cnn);
            cmd.CommandType = CommandType.Text;
            int i = cmd.ExecuteNonQuery();
            if(isLast)
            {
                cnn.Close();
            }
            
            return i;
        }
    }
}
//select p.name,s.serial from sn s,programloader p where s.usernumber=p.number