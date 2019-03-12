using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Windows.Forms;

namespace WSNUart
{
    class TestData_BLL
    {
        public AccessOperate m_aceOP;
        public Dictionary<string, List<string>> m_myDic;
        //public List<string> m_listFiled;
        //public List<string> m_listValue;
        private string m_tableName;
        //public int m_length;
        //public string dbName;
        //private m_con
        public TestData_BLL()
        {
            m_tableName = "";
            //m_listFiled = new List<string>();
            //ini_Filed(ref m_listFiled);
            //m_length = m_listFiled.Count();

            //m_listValue = new List<string>();
            //ini_strFieldValue(ref m_listValue, m_length);
            m_aceOP = new AccessOperate();            
            m_myDic = new Dictionary<string, List<string>>();
        }

        public void setTableName(string tableName) 
        {
            m_tableName = tableName;
        }

        public string getTableName()
        {
            return m_tableName;
        }
        //void ini_Filed(ref List<string> strFiled) 
        //{
        //    strFiled.Add("id");
        //    strFiled.Add("recvPkgCount");
        //    strFiled.Add("lossPkgRatio");
        //    strFiled.Add("avgRSSI");
        //    strFiled.Add("voltDiff");
        //    strFiled.Add("avgHopCount");
        //    strFiled.Add("avgTimeDiff");
        //}
        
        void ini_strFieldValue(ref List<string> listValue, int count)
        {
            for (int i = 0; i < count; i++) 
            {
                listValue.Add("");
            }
        }       

        public string LinkString(string prefix, string suffix)      // 连接字符串
        {
            string IntactStr = "";
            StringBuilder stringbuilder = new StringBuilder();
            stringbuilder.Append(prefix);
            stringbuilder.Append(suffix);
            IntactStr = stringbuilder.ToString();
            stringbuilder.Clear();
            return IntactStr;
        }

        public string getFiledStrCon(List<string> listFiledStr)                       // sql语句中 作为 字段名 的 字符串 不加单引号 
        {        
            string strFiledContact = "";
            int n = listFiledStr.Count;
            for (int i = 0; i < n; ++i)
            {
                strFiledContact +=  listFiledStr[i];
                if (i != n -1)
                {
                    strFiledContact += ",";
                }
            }
            return strFiledContact;
        }

        public string getValueStrCon(List<string> listValueStr)                     // sql 语句中 作为 字段值 的 字符串 要加单引号 
        {
            string strValueContact = "";           
            int n = listValueStr.Count;
            for (int i = 0; i < n; ++i)
            {
                strValueContact += ("\'" + listValueStr[i] + "\'");
                if (i != n - 1)
                {
                    strValueContact += ",";
                }
            }
            return strValueContact;
        }
        
        string LongStr(string [] strArray)                              // {1} 字段名
        {
            string strSnowball = "";
            foreach (string str in strArray)
            {
                if (str == strArray[0])                                        // 字符串数组第一个元素 特别处理
                {
                    strSnowball = LinkString(str, ",");
                }
                else 
                {
                    if (str != strArray[strArray.Length - 1])                    //不是第一个元素也不是最后一个数组元素
                    {
                        strSnowball = LinkString(LinkString(strSnowball, str), ",");    //  strtemp = LinkString(strtemp, str); strtemp = LinkString(strtemp, ",");
                    }
                    else                                                        // 最后一个字符串元素
                    {
                        strSnowball = LinkString(strSnowball, str);
                    }
                }
            }
            return strSnowball;
        }

        public int InsertData(string TableName, string FieldNameCon, string FieldValueCon, string connString)
        {
            string sql = string.Format("insert into {0}({1}) values ({2})", TableName, FieldNameCon, FieldValueCon);
            int isSuccess = m_aceOP.myExecuteNoQuery(sql, connString);
            if (isSuccess > 0)
            { ;}
            else
            {
                MessageBox.Show("增加记录失败！", "提示");       // 插入记录失败的操作
            }
            return isSuccess;
        }
 
        void UpdataData(string TableName, string FieldNameEqualValue, string MajorKey, string MajorValue)
        {
            string sql = string.Format("update {0} set {1} where {2}='{3}'",TableName, FieldNameEqualValue, MajorKey, MajorValue);
            int flag = m_aceOP.myExecuteNoQuery(sql);
            if (flag > 0)
            {
                MessageBox.Show("更新记录成功！", "更新数据提示");        // 插入成功的操作
            }
            else
            {
                MessageBox.Show("更新记录失败！", "更新数据提示");        // 插入记录失败的操作
            }
        }

        int SelectWhere(string TableName, string MajorKey, string MajorValue) 
        {
            string sql = string.Format("select * from {0} where {1}='{2}'", TableName, MajorKey, MajorValue);
            int RefLineNumber = m_aceOP.myExecuteNoQuery(sql);
            return RefLineNumber;
        }

        string LongStrUpdate(string[] strArrayFN, string[] strArrayFV)
        {
            string strSnowball = "";
            int i = 0;
            int n = strArrayFN.Length;
            for (i = 1; i < n; ++i)
            {
                strSnowball += strArrayFN[i];
                strSnowball += "=";
                strSnowball += "\'";
                strSnowball += strArrayFV[i];
                strSnowball += "\'";
                if (i != n - 1)
                {
                    strSnowball += ",";
                }
            }
            return strSnowball;
        }
    }
}