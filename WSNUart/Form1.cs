/* USST WSN 上位机测试程序
 * 170516 去掉开启串口时上位机发送‘2’以及后续两个字节的数据
 *170518 字节拼接改用.net中的方法
 *170519 timeStap
 *170818 布局界面 修改项目名，增加usst图标
 *170823 计算丢包数量，可写数据库新表
 *3.1.3 170827 多线程串口通信优化
 *3.1.4 170829 线程方法中增加Sleep(0)
 *导出文件
 */
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using System.IO.Ports;
using System.Threading;
using System.Data.OleDb;
using System.IO;
using System.Security;
using System.Diagnostics;

namespace WSNUart
{
    public partial class Form1 : Form
    {
        const string strAuthor = "WING@USST";
        const string strVerNO = "3.3.0";
        public struct pkgFrame
        {
            public int nodeId;
            public int protId;             //LS170508协议类型
            public int rssi;
            public double volt;
            public ushort seqNO;
            public string time;
            public int hopCount;             //LS170502跳数
            public int timeDifference;     //LS170504传输时差（发送到接收时差）该值不为负数
            public string timeDifference_0x;     //LS170504传输时差（发送到接收时差）该值不为负数
        }

        public struct FormatPacket
        {
            public string nodeId;
            public string protId;
            public string rssi;
            public string volt;
            public string seqNO;
            public string hopCount;
            public string timeDifference;
            public int pktCnt; 
        }

        enum ByteOrder
        {
            CODEID = 0,
            PROTID = 1,        //LS170508协议类型
            NODEID = 2,
            SEQ_NO = 4,
            RSSI_VAL = 6,
            VOLT = 7,
            HOPCPUNT = 9,
            TIMEDIFFERENCE = 10,   // 4个字节            
            BYTECOUNT = 14        //下位机有效载荷部分共有BYTECOUNT个字节的数据
        };

        enum CONSTDEF
        {
            FULLVALUE =0x10000,
            
        };

        enum SUMMVIEWROWNO
        {
            AVGLOSSR=0,
            AVGRSSI=1,
            AVGVOLTDEC=2,
            AVGHOPS=3,
            AVGDELAY=4,
            PCKTOTAL=5,
        };

        enum SYSSTATE
        {
            S_IDLE = 0,
            S_START = 1,
            S_STANDBY = 2,
        };
        SYSSTATE m_state;

        const int SUMMVIEWROWS = 6;

        const byte CODEID_VALUE = 0x44;
        const string ACK = "264\n";             //LS1705101512
        const int DEFAULT_PKG_NUM = 10;

        const int VOLT_COLUMN = 3;

        //表的列名（字段名）
        const string COLUM_NODE_ID = "nodeId";
        const string COLUM_RSSI = "rssi";
        const string COLUM_SEQ_NO = "seqNO";
        const string COLUM_GET_TIME = "gettime";
        const string COLUM_VOLT = "volt";
        const string COLUM_HOP_COUNT = "hopCount";
        const string COLUM_TIME_DIFF = "timeDifference";
        const string COLUM_PROTID = "protocol";
        const string TABLE_INPUTDATA_NAME = "Results";            //buffer中的数据库表名
        const string TABLE_STATISTIC = "Results_Statistic";   //用于平均求和的统计表的表名

        const string TABLENAME_COMPUTE = "Compute";
        const string COLUM_NODE_ID_AVG = "nodeId";
        const string COLUM_LOSS_PKG_COUNT = "lostPktCount";   //接收的数据包数量
        const string COLUM_LOSS_PKG_RATIO = "lostPktRatio";
        const string COLUM_AVG_RSSI = "avgRSSI";
        const string COLUM_VOLT_DIFF = "voltDiff";
        const string COLUM_AVG_HOP = "avgHopCount";
        const string COLUM_AVG_TIME_DIFF = "avgTimeDiff";
        const string COLUM_AVG_PKTCNT = "pktCnt";
        const string COLUM_AVG_INITVOLT = "initVolt";
        const string COLUM_AVG_INITSN = "initSN";
        const string COLUM_AVG_RSSISUM = "RssiSum";
        const string COLUM_AVG_HOPSUM = "HopSum";
        const string COLUM_AVG_DELAYSUM = "DelaySum";
        const string COLUM_AVG_REPCNT = "repCnt";
        const string COLUM_AVG_LASTSN = "lastSN";

        const string TABLENAME_SUMMARY = "TestSummary";
        const string COLUM_SUM_TBNAME = "TableName";
        const string COLUM_SUM_LOSSCNT = "SumLossCnt";
        const string COLUM_SUM_AVGLOSSRATIO = "AvgLossR";
        const string COLUM_SUM_AVGRSSI = "AvgRSSI";
        const string COLUM_SUM_AVGVOLTDEC = "AvgVoltDec";
        const string COLUM_SUM_AVGHOP = "AvgHops";
        const string COLUM_SUM_AVGDELAY = "AvgDelay";
        const string COLUM_SUM_PCKCNT = "AvgPckCnt";
        const string COLUM_SUM_REPCNT = "repCnt";
        const string COLUM_SUM_SPAN = "TimeSpan";

        ///异常值过滤阈值
        const double EXMAXVOLT = 3.5;
        const double EXMINVOLT = 0.0;
        const int EXMAXNODEID = 0x100;
        const int EXMINNODEID = 0x0;

        Thread readThread;
        public SerialPort spInstance;
        public pkgFrame m_Result;
        protected ushort m_PkgCount = 0;

        private DataTable dtOri;
        private DataTable dtStatistics;          //存储统计值
        private DataTable m_dtSummary;
        private DataSet dataset;
        private string m_connStr;
        private int m_PacketNum;         
        private int m_SendPeriod;

        TestData_BLL testdata_bll_1;
        private string m_tableName;
        private List<string> m_listField, m_listFiled_Results;
        private List<string> m_listValue;
        private int m_length;

        public string strProvider;
        public string strDBName;
        public string dest_dbPath;
        public string strCustomerDirectory;
        public string strTemplateDb;
        
        volatile bool isSpOpen = false;
        volatile bool RecoEn = false;

        //Number of nodes having been finished collection.
        int m_DoneCntr;
        bool m_bPauseScroll;
        bool m_bDate4Save;
        string strCurrentPath;
        //Indicate whether have input data already, has effect on status strip labels display.
        bool m_bHasInput;
        

        public Form1()
        {
            InitializeComponent();
            initParamters();


            m_listValue = new List<string>();
            m_listFiled_Results = new List<string>();
            ini_Filed_Results(ref m_listFiled_Results);


            setTablename(TABLENAME_COMPUTE);
            testdata_bll_1 = new TestData_BLL();
            testdata_bll_1.setTableName(this.getTableName());
            m_listField = new List<string>();
            ini_Field(ref m_listField);
            m_length = m_listField.Count();
            m_listValue = new List<string>();
            //ini_strFieldValue(ref m_listValue, m_length);

            strProvider = "Microsoft.ACE.OLEDB.12.0";
            strCustomerDirectory = "WSNTestData";
            strTemplateDb = "templateDB.accdb";
            strCurrentPath = Environment.CurrentDirectory;
            dest_dbPath = string.Format("{0}\\{1}", strCurrentPath, strCustomerDirectory);
            strDBName = "WSN_" + CurrentTime() + ".accdb";
            if (!Directory.Exists(dest_dbPath))
            {
                Directory.CreateDirectory(dest_dbPath);
            }
            System.IO.File.Copy(strCurrentPath + "\\" + strTemplateDb, dest_dbPath + "\\" + strDBName, true);          //根据模板创建db文件
            
        }


        public string getconnStr()
        {
            return m_connStr;
        }

        public void setTablename(string tableName)
        {
            m_tableName = tableName;
        }
        public string getTableName()
        {
            return m_tableName;
        }

        void ini_Field(ref List<string> strField)
        {
            strField.Add(COLUM_NODE_ID);
            strField.Add(COLUM_LOSS_PKG_COUNT);
            strField.Add(COLUM_LOSS_PKG_RATIO);
            strField.Add(COLUM_AVG_RSSI);
            strField.Add(COLUM_VOLT_DIFF);
            strField.Add(COLUM_AVG_HOP);
            strField.Add(COLUM_AVG_TIME_DIFF);
        }
        //void ini_strFieldValue(ref List<string> listValue, int count)
        //{
        //    for (int i = 0; i < count; i++)
        //    {
        //        listValue.Add("");
        //    }
        //}
        void ini_Filed_Results(ref List<string> strFiled)
        {
            strFiled.Add(COLUM_NODE_ID);
            strFiled.Add(COLUM_RSSI);
            strFiled.Add(COLUM_SEQ_NO);
            strFiled.Add(COLUM_GET_TIME);
            strFiled.Add(COLUM_VOLT);
            strFiled.Add(COLUM_HOP_COUNT);
            strFiled.Add(COLUM_TIME_DIFF);
        }


        public void ReadSerialPort()
        {
            byte[] testResult = new byte[(int)ByteOrder.BYTECOUNT];

            while (true)
            {
                if (isSpOpen && RecoEn /*&& m_Result.seqNO < 20*m_PacketNum*/)
                {
                    try
                    {
                        testResult[(int)ByteOrder.CODEID] = (byte)spInstance.ReadByte();//串口读入字节到testResult
                        switch (testResult[(int)ByteOrder.CODEID])
                        {
                            case CODEID_VALUE:
                                {
                                    while (spInstance.BytesToRead < (int)ByteOrder.BYTECOUNT - 1) ;   //接收缓冲区中数据的字节数
                                    spInstance.Read(testResult, (int)ByteOrder.CODEID + 1, (int)ByteOrder.BYTECOUNT - 1); //testResult中第一个字节存储codeid
                                                                                                                          //spInstance.Write(ACK);
                                    m_Result.protId = testResult[(int)ByteOrder.PROTID];
                                    m_Result.nodeId = (ushort)((testResult[(int)ByteOrder.NODEID] << 8) + testResult[(int)ByteOrder.NODEID + 1]);
                                    if (m_Result.protId != 0x10)
                                    {
                                        if (testResult[(int)ByteOrder.RSSI_VAL] <= 128)
                                        {
                                            m_Result.rssi = -45 + testResult[(int)ByteOrder.RSSI_VAL] - 256;
                                        }
                                        else
                                        {
                                            m_Result.rssi = -45 + testResult[(int)ByteOrder.RSSI_VAL];
                                        }
                                    }
                                    else
                                    {
                                        m_Result.rssi = testResult[(int)ByteOrder.RSSI_VAL];
                                    }


                                    m_Result.volt = ((testResult[(int)ByteOrder.VOLT] << 8) + testResult[(int)ByteOrder.VOLT + 1]) * 3.75 / 8192;
                                    m_Result.hopCount = testResult[(int)ByteOrder.HOPCPUNT];
                                    m_Result.timeDifference = (testResult[(int)ByteOrder.TIMEDIFFERENCE] << 24) + (testResult[(int)ByteOrder.TIMEDIFFERENCE + 1] << 16) +
                                        (testResult[(int)ByteOrder.TIMEDIFFERENCE + 2] << 8) + testResult[(int)ByteOrder.TIMEDIFFERENCE + 3];
                                    //数组低字段数值高地址 同书写顺序

                                    Byte[] tempTimeDiff_0x = new Byte[] {
                                                            testResult[(int)ByteOrder.TIMEDIFFERENCE],
                                                            testResult[(int)ByteOrder.TIMEDIFFERENCE + 1],
                                                            testResult[(int)ByteOrder.TIMEDIFFERENCE + 2],
                                                            testResult[(int)ByteOrder.TIMEDIFFERENCE + 3]
                                                            };

                                    m_Result.timeDifference_0x = BitConverter.ToString(tempTimeDiff_0x, 0).Replace("-", "");

                                    m_Result.seqNO = (ushort)((testResult[(int)ByteOrder.SEQ_NO] << 8) + testResult[(int)ByteOrder.SEQ_NO + 1]);  //BitConverter.ToUInt16(tempSeqNo, 0);
                                    m_Result.time = DateTime.Now.ToString();

                                    m_PkgCount++;

                                    this.Invoke(new MethodInvoker(delegate
                                    {
                                        if (dtStatistics.Rows.Contains(m_Result.nodeId))
                                        {
                                            DataRow drSel = dtStatistics.Rows.Find(m_Result.nodeId);

                                            if ((int)drSel[COLUM_AVG_PKTCNT] < m_PacketNum)
                                            {
                                                if (isFulled(drSel,m_Result.seqNO))/*(int)drSel[COLUM_AVG_PKTCNT] == m_PacketNum)*///按序号距离判断
                                                {
                                                    updateVoltage(drSel, m_Result.volt);
                                                    m_DoneCntr++;
                                                }
                                                else
                                                {
                                                    updateStatItem(drSel);
                                                    appendData();
                                                }
                                                //updateStatItem(drSel);
                                                //appendData();
                                            }
                                            else
                                            {
                                                updateVoltage(drSel, m_Result.volt);
                                                m_DoneCntr++;
                                            }
                                        }
                                        else
                                        {
                                            addNewStatItem();
                                            appendData();
                                        }
                                        updateStatusInfo();

                                    }));

                                    for (int i = 0; i < (int)ByteOrder.BYTECOUNT; i++)
                                    {
                                        testResult[i] = 0;
                                    }

                                    checkFinish(dtStatistics.Rows.Count,m_DoneCntr);
                                    break;
                                }
                            default:
                                {
                                    break;
                                }
                        }
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show(e.Message);
                    }
                }
                else
                {
                    try
                    {
                        //strStatus = "线程已经休眠";
                        Thread.Sleep(Timeout.Infinite);
                    }
                    catch (ThreadInterruptedException)
                    {
                        //strStatus = "线程已经唤醒"; 
                    }
                }
                Thread.Sleep(0);
            }
        }





        public bool isFulled(DataRow drSel,int seqNO)
        {
            //int maxNo = Convert.ToInt32(drSel[COLUM_AVG_LASTSN]);//找最大
            int minNo = Convert.ToInt32(drSel[COLUM_AVG_INITSN]);
            if (seqNO - minNo < m_PacketNum)
            {
                //if (maxNo - minNo > m_PacketNum - 1)
                //{
                //    DataRow[] dr0 = dtOri.Select(string.Format("{0}={1} and {2}={3}", COLUM_NODE_ID, nodeId, COLUM_SEQ_NO, maxNo));
                //    try
                //    {
                //        foreach (DataRow row in dr0)
                //            dtOri.Rows.Remove(row);//移除dtOri中maxNo所在行
                //    }
                //    catch (Exception ex)
                //    {
                //        Console.WriteLine(string.Format("catch unsolve exception: {0}\r\n异常信息：{1}\r\n异常堆栈：{2}", ex.GetType(), ex.Message, ex.StackTrace));
                //    }
                //}
                return false;
            }
            return true;
        }



        public DataTable ToDataTable(DataRow[] rows)  
        {  
            if (rows == null || rows.Length == 0) return null;  
            DataTable tmp = rows[0].Table.Clone();  // 复制DataRow的表结构  
            foreach (DataRow row in rows)  
                tmp.Rows.Add(row.ItemArray);  // 将DataRow添加到DataTable中  

            return tmp;  
        } 


        public string CurrentTime() // 插入当前时间到时间字段
        {
            string strTime;
            DateTime dateTime = DateTime.Now;
            strTime = dateTime.ToString("yyyyMMdd_HHmmss");
            return strTime;
        }
        

        private void btOpenSer_Click(object sender, EventArgs e)  //打开串口
        {
            btOpenSer.Enabled = false;
            btCloseSer.Enabled = true;
            if (cbSerialPort.Text != "")
            {
                
                if (!isSpOpen)
                {
                    try
                    {
                        spInstance.PortName = cbSerialPort.Text;
                        spInstance.BaudRate = Convert.ToInt32(BaudListBox.Text);
                        spInstance.StopBits = StopBits.One;
                        spInstance.Open();
                        if (spInstance.IsOpen)
                        {
                            isSpOpen = true;
                            m_state = SYSSTATE.S_STANDBY;
                            try
                            {
                                readThread.Interrupt();
                            }
                            catch(SecurityException se)
                            {
                                MessageBox.Show("读串口线程未唤醒，原因："+se.Message);
                            }
                            showSatus("串口已打开");
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                
            }
            else
            {
                showSatus("未填串口参数，串口未开");
            }
        }

        private void initParamters()
        {
            btOpenSer.Enabled = false;          // 初始状态下，连接串口按钮禁用（选串口后使能）
            btCloseSer.Enabled = false;

            if (cbSerialPort.Text == "")
            {
                showSatus("请选择串口");
            }
            spInstance = new SerialPort();
            
            //ThreadStart threadDelegate = new ThreadStart(ReadSerialPort);
            //readThread = new Thread(threadDelegate);
            readThread = new Thread(ReadSerialPort);
            readThread.IsBackground = true;
            readThread.Start();
            isSpOpen = false;
            RecoEn = false;

            m_PacketNum = DEFAULT_PKG_NUM;      //默认发送的个数
            m_SendPeriod = 100;
            //tbPeriod.Text = m_SendPeriod.ToString();
            tbPeriod1.Text = m_SendPeriod.ToString();
            
            BaudListBox.SelectedIndex = 2;

            tbPkgCount.Text = m_PacketNum.ToString();
            initDataSet();
            initInDatagrid();
            initDgvStatistic();
            initSummView();

            m_bPauseScroll = false;
            PauseButn.Checked = m_bPauseScroll;
            switchStatus(0);
            m_bHasInput = false;
            m_state = SYSSTATE.S_IDLE;
        }

        private void initDataSet()
        {
            dtOri = new DataTable(TABLE_INPUTDATA_NAME);
            dtOri.Columns.Add(COLUM_NODE_ID, typeof(int));
            dtOri.Columns.Add(COLUM_SEQ_NO, typeof(int));
            dtOri.Columns.Add(COLUM_RSSI, typeof(int));
            dtOri.Columns.Add(COLUM_VOLT, typeof(double));
            dtOri.Columns.Add(COLUM_HOP_COUNT, typeof(int));
            dtOri.Columns.Add(COLUM_TIME_DIFF, typeof(long));
            dtOri.Columns.Add(COLUM_GET_TIME, typeof(string));
            //dtOri.PrimaryKey = new DataColumn[2]{ dtOri.Columns[COLUM_NODE_ID], dtOri.Columns[COLUM_SEQ_NO] };


            dtStatistics = new DataTable(TABLE_STATISTIC);
            dtStatistics.Columns.Add(COLUM_NODE_ID_AVG, typeof(int));
            dtStatistics.Columns.Add(COLUM_LOSS_PKG_COUNT, typeof(int));
            dtStatistics.Columns.Add(COLUM_LOSS_PKG_RATIO, typeof(double));
            dtStatistics.Columns.Add(COLUM_AVG_RSSI, typeof(double));
            dtStatistics.Columns.Add(COLUM_VOLT_DIFF, typeof(double));
            dtStatistics.Columns.Add(COLUM_AVG_HOP, typeof(double));
            dtStatistics.Columns.Add(COLUM_AVG_TIME_DIFF, typeof(double));
            dtStatistics.Columns.Add(COLUM_AVG_PKTCNT, typeof(int));
            dtStatistics.Columns.Add(COLUM_AVG_INITVOLT, typeof(double));
            dtStatistics.Columns.Add(COLUM_AVG_INITSN, typeof(int));
            dtStatistics.Columns.Add(COLUM_AVG_RSSISUM, typeof(double));
            dtStatistics.Columns.Add(COLUM_AVG_HOPSUM, typeof(double));
            dtStatistics.Columns.Add(COLUM_AVG_DELAYSUM, typeof(double));
            dtStatistics.Columns.Add(COLUM_AVG_REPCNT, typeof(int));
            dtStatistics.Columns.Add(COLUM_AVG_LASTSN, typeof(int));
            DataColumn[] keys = new DataColumn[1];
            keys[0] = dtStatistics.Columns[COLUM_NODE_ID_AVG];
            dtStatistics.PrimaryKey = keys;

            m_dtSummary = new DataTable(TABLENAME_SUMMARY);
            m_dtSummary.Columns.Add(COLUM_SUM_LOSSCNT, typeof(int));
            m_dtSummary.Columns.Add(COLUM_SUM_AVGLOSSRATIO, typeof(double));
            m_dtSummary.Columns.Add(COLUM_SUM_AVGRSSI, typeof(double));
            m_dtSummary.Columns.Add(COLUM_SUM_AVGVOLTDEC, typeof(double));
            m_dtSummary.Columns.Add(COLUM_SUM_AVGHOP, typeof(double));
            m_dtSummary.Columns.Add(COLUM_SUM_AVGDELAY, typeof(double));
            m_dtSummary.Columns.Add(COLUM_SUM_PCKCNT, typeof(int));
            m_dtSummary.Columns.Add(COLUM_SUM_TBNAME, typeof(string));
            m_dtSummary.Columns.Add(COLUM_SUM_REPCNT, typeof(int));

            dataset = new DataSet();
            dataset.Tables.Add(dtOri);
            dataset.Tables.Add(dtStatistics);     //在内存中建立dtStatistic表
            dataset.Tables.Add(m_dtSummary);
        }

        private void appendData()
        {
            DataRow dr;
            dr = dataset.Tables[TABLE_INPUTDATA_NAME].NewRow();      // dr data row
            dr[COLUM_NODE_ID] = m_Result.nodeId;
            dr[COLUM_RSSI] = m_Result.rssi;
            dr[COLUM_SEQ_NO] = m_Result.seqNO;
            dr[COLUM_GET_TIME] = m_Result.time;
            dr[COLUM_VOLT] = Math.Round(m_Result.volt, 5);
            dr[COLUM_HOP_COUNT] = m_Result.hopCount;
            dr[COLUM_TIME_DIFF] = m_Result.timeDifference;

            this.Invoke(new MethodInvoker(delegate
            {
                dataset.Tables[TABLE_INPUTDATA_NAME].Rows.Add(dr);
                if (!m_bPauseScroll)
                    dataGridView1.FirstDisplayedScrollingRowIndex = dataGridView1.Rows.Count - 1;
            }));

            m_listValue.Add(dr[COLUM_NODE_ID].ToString());
            m_listValue.Add(dr[COLUM_RSSI].ToString());
            m_listValue.Add(dr[COLUM_SEQ_NO].ToString());
            m_listValue.Add(dr[COLUM_GET_TIME].ToString());
            m_listValue.Add(dr[COLUM_VOLT].ToString());
            m_listValue.Add(dr[COLUM_HOP_COUNT].ToString());
            m_listValue.Add(dr[COLUM_TIME_DIFF].ToString());

            m_connStr = string.Format("Provider={0}; Data Source={1};", strProvider, dest_dbPath + "\\" + strDBName);

            testdata_bll_1.InsertData(TABLE_INPUTDATA_NAME, testdata_bll_1.getFiledStrCon(m_listFiled_Results), testdata_bll_1.getValueStrCon(m_listValue), getconnStr());


            m_listValue.Clear();
            if (this.dataGridView1.RowCount >= 0x100)
            {
                this.dataGridView1.Rows.RemoveAt(0);
                dataset.Tables[TABLE_INPUTDATA_NAME].Rows.RemoveAt(0);
            }
        }

        private bool saveResults()
        {
            string strSQL;
            using (OleDbConnection conn = new OleDbConnection(m_connStr))
            {
                try
                {
                    conn.Open();
                    showSatus("正在保存数据...");
                    for (int rowIdx = 0; rowIdx < m_dtSummary.Rows.Count - 1; rowIdx++)
                    {
                        OleDbCommand cmd = new OleDbCommand();
                        cmd.Connection = conn;
                        strSQL = COLUM_NODE_ID_AVG + ","
                            + COLUM_LOSS_PKG_COUNT + ","
                            + COLUM_LOSS_PKG_RATIO + ","
                            + COLUM_AVG_RSSI + ","
                            + COLUM_VOLT_DIFF + ","
                            + COLUM_AVG_HOP + ","
                            + COLUM_AVG_TIME_DIFF + ","
                            + COLUM_AVG_PKTCNT + ","
                            + COLUM_AVG_REPCNT;


                        cmd.CommandText = "INSERT INTO " + TABLENAME_COMPUTE + "(" + strSQL + ") VALUES ("
                            + "'" + "0000" + dtStatistics.Rows[rowIdx][COLUM_NODE_ID_AVG].ToString() + "'" + ","
                            + "'" + m_dtSummary.Rows[rowIdx][COLUM_SUM_LOSSCNT].ToString() + "'" + ","
                            + "'" + m_dtSummary.Rows[rowIdx][COLUM_SUM_AVGLOSSRATIO].ToString() + "'" + ","
                            + "'" + m_dtSummary.Rows[rowIdx][COLUM_SUM_AVGRSSI].ToString() + "'" + ","
                            + "'" + m_dtSummary.Rows[rowIdx][COLUM_SUM_AVGVOLTDEC].ToString() + "'" + ","
                            + "'" + m_dtSummary.Rows[rowIdx][COLUM_SUM_AVGHOP].ToString() + "'" + ","
                            + "'" + m_dtSummary.Rows[rowIdx][COLUM_SUM_AVGDELAY].ToString() + "'" + ","
                            + "'" + m_dtSummary.Rows[rowIdx][COLUM_SUM_PCKCNT].ToString() + "'" + ","
                            + "'" + m_dtSummary.Rows[rowIdx][COLUM_SUM_REPCNT].ToString() + "'" + ")";

                        cmd.ExecuteNonQuery();
                    } 
                    conn.Close();
                    showSatus("保存成功.");
                    return true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    showSatus("保存失败.");
                    return false;
                }
            }
        }

        private void initInDatagrid()
        {
            DataGridViewTextBoxColumn dgvc = new DataGridViewTextBoxColumn();
            dgvc.DataPropertyName = COLUM_NODE_ID;
            dgvc.HeaderText = "节点ID";
            dgvc.ValueType = typeof(int);
            dgvc.DefaultCellStyle.Format = "X4";
            dataGridView1.Columns.Add(dgvc);

            dgvc = new DataGridViewTextBoxColumn();
            dgvc.DataPropertyName = COLUM_SEQ_NO;
            dgvc.HeaderText = "序号";
            dgvc.ValueType = typeof(int);
            dgvc.DefaultCellStyle.Format = "X4";
            dataGridView1.Columns.Add(dgvc);

            dgvc = new DataGridViewTextBoxColumn();
            dgvc.DataPropertyName = COLUM_RSSI;
            dgvc.HeaderText = "RSSI";
            dgvc.ValueType = typeof(int);
            dataGridView1.Columns.Add(dgvc);

            dgvc = new DataGridViewTextBoxColumn();
            dgvc.DataPropertyName = COLUM_VOLT;
            dgvc.HeaderText = "电压";
            dgvc.ValueType = typeof(double);
            dgvc.DefaultCellStyle.Format = "F5";
            dataGridView1.Columns.Add(dgvc);

            dgvc = new DataGridViewTextBoxColumn();
            dgvc.DataPropertyName = COLUM_HOP_COUNT;
            dgvc.HeaderText = "跳数";
            dgvc.ValueType = typeof(int);
            dataGridView1.Columns.Add(dgvc);

            dgvc = new DataGridViewTextBoxColumn();
            dgvc.DataPropertyName = COLUM_TIME_DIFF;
            dgvc.HeaderText = "时延";
            dgvc.ValueType = typeof(int);
            dataGridView1.Columns.Add(dgvc);

            dgvc = new DataGridViewTextBoxColumn();
            dgvc.DataPropertyName = COLUM_GET_TIME;
            dgvc.HeaderText = "时间";
            dgvc.ValueType = typeof(string);
            //dgvc.Width = 120;
            dataGridView1.Columns.Add(dgvc);

            dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.DataSource = dataset.Tables[TABLE_INPUTDATA_NAME];
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            hideStaInfoTitle();
        }
                
        private void initDgvStatistic()      //初始化统计表 界面右边的表
        {
            DataGridViewTextBoxColumn dgvc;
            dgvc = new DataGridViewTextBoxColumn();
            dgvc.DataPropertyName = COLUM_NODE_ID_AVG;
            dgvc.HeaderText = "节点ID";
            dgvc.ValueType = typeof(int);
            dgvc.DefaultCellStyle.Format = "X4";
            dgvc.Width = 50;
            dgvStatistic.Columns.Add(dgvc);

            dgvc = new DataGridViewTextBoxColumn();
            dgvc.DataPropertyName = COLUM_LOSS_PKG_COUNT;
            dgvc.HeaderText = "丢包数量";
            dgvc.ValueType = typeof(int);
            dgvc.Width = 50;
            dgvStatistic.Columns.Add(dgvc);

            dgvc = new DataGridViewTextBoxColumn();
            dgvc.DataPropertyName = COLUM_LOSS_PKG_RATIO;
            dgvc.HeaderText = "丢包率(%)";
            dgvc.ValueType = typeof(double);
            dgvc.Width = 60;
            dgvStatistic.Columns.Add(dgvc);


            dgvc = new DataGridViewTextBoxColumn();
            dgvc.DataPropertyName = COLUM_AVG_RSSI;   //"RSSILV"
            dgvc.HeaderText = "RSSI平均值";
            dgvc.ValueType = typeof(double);
            dgvc.DefaultCellStyle.Format = "F2";
            dgvc.Width = 60;
            dgvStatistic.Columns.Add(dgvc);

            dgvc = new DataGridViewTextBoxColumn();
            dgvc.DataPropertyName = COLUM_VOLT_DIFF;
            dgvc.HeaderText = "节点压降";
            dgvc.ValueType = typeof(double);
            dgvc.DefaultCellStyle.Format = "F5";
            dgvc.Width = 60;
            dgvStatistic.Columns.Add(dgvc);

            dgvc = new DataGridViewTextBoxColumn();
            dgvc.DataPropertyName = COLUM_AVG_HOP;
            dgvc.HeaderText = "平均跳数";
            dgvc.ValueType = typeof(double);
            dgvc.DefaultCellStyle.Format = "F2";
            dgvc.Width = 60;
            dgvStatistic.Columns.Add(dgvc);

            dgvc = new DataGridViewTextBoxColumn();
            dgvc.DataPropertyName = COLUM_AVG_TIME_DIFF;
            dgvc.HeaderText = "平均时延";
            dgvc.ValueType = typeof(double);
            dgvc.DefaultCellStyle.Format = "F2";
            dgvc.Width = 60;
            dgvStatistic.Columns.Add(dgvc);

            dgvc = new DataGridViewTextBoxColumn();
            dgvc.DataPropertyName = COLUM_AVG_PKTCNT;
            dgvc.HeaderText = "接收数量";
            dgvc.ValueType = typeof(int);
            //dgvc.DefaultCellStyle.Format = "F2";
            dgvc.Width = 60;
            dgvStatistic.Columns.Add(dgvc);

            dgvc = new DataGridViewTextBoxColumn();
            dgvc.DataPropertyName = COLUM_AVG_REPCNT;
            dgvc.HeaderText = "重复数量";
            dgvc.ValueType = typeof(int);
            dgvc.Width = 60;
            dgvStatistic.Columns.Add(dgvc);



            dgvStatistic.AutoGenerateColumns = false;
            dgvStatistic.DataSource = dataset.Tables[TABLE_STATISTIC];
        }

        private void initSummView()
        {
            SummaryDGV.Rows.Add(SUMMVIEWROWS);
            SummaryDGV.Rows[(int)SUMMVIEWROWNO.AVGLOSSR].Cells[0].Value = "平均丢包率/%";
            SummaryDGV.Rows[(int)SUMMVIEWROWNO.AVGLOSSR].Cells[1].Style.Format = "F3";
            SummaryDGV.Rows[(int)SUMMVIEWROWNO.AVGRSSI].Cells[0].Value = "平均RSS/dbm";
            SummaryDGV.Rows[(int)SUMMVIEWROWNO.AVGVOLTDEC].Cells[0].Value = "平均压降/V";
            SummaryDGV.Rows[(int)SUMMVIEWROWNO.AVGVOLTDEC].Cells[1].Style.Format = "F5";
            SummaryDGV.Rows[(int)SUMMVIEWROWNO.AVGHOPS].Cells[0].Value = "平均跳数";
            SummaryDGV.Rows[(int)SUMMVIEWROWNO.AVGDELAY].Cells[0].Value = "平均时延/ms";
            SummaryDGV.Rows[(int)SUMMVIEWROWNO.PCKTOTAL].Cells[0].Value = "包总数";
            SummaryDGV.Rows[(int)SUMMVIEWROWNO.PCKTOTAL].Cells[1].Style.Format = "D";
        }

        private void cbSerialPort_DropDown(object sender, EventArgs e)
        {
            cbSerialPort.Items.Clear();
            cbSerialPort.Items.AddRange(SerialPort.GetPortNames());
        }

        private void btCloseSer_Click(object sender, EventArgs e)
        {
            try
            {
                if (isSpOpen)
                {
                    spInstance.Close();
                    if (!spInstance.IsOpen)
                    {
                        isSpOpen = false;
                        btOpenSer.Enabled = true;
                        btCloseSer.Enabled = false;
                        m_state = SYSSTATE.S_IDLE;
                        showSatus("串口已关闭");
                    }
                }
                else
                {
                    showSatus("串口尚未打开.");
                    return;
                }
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
            }
        }
      

        private void cbSerialPort_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbSerialPort.Text != "")
            {
                btOpenSer.Enabled = true;
                btCloseSer.Enabled = false;
            }
            else
            {
                btOpenSer.Enabled = false;
            }
        }

        private void dataGridView1_RowStateChanged(object sender, DataGridViewRowStateChangedEventArgs e)
        {
            e.Row.HeaderCell.Value = string.Format("{0}", e.Row.Index + 1);
        }

        private void dgvStatistic_RowStateChanged(object sender, DataGridViewRowStateChangedEventArgs e)
        {
            e.Row.HeaderCell.Value = string.Format("{0}", e.Row.Index + 1);
        }

        private void tbPkgCount_TextChanged(object sender, EventArgs e)
        {
            try
            {
                m_PacketNum = Convert.ToInt32(tbPkgCount.Text.Trim());
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        
        private void showAboutInfo()
        {
            AboutBox1 newBox=new AboutBox1();
            newBox.ShowDialog(this);
            newBox.Dispose();
        }

        public void DeleteFile(string path)
        {
            FileAttributes attr = File.GetAttributes(path);
            if (attr == FileAttributes.Directory)
            {
                Directory.Delete(path, true);
            }
            else
            {
                File.Delete(path);
            }
        }


        /// <summary>
        /// 在读取数据过程中，向统计表中添加新项目。
        /// </summary>
        public void addNewStatItem()
        {
            DataRow drNew = dataset.Tables[TABLE_STATISTIC].NewRow();
            drNew[COLUM_NODE_ID_AVG] = m_Result.nodeId;
            drNew[COLUM_AVG_INITSN] = m_Result.seqNO;
            drNew[COLUM_AVG_RSSISUM] = m_Result.rssi;
            drNew[COLUM_AVG_INITVOLT] = m_Result.volt;
            drNew[COLUM_AVG_HOPSUM] = m_Result.hopCount;
            drNew[COLUM_AVG_DELAYSUM] = m_Result.timeDifference;
            drNew[COLUM_AVG_PKTCNT] = 1;
            drNew[COLUM_AVG_REPCNT] = 0;
            //Set initial values, in case the input process is broken by a user.
            drNew[COLUM_LOSS_PKG_COUNT] = 0;
            drNew[COLUM_LOSS_PKG_RATIO] = 0;
            drNew[COLUM_AVG_RSSI] = m_Result.rssi;
            drNew[COLUM_VOLT_DIFF] = 0;
            drNew[COLUM_AVG_HOP] = m_Result.hopCount;
            drNew[COLUM_AVG_TIME_DIFF] = m_Result.timeDifference;
            drNew[COLUM_AVG_LASTSN] = m_Result.seqNO;
            try
            {
                dataset.Tables[TABLE_STATISTIC].Rows.Add(drNew);
            }
            catch(Exception ex)
            {
                Console.WriteLine(string.Format("catch unsolve exception: {0}\r\n异常信息：{1}\r\n异常堆栈：{2}", ex.GetType(), ex.Message, ex.StackTrace));
            }
            
        }

        /// <summary>
        /// 更新统计数据。
        /// </summary>
        /// <param name="curRow">在统计表中当前需要更新的行。May remove curRow from dtStatistics </param>
        private void updateStatItem(DataRow curRow)
        {
            //增加COLUM_AVG_REPCNT
            int lastNo, curNo, curCnt;
            lastNo = (int)curRow[COLUM_AVG_LASTSN];
            curNo = m_Result.seqNO;
            curCnt = (int)curRow[COLUM_AVG_PKTCNT];
            int repCnt = (int)curRow[COLUM_AVG_REPCNT];
            
            if (curNo > lastNo)
            {
                curRow[COLUM_LOSS_PKG_COUNT] = curNo - Convert.ToInt32(curRow[COLUM_AVG_INITSN]) - curCnt;
                curRow[COLUM_AVG_LASTSN] = curNo;
                curCnt += 1;
                curRow[COLUM_AVG_PKTCNT] = curCnt;               
            }
            else if (curNo < lastNo)
            {
                if ((lastNo - curNo) > 60000)
                    //循环计数
                    curRow[COLUM_LOSS_PKG_COUNT] = Convert.ToInt32(curRow[COLUM_LOSS_PKG_COUNT]) + curNo + (int)(CONSTDEF.FULLVALUE) - lastNo - 1;
                else
                    //对于多跳传输，延迟接收的情况
                    curRow[COLUM_LOSS_PKG_COUNT] = (int)curRow[COLUM_LOSS_PKG_COUNT] - 1;
                if ((int)curRow[COLUM_LOSS_PKG_COUNT] < 0)
                    //Should never happen!
                    curRow[COLUM_LOSS_PKG_COUNT] = 0;
            }
            

        
            if (curNo != lastNo)//非重复
            {
                curRow[COLUM_AVG_RSSISUM] = (double)curRow[COLUM_AVG_RSSISUM] + m_Result.rssi;
                curRow[COLUM_AVG_HOPSUM] = (double)curRow[COLUM_AVG_HOPSUM] + m_Result.hopCount;
                curRow[COLUM_AVG_DELAYSUM] = (double)curRow[COLUM_AVG_DELAYSUM] + m_Result.timeDifference;
                //Average values.
                curRow[COLUM_LOSS_PKG_RATIO] = Math.Round(Convert.ToDouble(curRow[COLUM_LOSS_PKG_COUNT]) / (curCnt + (int)curRow[COLUM_LOSS_PKG_COUNT]) * 100, 3);
                curRow[COLUM_AVG_RSSI] = Math.Round((double)curRow[COLUM_AVG_RSSISUM] / curCnt, 2);
                curRow[COLUM_AVG_HOP] = Math.Round((double)curRow[COLUM_AVG_HOPSUM] / curCnt, 2);
                curRow[COLUM_AVG_TIME_DIFF] = Math.Round((double)curRow[COLUM_AVG_DELAYSUM] / curCnt, 2);
                curRow[COLUM_VOLT_DIFF] = Math.Round((double)curRow[COLUM_AVG_INITVOLT] - m_Result.volt, 5);
            }
            else
            {
                repCnt += 1;
                curRow[COLUM_AVG_REPCNT] = repCnt;

                if (m_Result.protId == 0x10)
                {
                    //replace rssi with new value
                    int lastVal = 0;
                    if (findLastRSSI(ref lastVal))
                    {
                        lastVal = m_Result.rssi - lastVal;
                        if (lastVal < 0)
                        {
                            showSatus("丢包数量累计异常.");
                        }                            
                        else
                        {
                            curRow[COLUM_AVG_RSSISUM] = (double)curRow[COLUM_AVG_RSSISUM] + lastVal;
                            curRow[COLUM_AVG_RSSI] = Math.Round((double)curRow[COLUM_AVG_RSSISUM] / curCnt, 2);
                        }
                    }
                    else
                    { 
                        showSatus("查找RSSI前值异常。");
                    }                        
                }
            }
        }

        private static DataTable SelectDistinct(DataTable SourceTable, params string[] FieldNames)
        {
            object[] lastValues;
            DataTable newTable;
            DataRow[] orderedRows;

            if (FieldNames == null || FieldNames.Length == 0)
                throw new ArgumentNullException("FieldNames");

            lastValues = new object[FieldNames.Length];
            newTable = new DataTable();

            foreach (string fieldName in FieldNames)
                newTable.Columns.Add(fieldName, SourceTable.Columns[fieldName].DataType);

            orderedRows = SourceTable.Select("", string.Join(", ", FieldNames));

            foreach (DataRow row in orderedRows)
            {
                if (!fieldValuesAreEqual(lastValues, row, FieldNames))
                {
                    newTable.Rows.Add(createRowClone(row, newTable.NewRow(), FieldNames));
                    setLastValues(lastValues, row, FieldNames);
                }
            }

            return newTable;
        }

        private static bool fieldValuesAreEqual(object[] lastValues, DataRow currentRow, string[] fieldNames)
        {
            bool areEqual = true;
            for (int i = 0; i < fieldNames.Length; i++)
            {
                if (lastValues[i] == null || !lastValues[i].Equals(currentRow[fieldNames[i]]))
                {
                    areEqual = false;
                    break;
                }
            }
            return areEqual;
        }

        private static DataRow createRowClone(DataRow sourceRow, DataRow newRow, string[] fieldNames)
        {
            foreach (string field in fieldNames)
                newRow[field] = sourceRow[field];
            return newRow;
        }

        private static void setLastValues(object[] lastValues, DataRow sourceRow, string[] fieldNames)
        {
            for (int i = 0; i < fieldNames.Length; i++)
                lastValues[i] = sourceRow[fieldNames[i]];
        }


        private bool isRepeat(DataRow curRow, int seqNo)//时间复杂度
        {
            DataRow[] dr = dtOri.Select(string.Format("{0}={1} and {2}={3}", COLUM_NODE_ID, curRow[COLUM_NODE_ID_AVG], COLUM_SEQ_NO, seqNo));
            if(dr.Length>=1)
            { 
                return true;
            }
            return false;
        }


        private bool findLastRSSI(ref int lastValue)
        {
            bool retVal = false;
            for (int i =dtOri.Rows.Count-1;i>=0;i--)
            {
                if((int)dtOri.Rows[i][COLUM_NODE_ID]==m_Result.nodeId)
                {
                    lastValue = (int)dtOri.Rows[i][COLUM_RSSI];
                    retVal = true;
                    break;
                }
            }
            return retVal;
        }



        private void updateVoltage(DataRow lastRow,double lastVolt)
        {
            lastRow[COLUM_VOLT_DIFF] = Math.Round((double)lastRow[COLUM_AVG_INITVOLT] - lastVolt,5);
        }

        private void 关于ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            showAboutInfo();
        }

        /// <summary>
        /// 显示状态信息
        /// </summary>
        private void showSatus(string statusInfo)
        {
            PromptLabel.Text = statusInfo;
        }

        /// <summary>
        /// 启动新测试。
        /// </summary>
        private void startTest()
        {
            if(m_bDate4Save)
            {
                if (MessageBox.Show("前次数据尚未保存，如果继续，数据将丢失，要继续吗？", "重要提示", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
                    return;
            }
            if (btOpenSer.Enabled == false)
            {
                if (!spInstance.IsOpen)
                {
                    showSatus("请打开串口.");
                }
                else
                {
                    switchStatus(1);
                    resetDataSet();
                    RecoEn = true;
                    m_state = SYSSTATE.S_START;
                    showSatus("正在输入数据...");
                    try
                    {
                        readThread.Interrupt();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
            }
        }

        private void StartTestMenuItem_Click(object sender, EventArgs e)
        {
            startTest();
        }

        private void StopTestMenuItem_Click(object sender, EventArgs e)
        {
            stopTest();
            
        }

        /// <summary>
        /// 退出系统
        /// </summary>
        private bool exitSys()
        {
            if (MessageBox.Show("确认要退出吗？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                if (m_bDate4Save)
                {
                    if (MessageBox.Show("测试数据未保存，需要在退出前保存数据吗？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        saveCurrentData();
                    }
                }
                //Processing befor exit.
                try
                {
                    readThread.Abort();
                    readThread.Join();
                    if (spInstance.IsOpen)
                        spInstance.Close();
                    dtOri.Dispose();
                    dtStatistics.Dispose();
                    m_dtSummary.Dispose();
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                }

            }
            else
            {
                return false;
            }
            return true;
        }

        private void ExitSYSButton_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void SaveDFileMenuItem_Click(object sender, EventArgs e)
        {
            saveCurrentData();
        }

        private void SaveDFileButton_Click(object sender, EventArgs e)
        {
            saveCurrentData();
        }

        /// <summary>
        /// 切换系统状态，控制界面控件的使能状态。
        /// </summary>
        /// <param name="iStatus">系统状态，0：停止测试；1：开始测试</param>
        private void switchStatus(int iStatus)
        {
            switch (iStatus)
            {
                case 0:
                    StartButton.Enabled = true;
                    StartTestMenuItem.Enabled = true;
                    StopButton.Enabled = false;
                    StopTestMenuItem.Enabled = false;
                    btCloseSer.Enabled = true;
                    break;
                case 1:
                    StartButton.Enabled = false;
                    StartTestMenuItem.Enabled = false;
                    StopButton.Enabled = true;
                    StopTestMenuItem.Enabled = true;
                    btCloseSer.Enabled = false;
                    break;
            }
            this.Refresh();
        }

        /// <summary>
        /// 停止当前测试。
        /// </summary>
        private void stopTest()
        {
            if (isSpOpen)
            {
                RecoEn = false;
                switchStatus(0);
                m_state = SYSSTATE.S_STANDBY;
                showSatus("测试结束.");
            }
        }

        private void StopButton_Click(object sender, EventArgs e)
        {
            doStatistics();
            stopTest();
        }

        private void StartButton_Click(object sender, EventArgs e)
        {
            startTest();
        }

        /// <summary>
        /// 更新状态栏内的显示信息
        /// </summary>
        private void updateStatusInfo()
        {
            if(!m_bHasInput)
            {
                m_bHasInput = true;
                showStaInfoTitle();
            }
            TimeVauleLbl.Text = m_Result.time;
            CurrNodeIDLbl.Text = m_Result.nodeId.ToString("X4");
            PckCntValueLbl.Text = m_PkgCount.ToString();
            NodeCntValueLbl.Text = dtStatistics.Rows.Count.ToString();
        }

        private void showStaInfoTitle()
        {
            PckCntLabel.Visible = true;
            NodeCntLabel.Visible = true;
            CurrNodeLbl.Visible = true;
            TimeLabel.Visible = true;
        }

        private void hideStaInfoTitle()
        {
            PckCntLabel.Visible = false;
            NodeCntLabel.Visible = false;
            CurrNodeLbl.Visible = false;
            TimeLabel.Visible = false;
        }

        /// <summary>
        /// 对当前数据进行统计。
        /// </summary>
        private void doStatistics()
        {
            //int iNdIdNum = dtStatistics.Rows.Count;
            double avgLossR = 0;
            double avgRssi = 0;
            double avgVoltDec=0;
            double avgHop = 0;
            double avgDelay = 0;
            int pckTotal = 0;
            int lostCnt = 0;
            int repCnt = 0;
            int recPkg = 0;

            double minVolt = 0;
            double maxVolt = 0;


            // sql数据源
            string selectStr = "SELECT DISTINCT " + COLUM_NODE_ID + " FROM " + TABLE_INPUTDATA_NAME;
            DataTable datadistinct = new AccessOperate(m_connStr).MyExecuteDataSet(selectStr).Tables[0];
            int tableLen = dtStatistics.Rows.Count;
            datadistinct.Dispose();


            DataTable[] diffdtOri = new DataTable[tableLen];//空表
            selectStr = "SELECT " + COLUM_NODE_ID + ", " + COLUM_RSSI + ", " + COLUM_SEQ_NO + ", " + COLUM_GET_TIME + ", " + COLUM_VOLT + ", " + COLUM_HOP_COUNT + ", " + COLUM_TIME_DIFF + " FROM " + TABLE_INPUTDATA_NAME;
            DataTable datasrc = new AccessOperate(m_connStr).MyExecuteDataSet(selectStr).Tables[0];


            for (int k = 0; k < tableLen; k++)
            {
                DataRow[] dtmp = datasrc.Select(string.Format(COLUM_NODE_ID + "='{0}'", dtStatistics.Rows[k][COLUM_NODE_ID]), COLUM_SEQ_NO + " ASC");
                diffdtOri[k] = datasrc.Clone();
                foreach (DataColumn col in diffdtOri[k].Columns)
                {
                    if (col.ColumnName == COLUM_RSSI || col.ColumnName == COLUM_HOP_COUNT)
                    {
                        col.DataType = typeof(Single);
                    }
                }
                diffdtOri[k] = ToDataTable(dtmp);//已排序
            }




            for (int k = 0; k < tableLen; k++)
            {
                int avgId = Convert.ToInt32(diffdtOri[k].Rows[0][COLUM_NODE_ID]);
                selectStr = string.Format("{0}='{1}'", COLUM_NODE_ID, avgId);                

                //DataRow[] drSeq = diffdtOri[k].Select(selectStr, COLUM_SEQ_NO + " ASC");
                lostCnt = 0;
                repCnt = 0;


                for (int i = 0; i < diffdtOri[k].Rows.Count - 1; i++)
                {                    
                    if (Convert.ToInt32(diffdtOri[k].Rows[i][COLUM_SEQ_NO]) == Convert.ToInt32(diffdtOri[k].Rows[i + 1][COLUM_SEQ_NO]))// 有重复
                    {
                        repCnt += 1;
                    }
                    else if (diffdtOri[k].Rows.Count == m_PacketNum && Convert.ToInt32(diffdtOri[k].Rows[i + 1][COLUM_SEQ_NO]) > Convert.ToInt32(diffdtOri[k].Rows[i][COLUM_SEQ_NO]))
                    {
                        lostCnt += Convert.ToInt32(diffdtOri[k].Rows[i + 1][COLUM_SEQ_NO]) - Convert.ToInt32(diffdtOri[k].Rows[i][COLUM_SEQ_NO]) - 1;
                    }

                }


                DataTable distinctTable = SelectDistinct(diffdtOri[k],COLUM_SEQ_NO);
                recPkg = distinctTable.Rows.Count;
                if (diffdtOri[k].Rows.Count > m_PacketNum)
                {
                    lostCnt = m_PacketNum - recPkg;
                }                
                distinctTable.Dispose();
                maxVolt = Convert.ToDouble(diffdtOri[k].Compute("Max(" + COLUM_VOLT + ")", selectStr));
                minVolt = Convert.ToDouble(diffdtOri[k].Compute("Min(" + COLUM_VOLT + ")", selectStr));

                pckTotal = lostCnt + recPkg + repCnt;//==m_PacketNum
                avgLossR = Convert.ToDouble(lostCnt) / Convert.ToDouble(m_PacketNum/*recPkg + lostCnt*/) * 100.0;
                avgRssi = Convert.ToDouble(diffdtOri[k].Compute("Avg(" + COLUM_RSSI + ")", selectStr));
                avgVoltDec = maxVolt - minVolt;
                avgHop = Convert.ToDouble(diffdtOri[k].Compute("Avg(" + COLUM_HOP_COUNT + ")", selectStr));
                avgDelay = Convert.ToDouble(diffdtOri[k].Compute("Avg(" + COLUM_TIME_DIFF + ")", selectStr));
                
                

                DataRow drNew = m_dtSummary.NewRow();
                drNew[COLUM_SUM_LOSSCNT] = lostCnt;
                drNew[COLUM_SUM_AVGLOSSRATIO] = avgLossR;
                drNew[COLUM_SUM_AVGRSSI] = avgRssi;
                drNew[COLUM_SUM_AVGVOLTDEC] = avgVoltDec;
                drNew[COLUM_SUM_AVGHOP] = avgHop;
                drNew[COLUM_SUM_AVGDELAY] = avgDelay;
                drNew[COLUM_SUM_PCKCNT] = recPkg;
                drNew[COLUM_SUM_REPCNT] = repCnt;
                drNew[COLUM_SUM_TBNAME] = "";
                m_dtSummary.Rows.Add(drNew);
                updateSummView(drNew);
            }
            m_bDate4Save = true;
            showSatus("计算完成");
        }

        private void updateSummView(DataRow curRow)
        {
            SummaryDGV.Rows[(int)SUMMVIEWROWNO.AVGLOSSR].Cells[1].Value = curRow[COLUM_SUM_AVGLOSSRATIO];
            SummaryDGV.Rows[(int)SUMMVIEWROWNO.AVGRSSI].Cells[1].Value = curRow[COLUM_SUM_AVGRSSI];
            SummaryDGV.Rows[(int)SUMMVIEWROWNO.AVGVOLTDEC].Cells[1].Value = curRow[COLUM_SUM_AVGVOLTDEC];
            SummaryDGV.Rows[(int)SUMMVIEWROWNO.AVGHOPS].Cells[1].Value = curRow[COLUM_SUM_AVGHOP];
            SummaryDGV.Rows[(int)SUMMVIEWROWNO.AVGDELAY].Cells[1].Value = curRow[COLUM_SUM_AVGDELAY];
            SummaryDGV.Rows[(int)SUMMVIEWROWNO.PCKTOTAL].Cells[1].Value = curRow[COLUM_SUM_PCKCNT];
        }

        /// <summary>
        /// 检查是否完成测试，如果完成，则进行统计和结束处理。
        /// </summary>
        /// <param name="iTotal">接收到数据的节点总数</param>
        /// <param name="iDone">已完成的数量</param>
        private bool checkFinish(int iTotal, int iDone)
        {            
            if (iTotal == iDone)
            {
                this.Invoke(new MethodInvoker(delegate
                {
                    doStatistics();
                    stopTest();
                }));                
                return true;
            }
            return false;
        }

        public void resetDataSet()
        {
            dataset.Tables[TABLE_INPUTDATA_NAME].Clear();
            dataset.Tables[TABLE_STATISTIC].Clear();
            m_DoneCntr = 0;
            m_bDate4Save = false;
        }

        private void ExitMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void PauseButn_Click(object sender, EventArgs e)
        {
            m_bPauseScroll = PauseButn.Checked;
        }

        /// <summary>
        /// 保存当前测试结果。
        /// </summary>
        private void saveCurrentData()
        {
            if(m_state==SYSSTATE.S_START)
            {
                MessageBox.Show("正在测试，请稍后再试。","提示",MessageBoxButtons.OK,MessageBoxIcon.Information);
                return;
            }
            if (m_bDate4Save)
            {
                //strDBName = "WSN_" + CurrentTime() + ".accdb";
                //System.IO.File.Copy(strCurrentPath + "\\" + strTemplateDb, dest_dbPath + "\\" + strDBName, true);          //根据模板创建db文件
                m_connStr = string.Format("Provider={0}; Data Source={1};", strProvider, dest_dbPath + "\\" + strDBName);

                if (saveResults())
                {
                    //update summmary field of table name
                    updateSummTbName();
                    m_bDate4Save = false;
                }
            }
            else
                showSatus("无数据需要保存或数据已经保存。");
        }

        /// <summary>
        /// 更新汇总表中数据库名字段。
        /// </summary>
        private void updateSummTbName()
        {
            int curIdx = m_dtSummary.Rows.Count - 1;
            m_dtSummary.Rows[curIdx][COLUM_SUM_TBNAME] = strDBName;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if(!exitSys())
                e.Cancel = true;
        }

        private void NewDFileButton_Click(object sender, EventArgs e)
        {
            setNewDBFolder();
        }

        /// <summary>
        /// 设置新的数据存储目录。
        /// </summary>
        private void setNewDBFolder()
        {
            if(m_state==SYSSTATE.S_START)
            {
                MessageBox.Show("正在输入数据，请稍后再试。","提示",MessageBoxButtons.OK,MessageBoxIcon.Information);
                return;
            }
            NewFileBrowser.RootFolder = Environment.SpecialFolder.MyComputer;
            if (NewFileBrowser.ShowDialog() == DialogResult.OK)
            {
                dest_dbPath = NewFileBrowser.SelectedPath;
            }
        }

        private void OpenDFileButton_Click(object sender, EventArgs e)
        {
            openDBFile();
        }

        /// <summary>
        /// 打开数据库文件。
        /// </summary>
        private void openDBFile()
        {
            openDBFileDialog.InitialDirectory = dest_dbPath;
            if(openDBFileDialog.ShowDialog()==DialogResult.OK)
            {
                Process MyProcess = new Process();
                MyProcess.StartInfo.FileName = openDBFileDialog.FileName;
                MyProcess.StartInfo.Verb = "Open";
                MyProcess.StartInfo.CreateNoWindow = true;
                MyProcess.Start();
            }
        }

        private void NewDFileMenuItem_Click(object sender, EventArgs e)
        {
            setNewDBFolder();
        }

        private void OpenDFileMenuItem_Click(object sender, EventArgs e)
        {
            openDBFile();
        }
    }    

}