using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI;
using NPOI.HPSF;
using NPOI.HSSF;
using NPOI.HSSF.UserModel;
using NPOI.POIFS;
using NPOI.Util;
using NPOI.HSSF.Util;
using NPOI.HSSF.Extractor;
using System.IO;
using System.Data.SqlClient;
using NPOI.SS.UserModel;
using System.Configuration;
using NPOI.XSSF.UserModel;
using FastReport;
using FastReport.Data;

namespace TKMOC
{
    public partial class frmMOCNO : Form
    {

        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        SqlDataAdapter adapter2= new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder2 = new SqlCommandBuilder();
        SqlDataAdapter adapter3 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder3 = new SqlCommandBuilder();
        SqlDataAdapter adapter4 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder4 = new SqlCommandBuilder();
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();
        DataTable dt = new DataTable();
        SqlTransaction tran;
        int result;

        string tablename = null;
        int rownum = 0;

        string OLDTA001;
        string OLDTA002;
        string NEWTA001;
        string NEWTA002;
        string ORINO;
        string BEFORENO;
        string AFTERNO;

        public frmMOCNO()
        {
            InitializeComponent();

            comboBox1load();
        }

        #region FUNCTION

        private void frmMOCNO_Load(object sender, EventArgs e)
        {
            dataGridView3.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;      //奇數列顏色

            //先建立個 CheckBox 欄
            DataGridViewCheckBoxColumn cbCol = new DataGridViewCheckBoxColumn();
            cbCol.Width = 50;   //設定寬度
            cbCol.HeaderText = "選擇";
            cbCol.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;   //置中
            cbCol.TrueValue = true;
            cbCol.FalseValue = false;
            dataGridView3.Columns.Insert(0, cbCol);

            #region 建立全选 CheckBox

            ////建立个矩形，等下计算 CheckBox 嵌入 GridView 的位置
            //Rectangle rect = dataGridView3.GetCellDisplayRectangle(0, -1, true);
            //rect.X = rect.Location.X + rect.Width / 4 - 18;
            //rect.Y = rect.Location.Y + (rect.Height / 2 - 9);

            //CheckBox cbHeader = new CheckBox();
            //cbHeader.Name = "checkboxHeader";
            //cbHeader.Size = new Size(18, 18);
            //cbHeader.Location = rect.Location;

            ////全选要设定的事件
            //cbHeader.CheckedChanged += new EventHandler(cbHeader_CheckedChanged);

            ////将 CheckBox 加入到 dataGridView
            //dataGridView3.Controls.Add(cbHeader);


            #endregion
        }

        public void comboBox1load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MD001,MD002 FROM [TK].dbo.CMSMD    WHERE MD002 LIKE '新廠%' ORDER BY MD001  ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MD001", typeof(string));
            dt.Columns.Add("MD002", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "MD001";
            comboBox1.DisplayMember = "MD002";
            sqlConn.Close();


        }

        public void SEARCH()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();
                              
                sbSql.AppendFormat(@"  SELECT TA001 AS '製令',TA002 AS '單號',TA003 AS '生產日',TA006 AS '品號',TA034 AS '品名',TA015 AS '生產量',TA007 AS '單位',TA021 AS '線別',TA026 AS '訂單',TA027 AS '單號',TA028 AS '序號'");
                sbSql.AppendFormat(@"  FROM [test].dbo.MOCTA");
                sbSql.AppendFormat(@"  WHERE TA003>='{0}' AND TA003<='{1}' ",dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  ORDER BY TA001,TA002,TA003");
                sbSql.AppendFormat(@"  ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds");
                sqlConn.Close();


                if (ds.Tables["TEMPds"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds.Tables["TEMPds"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView1.DataSource = ds.Tables["TEMPds"];
                        dataGridView1.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                    }
                }

            }
            catch
            {

            }
            finally
            {

            }
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            SETNULL();

            DateTime dt = new DateTime();
            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    dt = Convert.ToDateTime(row.Cells["生產日"].Value.ToString().Substring(0,4)+"/"+ row.Cells["生產日"].Value.ToString().Substring(4, 2) + "/" + row.Cells["生產日"].Value.ToString().Substring(6, 2));
                    dateTimePicker3.Value = dt;
                    textBox1.Text = row.Cells["製令"].Value.ToString().Trim();
                    textBox2.Text = row.Cells["單號"].Value.ToString().Trim();
                    OLDTA001 = row.Cells["製令"].Value.ToString().Trim();
                    OLDTA002 = row.Cells["單號"].Value.ToString().Trim();

                    SEARCHMOCNO();

                }
                else
                {
                                     

                }
            }
        }

        public void SEARCHMOCNO()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

               
                sbSql.AppendFormat(@"  SELECT  [ORINO] AS '最初單' ,[BEFORENO]  AS '舊單',[AFTERNO] AS '新單',[DTIMES]  AS '時間'");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCNO]");
                sbSql.AppendFormat(@"  WHERE [ORINO] IN (SELECT [ORINO] FROM [TKMOC].[dbo].[MOCNO] WHERE [AFTERNO]='{0}')",OLDTA001+OLDTA002);
                sbSql.AppendFormat(@"  ORDER BY DTIMES");
                sbSql.AppendFormat(@"  ");

                adapter3 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder3 = new SqlCommandBuilder(adapter3);
                sqlConn.Open();
                ds3.Clear();
                adapter3.Fill(ds3, "TEMPds3");
                sqlConn.Close();


                if (ds3.Tables["TEMPds3"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else
                {
                    if (ds3.Tables["TEMPds3"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView2.DataSource = ds3.Tables["TEMPds3"];
                        dataGridView2.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                    }
                }

            }
            catch
            {

            }
            finally
            {

            }
        }

        public void CHANGEMOCTAMOCTB()
        {
            if (!string.IsNullOrEmpty(NEWTA001)&& !string.IsNullOrEmpty(NEWTA002))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();

                   
                    sbSql.AppendFormat(" UPDATE [test].dbo.MOCTA SET TA001='{0}',TA002='{1}',TA003='{2}',TA009='{2}',TA010='{2}'", NEWTA001, NEWTA002,dateTimePicker4.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(" WHERE TA001='{0}' AND TA002='{1}'", OLDTA001, OLDTA002);
                    sbSql.AppendFormat(" ");
                    sbSql.AppendFormat(" UPDATE [test].dbo.MOCTB SET TB001='{0}',TB002='{1}'", NEWTA001, NEWTA002);
                    sbSql.AppendFormat(" WHERE TB001='{0}' AND TB002='{1}'", OLDTA001, OLDTA002);
                    sbSql.AppendFormat(" ");
               

                    cmd.Connection = sqlConn;
                    cmd.CommandTimeout = 60;
                    cmd.CommandText = sbSql.ToString();
                    cmd.Transaction = tran;
                    result = cmd.ExecuteNonQuery();

                    if (result == 0)
                    {
                        tran.Rollback();    //交易取消


                    }
                    else
                    {
                        tran.Commit();      //執行交易  
                        ADDMOCNO();

                    }

                }
                catch
                {

                }

                finally
                {
                    sqlConn.Close();
                }
            }
           
        }

        public void ADDMOCNO()
        {
            SERACHMOCNO();

            if(string.IsNullOrEmpty(ORINO))
            {
                INSERTMOCNO(OLDTA001 + OLDTA002, OLDTA001+OLDTA002,NEWTA001+NEWTA002);
            }
            else if(!string.IsNullOrEmpty(ORINO))
            {
                INSERTMOCNO(ORINO, OLDTA001 + OLDTA002, NEWTA001 + NEWTA002);
            }
        }
        public string SERACHMOCNO()
        {
            ORINO = null;

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT TOP 1 [ORINO] ,[BEFORENO] ,[AFTERNO]");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCNO]");
                sbSql.AppendFormat(@"  WHERE [AFTERNO]='{0}'", OLDTA001+ OLDTA002);
                sbSql.AppendFormat(@"  ORDER BY DTIMES");
                sbSql.AppendFormat(@"  ");

                adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                sqlConn.Open();
                ds2.Clear();
                adapter2.Fill(ds2, "TEMPds2");
                sqlConn.Close();


                if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                {
                    
                }
                else
                {
                    if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                    {
                        ORINO= ds2.Tables["TEMPds2"].Rows[0]["ORINO"].ToString();
                        
                    }
                }

            }
            catch
            {

            }
            finally
            {

            }
            return null;
        }

        public void INSERTMOCNO(string ORINO, string BEFORENO, string AFTERNO)
        {
            if (!string.IsNullOrEmpty(ORINO) && !string.IsNullOrEmpty(BEFORENO) && !string.IsNullOrEmpty(AFTERNO))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();

                    sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[MOCNO] ([ORINO] ,[BEFORENO] ,[AFTERNO],[DTIMES]) VALUES('{0}','{1}','{2}',Getdate() )",ORINO,BEFORENO,AFTERNO);


                    cmd.Connection = sqlConn;
                    cmd.CommandTimeout = 60;
                    cmd.CommandText = sbSql.ToString();
                    cmd.Transaction = tran;
                    result = cmd.ExecuteNonQuery();

                    if (result == 0)
                    {
                        tran.Rollback();    //交易取消


                    }
                    else
                    {
                        tran.Commit();      //執行交易  
                        

                    }

                }
                catch
                {

                }

                finally
                {
                    sqlConn.Close();
                }
            }

        }
        private void dateTimePicker4_ValueChanged(object sender, EventArgs e)
        {
            SETNULL();

            textBox4.Text=GETMAXNO();

            NEWTA001 = OLDTA001;
            textBox3.Text = NEWTA001;
        }
        public void SETNULL()
        {
            NEWTA001 = null;
            NEWTA002 = null;
            textBox3.Text = null;
            textBox4.Text = null;
        }
        public string GETMAXNO()
        {
           

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds1.Clear();

                sbSql.AppendFormat(@"  SELECT ISNULL(MAX(TA002),'00000000000') AS TA002");
                sbSql.AppendFormat(@"  FROM [test].[dbo].[MOCTA] ");
                //sbSql.AppendFormat(@"  WHERE  TC001='{0}' AND TC003='{1}'", "A542","20170119");
                sbSql.AppendFormat(@"  WHERE  TA001='{0}' AND TA003='{1}'", OLDTA001, dateTimePicker4.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                {
                    return null;
                }
                else
                {
                    if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        NEWTA002 = SETTA002(ds1.Tables["TEMPds1"].Rows[0]["TA002"].ToString());
                      
                        return NEWTA002;

                    }
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public string SETTA002(string TA002)
        {

            if (TA002.Equals("00000000000"))
            {
                return dateTimePicker4.Value.ToString("yyyyMMdd") + "001";
            }

            else
            {
                int serno = Convert.ToInt16(TA002.Substring(8, 3));
                serno = serno + 1;
                string temp = serno.ToString();
                temp = temp.PadLeft(3, '0');
                return dateTimePicker4.Value.ToString("yyyyMMdd") + temp.ToString();
            }
        }

        public void SEARCHMULTI()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT TA001 AS '製令',TA002 AS '單號',TA003 AS '生產日',TA006 AS '品號',TA034 AS '品名',TA015 AS '生產量',TA007 AS '單位',TA021 AS '線別',TA026 AS '訂單',TA027 AS '單號',TA028 AS '序號'");
                sbSql.AppendFormat(@"  FROM [test].dbo.MOCTA");
                sbSql.AppendFormat(@"  WHERE TA003>='{0}' AND TA003<='{1}' ", dateTimePicker5.Value.ToString("yyyyMMdd"), dateTimePicker6.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND TA021='{0}' ",comboBox1.SelectedValue.ToString());
                sbSql.AppendFormat(@"  ORDER BY TA001,TA002,TA003");
                sbSql.AppendFormat(@"  ");

                adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder4 = new SqlCommandBuilder(adapter4);
                sqlConn.Open();
                ds4.Clear();
                adapter4.Fill(ds4, "TEMPds4");
                sqlConn.Close();


                if (ds4.Tables["TEMPds4"].Rows.Count == 0)
                {
                    dataGridView3.DataSource = null;
                }
                else
                {
                    if (ds4.Tables["TEMPds4"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView3.DataSource = ds4.Tables["TEMPds4"];
                        dataGridView3.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                    }
                }

            }
            catch
            {

            }
            finally
            {

            }
        }

        private void cbHeader_CheckedChanged(object sender, EventArgs e)
        {
            foreach (DataGridViewRow dr in dataGridView3.Rows)
            {
                dr.Cells[0].Value = ((CheckBox)dataGridView3.Controls.Find("checkboxHeader", true)[0]).Checked;

            }

        }

        public void CHANGEMULTI()
        {
            foreach (DataGridViewRow dr in this.dataGridView3.Rows)
            {
                if (dr.Cells[0].Value != null && (bool)dr.Cells[0].Value)
                {
                    OLDTA001= ((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString();
                    OLDTA002 = ((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString();

                    //MessageBox.Show(OLDTA001+"-"+ OLDTA002);
                }
                else
                {
                    OLDTA001 = null;
                    OLDTA002 = null;
                }
            }

        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCH();
        }

        private void button2_Click(object sender, EventArgs e)
        {         

            DialogResult dialogResult = MessageBox.Show("要修改嗎?", "要修改嗎?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                CHANGEMOCTAMOCTB();
                SEARCH();

            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SEARCHMULTI();
        }



        #endregion

        private void button4_Click(object sender, EventArgs e)
        {
            CHANGEMULTI();
        }
    }
}
