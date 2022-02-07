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
using TKITDLL;

namespace TKMOC
{
    public partial class frmMOCCHANGE : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlDataAdapter adapter2 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder2 = new SqlCommandBuilder();
        SqlDataAdapter adapter3 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder3 = new SqlCommandBuilder();
        SqlDataAdapter adapter4= new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder4 = new SqlCommandBuilder();
        SqlDataAdapter adapter5 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder5 = new SqlCommandBuilder();

        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();
        DataSet ds5 = new DataSet();

        DataSet dsCHECKMOCTDMOCTG = new DataSet();
        DataTable dt = new DataTable();
        SqlTransaction tran;
        int result;

        string tablename = null;
        int rownum = 0;

        string TA001;
        string TA002;
        string OLDMB001;
        string NEWMB001;

        public frmMOCCHANGE()
        {
            InitializeComponent();

            comboBox1load();
            comboBox1load2("");
            comboBox1load3();
        }

        private void frmMOCCHANGE_Load(object sender, EventArgs e)
        {
            //dataGridView1
            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;      //奇數列顏色

            //先建立個 CheckBox 欄
            DataGridViewCheckBoxColumn cbCol = new DataGridViewCheckBoxColumn();
            cbCol.Width = 50;   //設定寬度
            cbCol.HeaderText = "選擇";
            cbCol.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;   //置中
            cbCol.TrueValue = true;
            cbCol.FalseValue = false;
            dataGridView1.Columns.Insert(0, cbCol);

            //dataGridView3
            dataGridView3.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;      //奇數列顏色

            //先建立個 CheckBox 欄
            DataGridViewCheckBoxColumn cbCol3 = new DataGridViewCheckBoxColumn();
            cbCol3.Width = 50;   //設定寬度
            cbCol3.HeaderText = "選擇";
            cbCol3.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;   //置中
            cbCol3.TrueValue = true;
            cbCol3.FalseValue = false;
            dataGridView3.Columns.Insert(0, cbCol3);

            //dataGridView5
            dataGridView5.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;      //奇數列顏色

            //先建立個 CheckBox 欄
            DataGridViewCheckBoxColumn cbCol5 = new DataGridViewCheckBoxColumn();
            cbCol5.Width = 50;   //設定寬度
            cbCol5.HeaderText = "選擇";
            cbCol5.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;   //置中
            cbCol5.TrueValue = true;
            cbCol5.FalseValue = false;
            dataGridView5.Columns.Insert(0, cbCol5);

            //先建立個 CheckBox 欄
            DataGridViewCheckBoxColumn cbCol7 = new DataGridViewCheckBoxColumn();
            cbCol7.Width = 50;   //設定寬度
            cbCol7.HeaderText = "選擇";
            cbCol7.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;   //置中
            cbCol7.TrueValue = true;
            cbCol7.FalseValue = false;
            dataGridView7.Columns.Insert(0, cbCol7);

            //先建立個 CheckBox 欄
            DataGridViewCheckBoxColumn cbCol9 = new DataGridViewCheckBoxColumn();
            cbCol9.Width = 50;   //設定寬度
            cbCol9.HeaderText = "選擇";
            cbCol9.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;   //置中
            cbCol9.TrueValue = true;
            cbCol9.FalseValue = false;
            dataGridView9.Columns.Insert(0, cbCol9);
        }

        #region FUNCTION
        public void comboBox1load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@" SELECT BOMMB.MB001,RTRIM(LTRIM(BOMMB.MB001))+' '+INVMB.MB002 AS MB002 FROM [TK].dbo.BOMMB,[TK] .dbo.INVMB WHERE BOMMB.MB001=INVMB.MB001 AND INVMB.MB002 NOT LIKE '%停%' AND  (BOMMB.MB001 LIKE '1%' OR BOMMB.MB001 LIKE '208%' ) GROUP BY BOMMB.MB001,INVMB.MB002 ORDER BY BOMMB.MB001 ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MB001", typeof(string));
            dt.Columns.Add("MB002", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "MB001";
            comboBox1.DisplayMember = "MB002";
            sqlConn.Close();


        }

        public void comboBox1load2(string MB001)
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT BOMMB.MB004,RTRIM(LTRIM(BOMMB.MB004))+' '+INVMB.MB002 AS MB002 FROM [TK].dbo.BOMMB,[TK] .dbo.INVMB WHERE BOMMB.MB004=INVMB.MB001  AND (BOMMB.MB001 LIKE '1%' OR BOMMB.MB001 LIKE '208%' ) AND BOMMB.MB001='{0}' GROUP BY BOMMB.MB004,RTRIM(LTRIM(BOMMB.MB004))+' '+INVMB.MB002 ", MB001);
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MB004", typeof(string));
            dt.Columns.Add("MB002", typeof(string));
            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "MB004";
            comboBox2.DisplayMember = "MB002";
            sqlConn.Close();


        }

        public void comboBox1load3()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MC001,MC002 FROM [TK].dbo.CMSMC    WHERE MC001 IN ('20006','20004','20005') ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MC001", typeof(string));
            dt.Columns.Add("MC002", typeof(string));
            da.Fill(dt);
            comboBox3.DataSource = dt.DefaultView;
            comboBox3.ValueMember = "MC001";
            comboBox3.DisplayMember = "MC001";
            sqlConn.Close();


        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox1load2(comboBox1.SelectedValue.ToString().Trim());
        }


        public void SEARCH()
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  
                                    SELECT *
                                    FROM (
                                    SELECT TA001 AS '製令',TA002 AS '單號',TA003 AS '生產日',TA006 AS '品號',TA034 AS '品名',TA015 AS '生產量',TA007 AS '單位',MD002 AS '線別',TA026 AS '訂單',TA027 AS '訂單單號',TA028 AS '序號'
                                    FROM [TK].dbo.MOCTA,[TK].dbo.CMSMD
                                    WHERE TA021=MD001
                                    AND TA003>='{0}' AND TA003<='{1}' 

                                    UNION ALL 
                                    SELECT TA001 AS '製令',TA002 AS '單號',TA003 AS '生產日',TA006 AS '品號',TA034 AS '品名',TA015 AS '生產量',TA007 AS '單位','' AS '線別',TA026 AS '訂單',TA027 AS '訂單單號',TA028 AS '序號'
                                    FROM [TK].dbo.MOCTA
                                    WHERE TA001='A512'
                                    AND TA003>='{0}' AND TA003<='{1}' 
                                    ) AS TEMP 
                                    ORDER BY 製令,單號,生產日
                                    ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));

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

        private void cbHeader_CheckedChanged(object sender, EventArgs e)
        {
            foreach (DataGridViewRow dr in dataGridView1.Rows)
            {
                dr.Cells[0].Value = ((CheckBox)dataGridView1.Controls.Find("checkboxHeader", true)[0]).Checked;

            }

        }
        public void CHANGEMULTI()
        {
            OLDMB001 = comboBox1.SelectedValue.ToString().Trim();
            NEWMB001 = comboBox2.SelectedValue.ToString().Trim();

            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {
                DataGridViewCheckBoxCell cbx = (DataGridViewCheckBoxCell)dr.Cells[0];

                if ((bool)cbx.FormattedValue)
                {
                    TA001 = ((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString();
                    TA002 = ((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString();

                    //MessageBox.Show(TA001 + "-"+ TA002);
                    if (!string.IsNullOrEmpty(TA001) && !string.IsNullOrEmpty(TA002) && !string.IsNullOrEmpty(OLDMB001) && !string.IsNullOrEmpty(NEWMB001))
                    {
                        UPDATEMOCTB(TA001, TA002, OLDMB001, NEWMB001);
                    }
                }
                else
                {
                    TA001 = null;
                    TA002 = null;
                }
            }

        }

        public void UPDATEMOCTB(string TA001, string TA002, string OLDMB001, string NEWMB001)
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();


                sbSql.AppendFormat(" UPDATE [TK].dbo.MOCTB ");
                sbSql.AppendFormat(" SET TB003=INVMB.MB001,TB012=INVMB.MB002,TB013=INVMB.MB003");
                sbSql.AppendFormat(" FROM [TK].dbo.INVMB");
                sbSql.AppendFormat(" WHERE INVMB.MB001='{0}'",NEWMB001);
                sbSql.AppendFormat(" AND TB001='{0}' AND TB002='{1}' AND TB003='{2}' ",TA001, TA002,OLDMB001);
                sbSql.AppendFormat(" ");
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

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    TA001 = row.Cells["製令"].Value.ToString().Trim();
                    TA002 = row.Cells["單號"].Value.ToString().Trim();
                    
                    SEARCHMOCTB(TA001, TA002);

                }
                else
                {
                    SEARCHMOCTB("", "");

                }
            }
        }

        public void SEARCHMOCTB(string ta001,string TA002)
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT TB003 AS '材料品號',TB012 AS '材料品名',TB004 AS '需領用量'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.MOCTB");
                sbSql.AppendFormat(@"  WHERE TB001='{0}' AND TB002='{1}'",TA001,TA002);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                sqlConn.Open();
                ds2.Clear();
                adapter2.Fill(ds2, "ds2");
                sqlConn.Close();


                if (ds2.Tables["ds2"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else
                {
                    if (ds2.Tables["ds2"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView2.DataSource = ds2.Tables["ds2"];
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

        public void SEARCH2(string SDAY,string EDAY)
        {
            SqlConnection sqlConn = new SqlConnection();
            string connectionString;
            StringBuilder sbSql = new StringBuilder();
          
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
          
            DataSet ds = new DataSet();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

               
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  
                                    SELECT *
                                    FROM (
                                    SELECT TA001 AS '製令',TA002 AS '單號',TA003 AS '生產日',TA006 AS '品號',TA034 AS '品名',TA015 AS '生產量',TA007 AS '單位',MD002 AS '線別',TA026 AS '訂單',TA027 AS '訂單單號',TA028 AS '序號'
                                    FROM [TK].dbo.MOCTA,[TK].dbo.CMSMD
                                    WHERE TA021=MD001
                                    AND TA003>='{0}' AND TA003<='{1}' 

                                    UNION ALL 
                                    SELECT TA001 AS '製令',TA002 AS '單號',TA003 AS '生產日',TA006 AS '品號',TA034 AS '品名',TA015 AS '生產量',TA007 AS '單位','' AS '線別',TA026 AS '訂單',TA027 AS '訂單單號',TA028 AS '序號'
                                    FROM [TK].dbo.MOCTA
                                    WHERE TA001='A512'
                                    AND TA003>='{0}' AND TA003<='{1}' 
                                    ) AS TEMP 
                                    ORDER BY 製令,單號,生產日
                                    ", SDAY, EDAY);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds");
                sqlConn.Close();


                if (ds.Tables["TEMPds"].Rows.Count == 0)
                {
                    dataGridView3.DataSource = null;
                }
                else
                {
                    if (ds.Tables["TEMPds"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView3.DataSource = ds.Tables["TEMPds"];
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

        public void SEARCH3(string SDAY, string EDAY)
        {
            SqlConnection sqlConn = new SqlConnection();
            string connectionString;
            StringBuilder sbSql = new StringBuilder();

            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

            DataSet ds = new DataSet();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

               
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"
                                    SELECT *
                                    FROM (
                                    SELECT TA001 AS '製令',TA002 AS '單號',TB003 AS '品號',TB012 AS '品名',TB004 AS '需領用量',TB009 AS '庫別' ,MD002 AS '線別'
                                    FROM [TK].dbo.MOCTA,[TK].dbo.MOCTB,[TK].dbo.CMSMD
                                    WHERE TA021=MD001
                                    AND TA001=TB001 AND TA002=TB002
                                    AND TB003='106061011'
                                    AND TA003>='{0}' AND TA003<='{1}'
                                    UNION ALL 
                                    SELECT TA001 AS '製令',TA002 AS '單號',TB003 AS '品號',TB012 AS '品名',TB004 AS '需領用量',TB009 AS '庫別' ,'' AS '線別'
                                    FROM [TK].dbo.MOCTA,[TK].dbo.MOCTB
                                    WHERE TA001='A512'
                                    AND TA001=TB001 AND TA002=TB002
                                    AND TB003='106061011'
                                    AND TA003>='{0}' AND TA003<='{1}'
                                    ) AS TEMP
                    
                                    ", SDAY, EDAY);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds");
                sqlConn.Close();


                if (ds.Tables["TEMPds"].Rows.Count == 0)
                {
                    dataGridView5.DataSource = null;
                }
                else
                {
                    if (ds.Tables["TEMPds"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView5.DataSource = ds.Tables["TEMPds"];
                        dataGridView5.AutoResizeColumns();
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
        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView3.CurrentRow != null)
            {
                int rowindex = dataGridView3.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView3.Rows[rowindex];
                    string TA001 = row.Cells["製令"].Value.ToString().Trim();
                    string TA002 = row.Cells["單號"].Value.ToString().Trim();

                    SEARCHMOCTB2(TA001, TA002);

                }
                else
                {
                    SEARCHMOCTB2("", "");

                }
            }
        }

        public void SEARCHMOCTB2(string TA001, string TA002)
        {
            SqlConnection sqlConn = new SqlConnection();
            string connectionString;
            StringBuilder sbSql = new StringBuilder();

            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

            DataSet ds = new DataSet();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT TA009 AS '預計開工',TA012 AS '實際開工' ");
                sbSql.AppendFormat(@"  FROM [TK].dbo.MOCTA");
                sbSql.AppendFormat(@"  WHERE TA001='{0}' AND TA002='{1}'", TA001, TA002);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    dataGridView4.DataSource = null;
                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView4.DataSource = ds.Tables["ds"];
                        dataGridView4.AutoResizeColumns();
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
        public void SEARCHMOCTB9(string TA001, string TA002)
        {
            SqlConnection sqlConn = new SqlConnection();
            string connectionString;
            StringBuilder sbSql = new StringBuilder();

            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

            DataSet ds = new DataSet();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

             
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT TA010 AS '預計完工' ");
                sbSql.AppendFormat(@"  FROM [TK].dbo.MOCTA");
                sbSql.AppendFormat(@"  WHERE TA001='{0}' AND TA002='{1}'", TA001, TA002);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    dataGridView10.DataSource = null;
                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView10.DataSource = ds.Tables["ds"];
                        dataGridView10.AutoResizeColumns();
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

        public void CHANGEMULTI2()
        {
            string NEWDATES = dateTimePicker5.Value.ToString("yyyyMMdd");

            foreach (DataGridViewRow dr in this.dataGridView3.Rows)
            {
                DataGridViewCheckBoxCell cbx = (DataGridViewCheckBoxCell)dr.Cells[0];

                if ((bool)cbx.FormattedValue)
                {
                    string TA001 = ((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString();
                    string TA002 = ((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString();

                    //MessageBox.Show(OLDTA001+"-"+ OLDTA002);
                    if (!string.IsNullOrEmpty(TA001) && !string.IsNullOrEmpty(TA002) && !string.IsNullOrEmpty(NEWDATES) )
                    {
                        UPDATEMOCTA(TA001, TA002, NEWDATES);
                    }
                }
                else
                {
                    TA001 = null;
                    TA002 = null;
                }
            }

        }

        public void CHANGEMULTI5()
        {
            string NEWDATES = dateTimePicker12.Value.ToString("yyyyMMdd");

            foreach (DataGridViewRow dr in this.dataGridView9.Rows)
            {
                DataGridViewCheckBoxCell cbx = (DataGridViewCheckBoxCell)dr.Cells[0];

                if ((bool)cbx.FormattedValue)
                {
                    string TA001 = ((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString();
                    string TA002 = ((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString();

                    //MessageBox.Show(OLDTA001+"-"+ OLDTA002);
                    if (!string.IsNullOrEmpty(TA001) && !string.IsNullOrEmpty(TA002) && !string.IsNullOrEmpty(NEWDATES))
                    {
                        UPDATEMOCTATA010(TA001, TA002, NEWDATES);
                    }
                }
                else
                {
                    TA001 = null;
                    TA002 = null;
                }
            }

        }

        public void CHANGEMOCTATA012(string YYYYMM)
        {
            UPDATEMOCTATA012(YYYYMM);

        }

        public void UPDATEMOCTA(string TA001, string TA002, string TA009)
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();



                sbSql.AppendFormat(" UPDATE [TK].dbo.MOCTA");
                sbSql.AppendFormat(" SET TA009='{0}',TA012='{0}'", TA009);
                sbSql.AppendFormat(" WHERE TA001='{0}' AND TA002='{1}'",TA001,TA002);
                sbSql.AppendFormat(" ");
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

        public void UPDATEMOCTATA010(string TA001, string TA002, string TA009)
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();



                sbSql.AppendFormat(" UPDATE [TK].dbo.MOCTA");
                sbSql.AppendFormat(" SET TA010='{0}'", TA009);
                sbSql.AppendFormat(" WHERE TA001='{0}' AND TA002='{1}'", TA001, TA002);
                sbSql.AppendFormat(" ");
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

        public void UPDATEMOCTATA012(string YYYYMM)
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();



                sbSql.AppendFormat(@" 
                                    UPDATE [TK].dbo.MOCTA
                                    SET TA012=(SELECT TOP 1 TC003 FROM  [TK].dbo.MOCTC,[TK].dbo.MOCTE WHERE TC001=TE001 AND TC002=TE002 AND TA001=TE011 AND TA002=TE012 ORDER BY TC003 )
                                    WHERE TA012<>(SELECT TOP 1 TC003 FROM  [TK].dbo.MOCTC,[TK].dbo.MOCTE WHERE TC001=TE001 AND TC002=TE002 AND TA001=TE011 AND TA002=TE012 ORDER BY TC003 )
                                    AND TA003 LIKE '{0}%'
                                    ",YYYYMM);



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

        public void CHANGEMULTI3()
        {
            foreach (DataGridViewRow dr in this.dataGridView5.Rows)
            {
                DataGridViewCheckBoxCell cbx = (DataGridViewCheckBoxCell)dr.Cells[0];

                if ((bool)cbx.FormattedValue)
                {
                    string TB001 = ((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString();
                    string TB002 = ((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString();
                    string TB009 = comboBox3.Text;

                    if(!string.IsNullOrEmpty(TB001)&& !string.IsNullOrEmpty(TB002) && !string.IsNullOrEmpty(TB009) )
                    {
                        UPDATEMOCTB(TB001.Trim(), TB002.Trim(), TB009.Trim());
                    }

                    //MessageBox.Show(OLDTA001+"-"+ OLDTA002);

                }
                else
                {
                  
                }
            }

        }

        public void UPDATEMOCTB(string TB001, string TB002, string TB009)
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();



                sbSql.AppendFormat(" UPDATE [TK].dbo.MOCTB");
                sbSql.AppendFormat(" SET TB009='{0}'", TB009);
                sbSql.AppendFormat(" WHERE TB003='106061011' AND TB001='{0}' AND TB002='{1}'", TB001, TB002);
                sbSql.AppendFormat(" ");
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


        public void SEARCH4(string SDATES,string EDATES)
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  
                                   SELECT *
                                    FROM (
                                    SELECT TA001 AS '製令',TA002 AS '單號',TA003 AS '生產日',TA006 AS '品號',TA034 AS '品名',TA015 AS '生產量',TA007 AS '單位',MD002 AS '線別',TA026 AS '訂單',TA027 AS '訂單單號',TA028 AS '序號'
                                    FROM [TK].dbo.MOCTA,[TK].dbo.CMSMD
                                    WHERE TA021=MD001
                                    AND TA003>='{0}' AND TA003<='{1}' 

                                    UNION ALL 
                                    SELECT TA001 AS '製令',TA002 AS '單號',TA003 AS '生產日',TA006 AS '品號',TA034 AS '品名',TA015 AS '生產量',TA007 AS '單位','' AS '線別',TA026 AS '訂單',TA027 AS '訂單單號',TA028 AS '序號'
                                    FROM [TK].dbo.MOCTA
                                    WHERE TA001='A512'
                                    AND TA003>='{0}' AND TA003<='{1}' 
                                    ) AS TEMP 
                                    ORDER BY 製令,單號,生產日
                                    ", SDATES, EDATES);

                adapter5 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder5 = new SqlCommandBuilder(adapter5);
                sqlConn.Open();
                ds5.Clear();
                adapter5.Fill(ds5, "ds5");
                sqlConn.Close();


                if (ds5.Tables["ds5"].Rows.Count == 0)
                {
                    dataGridView7.DataSource = null;
                }
                else
                {
                    if (ds5.Tables["ds5"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView7.DataSource = ds5.Tables["ds5"];
                        dataGridView7.AutoResizeColumns();
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

        public void SEARCH5(string SDAY, string EDAY)
        {
            SqlConnection sqlConn = new SqlConnection();
            string connectionString;
            StringBuilder sbSql = new StringBuilder();

            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

            DataSet ds = new DataSet();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

               
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  
                                    SELECT *
                                    FROM (
                                    SELECT TA001 AS '製令',TA002 AS '單號',TA003 AS '生產日',TA006 AS '品號',TA034 AS '品名',TA015 AS '生產量',TA007 AS '單位',MD002 AS '線別',TA026 AS '訂單',TA027 AS '訂單單號',TA028 AS '序號'
                                    FROM [TK].dbo.MOCTA,[TK].dbo.CMSMD
                                    WHERE TA021=MD001
                                    AND TA003>='{0}' AND TA003<='{1}' 

                                    UNION ALL 
                                    SELECT TA001 AS '製令',TA002 AS '單號',TA003 AS '生產日',TA006 AS '品號',TA034 AS '品名',TA015 AS '生產量',TA007 AS '單位','' AS '線別',TA026 AS '訂單',TA027 AS '訂單單號',TA028 AS '序號'
                                    FROM [TK].dbo.MOCTA
                                    WHERE TA001='A512'
                                    AND TA003>='{0}' AND TA003<='{1}' 
                                    ) AS TEMP 
                                    ORDER BY 製令,單號,生產日
                                    ", SDAY, EDAY);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds");
                sqlConn.Close();


                if (ds.Tables["TEMPds"].Rows.Count == 0)
                {
                    dataGridView9.DataSource = null;
                }
                else
                {
                    if (ds.Tables["TEMPds"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView9.DataSource = ds.Tables["TEMPds"];
                        dataGridView9.AutoResizeColumns();
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

        public void SEARCHMOCTATA012(string YYYYMM)
        {
            SqlConnection sqlConn = new SqlConnection();
            string connectionString;
            StringBuilder sbSql = new StringBuilder();

            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

            DataSet ds = new DataSet();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

              
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  
                                   SELECT TA001 AS '製令單別',TA002 AS '製令單號',TA012 AS '目前的實際開工日'
                                    ,(SELECT TOP 1 TC003 FROM  [TK].dbo.MOCTC,[TK].dbo.MOCTE WHERE TC001=TE001 AND TC002=TE002 AND TA001=TE011 AND TA002=TE012 ORDER BY TC003 ) AS '最早的領料日'
                                    FROM [TK].dbo.MOCTA
                                    WHERE TA012<>(SELECT TOP 1 TC003 FROM  [TK].dbo.MOCTC,[TK].dbo.MOCTE WHERE TC001=TE001 AND TC002=TE002 AND TA001=TE011 AND TA002=TE012 ORDER BY TC003 )
                                    AND TA003 LIKE '{0}%'

                                    ", YYYYMM);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds");
                sqlConn.Close();


                if (ds.Tables["TEMPds"].Rows.Count == 0)
                {
                    dataGridView11.DataSource = null;
                }
                else
                {
                    if (ds.Tables["TEMPds"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView11.DataSource = ds.Tables["TEMPds"];
                        dataGridView11.AutoResizeColumns();
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

        public void CHANGEMULTI4()
        {
            foreach (DataGridViewRow dr in this.dataGridView7.Rows)
            {
                DataGridViewCheckBoxCell cbx = (DataGridViewCheckBoxCell)dr.Cells[0];

                if ((bool)cbx.FormattedValue)
                {
                    string TB001 = ((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString();
                    string TB002 = ((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString();

                    if (!string.IsNullOrEmpty(TB001) && !string.IsNullOrEmpty(TB002) )
                    {
                        UPDATEMOCTB2(TB001.Trim(), TB002.Trim());
                    }

                    //MessageBox.Show(OLDTA001+"-"+ OLDTA002);

                }
                else
                {

                }
            }

        }

        public void UPDATEMOCTB2(string TB001, string TB002)
        {
            string SQLLIKE = SEARCHMOCCHANGE();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();



                sbSql.AppendFormat(@" 
                                    UPDATE [TK].dbo.MOCTB
                                    SET TB004=ROUND(TB004,0)
                                    WHERE ( {2} )
                                    AND TB001='{0}' AND TB002='{1}'
                                    ", TB001, TB002, SQLLIKE.ToString());



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

        public string SEARCHMOCCHANGE()
        {
            StringBuilder MB001 = new StringBuilder();
            MB001.Clear();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  
                                    SELECT [MB001] FROM [TKMOC].[dbo].[MOCCHANGE]
                                    ");

                adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder4 = new SqlCommandBuilder(adapter4);
                sqlConn.Open();
                ds4.Clear();
                adapter4.Fill(ds4, "ds4");
                sqlConn.Close();


                if (ds4.Tables["ds4"].Rows.Count >= 1)
                {
                    for (int i = 0; i < ds4.Tables["ds4"].Rows.Count; i++)
                    {
                        MB001.AppendFormat(@" (TB003 LIKE '{0}%') OR ", ds4.Tables["ds4"].Rows[i]["MB001"].ToString());
                    }

                    MB001.AppendFormat(@" (TB003 LIKE 'NA%') ");
                    return MB001.ToString();

                }
                else
                {
                    return "";
                }

            }
            catch
            {
                return "";
            }
            finally
            {

            }
        }
        private void dataGridView7_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView7.CurrentRow != null)
            {
                int rowindex = dataGridView7.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView7.Rows[rowindex];
                    TA001 = row.Cells["製令"].Value.ToString().Trim();
                    TA002 = row.Cells["單號"].Value.ToString().Trim();

                    SEARCHMOCTB3(TA001, TA002);

                }
                else
                {
                    SEARCHMOCTB3("", "");

                }
            }
        }

        public void SEARCHMOCTB3(string ta001, string TA002)
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  
                                SELECT  LTRIM(RTRIM(TB003)) AS '材料品號',LTRIM(RTRIM(TB012)) AS '材料品名',TB004 AS '需領用量'
                                FROM [TK].dbo.MOCTB
                                WHERE TB001='{0}' AND TB002='{1}'
                                ", TA001, TA002);

                adapter3 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder3 = new SqlCommandBuilder(adapter3);
                sqlConn.Open();
                ds3.Clear();
                adapter3.Fill(ds3, "ds3");
                sqlConn.Close();


                if (ds3.Tables["ds3"].Rows.Count == 0)
                {
                    dataGridView8.DataSource = null;
                }
                else
                {
                    if (ds3.Tables["ds3"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView8.DataSource = ds3.Tables["ds3"];
                        dataGridView8.AutoResizeColumns();

                        dataGridView8.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);
                        dataGridView8.DefaultCellStyle.Font = new Font("Tahoma", 10);
                        dataGridView8.Columns[0].Width = 120;
                        dataGridView8.Columns[1].Width = 100;
                        dataGridView8.Columns[2].Width = 100;
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
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            //checkBox1-製一線
            if (checkBox1.Checked)
            {
                dataGridView1checkBox1True();
            }
            else
            {
                dataGridView1checkBox1False();
            }
        }

        public void dataGridView1checkBox1True()
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if(dataGridView1.Rows[i].Cells["線別"].Value.ToString().Trim().Equals("製一線"))
                {
                    dataGridView1.Rows[i].Cells[0].Value = 1;
                }
                
            }
        }
        public void dataGridView1checkBox1False()
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1.Rows[i].Cells["線別"].Value.ToString().Trim().Equals("製一線"))
                {
                    dataGridView1.Rows[i].Cells[0].Value = 0;
                }

            }
        }
        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            //checkBox2-製二線
            if (checkBox2.Checked)
            {
                dataGridView1checkBox2True();
            }
            else
            {
                dataGridView1checkBox2False();
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            //checkBox3-包裝線
            if (checkBox3.Checked)
            {
                dataGridView1checkBox3True();
            }
            else
            {
                dataGridView1checkBox3False();
            }
        }
        public void dataGridView1checkBox2True()
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1.Rows[i].Cells["線別"].Value.ToString().Trim().Equals("製二線"))
                {
                    dataGridView1.Rows[i].Cells[0].Value = 1;
                }

            }
        }
        public void dataGridView1checkBox2False()
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1.Rows[i].Cells["線別"].Value.ToString().Trim().Equals("製二線"))
                {
                    dataGridView1.Rows[i].Cells[0].Value = 0;
                }

            }
        }
        public void dataGridView1checkBox3True()
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1.Rows[i].Cells["線別"].Value.ToString().Trim().Equals("包裝線"))
                {
                    dataGridView1.Rows[i].Cells[0].Value = 1;
                }

            }
        }
        public void dataGridView1checkBox3False()
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1.Rows[i].Cells["線別"].Value.ToString().Trim().Equals("包裝線"))
                {
                    dataGridView1.Rows[i].Cells[0].Value = 0;
                }

            }
        }
        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            //checkBox4-製一線
            if (checkBox4.Checked)
            {
                dataGridView3checkBox4True();
            }
            else
            {
                dataGridView3checkBox4False();
            }
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            //checkBox5-製二線
            if (checkBox5.Checked)
            {
                dataGridView3checkBox5True();
            }
            else
            {
                dataGridView3checkBox5False();
            }
        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            //checkBox6-包裝線
            if (checkBox6.Checked)
            {
                dataGridView3checkBox6True();
            }
            else
            {
                dataGridView3checkBox6False();
            }
        }

        public void dataGridView3checkBox4True()
        {
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                if (dataGridView3.Rows[i].Cells["線別"].Value.ToString().Trim().Equals("製一線"))
                {
                    dataGridView3.Rows[i].Cells[0].Value = 1;
                }

            }
        }
        public void dataGridView3checkBox4False()
        {
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                if (dataGridView3.Rows[i].Cells["線別"].Value.ToString().Trim().Equals("製一線"))
                {
                    dataGridView3.Rows[i].Cells[0].Value = 0;
                }

            }
        }
        public void dataGridView3checkBox5True()
        {
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                if (dataGridView3.Rows[i].Cells["線別"].Value.ToString().Trim().Equals("製二線"))
                {
                    dataGridView3.Rows[i].Cells[0].Value = 1;
                }

            }
        }
        public void dataGridView3checkBox5False()
        {
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                if (dataGridView3.Rows[i].Cells["線別"].Value.ToString().Trim().Equals("製二線"))
                {
                    dataGridView3.Rows[i].Cells[0].Value = 0;
                }

            }
        }
        public void dataGridView3checkBox6True()
        {
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                if (dataGridView3.Rows[i].Cells["線別"].Value.ToString().Trim().Equals("包裝線"))
                {
                    dataGridView3.Rows[i].Cells[0].Value = 1;
                }

            }
        }
        public void dataGridView3checkBox6False()
        {
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                if (dataGridView3.Rows[i].Cells["線別"].Value.ToString().Trim().Equals("包裝線"))
                {
                    dataGridView3.Rows[i].Cells[0].Value = 0;
                }

            }
        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {
            //checkBox7-製一線
            if (checkBox7.Checked)
            {
                dataGridView5checkBox7True();
            }
            else
            {
                dataGridView5checkBox7False();
            }
        }

        private void checkBox8_CheckedChanged(object sender, EventArgs e)
        {
            //checkBox8-製二線
            if (checkBox8.Checked)
            {
                dataGridView5checkBox8True();
            }
            else
            {
                dataGridView5checkBox8False();
            }
        }

        private void checkBox9_CheckedChanged(object sender, EventArgs e)
        {

            //checkBox9-包裝線
            if (checkBox9.Checked)
            {
                dataGridView5checkBox9True();
            }
            else
            {
                dataGridView5checkBox9False();
            }
        }

        public void dataGridView5checkBox7True()
        {
            for (int i = 0; i < dataGridView5.Rows.Count; i++)
            {
                if (dataGridView5.Rows[i].Cells["線別"].Value.ToString().Trim().Equals("製一線"))
                {
                    dataGridView5.Rows[i].Cells[0].Value = 1;
                }

            }
        }
        public void dataGridView5checkBox7False()
        {
            for (int i = 0; i < dataGridView5.Rows.Count; i++)
            {
                if (dataGridView5.Rows[i].Cells["線別"].Value.ToString().Trim().Equals("製一線"))
                {
                    dataGridView5.Rows[i].Cells[0].Value = 0;
                }

            }
        }
        public void dataGridView5checkBox8True()
        {
            for (int i = 0; i < dataGridView5.Rows.Count; i++)
            {
                if (dataGridView5.Rows[i].Cells["線別"].Value.ToString().Trim().Equals("製二線"))
                {
                    dataGridView5.Rows[i].Cells[0].Value = 1;
                }

            }
        }
        public void dataGridView5checkBox8False()
        {
            for (int i = 0; i < dataGridView5.Rows.Count; i++)
            {
                if (dataGridView5.Rows[i].Cells["線別"].Value.ToString().Trim().Equals("製二線"))
                {
                    dataGridView5.Rows[i].Cells[0].Value = 0;
                }

            }
        }
        public void dataGridView5checkBox9True()
        {
            for (int i = 0; i < dataGridView5.Rows.Count; i++)
            {
                if (dataGridView5.Rows[i].Cells["線別"].Value.ToString().Trim().Equals("包裝線"))
                {
                    dataGridView5.Rows[i].Cells[0].Value = 1;
                }

            }
        }
        public void dataGridView5checkBox9False()
        {
            for (int i = 0; i < dataGridView5.Rows.Count; i++)
            {
                if (dataGridView5.Rows[i].Cells["線別"].Value.ToString().Trim().Equals("包裝線"))
                {
                    dataGridView5.Rows[i].Cells[0].Value = 0;
                }

            }
        }

        private void checkBox10_CheckedChanged(object sender, EventArgs e)
        {
            //checkBox10-製一線
            if (checkBox10.Checked)
            {
                dataGridView7checkBox10True();
            }
            else
            {
                dataGridView7checkBox10False();
            }
        }

        private void checkBox11_CheckedChanged(object sender, EventArgs e)
        {
            //checkBox11-製二線
            if (checkBox11.Checked)
            {
                dataGridView7checkBox11True();
            }
            else
            {
                dataGridView7checkBox11False();
            }
        }

        private void checkBox12_CheckedChanged(object sender, EventArgs e)
        {
            //checkBox12-包裝線
            if (checkBox12.Checked)
            {
                dataGridView7checkBox12True();
            }
            else
            {
                dataGridView7checkBox12False();
            }
        }

        public void dataGridView7checkBox10True()
        {
            for (int i = 0; i < dataGridView7.Rows.Count; i++)
            {
                if (dataGridView7.Rows[i].Cells["線別"].Value.ToString().Trim().Equals("製一線"))
                {
                    dataGridView7.Rows[i].Cells[0].Value = 1;
                }

            }
        }
        public void dataGridView7checkBox10False()
        {
            for (int i = 0; i < dataGridView7.Rows.Count; i++)
            {
                if (dataGridView7.Rows[i].Cells["線別"].Value.ToString().Trim().Equals("製一線"))
                {
                    dataGridView7.Rows[i].Cells[0].Value = 0;
                }

            }
        }
        public void dataGridView7checkBox11True()
        {
            for (int i = 0; i < dataGridView7.Rows.Count; i++)
            {
                if (dataGridView7.Rows[i].Cells["線別"].Value.ToString().Trim().Equals("製二線"))
                {
                    dataGridView7.Rows[i].Cells[0].Value = 1;
                }

            }
        }
        public void dataGridView7checkBox11False()
        {
            for (int i = 0; i < dataGridView7.Rows.Count; i++)
            {
                if (dataGridView7.Rows[i].Cells["線別"].Value.ToString().Trim().Equals("製二線"))
                {
                    dataGridView7.Rows[i].Cells[0].Value = 0;
                }

            }
        }
        public void dataGridView7checkBox12True()
        {
            for (int i = 0; i < dataGridView7.Rows.Count; i++)
            {
                if (dataGridView7.Rows[i].Cells["線別"].Value.ToString().Trim().Equals("包裝線"))
                {
                    dataGridView7.Rows[i].Cells[0].Value = 1;
                }

            }
        }
        public void dataGridView7checkBox12False()
        {
            for (int i = 0; i < dataGridView7.Rows.Count; i++)
            {
                if (dataGridView7.Rows[i].Cells["線別"].Value.ToString().Trim().Equals("包裝線"))
                {
                    dataGridView7.Rows[i].Cells[0].Value = 0;
                }

            }
        }
        public void SETCHECK()
        {
            checkBox1.Checked = false;
            checkBox2.Checked = false;
            checkBox3.Checked = false;
            checkBox4.Checked = false;
            checkBox5.Checked = false;
            checkBox6.Checked = false;
            checkBox7.Checked = false;
            checkBox8.Checked = false;
            checkBox9.Checked = false;
            checkBox10.Checked = false;
            checkBox11.Checked = false;
            checkBox12.Checked = false;
            checkBox13.Checked = false;
            checkBox14.Checked = false;
            checkBox15.Checked = false;
        }

        private void checkBox13_CheckedChanged(object sender, EventArgs e)
        {
            //checkBox4-製一線
            if (checkBox13.Checked)
            {
                dataGridView9checkBox13True();
            }
            else
            {
                dataGridView9checkBox13False();
            }
        }

        private void checkBox14_CheckedChanged(object sender, EventArgs e)
        {
            //checkBox5-製二線
            if (checkBox14.Checked)
            {
                dataGridView9checkBox14True();
            }
            else
            {
                dataGridView9checkBox14False();
            }
        }

        private void checkBox15_CheckedChanged(object sender, EventArgs e)
        {
            //checkBox9-包裝線
            if (checkBox15.Checked)
            {
                dataGridView9checkBox15True();
            }
            else
            {
                dataGridView9checkBox15False();
            }
        }

        public void dataGridView9checkBox13True()
        {
            for (int i = 0; i < dataGridView9.Rows.Count; i++)
            {
                if (dataGridView9.Rows[i].Cells["線別"].Value.ToString().Trim().Equals("製一線"))
                {
                    dataGridView9.Rows[i].Cells[0].Value = 1;
                }

            }
        }
        public void dataGridView9checkBox13False()
        {
            for (int i = 0; i < dataGridView9.Rows.Count; i++)
            {
                if (dataGridView9.Rows[i].Cells["線別"].Value.ToString().Trim().Equals("製一線"))
                {
                    dataGridView9.Rows[i].Cells[0].Value = 0;
                }

            }
        }
        public void dataGridView9checkBox14True()
        {
            for (int i = 0; i < dataGridView9.Rows.Count; i++)
            {
                if (dataGridView9.Rows[i].Cells["線別"].Value.ToString().Trim().Equals("製二線"))
                {
                    dataGridView9.Rows[i].Cells[0].Value = 1;
                }

            }
        }
        public void dataGridView9checkBox14False()
        {
            for (int i = 0; i < dataGridView9.Rows.Count; i++)
            {
                if (dataGridView9.Rows[i].Cells["線別"].Value.ToString().Trim().Equals("製二線"))
                {
                    dataGridView9.Rows[i].Cells[0].Value = 0;
                }

            }
        }
        public void dataGridView9checkBox15True()
        {
            for (int i = 0; i < dataGridView9.Rows.Count; i++)
            {
                if (dataGridView9.Rows[i].Cells["線別"].Value.ToString().Trim().Equals("包裝線"))
                {
                    dataGridView9.Rows[i].Cells[0].Value = 1;
                }

            }
        }
        public void dataGridView9checkBox15False()
        {
            for (int i = 0; i < dataGridView9.Rows.Count; i++)
            {
                if (dataGridView9.Rows[i].Cells["線別"].Value.ToString().Trim().Equals("包裝線"))
                {
                    dataGridView9.Rows[i].Cells[0].Value = 0;
                }

            }
        }

        private void dataGridView9_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView9.CurrentRow != null)
            {
                int rowindex = dataGridView9.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView9.Rows[rowindex];
                    string TA001 = row.Cells["製令"].Value.ToString().Trim();
                    string TA002 = row.Cells["單號"].Value.ToString().Trim();

                    SEARCHMOCTB9(TA001, TA002);

                }
                else
                {
                    SEARCHMOCTB9("", "");

                }
            }
        }

        private void checkBox16_CheckedChanged(object sender, EventArgs e)
        {
            //手工線
            if (checkBox16.Checked)
            {
                dataGridView1checkBox16True();
            }
            else
            {
                dataGridView1checkBox16False();
            }
        }

        public void dataGridView1checkBox16True()
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1.Rows[i].Cells["線別"].Value.ToString().Trim().Equals("手工線"))
                {
                    dataGridView1.Rows[i].Cells[0].Value = 1;
                }

            }
        }

        public void dataGridView1checkBox16False()
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1.Rows[i].Cells["線別"].Value.ToString().Trim().Equals("手工線"))
                {
                    dataGridView1.Rows[i].Cells[0].Value =0;
                }

            }
        }
        private void checkBox17_CheckedChanged(object sender, EventArgs e)
        {
            //手工線
            if (checkBox17.Checked)
            {
                dataGridView3checkBox17True();
            }
            else
            {
                dataGridView3checkBox17False();
            }

           
        }
        public void dataGridView3checkBox17True()
        {
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                if (dataGridView3.Rows[i].Cells["線別"].Value.ToString().Trim().Equals("手工線"))
                {
                    dataGridView3.Rows[i].Cells[0].Value = 1;
                }

            }
        }
        public void dataGridView3checkBox17False()
        {
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                if (dataGridView3.Rows[i].Cells["線別"].Value.ToString().Trim().Equals("手工線"))
                {
                    dataGridView3.Rows[i].Cells[0].Value = 0;
                }

            }
        }
        private void checkBox18_CheckedChanged(object sender, EventArgs e)
        {
            //手工線
            if (checkBox18.Checked)
            {
                dataGridView5checkBox18True();
            }
            else
            {
                dataGridView5checkBox18False();
            }
            
        }


        public void dataGridView5checkBox18True()
        {
            for (int i = 0; i < dataGridView5.Rows.Count; i++)
            {
                if (dataGridView5.Rows[i].Cells["線別"].Value.ToString().Trim().Equals("手工線"))
                {
                    dataGridView5.Rows[i].Cells[0].Value = 1;
                }

            }
        }
        public void dataGridView5checkBox18False()
        {
            for (int i = 0; i < dataGridView5.Rows.Count; i++)
            {
                if (dataGridView5.Rows[i].Cells["線別"].Value.ToString().Trim().Equals("手工線"))
                {
                    dataGridView5.Rows[i].Cells[0].Value = 0;
                }

            }
        }

        private void checkBox19_CheckedChanged(object sender, EventArgs e)
        {
            //手工線
            if (checkBox19.Checked)
            {
                dataGridView7checkBox19True();
            }
            else
            {
                dataGridView7checkBox19False();
            }
            
        }
        public void dataGridView7checkBox19True()
        {
            for (int i = 0; i < dataGridView7.Rows.Count; i++)
            {
                if (dataGridView7.Rows[i].Cells["線別"].Value.ToString().Trim().Equals("手工線"))
                {
                    dataGridView7.Rows[i].Cells[0].Value = 1;
                }

            }
        }
        public void dataGridView7checkBox19False()
        {
            for (int i = 0; i < dataGridView7.Rows.Count; i++)
            {
                if (dataGridView7.Rows[i].Cells["線別"].Value.ToString().Trim().Equals("手工線"))
                {
                    dataGridView7.Rows[i].Cells[0].Value = 0;
                }

            }
        }
        private void checkBox20_CheckedChanged(object sender, EventArgs e)
        {

            //手工線
            if (checkBox20.Checked)
            {
                dataGridView9checkBox20True();
            }
            else
            {
                dataGridView9checkBox20False();
            }
         
        }
        public void dataGridView9checkBox20True()
        {
            for (int i = 0; i < dataGridView9.Rows.Count; i++)
            {
                if (dataGridView9.Rows[i].Cells["線別"].Value.ToString().Trim().Equals("手工線"))
                {
                    dataGridView9.Rows[i].Cells[0].Value = 1;
                }

            }
        }
        public void dataGridView9checkBox20False()
        {
            for (int i = 0; i < dataGridView9.Rows.Count; i++)
            {
                if (dataGridView9.Rows[i].Cells["線別"].Value.ToString().Trim().Equals("手工線"))
                {
                    dataGridView9.Rows[i].Cells[0].Value = 0;
                }

            }
        }
        private void checkBox21_CheckedChanged(object sender, EventArgs e)
        {
            //
            if (checkBox21.Checked)
            {
                dataGridView1checkBox21True();
            }
            else
            {
                dataGridView1checkBox21False();
            }
        }

        public void dataGridView1checkBox21True()
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1.Rows[i].Cells["製令"].Value.ToString().Trim().Equals("A512"))
                {
                    dataGridView1.Rows[i].Cells[0].Value = 1;
                }

            }
        }
        public void dataGridView1checkBox21False()
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1.Rows[i].Cells["製令"].Value.ToString().Trim().Equals("A512"))
                {
                    dataGridView1.Rows[i].Cells[0].Value = 0;
                }

            }
        }
        private void checkBox22_CheckedChanged(object sender, EventArgs e)
        {
            //
            if (checkBox22.Checked)
            {
                dataGridView3checkBox22True();
            }
            else
            {
                dataGridView3checkBox22False();
            }
        }
        public void dataGridView3checkBox22True()
        {
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                if (dataGridView3.Rows[i].Cells["製令"].Value.ToString().Trim().Equals("A512"))
                {
                    dataGridView3.Rows[i].Cells[0].Value = 1;
                }

            }
        }
        public void dataGridView3checkBox22False()
        {
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                if (dataGridView3.Rows[i].Cells["製令"].Value.ToString().Trim().Equals("A512"))
                {
                    dataGridView3.Rows[i].Cells[0].Value = 0;
                }

            }
        }

        private void checkBox23_CheckedChanged(object sender, EventArgs e)
        {
            //
            if (checkBox23.Checked)
            {
                dataGridView5checkBox23True();
            }
            else
            {
                dataGridView5checkBox23False();
            }
        }

        public void dataGridView5checkBox23True()
        {
            for (int i = 0; i < dataGridView5.Rows.Count; i++)
            {
                if (dataGridView5.Rows[i].Cells["製令"].Value.ToString().Trim().Equals("A512"))
                {
                    dataGridView5.Rows[i].Cells[0].Value = 1;
                }

            }
        }
        public void dataGridView5checkBox23False()
        {
            for (int i = 0; i < dataGridView5.Rows.Count; i++)
            {
                if (dataGridView5.Rows[i].Cells["製令"].Value.ToString().Trim().Equals("A512"))
                {
                    dataGridView5.Rows[i].Cells[0].Value = 0;
                }

            }
        }
        private void checkBox24_CheckedChanged(object sender, EventArgs e)
        {
            //
            if (checkBox24.Checked)
            {
                dataGridView7checkBox24True();
            }
            else
            {
                dataGridView7checkBox24False();
            }
        }
        public void dataGridView7checkBox24True()
        {
            for (int i = 0; i < dataGridView7.Rows.Count; i++)
            {
                if (dataGridView7.Rows[i].Cells["製令"].Value.ToString().Trim().Equals("A512"))
                {
                    dataGridView7.Rows[i].Cells[0].Value = 1;
                }

            }
        }
        public void dataGridView7checkBox24False()
        {
            for (int i = 0; i < dataGridView7.Rows.Count; i++)
            {
                if (dataGridView7.Rows[i].Cells["製令"].Value.ToString().Trim().Equals("A512"))
                {
                    dataGridView7.Rows[i].Cells[0].Value = 0;
                }

            }
        }
        private void checkBox25_CheckedChanged(object sender, EventArgs e)
        {
            //
            if (checkBox25.Checked)
            {
                dataGridView9checkBox25True();
            }
            else
            {
                dataGridView9checkBox25False();
            }
        }
        public void dataGridView9checkBox25True()
        {
            for (int i = 0; i < dataGridView9.Rows.Count; i++)
            {
                if (dataGridView9.Rows[i].Cells["製令"].Value.ToString().Trim().Equals("A512"))
                {
                    dataGridView9.Rows[i].Cells[0].Value = 1;
                }

            }
        }
        public void dataGridView9checkBox25False()
        {
            for (int i = 0; i < dataGridView9.Rows.Count; i++)
            {
                if (dataGridView9.Rows[i].Cells["製令"].Value.ToString().Trim().Equals("A512"))
                {
                    dataGridView9.Rows[i].Cells[0].Value = 0;
                }

            }
        }

        #endregion

        #region BUTTON
        private void button3_Click(object sender, EventArgs e)
        {
            SEARCH();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要修改嗎?", "要修改嗎?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                CHANGEMULTI();
                SEARCH();
                SETCHECK();

            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }




        private void button1_Click(object sender, EventArgs e)
        {
            SEARCH2(dateTimePicker3.Value.ToString("yyyyMMdd"), dateTimePicker4.Value.ToString("yyyyMMdd"));
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要修改嗎?", "要修改嗎?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                CHANGEMULTI2();
                SEARCH2(dateTimePicker3.Value.ToString("yyyyMMdd"), dateTimePicker4.Value.ToString("yyyyMMdd"));
                SETCHECK();

            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }
        private void button5_Click(object sender, EventArgs e)
        {
            SEARCH3(dateTimePicker6.Value.ToString("yyyyMMdd"), dateTimePicker7.Value.ToString("yyyyMMdd"));
        }

        private void button6_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要修改嗎?", "要修改嗎?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                CHANGEMULTI3();
                SEARCH3(dateTimePicker6.Value.ToString("yyyyMMdd"), dateTimePicker7.Value.ToString("yyyyMMdd"));
                SETCHECK();

            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }


        private void button7_Click(object sender, EventArgs e)
        {
            SEARCH4(dateTimePicker8.Value.ToString("yyyyMMdd"), dateTimePicker9.Value.ToString("yyyyMMdd"));
        }

        private void button8_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要修改嗎?", "要修改嗎?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                CHANGEMULTI4();
                SEARCH4(dateTimePicker8.Value.ToString("yyyyMMdd"), dateTimePicker9.Value.ToString("yyyyMMdd"));
                SETCHECK();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }





        private void button9_Click(object sender, EventArgs e)
        {
            SEARCH5(dateTimePicker10.Value.ToString("yyyyMMdd"), dateTimePicker11.Value.ToString("yyyyMMdd"));
        }

        private void button10_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要修改嗎?", "要修改嗎?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                CHANGEMULTI5();
                SEARCH5(dateTimePicker10.Value.ToString("yyyyMMdd"), dateTimePicker11.Value.ToString("yyyyMMdd"));
                SETCHECK();

            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            SEARCHMOCTATA012(dateTimePicker13.Value.ToString("yyyyMM"));
        }

        private void button12_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要修改嗎?", "要修改嗎?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                CHANGEMOCTATA012(dateTimePicker13.Value.ToString("yyyyMM"));
                SEARCHMOCTATA012(dateTimePicker13.Value.ToString("yyyyMM"));


            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }









        #endregion

       
    }
    
}
