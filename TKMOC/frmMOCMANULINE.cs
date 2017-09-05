using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using NPOI.SS.UserModel;
using System.Configuration;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using System.Reflection;
using System.Threading;
using System.Globalization;

namespace TKMOC
{
    public partial class frmMOCMANULINE : Form
    {
        private ComponentResourceManager _ResourceManager = new ComponentResourceManager();
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        SqlDataAdapter adapter2 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder2 = new SqlCommandBuilder();
        SqlDataAdapter adapter3 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder3 = new SqlCommandBuilder();
        SqlDataAdapter adapter4 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder4 = new SqlCommandBuilder();
        SqlDataAdapter adapter5 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder5 = new SqlCommandBuilder();
        SqlDataAdapter adapter6 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder6 = new SqlCommandBuilder();
        SqlDataAdapter adapter7 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder7 = new SqlCommandBuilder();
        SqlDataAdapter adapter8 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder8 = new SqlCommandBuilder();
        SqlDataAdapter adapter9 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder9 = new SqlCommandBuilder();
        SqlDataAdapter adapter10 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder10 = new SqlCommandBuilder();
        SqlDataAdapter adapter11 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder11 = new SqlCommandBuilder();
        SqlDataAdapter adapter12 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder12 = new SqlCommandBuilder();
        SqlDataAdapter adapter13 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder13 = new SqlCommandBuilder();
        SqlDataAdapter adapter14 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder14 = new SqlCommandBuilder();
        SqlDataAdapter adapter15 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder15 = new SqlCommandBuilder();
        SqlDataAdapter adapter16 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder16 = new SqlCommandBuilder();
        SqlDataAdapter adapter17 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder17 = new SqlCommandBuilder();
        SqlDataAdapter adapter18 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder18 = new SqlCommandBuilder();
        SqlDataAdapter adapter19 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder19 = new SqlCommandBuilder();
        SqlDataAdapter adapter20= new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder20 = new SqlCommandBuilder();

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();
        DataSet ds5 = new DataSet();
        DataSet ds6 = new DataSet();
        DataSet ds7 = new DataSet();
        DataSet ds8 = new DataSet();
        DataSet ds9 = new DataSet();
        DataSet ds10= new DataSet();
        DataSet ds13 = new DataSet();
        DataSet ds14 = new DataSet();
        DataSet ds15 = new DataSet();
        DataSet ds16 = new DataSet();
        DataSet ds17 = new DataSet();
        DataSet ds18 = new DataSet();
        DataSet ds19 = new DataSet();
        DataSet ds20 = new DataSet();
        DataSet ds21 = new DataSet();

        DataSet dsBOMMC = new DataSet();
        DataSet dsBOMMD = new DataSet();

        DataTable dt = new DataTable();
        string tablename = null;
        int result;
        string MANU= "新廠製二組";

        string ID1;
        DateTime dt1;
        string DELID1;
        string DELMOCTA001A;
        string DELMOCTA002A;
        string IN1="20001";
        string ID2;
        DateTime dt2;
        string DELID2;
        string DELMOCTA001B;
        string DELMOCTA002B;
        string IN2 = "20001";
        string ID3;
        DateTime dt3;
        string DELID3;
        string DELMOCTA001C;
        string DELMOCTA002C;
        string IN3 = "20001";
        string ID4;
        DateTime dt4;
        string DELID4;
        string DELMOCTA001D;
        string DELMOCTA002D;
        string IN4 = "20001";
        DateTime dt5;
        string DELID5;
        string DELMOCTA001E;
        string DELMOCTA002E;

        string TA001 = "A510";
        string TA002;
        string TA029;
        string MB001;
        string MB002;
        string MB003;
        string MB001B;
        string MB002B;
        string MB003B;
        string MB001C;
        string MB002C;
        string MB003C;
        string MB001D;
        string MB002D;
        string MB003D;
        string MB001E;
        string MB002E;
        string MB003E;
        decimal BAR;
        decimal BOX;
        decimal BAR2;
        decimal BAR3;
        decimal SUM1;
        decimal SUM2;
        decimal SUM3;
        decimal SUM4;

        string BOMVARSION;
        string UNIT;
        decimal BOMBAR;
        int BOXNUMERB;
        int MOCBOX;

        string SUBID;
        string SUBBAR;
        string SUBNUM;
        string SUBBOX;
        string SUBPACKAGE;
        string SUBID2;
        string SUBBAR2;
        string SUBNUM2;
        string SUBBOX2;
        string SUBPACKAGE2;
        string SUBID3;
        string SUBBAR3;
        string SUBNUM3;
        string SUBBOX3;
        string SUBPACKAGE3;
        string SUBID4;
        string SUBBAR4;
        string SUBNUM4;
        string SUBBOX4;
        string SUBPACKAGE4;

        public class MOCTADATA
        {
            public string COMPANY;
            public string CREATOR;
            public string USR_GROUP;
            public string CREATE_DATE;
            public string MODIFIER;
            public string MODI_DATE;
            public string FLAG;
            public string CREATE_TIME;
            public string MODI_TIME;
            public string TRANS_TYPE;
            public string TRANS_NAME;
            public string sync_count;
            public string DataGroup;
            public string TA001;
            public string TA002;
            public string TA003;
            public string TA004;
            public string TA005;
            public string TA006;
            public string TA007;
            public string TA009;
            public string TA010;
            public string TA011;
            public string TA012;
            public string TA013;
            public string TA014;
            public string TA015;
            public string TA016;
            public string TA017;
            public string TA018;
            public string TA019;
            public string TA020;
            public string TA021;
            public string TA022;
            public string TA024;
            public string TA025;
            public string TA029;
            public string TA030;
            public string TA031;
            public string TA034;
            public string TA035;
            public string TA040;
            public string TA041;
            public string TA043;
            public string TA044;
            public string TA045;
            public string TA046;
            public string TA047;
            public string TA049;
            public string TA050;
            public string TA200;
        }

        public class MOCTBDATA
        {
            public string COMPANY;
            public string CREATOR;
            public string USR_GROUP;
            public string CREATE_DATE;
            public string MODIFIER;
            public string MODI_DATE;
            public string FLAG;
            public string CREATE_TIME;
            public string MODI_TIME;
            public string TRANS_TYPE;
            public string TRANS_NAME;
            public string sync_count;
            public string DataGroup;
            
        }

        Thread TD;
        public frmMOCMANULINE()
        {
            InitializeComponent();
            comboBox1load();
            comboBox2load();
            comboBox3load();
            comboBox4load();
            comboBox5load();
            comboBox6load();
            comboBox7load();
            comboBox8load();
            SETIN();
        }
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        #region FUNCTION
        public void comboBox1load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MD001,MD002 FROM [TK].dbo.CMSMD    WHERE MD002 LIKE '新廠製二組%'   ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MD001", typeof(string));
            dt.Columns.Add("MD002", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "MD002";
            comboBox1.DisplayMember = "MD002";
            sqlConn.Close();


        }

        public void comboBox2load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MD001,MD002 FROM [TK].dbo.CMSMD    WHERE MD002 LIKE '新廠包裝線%'   ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MD001", typeof(string));
            dt.Columns.Add("MD002", typeof(string));
            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "MD002";
            comboBox2.DisplayMember = "MD002";
            sqlConn.Close();


        }

        public void comboBox3load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MD001,MD002 FROM [TK].dbo.CMSMD    WHERE MD002 LIKE '新廠製一組%'   ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MD001", typeof(string));
            dt.Columns.Add("MD002", typeof(string));
            da.Fill(dt);
            comboBox3.DataSource = dt.DefaultView;
            comboBox3.ValueMember = "MD002";
            comboBox3.DisplayMember = "MD002";
            sqlConn.Close();


        }
        public void comboBox4load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MD001,MD002 FROM [TK].dbo.CMSMD    WHERE MD002 LIKE '新廠製三組(手工)%'   ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MD001", typeof(string));
            dt.Columns.Add("MD002", typeof(string));
            da.Fill(dt);
            comboBox4.DataSource = dt.DefaultView;
            comboBox4.ValueMember = "MD002";
            comboBox4.DisplayMember = "MD002";
            sqlConn.Close();


        }

        public void comboBox5load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MC001 ,MC001+MC002 AS 'MC002' FROM [TK].dbo.CMSMC WHERE MC001 LIKE '2000%'  ORDER BY MC001 ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MC001", typeof(string));
            dt.Columns.Add("MC002", typeof(string));

            da.Fill(dt);
            comboBox5.DataSource = dt.DefaultView;
            comboBox5.ValueMember = "MC001";
            comboBox5.DisplayMember = "MC002";
            sqlConn.Close();

           

        }
        public void comboBox6load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MC001 ,MC001+MC002 AS 'MC002' FROM [TK].dbo.CMSMC WHERE MC001 LIKE '2000%'  ORDER BY MC001 ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MC001", typeof(string));
            dt.Columns.Add("MC002", typeof(string));

            da.Fill(dt);
            comboBox6.DataSource = dt.DefaultView;
            comboBox6.ValueMember = "MC001";
            comboBox6.DisplayMember = "MC002";
            sqlConn.Close();

           
        }
        public void comboBox7load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MC001 ,MC001+MC002 AS 'MC002' FROM [TK].dbo.CMSMC WHERE MC001 LIKE '2000%'  ORDER BY MC001 ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MC001", typeof(string));
            dt.Columns.Add("MC002", typeof(string));

            da.Fill(dt);
            comboBox7.DataSource = dt.DefaultView;
            comboBox7.ValueMember = "MC001";
            comboBox7.DisplayMember = "MC002";
            sqlConn.Close();

           


        }
        public void comboBox8load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MC001 ,MC001+MC002 AS 'MC002' FROM [TK].dbo.CMSMC WHERE MC001 LIKE '2000%'  ORDER BY MC001 ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MC001", typeof(string));
            dt.Columns.Add("MC002", typeof(string));

            da.Fill(dt);
            comboBox8.DataSource = dt.DefaultView;
            comboBox8.ValueMember = "MC001";
            comboBox8.DisplayMember = "MC002";
            sqlConn.Close();

           


        }

        public void SEARCHMOCMANULINE()
        {
            if(MANU.Equals("新廠製二組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();


                    sbSql.AppendFormat(@"  SELECT ");
                    sbSql.AppendFormat(@"  [MANU] AS '線別',CONVERT(varchar(100),[MANUDATE],112) AS '生產日',[MB001] AS '品號',[MB002] AS '品名'");
                    sbSql.AppendFormat(@"  ,[MB003] AS '規格',[BAR] AS '桶數',[NUM] AS '數量',[CLINET] AS '客戶'");
                    sbSql.AppendFormat(@"  ,[ID]");
                    sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINE]");
                    sbSql.AppendFormat(@"  WHERE [MANU]='{0}' ", MANU);
                    sbSql.AppendFormat(@"  AND CONVERT(varchar(100),[MANUDATE],112) LIKE '{0}%'", dateTimePicker1.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  ORDER BY [MANUDATE],[SERNO]");
                    sbSql.AppendFormat(@"  ");

                    adapter1= new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                    sqlConn.Open();
                    ds1.Clear();
                    adapter1.Fill(ds1, "TEMPds1");
                    sqlConn.Close();


                    if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                    {

                    }
                    else
                    {
                        if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                        {
                            //dataGridView1.Rows.Clear();
                            dataGridView1.DataSource = ds1.Tables["TEMPds1"];
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

            else if (MANU.Equals("新廠包裝線"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();


                    sbSql.AppendFormat(@"  SELECT ");
                    sbSql.AppendFormat(@"  [MANU] AS '線別',CONVERT(varchar(100),[MANUDATE],112) AS '生產日',[MB001] AS '品號',[MB002] AS '品名'");
                    sbSql.AppendFormat(@"  ,[MB003] AS '規格',[BOX] AS '箱數',[PACKAGE] AS '包裝數',[CLINET] AS '客戶',[MANUHOUR] AS '生產時間',[OUTDATE] AS '交期'");
                    sbSql.AppendFormat(@"  ,[ID]");
                    sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINE]");
                    sbSql.AppendFormat(@"  WHERE [MANU]='{0}' ", MANU);
                    sbSql.AppendFormat(@"  AND CONVERT(varchar(100),[MANUDATE],112) LIKE '{0}%'", dateTimePicker3.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  ORDER BY [MANUDATE],[SERNO]");
                    sbSql.AppendFormat(@"  ");

                    adapter7 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder7 = new SqlCommandBuilder(adapter7);
                    sqlConn.Open();
                    ds5.Clear();
                    adapter7.Fill(ds5, "TEMPds5");
                    sqlConn.Close();


                    if (ds5.Tables["TEMPds5"].Rows.Count == 0)
                    {

                    }
                    else
                    {
                        if (ds5.Tables["TEMPds5"].Rows.Count >= 1)
                        {
                            //dataGridView1.Rows.Clear();
                            dataGridView3.DataSource = ds5.Tables["TEMPds5"];
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
            else if (MANU.Equals("新廠製一組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();


                    sbSql.AppendFormat(@"  SELECT ");
                    sbSql.AppendFormat(@"  [MANU] AS '線別',CONVERT(varchar(100),[MANUDATE],112) AS '生產日',[MB001] AS '品號',[MB002] AS '品名'");
                    sbSql.AppendFormat(@"  ,[MB003] AS '規格',[BAR] AS '桶數',[NUM] AS '數量',[CLINET] AS '客戶'");
                    sbSql.AppendFormat(@"  ,[ID]");
                    sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINE]");
                    sbSql.AppendFormat(@"  WHERE [MANU]='{0}' ", MANU);
                    sbSql.AppendFormat(@"  AND CONVERT(varchar(100),[MANUDATE],112) LIKE '{0}%'", dateTimePicker6.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  ORDER BY [MANUDATE],[SERNO]");
                    sbSql.AppendFormat(@"  ");

                    adapter9 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder9 = new SqlCommandBuilder(adapter9);
                    sqlConn.Open();
                    ds7.Clear();
                    adapter9.Fill(ds7, "TEMPds7");
                    sqlConn.Close();


                    if (ds7.Tables["TEMPds7"].Rows.Count == 0)
                    {

                    }
                    else
                    {
                        if (ds7.Tables["TEMPds7"].Rows.Count >= 1)
                        {
                            //dataGridView1.Rows.Clear();
                            dataGridView5.DataSource = ds7.Tables["TEMPds7"];
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
            else if (MANU.Equals("新廠製三組(手工)"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();


                    sbSql.AppendFormat(@"  SELECT ");
                    sbSql.AppendFormat(@"  [MANU] AS '線別',CONVERT(varchar(100),[MANUDATE],112) AS '生產日',[MB001] AS '品號',[MB002] AS '品名'");
                    sbSql.AppendFormat(@"  ,[MB003] AS '規格',[BAR] AS '桶數',[NUM] AS '數量',[CLINET] AS '客戶'");
                    sbSql.AppendFormat(@"  ,[ID]");
                    sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINE]");
                    sbSql.AppendFormat(@"  WHERE [MANU]='{0}' ", MANU);
                    sbSql.AppendFormat(@"  AND CONVERT(varchar(100),[MANUDATE],112) LIKE '{0}%'", dateTimePicker8.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  ORDER BY [MANUDATE],[SERNO]");
                    sbSql.AppendFormat(@"  ");

                    adapter10 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder10 = new SqlCommandBuilder(adapter10);
                    sqlConn.Open();
                    ds8.Clear();
                    adapter10.Fill(ds8, "TEMPds8");
                    sqlConn.Close();


                    if (ds8.Tables["TEMPds8"].Rows.Count == 0)
                    {

                    }
                    else
                    {
                        if (ds8.Tables["TEMPds8"].Rows.Count >= 1)
                        {
                            //dataGridView1.Rows.Clear();
                            dataGridView7.DataSource = ds8.Tables["TEMPds8"];
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

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            SEARCHMB001();
            SEARCHBOMMD();
        }

        public void SEARCHMB001()
        {
            if (MANU.Equals("新廠製二組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"  SELECT MB001,MB002,MB003,MC004,MB017            ");
                    sbSql.AppendFormat(@"  FROM [TK].dbo.INVMB,[TK].dbo.BOMMC");
                    sbSql.AppendFormat(@"  WHERE MB001=MC001");
                    sbSql.AppendFormat(@"  AND MB001='{0}'", textBox1.Text);
                    sbSql.AppendFormat(@"  ");

                    adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                    sqlConn.Open();
                    ds2.Clear();
                    adapter2.Fill(ds2, "TEMPds2");
                    sqlConn.Close();


                    if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                    {
                        SETNULL1();
                    }
                    else
                    {
                        if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                        {
                            textBox2.Text = ds2.Tables["TEMPds2"].Rows[0]["MB002"].ToString();
                            textBox3.Text = ds2.Tables["TEMPds2"].Rows[0]["MB003"].ToString();
                            textBox32.Text = ds2.Tables["TEMPds2"].Rows[0]["MC004"].ToString();
                            //comboBox5.SelectedValue = ds2.Tables["TEMPds2"].Rows[0]["MB017"].ToString();
                            //label51.Text = ds2.Tables["TEMPds2"].Rows[0]["MB017"].ToString();
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
            else if (MANU.Equals("新廠包裝線"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"  SELECT MB001,MB002,MB003,MC004 ,MB017 ");
                    sbSql.AppendFormat(@"  FROM [TK].dbo.INVMB,[TK].dbo.BOMMC");
                    sbSql.AppendFormat(@"  WHERE MB001=MC001");
                    sbSql.AppendFormat(@"  AND MB001='{0}'", textBox7.Text);
                    sbSql.AppendFormat(@"  ");

                    adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                    sqlConn.Open();
                    ds2.Clear();
                    adapter2.Fill(ds2, "TEMPds2");
                    sqlConn.Close();


                    if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                    {
                        SETNULL1();
                    }
                    else
                    {
                        if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                        {
                            textBox10.Text = ds2.Tables["TEMPds2"].Rows[0]["MB002"].ToString();
                            textBox11.Text = ds2.Tables["TEMPds2"].Rows[0]["MB003"].ToString();
                            textBox33.Text = ds2.Tables["TEMPds2"].Rows[0]["MC004"].ToString();
                            //comboBox6.SelectedValue = ds2.Tables["TEMPds2"].Rows[0]["MB017"].ToString();
                            //label52.Text = ds2.Tables["TEMPds2"].Rows[0]["MB017"].ToString();

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

            else if (MANU.Equals("新廠製一組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"  SELECT MB001,MB002,MB003,MC004 ,MB017 ");
                    sbSql.AppendFormat(@"  FROM [TK].dbo.INVMB,[TK].dbo.BOMMC");
                    sbSql.AppendFormat(@"  WHERE MB001=MC001");
                    sbSql.AppendFormat(@"  AND MB001='{0}'", textBox14.Text);
                    sbSql.AppendFormat(@"  ");

                    adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                    sqlConn.Open();
                    ds2.Clear();
                    adapter2.Fill(ds2, "TEMPds2");
                    sqlConn.Close();


                    if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                    {
                        SETNULL4();
                    }
                    else
                    {
                        if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                        {
                            textBox17.Text = ds2.Tables["TEMPds2"].Rows[0]["MB002"].ToString();
                            textBox18.Text = ds2.Tables["TEMPds2"].Rows[0]["MB003"].ToString();
                            textBox34.Text = ds2.Tables["TEMPds2"].Rows[0]["MC004"].ToString();
                            //comboBox7.SelectedValue = ds2.Tables["TEMPds2"].Rows[0]["MB017"].ToString();
                            //label53.Text = ds2.Tables["TEMPds2"].Rows[0]["MB017"].ToString();
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
            else if (MANU.Equals("新廠製三組(手工)"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"  SELECT MB001,MB002,MB003,MC004 ,MB017 ");
                    sbSql.AppendFormat(@"  FROM [TK].dbo.INVMB,[TK].dbo.BOMMC");
                    sbSql.AppendFormat(@"  WHERE MB001=MC001");
                    sbSql.AppendFormat(@"  AND MB001='{0}'", textBox20.Text);
                    sbSql.AppendFormat(@"  ");

                    adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                    sqlConn.Open();
                    ds2.Clear();
                    adapter2.Fill(ds2, "TEMPds2");
                    sqlConn.Close();


                    if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                    {
                        SETNULL4();
                    }
                    else
                    {
                        if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                        {
                            textBox24.Text = ds2.Tables["TEMPds2"].Rows[0]["MB002"].ToString();
                            textBox25.Text = ds2.Tables["TEMPds2"].Rows[0]["MB003"].ToString();
                            textBox35.Text = ds2.Tables["TEMPds2"].Rows[0]["MC004"].ToString();
                            //comboBox8.SelectedValue = ds2.Tables["TEMPds2"].Rows[0]["MB017"].ToString();
                            //label54.Text = ds2.Tables["TEMPds2"].Rows[0]["MB017"].ToString();
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
        }

        public void SETNULL1()
        {
            //textBox1.Text = null;
            textBox2.Text = null;
            textBox3.Text = null;
            textBox4.Text = null;
            textBox5.Text = null;
            
        }
       
        public void ADDMOCMANULINE()
        {
            if(MANU.Equals("新廠製二組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();


                    sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[MOCMANULINE]");
                    sbSql.AppendFormat(" ([ID],[MANU],[MANUDATE],[MB001],[MB002],[MB003],[BAR],[NUM],[CLINET])");
                    sbSql.AppendFormat(" VALUES ({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}')", "NEWID()", comboBox1.Text, dateTimePicker2.Value.ToString("yyyy/MM/dd"), textBox1.Text, textBox2.Text, textBox3.Text, textBox4.Text, textBox5.Text, textBox6.Text);
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
            else if (MANU.Equals("新廠包裝線"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();


                    sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[MOCMANULINE]");
                    sbSql.AppendFormat(" ([ID],[MANU],[MANUDATE],[MB001],[MB002],[MB003],[CLINET],[MANUHOUR],[BOX],[PACKAGE],[OUTDATE])");
                    sbSql.AppendFormat(" VALUES ({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}')", "NEWID()", comboBox2.Text, dateTimePicker4.Value.ToString("yyyy/MM/dd"), textBox7.Text, textBox10.Text, textBox11.Text, textBox9.Text, textBox13.Text, textBox8.Text, textBox12.Text, dateTimePicker5.Value.ToString("yyyy/MM/dd"));
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
            else if (MANU.Equals("新廠製一組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();


                    sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[MOCMANULINE]");
                    sbSql.AppendFormat(" ([ID],[MANU],[MANUDATE],[MB001],[MB002],[MB003],[BAR],[NUM],[CLINET])");
                    sbSql.AppendFormat(" VALUES ({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}')", "NEWID()", comboBox3.Text, dateTimePicker7.Value.ToString("yyyy/MM/dd"), textBox14.Text, textBox17.Text, textBox18.Text, textBox15.Text, textBox19.Text, textBox16.Text);
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
            else if (MANU.Equals("新廠製三組(手工)"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();


                    sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[MOCMANULINE]");
                    sbSql.AppendFormat(" ([ID],[MANU],[MANUDATE],[MB001],[MB002],[MB003],[BAR],[NUM],[CLINET])");
                    sbSql.AppendFormat(" VALUES ({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}')", "NEWID()", comboBox4.Text, dateTimePicker9.Value.ToString("yyyy/MM/dd"), textBox20.Text, textBox24.Text, textBox25.Text, textBox21.Text, textBox23.Text, textBox22.Text);
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

            SEARCHMOCMANULINE();
        }
        public void SETNULL2()
        {
            textBox1.Text = null;
            textBox2.Text = null;
            textBox3.Text = null;
            textBox4.Text = null;
            textBox5.Text = null;
            textBox6.Text = null;
        }
        public void SETNULL3()
        {
            textBox7.Text = null;
            textBox8.Text = null;
            textBox9.Text = null;
            textBox10.Text = null;
            textBox11.Text = null;
            textBox12.Text = null;
            textBox13.Text = "0";
        }
        public void SETNULL4()
        {
            textBox14.Text = null;
            textBox15.Text = null;
            textBox16.Text = null;
            textBox17.Text = null;
            textBox18.Text = null;
            textBox19.Text = null;
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            textBox40.Text = null;
            textBox41.Text = null;

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    ID1 = row.Cells["ID"].Value.ToString();
                    dt1=Convert.ToDateTime (row.Cells["生產日"].Value.ToString().Substring(0,4)+"/"+row.Cells["生產日"].Value.ToString().Substring(4, 2)+"/"+row.Cells["生產日"].Value.ToString().Substring(6, 2));
                    MB001= row.Cells["品號"].Value.ToString();
                    MB002 = row.Cells["品名"].Value.ToString();
                    MB003 = row.Cells["規格"].Value.ToString();
                    BAR = Convert.ToDecimal(row.Cells["桶數"].Value.ToString());
                    SUM1 = Convert.ToDecimal(row.Cells["數量"].Value.ToString());
                    TA029 = row.Cells["客戶"].Value.ToString();

                    SUBID = row.Cells["ID"].Value.ToString();
                    SUBBAR = row.Cells["桶數"].Value.ToString();
                    SUBNUM = row.Cells["數量"].Value.ToString();
                    SUBBOX= null;
                    SUBPACKAGE = null;

                    SEARCHMB017();
                    SEARCHMOCMANULINERESULT();

                    SEARCHMOCMANULINECOP();

;
                }
                else
                {
                    ID1 = null;
                    SUBID = null;
                    SUBBAR = null;
                    SUBNUM = null;
                    SUBBOX = null;
                    SUBPACKAGE = null;

                }
            }
        }
        
        public void DELMOCMANULINE()
        {
            if (MANU.Equals("新廠製二組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();
                    sbSql.AppendFormat("  DELETE [TKMOC].[dbo].[MOCMANULINE]");
                    sbSql.AppendFormat("  WHERE ID='{0}'", ID1);
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

            else if (MANU.Equals("新廠包裝線"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();
                    sbSql.AppendFormat("  DELETE [TKMOC].[dbo].[MOCMANULINE]");
                    sbSql.AppendFormat("  WHERE ID='{0}'", ID2);
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
            else if (MANU.Equals("新廠製一組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();
                    sbSql.AppendFormat("  DELETE [TKMOC].[dbo].[MOCMANULINE]");
                    sbSql.AppendFormat("  WHERE ID='{0}'", ID3);
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
            else if (MANU.Equals("新廠製三組(手工)"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();
                    sbSql.AppendFormat("  DELETE [TKMOC].[dbo].[MOCMANULINE]");
                    sbSql.AppendFormat("  WHERE ID='{0}'", ID4);
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

            SEARCHMOCMANULINE();
        }

        public void ADDMOCMANULINERESULT()
        {
            if (MANU.Equals("新廠製二組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();


                    sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[MOCMANULINERESULT]");
                    sbSql.AppendFormat(" ([SID],[MOCTA001],[MOCTA002])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}')", ID1, TA001, TA002);
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
            else if (MANU.Equals("新廠包裝線"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();


                    sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[MOCMANULINERESULT]");
                    sbSql.AppendFormat(" ([SID],[MOCTA001],[MOCTA002])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}')", ID2, TA001, TA002);
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
                
            else if(MANU.Equals("新廠製一組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();


                    sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[MOCMANULINERESULT]");
                    sbSql.AppendFormat(" ([SID],[MOCTA001],[MOCTA002])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}')", ID3, TA001, TA002);
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
            else if (MANU.Equals("新廠製三組(手工)"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();


                    sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[MOCMANULINERESULT]");
                    sbSql.AppendFormat(" ([SID],[MOCTA001],[MOCTA002])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}')", ID4, TA001, TA002);
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
        }

        public void ADDMOCTATB()
        {
            MOCTADATA MOCTA = new MOCTADATA();
            MOCTA = SETMOCTA();
            string MOCMB001 = null;
            decimal MOCTA004 = 0; ;
            string TA026 = null;
            string TA027 = null;

            if (MANU.Equals("新廠製二組"))
            {
                MOCMB001 = MB001;
                MOCTA004 = BAR;
                TA026 = textBox40.Text;
                TA027 = textBox41.Text;
            }
            else if (MANU.Equals("新廠包裝線"))
            {
                MOCMB001 = MB001B;
                MOCTA004 = BOX;
                TA026 = textBox42.Text;
                TA027 = textBox43.Text;
            }
            else if (MANU.Equals("新廠製一組"))
            {
                MOCMB001 = MB001C;
                MOCTA004 = BAR2;
                TA026 = textBox44.Text;
                TA027 = textBox45.Text;
            }
            else if (MANU.Equals("新廠製三組(手工)"))
            {
                MOCMB001 = MB001D;
                MOCTA004 = BAR3;
                TA026 = textBox46.Text;
                TA027 = textBox47.Text;
            }
            else if (MANU.Equals("水麵"))
            {
                MOCMB001 = MB001E;
                MOCTA004 = Convert.ToDecimal(textBox31.Text)/ BOMBAR;
            }

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();


                sbSql.AppendFormat(" INSERT INTO [TK].[dbo].[MOCTA]");
                sbSql.AppendFormat(" ([COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE]");
                sbSql.AppendFormat(" ,[TRANS_NAME],[sync_count],[DataGroup],[TA001],[TA002],[TA003],[TA004],[TA005],[TA006],[TA007]");
                sbSql.AppendFormat(" ,[TA009],[TA010],[TA011],[TA012],[TA013],[TA014],[TA015],[TA016],[TA017],[TA018]");
                sbSql.AppendFormat(" ,[TA019],[TA020],[TA021],[TA022],[TA024],[TA025],[TA029],[TA030],[TA031],[TA034],[TA035]");
                sbSql.AppendFormat(" ,[TA040],[TA041],[TA043],[TA044],[TA045],[TA046],[TA047],[TA049],[TA050],[TA200]");
                sbSql.AppendFormat(" ,[TA026],[TA027]");
                sbSql.AppendFormat(" )");
                sbSql.AppendFormat(" VALUES");
                sbSql.AppendFormat(" ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}',",MOCTA.COMPANY, MOCTA.CREATOR, MOCTA.USR_GROUP, MOCTA.CREATE_DATE, MOCTA.MODIFIER, MOCTA.MODI_DATE, MOCTA.FLAG, MOCTA.CREATE_TIME, MOCTA.MODI_TIME, MOCTA.TRANS_TYPE);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}',", MOCTA.TRANS_NAME, MOCTA.sync_count, MOCTA.DataGroup, MOCTA.TA001, MOCTA.TA002, MOCTA.TA003, MOCTA.TA004, MOCTA.TA005, MOCTA.TA006, MOCTA.TA007);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}',", MOCTA.TA009, MOCTA.TA010, MOCTA.TA011, MOCTA.TA012, MOCTA.TA013, MOCTA.TA014, MOCTA.TA015, MOCTA.TA016, MOCTA.TA017, MOCTA.TA018);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}',", MOCTA.TA019, MOCTA.TA020, MOCTA.TA021, MOCTA.TA022, MOCTA.TA024, MOCTA.TA025,MOCTA.TA029, MOCTA.TA030, MOCTA.TA031, MOCTA.TA034, MOCTA.TA035);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}',", MOCTA.TA040, MOCTA.TA041, MOCTA.TA043, MOCTA.TA044, MOCTA.TA045, MOCTA.TA046, MOCTA.TA047, MOCTA.TA049, MOCTA.TA050, MOCTA.TA200);
                sbSql.AppendFormat(" '{0}','{1}'",TA026,TA027);
                sbSql.AppendFormat(" )");
                sbSql.AppendFormat(" ");                
                sbSql.AppendFormat(" INSERT INTO [TK].dbo.[MOCTB]");
                sbSql.AppendFormat(" ([COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE]");
                sbSql.AppendFormat(" ,[TRANS_NAME],[sync_count],[DataGroup],[TB001],[TB002],[TB003],[TB004],[TB005],[TB006],[TB007]");
                sbSql.AppendFormat(" ,[TB009],[TB011],[TB012],[TB013],[TB014],[TB018],[TB019],[TB020],[TB022],[TB024]");
                sbSql.AppendFormat(" ,[TB025],[TB026],[TB027],[TB029],[TB030],[TB031],[TB501],[TB554],[TB556],[TB560])");
                sbSql.AppendFormat(" (SELECT ");
                sbSql.AppendFormat(" '{0}' [COMPANY],'{1}' [CREATOR],'{2}' [USR_GROUP],'{3}' [CREATE_DATE],'{4}' [MODIFIER],'{5}' [MODI_DATE],'{6}' [FLAG],'{7}' [CREATE_TIME],'{8}' [MODI_TIME],'{9}' [TRANS_TYPE]", MOCTA.COMPANY, MOCTA.CREATOR, MOCTA.USR_GROUP, MOCTA.CREATE_DATE, MOCTA.MODIFIER, MOCTA.MODI_DATE, MOCTA.FLAG, MOCTA.CREATE_TIME, MOCTA.MODI_TIME, MOCTA.TRANS_TYPE);
                sbSql.AppendFormat(" ,'{0}' [TRANS_NAME],{1} [sync_count],'{2}' [DataGroup],'{3}' [TB001],'{4}' [TB002],[BOMMD].MD003 [TB003],ROUND({5}*[BOMMD].MD006/[BOMMD].MD007*(1+[BOMMD].MD008),3) [TB004],0 [TB005],'****' [TB006],[INVMB].MB004  [TB007]", MOCTA.TRANS_NAME, MOCTA.sync_count, MOCTA.DataGroup, MOCTA.TA001, MOCTA.TA002, MOCTA004);
                sbSql.AppendFormat(" ,[INVMB].MB017 [TB009],'1' [TB011],[INVMB].MB002 [TB012],[INVMB].MB003 [TB013],[BOMMD].MD001 [TB014],'N' [TB018],'0' [TB019],'0' [TB020],'2' [TB022],'0' [TB024]");
                sbSql.AppendFormat(" ,'****' [TB025],'0' [TB026],'1' [TB027],'0' [TB029],'0' [TB030],'0' [TB031],'0' [TB501],'N' [TB554],'0' [TB556],'0' [TB560]");
                sbSql.AppendFormat(" FROM [TK].dbo.[BOMMD],[TK].dbo.[INVMB]");
                sbSql.AppendFormat(" WHERE [BOMMD].MD003=[INVMB].MB001");
                sbSql.AppendFormat(" AND MD001='{0}' AND ISNULL(MD012,'')='' )", MOCMB001);
                sbSql.AppendFormat(" ");
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

        public MOCTADATA SETMOCTA()
        {
            if (MANU.Equals("新廠製二組"))
            {
                SEARCHBOMMC();

                MOCTADATA MOCTA = new MOCTADATA();
                MOCTA.COMPANY = "TK";
                MOCTA.CREATOR = "000002";
                MOCTA.USR_GROUP = "103000";
                MOCTA.CREATE_DATE = dt1.ToString("yyyyMMdd");
                MOCTA.MODIFIER = "000002";
                MOCTA.MODI_DATE = dt1.ToString("yyyyMMdd");
                MOCTA.FLAG = "0";
                MOCTA.CREATE_TIME = DateTime.Now.ToString("HH:mm:dd");
                MOCTA.MODI_TIME = DateTime.Now.ToString("HH:mm:dd");
                MOCTA.TRANS_TYPE = "P001";
                MOCTA.TRANS_NAME = "MOCMI02";
                MOCTA.sync_count = "0";
                MOCTA.DataGroup = "103000";
                MOCTA.TA001 = "A510";
                MOCTA.TA002 = TA002;
                MOCTA.TA003 = dt1.ToString("yyyyMMdd");
                MOCTA.TA004 = dt1.ToString("yyyyMMdd");
                MOCTA.TA005 = BOMVARSION;
                MOCTA.TA006 = MB001;
                MOCTA.TA007 = UNIT;
                MOCTA.TA009 = dt1.ToString("yyyyMMdd");
                MOCTA.TA010 = dt1.ToString("yyyyMMdd");
                MOCTA.TA011 = "1";
                MOCTA.TA012 = dt1.ToString("yyyyMMdd");
                MOCTA.TA013 = "N";
                //MOCTA.TA014 = dt1.ToString("yyyyMMdd");
                MOCTA.TA014 = "";
                //MOCTA.TA015 = (BAR * BOMBAR).ToString();
                MOCTA.TA015 = SUM1.ToString();
                MOCTA.TA016 = "0";
                MOCTA.TA017 = "0";
                MOCTA.TA018 = "0";
                MOCTA.TA019 = "20";
                MOCTA.TA020 = IN1;
                MOCTA.TA021 = "02";
                MOCTA.TA022 = "0";
                MOCTA.TA024 = "A510";
                MOCTA.TA025 = TA002;
                MOCTA.TA029 = TA029;
                MOCTA.TA030 = "1";
                MOCTA.TA031 = "0";
                MOCTA.TA034 = MB002;
                MOCTA.TA035 = MB003;
                MOCTA.TA040 = dt1.ToString("yyyyMMdd");
                MOCTA.TA041 = "";
                MOCTA.TA043 = "1";
                MOCTA.TA044 = "N";
                MOCTA.TA045 = SUM1.ToString();
                MOCTA.TA046 = SUM1.ToString();
                MOCTA.TA047 = "0";
                MOCTA.TA049 = "0";
                MOCTA.TA050 = "0";
                MOCTA.TA200 = "1";


                return MOCTA;
            }

            else if (MANU.Equals("新廠包裝線"))
            {
                SEARCHBOMMC();

                MOCTADATA MOCTA = new MOCTADATA();
                MOCTA.COMPANY = "TK";
                MOCTA.CREATOR = "000002";
                MOCTA.USR_GROUP = "103000";
                MOCTA.CREATE_DATE = dt2.ToString("yyyyMMdd");
                MOCTA.MODIFIER = "000002";
                MOCTA.MODI_DATE = dt2.ToString("yyyyMMdd");
                MOCTA.FLAG = "0";
                MOCTA.CREATE_TIME = DateTime.Now.ToString("HH:mm:dd");
                MOCTA.MODI_TIME = DateTime.Now.ToString("HH:mm:dd");
                MOCTA.TRANS_TYPE = "P001";
                MOCTA.TRANS_NAME = "MOCMI02";
                MOCTA.sync_count = "0";
                MOCTA.DataGroup = "103000";
                MOCTA.TA001 = "A510";
                MOCTA.TA002 = TA002;
                MOCTA.TA003 = dt2.ToString("yyyyMMdd");
                MOCTA.TA004 = dt2.ToString("yyyyMMdd");
                MOCTA.TA005 = BOMVARSION;
                MOCTA.TA006 = MB001B;
                MOCTA.TA007 = UNIT;
                MOCTA.TA009 = dt2.ToString("yyyyMMdd");
                MOCTA.TA010 = dt2.ToString("yyyyMMdd");
                MOCTA.TA011 = "1";
                MOCTA.TA012 = dt2.ToString("yyyyMMdd");
                MOCTA.TA013 = "N";
                // MOCTA.TA014 = dt2.ToString("yyyyMMdd");
                MOCTA.TA014 = "";
                //MOCTA.TA015 = (BOX * BOMBAR).ToString();
                MOCTA.TA015 = SUM2.ToString();
                MOCTA.TA016 = "0";
                MOCTA.TA017 = "0";
                MOCTA.TA018 = "0";
                MOCTA.TA019 = "20";
                MOCTA.TA020 = IN2;
                MOCTA.TA021 = "09";
                MOCTA.TA022 = "0";
                MOCTA.TA024 = "A510";
                MOCTA.TA025 = TA002;
                MOCTA.TA029 = TA029;
                MOCTA.TA030 = "1";
                MOCTA.TA031 = "0";
                MOCTA.TA034 = MB002B;
                MOCTA.TA035 = MB003B;
                MOCTA.TA040 = dt2.ToString("yyyyMMdd");
                MOCTA.TA041 = "";
                MOCTA.TA043 = "1";
                MOCTA.TA044 = "N";
                MOCTA.TA045 = SUM2.ToString();
                MOCTA.TA046 = SUM2.ToString();
                MOCTA.TA047 = "0";
                MOCTA.TA049 = "0";
                MOCTA.TA050 = "0";
                MOCTA.TA200 = "1";


                return MOCTA;
            }

            else if (MANU.Equals("新廠製一組"))
            {
                SEARCHBOMMC();

                MOCTADATA MOCTA = new MOCTADATA();
                MOCTA.COMPANY = "TK";
                MOCTA.CREATOR = "000002";
                MOCTA.USR_GROUP = "103000";
                MOCTA.CREATE_DATE = dt3.ToString("yyyyMMdd");
                MOCTA.MODIFIER = "000002";
                MOCTA.MODI_DATE = dt3.ToString("yyyyMMdd");
                MOCTA.FLAG = "0";
                MOCTA.CREATE_TIME = DateTime.Now.ToString("HH:mm:dd");
                MOCTA.MODI_TIME = DateTime.Now.ToString("HH:mm:dd");
                MOCTA.TRANS_TYPE = "P001";
                MOCTA.TRANS_NAME = "MOCMI02";
                MOCTA.sync_count = "0";
                MOCTA.DataGroup = "103000";
                MOCTA.TA001 = "A510";
                MOCTA.TA002 = TA002;
                MOCTA.TA003 = dt3.ToString("yyyyMMdd");
                MOCTA.TA004 = dt3.ToString("yyyyMMdd");
                MOCTA.TA005 = BOMVARSION;
                MOCTA.TA006 = MB001C;
                MOCTA.TA007 = UNIT;
                MOCTA.TA009 = dt3.ToString("yyyyMMdd");
                MOCTA.TA010 = dt3.ToString("yyyyMMdd");
                MOCTA.TA011 = "1";
                MOCTA.TA012 = dt3.ToString("yyyyMMdd");
                MOCTA.TA013 = "N";
                //MOCTA.TA014 = dt3.ToString("yyyyMMdd");
                MOCTA.TA014 = "";
                //MOCTA.TA015 = (BAR2 * BOMBAR).ToString();
                MOCTA.TA015 = SUM3.ToString();
                MOCTA.TA016 = "0";
                MOCTA.TA017 = "0";
                MOCTA.TA018 = "0";
                MOCTA.TA019 = "20";
                MOCTA.TA020 = IN3;
                MOCTA.TA021 = "03";
                MOCTA.TA022 = "0";
                MOCTA.TA024 = "A510";
                MOCTA.TA025 = TA002;
                MOCTA.TA029 = TA029;
                MOCTA.TA030 = "1";
                MOCTA.TA031 = "0";
                MOCTA.TA034 = MB002C;
                MOCTA.TA035 = MB003C;
                MOCTA.TA040 = dt3.ToString("yyyyMMdd");
                MOCTA.TA041 = "";
                MOCTA.TA043 = "1";
                MOCTA.TA044 = "N";
                MOCTA.TA045 = SUM3.ToString();
                MOCTA.TA046 = SUM3.ToString();
                MOCTA.TA047 = "0";
                MOCTA.TA049 = "0";
                MOCTA.TA050 = "0";
                MOCTA.TA200 = "1";


                return MOCTA;
            }
            else if (MANU.Equals("新廠製三組(手工)"))
            {
                SEARCHBOMMC();

                MOCTADATA MOCTA = new MOCTADATA();
                MOCTA.COMPANY = "TK";
                MOCTA.CREATOR = "000002";
                MOCTA.USR_GROUP = "103000";
                MOCTA.CREATE_DATE = dt4.ToString("yyyyMMdd");
                MOCTA.MODIFIER = "000002";
                MOCTA.MODI_DATE = dt4.ToString("yyyyMMdd");
                MOCTA.FLAG = "0";
                MOCTA.CREATE_TIME = DateTime.Now.ToString("HH:mm:dd");
                MOCTA.MODI_TIME = DateTime.Now.ToString("HH:mm:dd");
                MOCTA.TRANS_TYPE = "P001";
                MOCTA.TRANS_NAME = "MOCMI02";
                MOCTA.sync_count = "0";
                MOCTA.DataGroup = "103000";
                MOCTA.TA001 = "A510";
                MOCTA.TA002 = TA002;
                MOCTA.TA003 = dt4.ToString("yyyyMMdd");
                MOCTA.TA004 = dt4.ToString("yyyyMMdd");
                MOCTA.TA005 = BOMVARSION;
                MOCTA.TA006 = MB001D;
                MOCTA.TA007 = UNIT;
                MOCTA.TA009 = dt4.ToString("yyyyMMdd");
                MOCTA.TA010 = dt4.ToString("yyyyMMdd");
                MOCTA.TA011 = "1";
                MOCTA.TA012 = dt4.ToString("yyyyMMdd");
                MOCTA.TA013 = "N";
                //MOCTA.TA014 = dt4.ToString("yyyyMMdd");
                MOCTA.TA014 = "";
                //MOCTA.TA015 = (BAR3 * BOMBAR).ToString();
                MOCTA.TA015 = SUM4.ToString();
                MOCTA.TA016 = "0";
                MOCTA.TA017 = "0";
                MOCTA.TA018 = "0";
                MOCTA.TA019 = "20";
                MOCTA.TA020 = IN4;
                MOCTA.TA021 = "04";
                MOCTA.TA022 = "0";
                MOCTA.TA024 = "A510";
                MOCTA.TA025 = TA002;
                MOCTA.TA029 = TA029;
                MOCTA.TA030 = "1";
                MOCTA.TA031 = "0";
                MOCTA.TA034 = MB002D;
                MOCTA.TA035 = MB003D;
                MOCTA.TA040 = dt4.ToString("yyyyMMdd");
                MOCTA.TA041 = "";
                MOCTA.TA043 = "1";
                MOCTA.TA044 = "N";
                MOCTA.TA045 = SUM4.ToString();
                MOCTA.TA046 = SUM4.ToString();
                MOCTA.TA047 = "0";
                MOCTA.TA049 = "0";
                MOCTA.TA050 = "0";
                MOCTA.TA200 = "1";


                return MOCTA;
            }
            else if (MANU.Equals("水麵"))
            {
                SEARCHBOMMC();

                MOCTADATA MOCTA = new MOCTADATA();
                MOCTA.COMPANY = "TK";
                MOCTA.CREATOR = "000002";
                MOCTA.USR_GROUP = "103000";
                MOCTA.CREATE_DATE = dt5.ToString("yyyyMMdd");
                MOCTA.MODIFIER = "000002";
                MOCTA.MODI_DATE = dt5.ToString("yyyyMMdd");
                MOCTA.FLAG = "0";
                MOCTA.CREATE_TIME = DateTime.Now.ToString("HH:mm:dd");
                MOCTA.MODI_TIME = DateTime.Now.ToString("HH:mm:dd");
                MOCTA.TRANS_TYPE = "P001";
                MOCTA.TRANS_NAME = "MOCMI02";
                MOCTA.sync_count = "0";
                MOCTA.DataGroup = "103000";
                MOCTA.TA001 = "A510";
                MOCTA.TA002 = TA002;
                MOCTA.TA003 = dt5.ToString("yyyyMMdd");
                MOCTA.TA004 = dt5.ToString("yyyyMMdd");
                MOCTA.TA005 = BOMVARSION;
                MOCTA.TA006 = MB001E;
                MOCTA.TA007 = UNIT;
                MOCTA.TA009 = dt5.ToString("yyyyMMdd");
                MOCTA.TA010 = dt5.ToString("yyyyMMdd");
                MOCTA.TA011 = "1";
                MOCTA.TA012 = dt5.ToString("yyyyMMdd");
                MOCTA.TA013 = "N";
                //MOCTA.TA014 = dt5.ToString("yyyyMMdd");
                MOCTA.TA014 = "";
                MOCTA.TA015 = textBox31.Text;
                MOCTA.TA016 = "0";
                MOCTA.TA017 = "0";
                MOCTA.TA018 = "0";
                MOCTA.TA019 = "20";
                MOCTA.TA020 = textBox36.Text;
                MOCTA.TA021 = textBox27.Text;
                MOCTA.TA022 = "0";
                MOCTA.TA024 = "A510";
                MOCTA.TA025 = TA002;
                MOCTA.TA029 = "";
                MOCTA.TA030 = "1";
                MOCTA.TA031 = "0";
                MOCTA.TA034 = MB002E;
                MOCTA.TA035 = MB003E;
                MOCTA.TA040 = dt4.ToString("yyyyMMdd");
                MOCTA.TA041 = "";
                MOCTA.TA043 = "1";
                MOCTA.TA044 = "N";
                MOCTA.TA045 = (BAR3 * BOMBAR).ToString();
                MOCTA.TA046 = (BAR3 * BOMBAR).ToString();
                MOCTA.TA047 = "0";
                MOCTA.TA049 = "0";
                MOCTA.TA050 = "0";
                MOCTA.TA200 = "1";


                return MOCTA;
            }
            return null;
            
        }

        public void SEARCHBOMMC()
        {
            BOMVARSION = null;
            UNIT = null;
            BOMBAR = 0;

            if (MANU.Equals("新廠製二組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"  SELECT ");
                    sbSql.AppendFormat(@"  [MC001],[MC002],[MC003],[MC004],[MC005],[MC006],[MC007],[MC008],[MC009],[MC010],[MC011],[MC012],[MC013],[MC014],[MC015],[MC016],[MC017],[MC018],[MC019],[MC020],[MC021],[MC022],[MC023],[MC024],[MC025],[MC026],[MC027]");
                    sbSql.AppendFormat(@"  ,INVMB.MB004");
                    sbSql.AppendFormat(@"  FROM [TK].[dbo].[BOMMC]");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.[INVMB] ON MB001=MC001");
                    sbSql.AppendFormat(@"  WHERE  [MC001]='{0}'", MB001);
                    sbSql.AppendFormat(@"  ");

                    adapter5 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder5 = new SqlCommandBuilder(adapter5);
                    sqlConn.Open();
                    dsBOMMC.Clear();
                    adapter5.Fill(dsBOMMC, "dsBOMMC");
                    sqlConn.Close();


                    if (dsBOMMC.Tables["dsBOMMC"].Rows.Count == 0)
                    {
                        BOMVARSION = null;
                        UNIT = null;
                        BOMBAR = 0;
                    }
                    else
                    {
                        if (dsBOMMC.Tables["dsBOMMC"].Rows.Count >= 1)
                        {
                            BOMVARSION = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MC009"].ToString();
                            //UNIT = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MC002"].ToString();
                            UNIT = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MB004"].ToString();
                            BOMBAR = Convert.ToDecimal(dsBOMMC.Tables["dsBOMMC"].Rows[0]["MC004"].ToString());

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
            else if (MANU.Equals("新廠包裝線"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"  SELECT ");
                    sbSql.AppendFormat(@"  [MC001],[MC002],[MC003],[MC004],[MC005],[MC006],[MC007],[MC008],[MC009],[MC010],[MC011],[MC012],[MC013],[MC014],[MC015],[MC016],[MC017],[MC018],[MC019],[MC020],[MC021],[MC022],[MC023],[MC024],[MC025],[MC026],[MC027]");
                    sbSql.AppendFormat(@"  ,INVMB.MB004");
                    sbSql.AppendFormat(@"  FROM [TK].[dbo].[BOMMC]");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.[INVMB] ON MB001=MC001");
                    sbSql.AppendFormat(@"  WHERE  [MC001]='{0}'", MB001B);
                    sbSql.AppendFormat(@"  ");

                    adapter5 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder5 = new SqlCommandBuilder(adapter5);
                    sqlConn.Open();
                    dsBOMMC.Clear();
                    adapter5.Fill(dsBOMMC, "dsBOMMC");
                    sqlConn.Close();


                    if (dsBOMMC.Tables["dsBOMMC"].Rows.Count == 0)
                    {
                        BOMVARSION = null;
                        UNIT = null;
                        BOMBAR = 0;
                    }
                    else
                    {
                        if (dsBOMMC.Tables["dsBOMMC"].Rows.Count >= 1)
                        {
                            BOMVARSION = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MC009"].ToString();
                            //UNIT = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MC002"].ToString();
                            UNIT = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MB004"].ToString();
                            BOMBAR = Convert.ToDecimal(dsBOMMC.Tables["dsBOMMC"].Rows[0]["MC004"].ToString());

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
            else if (MANU.Equals("新廠製一組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"  SELECT ");
                    sbSql.AppendFormat(@"  [MC001],[MC002],[MC003],[MC004],[MC005],[MC006],[MC007],[MC008],[MC009],[MC010],[MC011],[MC012],[MC013],[MC014],[MC015],[MC016],[MC017],[MC018],[MC019],[MC020],[MC021],[MC022],[MC023],[MC024],[MC025],[MC026],[MC027]");
                    sbSql.AppendFormat(@"  ,INVMB.MB004");
                    sbSql.AppendFormat(@"  FROM [TK].[dbo].[BOMMC]");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.[INVMB] ON MB001=MC001");
                    sbSql.AppendFormat(@"  WHERE  [MC001]='{0}'", MB001C);
                    sbSql.AppendFormat(@"  ");

                    adapter5 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder5 = new SqlCommandBuilder(adapter5);
                    sqlConn.Open();
                    dsBOMMC.Clear();
                    adapter5.Fill(dsBOMMC, "dsBOMMC");
                    sqlConn.Close();


                    if (dsBOMMC.Tables["dsBOMMC"].Rows.Count == 0)
                    {
                        BOMVARSION = null;
                        UNIT = null;
                        BOMBAR = 0;
                    }
                    else
                    {
                        if (dsBOMMC.Tables["dsBOMMC"].Rows.Count >= 1)
                        {
                            BOMVARSION = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MC009"].ToString();
                            //UNIT = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MC002"].ToString();
                            UNIT = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MB004"].ToString();
                            BOMBAR = Convert.ToDecimal(dsBOMMC.Tables["dsBOMMC"].Rows[0]["MC004"].ToString());

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
            else if (MANU.Equals("新廠製三組(手工)"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"  SELECT ");
                    sbSql.AppendFormat(@"  [MC001],[MC002],[MC003],[MC004],[MC005],[MC006],[MC007],[MC008],[MC009],[MC010],[MC011],[MC012],[MC013],[MC014],[MC015],[MC016],[MC017],[MC018],[MC019],[MC020],[MC021],[MC022],[MC023],[MC024],[MC025],[MC026],[MC027]");
                    sbSql.AppendFormat(@"  ,INVMB.MB004");
                    sbSql.AppendFormat(@"  FROM [TK].[dbo].[BOMMC]");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.[INVMB] ON MB001=MC001");
                    sbSql.AppendFormat(@"  WHERE  [MC001]='{0}'", MB001D);
                    sbSql.AppendFormat(@"  ");

                    adapter5 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder5 = new SqlCommandBuilder(adapter5);
                    sqlConn.Open();
                    dsBOMMC.Clear();
                    adapter5.Fill(dsBOMMC, "dsBOMMC");
                    sqlConn.Close();


                    if (dsBOMMC.Tables["dsBOMMC"].Rows.Count == 0)
                    {
                        BOMVARSION = null;
                        UNIT = null;
                        BOMBAR = 0;
                    }
                    else
                    {
                        if (dsBOMMC.Tables["dsBOMMC"].Rows.Count >= 1)
                        {
                            BOMVARSION = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MC009"].ToString();
                            //UNIT = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MC002"].ToString();
                            UNIT = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MB004"].ToString();
                            BOMBAR = Convert.ToDecimal(dsBOMMC.Tables["dsBOMMC"].Rows[0]["MC004"].ToString());

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
            else if (MANU.Equals("水麵"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"  SELECT ");
                    sbSql.AppendFormat(@"  [MC001],[MC002],[MC003],[MC004],[MC005],[MC006],[MC007],[MC008],[MC009],[MC010],[MC011],[MC012],[MC013],[MC014],[MC015],[MC016],[MC017],[MC018],[MC019],[MC020],[MC021],[MC022],[MC023],[MC024],[MC025],[MC026],[MC027]");
                    sbSql.AppendFormat(@"  ,INVMB.MB004");
                    sbSql.AppendFormat(@"  FROM [TK].[dbo].[BOMMC]");
                    sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.[INVMB] ON MB001=MC001");
                    sbSql.AppendFormat(@"  WHERE  [MC001]='{0}'", MB001E);
                    sbSql.AppendFormat(@"  ");

                    adapter5 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder5 = new SqlCommandBuilder(adapter5);
                    sqlConn.Open();
                    dsBOMMC.Clear();
                    adapter5.Fill(dsBOMMC, "dsBOMMC");
                    sqlConn.Close();


                    if (dsBOMMC.Tables["dsBOMMC"].Rows.Count == 0)
                    {
                        BOMVARSION = null;
                        UNIT = null;
                        BOMBAR = 0;
                    }
                    else
                    {
                        if (dsBOMMC.Tables["dsBOMMC"].Rows.Count >= 1)
                        {
                            BOMVARSION = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MC009"].ToString();
                            //UNIT = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MC002"].ToString();
                            UNIT = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MB004"].ToString();
                            BOMBAR = Convert.ToDecimal(dsBOMMC.Tables["dsBOMMC"].Rows[0]["MC004"].ToString());

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



        }
        public void SEARCHMOCMANULINERESULT()
        {
            if (MANU.Equals("新廠製二組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();


                    sbSql.AppendFormat(@"  SELECT  [MOCTA001] AS '製令',[MOCTA002]  AS '單號',[SID]");
                    sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINERESULT]");
                    sbSql.AppendFormat(@"  WHERE [SID]='{0}'", ID1);
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");

                    adapter3 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder3 = new SqlCommandBuilder(adapter3);
                    sqlConn.Open();
                    ds3.Clear();
                    adapter3.Fill(ds3, "TEMPds3");
                    sqlConn.Close();


                    if (ds3.Tables["TEMPds3"].Rows.Count == 0)
                    {

                    }
                    else
                    {
                        if (ds3.Tables["TEMPds3"].Rows.Count >= 1)
                        {

                            dataGridView2.DataSource = ds3.Tables["TEMPds3"];
                            dataGridView2.AutoResizeColumns();
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
            else if  (MANU.Equals("新廠包裝線"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();


                    sbSql.AppendFormat(@"  SELECT  [MOCTA001] AS '製令',[MOCTA002]  AS '單號',[SID]");
                    sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINERESULT]");
                    sbSql.AppendFormat(@"  WHERE [SID]='{0}'", ID2);
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");

                    adapter8 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder8 = new SqlCommandBuilder(adapter8);
                    sqlConn.Open();
                    ds6.Clear();
                    adapter8.Fill(ds6, "TEMPds6");
                    sqlConn.Close();


                    if (ds6.Tables["TEMPds6"].Rows.Count == 0)
                    {

                    }
                    else
                    {
                        if (ds6.Tables["TEMPds6"].Rows.Count >= 1)
                        {

                            dataGridView4.DataSource = ds6.Tables["TEMPds6"];
                            dataGridView4.AutoResizeColumns();
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
            else if (MANU.Equals("新廠製一組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();


                    sbSql.AppendFormat(@"  SELECT  [MOCTA001] AS '製令',[MOCTA002]  AS '單號',[SID]");
                    sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINERESULT]");
                    sbSql.AppendFormat(@"  WHERE [SID]='{0}'", ID3);
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");

                    adapter11 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder11 = new SqlCommandBuilder(adapter11);
                    sqlConn.Open();
                    ds9.Clear();
                    adapter11.Fill(ds9, "TEMPds9");
                    sqlConn.Close();


                    if (ds9.Tables["TEMPds9"].Rows.Count == 0)
                    {

                    }
                    else
                    {
                        if (ds9.Tables["TEMPds9"].Rows.Count >= 1)
                        {

                            dataGridView6.DataSource = ds9.Tables["TEMPds9"];
                            dataGridView6.AutoResizeColumns();
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
            else if (MANU.Equals("新廠製三組(手工)"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();


                    sbSql.AppendFormat(@"  SELECT  [MOCTA001] AS '製令',[MOCTA002]  AS '單號',[SID]");
                    sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINERESULT]");
                    sbSql.AppendFormat(@"  WHERE [SID]='{0}'", ID4);
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");

                    adapter12 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder12 = new SqlCommandBuilder(adapter12);
                    sqlConn.Open();
                    ds10.Clear();
                    adapter12.Fill(ds10, "TEMPds10");
                    sqlConn.Close();


                    if (ds10.Tables["TEMPds10"].Rows.Count == 0)
                    {

                    }
                    else
                    {
                        if (ds10.Tables["TEMPds10"].Rows.Count >= 1)
                        {

                            dataGridView8.DataSource = ds10.Tables["TEMPds10"];
                            dataGridView8.AutoResizeColumns();
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

        }

        public string GETMAXTA002(string TA001)
        {
            string TA002;

            if (MANU.Equals("新廠製二組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    StringBuilder sbSql = new StringBuilder();
                    sbSql.Clear();
                    sbSqlQuery.Clear();
                    ds4.Clear();

                    sbSql.AppendFormat(@"  SELECT ISNULL(MAX(TA002),'00000000000') AS TA002");
                    sbSql.AppendFormat(@"  FROM [TK].[dbo].[MOCTA] ");
                    //sbSql.AppendFormat(@"  WHERE  TC001='{0}' AND TC003='{1}'", "A542","20170119");
                    sbSql.AppendFormat(@"  WHERE  TA001='{0}' AND TA003='{1}'", TA001, dt1.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");

                    adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder4 = new SqlCommandBuilder(adapter4);
                    sqlConn.Open();
                    ds4.Clear();
                    adapter4.Fill(ds4, "TEMPds4");
                    sqlConn.Close();


                    if (ds4.Tables["TEMPds4"].Rows.Count == 0)
                    {
                        return null;
                    }
                    else
                    {
                        if (ds4.Tables["TEMPds4"].Rows.Count >= 1)
                        {
                            TA002 = SETTA002(ds4.Tables["TEMPds4"].Rows[0]["TA002"].ToString());
                            return TA002;

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
            else if (MANU.Equals("新廠包裝線"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    StringBuilder sbSql = new StringBuilder();
                    sbSql.Clear();
                    sbSqlQuery.Clear();
                    ds4.Clear();

                    sbSql.AppendFormat(@"  SELECT ISNULL(MAX(TA002),'00000000000') AS TA002");
                    sbSql.AppendFormat(@"  FROM [TK].[dbo].[MOCTA] ");
                    //sbSql.AppendFormat(@"  WHERE  TC001='{0}' AND TC003='{1}'", "A542","20170119");
                    sbSql.AppendFormat(@"  WHERE  TA001='{0}' AND TA003='{1}'", TA001, dt2.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");

                    adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder4 = new SqlCommandBuilder(adapter4);
                    sqlConn.Open();
                    ds4.Clear();
                    adapter4.Fill(ds4, "TEMPds4");
                    sqlConn.Close();


                    if (ds4.Tables["TEMPds4"].Rows.Count == 0)
                    {
                        return null;
                    }
                    else
                    {
                        if (ds4.Tables["TEMPds4"].Rows.Count >= 1)
                        {
                            TA002 = SETTA002(ds4.Tables["TEMPds4"].Rows[0]["TA002"].ToString());
                            return TA002;

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
            else if (MANU.Equals("新廠製一組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    StringBuilder sbSql = new StringBuilder();
                    sbSql.Clear();
                    sbSqlQuery.Clear();
                    ds4.Clear();

                    sbSql.AppendFormat(@"  SELECT ISNULL(MAX(TA002),'00000000000') AS TA002");
                    sbSql.AppendFormat(@"  FROM [TK].[dbo].[MOCTA] ");
                    //sbSql.AppendFormat(@"  WHERE  TC001='{0}' AND TC003='{1}'", "A542","20170119");
                    sbSql.AppendFormat(@"  WHERE  TA001='{0}' AND TA003='{1}'", TA001, dt3.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");

                    adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder4 = new SqlCommandBuilder(adapter4);
                    sqlConn.Open();
                    ds4.Clear();
                    adapter4.Fill(ds4, "TEMPds4");
                    sqlConn.Close();


                    if (ds4.Tables["TEMPds4"].Rows.Count == 0)
                    {
                        return null;
                    }
                    else
                    {
                        if (ds4.Tables["TEMPds4"].Rows.Count >= 1)
                        {
                            TA002 = SETTA002(ds4.Tables["TEMPds4"].Rows[0]["TA002"].ToString());
                            return TA002;

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
            else if (MANU.Equals("新廠製三組(手工)"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    StringBuilder sbSql = new StringBuilder();
                    sbSql.Clear();
                    sbSqlQuery.Clear();
                    ds4.Clear();

                    sbSql.AppendFormat(@"  SELECT ISNULL(MAX(TA002),'00000000000') AS TA002");
                    sbSql.AppendFormat(@"  FROM [TK].[dbo].[MOCTA] ");
                    //sbSql.AppendFormat(@"  WHERE  TC001='{0}' AND TC003='{1}'", "A542","20170119");
                    sbSql.AppendFormat(@"  WHERE  TA001='{0}' AND TA003='{1}'", TA001, dt4.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");

                    adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder4 = new SqlCommandBuilder(adapter4);
                    sqlConn.Open();
                    ds4.Clear();
                    adapter4.Fill(ds4, "TEMPds4");
                    sqlConn.Close();


                    if (ds4.Tables["TEMPds4"].Rows.Count == 0)
                    {
                        return null;
                    }
                    else
                    {
                        if (ds4.Tables["TEMPds4"].Rows.Count >= 1)
                        {
                            TA002 = SETTA002(ds4.Tables["TEMPds4"].Rows[0]["TA002"].ToString());
                            return TA002;

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
            else if (MANU.Equals("水麵"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    StringBuilder sbSql = new StringBuilder();
                    sbSql.Clear();
                    sbSqlQuery.Clear();
                    ds4.Clear();

                    sbSql.AppendFormat(@"  SELECT ISNULL(MAX(TA002),'00000000000') AS TA002");
                    sbSql.AppendFormat(@"  FROM [TK].[dbo].[MOCTA] ");
                    //sbSql.AppendFormat(@"  WHERE  TC001='{0}' AND TC003='{1}'", "A542","20170119");
                    sbSql.AppendFormat(@"  WHERE  TA001='{0}' AND TA003='{1}'", TA001, dt5.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");

                    adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder4 = new SqlCommandBuilder(adapter4);
                    sqlConn.Open();
                    ds4.Clear();
                    adapter4.Fill(ds4, "TEMPds4");
                    sqlConn.Close();


                    if (ds4.Tables["TEMPds4"].Rows.Count == 0)
                    {
                        return null;
                    }
                    else
                    {
                        if (ds4.Tables["TEMPds4"].Rows.Count >= 1)
                        {
                            TA002 = SETTA002(ds4.Tables["TEMPds4"].Rows[0]["TA002"].ToString());
                            return TA002;

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
            return null;

        }
        public string SETTA002(string TA002)
        {

            if (MANU.Equals("新廠製二組"))
            {
                if (TA002.Equals("00000000000"))
                {
                    return dt1.ToString("yyyyMMdd") + "001";
                }

                else
                {
                    int serno = Convert.ToInt16(TA002.Substring(8, 3));
                    serno = serno + 1;
                    string temp = serno.ToString();
                    temp = temp.PadLeft(3, '0');
                    return dt1.ToString("yyyyMMdd") + temp.ToString();
                }
            }

            else if (MANU.Equals("新廠包裝線"))
            {
                if (TA002.Equals("00000000000"))
                {
                    return dt2.ToString("yyyyMMdd") + "001";
                }

                else
                {
                    int serno = Convert.ToInt16(TA002.Substring(8, 3));
                    serno = serno + 1;
                    string temp = serno.ToString();
                    temp = temp.PadLeft(3, '0');
                    return dt2.ToString("yyyyMMdd") + temp.ToString();
                }
            }

            else if (MANU.Equals("新廠製一組"))
            {
                if (TA002.Equals("00000000000"))
                {
                    return dt3.ToString("yyyyMMdd") + "001";
                }

                else
                {
                    int serno = Convert.ToInt16(TA002.Substring(8, 3));
                    serno = serno + 1;
                    string temp = serno.ToString();
                    temp = temp.PadLeft(3, '0');
                    return dt3.ToString("yyyyMMdd") + temp.ToString();
                }
            }
            else if (MANU.Equals("新廠製三組(手工)"))
            {
                if (TA002.Equals("00000000000"))
                {
                    return dt4.ToString("yyyyMMdd") + "001";
                }

                else
                {
                    int serno = Convert.ToInt16(TA002.Substring(8, 3));
                    serno = serno + 1;
                    string temp = serno.ToString();
                    temp = temp.PadLeft(3, '0');
                    return dt4.ToString("yyyyMMdd") + temp.ToString();
                }
            }
            else if (MANU.Equals("水麵"))
            {
                if (TA002.Equals("00000000000"))
                {
                    return dt5.ToString("yyyyMMdd") + "001";
                }

                else
                {
                    int serno = Convert.ToInt16(TA002.Substring(8, 3));
                    serno = serno + 1;
                    string temp = serno.ToString();
                    temp = temp.PadLeft(3, '0');
                    return dt5.ToString("yyyyMMdd") + temp.ToString();
                }
            }

            return null;
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage1"])
            {
                MessageBox.Show("新廠製二組");
                MANU = "新廠製二組";
            }
            else if(tabControl1.SelectedTab == tabControl1.TabPages["tabPage2"])
            {
                MessageBox.Show("新廠製一組");
                MANU = "新廠製一組";
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage3"])
            {
                MessageBox.Show("新廠製三組(手工)");
                MANU = "新廠製三組(手工)";
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage4"])
            {
                MessageBox.Show("新廠包裝線");
                MANU = "新廠包裝線";
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage5"])
            {
                MessageBox.Show("水麵");
                MANU = "水麵";
            }
        }
        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            SEARCHMB001();
        }
        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            textBox42.Text = null;
            textBox43.Text = null;

            if (dataGridView3.CurrentRow != null)
            {
                int rowindex = dataGridView3.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView3.Rows[rowindex];
                    ID2 = row.Cells["ID"].Value.ToString();
                    dt2 = Convert.ToDateTime(row.Cells["生產日"].Value.ToString().Substring(0, 4) + "/" + row.Cells["生產日"].Value.ToString().Substring(4, 2) + "/" + row.Cells["生產日"].Value.ToString().Substring(6, 2));
                    MB001B = row.Cells["品號"].Value.ToString();
                    MB002B = row.Cells["品名"].Value.ToString();
                    MB003B = row.Cells["規格"].Value.ToString();
                    BOX = Convert.ToDecimal(row.Cells["包裝數"].Value.ToString());
                    SUM2 = Convert.ToDecimal(row.Cells["包裝數"].Value.ToString());
                    TA029 = row.Cells["客戶"].Value.ToString();

                    SUBID2 = row.Cells["ID"].Value.ToString();
                    SUBBAR2 = "";
                    SUBNUM2 = "";
                    SUBBOX2 = row.Cells["箱數"].Value.ToString();
                    SUBPACKAGE2 = row.Cells["包裝數"].Value.ToString();

                    SEARCHMOCMANULINERESULT();
                    SEARCHMOCMANULINECOP();
                    
                }
                else
                {
                    ID2 = null;
                    SUBID2 = null;
                    SUBBAR2 = null;
                    SUBNUM2 = null;
                    SUBBOX2= null;
                    SUBPACKAGE2 = null;

                }
            }
        }

        private void textBox14_TextChanged(object sender, EventArgs e)
        {
            SEARCHMB001();
            SEARCHBOMMD();
        }
        private void dataGridView5_SelectionChanged(object sender, EventArgs e)
        {
            textBox44.Text = null;
            textBox45.Text = null;

            if (dataGridView5.CurrentRow != null)
            {
                int rowindex = dataGridView5.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView5.Rows[rowindex];
                    ID3 = row.Cells["ID"].Value.ToString();
                    dt3 = Convert.ToDateTime(row.Cells["生產日"].Value.ToString().Substring(0, 4) + "/" + row.Cells["生產日"].Value.ToString().Substring(4, 2) + "/" + row.Cells["生產日"].Value.ToString().Substring(6, 2));
                    MB001C= row.Cells["品號"].Value.ToString();
                    MB002C = row.Cells["品名"].Value.ToString();
                    MB003C = row.Cells["規格"].Value.ToString();
                    BAR2 = Convert.ToDecimal(row.Cells["桶數"].Value.ToString());
                    SUM3 = Convert.ToDecimal(row.Cells["數量"].Value.ToString());
                    TA029 = row.Cells["客戶"].Value.ToString();

                    SUBID3 = row.Cells["ID"].Value.ToString();
                    SUBBAR3 = row.Cells["桶數"].Value.ToString();
                    SUBNUM3 = row.Cells["數量"].Value.ToString();
                    SUBBOX3 = null;
                    SUBPACKAGE3 = null;

                    SEARCHMOCMANULINERESULT();
                    SEARCHMOCMANULINECOP();
                    
                }
                else
                {
                    ID3 = null;
                    SUBID3 = null;
                    SUBBAR3 = null;
                    SUBNUM3 = null;
                    SUBBOX3 = null;
                    SUBPACKAGE3 = null;

                }
            }
        }
        private void textBox20_TextChanged(object sender, EventArgs e)
        {
            SEARCHMB001();
            SEARCHBOMMD();
        }
        private void dataGridView7_SelectionChanged(object sender, EventArgs e)
        {

            textBox46.Text = null;
            textBox47.Text = null;

            if (dataGridView7.CurrentRow != null)
            {
                int rowindex = dataGridView7.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView7.Rows[rowindex];
                    ID4 = row.Cells["ID"].Value.ToString();
                    dt4 = Convert.ToDateTime(row.Cells["生產日"].Value.ToString().Substring(0, 4) + "/" + row.Cells["生產日"].Value.ToString().Substring(4, 2) + "/" + row.Cells["生產日"].Value.ToString().Substring(6, 2));
                    MB001D = row.Cells["品號"].Value.ToString();
                    MB002D = row.Cells["品名"].Value.ToString();
                    MB003D = row.Cells["規格"].Value.ToString();
                    BAR3 = Convert.ToDecimal(row.Cells["桶數"].Value.ToString());
                    SUM4 = Convert.ToDecimal(row.Cells["數量"].Value.ToString());
                    TA029 = row.Cells["客戶"].Value.ToString();

                    SUBID4 = row.Cells["ID"].Value.ToString();
                    SUBBAR4 = row.Cells["桶數"].Value.ToString();
                    SUBNUM4 = row.Cells["數量"].Value.ToString();
                    SUBBOX4 = null;
                    SUBPACKAGE4 = null;

                    SEARCHMOCMANULINERESULT();
                    SEARCHMOCMANULINECOP();
                    
                }
                else
                {
                    ID4 = null;
                    SUBID4 = null;
                    SUBBAR4= null;
                    SUBNUM4 = null;
                    SUBBOX4 = null;
                    SUBPACKAGE4 = null;

                }
            }
        }
        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];
                    DELID1 = row.Cells["SID"].Value.ToString();
                    DELMOCTA001A= row.Cells["製令"].Value.ToString();
                    DELMOCTA002A = row.Cells["單號"].Value.ToString();

                }
                else
                {
                    DELID1 = null;

                }
            }
        }
        private void dataGridView4_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView4.CurrentRow != null)
            {
                int rowindex = dataGridView4.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView4.Rows[rowindex];
                    DELID2 = row.Cells["SID"].Value.ToString();
                    DELMOCTA001B = row.Cells["製令"].Value.ToString();
                    DELMOCTA002B = row.Cells["單號"].Value.ToString();



                }
                else
                {
                    DELID2 = null;

                }
            }
        }

        private void dataGridView6_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView6.CurrentRow != null)
            {
                int rowindex = dataGridView6.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView6.Rows[rowindex];
                    DELID3 = row.Cells["SID"].Value.ToString();
                    DELMOCTA001C = row.Cells["製令"].Value.ToString();
                    DELMOCTA002C = row.Cells["單號"].Value.ToString();



                }
                else
                {
                    DELID3 = null;

                }
            }
        }

        private void dataGridView8_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView8.CurrentRow != null)
            {
                int rowindex = dataGridView8.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView8.Rows[rowindex];
                    DELID4 = row.Cells["SID"].Value.ToString();
                    DELMOCTA001D = row.Cells["製令"].Value.ToString();
                    DELMOCTA002D = row.Cells["單號"].Value.ToString();



                }
                else
                {
                    DELID4 = null;

                }
            }
        }

        public void DELMOCMANULINERESULT()
        {
            if (MANU.Equals("新廠製二組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();
                    sbSql.AppendFormat("  DELETE [TKMOC].[dbo].[MOCMANULINERESULT]");
                    sbSql.AppendFormat("  WHERE SID='{0}' ", DELID1);
                    sbSql.AppendFormat("  AND [MOCTA001] ='{0}' AND [MOCTA002]='{1}'",DELMOCTA001A, DELMOCTA002A);
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

            else if (MANU.Equals("新廠包裝線"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();
                    sbSql.AppendFormat("  DELETE [TKMOC].[dbo].[MOCMANULINERESULT]");
                    sbSql.AppendFormat("  WHERE SID='{0}'", DELID2);
                    sbSql.AppendFormat("  AND [MOCTA001] ='{0}' AND [MOCTA002]='{1}'", DELMOCTA001B, DELMOCTA002B);
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
            else if (MANU.Equals("新廠製一組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();
                    sbSql.AppendFormat("  DELETE [TKMOC].[dbo].[MOCMANULINERESULT]");
                    sbSql.AppendFormat("  WHERE SID='{0}'", DELID3);
                    sbSql.AppendFormat("  AND [MOCTA001] ='{0}' AND [MOCTA002]='{1}'", DELMOCTA001C, DELMOCTA002C);
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
            else if (MANU.Equals("新廠製三組(手工)"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();
                    sbSql.AppendFormat("  DELETE [TKMOC].[dbo].[MOCMANULINERESULT]");
                    sbSql.AppendFormat("  WHERE SID='{0}'", DELID4);
                    sbSql.AppendFormat("  AND [MOCTA001] ='{0}' AND [MOCTA002]='{1}'", DELMOCTA001D, DELMOCTA002D);
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

        }

        public void SEARCHMOCTB()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT TA003  AS '日期',[TA021] AS '線別號',[MD002] AS '線別',TB003 AS '品號',TB012 AS '品名',SUM(TB004)  AS '總數量',TB009  AS '入庫別'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.MOCTB, [TK].dbo.MOCTA,[TK].dbo.CMSMD");
                sbSql.AppendFormat(@"  WHERE TA001=TB001 AND TA002=TB002");
                sbSql.AppendFormat(@"  AND [TA021]=MD001");
                sbSql.AppendFormat(@"  AND TB012 LIKE '%水麵%'");
                sbSql.AppendFormat(@"  AND TA003='{0}'",dateTimePicker10.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  GROUP BY TB003,TB012,TB009,TA003,[TA021],[MD002] ");
                sbSql.AppendFormat(@"  ORDER BY TA003,[TA021],TB003");
                sbSql.AppendFormat(@"  ");

                adapter13 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder13 = new SqlCommandBuilder(adapter13);
                sqlConn.Open();
                ds13.Clear();
                adapter13.Fill(ds13, "TEMPds13");
                sqlConn.Close();


                if (ds13.Tables["TEMPds13"].Rows.Count == 0)
                {
                    dataGridView9.DataSource = null;
                }
                else
                {
                    if (ds13.Tables["TEMPds13"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView9.DataSource = ds13.Tables["TEMPds13"];
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

        private void dataGridView9_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView9.CurrentRow != null)
            {
                int rowindex = dataGridView9.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView9.Rows[rowindex];
                    textBox26.Text = row.Cells["日期"].Value.ToString();
                    textBox27.Text = row.Cells["線別號"].Value.ToString();
                    textBox28.Text = row.Cells["線別"].Value.ToString();
                    textBox29.Text = row.Cells["品號"].Value.ToString();
                    textBox30.Text = row.Cells["品名"].Value.ToString();
                    textBox31.Text = row.Cells["總數量"].Value.ToString();
                    textBox36.Text = row.Cells["入庫別"].Value.ToString();
                    dt5 = Convert.ToDateTime(row.Cells["日期"].Value.ToString().Substring(0,4)+"/"+row.Cells["日期"].Value.ToString().Substring(4, 2)+"/"+ row.Cells["日期"].Value.ToString().Substring(6, 2));

                    MB001E = row.Cells["品號"].Value.ToString();
                    MB002E = row.Cells["品名"].Value.ToString();                   

                    SEARCHMOCMANULINETOATL();
                }
                else
                {
                    textBox26.Text = null;
                    textBox27.Text = null;
                    textBox28.Text = null;
                    textBox29.Text = null;
                    textBox30.Text = null;
                    textBox31.Text = null;

                }
            }
        }

        public void SEARCHMOCMANULINETOATL()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT  [MOCTA001] AS '單別',[MOCTA002] AS '製令'");
                sbSql.AppendFormat(@"  ,[TA003] AS '日期',[TA021] AS '線別號',[TA021N] AS '線別',[TB003] AS '品號',[TB012] AS '品名',[TB004] AS '總數量',[TB009] AS '入庫別'");
                sbSql.AppendFormat(@"  ,[ID]");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINETOATL]");
                sbSql.AppendFormat(@"  WHERE [TA003]='{0}' AND [TA021]='{1}' AND [TB003]='{2}'   AND [TB004]='{3}' ", textBox26.Text, textBox27.Text, textBox29.Text, textBox31.Text);
                sbSql.AppendFormat(@"  ");

                adapter14 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder14 = new SqlCommandBuilder(adapter14);
                sqlConn.Open();
                ds14.Clear();
                adapter14.Fill(ds14, "TEMPds14");
                sqlConn.Close();


                if (ds14.Tables["TEMPds14"].Rows.Count == 0)
                {
                    dataGridView10.DataSource = null;
                }
                else
                {
                    if (ds14.Tables["TEMPds14"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView10.DataSource = ds14.Tables["TEMPds14"];
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

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            //CALPRODUCT();
        }

        public void CALPRODUCT()
        {
            try
            {
                if (MANU.Equals("新廠製二組"))
                {
                    textBox5.Text = (Convert.ToDecimal(textBox32.Text) * Convert.ToDecimal(textBox4.Text)).ToString();
                }

                else if (MANU.Equals("新廠包裝線"))
                {
                    
                    textBox12.Text = (Convert.ToDecimal(textBox33.Text) * Convert.ToDecimal(textBox8.Text) ).ToString();
                }
                else if (MANU.Equals("新廠製一組"))
                {
                    textBox19.Text = (Convert.ToDecimal(textBox34.Text) * Convert.ToDecimal(textBox15.Text)).ToString();
                }
                else if (MANU.Equals("新廠製三組(手工)"))
                {
                    textBox23.Text = (Convert.ToDecimal(textBox35.Text) * Convert.ToDecimal(textBox21.Text)).ToString();
                }
            }
            catch
            {
                //MessageBox.Show("請填數字");
            }
            finally
            {

            }
            
        }

        public void CALPRODUCTDETAIL()
        {
            try
            {
                if (MANU.Equals("新廠製二組"))
                {
                    textBox4.Text = Math.Round(Convert.ToDecimal(textBox5.Text)/ Convert.ToDecimal(textBox32.Text), 4).ToString();
                }

                else if (MANU.Equals("新廠包裝線"))
                {
                    SEARCHMB001BOX();
                    textBox8.Text = Math.Round(Convert.ToDecimal(textBox12.Text) / Convert.ToDecimal(textBox33.Text)/BOXNUMERB, 4).ToString();
                }
                else if (MANU.Equals("新廠製一組"))
                {
                    textBox15.Text = Math.Round(Convert.ToDecimal(textBox19.Text) / Convert.ToDecimal(textBox34.Text), 4).ToString();
                }
                else if (MANU.Equals("新廠製三組(手工)"))
                {
                    textBox21.Text = Math.Round(Convert.ToDecimal(textBox23.Text) / Convert.ToDecimal(textBox35.Text), 4).ToString();
                }
            }
            catch
            {
                //MessageBox.Show("請填數字");
            }
            finally
            {

            }

        }
        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            //CALPRODUCT();
        }

        private void textBox15_TextChanged(object sender, EventArgs e)
        {
            //CALPRODUCT();
        }

        private void textBox21_TextChanged(object sender, EventArgs e)
        {
            //CALPRODUCT();
        }

        public void ADDMOCMANULINETOATL()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();


                sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[MOCMANULINETOATL]");
                sbSql.AppendFormat(" ([ID],[TA003],[TA021],[TA021N],[TB003],[TB012],[TB004],[TB009],[MOCTA001],[MOCTA002])");
                sbSql.AppendFormat(" VALUES ({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}')", "NEWID()",textBox26.Text,textBox27.Text,textBox28.Text,textBox29.Text, textBox30.Text, textBox31.Text,textBox36.Text, TA001, TA002);
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

        public void DELMOCMANULINETOATL()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat("  DELETE [TKMOC].[dbo].[MOCMANULINETOATL]");
                sbSql.AppendFormat("  WHERE ID='{0}'", DELID5);          
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

        private void dataGridView10_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView10.CurrentRow != null)
            {
                int rowindex = dataGridView10.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView10.Rows[rowindex];
                    DELID5 = row.Cells["ID"].Value.ToString();
                    DELMOCTA001E = row.Cells["單別"].Value.ToString();
                    DELMOCTA002E = row.Cells["製令"].Value.ToString();



                }
                else
                {
                    DELID5 = null;

                }
            }
        }

        public void SETIN()
        {
            label51.Text = "20001";
            label52.Text = "20001";
            label53.Text = "20001";
            label54.Text = "20001";
            IN1 = "20001";
            IN2 = "20001";
            IN3 = "20001";
            IN4 = "20001";

        }
        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            label51.Text = comboBox5.SelectedValue.ToString();
            IN1= comboBox5.SelectedValue.ToString();
        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            label52.Text = comboBox6.SelectedValue.ToString();
            IN2 = comboBox6.SelectedValue.ToString();
        }

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {
            label53.Text = comboBox7.SelectedValue.ToString();
            IN3 = comboBox7.SelectedValue.ToString();

        }

        private void comboBox8_SelectedIndexChanged(object sender, EventArgs e)
        {
            label54.Text = comboBox8.SelectedValue.ToString();
            IN4 = comboBox8.SelectedValue.ToString();

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            CALPRODUCTDETAIL();
        }
        private void textBox12_TextChanged(object sender, EventArgs e)
        {
            CALPRODUCTDETAIL();
        }
        private void textBox19_TextChanged(object sender, EventArgs e)
        {
            CALPRODUCTDETAIL();
        }

        private void textBox23_TextChanged(object sender, EventArgs e)
        {
            CALPRODUCTDETAIL();
        }

        public void SEARCHBOMMD()
        {
            if (MANU.Equals("新廠製二組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();


                    sbSql.AppendFormat(@"  SELECT MD001,MD003,MB002,CONVERT(decimal(18,2), MD006/MD007) AS MD006");
                    sbSql.AppendFormat(@"  FROM [TK].dbo.BOMMD,[TK].dbo.INVMB");
                    sbSql.AppendFormat(@"  WHERE MD003=MB001");
                    sbSql.AppendFormat(@"  AND MB002 LIKE '%低筋%'");
                    sbSql.AppendFormat(@"  AND MD001='{0}'", textBox1.Text);
                    sbSql.AppendFormat(@"  ");


                    adapter15 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder15 = new SqlCommandBuilder(adapter15);
                    sqlConn.Open();
                    ds15.Clear();
                    adapter15.Fill(ds15, "TEMPds15");
                    sqlConn.Close();


                    if (ds15.Tables["TEMPds15"].Rows.Count == 0)
                    {
                        SETNULL5();
                    }
                    else
                    {
                        if (ds15.Tables["TEMPds15"].Rows.Count >= 1)
                        {
                            textBox37.Text = ds15.Tables["TEMPds15"].Rows[0]["MD006"].ToString();
                         ;
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
            else if (MANU.Equals("新廠包裝線"))
            {
                try
                {
                   

                }
                catch
                {

                }
                finally
                {

                }


            }

            else if (MANU.Equals("新廠製一組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"  SELECT MD001,MD003,MB002,CONVERT(decimal(18,2), MD006/MD007) AS MD006");
                    sbSql.AppendFormat(@"  FROM [TK].dbo.BOMMD,[TK].dbo.INVMB");
                    sbSql.AppendFormat(@"  WHERE MD003=MB001");
                    sbSql.AppendFormat(@"  AND MB002 LIKE '%低筋%'");
                    sbSql.AppendFormat(@"  AND MD001='{0}'", textBox14.Text);
                    sbSql.AppendFormat(@"  ");

                    adapter16 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder16 = new SqlCommandBuilder(adapter16);
                    sqlConn.Open();
                    ds16.Clear();
                    adapter16.Fill(ds16, "TEMPds16");
                    sqlConn.Close();


                    if (ds16.Tables["TEMPds16"].Rows.Count == 0)
                    {
                        SETNULL5(); 
                    }
                    else
                    {
                        if (ds16.Tables["TEMPds16"].Rows.Count >= 1)
                        {
                            textBox38.Text = ds16.Tables["TEMPds16"].Rows[0]["MD006"].ToString();
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
            else if (MANU.Equals("新廠製三組(手工)"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"  SELECT MD001,MD003,MB002,CONVERT(decimal(18,2), MD006/MD007) AS MD006");
                    sbSql.AppendFormat(@"  FROM [TK].dbo.BOMMD,[TK].dbo.INVMB");
                    sbSql.AppendFormat(@"  WHERE MD003=MB001");
                    sbSql.AppendFormat(@"  AND MB002 LIKE '%低筋%'");
                    sbSql.AppendFormat(@"  AND MD001='{0}'", textBox20.Text);
                    sbSql.AppendFormat(@"  ");

                    adapter17 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder17 = new SqlCommandBuilder(adapter17);
                    sqlConn.Open();
                    ds17.Clear();
                    adapter17.Fill(ds17, "TEMPds17");
                    sqlConn.Close();


                    if (ds17.Tables["TEMPds17"].Rows.Count == 0)
                    {
                        SETNULL5();
                    }
                    else
                    {
                        if (ds17.Tables["TEMPds17"].Rows.Count >= 1)
                        {
                            textBox39.Text = ds17.Tables["TEMPds17"].Rows[0]["MD006"].ToString();
                            
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
        }

        public void SETNULL5()
        {
            //textBox1.Text = null;

            textBox37.Text = null;
            textBox38.Text = null;
            textBox39.Text = null;
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            dateTimePicker2.Value = dateTimePicker1.Value;
        }

        private void dateTimePicker3_ValueChanged(object sender, EventArgs e)
        {
            dateTimePicker4.Value = dateTimePicker3.Value;
        }

        private void dateTimePicker6_ValueChanged(object sender, EventArgs e)
        {
            dateTimePicker7.Value = dateTimePicker6.Value;
        }

        private void dateTimePicker8_ValueChanged(object sender, EventArgs e)
        {
            dateTimePicker9.Value = dateTimePicker8.Value;
        }

        public void SEARCHMB017()
        {
            if (MANU.Equals("新廠製二組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"  SELECT MB001,MB002,MB003,MC004,MB017            ");
                    sbSql.AppendFormat(@"  FROM [TK].dbo.INVMB,[TK].dbo.BOMMC");
                    sbSql.AppendFormat(@"  WHERE MB001=MC001");
                    sbSql.AppendFormat(@"  AND MB001='{0}'", MB001);
                    sbSql.AppendFormat(@"  ");

                    adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                    sqlConn.Open();
                    ds2.Clear();
                    adapter2.Fill(ds2, "TEMPds2");
                    sqlConn.Close();


                    if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                    {
                        SETNULL1();
                    }
                    else
                    {
                        if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                        {
                            comboBox5.SelectedValue = ds2.Tables["TEMPds2"].Rows[0]["MB017"].ToString();
                            label51.Text = ds2.Tables["TEMPds2"].Rows[0]["MB017"].ToString();
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
            else if (MANU.Equals("新廠包裝線"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"  SELECT MB001,MB002,MB003,MC004 ,MB017 ");
                    sbSql.AppendFormat(@"  FROM [TK].dbo.INVMB,[TK].dbo.BOMMC");
                    sbSql.AppendFormat(@"  WHERE MB001=MC001");
                    sbSql.AppendFormat(@"  AND MB001='{0}'", MB001);
                    sbSql.AppendFormat(@"  ");

                    adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                    sqlConn.Open();
                    ds2.Clear();
                    adapter2.Fill(ds2, "TEMPds2");
                    sqlConn.Close();


                    if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                    {
                        SETNULL1();
                    }
                    else
                    {
                        if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                        {
                            comboBox6.SelectedValue = ds2.Tables["TEMPds2"].Rows[0]["MB017"].ToString();
                            label52.Text = ds2.Tables["TEMPds2"].Rows[0]["MB017"].ToString();

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

            else if (MANU.Equals("新廠製一組"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"  SELECT MB001,MB002,MB003,MC004 ,MB017 ");
                    sbSql.AppendFormat(@"  FROM [TK].dbo.INVMB,[TK].dbo.BOMMC");
                    sbSql.AppendFormat(@"  WHERE MB001=MC001");
                    sbSql.AppendFormat(@"  AND MB001='{0}'", MB001);
                    sbSql.AppendFormat(@"  ");

                    adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                    sqlConn.Open();
                    ds2.Clear();
                    adapter2.Fill(ds2, "TEMPds2");
                    sqlConn.Close();


                    if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                    {
                        SETNULL4();
                    }
                    else
                    {
                        if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                        {
                            comboBox7.SelectedValue = ds2.Tables["TEMPds2"].Rows[0]["MB017"].ToString();
                            label53.Text = ds2.Tables["TEMPds2"].Rows[0]["MB017"].ToString();
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
            else if (MANU.Equals("新廠製三組(手工)"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"  SELECT MB001,MB002,MB003,MC004 ,MB017 ");
                    sbSql.AppendFormat(@"  FROM [TK].dbo.INVMB,[TK].dbo.BOMMC");
                    sbSql.AppendFormat(@"  WHERE MB001=MC001");
                    sbSql.AppendFormat(@"  AND MB001='{0}'", MB001);
                    sbSql.AppendFormat(@"  ");

                    adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                    sqlConn.Open();
                    ds2.Clear();
                    adapter2.Fill(ds2, "TEMPds2");
                    sqlConn.Close();


                    if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                    {
                        SETNULL4();
                    }
                    else
                    {
                        if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                        {
                            comboBox8.SelectedValue = ds2.Tables["TEMPds2"].Rows[0]["MB017"].ToString();
                            label54.Text = ds2.Tables["TEMPds2"].Rows[0]["MB017"].ToString();
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
        }

        public void UPDATEMOCMANULINE()
        {
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage1"])
            {
                frmMOCMANULINESub MOCMANULINESub = new frmMOCMANULINESub(ID1);
                MOCMANULINESub.ShowDialog();
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage4"])
            {
                frmMOCMANULINESub MOCMANULINESub = new frmMOCMANULINESub(ID2);
                MOCMANULINESub.ShowDialog();
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage2"])
            {
                frmMOCMANULINESub MOCMANULINESub = new frmMOCMANULINESub(ID3);
                MOCMANULINESub.ShowDialog();
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage3"])
            {
                frmMOCMANULINESub MOCMANULINESub = new frmMOCMANULINESub(ID4);
                MOCMANULINESub.ShowDialog();
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage5"])
            {
                
            }

            

        }

        public void CHECKMOCTAB()
        {
            string CHECKID = null;

            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage1"])
            {
                CHECKID = ID1;
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage4"])
            {
                CHECKID = ID2;
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage2"])
            {
                CHECKID = ID3;
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage3"])
            {
                CHECKID = ID4;
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage5"])
            {

            }

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT	MOCTA001,MOCTA002");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINERESULT]");
                sbSql.AppendFormat(@"  WHERE [SID]='{0}'",CHECKID);
                sbSql.AppendFormat(@"  UNION ALL");
                sbSql.AppendFormat(@"  SELECT	TA001,TA002");
                sbSql.AppendFormat(@"  FROM [TK].[dbo].[MOCTA]");
                sbSql.AppendFormat(@"  WHERE EXISTS (SELECT [MOCTA001],[MOCTA002] FROM [TKMOC].[dbo].[MOCMANULINERESULT] WHERE [SID]='{0}' AND TA001=MOCTA001 AND TA002=MOCTA002)", CHECKID);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter19 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder19 = new SqlCommandBuilder(adapter19);
                sqlConn.Open();
                ds19.Clear();
                adapter19.Fill(ds19, "TEMPds19");
                sqlConn.Close();


                if (ds19.Tables["TEMPds19"].Rows.Count == 0)
                {
                    UPDATEMOCMANULINE();
                }
                else
                {
                    MessageBox.Show("有製令未刪除，請檢查一下");
                }

            }
            catch
            {

            }
            finally
            {

            }
        }

        public void SEARCHMB001BOX()
        {
            
            if (MANU.Equals("新廠包裝線"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"  SELECT TOP 1 MD001,MD003,MB001,MB002,ISNULL(MD007,1) AS MD007,ISNULL(MD010,1) AS MD010");
                    sbSql.AppendFormat(@"  FROM [TK].dbo.BOMMD,[TK].dbo.INVMB");
                    sbSql.AppendFormat(@"  WHERE MD003=MB001");
                    sbSql.AppendFormat(@"  AND MB002 LIKE '%箱%'");
                    sbSql.AppendFormat(@"  AND MD003 LIKE '2%'");
                    sbSql.AppendFormat(@"  AND MD001='{0}'", textBox7.Text);
                    sbSql.AppendFormat(@"  ");

                    adapter20 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder20 = new SqlCommandBuilder(adapter20);
                    sqlConn.Open();
                    ds20.Clear();
                    adapter20.Fill(ds20, "TEMPds20");
                    sqlConn.Close();


                    if (ds20.Tables["TEMPds20"].Rows.Count == 0)
                    {
                        BOXNUMERB = 1;
                    }
                    else
                    {
                        if (ds20.Tables["TEMPds20"].Rows.Count >= 1)
                        {
                            BOXNUMERB = (Convert.ToInt32(ds20.Tables["TEMPds20"].Rows[0]["MD007"].ToString())/ Convert.ToInt32(ds20.Tables["TEMPds20"].Rows[0]["MD010"].ToString()));
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
            
        }
        public void SEARCHMOCMANULINECOP()
        {
            string SOURCEID = null;

            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage1"])
            {
                SOURCEID = ID1;
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage4"])
            {
                SOURCEID = ID2;
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage2"])
            {
                SOURCEID = ID3;
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage3"])
            {
                SOURCEID = ID4;
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage5"])
            {

            }


            try
            {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);

            sbSql.Clear();
            sbSqlQuery.Clear();

            sbSql.AppendFormat(@"  SELECT [MANU] AS '組別',[TC001] AS '訂單單別',[TC002] AS '訂單單號',[SID] AS '來源',[ID]");
            sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINECOP]");
            sbSql.AppendFormat(@"  WHERE [SID]='{0}'", SOURCEID);
            sbSql.AppendFormat(@"  ORDER BY [MANU],[TC001],[TC002]");
            sbSql.AppendFormat(@"  ");
            sbSql.AppendFormat(@"  ");


            adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

            sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
            sqlConn.Open();
            ds21.Clear();
            adapter2.Fill(ds21, "TEMPds21");
            sqlConn.Close();


                if (ds21.Tables["TEMPds21"].Rows.Count == 0)
                {

                }
                else
                {
                    if (ds21.Tables["TEMPds21"].Rows.Count >= 1)
                    {
                        if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage1"])
                        {
                            dataGridView11.DataSource = ds21.Tables["TEMPds21"];
                            dataGridView11.AutoResizeColumns();
                        }
                        else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage4"])
                        {
                            dataGridView12.DataSource = ds21.Tables["TEMPds21"];
                            dataGridView12.AutoResizeColumns();
                        }
                        else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage2"])
                        {
                            dataGridView13.DataSource = ds21.Tables["TEMPds21"];
                            dataGridView13.AutoResizeColumns();
                        }
                        else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage3"])
                        {
                            dataGridView14.DataSource = ds21.Tables["TEMPds21"];
                            dataGridView14.AutoResizeColumns();
                        }
                        else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage5"])
                        {

                        }
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


        private void dataGridView11_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView11.CurrentRow != null)
            {
                int rowindex = dataGridView11.CurrentRow.Index;
                
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView11.Rows[rowindex];
                    textBox40.Text = row.Cells["訂單單別"].Value.ToString();
                    textBox41.Text = row.Cells["訂單單號"].Value.ToString();                    
                }
                else
                {
                    textBox40.Text = null;
                    textBox41.Text = null;                   

                }
            }
            else
            {
                textBox40.Text = null;
                textBox41.Text = null;

            }
        }

        private void dataGridView12_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView12.CurrentRow != null)
            {
                int rowindex = dataGridView12.CurrentRow.Index;

                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView12.Rows[rowindex];
                    textBox42.Text = row.Cells["訂單單別"].Value.ToString();
                    textBox43.Text = row.Cells["訂單單號"].Value.ToString();
                }
                else
                {
                    textBox42.Text = null;
                    textBox43.Text = null;

                }
            }
            else
            {
                textBox42.Text = null;
                textBox43.Text = null;

            }
        }

        private void dataGridView13_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView13.CurrentRow != null)
            {
                int rowindex = dataGridView13.CurrentRow.Index;

                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView13.Rows[rowindex];
                    textBox44.Text = row.Cells["訂單單別"].Value.ToString();
                    textBox45.Text = row.Cells["訂單單號"].Value.ToString();
                }
                else
                {
                    textBox44.Text = null;
                    textBox45.Text = null;

                }
            }
            else
            {
                textBox44.Text = null;
                textBox45.Text = null;

            }
        }

        private void dataGridView14_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView14.CurrentRow != null)
            {
                int rowindex = dataGridView14.CurrentRow.Index;

                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView14.Rows[rowindex];
                    textBox46.Text = row.Cells["訂單單別"].Value.ToString();
                    textBox47.Text = row.Cells["訂單單號"].Value.ToString();
                }
                else
                {
                    textBox46.Text = null;
                    textBox47.Text = null;

                }
            }
            else
            {
                textBox46.Text = null;
                textBox47.Text = null;

            }
        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCHMOCMANULINE();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox2.Text))
            {
                ADDMOCMANULINE();
                SETNULL2();
            }
            else
            {
                MessageBox.Show("品名錯誤");
            }
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELMOCMANULINE();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }
        private void button4_Click(object sender, EventArgs e)
        {
            
            frmSUBMOCMANULINE SUBfrmSUBMOCMANULINE = new frmSUBMOCMANULINE();
            SUBfrmSUBMOCMANULINE.ShowDialog();
            textBox1.Text = SUBfrmSUBMOCMANULINE.TextBoxMsg;
  
        }


        private void button5_Click(object sender, EventArgs e)
        {

            TA002 = GETMAXTA002(TA001);
            ADDMOCMANULINERESULT();
            ADDMOCTATB();
            SEARCHMOCMANULINERESULT();

            MessageBox.Show("完成");
        }
        private void button6_Click(object sender, EventArgs e)
        {
            SEARCHMOCMANULINE();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            frmSUBMOCMANULINE SUBfrmSUBMOCMANULINE = new frmSUBMOCMANULINE();
            SUBfrmSUBMOCMANULINE.ShowDialog();
            textBox7.Text = SUBfrmSUBMOCMANULINE.TextBoxMsg;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox7.Text))
            {
                ADDMOCMANULINE();
                SETNULL3();
            }
            else
            {
                MessageBox.Show("品名錯誤");
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELMOCMANULINE();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }
        private void button10_Click(object sender, EventArgs e)
        {
            TA002 = GETMAXTA002(TA001);
            ADDMOCMANULINERESULT();
            ADDMOCTATB();
            SEARCHMOCMANULINERESULT();

            MessageBox.Show("完成");
        }
        private void button11_Click(object sender, EventArgs e)
        {
            SEARCHMOCMANULINE();
        }
        private void button12_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox14.Text))
            {
                ADDMOCMANULINE();
                SETNULL4();
            }
            else
            {
                MessageBox.Show("品名錯誤");
            }
        }
        private void button13_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELMOCMANULINE();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            frmSUBMOCMANULINE SUBfrmSUBMOCMANULINE = new frmSUBMOCMANULINE();
            SUBfrmSUBMOCMANULINE.ShowDialog();
            textBox14.Text = SUBfrmSUBMOCMANULINE.TextBoxMsg;
        }

        private void button15_Click(object sender, EventArgs e)
        {
            TA002 = GETMAXTA002(TA001);
            ADDMOCMANULINERESULT();
            ADDMOCTATB();
            SEARCHMOCMANULINERESULT();

            MessageBox.Show("完成");
        }
        private void button16_Click(object sender, EventArgs e)
        {
            SEARCHMOCMANULINE();
        }
        private void button17_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox20.Text))
            {
                ADDMOCMANULINE();
                SETNULL4();
            }
            else
            {
                MessageBox.Show("品名錯誤");
            }
        }
        private void button19_Click(object sender, EventArgs e)
        {
            frmSUBMOCMANULINE SUBfrmSUBMOCMANULINE = new frmSUBMOCMANULINE();
            SUBfrmSUBMOCMANULINE.ShowDialog();
            textBox20.Text = SUBfrmSUBMOCMANULINE.TextBoxMsg;
        }

        private void button18_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELMOCMANULINE();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }

        private void button20_Click(object sender, EventArgs e)
        {

            TA002 = GETMAXTA002(TA001);
            ADDMOCMANULINERESULT();
            ADDMOCTATB();
            SEARCHMOCMANULINERESULT();

            MessageBox.Show("完成");
        }


        private void button21_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELMOCMANULINERESULT();
                SEARCHMOCMANULINE();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }

        private void button24_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELMOCMANULINERESULT();
                SEARCHMOCMANULINE();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }

        private void button22_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELMOCMANULINERESULT();
                SEARCHMOCMANULINE();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }

        private void button23_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELMOCMANULINERESULT();
                SEARCHMOCMANULINE();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }


        private void button25_Click(object sender, EventArgs e)
        {
            SEARCHMOCTB();
        }



        private void button26_Click(object sender, EventArgs e)
        {
            TA002 = GETMAXTA002(TA001);
            ADDMOCMANULINETOATL();
            ADDMOCTATB();
            SEARCHMOCMANULINETOATL();

            MessageBox.Show("完成");
        }

        private void button27_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELMOCMANULINETOATL();
                SEARCHMOCMANULINETOATL();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }

        private void button28_Click(object sender, EventArgs e)
        {
            frmSUBMOCCOPMA SUBfrmSUBMOCCOPMA = new frmSUBMOCCOPMA();
            SUBfrmSUBMOCCOPMA.ShowDialog();
            textBox6.Text = SUBfrmSUBMOCCOPMA.TextBoxMsg;
        }
        

        private void button29_Click(object sender, EventArgs e)
        {
            frmSUBMOCCOPMA SUBfrmSUBMOCCOPMA = new frmSUBMOCCOPMA();
            SUBfrmSUBMOCCOPMA.ShowDialog();
            textBox9.Text = SUBfrmSUBMOCCOPMA.TextBoxMsg;
        }

        private void button30_Click(object sender, EventArgs e)
        {
            frmSUBMOCCOPMA SUBfrmSUBMOCCOPMA = new frmSUBMOCCOPMA();
            SUBfrmSUBMOCCOPMA.ShowDialog();
            textBox16.Text = SUBfrmSUBMOCCOPMA.TextBoxMsg;
        }

        private void button31_Click(object sender, EventArgs e)
        {
            frmSUBMOCCOPMA SUBfrmSUBMOCCOPMA = new frmSUBMOCCOPMA();
            SUBfrmSUBMOCCOPMA.ShowDialog();
            textBox22.Text = SUBfrmSUBMOCCOPMA.TextBoxMsg;
        }

        private void button32_Click(object sender, EventArgs e)
        {
            CHECKMOCTAB();
            SEARCHMOCMANULINE();
        }
        private void button35_Click(object sender, EventArgs e)
        {
            CHECKMOCTAB();
            SEARCHMOCMANULINE();
        }

        private void button34_Click(object sender, EventArgs e)
        {
            CHECKMOCTAB();
            SEARCHMOCMANULINE();
        }

        private void button33_Click(object sender, EventArgs e)
        {
            CHECKMOCTAB();
            SEARCHMOCMANULINE();
        }

        private void button36_Click(object sender, EventArgs e)
        {
            frmMOCMANULINECOP SUBfrmMOCMANULINECOP = new frmMOCMANULINECOP(SUBID,SUBBAR,SUBNUM,SUBBOX,SUBPACKAGE);
            if (!string.IsNullOrEmpty(SUBID))
            {
                SUBfrmMOCMANULINECOP.ShowDialog();
            }

            SEARCHMOCMANULINECOP();
        }
        private void button37_Click(object sender, EventArgs e)
        {

            frmMOCMANULINECOP SUBfrmMOCMANULINECOP = new frmMOCMANULINECOP(SUBID2, SUBBAR2, SUBNUM2, SUBBOX2, SUBPACKAGE2);
            if (!string.IsNullOrEmpty(SUBID2))
            {
                SUBfrmMOCMANULINECOP.ShowDialog();
            }

            SEARCHMOCMANULINECOP();
        }

        private void button38_Click(object sender, EventArgs e)
        {

            frmMOCMANULINECOP SUBfrmMOCMANULINECOP = new frmMOCMANULINECOP(SUBID3, SUBBAR3, SUBNUM3, SUBBOX3, SUBPACKAGE3);
            if (!string.IsNullOrEmpty(SUBID3))
            {
                SUBfrmMOCMANULINECOP.ShowDialog();
            }

            SEARCHMOCMANULINECOP();
        }

        private void button39_Click(object sender, EventArgs e)
        {

            frmMOCMANULINECOP SUBfrmMOCMANULINECOP = new frmMOCMANULINECOP(SUBID4, SUBBAR4, SUBNUM4, SUBBOX4, SUBPACKAGE4);
            if (!string.IsNullOrEmpty(SUBID4))
            {
                SUBfrmMOCMANULINECOP.ShowDialog();
            }

            SEARCHMOCMANULINECOP();
        }


        #endregion

       
    }
}
