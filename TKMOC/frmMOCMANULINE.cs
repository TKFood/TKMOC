﻿using System;
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

        DataSet dsBOMMC = new DataSet();
        DataSet dsBOMMD = new DataSet();

        DataTable dt = new DataTable();
        string tablename = null;
        int result;
        string MANU= "新廠製二組";

        string ID1;
        DateTime dt1;
        string DELID1;
        string ID2;
        DateTime dt2;
        string DELID2;
        string ID3;
        DateTime dt3;
        string DELID3;
        string ID4;
        DateTime dt4;
        string DELID4;

        string TA001 = "A510";
        string TA002;
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
        decimal BAR;
        decimal BOX;
        decimal BAR2;
        decimal BAR3;

        string BOMVARSION;
        string UNIT;
        decimal BOMBAR;

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
            Sequel.AppendFormat(@"SELECT MD001,MD002 FROM [TK].dbo.CMSMD    WHERE MD002 LIKE '新廠製四組(包裝)%'   ");
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
            Sequel.AppendFormat(@"SELECT MC001 FROM [TK].dbo.CMSMC WHERE MC001 LIKE '2000%'   ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MC001", typeof(string));
         
            da.Fill(dt);
            comboBox5.DataSource = dt.DefaultView;
            comboBox5.ValueMember = "MC001";
            comboBox5.DisplayMember = "MC001";
            sqlConn.Close();


        }
        public void comboBox6load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MC001 FROM [TK].dbo.CMSMC WHERE MC001 LIKE '2000%'   ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MC001", typeof(string));
           
            da.Fill(dt);
            comboBox6.DataSource = dt.DefaultView;
            comboBox6.ValueMember = "MC001";
            comboBox6.DisplayMember = "MC001";
            sqlConn.Close();


        }
        public void comboBox7load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MC001 FROM [TK].dbo.CMSMC WHERE MC001 LIKE '2000%'   ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MC001", typeof(string));
            
            da.Fill(dt);
            comboBox7.DataSource = dt.DefaultView;
            comboBox7.ValueMember = "MC001";
            comboBox7.DisplayMember = "MC001";
            sqlConn.Close();


        }
        public void comboBox8load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MC001 FROM [TK].dbo.CMSMC WHERE MC001 LIKE '2000%'   ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MC001", typeof(string));
           
            da.Fill(dt);
            comboBox8.DataSource = dt.DefaultView;
            comboBox8.ValueMember = "MC001";
            comboBox8.DisplayMember = "MC001";
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
                    sbSql.AppendFormat(@"  AND CONVERT(varchar(100),[MANUDATE],112) LIKE '{0}%'", dateTimePicker1.Value.ToString("yyyyMM"));
                    sbSql.AppendFormat(@"  ORDER BY [MANUDATE],[ID]");
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

            else if (MANU.Equals("新廠製四組(包裝)"))
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
                    sbSql.AppendFormat(@"  AND CONVERT(varchar(100),[MANUDATE],112) LIKE '{0}%'", dateTimePicker3.Value.ToString("yyyyMM"));
                    sbSql.AppendFormat(@"  ORDER BY [MANUDATE],[ID]");
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
                    sbSql.AppendFormat(@"  AND CONVERT(varchar(100),[MANUDATE],112) LIKE '{0}%'", dateTimePicker6.Value.ToString("yyyyMM"));
                    sbSql.AppendFormat(@"  ORDER BY [MANUDATE],[ID]");
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
                    sbSql.AppendFormat(@"  AND CONVERT(varchar(100),[MANUDATE],112) LIKE '{0}%'", dateTimePicker8.Value.ToString("yyyyMM"));
                    sbSql.AppendFormat(@"  ORDER BY [MANUDATE],[ID]");
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

                    sbSql.AppendFormat(@"  SELECT MB001,MB002,MB003");
                    sbSql.AppendFormat(@"  FROM [TK].dbo.INVMB");
                    sbSql.AppendFormat(@"  WHERE MB001='{0}'", textBox1.Text);

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
            else if (MANU.Equals("新廠製四組(包裝)"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"  SELECT MB001,MB002,MB003");
                    sbSql.AppendFormat(@"  FROM [TK].dbo.INVMB");
                    sbSql.AppendFormat(@"  WHERE MB001='{0}'", textBox7.Text);

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

                    sbSql.AppendFormat(@"  SELECT MB001,MB002,MB003");
                    sbSql.AppendFormat(@"  FROM [TK].dbo.INVMB");
                    sbSql.AppendFormat(@"  WHERE MB001='{0}'", textBox14.Text);

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

                    sbSql.AppendFormat(@"  SELECT MB001,MB002,MB003");
                    sbSql.AppendFormat(@"  FROM [TK].dbo.INVMB");
                    sbSql.AppendFormat(@"  WHERE MB001='{0}'", textBox20.Text);

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
            else if (MANU.Equals("新廠製四組(包裝)"))
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
            textBox13.Text = null;
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
                    SEARCHMOCMANULINERESULT();
;
                }
                else
                {
                    ID1 = null;

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

            else if (MANU.Equals("新廠製四組(包裝)"))
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
            else if (MANU.Equals("新廠製四組(包裝)"))
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

            if (MANU.Equals("新廠製二組"))
            {
                MOCMB001 = MB001;
                MOCTA004 = BAR;
            }
            else if (MANU.Equals("新廠製四組(包裝)"))
            {
                MOCMB001 = MB001B;
                MOCTA004 = BOX;
            }
            else if (MANU.Equals("新廠製一組"))
            {
                MOCMB001 = MB001C;
                MOCTA004 = BAR2;
            }
            else if (MANU.Equals("新廠製三組(手工)"))
            {
                MOCMB001 = MB001D;
                MOCTA004 = BAR3;
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
                sbSql.AppendFormat(" ,[TA019],[TA020],[TA021],[TA022],[TA024],[TA025],[TA030],[TA031],[TA034],[TA035]");
                sbSql.AppendFormat(" ,[TA040],[TA041],[TA043],[TA044],[TA045],[TA046],[TA047],[TA049],[TA050],[TA200])");
                sbSql.AppendFormat(" VALUES");
                sbSql.AppendFormat(" ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}',",MOCTA.COMPANY, MOCTA.CREATOR, MOCTA.USR_GROUP, MOCTA.CREATE_DATE, MOCTA.MODIFIER, MOCTA.MODI_DATE, MOCTA.FLAG, MOCTA.CREATE_TIME, MOCTA.MODI_TIME, MOCTA.TRANS_TYPE);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}',", MOCTA.TRANS_NAME, MOCTA.sync_count, MOCTA.DataGroup, MOCTA.TA001, MOCTA.TA002, MOCTA.TA003, MOCTA.TA004, MOCTA.TA005, MOCTA.TA006, MOCTA.TA007);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}',", MOCTA.TA009, MOCTA.TA010, MOCTA.TA011, MOCTA.TA012, MOCTA.TA013, MOCTA.TA014, MOCTA.TA015, MOCTA.TA016, MOCTA.TA017, MOCTA.TA018);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}',", MOCTA.TA019, MOCTA.TA020, MOCTA.TA021, MOCTA.TA022, MOCTA.TA024, MOCTA.TA025, MOCTA.TA030, MOCTA.TA031, MOCTA.TA034, MOCTA.TA035);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}')", MOCTA.TA040, MOCTA.TA041, MOCTA.TA043, MOCTA.TA044, MOCTA.TA045, MOCTA.TA046, MOCTA.TA047, MOCTA.TA049, MOCTA.TA050, MOCTA.TA200);
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" INSERT INTO [TK].dbo.[MOCTB]");
                sbSql.AppendFormat(" ([COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE]");
                sbSql.AppendFormat(" ,[TRANS_NAME],[sync_count],[DataGroup],[TB001],[TB002],[TB003],[TB004],[TB005],[TB006],[TB007]");
                sbSql.AppendFormat(" ,[TB009],[TB011],[TB012],[TB013],[TB014],[TB018],[TB019],[TB020],[TB022],[TB024]");
                sbSql.AppendFormat(" ,[TB025],[TB026],[TB027],[TB029],[TB030],[TB031],[TB501],[TB554],[TB556],[TB560])");
                sbSql.AppendFormat(" (SELECT ");
                sbSql.AppendFormat(" '{0}' [COMPANY],'{1}' [CREATOR],'{2}' [USR_GROUP],'{3}' [CREATE_DATE],'{4}' [MODIFIER],'{5}' [MODI_DATE],'{6}' [FLAG],'{7}' [CREATE_TIME],'{8}' [MODI_TIME],'{9}' [TRANS_TYPE]", MOCTA.COMPANY, MOCTA.CREATOR, MOCTA.USR_GROUP, MOCTA.CREATE_DATE, MOCTA.MODIFIER, MOCTA.MODI_DATE, MOCTA.FLAG, MOCTA.CREATE_TIME, MOCTA.MODI_TIME, MOCTA.TRANS_TYPE);
                sbSql.AppendFormat(" ,'{0}' [TRANS_NAME],{1} [sync_count],'{2}' [DataGroup],'{3}' [TB001],'{4}' [TB002],[BOMMD].MD003 [TB003],{5}*[BOMMD].MD006 [TB004],{5}*[BOMMD].MD006 [TB005],'****' [TB006],[BOMMD].MD004 [TB007]", MOCTA.TRANS_NAME, MOCTA.sync_count, MOCTA.DataGroup, MOCTA.TA001, MOCTA.TA002, MOCTA004);
                sbSql.AppendFormat(" ,'20006' [TB009],'1' [TB011],[INVMB].MB002 [TB012],[INVMB].MB003 [TB013],[BOMMD].MD001 [TB014],'N' [TB018],'0' [TB019],'0' [TB020],'2' [TB022],'0' [TB024]");
                sbSql.AppendFormat(" ,'****' [TB025],'0' [TB026],'1' [TB027],'0' [TB029],'0' [TB030],'0' [TB031],'0' [TB501],'N' [TB554],'0' [TB556],'0' [TB560]");
                sbSql.AppendFormat(" FROM [TK].dbo.[BOMMD],[TK].dbo.[INVMB]");
                sbSql.AppendFormat(" WHERE [BOMMD].MD003=[INVMB].MB001");
                sbSql.AppendFormat(" AND MD001='{0}')", MOCMB001);
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
                MOCTA.TA014 = dt1.ToString("yyyyMMdd");
                MOCTA.TA015 = (BAR * BOMBAR).ToString();
                MOCTA.TA016 = "0";
                MOCTA.TA017 = "0";
                MOCTA.TA018 = "0";
                MOCTA.TA019 = "20";
                MOCTA.TA020 = "20001";
                MOCTA.TA021 = "03";
                MOCTA.TA022 = "0";
                MOCTA.TA024 = "A510";
                MOCTA.TA025 = TA002;
                MOCTA.TA030 = "1";
                MOCTA.TA031 = "0";
                MOCTA.TA034 = MB002;
                MOCTA.TA035 = MB003;
                MOCTA.TA040 = dt1.ToString("yyyyMMdd");
                MOCTA.TA041 = "";
                MOCTA.TA043 = "1";
                MOCTA.TA044 = "N";
                MOCTA.TA045 = (BAR * BOMBAR).ToString();
                MOCTA.TA046 = (BAR * BOMBAR).ToString();
                MOCTA.TA047 = "0";
                MOCTA.TA049 = "0";
                MOCTA.TA050 = "0";
                MOCTA.TA200 = "1";


                return MOCTA;
            }

            else if (MANU.Equals("新廠製四組(包裝)"))
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
                MOCTA.TA014 = dt2.ToString("yyyyMMdd");
                MOCTA.TA015 = (BOX * BOMBAR).ToString();
                MOCTA.TA016 = "0";
                MOCTA.TA017 = "0";
                MOCTA.TA018 = "0";
                MOCTA.TA019 = "20";
                MOCTA.TA020 = "20001";
                MOCTA.TA021 = "07";
                MOCTA.TA022 = "0";
                MOCTA.TA024 = "A510";
                MOCTA.TA025 = TA002;
                MOCTA.TA030 = "1";
                MOCTA.TA031 = "0";
                MOCTA.TA034 = MB002B;
                MOCTA.TA035 = MB003B;
                MOCTA.TA040 = dt2.ToString("yyyyMMdd");
                MOCTA.TA041 = "";
                MOCTA.TA043 = "1";
                MOCTA.TA044 = "N";
                MOCTA.TA045 = (BOX * BOMBAR).ToString();
                MOCTA.TA046 = (BOX * BOMBAR).ToString();
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
                MOCTA.TA014 = dt3.ToString("yyyyMMdd");
                MOCTA.TA015 = (BAR2 * BOMBAR).ToString();
                MOCTA.TA016 = "0";
                MOCTA.TA017 = "0";
                MOCTA.TA018 = "0";
                MOCTA.TA019 = "20";
                MOCTA.TA020 = "20001";
                MOCTA.TA021 = "03";
                MOCTA.TA022 = "0";
                MOCTA.TA024 = "A510";
                MOCTA.TA025 = TA002;
                MOCTA.TA030 = "1";
                MOCTA.TA031 = "0";
                MOCTA.TA034 = MB002C;
                MOCTA.TA035 = MB003C;
                MOCTA.TA040 = dt3.ToString("yyyyMMdd");
                MOCTA.TA041 = "";
                MOCTA.TA043 = "1";
                MOCTA.TA044 = "N";
                MOCTA.TA045 = (BAR2 * BOMBAR).ToString();
                MOCTA.TA046 = (BAR2 * BOMBAR).ToString();
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
                MOCTA.TA014 = dt4.ToString("yyyyMMdd");
                MOCTA.TA015 = (BAR3 * BOMBAR).ToString();
                MOCTA.TA016 = "0";
                MOCTA.TA017 = "0";
                MOCTA.TA018 = "0";
                MOCTA.TA019 = "20";
                MOCTA.TA020 = "20001";
                MOCTA.TA021 = "04";
                MOCTA.TA022 = "0";
                MOCTA.TA024 = "A510";
                MOCTA.TA025 = TA002;
                MOCTA.TA030 = "1";
                MOCTA.TA031 = "0";
                MOCTA.TA034 = MB002D;
                MOCTA.TA035 = MB003D;
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
                    sbSql.AppendFormat(@"  FROM [TK].[dbo].[BOMMC]");
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
                            UNIT = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MC002"].ToString();
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
            else if (MANU.Equals("新廠製四組(包裝)"))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"  SELECT ");
                    sbSql.AppendFormat(@"  [MC001],[MC002],[MC003],[MC004],[MC005],[MC006],[MC007],[MC008],[MC009],[MC010],[MC011],[MC012],[MC013],[MC014],[MC015],[MC016],[MC017],[MC018],[MC019],[MC020],[MC021],[MC022],[MC023],[MC024],[MC025],[MC026],[MC027]");
                    sbSql.AppendFormat(@"  FROM [TK].[dbo].[BOMMC]");
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
                            UNIT = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MC002"].ToString();
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
                    sbSql.AppendFormat(@"  FROM [TK].[dbo].[BOMMC]");
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
                            UNIT = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MC002"].ToString();
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
                    sbSql.AppendFormat(@"  FROM [TK].[dbo].[BOMMC]");
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
                            UNIT = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MC002"].ToString();
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
            else if  (MANU.Equals("新廠製四組(包裝)"))
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
            else if (MANU.Equals("新廠製四組(包裝)"))
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

            else if (MANU.Equals("新廠製四組(包裝)"))
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
                MessageBox.Show("新廠製四組(包裝)");
                MANU = "新廠製四組(包裝)";
            }
        }
        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            SEARCHMB001();
        }
        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
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
                    BOX = Convert.ToDecimal(row.Cells["箱數"].Value.ToString());
                    SEARCHMOCMANULINERESULT();
                    ;
                }
                else
                {
                    ID2 = null;

                }
            }
        }

        private void textBox14_TextChanged(object sender, EventArgs e)
        {
            SEARCHMB001();
        }
        private void dataGridView5_SelectionChanged(object sender, EventArgs e)
        {
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
                    SEARCHMOCMANULINERESULT();
                    ;
                }
                else
                {
                    ID3 = null;

                }
            }
        }
        private void textBox20_TextChanged(object sender, EventArgs e)
        {
            SEARCHMB001();
        }
        private void dataGridView7_SelectionChanged(object sender, EventArgs e)
        {
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
                    SEARCHMOCMANULINERESULT();
                    ;
                }
                else
                {
                    ID4 = null;

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
                    sbSql.AppendFormat("  WHERE SID='{0}'", DELID1);
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

            else if (MANU.Equals("新廠製四組(包裝)"))
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

        #endregion




    }
}
