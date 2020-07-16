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
using Calendar.NET;
using FastReport;
using FastReport.Data;

namespace TKMOC
{
    public partial class frmMOCTADAILY : Form
    {
        private ComponentResourceManager _ResourceManager = new ComponentResourceManager();
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();

        int result;
        string STATUS = null;
        string ID = null;
        string TA001 = null;
        string TA002 = null;
        string TA021 = null;

        public frmMOCTADAILY()
        {
            InitializeComponent();

            comboBox1load();
        }

        public class DATAMOCTADAILY
        {
            public string TEMPERAVG;
            public string TEMPERMIN;
            public string TEMPERMAX;
            public string HUMIAVG;
            public string HUMIMIN;
            public string HUMIMAX;
            public string BSPEED;
            public string B1ARAVG;
            public string B1ARMIN;
            public string B1ARMAX;
            public string B1BRAVG;
            public string B1BRMIN;
            public string B1BRMAX;
            public string B1AMAVG;
            public string B1AMMIN;
            public string B1AMMAX;
            public string B1BMAVG;
            public string B1BMMIN;
            public string B1BMMAX;
            public string B1ALAVG;
            public string B1ALMIN;
            public string B1ALMAX;
            public string B1BLAVG;
            public string B1BLMIN;
            public string B1BLMAX;

            public string B2ARAVG;
            public string B2ARMIN;
            public string B2ARMAX;
            public string B2BRAVG;
            public string B2BRMIN;
            public string B2BRMAX;
            public string B2AMAVG;
            public string B2AMMIN;
            public string B2AMMAX;
            public string B2BMAVG;
            public string B2BMMIN;
            public string B2BMMAX;
            public string B2ALAVG;
            public string B2ALMIN;
            public string B2ALMAX;
            public string B2BLAVG;
            public string B2BLMIN;
            public string B2BLMAX;

            public string B3ARAVG;
            public string B3ARMIN;
            public string B3ARMAX;
            public string B3BRAVG;
            public string B3BRMIN;
            public string B3BRMAX;
            public string B3AMAVG;
            public string B3AMMIN;
            public string B3AMMAX;
            public string B3BMAVG;
            public string B3BMMIN;
            public string B3BMMAX;
            public string B3ALAVG;
            public string B3ALMIN;
            public string B3ALMAX;
            public string B3BLAVG;
            public string B3BLMIN;
            public string B3BLMAX;

            public string B4ARAVG;
            public string B4ARMIN;
            public string B4ARMAX;
            public string B4BRAVG;
            public string B4BRMIN;
            public string B4BRMAX;
            public string B4AMAVG;
            public string B4AMMIN;
            public string B4AMMAX;
            public string B4BMAVG;
            public string B4BMMIN;
            public string B4BMMAX;
            public string B4ALAVG;
            public string B4ALMIN;
            public string B4ALMAX;
            public string B4BLAVG;
            public string B4BLMIN;
            public string B4BLMAX;

            public string B5ARAVG;
            public string B5ARMIN;
            public string B5ARMAX;
            public string B5BRAVG;
            public string B5BRMIN;
            public string B5BRMAX;
            public string B5AMAVG;
            public string B5AMMIN;
            public string B5AMMAX;
            public string B5BMAVG;
            public string B5BMMIN;
            public string B5BMMAX;
            public string B5ALAVG;
            public string B5ALMIN;
            public string B5ALMAX;
            public string B5BLAVG;
            public string B5BLMIN;
            public string B5BLMAX;





        }
        #region FUNCTION

        public void comboBox1load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MD001,MD002 FROM [TK].dbo.CMSMD    WHERE MD002 LIKE '新廠%'   ");
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

        public void SEARCH(string IDDATE)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();

            DataSet ds1 = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  SELECT ");
                sbSql.AppendFormat(@"  [TA001] AS '製令',[TA002] AS '單號',[TA021] AS '線別',[MB001] AS '品號',[MB002] AS '品名',[MB003] AS '規格',[TA017] AS '生產量',[NUM] AS '入庫量',[NGNUM] AS '未熟量',CONVERT(VARCHAR, [SDATES], 120) AS '開始時間',CONVERT(VARCHAR,[EDATES], 120)  AS '結束時間'");
                sbSql.AppendFormat(@"  ,[TEMPERAVG] AS '溫度-平均',[TEMPERMIN] AS '溫度-最小',[TEMPERMAX] AS '溫度-最大',[HUMIAVG] AS '溼度-平均',[HUMIMIN] AS '溼度-最小',[HUMIMAX] AS '溼度-最大'");
                sbSql.AppendFormat(@"  ,[ASPEED] AS '大線爐速',[BSPEED] AS '小線爐速'");
                sbSql.AppendFormat(@"  ,[A1AAVG] AS '大線-1段-上爐-平均',[A1AMIN] AS '大線-1段-上爐-最小',[A1AMAX] AS '大線-1段-上爐-最大'");
                sbSql.AppendFormat(@"  ,[A1BAVG] AS '大線-1段-下爐-平均',[A1BMIN] AS '大線-1段-下爐-最小',[A1BMAX] AS '大線-1段-下爐-最大'");
                sbSql.AppendFormat(@"  ,[A2AAVG] AS '大線-2段-上爐-平均',[A2AMIN] AS '大線-2段-上爐-最小',[A2AMAX] AS '大線-2段-上爐-最大'");
                sbSql.AppendFormat(@"  ,[A2BAVG] AS '大線-2段-下爐-平均',[A2BMIN] AS '大線-2段-下爐-最小',[A2BMAX] AS '大線-2段-下爐-最大'");
                sbSql.AppendFormat(@"  ,[A3AAVG] AS '大線-3段-上爐-平均',[A3AMIN] AS '大線-3段-上爐-最小',[A3AMAX] AS '大線-3段-上爐-最大'");
                sbSql.AppendFormat(@"  ,[A3BAVG] AS '大線-3段-下爐-平均',[A3BMIN] AS '大線-3段-下爐-最小',[A3BMAX] AS '大線-3段-下爐-最大'");
                sbSql.AppendFormat(@"  ,[A4AAVG] AS '大線-4段-上爐-平均',[A4AMIN] AS '大線-4段-上爐-最小',[A4AMAX] AS '大線-4段-上爐-最大'");
                sbSql.AppendFormat(@"  ,[A4BAVG] AS '大線-4段-下爐-平均',[A4BMIN] AS '大線-4段-下爐-最小',[A4BMAX] AS '大線-4段-下爐-最大'");
                sbSql.AppendFormat(@"  ,[B1ARAVG] AS '小線-1段-上右爐-平均',[B1ARMIN] AS '小線-1段-上右爐-最小',[B1ARMAX] AS '小線-1段-上右爐-最大'");
                sbSql.AppendFormat(@"  ,[B1BRAVG] AS '小線-1段-下右爐-平均',[B1BRMIN] AS '小線-1段-下右爐-最小',[B1BRMAX] AS '小線-1段-下右爐-最大'");
                sbSql.AppendFormat(@"  ,[B1AMAVG] AS '小線-1段-上中爐-平均',[B1AMMIN] AS '小線-1段-上中爐-最小',[B1AMMAX] AS '小線-1段-上中爐-最大'");
                sbSql.AppendFormat(@"  ,[B1BMAVG] AS '小線-1段-下中爐-平均',[B1BMMIN] AS '小線-1段-下中爐-最小',[B1BMMAX] AS '小線-1段-下中爐-最大'");
                sbSql.AppendFormat(@"  ,[B1ALAVG] AS '小線-1段-上左爐-平均',[B1ALMIN] AS '小線-1段-上左爐-最小',[B1ALMAX] AS '小線-1段-上左爐-最大'");
                sbSql.AppendFormat(@"  ,[B1BLAVG] AS '小線-1段-下左爐-平均',[B1BLMIN] AS '小線-1段-下左爐-最小',[B1BLMAX] AS '小線-1段-下左爐-最大'");
                sbSql.AppendFormat(@"  ,[B2ARAVG] AS '小線-2段-上右爐-平均',[B2ARMIN] AS '小線-2段-上右爐-最小',[B2ARNAX] AS '小線-2段-上右爐-最大'");
                sbSql.AppendFormat(@"  ,[B2BRAVG] AS '小線-2段-下右爐-平均',[B2BRMIN] AS '小線-2段-下右爐-最小',[B2BRMAX] AS '小線-2段-下右爐-最大'");
                sbSql.AppendFormat(@"  ,[B2AMAVG] AS '小線-2段-上中爐-平均',[B2AMMIN] AS '小線-2段-上中爐-最小',[B2AMMAX] AS '小線-2段-上中爐-最大'");
                sbSql.AppendFormat(@"  ,[B2BMAVG] AS '小線-2段-下中爐-平均',[B2BMMIN] AS '小線-2段-下中爐-最小',[B2BMMAX] AS '小線-2段-下中爐-最大'");
                sbSql.AppendFormat(@"  ,[B2ALAVG] AS '小線-2段-上左爐-平均',[B2ALMIN] AS '小線-2段-上左爐-最小',[B2ALMAX] AS '小線-2段-上左爐-最大'");
                sbSql.AppendFormat(@"  ,[B2BLAVG] AS '小線-2段-下左爐-平均',[B2BLMIN] AS '小線-2段-下左爐-最小',[B2BLMAX] AS '小線-2段-下左爐-最大'");
                sbSql.AppendFormat(@"  ,[B3ARAVG] AS '小線-3段-上右爐-平均',[B3ARMIN] AS '小線-3段-上右爐-最小',[B3ARMAX] AS '小線-3段-上右爐-最大'");
                sbSql.AppendFormat(@"  ,[B3BRAVG] AS '小線-3段-下右爐-平均',[B3BRMIN] AS '小線-3段-下右爐-最小',[B3BRMAX] AS '小線-3段-下右爐-最大'");
                sbSql.AppendFormat(@"  ,[B3AMAVG] AS '小線-3段-上中爐-平均',[B3AMMIN] AS '小線-3段-上中爐-最小',[B3AMMAX] AS '小線-3段-上中爐-最大'");
                sbSql.AppendFormat(@"  ,[B3BMAVG] AS '小線-3段-下中爐-平均',[B3BMMIN] AS '小線-3段-下中爐-最小',[B3BMMAX] AS '小線-3段-下中爐-最大'");
                sbSql.AppendFormat(@"  ,[B3ALAVG] AS '小線-3段-上左爐-平均',[B3ALMIN] AS '小線-3段-上左爐-最小',[B3ALMAX] AS '小線-3段-上左爐-最大'");
                sbSql.AppendFormat(@"  ,[B3BLAVG] AS '小線-3段-下左爐-平均',[B3BLMIN] AS '小線-3段-下左爐-最小',[B3BLMAX] AS '小線-3段-下左爐-最大'");
                sbSql.AppendFormat(@"  ,[B4ARAVG] AS '小線-4段-上右爐-平均',[B4ARMIN] AS '小線-4段-上右爐-最小',[B4ARMAX] AS '小線-4段-上右爐-最大'");
                sbSql.AppendFormat(@"  ,[B4BRAVG] AS '小線-4段-下右爐-平均',[B4BRMIN] AS '小線-4段-下右爐-最小',[B4BRMAX] AS '小線-4段-下右爐-最大'");
                sbSql.AppendFormat(@"  ,[B4AMAVG] AS '小線-4段-上中爐-平均',[B4AMMIN] AS '小線-4段-上中爐-最小',[B4AMMAX] AS '小線-4段-上中爐-最大'");
                sbSql.AppendFormat(@"  ,[B4BMAVG] AS '小線-4段-下中爐-平均',[B4BMMIN] AS '小線-4段-下中爐-最小',[B4BMMAX] AS '小線-4段-下中爐-最大'");
                sbSql.AppendFormat(@"  ,[B4ALAVG] AS '小線-4段-上左爐-平均',[B4ALMIN] AS '小線-4段-上左爐-最小',[B4ALMAX] AS '小線-4段-上左爐-最大'");
                sbSql.AppendFormat(@"  ,[B4BLAVG] AS '小線-4段-下左爐-平均',[B4BLMIN] AS '小線-4段-下左爐-最小',[B4BLMAX] AS '小線-4段-下左爐-最大'");
                sbSql.AppendFormat(@"  ,[B5ARAVG] AS '小線-5段-上右爐-平均',[B5ARMIN] AS '小線-5段-上右爐-最小',[B5ARMAX] AS '小線-5段-上右爐-最大'");
                sbSql.AppendFormat(@"  ,[B5BRAVG] AS '小線-5段-下右爐-平均',[B5BRMIN] AS '小線-5段-下右爐-最小',[B5BRMAX] AS '小線-5段-下右爐-最大'");
                sbSql.AppendFormat(@"  ,[B5AMAVG] AS '小線-5段-上中爐-平均',[B5AMMIN] AS '小線-5段-上中爐-最小',[B5AMMAX] AS '小線-5段-上中爐-最大'");
                sbSql.AppendFormat(@"  ,[B5BMAVG] AS '小線-5段-下中爐-平均',[B5BMMIN] AS '小線-5段-下中爐-最小',[B5BMMAX] AS '小線-5段-下中爐-最大'");
                sbSql.AppendFormat(@"  ,[B5ALAVG] AS '小線-5段-上左爐-平均',[B5ALMIN] AS '小線-5段-上左爐-最小',[B5ALMAX] AS '小線-5段-上左爐-最大'");
                sbSql.AppendFormat(@"  ,[B5BLAVG] AS '小線-5段-下左爐-平均',[B5BLMIN] AS '小線-5段-下左爐-最小',[B5BLMAX] AS '小線-5段-下左爐-最大'");
                sbSql.AppendFormat(@"  ,[ID]");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCTADAILY]");
                sbSql.AppendFormat(@"  WHERE CONVERT(VARCHAR, [SDATES], 112)='{0}'", IDDATE);
                sbSql.AppendFormat(@"  ORDER BY  [TA001],[TA002],[TA021]");
                sbSql.AppendFormat(@"  ");


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {                        //dataGridView1.Rows.Clear();
                        dataGridView1.DataSource = ds1.Tables["ds1"];
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
                sqlConn.Close();
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox1.Text) && !string.IsNullOrEmpty(textBox2.Text))
            {
                SEARCHMOCTA(textBox1.Text.Trim(), textBox2.Text.Trim());
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox1.Text) && !string.IsNullOrEmpty(textBox2.Text))
            {
                SEARCHMOCTA(textBox1.Text.Trim(), textBox2.Text.Trim());
            }
        }

        public void SEARCHMOCTA(string TA001, string TA002)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();

            DataSet ds1 = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@" SELECT TA006,TA015,TA017,TA021,TA034,TA035 FROM [TK].dbo.MOCTA WHERE TA001='{0}' AND TA002='{1}' ", TA001, TA002);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count == 0)
                {

                }
                else
                {
                    if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        textBox111.Text = ds1.Tables["ds1"].Rows[0]["TA006"].ToString();
                        textBox112.Text = ds1.Tables["ds1"].Rows[0]["TA034"].ToString();
                        textBox113.Text = ds1.Tables["ds1"].Rows[0]["TA035"].ToString();
                        textBox121.Text = ds1.Tables["ds1"].Rows[0]["TA015"].ToString();
                        textBox122.Text = ds1.Tables["ds1"].Rows[0]["TA017"].ToString();
                        textBox123.Text ="";

                        comboBox1.SelectedValue = ds1.Tables["ds1"].Rows[0]["TA021"].ToString();

                    }
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

        public void SETTEXTBOX1()
        {
            textBox1.ReadOnly = false;
            textBox2.ReadOnly = false;

            textBox1.Text = null;
            textBox2.Text = null;

            textBox111.Text = null;
            textBox112.Text = null;
            textBox113.Text = null;
            textBox121.Text = null;
            textBox122.Text = null;
            textBox123.Text = null;


            textBox211.Text = "0";
            textBox231.Text = "0";
            textBox232.Text = "0";
            textBox233.Text = "0";
            textBox234.Text = "0";
            textBox235.Text = "0";
            textBox236.Text = "0";
            textBox241.Text = "0";
            textBox242.Text = "0";
            textBox243.Text = "0";
            textBox244.Text = "0";
            textBox245.Text = "0";
            textBox246.Text = "0";
            textBox251.Text = "0";
            textBox252.Text = "0";
            textBox253.Text = "0";
            textBox254.Text = "0";
            textBox255.Text = "0";
            textBox256.Text = "0";
            textBox261.Text = "0";
            textBox262.Text = "0";
            textBox263.Text = "0";
            textBox264.Text = "0";
            textBox265.Text = "0";
            textBox266.Text = "0";


            textBox311.Text = null;

            textBox331.Text = null;
            textBox332.Text = null;
            textBox333.Text = null;
            textBox334.Text = null;
            textBox335.Text = null;
            textBox336.Text = null;
            textBox341.Text = null;
            textBox342.Text = null;
            textBox343.Text = null;
            textBox344.Text = null;
            textBox345.Text = null;
            textBox346.Text = null;
            textBox351.Text = null;
            textBox352.Text = null;
            textBox353.Text = null;
            textBox354.Text = null;
            textBox355.Text = null;
            textBox356.Text = null;
            textBox361.Text = null;
            textBox362.Text = null;
            textBox363.Text = null;
            textBox364.Text = null;
            textBox365.Text = null;
            textBox366.Text = null;
            textBox371.Text = null;
            textBox372.Text = null;
            textBox373.Text = null;
            textBox374.Text = null;
            textBox375.Text = null;
            textBox376.Text = null;
        }

        public void SETTEXTBOX2()
        {
            textBox1.ReadOnly = true;
            textBox2.ReadOnly = true;

            textBox1.Text = null;
            textBox2.Text = null;
        }

        public void ADDMOCTADAILY(string ID,string TA001, string TA002, string TA021, string MB001, string MB002, string MB003, string TA017, string NUM, string NGNUM, string SDATES, string EDATES
            , string ASPEED
            , string A1AAVG, string A1AMIN, string A1AMAX, string A1BAVG, string A1BMIN, string A1BMAX
            , string A2AAVG, string A2AMIN, string A2AMAX, string A2BAVG, string A2BMIN, string A2BMAX
            , string A3AAVG, string A3AMIN, string A3AMAX, string A3BAVG, string A3BMIN, string A3BMAX
            , string A4AAVG, string A4AMIN, string A4AMAX, string A4BAVG, string A4BMIN, string A4BMAX
            )
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat("  INSERT INTO  [TKMOC].[dbo].[MOCTADAILY] ");
                sbSql.AppendFormat("  ([ID],[TA001],[TA002],[TA021],[MB001],[MB002],[MB003],[TA017],[NUM],[NGNUM],[SDATES],[EDATES] ");
                sbSql.AppendFormat("  ,[ASPEED] ");
                sbSql.AppendFormat("  ,[A1AAVG],[A1AMIN],[A1AMAX],[A1BAVG],[A1BMIN],[A1BMAX]");
                sbSql.AppendFormat("  ,[A2AAVG],[A2AMIN],[A2AMAX],[A2BAVG],[A2BMIN],[A2BMAX]");
                sbSql.AppendFormat("  ,[A3AAVG],[A3AMIN],[A3AMAX],[A3BAVG],[A3BMIN],[A3BMAX]");
                sbSql.AppendFormat("  ,[A4AAVG],[A4AMIN],[A4AMAX],[A4BAVG],[A4BMIN],[A4BMAX]");
                sbSql.AppendFormat("  )");
                sbSql.AppendFormat("  VALUES");
                sbSql.AppendFormat("  ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}'", ID,TA001, TA002, TA021, MB001, MB002, MB003, TA017, NUM, NGNUM, SDATES, EDATES);
                sbSql.AppendFormat("  ,'{0}'", ASPEED);
                sbSql.AppendFormat("  ,'{0}','{1}','{2}','{3}','{4}','{5}'", A1AAVG, A1AMIN, A1AMAX, A1BAVG, A1BMIN, A1BMAX);
                sbSql.AppendFormat("  ,'{0}','{1}','{2}','{3}','{4}','{5}'", A2AAVG, A2AMIN, A2AMAX, A2BAVG, A2BMIN, A2BMAX);
                sbSql.AppendFormat("  ,'{0}','{1}','{2}','{3}','{4}','{5}'", A3AAVG, A3AMIN, A3AMAX, A3BAVG, A3BMIN, A3BMAX);
                sbSql.AppendFormat("  ,'{0}','{1}','{2}','{3}','{4}','{5}'", A4AAVG, A4AMIN, A4AMAX, A4BAVG, A4BMIN, A4BMAX);
                sbSql.AppendFormat("  )");
                sbSql.AppendFormat("  ");
                sbSql.AppendFormat("  ");
                sbSql.AppendFormat("  ");

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

        public void UPDATEMOCTADAILY(string ID,string NUM,string NGNUM,string SDATES,string EDATES
            , string ASPEED
            , string A1AAVG, string A1AMIN, string A1AMAX, string A1BAVG, string A1BMIN, string A1BMAX
            , string A2AAVG, string A2AMIN, string A2AMAX, string A2BAVG, string A2BMIN, string A2BMAX
            , string A3AAVG, string A3AMIN, string A3AMAX, string A3BAVG, string A3BMIN, string A3BMAX
            , string A4AAVG, string A4AMIN, string A4AMAX, string A4BAVG, string A4BMIN, string A4BMAX
            )
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();


                sbSql.AppendFormat("  UPDATE [TKMOC].[dbo].[MOCTADAILY]");
                sbSql.AppendFormat("  SET [NUM]='{0}',[NGNUM]='{1}',[SDATES]='{2}',[EDATES]='{3}'", NUM,NGNUM,SDATES,EDATES);
                sbSql.AppendFormat("  ,[ASPEED]='{0}'", ASPEED);
                sbSql.AppendFormat("  ,A1AAVG='{0}',A1AMIN='{1}',A1AMAX='{2}',A1BAVG='{3}',A1BMIN='{4}',A1BMAX='{5}'", A1AAVG, A1AMIN, A1AMAX, A1BAVG, A1BMIN, A1BMAX);
                sbSql.AppendFormat("  ,A2AAVG='{0}',A2AMIN='{1}',A2AMAX='{2}',A2BAVG='{3}',A2BMIN='{4}',A2BMAX='{5}'", A2AAVG, A2AMIN, A2AMAX, A2BAVG, A2BMIN, A2BMAX);
                sbSql.AppendFormat("  ,A3AAVG='{0}',A3AMIN='{1}',A3AMAX='{2}',A3BAVG='{3}',A3BMIN='{4}',A3BMAX='{5}'", A3AAVG, A3AMIN, A3AMAX, A3BAVG, A3BMIN, A3BMAX);
                sbSql.AppendFormat("  ,A4AAVG='{0}',A4AMIN='{1}',A4AMAX='{2}',A4BAVG='{3}',A4BMIN='{4}',A4BMAX='{5}'", A4AAVG, A4AMIN, A4AMAX, A4BAVG, A4BMIN, A4BMAX);
                sbSql.AppendFormat("  WHERE ID='{0}'",ID);
                sbSql.AppendFormat("  ");
                sbSql.AppendFormat("  ");

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

        public void DELMOCTADAILY(string ID)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();


                sbSql.AppendFormat("  DELETE [TKMOC].[dbo].[MOCTADAILY]");
                sbSql.AppendFormat("  WHERE ID='{0}'", ID);
                sbSql.AppendFormat("  ");
                sbSql.AppendFormat("  ");

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
            if (dataGridView1.Rows.Count >= 1)
            {
                int rowindex = dataGridView1.CurrentRow.Index;

                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    ID = row.Cells["ID"].Value.ToString();
                    textBox151.Text = row.Cells["ID"].Value.ToString();

                    textBox1.Text = row.Cells["製令"].Value.ToString();
                    textBox2.Text = row.Cells["單號"].Value.ToString();

                    textBox111.Text = row.Cells["品號"].Value.ToString();
                    textBox112.Text = row.Cells["品名"].Value.ToString();
                    textBox113.Text = row.Cells["規格"].Value.ToString();
                    textBox121.Text = row.Cells["生產量"].Value.ToString();
                    textBox122.Text= row.Cells["入庫量"].Value.ToString();
                    textBox123.Text = row.Cells["未熟量"].Value.ToString();

                    comboBox1.Text = row.Cells["線別"].Value.ToString();

                    dateTimePicker2.Value = Convert.ToDateTime(row.Cells["開始時間"].Value.ToString());
                    dateTimePicker3.Value = Convert.ToDateTime(row.Cells["結束時間"].Value.ToString());

                    textBox211.Text = row.Cells["大線爐速"].Value.ToString();
                    textBox231.Text = row.Cells["大線-1段-上爐-平均"].Value.ToString();
                    textBox232.Text = row.Cells["大線-1段-上爐-最小"].Value.ToString();
                    textBox233.Text = row.Cells["大線-1段-上爐-最大"].Value.ToString();
                    textBox234.Text = row.Cells["大線-1段-下爐-平均"].Value.ToString();
                    textBox235.Text = row.Cells["大線-1段-下爐-最小"].Value.ToString();
                    textBox236.Text = row.Cells["大線-1段-下爐-最大"].Value.ToString();
                    textBox241.Text = row.Cells["大線-2段-上爐-平均"].Value.ToString();
                    textBox242.Text = row.Cells["大線-2段-上爐-最小"].Value.ToString();
                    textBox243.Text = row.Cells["大線-2段-上爐-最大"].Value.ToString();
                    textBox244.Text = row.Cells["大線-2段-下爐-平均"].Value.ToString();
                    textBox245.Text = row.Cells["大線-2段-下爐-最小"].Value.ToString();
                    textBox246.Text = row.Cells["大線-2段-下爐-最大"].Value.ToString();
                    textBox251.Text = row.Cells["大線-3段-上爐-平均"].Value.ToString();
                    textBox252.Text = row.Cells["大線-3段-上爐-最小"].Value.ToString();
                    textBox253.Text = row.Cells["大線-3段-上爐-最大"].Value.ToString();
                    textBox254.Text = row.Cells["大線-3段-下爐-平均"].Value.ToString();
                    textBox255.Text = row.Cells["大線-3段-下爐-最小"].Value.ToString();
                    textBox256.Text = row.Cells["大線-3段-下爐-最大"].Value.ToString();
                    textBox261.Text = row.Cells["大線-4段-上爐-平均"].Value.ToString();
                    textBox262.Text = row.Cells["大線-4段-上爐-最小"].Value.ToString();
                    textBox263.Text = row.Cells["大線-4段-上爐-最大"].Value.ToString();
                    textBox264.Text = row.Cells["大線-4段-下爐-平均"].Value.ToString();
                    textBox265.Text = row.Cells["大線-4段-下爐-最小"].Value.ToString();
                    textBox266.Text = row.Cells["大線-4段-下爐-最大"].Value.ToString();

                    textBox311.Text = row.Cells["小線爐速"].Value.ToString();

                    textBox331.Text = row.Cells["小線-1段-上右爐-平均"].Value.ToString();
                    textBox332.Text = row.Cells["小線-1段-下右爐-平均"].Value.ToString();
                    textBox333.Text = row.Cells["小線-1段-上中爐-平均"].Value.ToString();
                    textBox334.Text = row.Cells["小線-1段-下中爐-平均"].Value.ToString();
                    textBox335.Text = row.Cells["小線-1段-上左爐-平均"].Value.ToString();
                    textBox336.Text = row.Cells["小線-1段-下左爐-平均"].Value.ToString();
                    textBox341.Text = row.Cells["小線-2段-上右爐-平均"].Value.ToString();
                    textBox342.Text = row.Cells["小線-2段-下右爐-平均"].Value.ToString();
                    textBox343.Text = row.Cells["小線-2段-上中爐-平均"].Value.ToString();
                    textBox344.Text = row.Cells["小線-2段-下中爐-平均"].Value.ToString();
                    textBox345.Text = row.Cells["小線-2段-上左爐-平均"].Value.ToString();
                    textBox346.Text = row.Cells["小線-2段-下左爐-平均"].Value.ToString();
                    textBox351.Text = row.Cells["小線-3段-上右爐-平均"].Value.ToString();
                    textBox352.Text = row.Cells["小線-3段-下右爐-平均"].Value.ToString();
                    textBox353.Text = row.Cells["小線-3段-上中爐-平均"].Value.ToString();
                    textBox354.Text = row.Cells["小線-3段-下中爐-平均"].Value.ToString();
                    textBox355.Text = row.Cells["小線-3段-上左爐-平均"].Value.ToString();
                    textBox356.Text = row.Cells["小線-3段-下左爐-平均"].Value.ToString();
                    textBox361.Text = row.Cells["小線-4段-上右爐-平均"].Value.ToString();
                    textBox362.Text = row.Cells["小線-4段-下右爐-平均"].Value.ToString();
                    textBox363.Text = row.Cells["小線-4段-上中爐-平均"].Value.ToString();
                    textBox364.Text = row.Cells["小線-4段-下中爐-平均"].Value.ToString();
                    textBox365.Text = row.Cells["小線-4段-上左爐-平均"].Value.ToString();
                    textBox366.Text = row.Cells["小線-4段-下左爐-平均"].Value.ToString();
                    textBox371.Text = row.Cells["小線-5段-上右爐-平均"].Value.ToString();
                    textBox372.Text = row.Cells["小線-5段-下右爐-平均"].Value.ToString();
                    textBox373.Text = row.Cells["小線-5段-上中爐-平均"].Value.ToString();
                    textBox374.Text = row.Cells["小線-5段-下中爐-平均"].Value.ToString();
                    textBox375.Text = row.Cells["小線-5段-上左爐-平均"].Value.ToString();
                    textBox376.Text = row.Cells["小線-5段-下左爐-平均"].Value.ToString();
                }
            }
            else
            {
                
                ID = null;

                textBox111.Text = null;
                textBox112.Text = null;
                textBox113.Text = null;
                textBox121.Text = null;
                textBox122.Text = null;
                textBox123.Text = null;


                textBox211.Text = null;
                textBox231.Text = null;
                textBox232.Text = null;
                textBox233.Text = null;
                textBox234.Text = null;
                textBox235.Text = null;
                textBox236.Text = null;
                textBox241.Text = null;
                textBox242.Text = null;
                textBox243.Text = null;
                textBox244.Text = null;
                textBox245.Text = null;
                textBox246.Text = null;
                textBox251.Text = null;
                textBox252.Text = null;
                textBox253.Text = null;
                textBox254.Text = null;
                textBox255.Text = null;
                textBox256.Text = null;
                textBox261.Text = null;
                textBox262.Text = null;
                textBox263.Text = null;
                textBox264.Text = null;
                textBox265.Text = null;
                textBox266.Text = null;


                textBox311.Text = null;

                textBox331.Text = null;
                textBox332.Text = null;
                textBox333.Text = null;
                textBox334.Text = null;
                textBox335.Text = null;
                textBox336.Text = null; 
                textBox341.Text = null;
                textBox342.Text = null;
                textBox343.Text = null;
                textBox344.Text = null; 
                textBox345.Text = null;
                textBox346.Text = null;
                textBox351.Text = null;
                textBox352.Text = null;
                textBox353.Text = null;
                textBox354.Text = null;
                textBox355.Text = null;
                textBox356.Text = null;
                textBox361.Text = null;
                textBox362.Text = null;
                textBox363.Text = null;
                textBox364.Text = null;
                textBox365.Text = null;
                textBox366.Text = null;
                textBox371.Text = null;
                textBox372.Text = null;
                textBox373.Text = null;
                textBox374.Text = null;
                textBox375.Text = null;
                textBox376.Text = null;

            }
        }

        public void UPDATEMOCTADAILYDETAIL(string ID,string SDATES,string EDATES)
        {
            DATAMOCTADAILY MOCTADAILY = new DATAMOCTADAILY();
            MOCTADAILY = CALSDEEP(MOCTADAILY, SDATES, EDATES);
            MOCTADAILY = CALTEMPERHUMI(MOCTADAILY,SDATES,EDATES);
            MOCTADAILY = CALTEMPERAVG(MOCTADAILY, SDATES, EDATES);

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat("  UPDATE [TKMOC].[dbo].[MOCTADAILY]");
                sbSql.AppendFormat("  SET ");
                sbSql.AppendFormat("  [BSPEED]='{0}'", MOCTADAILY.BSPEED);
                sbSql.AppendFormat("  ,[TEMPERAVG]='{0}',[TEMPERMIN]='{1}',[TEMPERMAX]='{2}'", MOCTADAILY.TEMPERAVG, MOCTADAILY.TEMPERMIN, MOCTADAILY.TEMPERMAX);
                sbSql.AppendFormat("  ,[HUMIAVG]='{0}',[HUMIMIN]='{1}',[HUMIMAX]='{2}'", MOCTADAILY.HUMIAVG, MOCTADAILY.HUMIMIN, MOCTADAILY.HUMIMAX);
                sbSql.AppendFormat("  ,[B1ARAVG]='{0}',[B1ARMIN]='{1}',[B1ARMAX]='{2}',[B1BRAVG]='{3}',[B1BRMIN]='{4}',[B1BRMAX]='{5}',[B1AMAVG]='{6}',[B1AMMIN]='{7}',[B1AMMAX]='{8}',[B1BMAVG]='{9}',[B1BMMIN]='{10}',[B1BMMAX]='{11}',[B1ALAVG]='{12}',[B1ALMIN]='{13}',[B1ALMAX]='{14}',[B1BLAVG]='{15}',[B1BLMIN]='{16}',[B1BLMAX]='{17}'", MOCTADAILY.B1ARAVG, MOCTADAILY.B1ARMIN, MOCTADAILY.B1ARMAX, MOCTADAILY.B1BRAVG, MOCTADAILY.B1BRMIN, MOCTADAILY.B1BRMAX, MOCTADAILY.B1AMAVG, MOCTADAILY.B1AMMIN, MOCTADAILY.B1AMMAX, MOCTADAILY.B1BMAVG, MOCTADAILY.B1BMMIN, MOCTADAILY.B1BMMAX, MOCTADAILY.B1ALAVG, MOCTADAILY.B1ALMIN, MOCTADAILY.B1ALMAX, MOCTADAILY.B1BLAVG, MOCTADAILY.B1BLMIN, MOCTADAILY.B1BLMAX);
                sbSql.AppendFormat("  ,[B2ARAVG]='{0}',[B2ARMIN]='{1}',[B2ARNAX]='{2}',[B2BRAVG]='{3}',[B2BRMIN]='{4}',[B2BRMAX]='{5}',[B2AMAVG]='{6}',[B2AMMIN]='{7}',[B2AMMAX]='{8}',[B2BMAVG]='{9}',[B2BMMIN]='{10}',[B2BMMAX]='{11}',[B2ALAVG]='{12}',[B2ALMIN]='{13}',[B2ALMAX]='{14}',[B2BLAVG]='{15}',[B2BLMIN]='{16}',[B2BLMAX]='{17}'", MOCTADAILY.B2ARAVG, MOCTADAILY.B2ARMIN, MOCTADAILY.B2ARMAX, MOCTADAILY.B2BRAVG, MOCTADAILY.B2BRMIN, MOCTADAILY.B2BRMAX, MOCTADAILY.B2AMAVG, MOCTADAILY.B2AMMIN, MOCTADAILY.B2AMMAX, MOCTADAILY.B2BMAVG, MOCTADAILY.B2BMMIN, MOCTADAILY.B2BMMAX, MOCTADAILY.B2ALAVG, MOCTADAILY.B2ALMIN, MOCTADAILY.B2ALMAX, MOCTADAILY.B2BLAVG, MOCTADAILY.B2BLMIN, MOCTADAILY.B2BLMAX);
                sbSql.AppendFormat("  ,[B3ARAVG]='{0}',[B3ARMIN]='{1}',[B3ARMAX]='{2}',[B3BRAVG]='{3}',[B3BRMIN]='{4}',[B3BRMAX]='{5}',[B3AMAVG]='{6}',[B3AMMIN]='{7}',[B3AMMAX]='{8}',[B3BMAVG]='{9}',[B3BMMIN]='{10}',[B3BMMAX]='{11}',[B3ALAVG]='{12}',[B3ALMIN]='{13}',[B3ALMAX]='{14}',[B3BLAVG]='{15}',[B3BLMIN]='{16}',[B3BLMAX]='{17}'", MOCTADAILY.B3ARAVG, MOCTADAILY.B3ARMIN, MOCTADAILY.B3ARMAX, MOCTADAILY.B3BRAVG, MOCTADAILY.B3BRMIN, MOCTADAILY.B3BRMAX, MOCTADAILY.B3AMAVG, MOCTADAILY.B3AMMIN, MOCTADAILY.B3AMMAX, MOCTADAILY.B3BMAVG, MOCTADAILY.B3BMMIN, MOCTADAILY.B3BMMAX, MOCTADAILY.B3ALAVG, MOCTADAILY.B3ALMIN, MOCTADAILY.B3ALMAX, MOCTADAILY.B3BLAVG, MOCTADAILY.B3BLMIN, MOCTADAILY.B3BLMAX);
                sbSql.AppendFormat("  ,[B4ARAVG]='{0}',[B4ARMIN]='{1}',[B4ARMAX]='{2}',[B4BRAVG]='{3}',[B4BRMIN]='{4}',[B4BRMAX]='{5}',[B4AMAVG]='{6}',[B4AMMIN]='{7}',[B4AMMAX]='{8}',[B4BMAVG]='{9}',[B4BMMIN]='{10}',[B4BMMAX]='{11}',[B4ALAVG]='{12}',[B4ALMIN]='{13}',[B4ALMAX]='{14}',[B4BLAVG]='{15}',[B4BLMIN]='{16}',[B4BLMAX]='{17}'", MOCTADAILY.B4ARAVG, MOCTADAILY.B4ARMIN, MOCTADAILY.B4ARMAX, MOCTADAILY.B4BRAVG, MOCTADAILY.B4BRMIN, MOCTADAILY.B4BRMAX, MOCTADAILY.B4AMAVG, MOCTADAILY.B4AMMIN, MOCTADAILY.B4AMMAX, MOCTADAILY.B4BMAVG, MOCTADAILY.B4BMMIN, MOCTADAILY.B4BMMAX, MOCTADAILY.B4ALAVG, MOCTADAILY.B4ALMIN, MOCTADAILY.B4ALMAX, MOCTADAILY.B4BLAVG, MOCTADAILY.B4BLMIN, MOCTADAILY.B4BLMAX);
                sbSql.AppendFormat("  ,[B5ARAVG]='{0}',[B5ARMIN]='{1}',[B5ARMAX]='{2}',[B5BRAVG]='{3}',[B5BRMIN]='{4}',[B5BRMAX]='{5}',[B5AMAVG]='{6}',[B5AMMIN]='{7}',[B5AMMAX]='{8}',[B5BMAVG]='{9}',[B5BMMIN]='{10}',[B5BMMAX]='{11}',[B5ALAVG]='{12}',[B5ALMIN]='{13}',[B5ALMAX]='{14}',[B5BLAVG]='{15}',[B5BLMIN]='{16}',[B5BLMAX]='{17}'", MOCTADAILY.B5ARAVG, MOCTADAILY.B5ARMIN, MOCTADAILY.B5ARMAX, MOCTADAILY.B5BRAVG, MOCTADAILY.B5BRMIN, MOCTADAILY.B5BRMAX, MOCTADAILY.B5AMAVG, MOCTADAILY.B5AMMIN, MOCTADAILY.B5AMMAX, MOCTADAILY.B5BMAVG, MOCTADAILY.B5BMMIN, MOCTADAILY.B5BMMAX, MOCTADAILY.B5ALAVG, MOCTADAILY.B5ALMIN, MOCTADAILY.B5ALMAX, MOCTADAILY.B5BLAVG, MOCTADAILY.B5BLMIN, MOCTADAILY.B5BLMAX);
                sbSql.AppendFormat("  WHERE ID='{0}'", ID);
                sbSql.AppendFormat("  ");
                sbSql.AppendFormat("  ");

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

            //MessageBox.Show(MOCTADAILY.TEMPERAVG+" "+ MOCTADAILY.B1ALAVG);
        }

        public DATAMOCTADAILY CALSDEEP(DATAMOCTADAILY MOCTADAILY, string SDATES, string EDATES)
        {

            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();

            DataSet ds1 = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbMOC"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  SELECT AVG(CONVERT(decimal(16,4),控項_1)) AS 'BSPEED'");
                sbSql.AppendFormat(@"  FROM [TK_FOOD].[dbo].[log_table]");
                sbSql.AppendFormat(@"  WHERE [機台名稱]='烤爐_小線' AND [類型]='速度段'");
                sbSql.AppendFormat(@"  AND CONVERT(VARCHAR,[日期時間], 120)>='{0}' AND CONVERT(VARCHAR,[日期時間], 120)<='{1}'", SDATES, EDATES);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count == 0)
                {

                }
                else
                {
                    if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        MOCTADAILY.BSPEED = ds1.Tables["ds1"].Rows[0]["BSPEED"].ToString();
                        

                    }
                }

            }
            catch
            {

            }
            finally
            {
                sqlConn.Close();
            }

            return MOCTADAILY;
        }

        public DATAMOCTADAILY CALTEMPERHUMI(DATAMOCTADAILY MOCTADAILY, string SDATES, string EDATES)
        {

            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();

            DataSet ds1 = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbMOC"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  SELECT AVG(CONVERT(decimal(16,4),[控項_1])) AS TEMPERAVG,MIN(CONVERT(decimal(16,4),[控項_1])) AS TEMPERMIN,MAX(CONVERT(decimal(16,4),[控項_1])) AS TEMPERMAX,AVG(CONVERT(decimal(16,4),[控項_4])) AS HUMIAVG,MIN(CONVERT(decimal(16,4),[控項_4])) AS HUMIMIN,MAX(CONVERT(decimal(16,4),[控項_4])) AS HUMIMAX");
                sbSql.AppendFormat(@"  FROM [TK_FOOD].[dbo].[log_table]");
                sbSql.AppendFormat(@"  WHERE [機台名稱]='溫濕度6'");
                sbSql.AppendFormat(@"  AND CONVERT(VARCHAR,[日期時間], 120)>='{0}' AND CONVERT(VARCHAR,[日期時間], 120)<='{1}'",SDATES,EDATES);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count == 0)
                {

                }
                else
                {
                    if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        MOCTADAILY.TEMPERAVG = ds1.Tables["ds1"].Rows[0]["TEMPERAVG"].ToString();
                        MOCTADAILY.TEMPERMIN = ds1.Tables["ds1"].Rows[0]["TEMPERMIN"].ToString();
                        MOCTADAILY.TEMPERMAX = ds1.Tables["ds1"].Rows[0]["TEMPERMAX"].ToString();
                        MOCTADAILY.HUMIAVG = ds1.Tables["ds1"].Rows[0]["HUMIAVG"].ToString();
                        MOCTADAILY.HUMIMIN = ds1.Tables["ds1"].Rows[0]["HUMIMIN"].ToString();
                        MOCTADAILY.HUMIMAX = ds1.Tables["ds1"].Rows[0]["HUMIMAX"].ToString();

                    }
                }

            }
            catch
            {

            }
            finally
            {
                sqlConn.Close();
            }

            return MOCTADAILY;
        }

        public DATAMOCTADAILY CALTEMPERAVG(DATAMOCTADAILY MOCTADAILY, string SDATES, string EDATES)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();

            DataSet ds1 = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbMOC"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  SELECT '1' AS NOS");
                sbSql.AppendFormat(@"  ,AVG(CONVERT(decimal(16,4),[控項_2])) AS 'BARAVG',MIN(CONVERT(decimal(16,4),[控項_2])) AS 'BARMIN',MAX(CONVERT(decimal(16,4),[控項_2])) AS 'BARMAX'");
                sbSql.AppendFormat(@"  ,AVG(CONVERT(decimal(16,4),[控項_6]))  AS 'BBRAVG',MIN(CONVERT(decimal(16,4),[控項_6]))  AS 'BBRMIN',MAX(CONVERT(decimal(16,4),[控項_6]))  AS 'BBRMAX'");
                sbSql.AppendFormat(@"  ,AVG(CONVERT(decimal(16,4),[控項_10]))  AS 'BAMAVG',MIN(CONVERT(decimal(16,4),[控項_10]))  AS 'BAMMIN',MAX(CONVERT(decimal(16,4),[控項_10]))  AS 'BAMMAX'");
                sbSql.AppendFormat(@"  ,AVG(CONVERT(decimal(16,4),[控項_14]))  AS 'BBMAVG',MIN(CONVERT(decimal(16,4),[控項_14]))  AS 'BBMMIN',MAX(CONVERT(decimal(16,4),[控項_14]))  AS 'BBMMAX'");
                sbSql.AppendFormat(@"  ,AVG(CONVERT(decimal(16,4),[控項_18]))  AS 'BALAVG',MIN(CONVERT(decimal(16,4),[控項_18]))  AS 'BALMIN',MAX(CONVERT(decimal(16,4),[控項_18]))  AS 'BALMAX'");
                sbSql.AppendFormat(@"  ,AVG(CONVERT(decimal(16,4),[控項_22]))  AS 'BBLAVG',MIN(CONVERT(decimal(16,4),[控項_22]))  AS 'BBLMIN',MAX(CONVERT(decimal(16,4),[控項_22]))  AS 'BBLMAX' ");
                sbSql.AppendFormat(@"  FROM [TK_FOOD].[dbo].[log_table]");
                sbSql.AppendFormat(@"  WHERE [機台名稱]='烤爐_小線' AND [類型]='第一段'");
                sbSql.AppendFormat(@"  AND CONVERT(VARCHAR,[日期時間], 120)>='{0}' AND CONVERT(VARCHAR,[日期時間], 120)<='{1}'", SDATES, EDATES);
                sbSql.AppendFormat(@"  UNION ALL");
                sbSql.AppendFormat(@"  SELECT '2' AS NOS");
                sbSql.AppendFormat(@"  ,AVG(CONVERT(decimal(16,4),[控項_2])) AS 'BARAVG',MIN(CONVERT(decimal(16,4),[控項_2])) AS 'BARMIN',MAX(CONVERT(decimal(16,4),[控項_2])) AS 'BARMAX'");
                sbSql.AppendFormat(@"  ,AVG(CONVERT(decimal(16,4),[控項_6]))  AS 'BBRAVG',MIN(CONVERT(decimal(16,4),[控項_6]))  AS 'BBRMIN',MAX(CONVERT(decimal(16,4),[控項_6]))  AS 'BBRMAX'");
                sbSql.AppendFormat(@"  ,AVG(CONVERT(decimal(16,4),[控項_10]))  AS 'BAMAVG',MIN(CONVERT(decimal(16,4),[控項_10]))  AS 'BAMMIN',MAX(CONVERT(decimal(16,4),[控項_10]))  AS 'BAMMAX'");
                sbSql.AppendFormat(@"  ,AVG(CONVERT(decimal(16,4),[控項_14]))  AS 'BBMAVG',MIN(CONVERT(decimal(16,4),[控項_14]))  AS 'BBMMIN',MAX(CONVERT(decimal(16,4),[控項_14]))  AS 'BBMMAX'");
                sbSql.AppendFormat(@"  ,AVG(CONVERT(decimal(16,4),[控項_18]))  AS 'BALAVG',MIN(CONVERT(decimal(16,4),[控項_18]))  AS 'BALMIN',MAX(CONVERT(decimal(16,4),[控項_18]))  AS 'BALMAX'");
                sbSql.AppendFormat(@"  ,AVG(CONVERT(decimal(16,4),[控項_22]))  AS 'BBLAVG',MIN(CONVERT(decimal(16,4),[控項_22]))  AS 'BBLMIN',MAX(CONVERT(decimal(16,4),[控項_22]))  AS 'BBLMAX' ");
                sbSql.AppendFormat(@"  FROM [TK_FOOD].[dbo].[log_table]");
                sbSql.AppendFormat(@"  WHERE [機台名稱]='烤爐_小線' AND [類型]='第二段'");
                sbSql.AppendFormat(@"  AND CONVERT(VARCHAR,[日期時間], 120)>='{0}' AND CONVERT(VARCHAR,[日期時間], 120)<='{1}'", SDATES, EDATES);
                sbSql.AppendFormat(@"  UNION ALL");
                sbSql.AppendFormat(@"  SELECT '3' AS NOS");
                sbSql.AppendFormat(@"  ,AVG(CONVERT(decimal(16,4),[控項_2])) AS 'BARAVG',MIN(CONVERT(decimal(16,4),[控項_2])) AS 'BARMIN',MAX(CONVERT(decimal(16,4),[控項_2])) AS 'BARMAX'");
                sbSql.AppendFormat(@"  ,AVG(CONVERT(decimal(16,4),[控項_6]))  AS 'BBRAVG',MIN(CONVERT(decimal(16,4),[控項_6]))  AS 'BBRMIN',MAX(CONVERT(decimal(16,4),[控項_6]))  AS 'BBRMAX'");
                sbSql.AppendFormat(@"  ,AVG(CONVERT(decimal(16,4),[控項_10]))  AS 'BAMAVG',MIN(CONVERT(decimal(16,4),[控項_10]))  AS 'BAMMIN',MAX(CONVERT(decimal(16,4),[控項_10]))  AS 'BAMMAX'");
                sbSql.AppendFormat(@"  ,AVG(CONVERT(decimal(16,4),[控項_14]))  AS 'BBMAVG',MIN(CONVERT(decimal(16,4),[控項_14]))  AS 'BBMMIN',MAX(CONVERT(decimal(16,4),[控項_14]))  AS 'BBMMAX'");
                sbSql.AppendFormat(@"  ,AVG(CONVERT(decimal(16,4),[控項_18]))  AS 'BALAVG',MIN(CONVERT(decimal(16,4),[控項_18]))  AS 'BALMIN',MAX(CONVERT(decimal(16,4),[控項_18]))  AS 'BALMAX'");
                sbSql.AppendFormat(@"  ,AVG(CONVERT(decimal(16,4),[控項_22]))  AS 'BBLAVG',MIN(CONVERT(decimal(16,4),[控項_22]))  AS 'BBLMIN',MAX(CONVERT(decimal(16,4),[控項_22]))  AS 'BBLMAX' ");
                sbSql.AppendFormat(@"  FROM [TK_FOOD].[dbo].[log_table]");
                sbSql.AppendFormat(@"  WHERE [機台名稱]='烤爐_小線' AND [類型]='第三段'");
                sbSql.AppendFormat(@"  AND CONVERT(VARCHAR,[日期時間], 120)>='{0}' AND CONVERT(VARCHAR,[日期時間], 120)<='{1}'", SDATES, EDATES);
                sbSql.AppendFormat(@"  UNION ALL");
                sbSql.AppendFormat(@"  SELECT '4' AS NOS");
                sbSql.AppendFormat(@"  ,AVG(CONVERT(decimal(16,4),[控項_2])) AS 'BARAVG',MIN(CONVERT(decimal(16,4),[控項_2])) AS 'BARMIN',MAX(CONVERT(decimal(16,4),[控項_2])) AS 'BARMAX'");
                sbSql.AppendFormat(@"  ,AVG(CONVERT(decimal(16,4),[控項_6]))  AS 'BBRAVG',MIN(CONVERT(decimal(16,4),[控項_6]))  AS 'BBRMIN',MAX(CONVERT(decimal(16,4),[控項_6]))  AS 'BBRMAX'");
                sbSql.AppendFormat(@"  ,AVG(CONVERT(decimal(16,4),[控項_10]))  AS 'BAMAVG',MIN(CONVERT(decimal(16,4),[控項_10]))  AS 'BAMMIN',MAX(CONVERT(decimal(16,4),[控項_10]))  AS 'BAMMAX'");
                sbSql.AppendFormat(@"  ,AVG(CONVERT(decimal(16,4),[控項_14]))  AS 'BBMAVG',MIN(CONVERT(decimal(16,4),[控項_14]))  AS 'BBMMIN',MAX(CONVERT(decimal(16,4),[控項_14]))  AS 'BBMMAX'");
                sbSql.AppendFormat(@"  ,AVG(CONVERT(decimal(16,4),[控項_18]))  AS 'BALAVG',MIN(CONVERT(decimal(16,4),[控項_18]))  AS 'BALMIN',MAX(CONVERT(decimal(16,4),[控項_18]))  AS 'BALMAX'");
                sbSql.AppendFormat(@"  ,AVG(CONVERT(decimal(16,4),[控項_22]))  AS 'BBLAVG',MIN(CONVERT(decimal(16,4),[控項_22]))  AS 'BBLMIN',MAX(CONVERT(decimal(16,4),[控項_22]))  AS 'BBLMAX' ");
                sbSql.AppendFormat(@"  FROM [TK_FOOD].[dbo].[log_table]");
                sbSql.AppendFormat(@"  WHERE [機台名稱]='烤爐_小線' AND [類型]='第四段'");
                sbSql.AppendFormat(@"  AND CONVERT(VARCHAR,[日期時間], 120)>='{0}' AND CONVERT(VARCHAR,[日期時間], 120)<='{1}'", SDATES, EDATES);
                sbSql.AppendFormat(@"  UNION ALL");
                sbSql.AppendFormat(@"  SELECT '5' AS NOS");
                sbSql.AppendFormat(@"  ,AVG(CONVERT(decimal(16,4),[控項_2])) AS 'BARAVG',MIN(CONVERT(decimal(16,4),[控項_2])) AS 'BARMIN',MAX(CONVERT(decimal(16,4),[控項_2])) AS 'BARMAX'");
                sbSql.AppendFormat(@"  ,AVG(CONVERT(decimal(16,4),[控項_6]))  AS 'BBRAVG',MIN(CONVERT(decimal(16,4),[控項_6]))  AS 'BBRMIN',MAX(CONVERT(decimal(16,4),[控項_6]))  AS 'BBRMAX'");
                sbSql.AppendFormat(@"  ,AVG(CONVERT(decimal(16,4),[控項_10]))  AS 'BAMAVG',MIN(CONVERT(decimal(16,4),[控項_10]))  AS 'BAMMIN',MAX(CONVERT(decimal(16,4),[控項_10]))  AS 'BAMMAX'");
                sbSql.AppendFormat(@"  ,AVG(CONVERT(decimal(16,4),[控項_14]))  AS 'BBMAVG',MIN(CONVERT(decimal(16,4),[控項_14]))  AS 'BBMMIN',MAX(CONVERT(decimal(16,4),[控項_14]))  AS 'BBMMAX'");
                sbSql.AppendFormat(@"  ,AVG(CONVERT(decimal(16,4),[控項_18]))  AS 'BALAVG',MIN(CONVERT(decimal(16,4),[控項_18]))  AS 'BALMIN',MAX(CONVERT(decimal(16,4),[控項_18]))  AS 'BALMAX'");
                sbSql.AppendFormat(@"  ,AVG(CONVERT(decimal(16,4),[控項_22]))  AS 'BBLAVG',MIN(CONVERT(decimal(16,4),[控項_22]))  AS 'BBLMIN',MAX(CONVERT(decimal(16,4),[控項_22]))  AS 'BBLMAX' ");
                sbSql.AppendFormat(@"  FROM [TK_FOOD].[dbo].[log_table]");
                sbSql.AppendFormat(@"  WHERE [機台名稱]='烤爐_小線' AND [類型]='第五段'");
                sbSql.AppendFormat(@"  AND CONVERT(VARCHAR,[日期時間], 120)>='{0}' AND CONVERT(VARCHAR,[日期時間], 120)<='{1}'", SDATES, EDATES);  
                sbSql.AppendFormat(@"  ORDER BY NOS");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count == 0)
                {

                }
                else
                {
                    if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        MOCTADAILY.B1ARAVG = ds1.Tables["ds1"].Rows[0]["BARAVG"].ToString();
                        MOCTADAILY.B1ARMIN = ds1.Tables["ds1"].Rows[0]["BARMIN"].ToString();
                        MOCTADAILY.B1ARMAX = ds1.Tables["ds1"].Rows[0]["BARMAX"].ToString();
                        MOCTADAILY.B1BRAVG = ds1.Tables["ds1"].Rows[0]["BBRAVG"].ToString();
                        MOCTADAILY.B1BRMIN = ds1.Tables["ds1"].Rows[0]["BBRMIN"].ToString();
                        MOCTADAILY.B1BRMAX = ds1.Tables["ds1"].Rows[0]["BBRMAX"].ToString();
                        MOCTADAILY.B1AMAVG = ds1.Tables["ds1"].Rows[0]["BAMAVG"].ToString();
                        MOCTADAILY.B1AMMIN = ds1.Tables["ds1"].Rows[0]["BAMMIN"].ToString();
                        MOCTADAILY.B1AMMAX = ds1.Tables["ds1"].Rows[0]["BAMMAX"].ToString();
                        MOCTADAILY.B1BMAVG = ds1.Tables["ds1"].Rows[0]["BBMAVG"].ToString();
                        MOCTADAILY.B1BMMIN = ds1.Tables["ds1"].Rows[0]["BBMMIN"].ToString();
                        MOCTADAILY.B1BMMAX = ds1.Tables["ds1"].Rows[0]["BBMMAX"].ToString();
                        MOCTADAILY.B1ALAVG = ds1.Tables["ds1"].Rows[0]["BALAVG"].ToString();
                        MOCTADAILY.B1ALMIN = ds1.Tables["ds1"].Rows[0]["BALMIN"].ToString();
                        MOCTADAILY.B1ALMAX = ds1.Tables["ds1"].Rows[0]["BALMAX"].ToString();
                        MOCTADAILY.B1BLAVG = ds1.Tables["ds1"].Rows[0]["BBLAVG"].ToString();
                        MOCTADAILY.B1BLMIN = ds1.Tables["ds1"].Rows[0]["BBLMIN"].ToString();
                        MOCTADAILY.B1BLMAX = ds1.Tables["ds1"].Rows[0]["BBLMAX"].ToString();

                        MOCTADAILY.B2ARAVG = ds1.Tables["ds1"].Rows[1]["BARAVG"].ToString();
                        MOCTADAILY.B2ARMIN = ds1.Tables["ds1"].Rows[1]["BARMIN"].ToString();
                        MOCTADAILY.B2ARMAX = ds1.Tables["ds1"].Rows[1]["BARMAX"].ToString();
                        MOCTADAILY.B2BRAVG = ds1.Tables["ds1"].Rows[1]["BBRAVG"].ToString();
                        MOCTADAILY.B2BRMIN = ds1.Tables["ds1"].Rows[1]["BBRMIN"].ToString();
                        MOCTADAILY.B2BRMAX = ds1.Tables["ds1"].Rows[1]["BBRMAX"].ToString();
                        MOCTADAILY.B2AMAVG = ds1.Tables["ds1"].Rows[1]["BAMAVG"].ToString();
                        MOCTADAILY.B2AMMIN = ds1.Tables["ds1"].Rows[1]["BAMMIN"].ToString();
                        MOCTADAILY.B2AMMAX = ds1.Tables["ds1"].Rows[1]["BAMMAX"].ToString();
                        MOCTADAILY.B2BMAVG = ds1.Tables["ds1"].Rows[1]["BBMAVG"].ToString();
                        MOCTADAILY.B2BMMIN = ds1.Tables["ds1"].Rows[1]["BBMMIN"].ToString();
                        MOCTADAILY.B2BMMAX = ds1.Tables["ds1"].Rows[1]["BBMMAX"].ToString();
                        MOCTADAILY.B2ALAVG = ds1.Tables["ds1"].Rows[1]["BALAVG"].ToString();
                        MOCTADAILY.B2ALMIN = ds1.Tables["ds1"].Rows[1]["BALMIN"].ToString();
                        MOCTADAILY.B2ALMAX = ds1.Tables["ds1"].Rows[1]["BALMAX"].ToString();
                        MOCTADAILY.B2BLAVG = ds1.Tables["ds1"].Rows[1]["BBLAVG"].ToString();
                        MOCTADAILY.B2BLMIN = ds1.Tables["ds1"].Rows[1]["BBLMIN"].ToString();
                        MOCTADAILY.B2BLMAX = ds1.Tables["ds1"].Rows[1]["BBLMAX"].ToString();

                        MOCTADAILY.B3ARAVG = ds1.Tables["ds1"].Rows[2]["BARAVG"].ToString();
                        MOCTADAILY.B3ARMIN = ds1.Tables["ds1"].Rows[2]["BARMIN"].ToString();
                        MOCTADAILY.B3ARMAX = ds1.Tables["ds1"].Rows[2]["BARMAX"].ToString();
                        MOCTADAILY.B3BRAVG = ds1.Tables["ds1"].Rows[2]["BBRAVG"].ToString();
                        MOCTADAILY.B3BRMIN = ds1.Tables["ds1"].Rows[2]["BBRMIN"].ToString();
                        MOCTADAILY.B3BRMAX = ds1.Tables["ds1"].Rows[2]["BBRMAX"].ToString();
                        MOCTADAILY.B3AMAVG = ds1.Tables["ds1"].Rows[2]["BAMAVG"].ToString();
                        MOCTADAILY.B3AMMIN = ds1.Tables["ds1"].Rows[2]["BAMMIN"].ToString();
                        MOCTADAILY.B3AMMAX = ds1.Tables["ds1"].Rows[2]["BAMMAX"].ToString();
                        MOCTADAILY.B3BMAVG = ds1.Tables["ds1"].Rows[2]["BBMAVG"].ToString();
                        MOCTADAILY.B3BMMIN = ds1.Tables["ds1"].Rows[2]["BBMMIN"].ToString();
                        MOCTADAILY.B3BMMAX = ds1.Tables["ds1"].Rows[2]["BBMMAX"].ToString();
                        MOCTADAILY.B3ALAVG = ds1.Tables["ds1"].Rows[2]["BALAVG"].ToString();
                        MOCTADAILY.B3ALMIN = ds1.Tables["ds1"].Rows[2]["BALMIN"].ToString();
                        MOCTADAILY.B3ALMAX = ds1.Tables["ds1"].Rows[2]["BALMAX"].ToString();
                        MOCTADAILY.B3BLAVG = ds1.Tables["ds1"].Rows[2]["BBLAVG"].ToString();
                        MOCTADAILY.B3BLMIN = ds1.Tables["ds1"].Rows[2]["BBLMIN"].ToString();
                        MOCTADAILY.B3BLMAX = ds1.Tables["ds1"].Rows[2]["BBLMAX"].ToString();

                        MOCTADAILY.B4ARAVG = ds1.Tables["ds1"].Rows[3]["BARAVG"].ToString();
                        MOCTADAILY.B4ARMIN = ds1.Tables["ds1"].Rows[3]["BARMIN"].ToString();
                        MOCTADAILY.B4ARMAX = ds1.Tables["ds1"].Rows[3]["BARMAX"].ToString();
                        MOCTADAILY.B4BRAVG = ds1.Tables["ds1"].Rows[3]["BBRAVG"].ToString();
                        MOCTADAILY.B4BRMIN = ds1.Tables["ds1"].Rows[3]["BBRMIN"].ToString();
                        MOCTADAILY.B4BRMAX = ds1.Tables["ds1"].Rows[3]["BBRMAX"].ToString();
                        MOCTADAILY.B4AMAVG = ds1.Tables["ds1"].Rows[3]["BAMAVG"].ToString();
                        MOCTADAILY.B4AMMIN = ds1.Tables["ds1"].Rows[3]["BAMMIN"].ToString();
                        MOCTADAILY.B4AMMAX = ds1.Tables["ds1"].Rows[3]["BAMMAX"].ToString();
                        MOCTADAILY.B4BMAVG = ds1.Tables["ds1"].Rows[3]["BBMAVG"].ToString();
                        MOCTADAILY.B4BMMIN = ds1.Tables["ds1"].Rows[3]["BBMMIN"].ToString();
                        MOCTADAILY.B4BMMAX = ds1.Tables["ds1"].Rows[3]["BBMMAX"].ToString();
                        MOCTADAILY.B4ALAVG = ds1.Tables["ds1"].Rows[3]["BALAVG"].ToString();
                        MOCTADAILY.B4ALMIN = ds1.Tables["ds1"].Rows[3]["BALMIN"].ToString();
                        MOCTADAILY.B4ALMAX = ds1.Tables["ds1"].Rows[3]["BALMAX"].ToString();
                        MOCTADAILY.B4BLAVG = ds1.Tables["ds1"].Rows[3]["BBLAVG"].ToString();
                        MOCTADAILY.B4BLMIN = ds1.Tables["ds1"].Rows[3]["BBLMIN"].ToString();
                        MOCTADAILY.B4BLMAX = ds1.Tables["ds1"].Rows[3]["BBLMAX"].ToString();

                        MOCTADAILY.B5ARAVG = ds1.Tables["ds1"].Rows[4]["BARAVG"].ToString();
                        MOCTADAILY.B5ARMIN = ds1.Tables["ds1"].Rows[4]["BARMIN"].ToString();
                        MOCTADAILY.B5ARMAX = ds1.Tables["ds1"].Rows[4]["BARMAX"].ToString();
                        MOCTADAILY.B5BRAVG = ds1.Tables["ds1"].Rows[4]["BBRAVG"].ToString();
                        MOCTADAILY.B5BRMIN = ds1.Tables["ds1"].Rows[4]["BBRMIN"].ToString();
                        MOCTADAILY.B5BRMAX = ds1.Tables["ds1"].Rows[4]["BBRMAX"].ToString();
                        MOCTADAILY.B5AMAVG = ds1.Tables["ds1"].Rows[4]["BAMAVG"].ToString();
                        MOCTADAILY.B5AMMIN = ds1.Tables["ds1"].Rows[4]["BAMMIN"].ToString();
                        MOCTADAILY.B5AMMAX = ds1.Tables["ds1"].Rows[4]["BAMMAX"].ToString();
                        MOCTADAILY.B5BMAVG = ds1.Tables["ds1"].Rows[4]["BBMAVG"].ToString();
                        MOCTADAILY.B5BMMIN = ds1.Tables["ds1"].Rows[4]["BBMMIN"].ToString();
                        MOCTADAILY.B5BMMAX = ds1.Tables["ds1"].Rows[4]["BBMMAX"].ToString();
                        MOCTADAILY.B5ALAVG = ds1.Tables["ds1"].Rows[4]["BALAVG"].ToString();
                        MOCTADAILY.B5ALMIN = ds1.Tables["ds1"].Rows[4]["BALMIN"].ToString();
                        MOCTADAILY.B5ALMAX = ds1.Tables["ds1"].Rows[4]["BALMAX"].ToString();
                        MOCTADAILY.B5BLAVG = ds1.Tables["ds1"].Rows[4]["BBLAVG"].ToString();
                        MOCTADAILY.B5BLMIN = ds1.Tables["ds1"].Rows[4]["BBLMIN"].ToString();
                        MOCTADAILY.B5BLMAX = ds1.Tables["ds1"].Rows[4]["BBLMAX"].ToString();

                    }
                }

            }
            catch
            {

            }
            finally
            {
                sqlConn.Close();
            }

            return MOCTADAILY;
        }


        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            SEARCH(dateTimePicker1.Value.ToString("yyyyMMdd"));
        }
        private void button2_Click(object sender, EventArgs e)
        {
            STATUS = "ADD";
            label26.Text = "ADD";
            SETTEXTBOX1();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            STATUS = "EDIT";
            label26.Text = "EDIT";
        }

      
        private void button4_Click(object sender, EventArgs e)
        {
            if (STATUS.Equals("ADD"))
            {
                string ID = null;
                ID = Guid.NewGuid().ToString();

                ADDMOCTADAILY(ID, textBox1.Text.Trim(), textBox2.Text.Trim(), comboBox1.Text.ToString().Trim(), textBox111.Text.Trim(), textBox112.Text.Trim(), textBox113.Text.Trim(), textBox121.Text.Trim(), textBox122.Text.Trim(), textBox123.Text.Trim(), dateTimePicker2.Value.ToString("yyyy-MM-dd HH:mm:ss"), dateTimePicker3.Value.ToString("yyyy-MM-dd HH:mm:ss")
                    ,textBox211.Text.Trim()
                    ,textBox231.Text.Trim(), textBox232.Text.Trim(), textBox233.Text.Trim(), textBox234.Text.Trim(), textBox235.Text.Trim(), textBox236.Text.Trim()
                    ,textBox241.Text.Trim(), textBox242.Text.Trim(), textBox243.Text.Trim(), textBox244.Text.Trim(), textBox245.Text.Trim(), textBox246.Text.Trim()
                    ,textBox251.Text.Trim(), textBox252.Text.Trim(), textBox253.Text.Trim(), textBox254.Text.Trim(), textBox255.Text.Trim(), textBox256.Text.Trim()
                    ,textBox261.Text.Trim(), textBox262.Text.Trim(), textBox263.Text.Trim(), textBox264.Text.Trim(), textBox265.Text.Trim(), textBox266.Text.Trim()
                    );

                UPDATEMOCTADAILYDETAIL(ID, dateTimePicker2.Value.ToString("yyyy-MM-dd HH:mm:ss"), dateTimePicker3.Value.ToString("yyyy-MM-dd HH:mm:ss"));
            }
            else if (STATUS.Equals("EDIT"))
            {
                UPDATEMOCTADAILY(ID,textBox122.Text.Trim(),textBox123.Text.Trim(),dateTimePicker2.Value.ToString("yyyy-MM-dd HH:mm:ss"), dateTimePicker3.Value.ToString("yyyy-MM-dd HH:mm:ss")
                    , textBox211.Text.Trim()
                    , textBox231.Text.Trim(), textBox232.Text.Trim(), textBox233.Text.Trim(), textBox234.Text.Trim(), textBox235.Text.Trim(), textBox236.Text.Trim()
                    , textBox241.Text.Trim(), textBox242.Text.Trim(), textBox243.Text.Trim(), textBox244.Text.Trim(), textBox245.Text.Trim(), textBox246.Text.Trim()
                    , textBox251.Text.Trim(), textBox252.Text.Trim(), textBox253.Text.Trim(), textBox254.Text.Trim(), textBox255.Text.Trim(), textBox256.Text.Trim()
                    , textBox261.Text.Trim(), textBox262.Text.Trim(), textBox263.Text.Trim(), textBox264.Text.Trim(), textBox265.Text.Trim(), textBox266.Text.Trim()
                    );

                UPDATEMOCTADAILYDETAIL(textBox151.Text.Trim(),dateTimePicker2.Value.ToString("yyyy-MM-dd HH:mm:ss"), dateTimePicker3.Value.ToString("yyyy-MM-dd HH:mm:ss"));

            }

            SETTEXTBOX2();

            STATUS = null;
            label26.Text = "STATUS";
            SEARCH(dateTimePicker1.Value.ToString("yyyyMMdd"));
        }
        private void button5_Click(object sender, EventArgs e)
        {

            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELMOCTADAILY(ID);

                SEARCH(dateTimePicker1.Value.ToString("yyyyMMdd"));
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }


        #endregion


    }
}
