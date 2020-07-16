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
                sbSql.AppendFormat(@"  ,[B3AMAVG] AS '小線-3段-上右爐-平均',[B3AMMIN] AS '小線-3段-上中爐-最小',[B3AMMAX] AS '小線-3段-上中爐-最大'");
                sbSql.AppendFormat(@"  ,[B3BMAVG] AS '小線-3段-下中爐-平均',[B3BMMIN] AS '小線-3段-下中爐-最小',[B3BMMAX] AS '小線-3段-下中爐-最大'");
                sbSql.AppendFormat(@"  ,[B3ALAVG] AS '小線-3段-上中爐-平均',[B3ALMIN] AS '小線-3段-上左爐-最小',[B3ALMAX] AS '小線-3段-上左爐-最大'");
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

            }
        }

        public void UPDATEMOCTADAILYDETAIL(string TA001,string TA002,string TA021,string SDATES,string EDATES)
        {
            DATAMOCTADAILY MOCTADAILY = new DATAMOCTADAILY();
            MOCTADAILY = CALTEMPERHUMI(MOCTADAILY,SDATES,EDATES);
            MOCTADAILY = CALTEMPERAVG(MOCTADAILY, SDATES, EDATES);

            MessageBox.Show(MOCTADAILY.TEMPERAVG+" "+ MOCTADAILY.B1ALAVG);
        }

        public DATAMOCTADAILY CALTEMPERHUMI(DATAMOCTADAILY MOCTADAILY, string SDATES, string EDATES)
        {
            MOCTADAILY.TEMPERAVG = "30.6";
            return MOCTADAILY;
        }

        public DATAMOCTADAILY CALTEMPERAVG(DATAMOCTADAILY MOCTADAILY, string SDATES, string EDATES)
        {
            MOCTADAILY.B1ALAVG = "250";
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
                ADDMOCTADAILY(Guid.NewGuid().ToString(), textBox1.Text.Trim(), textBox2.Text.Trim(), comboBox1.Text.ToString().Trim(), textBox111.Text.Trim(), textBox112.Text.Trim(), textBox113.Text.Trim(), textBox121.Text.Trim(), textBox122.Text.Trim(), textBox123.Text.Trim(), dateTimePicker2.Value.ToString("yyyy-MM-dd HH:mm:ss"), dateTimePicker3.Value.ToString("yyyy-MM-dd HH:mm:ss")
                    ,textBox211.Text.Trim()
                    ,textBox231.Text.Trim(), textBox232.Text.Trim(), textBox233.Text.Trim(), textBox234.Text.Trim(), textBox235.Text.Trim(), textBox236.Text.Trim()
                    ,textBox241.Text.Trim(), textBox242.Text.Trim(), textBox243.Text.Trim(), textBox244.Text.Trim(), textBox245.Text.Trim(), textBox246.Text.Trim()
                    ,textBox251.Text.Trim(), textBox252.Text.Trim(), textBox253.Text.Trim(), textBox254.Text.Trim(), textBox255.Text.Trim(), textBox256.Text.Trim()
                    ,textBox261.Text.Trim(), textBox262.Text.Trim(), textBox263.Text.Trim(), textBox264.Text.Trim(), textBox265.Text.Trim(), textBox266.Text.Trim()
                    );
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

                UPDATEMOCTADAILYDETAIL(textBox1.Text.Trim(),textBox2.Text.Trim(),comboBox1.Text.Trim(),dateTimePicker2.Value.ToString("yyyy-MM-dd HH:mm:ss"), dateTimePicker3.Value.ToString("yyyy-MM-dd HH:mm:ss"));

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
