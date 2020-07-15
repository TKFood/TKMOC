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

        public frmMOCTADAILY()
        {
            InitializeComponent();

            comboBox1load();
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
        }

        public void SETTEXTBOX2()
        {
            textBox1.ReadOnly = true;
            textBox2.ReadOnly = true;

            textBox1.Text = null;
            textBox2.Text = null;
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

        public void  ADDMOCTADAILY(string TA001, string TA002, string TA021, string MB001, string MB002, string MB003, string TA017, string NUM, string NGNUM, string SDATES, string EDATES)
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
                sbSql.AppendFormat("  ([TA001],[TA002],[TA021],[MB001],[MB002],[MB003],[TA017],[NUM],[NGNUM],[SDATES],[EDATES]) ");
                sbSql.AppendFormat("  VALUES");
                sbSql.AppendFormat("  ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}')",TA001,TA002,TA021,MB001,MB002,MB003,TA017,NUM,NGNUM,SDATES,EDATES);
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

        public void UPDATEMOCTADAILY()
        {

        }
        #endregion

        private void button4_Click(object sender, EventArgs e)
        {
            if (STATUS.Equals("ADD"))
            {
                ADDMOCTADAILY(textBox1.Text.Trim(), textBox2.Text.Trim(), comboBox1.Text.ToString().Trim(), textBox111.Text.Trim(), textBox112.Text.Trim(), textBox113.Text.Trim(), textBox121.Text.Trim(), textBox122.Text.Trim(), textBox123.Text.Trim(), dateTimePicker2.Value.ToString("yyyy/MM/dd HH:mm:ss"), dateTimePicker3.Value.ToString("yyyy/MM/dd HH:mm:ss"));
            }
            else if (STATUS.Equals("EDIT"))
            {
               
            }

            SETTEXTBOX2();

            STATUS = null;
            label26.Text = "STATUS";
            SEARCH(dateTimePicker1.Value.ToString("yyyyMMdd"));
        }
    }
}
