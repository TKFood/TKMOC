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
    public partial class MOCDAILYWORKHRS : Form
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
     

        public MOCDAILYWORKHRS()
        {
            InitializeComponent();

            SETTIMES();
            comboBox1load();

        }
        public class DATACSTMB
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
            public string sync_date;
            public string sync_time;
            public string sync_mark;
            public string sync_count;
            public string DataUser;
            public string DataGroup;
            public string MB001;
            public string MB002;
            public string MB003;
            public string MB004;
            public string MB005;
            public string MB006;
            public string MB007;
            public string MB008;
            public string MB009;
            public string MB010;
            public string MB011;
            public string MB012;
            public string MB013;
            public string MB014;
            public string MB015;
            public string MB016;
            public string MB017;
            public string MB018;
            public string MB019;
            public string MB020;
            public string MB021;
            public string MB022;
            public string UDF01;
            public string UDF02;
            public string UDF03;
            public string UDF04;
            public string UDF05;
            public string UDF06;
            public string UDF07;
            public string UDF08;
            public string UDF09;
            public string UDF10;

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


                sbSql.AppendFormat(@"  SELECT  CONVERT(NVARCHAR,[DATS],112) AS '日期',[MANU] AS '產線別',[TA001] AS '製令單',[TA002] AS '製令單號',[MB001] AS '品號',[MB002] AS '品名',[NUMS] AS '入庫量',[MOCNUM] AS '預計生產量'");
                sbSql.AppendFormat(@"  ,CONVERT(NVARCHAR,[WORKSTART],114) AS '開始時間',CONVERT(NVARCHAR,[WORKEND],114) AS '結束時間',[WORKHRS] AS '工時',[WORKTIMES] AS '工時(分)',[AVGWORKHRS] AS '平均工時'");
                sbSql.AppendFormat(@"  ,[WATERNOODLESEMP] AS '水麵攪拌',CONVERT(NVARCHAR,[WATERNOODLESSTART],114) AS '水麵攪拌開始時間',CONVERT(NVARCHAR,[WATERNOODLESEND],114) AS '水麵攪拌結束時間',[WATERNOODLESTIMES] AS '水麵攪拌工時'");
                sbSql.AppendFormat(@"  ,[OILPASTRYEMP] AS '油酥攪拌',CONVERT(NVARCHAR,[OILPASTRYSTART],114) AS '油酥攪拌開始時間',CONVERT(NVARCHAR,[OILPASTRYEND],114) AS '油酥攪拌結束時間',[OILPASTRYTIMES] AS '油酥攪拌工時'");
                sbSql.AppendFormat(@"  ,[FOLDEMP] AS '摺疊',CONVERT(NVARCHAR,[FOLDSTART],114) AS '摺疊開始時間',CONVERT(NVARCHAR,[FOLDEND],114) AS '摺疊結束時間',[FOLDTIMES] AS '摺疊工時'");
                sbSql.AppendFormat(@"  ,[TYPECOOKEMP] AS '舖餅',CONVERT(NVARCHAR,[TYPECOOKSTART],114) AS '舖餅開始時間',CONVERT(NVARCHAR,[TYPECOOKEND],114) AS '舖餅結束時間',[TYPECOOKTIMES] AS '舖餅工時'");
                sbSql.AppendFormat(@"  ,[TYPEEMP] AS '成型/烘烤',CONVERT(NVARCHAR,[TYPESTART],114) AS '成型/烘烤開始時間',CONVERT(NVARCHAR,[TYPEEND],114) AS '成型/烘烤結束時間',[TYPETIMES] AS '成型/烘烤工時'");
                sbSql.AppendFormat(@"  ,[OVENCOOKEMP] AS '烤箱篩餅',CONVERT(NVARCHAR,[OVENCOOKSTART],114) AS '烤箱篩餅開始時間',CONVERT(NVARCHAR,[OVENCOOKEND],114) AS '烤箱篩餅結束時間',[OVENCOOKTIMES] AS '烤箱篩餅工時'");
                sbSql.AppendFormat(@"  ,[COLDCOOKEMP] AS '冷卻篩餅',CONVERT(NVARCHAR,[COLDCOOKSTART],114) AS '冷卻篩餅開始時間',CONVERT(NVARCHAR,[COLDCOOKEND],114) AS '冷卻篩餅結束時間',[COLDCOOKTIMES] AS '冷卻篩餅工時'");
                sbSql.AppendFormat(@"  ,[ARRAYEMP] AS '排餅/裝罐',CONVERT(NVARCHAR,[ARRAYSTART],114) AS '排餅/裝罐開始時間',CONVERT(NVARCHAR,[ARRAYEND],114) AS '排餅/裝罐結束時間',[ARRAYTIMES] AS '排餅/裝罐工時'");
                sbSql.AppendFormat(@"  ,[PACKEMP] AS '包裝機',CONVERT(NVARCHAR,[PACKSTART],114) AS '包裝機開始時間',CONVERT(NVARCHAR,[PACKEND],114) AS '包裝機結束時間',[PACKTIMES] AS '包裝機工時'");
                sbSql.AppendFormat(@"  ,[PACKPICKEMP] AS '包裝篩餅',CONVERT(NVARCHAR,[PACKPICKSTART],114) AS '包裝篩餅開始時間',CONVERT(NVARCHAR,[PACKPICKEND],114) AS '包裝篩餅結束時間',[PACKPICKTIMES] AS '包裝篩餅工時'");
                sbSql.AppendFormat(@"  ,[BOXSEMP] AS '裝箱',CONVERT(NVARCHAR,[BOXSSTART],114) AS '裝箱開始時間',CONVERT(NVARCHAR,[BOXSEND],114) AS '裝箱結束時間',[BOXSTIMES] AS '裝箱工時'");
                sbSql.AppendFormat(@"  ,[HANDCOOKEMP] AS '撿餅',CONVERT(NVARCHAR,[HANDCOOKSTART],114) AS '撿餅開始時間',CONVERT(NVARCHAR,[HANDCOOKEND],114) AS '撿餅結束時間',[HANDCOOKTIMES] AS '撿餅工時'");
                sbSql.AppendFormat(@"  ,[SCALESWEIGHTEMP] AS '秤重',CONVERT(NVARCHAR,[SCALESWEIGHTSTART],114) AS '秤重開始時間',CONVERT(NVARCHAR,[SCALESWEIGHTEND],114) AS '秤重結束時間',[SCALESWEIGHTTIMES] AS '秤重工時'");
                sbSql.AppendFormat(@"  ,[OUTBOXSEMP] AS '外裝箱',CONVERT(NVARCHAR,[OUTBOXSSTART],114) AS '外裝箱開始時間',CONVERT(NVARCHAR,[OUTBOXSEND],114) AS '外裝箱結束時間',[OUTBOXSTIMES] AS '外裝箱工時'");
                sbSql.AppendFormat(@"  ,[SEALEMP] AS '封箱',CONVERT(NVARCHAR,[SEALSTART],114) AS '封箱開始時間',CONVERT(NVARCHAR,[SEALEND],114) AS '封箱結束時間',[SEALTIMES] AS '封箱工時'");
                sbSql.AppendFormat(@"  ,[THROWEMP] AS '倒餅',CONVERT(NVARCHAR,[THROWSTART],114) AS '倒餅開始時間',CONVERT(NVARCHAR,[THROWEND],114) AS '倒餅結束時間',[THROWTIMES] AS '倒餅工時'");
                sbSql.AppendFormat(@"  ,[BOXPACKEMP] AS '封盒機',CONVERT(NVARCHAR,[BOXPACKSTART],114) AS '封盒機開始時間',CONVERT(NVARCHAR,[BOXPACKEND],114) AS '封盒機結束時間',[BOXPACKTIMES] AS '封盒機工時'");
                sbSql.AppendFormat(@"  ,[ID]");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCDAILYWORKHRS] ");
                sbSql.AppendFormat(@"  WHERE  CONVERT(NVARCHAR,[DATS],112)='{0}'", IDDATE);
                sbSql.AppendFormat(@"  ORDER BY [TA001],[TA002]");
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
        private void textBox11_TextChanged(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(textBox11.Text)&& !string.IsNullOrEmpty(textBox12.Text))
            {
                SEARCHMOCTA(textBox11.Text.Trim(), textBox12.Text.Trim());
            }
        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox11.Text) && !string.IsNullOrEmpty(textBox12.Text))
            {
                SEARCHMOCTA(textBox11.Text.Trim(), textBox12.Text.Trim());
            }
        }

        public void SEARCHMOCTA(string TA001,string TA002)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();

            DataSet ds1 = new DataSet();

            SETTEXT1();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@" SELECT TA006,TA034,TA015,TA017 FROM [TK].dbo.MOCTA WHERE TA001='{0}' AND TA002='{1}' ", TA001,TA002);
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
                        textBox21.Text = ds1.Tables["ds1"].Rows[0]["TA006"].ToString();
                        textBox22.Text = ds1.Tables["ds1"].Rows[0]["TA034"].ToString();
                        textBox23.Text = ds1.Tables["ds1"].Rows[0]["TA017"].ToString();
                        textBox24.Text = ds1.Tables["ds1"].Rows[0]["TA015"].ToString();

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

        public void ADDMOCDAILYWORKHRS(string ID, string DATS, string MANU, string TA001, string TA002, string MB001, string MB002, string NUMS, string MOCNUM
            , string WORKSTART, string WORKEND, string WORKHRS, string WORKTIMES, string AVGWORKHRS
            , string WATERNOODLESEMP, string WATERNOODLESSTART, string WATERNOODLESEND, string WATERNOODLESTIMES
            , string OILPASTRYEMP, string OILPASTRYSTART, string OILPASTRYEND, string OILPASTRYTIMES
            , string FOLDEMP, string FOLDSTART, string FOLDEND, string FOLDTIMES
            , string TYPECOOKEMP, string TYPECOOKSTART, string TYPECOOKEND, string TYPECOOKTIMES
            , string TYPEEMP, string TYPESTART, string TYPEEND, string TYPETIMES
            , string OVENCOOKEMP, string OVENCOOKSTART, string OVENCOOKEND, string OVENCOOKTIMES
            , string COLDCOOKEMP, string COLDCOOKSTART, string COLDCOOKEND, string COLDCOOKTIMES
            , string ARRAYEMP, string ARRAYSTART, string ARRAYEND, string ARRAYTIMES
            , string PACKEMP, string PACKSTART, string PACKEND, string PACKTIMES
            , string PACKPICKEMP, string PACKPICKSTART, string PACKPICKEND, string PACKPICKTIMES
            , string BOXSEMP, string BOXSSTART, string BOXSEND, string BOXSTIMES
            , string HANDCOOKEMP, string HANDCOOKSTART, string HANDCOOKEND, string HANDCOOKTIMES
            , string SCALESWEIGHTEMP, string SCALESWEIGHTSTART, string SCALESWEIGHTEND, string SCALESWEIGHTTIMES
            , string OUTBOXSEMP, string OUTBOXSSTART, string OUTBOXSEND, string OUTBOXSTIMES
            , string SEALEMP, string SEALSTART, string SEALEND, string SEALTIMES
            , string THROWEMP, string THROWSTART, string THROWEND, string THROWTIMES
            , string BOXPACKEMP, string BOXPACKSTART, string BOXPACKEND, string BOXPACKTIMES                            
            )
        {
            try
            {
                //add ZWAREWHOUSEPURTH
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat("  INSERT INTO [TKMOC].[dbo].[MOCDAILYWORKHRS]");
                sbSql.AppendFormat("  (");
                sbSql.AppendFormat("  [ID],[DATS],[MANU],[TA001],[TA002],[MB001],[MB002],[NUMS],[MOCNUM]");
                sbSql.AppendFormat("  ,[WORKSTART],[WORKEND],[WORKHRS],[WORKTIMES],[AVGWORKHRS]");
                sbSql.AppendFormat("  ,[WATERNOODLESEMP],[WATERNOODLESSTART],[WATERNOODLESEND],[WATERNOODLESTIMES]");
                sbSql.AppendFormat("  ,[OILPASTRYEMP],[OILPASTRYSTART],[OILPASTRYEND],[OILPASTRYTIMES]");
                sbSql.AppendFormat("  ,[FOLDEMP],[FOLDSTART],[FOLDEND],[FOLDTIMES]");
                sbSql.AppendFormat("  ,[TYPECOOKEMP],[TYPECOOKSTART],[TYPECOOKEND],[TYPECOOKTIMES]");
                sbSql.AppendFormat("  ,[TYPEEMP],[TYPESTART],[TYPEEND],[TYPETIMES]");
                sbSql.AppendFormat("  ,[OVENCOOKEMP],[OVENCOOKSTART],[OVENCOOKEND],[OVENCOOKTIMES]");
                sbSql.AppendFormat("  ,[COLDCOOKEMP],[COLDCOOKSTART],[COLDCOOKEND],[COLDCOOKTIMES]");
                sbSql.AppendFormat("  ,[ARRAYEMP],[ARRAYSTART],[ARRAYEND],[ARRAYTIMES]");
                sbSql.AppendFormat("  ,[PACKEMP],[PACKSTART],[PACKEND],[PACKTIMES]");
                sbSql.AppendFormat("  ,[PACKPICKEMP],[PACKPICKSTART],[PACKPICKEND],[PACKPICKTIMES]");
                sbSql.AppendFormat("  ,[BOXSEMP],[BOXSSTART],[BOXSEND],[BOXSTIMES]");
                sbSql.AppendFormat("  ,[HANDCOOKEMP],[HANDCOOKSTART],[HANDCOOKEND],[HANDCOOKTIMES]");
                sbSql.AppendFormat("  ,[SCALESWEIGHTEMP],[SCALESWEIGHTSTART],[SCALESWEIGHTEND],[SCALESWEIGHTTIMES]");
                sbSql.AppendFormat("  ,[OUTBOXSEMP],[OUTBOXSSTART],[OUTBOXSEND],[OUTBOXSTIMES]");
                sbSql.AppendFormat("  ,[SEALEMP],[SEALSTART],[SEALEND],[SEALTIMES]");
                sbSql.AppendFormat("  ,[THROWEMP],[THROWSTART],[THROWEND],[THROWTIMES]");
                sbSql.AppendFormat("  ,[BOXPACKEMP],[BOXPACKSTART],[BOXPACKEND],[BOXPACKTIMES]");
                sbSql.AppendFormat("  )");
                sbSql.AppendFormat("  VALUES");
                sbSql.AppendFormat("  (");
                sbSql.AppendFormat("  '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}',", ID, DATS, MANU, TA001, TA002, MB001, MB002, NUMS, MOCNUM);
                sbSql.AppendFormat("  '{0}','{1}','{2}','{3}','{4}',", WORKSTART, WORKEND, WORKHRS, WORKTIMES, AVGWORKHRS);
                sbSql.AppendFormat("  '{0}','{1}','{2}','{3}',", WATERNOODLESEMP, WATERNOODLESSTART, WATERNOODLESEND, WATERNOODLESTIMES);
                sbSql.AppendFormat("  '{0}','{1}','{2}','{3}',", OILPASTRYEMP, OILPASTRYSTART, OILPASTRYEND, OILPASTRYTIMES);
                sbSql.AppendFormat("  '{0}','{1}','{2}','{3}',", FOLDEMP, FOLDSTART, FOLDEND, FOLDTIMES);
                sbSql.AppendFormat("  '{0}','{1}','{2}','{3}',", TYPECOOKEMP, TYPECOOKSTART, TYPECOOKEND, TYPECOOKTIMES);
                sbSql.AppendFormat("  '{0}','{1}','{2}','{3}',", TYPEEMP, TYPESTART, TYPEEND, TYPETIMES);
                sbSql.AppendFormat("  '{0}','{1}','{2}','{3}',", OVENCOOKEMP, OVENCOOKSTART, OVENCOOKEND, OVENCOOKTIMES);
                sbSql.AppendFormat("  '{0}','{1}','{2}','{3}',", COLDCOOKEMP, COLDCOOKSTART, COLDCOOKEND, COLDCOOKTIMES);
                sbSql.AppendFormat("  '{0}','{1}','{2}','{3}',", ARRAYEMP, ARRAYSTART, ARRAYEND, ARRAYTIMES);
                sbSql.AppendFormat("  '{0}','{1}','{2}','{3}',", PACKEMP, PACKSTART, PACKEND, PACKTIMES);
                sbSql.AppendFormat("  '{0}','{1}','{2}','{3}',", PACKPICKEMP, PACKPICKSTART, PACKPICKEND, PACKPICKTIMES);
                sbSql.AppendFormat("  '{0}','{1}','{2}','{3}',", BOXSEMP, BOXSSTART, BOXSEND, BOXSTIMES);
                sbSql.AppendFormat("  '{0}','{1}','{2}','{3}',", HANDCOOKEMP, HANDCOOKSTART, HANDCOOKEND, HANDCOOKTIMES);
                sbSql.AppendFormat("  '{0}','{1}','{2}','{3}',", SCALESWEIGHTEMP, SCALESWEIGHTSTART, SCALESWEIGHTEND, SCALESWEIGHTTIMES);
                sbSql.AppendFormat("  '{0}','{1}','{2}','{3}',", OUTBOXSEMP, OUTBOXSSTART, OUTBOXSEND, OUTBOXSTIMES);
                sbSql.AppendFormat("  '{0}','{1}','{2}','{3}',", SEALEMP, SEALSTART, SEALEND, SEALTIMES);
                sbSql.AppendFormat("  '{0}','{1}','{2}','{3}',", THROWEMP, THROWSTART, THROWEND, THROWTIMES);
                sbSql.AppendFormat("  '{0}','{1}','{2}','{3}'", BOXPACKEMP, BOXPACKSTART, BOXPACKEND, BOXPACKTIMES);
                sbSql.AppendFormat("  )");
                sbSql.AppendFormat("  ");
                sbSql.AppendFormat("  ");
                sbSql.AppendFormat("  ");
                sbSql.AppendFormat("  ");
                sbSql.AppendFormat("  ");
                sbSql.AppendFormat("  ");
                sbSql.AppendFormat("  ");
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

        public void UPDATEMOCDAILYWORKHRS(
            string ID, string DATS, string MANU, string TA001, string TA002, string MB001, string MB002, string NUMS, string MOCNUM
            , string WORKSTART, string WORKEND, string WORKHRS, string WORKTIMES, string AVGWORKHRS
            , string WATERNOODLESEMP, string WATERNOODLESSTART, string WATERNOODLESEND, string WATERNOODLESTIMES
            , string OILPASTRYEMP, string OILPASTRYSTART, string OILPASTRYEND, string OILPASTRYTIMES
            , string FOLDEMP, string FOLDSTART, string FOLDEND, string FOLDTIMES
            , string TYPECOOKEMP, string TYPECOOKSTART, string TYPECOOKEND, string TYPECOOKTIMES
            , string TYPEEMP, string TYPESTART, string TYPEEND, string TYPETIMES
            , string OVENCOOKEMP, string OVENCOOKSTART, string OVENCOOKEND, string OVENCOOKTIMES
            , string COLDCOOKEMP, string COLDCOOKSTART, string COLDCOOKEND, string COLDCOOKTIMES
            , string ARRAYEMP, string ARRAYSTART, string ARRAYEND, string ARRAYTIMES
            , string PACKEMP, string PACKSTART, string PACKEND, string PACKTIMES
            , string PACKPICKEMP, string PACKPICKSTART, string PACKPICKEND, string PACKPICKTIMES
            , string BOXSEMP, string BOXSSTART, string BOXSEND, string BOXSTIMES
            , string HANDCOOKEMP, string HANDCOOKSTART, string HANDCOOKEND, string HANDCOOKTIMES
            , string SCALESWEIGHTEMP, string SCALESWEIGHTSTART, string SCALESWEIGHTEND, string SCALESWEIGHTTIMES
            , string OUTBOXSEMP, string OUTBOXSSTART, string OUTBOXSEND, string OUTBOXSTIMES
            , string SEALEMP, string SEALSTART, string SEALEND, string SEALTIMES
            , string THROWEMP, string THROWSTART, string THROWEND, string THROWTIMES
            , string BOXPACKEMP, string BOXPACKSTART, string BOXPACKEND, string BOXPACKTIMES
            )
        {
            try
            {
                //add ZWAREWHOUSEPURTH
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat("  UPDATE [TKMOC].[dbo].[MOCDAILYWORKHRS]");
                sbSql.AppendFormat("  SET [DATS]='{0}',[MANU]='{1}',[TA001]='{2}',[TA002]='{3}',[MB001]='{4}',[MB002]='{5}',[NUMS]='{6}',[MOCNUM]='{7}'", DATS, MANU, TA001, TA002, MB001, MB002, NUMS, MOCNUM);
                sbSql.AppendFormat("  ,[WORKSTART]='{0}',[WORKEND]='{1}',[WORKHRS]='{2}',[WORKTIMES]='{3}',[AVGWORKHRS]='{4}'", WORKSTART, WORKEND, WORKHRS, WORKTIMES, AVGWORKHRS);
                sbSql.AppendFormat("  ,[WATERNOODLESEMP]='{0}',[WATERNOODLESSTART]='{1}',[WATERNOODLESEND]='{2}',[WATERNOODLESTIMES]='{3}'", WATERNOODLESEMP, WATERNOODLESSTART, WATERNOODLESEND, WATERNOODLESTIMES);
                sbSql.AppendFormat("  ,[OILPASTRYEMP]='{0}',[OILPASTRYSTART]='{1}',[OILPASTRYEND]='{2}',[OILPASTRYTIMES]='{3}'", OILPASTRYEMP, OILPASTRYSTART, OILPASTRYEND, OILPASTRYTIMES);
                sbSql.AppendFormat("  ,[FOLDEMP]='{0}',[FOLDSTART]='{1}',[FOLDEND]='{2}',[FOLDTIMES]='{3}'", FOLDEMP, FOLDSTART, FOLDEND, FOLDTIMES);
                sbSql.AppendFormat("  ,[TYPECOOKEMP]='{0}',[TYPECOOKSTART]='{1}',[TYPECOOKEND]='{2}',[TYPECOOKTIMES]='{3}'", TYPECOOKEMP, TYPECOOKSTART, TYPECOOKEND, TYPECOOKTIMES);
                sbSql.AppendFormat("  ,[TYPEEMP]='{0}',[TYPESTART]='{1}',[TYPEEND]='{2}',[TYPETIMES]='{3}'", TYPEEMP, TYPESTART, TYPEEND, TYPETIMES);
                sbSql.AppendFormat("  ,[OVENCOOKEMP]='{0}',[OVENCOOKSTART]='{1}',[OVENCOOKEND]='{2}',[OVENCOOKTIMES]='{3}'", OVENCOOKEMP, OVENCOOKSTART, OVENCOOKEND, OVENCOOKTIMES);
                sbSql.AppendFormat("  ,[COLDCOOKEMP]='{0}',[COLDCOOKSTART]='{1}',[COLDCOOKEND]='{2}',[COLDCOOKTIMES]='{3}'", COLDCOOKEMP, COLDCOOKSTART, COLDCOOKEND, COLDCOOKTIMES);
                sbSql.AppendFormat("  ,[ARRAYEMP]='{0}',[ARRAYSTART]='{1}',[ARRAYEND]='{2}',[ARRAYTIMES]='{3}'", ARRAYEMP, ARRAYSTART, ARRAYEND, ARRAYTIMES);
                sbSql.AppendFormat("  ,[PACKEMP]='{0}',[PACKSTART]='{1}',[PACKEND]='{2}',[PACKTIMES]='{3}'", PACKEMP, PACKSTART, PACKEND, PACKTIMES);
                sbSql.AppendFormat("  ,[PACKPICKEMP]='{0}',[PACKPICKSTART]='{1}',[PACKPICKEND]='{2}',[PACKPICKTIMES]='{3}'", PACKPICKEMP, PACKPICKSTART, PACKPICKEND, PACKPICKTIMES);
                sbSql.AppendFormat("  ,[BOXSEMP]='{0}',[BOXSSTART]='{1}',[BOXSEND]='{2}',[BOXSTIMES]='{3}'", BOXSEMP, BOXSSTART, BOXSEND, BOXSTIMES);
                sbSql.AppendFormat("  ,[HANDCOOKEMP]='{0}',[HANDCOOKSTART]='{1}',[HANDCOOKEND]='{2}',[HANDCOOKTIMES]='{3}'", HANDCOOKEMP, HANDCOOKSTART, HANDCOOKEND, HANDCOOKTIMES);
                sbSql.AppendFormat("  ,[SCALESWEIGHTEMP]='{0}',[SCALESWEIGHTSTART]='{1}',[SCALESWEIGHTEND]='{2}',[SCALESWEIGHTTIMES]='{3}'", SCALESWEIGHTEMP, SCALESWEIGHTSTART, SCALESWEIGHTEND, SCALESWEIGHTTIMES);
                sbSql.AppendFormat("  ,[OUTBOXSEMP]='{0}',[OUTBOXSSTART]='{1}',[OUTBOXSEND]='{2}',[OUTBOXSTIMES]='{3}'", OUTBOXSEMP, OUTBOXSSTART, OUTBOXSEND, OUTBOXSTIMES);
                sbSql.AppendFormat("  ,[SEALEMP]='{0}',[SEALSTART]='{1}',[SEALEND]='{2}',[SEALTIMES]='{3}'", SEALEMP, SEALSTART, SEALEND, SEALTIMES);
                sbSql.AppendFormat("  ,[THROWEMP]='{0}',[THROWSTART]='{1}',[THROWEND]='{2}',[THROWTIMES]='{3}'", THROWEMP, THROWSTART, THROWEND, THROWTIMES);
                sbSql.AppendFormat("  ,[BOXPACKEMP]='{0}',[BOXPACKSTART]='{1}',[BOXPACKEND]='{2}',[BOXPACKTIMES]='{3}'", BOXPACKEMP, BOXPACKSTART, BOXPACKEND, BOXPACKTIMES);
                sbSql.AppendFormat("  WHERE [ID]='{0}'",ID);
                sbSql.AppendFormat("  ");
                sbSql.AppendFormat("  ");
                sbSql.AppendFormat("  ");
                sbSql.AppendFormat("  ");
                sbSql.AppendFormat("  ");
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

        public void DELMOCDAILYWORKHRS(string ID)
        {
            try
            {
                //add ZWAREWHOUSEPURTH
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat("  DELETE [TKMOC].[dbo].[MOCDAILYWORKHRS]");               
                sbSql.AppendFormat("  WHERE [ID]='{0}'", ID);
                sbSql.AppendFormat("  ");
                sbSql.AppendFormat("  ");
                sbSql.AppendFormat("  ");
                sbSql.AppendFormat("  ");
                sbSql.AppendFormat("  ");
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

      

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            SETSTARTTIMES(dateTimePicker2.Value);
            CALTIMES();
        }

        private void dateTimePicker3_ValueChanged(object sender, EventArgs e)
        {
            SETENDTIMES(dateTimePicker3.Value);
            CALTIMES();
        }

        private void dateTimePicker5_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker6_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker7_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker8_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker9_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker10_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker11_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker12_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker13_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker14_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker15_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker16_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker17_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker18_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker19_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker20_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker21_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker22_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker23_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker24_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker25_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker26_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker27_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker28_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker29_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker30_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker31_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker32_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker33_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker34_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker35_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void dateTimePicker36_ValueChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void textBox31_TextChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }
        private void textBox41_TextChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void textBox51_TextChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void textBox61_TextChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void textBox71_TextChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void textBox81_TextChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void textBox91_TextChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void textBox101_TextChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void textBox111_TextChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void textBox121_TextChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void textBox131_TextChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void textBox32_TextChanged(object sender, EventArgs e)
        {
            CALTIMES();

        }

        private void textBox42_TextChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void textBox52_TextChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void textBox62_TextChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void textBox72_TextChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }

        private void textBox82_TextChanged(object sender, EventArgs e)
        {
            CALTIMES();
        }
        public void CALTIMES()
        {
            TimeSpan TS = new TimeSpan();
           

            if (!string.IsNullOrEmpty(textBox31.Text) && Convert.ToDecimal(new TimeSpan(dateTimePicker3.Value.Ticks - dateTimePicker2.Value.Ticks).TotalHours)>0)
            {
                numericUpDown31.Value = Convert.ToDecimal(textBox31.Text) * Convert.ToDecimal(new TimeSpan(dateTimePicker3.Value.Ticks - dateTimePicker2.Value.Ticks).TotalHours);
            }

            if (!string.IsNullOrEmpty(textBox41.Text) && Convert.ToDecimal(new TimeSpan(dateTimePicker6.Value.Ticks - dateTimePicker5.Value.Ticks).TotalHours) > 0)
            {
                numericUpDown41.Value = Convert.ToDecimal(textBox41.Text) * Convert.ToDecimal(new TimeSpan(dateTimePicker6.Value.Ticks - dateTimePicker5.Value.Ticks).TotalHours);
            }
            if (!string.IsNullOrEmpty(textBox51.Text) && Convert.ToDecimal(new TimeSpan(dateTimePicker8.Value.Ticks - dateTimePicker7.Value.Ticks).TotalHours) > 0)
            {
                numericUpDown51.Value = Convert.ToDecimal(textBox51.Text) * Convert.ToDecimal(new TimeSpan(dateTimePicker8.Value.Ticks - dateTimePicker7.Value.Ticks).TotalHours);
               
            }
            if (!string.IsNullOrEmpty(textBox61.Text) && Convert.ToDecimal(new TimeSpan(dateTimePicker10.Value.Ticks - dateTimePicker9.Value.Ticks).TotalHours) > 0)
            {
                numericUpDown61.Value = Convert.ToDecimal(textBox61.Text) * Convert.ToDecimal(new TimeSpan(dateTimePicker10.Value.Ticks - dateTimePicker9.Value.Ticks).TotalHours);
               
            }
            if (!string.IsNullOrEmpty(textBox71.Text) && Convert.ToDecimal(new TimeSpan(dateTimePicker12.Value.Ticks - dateTimePicker11.Value.Ticks).TotalHours) > 0)
            {
                numericUpDown71.Value = Convert.ToDecimal(textBox71.Text) * Convert.ToDecimal(new TimeSpan(dateTimePicker12.Value.Ticks - dateTimePicker11.Value.Ticks).TotalHours);
                
            }
            if (!string.IsNullOrEmpty(textBox81.Text) && Convert.ToDecimal(new TimeSpan(dateTimePicker14.Value.Ticks - dateTimePicker13.Value.Ticks).TotalHours) > 0)
            {
                numericUpDown81.Value = Convert.ToDecimal(textBox81.Text) * Convert.ToDecimal(new TimeSpan(dateTimePicker14.Value.Ticks - dateTimePicker13.Value.Ticks).TotalHours);
               
            }
            if (!string.IsNullOrEmpty(textBox91.Text) && Convert.ToDecimal(new TimeSpan(dateTimePicker16.Value.Ticks - dateTimePicker15.Value.Ticks).TotalHours) > 0)
            {
                numericUpDown91.Value = Convert.ToDecimal(textBox91.Text) * Convert.ToDecimal(new TimeSpan(dateTimePicker16.Value.Ticks - dateTimePicker15.Value.Ticks).TotalHours);
               
            }
            if (!string.IsNullOrEmpty(textBox101.Text) && Convert.ToDecimal(new TimeSpan(dateTimePicker18.Value.Ticks - dateTimePicker17.Value.Ticks).TotalHours) > 0)
            {
                numericUpDown101.Value = Convert.ToDecimal(textBox101.Text) * Convert.ToDecimal(new TimeSpan(dateTimePicker18.Value.Ticks - dateTimePicker17.Value.Ticks).TotalHours);
                
            }
            if (!string.IsNullOrEmpty(textBox111.Text) && Convert.ToDecimal(new TimeSpan(dateTimePicker20.Value.Ticks - dateTimePicker19.Value.Ticks).TotalHours) > 0)
            {
                numericUpDown111.Value = Convert.ToDecimal(textBox111.Text) * Convert.ToDecimal(new TimeSpan(dateTimePicker20.Value.Ticks - dateTimePicker19.Value.Ticks).TotalHours);
               
            }
            if (!string.IsNullOrEmpty(textBox121.Text) && Convert.ToDecimal(new TimeSpan(dateTimePicker22.Value.Ticks - dateTimePicker21.Value.Ticks).TotalHours) > 0)
            {
                numericUpDown121.Value = Convert.ToDecimal(textBox121.Text) * Convert.ToDecimal(new TimeSpan(dateTimePicker22.Value.Ticks - dateTimePicker21.Value.Ticks).TotalHours);
                
            }
            if (!string.IsNullOrEmpty(textBox131.Text) && Convert.ToDecimal(new TimeSpan(dateTimePicker24.Value.Ticks - dateTimePicker23.Value.Ticks).TotalHours) > 0)
            {
                numericUpDown131.Value = Convert.ToDecimal(textBox131.Text) * Convert.ToDecimal(new TimeSpan(dateTimePicker24.Value.Ticks - dateTimePicker23.Value.Ticks).TotalHours);

              
            }
            if (!string.IsNullOrEmpty(textBox32.Text) && Convert.ToDecimal(new TimeSpan(dateTimePicker26.Value.Ticks - dateTimePicker25.Value.Ticks).TotalHours) > 0)
            {
                numericUpDown32.Value = Convert.ToDecimal(textBox32.Text) * Convert.ToDecimal(new TimeSpan(dateTimePicker26.Value.Ticks - dateTimePicker25.Value.Ticks).TotalHours);
              
            }
            if (!string.IsNullOrEmpty(textBox42.Text) && Convert.ToDecimal(new TimeSpan(dateTimePicker28.Value.Ticks - dateTimePicker27.Value.Ticks).TotalHours) > 0)
            {
                numericUpDown42.Value = Convert.ToDecimal(textBox42.Text) * Convert.ToDecimal(new TimeSpan(dateTimePicker28.Value.Ticks - dateTimePicker27.Value.Ticks).TotalHours);
               
            }
            if (!string.IsNullOrEmpty(textBox52.Text) && Convert.ToDecimal(new TimeSpan(dateTimePicker30.Value.Ticks - dateTimePicker29.Value.Ticks).TotalHours) > 0)
            {
                numericUpDown52.Value = Convert.ToDecimal(textBox52.Text) * Convert.ToDecimal(new TimeSpan(dateTimePicker30.Value.Ticks - dateTimePicker29.Value.Ticks).TotalHours);
              
            }
            if (!string.IsNullOrEmpty(textBox62.Text) && Convert.ToDecimal(new TimeSpan(dateTimePicker32.Value.Ticks - dateTimePicker31.Value.Ticks).TotalHours) > 0)
            {
                numericUpDown62.Value = Convert.ToDecimal(textBox62.Text) * Convert.ToDecimal(new TimeSpan(dateTimePicker32.Value.Ticks - dateTimePicker31.Value.Ticks).TotalHours);
                
            }
            if (!string.IsNullOrEmpty(textBox72.Text) && Convert.ToDecimal(new TimeSpan(dateTimePicker34.Value.Ticks - dateTimePicker33.Value.Ticks).TotalHours) > 0)
            {               
                numericUpDown72.Value = Convert.ToDecimal(textBox72.Text) * Convert.ToDecimal(new TimeSpan(dateTimePicker34.Value.Ticks - dateTimePicker33.Value.Ticks).TotalHours);
                            }
            if (!string.IsNullOrEmpty(textBox82.Text) && Convert.ToDecimal(new TimeSpan(dateTimePicker36.Value.Ticks - dateTimePicker35.Value.Ticks).TotalHours) > 0)
            {
                numericUpDown82.Value = Convert.ToDecimal(textBox82.Text) * Convert.ToDecimal(new TimeSpan(dateTimePicker36.Value.Ticks - dateTimePicker35.Value.Ticks).TotalHours);

            }

            numericUpDown11.Value = numericUpDown31.Value + numericUpDown41.Value + numericUpDown51.Value + numericUpDown61.Value + numericUpDown71.Value + numericUpDown81.Value + numericUpDown91.Value + numericUpDown101.Value + numericUpDown111.Value + numericUpDown121.Value + numericUpDown131.Value + numericUpDown32.Value + numericUpDown42.Value + numericUpDown52.Value + numericUpDown62.Value + numericUpDown72.Value + numericUpDown82.Value;

        }

        public void SETTEXT1()
        {
            textBox21.Text = null;
            textBox22.Text = null;
            textBox23.Text = null;
            textBox24.Text = null;
        }

        public void SETTIMES()
        {
            DateTime dt = Convert.ToDateTime(DateTime.Now.Year+"/"+ DateTime.Now.Month + "/" + DateTime.Now.Day+" "+ DateTime.Now.Hour+":"+ DateTime.Now.Minute);
            dateTimePicker2.Value = dt;
            dateTimePicker5.Value = dt;
            dateTimePicker7.Value = dt;
            dateTimePicker9.Value = dt;
            dateTimePicker11.Value = dt;
            dateTimePicker13.Value = dt;
            dateTimePicker15.Value = dt;
            dateTimePicker17.Value = dt;
            dateTimePicker19.Value = dt;
            dateTimePicker21.Value = dt;
            dateTimePicker23.Value = dt;
            dateTimePicker25.Value = dt;
            dateTimePicker27.Value = dt;
            dateTimePicker29.Value = dt;
            dateTimePicker31.Value = dt;
            dateTimePicker33.Value = dt;
            dateTimePicker35.Value = dt;

            dateTimePicker3.Value = dt;
            dateTimePicker6.Value = dt;
            dateTimePicker8.Value = dt;
            dateTimePicker10.Value = dt;
            dateTimePicker12.Value = dt;
            dateTimePicker14.Value = dt;
            dateTimePicker16.Value = dt;
            dateTimePicker18.Value = dt;
            dateTimePicker20.Value = dt;
            dateTimePicker22.Value = dt;
            dateTimePicker24.Value = dt;
            dateTimePicker26.Value = dt;
            dateTimePicker28.Value = dt;
            dateTimePicker30.Value = dt;
            dateTimePicker32.Value = dt;
            dateTimePicker34.Value = dt;
            dateTimePicker36.Value = dt;

         
        }

        public void SETSTARTTIMES(DateTime dt)
        {
            dateTimePicker2.Value = dt;
            dateTimePicker5.Value = dt;
            dateTimePicker7.Value = dt;
            dateTimePicker9.Value = dt;
            dateTimePicker11.Value = dt;
            dateTimePicker13.Value = dt;
            dateTimePicker15.Value = dt;
            dateTimePicker17.Value = dt;
            dateTimePicker19.Value = dt;
            dateTimePicker21.Value = dt;
            dateTimePicker23.Value = dt;
            dateTimePicker25.Value = dt;
            dateTimePicker27.Value = dt;
            dateTimePicker29.Value = dt;
            dateTimePicker31.Value = dt;
            dateTimePicker33.Value = dt;
            dateTimePicker35.Value = dt;
        }

        public void SETENDTIMES(DateTime dt)
        {
            dateTimePicker3.Value = dt;
            dateTimePicker6.Value = dt;
            dateTimePicker8.Value = dt;
            dateTimePicker10.Value = dt;
            dateTimePicker12.Value = dt;
            dateTimePicker14.Value = dt;
            dateTimePicker16.Value = dt;
            dateTimePicker18.Value = dt;
            dateTimePicker20.Value = dt;
            dateTimePicker22.Value = dt;
            dateTimePicker24.Value = dt;
            dateTimePicker26.Value = dt;
            dateTimePicker28.Value = dt;
            dateTimePicker30.Value = dt;
            dateTimePicker32.Value = dt;
            dateTimePicker34.Value = dt;
            dateTimePicker36.Value = dt;
        }


        public void SETTEXTBOX1()
        {
            textBox11.Text = null;
            textBox12.Text = null;
            textBox31.Text = null;
            textBox41.Text = null;
            textBox51.Text = null;
            textBox61.Text = null;
            textBox71.Text = null;
            textBox81.Text = null;
            textBox91.Text = null;
            textBox101.Text = null;
            textBox111.Text = null;
            textBox121.Text = null;
            textBox131.Text = null;
            textBox32.Text = null;
            textBox42.Text = null;
            textBox52.Text = null;
            textBox62.Text = null;
            textBox72.Text = null;
            textBox82.Text = null;

            textBox21.Text = null;
            textBox22.Text = null;
            textBox23.Text = null;
            textBox24.Text = null;
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

                    DateTime dt1 = Convert.ToDateTime(row.Cells["日期"].Value.ToString().Substring(0, 4) + "/" + row.Cells["日期"].Value.ToString().Substring(4, 2) + "/" + row.Cells["日期"].Value.ToString().Substring(6, 2));
                    dateTimePicker4.Value = dt1;
                   

                    comboBox1.Text = row.Cells["產線別"].Value.ToString();
                    textBox11.Text = row.Cells["製令單"].Value.ToString();
                    textBox12.Text = row.Cells["製令單號"].Value.ToString();

                    textBox31.Text = row.Cells["水麵攪拌"].Value.ToString();
                    dateTimePicker2.Value = Convert.ToDateTime("1911-1-1 " + row.Cells["水麵攪拌開始時間"].Value.ToString().Substring(0, 8));
                    dateTimePicker3.Value = Convert.ToDateTime("1911-1-1 " + row.Cells["水麵攪拌結束時間"].Value.ToString().Substring(0, 8));
                    numericUpDown31.Value = Convert.ToDecimal(row.Cells["水麵攪拌工時"].Value.ToString());

                    textBox41.Text = row.Cells["油酥攪拌"].Value.ToString();
                    dateTimePicker5.Value = Convert.ToDateTime("1911-1-1 " + row.Cells["油酥攪拌開始時間"].Value.ToString().Substring(0, 8));
                    dateTimePicker6.Value = Convert.ToDateTime("1911-1-1 " + row.Cells["油酥攪拌結束時間"].Value.ToString().Substring(0, 8));
                    numericUpDown41.Value = Convert.ToDecimal(row.Cells["油酥攪拌工時"].Value.ToString());

                    textBox51.Text = row.Cells["摺疊"].Value.ToString();
                    dateTimePicker7.Value = Convert.ToDateTime("1911-1-1 " + row.Cells["摺疊開始時間"].Value.ToString().Substring(0, 8));
                    dateTimePicker8.Value = Convert.ToDateTime("1911-1-1 " + row.Cells["摺疊結束時間"].Value.ToString().Substring(0, 8));
                    numericUpDown51.Value = Convert.ToDecimal(row.Cells["摺疊工時"].Value.ToString());

                    textBox61.Text = row.Cells["舖餅"].Value.ToString();
                    dateTimePicker9.Value = Convert.ToDateTime("1911-1-1 " + row.Cells["舖餅開始時間"].Value.ToString().Substring(0, 8));
                    dateTimePicker10.Value = Convert.ToDateTime("1911-1-1 " + row.Cells["舖餅結束時間"].Value.ToString().Substring(0, 8));
                    numericUpDown61.Value = Convert.ToDecimal(row.Cells["舖餅工時"].Value.ToString());

                    textBox71.Text = row.Cells["成型/烘烤"].Value.ToString();
                    dateTimePicker11.Value = Convert.ToDateTime("1911-1-1 " + row.Cells["烤箱篩餅開始時間"].Value.ToString().Substring(0, 8));
                    dateTimePicker12.Value = Convert.ToDateTime("1911-1-1 " + row.Cells["烤箱篩餅結束時間"].Value.ToString().Substring(0, 8));
                    numericUpDown71.Value = Convert.ToDecimal(row.Cells["烤箱篩餅工時"].Value.ToString());

                    textBox81.Text = row.Cells["烤箱篩餅"].Value.ToString();
                    dateTimePicker13.Value = Convert.ToDateTime("1911-1-1 " + row.Cells["冷卻篩餅開始時間"].Value.ToString().Substring(0, 8));
                    dateTimePicker14.Value = Convert.ToDateTime("1911-1-1 " + row.Cells["冷卻篩餅結束時間"].Value.ToString().Substring(0, 8));
                    numericUpDown81.Value = Convert.ToDecimal(row.Cells["冷卻篩餅工時"].Value.ToString());

                    textBox91.Text = row.Cells["冷卻篩餅"].Value.ToString();
                    dateTimePicker15.Value = Convert.ToDateTime("1911-1-1 " + row.Cells["冷卻篩餅開始時間"].Value.ToString().Substring(0, 8));
                    dateTimePicker16.Value = Convert.ToDateTime("1911-1-1 " + row.Cells["冷卻篩餅結束時間"].Value.ToString().Substring(0, 8));
                    numericUpDown91.Value = Convert.ToDecimal(row.Cells["冷卻篩餅工時"].Value.ToString());

                    textBox101.Text = row.Cells["排餅/裝罐"].Value.ToString();
                    dateTimePicker17.Value = Convert.ToDateTime("1911-1-1 " + row.Cells["排餅/裝罐開始時間"].Value.ToString().Substring(0, 8));
                    dateTimePicker18.Value = Convert.ToDateTime("1911-1-1 " + row.Cells["排餅/裝罐結束時間"].Value.ToString().Substring(0, 8));
                    numericUpDown101.Value = Convert.ToDecimal(row.Cells["排餅/裝罐工時"].Value.ToString());

                    textBox111.Text = row.Cells["包裝機"].Value.ToString();
                    dateTimePicker19.Value = Convert.ToDateTime("1911-1-1 " + row.Cells["包裝機開始時間"].Value.ToString().Substring(0, 8));
                    dateTimePicker20.Value = Convert.ToDateTime("1911-1-1 " + row.Cells["包裝機結束時間"].Value.ToString().Substring(0, 8));
                    numericUpDown111.Value = Convert.ToDecimal(row.Cells["包裝機工時"].Value.ToString());

                    textBox121.Text = row.Cells["包裝篩餅"].Value.ToString();
                    dateTimePicker21.Value = Convert.ToDateTime("1911-1-1 " + row.Cells["包裝篩餅開始時間"].Value.ToString().Substring(0, 8));
                    dateTimePicker22.Value = Convert.ToDateTime("1911-1-1 " + row.Cells["包裝篩餅結束時間"].Value.ToString().Substring(0, 8));
                    numericUpDown121.Value = Convert.ToDecimal(row.Cells["包裝篩餅工時"].Value.ToString());

                    textBox131.Text = row.Cells["裝箱"].Value.ToString();
                    dateTimePicker23.Value = Convert.ToDateTime("1911-1-1 " + row.Cells["裝箱開始時間"].Value.ToString().Substring(0, 8));
                    dateTimePicker24.Value = Convert.ToDateTime("1911-1-1 " + row.Cells["裝箱結束時間"].Value.ToString().Substring(0, 8));
                    numericUpDown131.Value = Convert.ToDecimal(row.Cells["裝箱工時"].Value.ToString());

                    textBox32.Text = row.Cells["撿餅"].Value.ToString();
                    dateTimePicker25.Value = Convert.ToDateTime("1911-1-1 " + row.Cells["撿餅開始時間"].Value.ToString().Substring(0, 8));
                    dateTimePicker26.Value = Convert.ToDateTime("1911-1-1 " + row.Cells["撿餅結束時間"].Value.ToString().Substring(0, 8));
                    numericUpDown32.Value = Convert.ToDecimal(row.Cells["撿餅工時"].Value.ToString());

                    textBox42.Text = row.Cells["秤重"].Value.ToString();
                    dateTimePicker27.Value = Convert.ToDateTime("1911-1-1 " + row.Cells["秤重開始時間"].Value.ToString().Substring(0, 8));
                    dateTimePicker28.Value = Convert.ToDateTime("1911-1-1 " + row.Cells["秤重結束時間"].Value.ToString().Substring(0, 8));
                    numericUpDown42.Value = Convert.ToDecimal(row.Cells["秤重工時"].Value.ToString());

                    textBox52.Text = row.Cells["外裝箱"].Value.ToString();
                    dateTimePicker29.Value = Convert.ToDateTime("1911-1-1 " + row.Cells["外裝箱開始時間"].Value.ToString().Substring(0, 8));
                    dateTimePicker30.Value = Convert.ToDateTime("1911-1-1 " + row.Cells["外裝箱結束時間"].Value.ToString().Substring(0, 8));
                    numericUpDown52.Value = Convert.ToDecimal(row.Cells["外裝箱工時"].Value.ToString());

                    textBox62.Text = row.Cells["封箱"].Value.ToString();
                    dateTimePicker31.Value = Convert.ToDateTime("1911-1-1 " + row.Cells["封箱開始時間"].Value.ToString().Substring(0, 8));
                    dateTimePicker32.Value = Convert.ToDateTime("1911-1-1 " + row.Cells["封箱結束時間"].Value.ToString().Substring(0, 8));
                    numericUpDown62.Value = Convert.ToDecimal(row.Cells["封箱工時"].Value.ToString());

                    textBox72.Text = row.Cells["倒餅"].Value.ToString();
                    dateTimePicker33.Value = Convert.ToDateTime("1911-1-1 " + row.Cells["倒餅開始時間"].Value.ToString().Substring(0, 8));
                    dateTimePicker34.Value = Convert.ToDateTime("1911-1-1 " + row.Cells["倒餅結束時間"].Value.ToString().Substring(0, 8));
                    numericUpDown72.Value = Convert.ToDecimal(row.Cells["倒餅工時"].Value.ToString());

                    textBox82.Text = row.Cells["封盒機"].Value.ToString();
                    dateTimePicker35.Value = Convert.ToDateTime("1911-1-1 " + row.Cells["封盒機開始時間"].Value.ToString().Substring(0, 8));
                    dateTimePicker36.Value = Convert.ToDateTime("1911-1-1 " + row.Cells["封盒機結束時間"].Value.ToString().Substring(0, 8));
                    numericUpDown82.Value = Convert.ToDecimal(row.Cells["封盒機工時"].Value.ToString());


                }
            }
            else
            {
                ID = null;
            }

          
        }


        public void SETFASTREPORT(string SDAY,string EDAY,string MB001)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL1(SDAY, EDAY, MB001);
            Report report1 = new Report();
            report1.Load(@"REPORT\生產工時記錄.frx");

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL1(string SDAY, string EDAY, string MB001)
        {
            StringBuilder SB = new StringBuilder();
            StringBuilder SBQUERY = new StringBuilder();

            if(!string.IsNullOrEmpty(MB001))
            {
                SBQUERY.AppendFormat(" AND MB001 LIKE '%{0}%'", MB001);
            }
            

            SB.AppendFormat(@"  SELECT  CONVERT(NVARCHAR,[DATS],112) AS '日期',[MANU] AS '產線別',[TA001] AS '製令單',[TA002] AS '製令單號',[MB001] AS '品號',[MB002] AS '品名',[NUMS] AS '入庫量',[MOCNUM] AS '預計生產量'");
            SB.AppendFormat(@"  ,CONVERT(NVARCHAR,[WORKSTART],114) AS '開始時間',CONVERT(NVARCHAR,[WORKEND],114) AS '結束時間',[WORKHRS] AS '工時',[WORKTIMES] AS '工時(分)',[AVGWORKHRS] AS '平均工時'");
            SB.AppendFormat(@"  ,[WATERNOODLESEMP] AS '水麵攪拌',CONVERT(NVARCHAR,[WATERNOODLESSTART],114) AS '水麵攪拌開始時間',CONVERT(NVARCHAR,[WATERNOODLESEND],114) AS '水麵攪拌結束時間',[WATERNOODLESTIMES] AS '水麵攪拌工時'");
            SB.AppendFormat(@"  ,[OILPASTRYEMP] AS '油酥攪拌',CONVERT(NVARCHAR,[OILPASTRYSTART],114) AS '油酥攪拌開始時間',CONVERT(NVARCHAR,[OILPASTRYEND],114) AS '油酥攪拌結束時間',[OILPASTRYTIMES] AS '油酥攪拌工時'");
            SB.AppendFormat(@"  ,[FOLDEMP] AS '摺疊',CONVERT(NVARCHAR,[FOLDSTART],114) AS '摺疊開始時間',CONVERT(NVARCHAR,[FOLDEND],114) AS '摺疊結束時間',[FOLDTIMES] AS '摺疊工時'");
            SB.AppendFormat(@"  ,[TYPECOOKEMP] AS '舖餅',CONVERT(NVARCHAR,[TYPECOOKSTART],114) AS '舖餅開始時間',CONVERT(NVARCHAR,[TYPECOOKEND],114) AS '舖餅結束時間',[TYPECOOKTIMES] AS '舖餅工時'");
            SB.AppendFormat(@"  ,[TYPEEMP] AS '成型/烘烤',CONVERT(NVARCHAR,[TYPESTART],114) AS '成型/烘烤開始時間',CONVERT(NVARCHAR,[TYPEEND],114) AS '成型/烘烤結束時間',[TYPETIMES] AS '成型/烘烤工時'");
            SB.AppendFormat(@"  ,[OVENCOOKEMP] AS '烤箱篩餅',CONVERT(NVARCHAR,[OVENCOOKSTART],114) AS '烤箱篩餅開始時間',CONVERT(NVARCHAR,[OVENCOOKEND],114) AS '烤箱篩餅結束時間',[OVENCOOKTIMES] AS '烤箱篩餅工時'");
            SB.AppendFormat(@"  ,[COLDCOOKEMP] AS '冷卻篩餅',CONVERT(NVARCHAR,[COLDCOOKSTART],114) AS '冷卻篩餅開始時間',CONVERT(NVARCHAR,[COLDCOOKEND],114) AS '冷卻篩餅結束時間',[COLDCOOKTIMES] AS '冷卻篩餅工時'");
            SB.AppendFormat(@"  ,[ARRAYEMP] AS '排餅/裝罐',CONVERT(NVARCHAR,[ARRAYSTART],114) AS '排餅/裝罐開始時間',CONVERT(NVARCHAR,[ARRAYEND],114) AS '排餅/裝罐結束時間',[ARRAYTIMES] AS '排餅/裝罐工時'");
            SB.AppendFormat(@"  ,[PACKEMP] AS '包裝機',CONVERT(NVARCHAR,[PACKSTART],114) AS '包裝機開始時間',CONVERT(NVARCHAR,[PACKEND],114) AS '包裝機結束時間',[PACKTIMES] AS '包裝機工時'");
            SB.AppendFormat(@"  ,[PACKPICKEMP] AS '包裝篩餅',CONVERT(NVARCHAR,[PACKPICKSTART],114) AS '包裝篩餅開始時間',CONVERT(NVARCHAR,[PACKPICKEND],114) AS '包裝篩餅結束時間',[PACKPICKTIMES] AS '包裝篩餅工時'");
            SB.AppendFormat(@"  ,[BOXSEMP] AS '裝箱',CONVERT(NVARCHAR,[BOXSSTART],114) AS '裝箱開始時間',CONVERT(NVARCHAR,[BOXSEND],114) AS '裝箱結束時間',[BOXSTIMES] AS '裝箱工時'");
            SB.AppendFormat(@"  ,[HANDCOOKEMP] AS '撿餅',CONVERT(NVARCHAR,[HANDCOOKSTART],114) AS '撿餅開始時間',CONVERT(NVARCHAR,[HANDCOOKEND],114) AS '撿餅結束時間',[HANDCOOKTIMES] AS '撿餅工時'");
            SB.AppendFormat(@"  ,[SCALESWEIGHTEMP] AS '秤重',CONVERT(NVARCHAR,[SCALESWEIGHTSTART],114) AS '秤重開始時間',CONVERT(NVARCHAR,[SCALESWEIGHTEND],114) AS '秤重結束時間',[SCALESWEIGHTTIMES] AS '秤重工時'");
            SB.AppendFormat(@"  ,[OUTBOXSEMP] AS '外裝箱',CONVERT(NVARCHAR,[OUTBOXSSTART],114) AS '外裝箱開始時間',CONVERT(NVARCHAR,[OUTBOXSEND],114) AS '外裝箱結束時間',[OUTBOXSTIMES] AS '外裝箱工時'");
            SB.AppendFormat(@"  ,[SEALEMP] AS '封箱',CONVERT(NVARCHAR,[SEALSTART],114) AS '封箱開始時間',CONVERT(NVARCHAR,[SEALEND],114) AS '封箱結束時間',[SEALTIMES] AS '封箱工時'");
            SB.AppendFormat(@"  ,[THROWEMP] AS '倒餅',CONVERT(NVARCHAR,[THROWSTART],114) AS '倒餅開始時間',CONVERT(NVARCHAR,[THROWEND],114) AS '倒餅結束時間',[THROWTIMES] AS '倒餅工時'");
            SB.AppendFormat(@"  ,[BOXPACKEMP] AS '封盒機',CONVERT(NVARCHAR,[BOXPACKSTART],114) AS '封盒機開始時間',CONVERT(NVARCHAR,[BOXPACKEND],114) AS '封盒機結束時間',[BOXPACKTIMES] AS '封盒機工時'");
            SB.AppendFormat(@"  ,[ID]");
            SB.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCDAILYWORKHRS] ");
            SB.AppendFormat(@"  WHERE  CONVERT(NVARCHAR,[DATS],112)>='{0}' AND  CONVERT(NVARCHAR,[DATS],112)<='{1}'",SDAY,EDAY);
            SB.AppendFormat(@"   {0}",SBQUERY.ToString());
            SB.AppendFormat(@"  ORDER BY [TA001],[TA002]");
            SB.AppendFormat(@"   ");
            SB.AppendFormat(@"   ");


            return SB;

        }

        public void ADDCSTMB(string MB001,string MB002,string MB003,string MB004,string MB005,string MB006, string MB007)
        {
            DATACSTMB CSTMB = new DATACSTMB();
            CSTMB = SETCSTMB();

            try
            {
                //add ZWAREWHOUSEPURTH
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat("  INSERT INTO  [TK].[dbo].[CSTMB]");
                sbSql.AppendFormat("  (");
                sbSql.AppendFormat("  [COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE]");
                sbSql.AppendFormat("  ,[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count],[DataUser],[DataGroup]");
                sbSql.AppendFormat("  ,[MB001],[MB002],[MB003],[MB004],[MB005],[MB006],[MB007],[MB008],[MB009],[MB010]");
                sbSql.AppendFormat("  ,[MB011],[MB012],[MB013],[MB014],[MB015],[MB016],[MB017],[MB018],[MB019],[MB020]");
                sbSql.AppendFormat("  ,[MB021],[MB022]");
                sbSql.AppendFormat("  ,[UDF01],[UDF02],[UDF03],[UDF04],[UDF05],[UDF06],[UDF07],[UDF08],[UDF09],[UDF10]");
                sbSql.AppendFormat("  )");
                sbSql.AppendFormat("  VALUES");
                sbSql.AppendFormat("  (");
                sbSql.AppendFormat("  '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}'", CSTMB.COMPANY, CSTMB.CREATOR, CSTMB.USR_GROUP, MB002, CSTMB.MODIFIER, CSTMB.MODI_DATE, CSTMB.FLAG, CSTMB.CREATE_TIME, CSTMB.MODI_TIME, CSTMB.TRANS_TYPE);
                sbSql.AppendFormat("  ,'{0}','{1}','{2}','{3}','{4}','{5}','{6}'", CSTMB.TRANS_NAME, CSTMB.sync_date, CSTMB.sync_time, CSTMB.sync_mark, CSTMB.sync_count, CSTMB.DataUser, CSTMB.DataGroup);
                sbSql.AppendFormat("  ,'{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}'", MB001, MB002,MB003, MB004, MB005, MB006, MB007, CSTMB.MB008, CSTMB.MB009, CSTMB.MB010);
                sbSql.AppendFormat("  ,'{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}'", CSTMB.MB011, CSTMB.MB012, CSTMB.MB013, CSTMB.MB014, CSTMB.MB015, CSTMB.MB016, CSTMB.MB017, CSTMB.MB018, CSTMB.MB019, CSTMB.MB020);
                sbSql.AppendFormat("  ,'{0}','{1}'", CSTMB.MB021, CSTMB.MB022);
                sbSql.AppendFormat("  ,'{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}'", CSTMB.UDF01, CSTMB.UDF02, CSTMB.UDF03, CSTMB.UDF04, CSTMB.UDF05, CSTMB.UDF06, CSTMB.UDF07, CSTMB.UDF08, CSTMB.UDF09, CSTMB.UDF10);
                sbSql.AppendFormat("  )");
                sbSql.AppendFormat("  ");
                sbSql.AppendFormat("  ");
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

        public void UPDATECSTMB(string MB001, string MB002, string MB003, string MB004, string MB005, string MB006, string MB007)
        {

        }

        public DATACSTMB SETCSTMB()
        {
            DATACSTMB CSTMB = new DATACSTMB();
            CSTMB.COMPANY = "TK";
            CSTMB.CREATOR = "100008";
            CSTMB.USR_GROUP = "103000";
            CSTMB.CREATE_DATE = "";
            CSTMB.MODIFIER = "";
            CSTMB.MODI_DATE = "";
            CSTMB.FLAG = "1";
            CSTMB.CREATE_TIME = DateTime.Now.ToString("HH:mm:ss");
            CSTMB.MODI_TIME = "";
            CSTMB.TRANS_TYPE = "P001";
            CSTMB.TRANS_NAME = "CSTI02";
            CSTMB.sync_date = "";
            CSTMB.sync_time = DateTime.Now.ToString("HH:mm:ss");
            CSTMB.sync_mark = "";
            CSTMB.sync_count = "0";
            CSTMB.DataUser = "";
            CSTMB.DataGroup = "103000";
            CSTMB.MB001 = "";
            CSTMB.MB002 = "";
            CSTMB.MB003 = "";
            CSTMB.MB004 = "";
            CSTMB.MB005 = "";
            CSTMB.MB006 = "0";
            CSTMB.MB007 = "";
            CSTMB.MB008 = "0";
            CSTMB.MB009 = "0";
            CSTMB.MB010 = "";
            CSTMB.MB011 = "";
            CSTMB.MB012 = "";
            CSTMB.MB013 = "0";
            CSTMB.MB014 = "0";
            CSTMB.MB015 = "0";
            CSTMB.MB016 = "0";
            CSTMB.MB017 = "";
            CSTMB.MB018 = "";
            CSTMB.MB019 = "";
            CSTMB.MB020 = "";
            CSTMB.MB021 = "";
            CSTMB.MB022 = "";
            CSTMB.UDF01 = "";
            CSTMB.UDF02 = "";
            CSTMB.UDF03 = "";
            CSTMB.UDF04 = "";
            CSTMB.UDF05 = "";
            CSTMB.UDF06 = "0";
            CSTMB.UDF07 = "0";
            CSTMB.UDF08 = "0";
            CSTMB.UDF09 = "0";
            CSTMB.UDF10 = "0";


            return CSTMB;
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
            label26.Text= "ADD";
            SETTEXTBOX1();

        }
        private void button3_Click(object sender, EventArgs e)
        {
            STATUS = "EDIT";
            label26.Text = "EDIT";

        }
        private void button4_Click(object sender, EventArgs e)
        {
            if(STATUS.Equals("ADD"))
            {
                ADDMOCDAILYWORKHRS(Guid.NewGuid().ToString(), dateTimePicker4.Value.ToString("yyyy/MM/dd"), comboBox1.Text.Trim(), textBox11.Text, textBox12.Text, textBox21.Text, textBox22.Text, textBox23.Text, textBox24.Text
                , dateTimePicker2.Value.ToString("HH:mm"), dateTimePicker3.Value.ToString("HH:mm"), numericUpDown11.Value.ToString(), "0","0"
                , textBox31.Text, dateTimePicker2.Value.ToString("HH:mm"), dateTimePicker3.Value.ToString("HH:mm"), numericUpDown31.Value.ToString()
                , textBox41.Text, dateTimePicker5.Value.ToString("HH:mm"), dateTimePicker6.Value.ToString("HH:mm"), numericUpDown41.Value.ToString()
                , textBox51.Text, dateTimePicker7.Value.ToString("HH:mm"), dateTimePicker8.Value.ToString("HH:mm"), numericUpDown51.Value.ToString()
                , textBox61.Text, dateTimePicker9.Value.ToString("HH:mm"), dateTimePicker10.Value.ToString("HH:mm"), numericUpDown61.Value.ToString()
                , textBox71.Text, dateTimePicker11.Value.ToString("HH:mm"), dateTimePicker12.Value.ToString("HH:mm"), numericUpDown71.Value.ToString()
                , textBox81.Text, dateTimePicker13.Value.ToString("HH:mm"), dateTimePicker14.Value.ToString("HH:mm"), numericUpDown81.Value.ToString()
                , textBox91.Text, dateTimePicker15.Value.ToString("HH:mm"), dateTimePicker16.Value.ToString("HH:mm"), numericUpDown91.Value.ToString()
                , textBox101.Text, dateTimePicker17.Value.ToString("HH:mm"), dateTimePicker18.Value.ToString("HH:mm"), numericUpDown101.Value.ToString()
                , textBox111.Text, dateTimePicker19.Value.ToString("HH:mm"), dateTimePicker20.Value.ToString("HH:mm"), numericUpDown111.Value.ToString()
                , textBox121.Text, dateTimePicker21.Value.ToString("HH:mm"), dateTimePicker22.Value.ToString("HH:mm"), numericUpDown121.Value.ToString()
                , textBox131.Text, dateTimePicker23.Value.ToString("HH:mm"), dateTimePicker24.Value.ToString("HH:mm"), numericUpDown131.Value.ToString()
                , textBox32.Text, dateTimePicker25.Value.ToString("HH:mm"), dateTimePicker26.Value.ToString("HH:mm"), numericUpDown32.Value.ToString()
                , textBox42.Text, dateTimePicker27.Value.ToString("HH:mm"), dateTimePicker28.Value.ToString("HH:mm"), numericUpDown42.Value.ToString()
                , textBox52.Text, dateTimePicker29.Value.ToString("HH:mm"), dateTimePicker30.Value.ToString("HH:mm"), numericUpDown52.Value.ToString()
                , textBox62.Text, dateTimePicker31.Value.ToString("HH:mm"), dateTimePicker32.Value.ToString("HH:mm"), numericUpDown62.Value.ToString()
                , textBox72.Text, dateTimePicker33.Value.ToString("HH:mm"), dateTimePicker34.Value.ToString("HH:mm"), numericUpDown72.Value.ToString()
                , textBox82.Text, dateTimePicker35.Value.ToString("HH:mm"), dateTimePicker36.Value.ToString("HH:mm"), numericUpDown82.Value.ToString()
                );

                ADDCSTMB(comboBox1.SelectedValue.ToString().Trim(),dateTimePicker4.Value.ToString("yyyyMMdd"),textBox11.Text.Trim(),textBox12.Text.Trim(),numericUpDown11.Value.ToString(),"0",textBox21.Text.Trim());
            }
            else if(STATUS.Equals("EDIT"))
            {
                UPDATEMOCDAILYWORKHRS(ID, dateTimePicker4.Value.ToString("yyyy/MM/dd"), comboBox1.Text.Trim(), textBox11.Text, textBox12.Text, textBox21.Text, textBox22.Text, textBox23.Text, textBox24.Text
               , dateTimePicker2.Value.ToString("HH:mm"), dateTimePicker3.Value.ToString("HH:mm"), numericUpDown11.Value.ToString(), "0", "0"
               , textBox31.Text, dateTimePicker2.Value.ToString("HH:mm"), dateTimePicker3.Value.ToString("HH:mm"), numericUpDown31.Value.ToString()
               , textBox41.Text, dateTimePicker5.Value.ToString("HH:mm"), dateTimePicker6.Value.ToString("HH:mm"), numericUpDown41.Value.ToString()
               , textBox51.Text, dateTimePicker7.Value.ToString("HH:mm"), dateTimePicker8.Value.ToString("HH:mm"), numericUpDown51.Value.ToString()
               , textBox61.Text, dateTimePicker9.Value.ToString("HH:mm"), dateTimePicker10.Value.ToString("HH:mm"), numericUpDown61.Value.ToString()
               , textBox71.Text, dateTimePicker11.Value.ToString("HH:mm"), dateTimePicker12.Value.ToString("HH:mm"), numericUpDown71.Value.ToString()
               , textBox81.Text, dateTimePicker13.Value.ToString("HH:mm"), dateTimePicker14.Value.ToString("HH:mm"), numericUpDown81.Value.ToString()
               , textBox91.Text, dateTimePicker15.Value.ToString("HH:mm"), dateTimePicker16.Value.ToString("HH:mm"), numericUpDown91.Value.ToString()
               , textBox101.Text, dateTimePicker17.Value.ToString("HH:mm"), dateTimePicker18.Value.ToString("HH:mm"), numericUpDown101.Value.ToString()
               , textBox111.Text, dateTimePicker19.Value.ToString("HH:mm"), dateTimePicker20.Value.ToString("HH:mm"), numericUpDown111.Value.ToString()
               , textBox121.Text, dateTimePicker21.Value.ToString("HH:mm"), dateTimePicker22.Value.ToString("HH:mm"), numericUpDown121.Value.ToString()
               , textBox131.Text, dateTimePicker23.Value.ToString("HH:mm"), dateTimePicker24.Value.ToString("HH:mm"), numericUpDown131.Value.ToString()
               , textBox32.Text, dateTimePicker25.Value.ToString("HH:mm"), dateTimePicker26.Value.ToString("HH:mm"), numericUpDown32.Value.ToString()
               , textBox42.Text, dateTimePicker27.Value.ToString("HH:mm"), dateTimePicker28.Value.ToString("HH:mm"), numericUpDown42.Value.ToString()
               , textBox52.Text, dateTimePicker29.Value.ToString("HH:mm"), dateTimePicker30.Value.ToString("HH:mm"), numericUpDown52.Value.ToString()
               , textBox62.Text, dateTimePicker31.Value.ToString("HH:mm"), dateTimePicker32.Value.ToString("HH:mm"), numericUpDown62.Value.ToString()
               , textBox72.Text, dateTimePicker33.Value.ToString("HH:mm"), dateTimePicker34.Value.ToString("HH:mm"), numericUpDown72.Value.ToString()
               , textBox82.Text, dateTimePicker35.Value.ToString("HH:mm"), dateTimePicker36.Value.ToString("HH:mm"), numericUpDown82.Value.ToString()
               );
            }

            STATUS = null;
            label26.Text = "STATUS";
            SEARCH(dateTimePicker1.Value.ToString("yyyyMMdd"));
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELMOCDAILYWORKHRS(ID);
                SEARCH(dateTimePicker1.Value.ToString("yyyyMMdd"));
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }





        private void button6_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(dateTimePicker39.Value.ToString("yyyyMMdd"),dateTimePicker40.Value.ToString("yyyyMMdd"),textBox1.Text.Trim());
        }






        #endregion

      
    }
}
