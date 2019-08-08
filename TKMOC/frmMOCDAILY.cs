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
using Calendar.NET;
using Excel = Microsoft.Office.Interop.Excel;
using FastReport;
using FastReport.Data;

namespace TKMOC
{
    public partial class frmMOCDAILY : Form
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
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        int rownum = 0;

        SqlTransaction tran;
       
        DataSet ds1 = new DataSet();
        int result;

        Report report1 = new Report();

        public frmMOCDAILY()
        {
            InitializeComponent();

            SEARCHMOCDAILYRECORDNG();
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
            comboBox1.ValueMember = "MD002";
            comboBox1.DisplayMember = "MD002";
            sqlConn.Close();


        }

        public void SETFASTREPORT()
        {
            StringBuilder SQL1 = new StringBuilder();
            StringBuilder SQL2 = new StringBuilder();

            SQL1 = SETSQL1();

            Report report1 = new Report();
            report1.Load(@"REPORT\生產報表-得料率報表.frx");

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL1()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(" SELECT ");
            SB.AppendFormat(" 線別,SUBSTRING(製令單號,1,8) AS '日期',品號,品名,規格,製令單別,製令單號,生產單位,預計產量,生產量,淨重,單片重,袋重,袋重比,蒸發率,原料用量,成品用量/1000 AS 成品用量,類別,領料是否扣袋重,成品是否扣袋重");
            SB.AppendFormat(" ,CASE WHEN 領料是否扣袋重 IN ('Y') AND 成品是否扣袋重 IN ('Y') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000)-(袋重比*原料用量))>0 THEN (((生產量*淨重*(1-袋重比)))/1000)/(原料用量*(1-(蒸發率/100))+(成品用量/1000)-(袋重比*原料用量)) ELSE 0 END  AS '領料扣成品扣的得料率'");
            SB.AppendFormat(" ,CASE WHEN 領料是否扣袋重 IN ('Y') AND 成品是否扣袋重 IN ('N') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000))>0 THEN (((生產量*淨重))/1000)/(原料用量*(1-蒸發率)+(成品用量/1000)-(袋重比*原料用量)) ELSE 0 END AS '領料扣成品不扣的得料率'");
            SB.AppendFormat(" ,CASE WHEN 領料是否扣袋重 IN ('N') AND 成品是否扣袋重 IN ('Y') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000)-(袋重比*原料用量))>0 THEN (((生產量*淨重*(1-袋重比)))/1000)/(原料用量*(1-(蒸發率/100))+(成品用量/1000)) ELSE 0 END AS '領料不扣成品扣的得料率'");
            SB.AppendFormat(" ,CASE WHEN 領料是否扣袋重 IN ('N') AND 成品是否扣袋重 IN ('N') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000))>0 THEN (((生產量*淨重)/1000)/(原料用量*(1-(蒸發率/100))+(成品用量/1000))) ELSE 0 END AS '領料不扣成品不扣的得料率'");
            SB.AppendFormat(" ,CASE WHEN 類別 IN ('半成品') AND 原料用量>0  AND 成品是否扣袋重 IN ('Y') THEN (生產量-(生產量*袋重比))/(原料用量*(1-蒸發率/100)) ELSE 0 END  AS '半成品得料率(成品扣袋重)'");
            SB.AppendFormat(" ,CASE WHEN 類別 IN ('半成品') AND 原料用量>0  AND 成品是否扣袋重 IN ('N') THEN (生產量)/(原料用量*(1-蒸發率/100)) ELSE 0 END  AS '半成品得料率(成品不扣袋重)'");
            SB.AppendFormat(" ,CASE WHEN 類別 IN ('個','試吃') AND 原料用量>0 AND (原料用量*(1-(蒸發率/100)))>0 THEN (生產量*淨重/1000)/(原料用量*(1-(蒸發率/100))) ELSE 0 END  AS '個/試吃得料率'");
            SB.AppendFormat(" ,CASE WHEN 類別 IN ('片') AND 原料用量>0 AND (原料用量*(1-(蒸發率/100))-(原料用量*袋重比))>0 THEN (生產量*淨重/1000)/(原料用量*(1-(蒸發率/100))-(原料用量*袋重比)) ELSE 0 END  AS '片得料率'");
            SB.AppendFormat(" ,CASE WHEN 類別 IN ('單包') AND 原料用量>0 THEN 生產量/原料用量  ELSE 0 END AS '單包得料率'");
            SB.AppendFormat(" ,CASE WHEN 類別 IN ('kg') AND (原料用量*(1-(蒸發率/100))-(原料用量*袋重比))>0 THEN ((生產量)/(原料用量*(1-(蒸發率/100))+(成品用量/1000)-(原料用量*袋重比))) ELSE 0 END AS 'kg得料率'");
            SB.AppendFormat(" ,(CASE WHEN 領料是否扣袋重 IN ('Y') AND 成品是否扣袋重 IN ('Y') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000)-(袋重比*原料用量))>0 THEN (((生產量*淨重*(1-袋重比)))/1000)/(原料用量*(1-(蒸發率/100))+(成品用量/1000)-(袋重比*原料用量)) ELSE 0 END)+(CASE WHEN 領料是否扣袋重 IN ('Y') AND 成品是否扣袋重 IN ('N') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000))>0 THEN (((生產量*淨重))/1000)/(原料用量*(1-蒸發率)+(成品用量/1000)-(袋重比*原料用量)) ELSE 0 END)+(CASE WHEN 領料是否扣袋重 IN ('N') AND 成品是否扣袋重 IN ('Y') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000)-(袋重比*原料用量))>0 THEN (((生產量*淨重*(1-袋重比)))/1000)/(原料用量*(1-(蒸發率/100))+(成品用量/1000)) ELSE 0 END)+(CASE WHEN 領料是否扣袋重 IN ('N') AND 成品是否扣袋重 IN ('N') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000))>0 THEN (((生產量*淨重)/1000)/(原料用量*(1-(蒸發率/100))+(成品用量/1000))) ELSE 0 END)+(CASE WHEN 類別 IN ('半成品') AND 原料用量>0  AND 成品是否扣袋重 IN ('Y') THEN (生產量-(生產量*袋重比))/(原料用量*(1-蒸發率/100)) ELSE 0 END)+(CASE WHEN 類別 IN ('半成品') AND 原料用量>0  AND 成品是否扣袋重 IN ('N') THEN (生產量)/(原料用量*(1-蒸發率/100)) ELSE 0 END)+(CASE WHEN 類別 IN ('個','試吃') AND 原料用量>0 AND (原料用量*(1-(蒸發率/100)))>0 THEN (生產量*淨重/1000)/(原料用量*(1-(蒸發率/100))) ELSE 0 END)+(CASE WHEN 類別 IN ('片') AND 原料用量>0 AND (原料用量*(1-(蒸發率/100))-(原料用量*袋重比))>0 THEN (生產量*淨重/1000)/(原料用量*(1-(蒸發率/100))-(原料用量*袋重比)) ELSE 0 END)+(CASE WHEN 類別 IN ('單包') AND 原料用量>0 THEN 生產量/原料用量  ELSE 0 END)+(CASE WHEN 類別 IN ('kg') AND (原料用量*(1-(蒸發率/100))-(原料用量*袋重比))>0 THEN ((生產量)/(原料用量*(1-(蒸發率/100))+(成品用量/1000)-(原料用量*袋重比))) ELSE 0 END) AS '得料率'");
            SB.AppendFormat(" ");
            SB.AppendFormat(" FROM(");
            SB.AppendFormat(" SELECT MD002 AS '線別',TA006 AS '品號',TA034 AS '品名',MB003 AS '規格' ,TA001 AS '製令單別',TA002 AS '製令單號',TA007 AS '生產單位',MB114 AS '類別',TA015 AS '預計產量',TA017 AS '生產量',INVMB.UDF07 AS '淨重',INVMB.UDF08 AS '單片重',INVMB.UDF09 AS '袋重',INVMB.UDF06 AS '蒸發率',MB112 AS '成品是否扣袋重',MB113 AS '領料是否扣袋重'");
            SB.AppendFormat(" ,(SELECT ISNULL(SUM(TB005),0) FROM [TK].dbo.MOCTB TB WHERE (TB.TB003 LIKE '1%' OR TB.TB003 LIKE '3%') AND TB.TB001=MOCTA.TA001 AND TB.TB002=MOCTA.TA002)  AS '原料用量'");
            SB.AppendFormat(" ,(SELECT ISNULL(SUM(TB005*MB.UDF07),0) FROM [TK].dbo.MOCTB TB,[TK].dbo.INVMB MB WHERE TB.TB003=MB.MB001 AND TB.TB003 LIKE '4%' AND TB.TB001=MOCTA.TA001 AND TB.TB002=MOCTA.TA002) AS '成品用量'");
            SB.AppendFormat(" ,CASE WHEN INVMB.UDF08>0 AND   INVMB.UDF09>0  THEN 1/(INVMB.UDF08+INVMB.UDF09)*INVMB.UDF09 ELSE 1 END  AS '袋重比'");
            SB.AppendFormat(" FROM [TK].dbo.INVMB,[TK].dbo.MOCTA,[TK].dbo.CMSMD");
            SB.AppendFormat(" WHERE TA006=MB001 AND TA021=MD001");
            SB.AppendFormat(" AND ISNULL(MB114,'')<>''");
            SB.AppendFormat(" AND TA003>='{0}' AND TA003<='{1}'",dateTimePicker1.Value.ToString("yyyyMMdd"),dateTimePicker2.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" ) AS TEMP");
            SB.AppendFormat(" WHERE 線別='{0}'",comboBox1.Text);
            SB.AppendFormat(" ORDER BY 線別,SUBSTRING(製令單號,1,8),品號");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");

            return SB;

        }


        public void SETNULL2()
        {
            textBox1.Text = null;
            textBox2.Text = null;
            textBox3.Text = null;
            textBox4.Text = null;
            

        }

        public void ADDMOCDAILYRECORDNG()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                
           
                sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[MOCDAILYRECORDNG]");
                sbSql.AppendFormat(" ([DATES],[MOCLINE],[NGCOOK],[NGCOOL],[NGPACKF],[NGPACKB])");
                sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}')", dateTimePicker3.Value.ToString("yyyy/MM/dd"),comboBox3.Text,textBox1.Text,textBox2.Text,textBox3.Text,textBox4.Text);
                sbSql.AppendFormat("  ");

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消

                    MessageBox.Show("失敗");
                }
                else
                {
                    tran.Commit();      //執行交易  

                    MessageBox.Show("成功");
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

        public void UPDATEMOCDAILYRECORDNG()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" UPDATE [TKMOC].[dbo].[MOCDAILYRECORDNG]");
                sbSql.AppendFormat(" SET [NGCOOK]='{0}',[NGCOOL]='{1}',[NGPACKF]='{2}',[NGPACKB]='{3}'",textBox1.Text,textBox2.Text,textBox3.Text,textBox4.Text);
                sbSql.AppendFormat(" WHERE [DATES]='{0}' AND [MOCLINE]='{1}'", dateTimePicker3.Value.ToString("yyyy/MM/dd"),comboBox3.Text);
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
                    MessageBox.Show("失敗");
                }
                else
                {
                    tran.Commit();      //執行交易  
                    MessageBox.Show("成功");

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

        private void dateTimePicker3_ValueChanged(object sender, EventArgs e)
        {
            SEARCHMOCDAILYRECORDNG();
        }

        public void SEARCHMOCDAILYRECORDNG()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@" SELECT CONVERT(NVARCHAR,[DATES],111) AS '日期',[MOCLINE] AS '線別',[NGCOOK] AS '可回收-烘焙不良品 	',[NGCOOL] AS '打餅區落地-冷卻不良品',[NGPACKF] AS '前端-包裝不良品',[NGPACKB] AS '後端落地-包裝不良品' ");
                sbSql.AppendFormat(@" FROM [TKMOC].[dbo].[MOCDAILYRECORDNG]");
                sbSql.AppendFormat(@" WHERE [MOCLINE]='{0}' AND CONVERT(NVARCHAR,[DATES],112) LIKE '{1}%' ", comboBox3.Text,dateTimePicker3.Value.ToString("yyyyMM"));
                sbSql.AppendFormat(@" ORDER BY CONVERT(NVARCHAR,[DATES],112)");
                sbSql.AppendFormat(@"  ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

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


        public void SETFASTREPORT2()
        {
            StringBuilder SQL1 = new StringBuilder();
            StringBuilder SQL2 = new StringBuilder();

            SQL1 = SETSQL2();

            Report report2 = new Report();
            report2.Load(@"REPORT\生產報表-每日得料率報表.frx");

            report2.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            TableDataSource table = report2.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            report2.Preview = previewControl2;
            report2.Show();
        }

        public StringBuilder SETSQL2()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(" SELECT 線別,SUBSTRING(製令單號,1,8) AS '日期'");
            SB.AppendFormat(" ,SUM(領料扣成品扣的投入量+領料扣成品不扣的投入量+領料不扣成品扣的投入量+領料不扣成品不扣的投入量+半成品得料率成品扣袋重的投入量+半成品得料率成品不扣袋重的投入量+個試吃的投入量+片的投入量+單包的投入量+kg的投入量) AS '應產出量 '");
            SB.AppendFormat(" ,SUM(領料扣成品扣的入庫量+領料扣成品不扣的入庫量+領料不扣成品扣的入庫量+領料不扣成品不扣的入庫量+半成品得料率成品扣袋重的入庫量+半成品得料率成品不扣袋重的入庫量+個試吃的入庫量+片的入庫量+單包的入庫量+kg的入庫量) AS '入庫淨重'");
            SB.AppendFormat(" ,CASE WHEN  SUM(領料扣成品扣的投入量+領料扣成品不扣的投入量+領料不扣成品扣的投入量+領料不扣成品不扣的投入量+半成品得料率成品扣袋重的投入量+半成品得料率成品不扣袋重的投入量+個試吃的投入量+片的投入量+單包的投入量+kg的投入量)>0 THEN SUM(領料扣成品扣的入庫量+領料扣成品不扣的入庫量+領料不扣成品扣的入庫量+領料不扣成品不扣的入庫量+半成品得料率成品扣袋重的入庫量+半成品得料率成品不扣袋重的入庫量+個試吃的入庫量+片的入庫量+單包的入庫量+kg的入庫量)/SUM(領料扣成品扣的投入量+領料扣成品不扣的投入量+領料不扣成品扣的投入量+領料不扣成品不扣的投入量+半成品得料率成品扣袋重的投入量+半成品得料率成品不扣袋重的投入量+個試吃的投入量+片的投入量+單包的投入量+kg的投入量) ELSE 0 END  AS '得料率(%)'");
            SB.AppendFormat(" ,ISNULL((SELECT [NGCOOK] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8)  AND MOCLINE=線別),0) AS '可回收-烘焙不良品 	'");
            SB.AppendFormat(" ,ISNULL((SELECT [NGCOOL] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8)  AND MOCLINE=線別),0) AS '打餅區落地-冷卻不良品'");
            SB.AppendFormat(" ,ISNULL((SELECT [NGPACKF] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8)  AND MOCLINE=線別),0) AS '前端-包裝不良品'");
            SB.AppendFormat(" ,ISNULL((SELECT [NGPACKB] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8)  AND MOCLINE=線別),0) AS '後端落地-包裝不良品'");
            SB.AppendFormat(" ,(ISNULL((SELECT [NGCOOK] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8)  AND MOCLINE=線別) ,0)+ISNULL((SELECT [NGCOOL] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8) AND MOCLINE=線別),0)+ISNULL((SELECT [NGPACKF] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8) AND MOCLINE=線別),0)+ISNULL((SELECT [NGPACKB] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8) AND MOCLINE=線別),0)) AS '不良合計'");
            SB.AppendFormat(" ,CASE WHEN  SUM(領料扣成品扣的投入量+領料扣成品不扣的投入量+領料不扣成品扣的投入量+領料不扣成品不扣的投入量+半成品得料率成品扣袋重的投入量+半成品得料率成品不扣袋重的投入量+個試吃的投入量+片的投入量+單包的投入量+kg的投入量)>0 THEN (ISNULL((SELECT [NGCOOK] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8) AND MOCLINE=線別),0)+ISNULL((SELECT [NGCOOL] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8) AND MOCLINE=線別),0)+ISNULL((SELECT [NGPACKF] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8) AND MOCLINE=線別),0)+ISNULL((SELECT [NGPACKB] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8) AND MOCLINE=線別),0))/SUM(領料扣成品扣的投入量+領料扣成品不扣的投入量+領料不扣成品扣的投入量+領料不扣成品不扣的投入量+半成品得料率成品扣袋重的投入量+半成品得料率成品不扣袋重的投入量+個試吃的投入量+片的投入量+單包的投入量+kg的投入量) ELSE 0 END  AS '報廢率(％)'");
            SB.AppendFormat(" ");
            SB.AppendFormat(" FROM ");
            SB.AppendFormat(" (");
            SB.AppendFormat(" SELECT ");
            SB.AppendFormat(" 線別,品號,品名,製令單別,製令單號,生產單位,類別,領料是否扣袋重,成品是否扣袋重,生產量,淨重,單片重,袋重,袋重比,蒸發率,原料用量,成品用量/1000 AS 成品用量");
            SB.AppendFormat(" ,CASE WHEN 領料是否扣袋重 IN ('Y') AND 成品是否扣袋重 IN ('Y') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000)-(袋重比*原料用量))>0 THEN (生產量*淨重*(1-袋重比)/1000) ELSE 0 END  AS '領料扣成品扣的入庫量'");
            SB.AppendFormat(" ,CASE WHEN 領料是否扣袋重 IN ('Y') AND 成品是否扣袋重 IN ('Y') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000)-(袋重比*原料用量))>0 THEN (原料用量*(1-(蒸發率/100))+(成品用量/1000)-(袋重比*原料用量)) ELSE 0 END  AS '領料扣成品扣的投入量'");
            SB.AppendFormat(" ,CASE WHEN 領料是否扣袋重 IN ('Y') AND 成品是否扣袋重 IN ('N') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000))>0 THEN ((生產量*淨重)/1000) ELSE 0 END AS '領料扣成品不扣的入庫量'");
            SB.AppendFormat(" ,CASE WHEN 領料是否扣袋重 IN ('Y') AND 成品是否扣袋重 IN ('N') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000))>0 THEN (原料用量*(1-蒸發率)+(成品用量/1000)-(袋重比*原料用量)) ELSE 0 END AS '領料扣成品不扣的投入量'");
            SB.AppendFormat(" ,CASE WHEN 領料是否扣袋重 IN ('N') AND 成品是否扣袋重 IN ('Y') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000)-(袋重比*原料用量))>0 THEN (((生產量*淨重*(1-袋重比)))/1000) ELSE 0 END AS '領料不扣成品扣的入庫量'");
            SB.AppendFormat(" ,CASE WHEN 領料是否扣袋重 IN ('N') AND 成品是否扣袋重 IN ('Y') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000)-(袋重比*原料用量))>0 THEN (原料用量*(1-(蒸發率/100))+(成品用量/1000)) ELSE 0 END AS '領料不扣成品扣的投入量'");
            SB.AppendFormat(" ,CASE WHEN 領料是否扣袋重 IN ('N') AND 成品是否扣袋重 IN ('N') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000))>0 THEN ((生產量*淨重)/1000) ELSE 0 END AS '領料不扣成品不扣的入庫量'");
            SB.AppendFormat(" ,CASE WHEN 領料是否扣袋重 IN ('N') AND 成品是否扣袋重 IN ('N') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000))>0 THEN (原料用量*(1-(蒸發率/100))+(成品用量/1000)) ELSE 0 END AS '領料不扣成品不扣的投入量'");
            SB.AppendFormat(" ,CASE WHEN 類別 IN ('半成品') AND 原料用量>0  AND 成品是否扣袋重 IN ('Y') THEN (生產量-(生產量*袋重比)) ELSE 0 END  AS '半成品得料率成品扣袋重的入庫量'");
            SB.AppendFormat(" ,CASE WHEN 類別 IN ('半成品') AND 原料用量>0  AND 成品是否扣袋重 IN ('Y') THEN (原料用量*(1-蒸發率/100)) ELSE 0 END  AS '半成品得料率成品扣袋重的投入量'");
            SB.AppendFormat(" ,CASE WHEN 類別 IN ('半成品') AND 原料用量>0  AND 成品是否扣袋重 IN ('N') THEN (生產量) ELSE 0 END  AS '半成品得料率成品不扣袋重的入庫量'");
            SB.AppendFormat(" ,CASE WHEN 類別 IN ('半成品') AND 原料用量>0  AND 成品是否扣袋重 IN ('N') THEN (原料用量*(1-蒸發率/100)) ELSE 0 END  AS '半成品得料率成品不扣袋重的投入量'");
            SB.AppendFormat(" ,CASE WHEN 類別 IN ('個','試吃') AND 原料用量>0 AND (原料用量*(1-(蒸發率/100)))>0 THEN (生產量*淨重/1000) ELSE 0 END  AS '個試吃的入庫量'");
            SB.AppendFormat(" ,CASE WHEN 類別 IN ('個','試吃') AND 原料用量>0 AND (原料用量*(1-(蒸發率/100)))>0 THEN (原料用量*(1-(蒸發率/100))) ELSE 0 END  AS '個試吃的投入量'");
            SB.AppendFormat(" ,CASE WHEN 類別 IN ('片') AND 原料用量>0 AND (原料用量*(1-(蒸發率/100))-(原料用量*袋重比))>0 THEN (生產量*淨重/1000) ELSE 0 END  AS '片的入庫量'");
            SB.AppendFormat(" ,CASE WHEN 類別 IN ('片') AND 原料用量>0 AND (原料用量*(1-(蒸發率/100))-(原料用量*袋重比))>0 THEN (原料用量*(1-(蒸發率/100))-(原料用量*袋重比)) ELSE 0 END  AS '片的投入量'");
            SB.AppendFormat(" ,CASE WHEN 類別 IN ('單包') AND 原料用量>0 THEN 生產量  ELSE 0 END AS '單包的入庫量'");
            SB.AppendFormat(" ,CASE WHEN 類別 IN ('單包') AND 原料用量>0 THEN 原料用量  ELSE 0 END AS '單包的投入量'");
            SB.AppendFormat(" ,CASE WHEN 類別 IN ('kg') AND (原料用量*(1-(蒸發率/100))-(原料用量*袋重比))>0 THEN (生產量) ELSE 0 END AS 'kg的入庫量'");
            SB.AppendFormat(" ,CASE WHEN 類別 IN ('kg') AND (原料用量*(1-(蒸發率/100))-(原料用量*袋重比))>0 THEN (原料用量*(1-(蒸發率/100))+(成品用量/1000)-(原料用量*袋重比)) ELSE 0 END AS 'kg的投入量'");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ,CASE WHEN 領料是否扣袋重 IN ('Y') AND 成品是否扣袋重 IN ('Y') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000)-(袋重比*原料用量))>0 THEN (((生產量*淨重*(1-袋重比)))/1000)/(原料用量*(1-(蒸發率/100))+(成品用量/1000)-(袋重比*原料用量)) ELSE 0 END  AS '領料扣成品扣的得料率'");
            SB.AppendFormat(" ,CASE WHEN 領料是否扣袋重 IN ('Y') AND 成品是否扣袋重 IN ('N') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000))>0 THEN (((生產量*淨重))/1000)/(原料用量*(1-蒸發率)+(成品用量/1000)-(袋重比*原料用量)) ELSE 0 END AS '領料扣成品不扣的得料率'");
            SB.AppendFormat(" ,CASE WHEN 領料是否扣袋重 IN ('N') AND 成品是否扣袋重 IN ('Y') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000)-(袋重比*原料用量))>0 THEN (((生產量*淨重*(1-袋重比)))/1000)/(原料用量*(1-(蒸發率/100))+(成品用量/1000)) ELSE 0 END AS '領料不扣成品扣的得料率'");
            SB.AppendFormat(" ,CASE WHEN 領料是否扣袋重 IN ('N') AND 成品是否扣袋重 IN ('N') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000))>0 THEN (((生產量*淨重)/1000)/(原料用量*(1-(蒸發率/100))+(成品用量/1000))) ELSE 0 END AS '領料不扣成品不扣的得料率'");
            SB.AppendFormat(" ,CASE WHEN 類別 IN ('半成品') AND 原料用量>0  AND 成品是否扣袋重 IN ('Y') THEN (生產量-(生產量*袋重比))/(原料用量*(1-蒸發率/100)) ELSE 0 END  AS '半成品得料率(成品扣袋重)'");
            SB.AppendFormat(" ,CASE WHEN 類別 IN ('半成品') AND 原料用量>0  AND 成品是否扣袋重 IN ('N') THEN (生產量)/(原料用量*(1-蒸發率/100)) ELSE 0 END  AS '半成品得料率(成品不扣袋重)'");
            SB.AppendFormat(" ,CASE WHEN 類別 IN ('個','試吃') AND 原料用量>0 AND (原料用量*(1-(蒸發率/100)))>0 THEN (生產量*淨重/1000)/(原料用量*(1-(蒸發率/100))) ELSE 0 END  AS '個/試吃得料率'");
            SB.AppendFormat(" ,CASE WHEN 類別 IN ('片') AND 原料用量>0 AND (原料用量*(1-(蒸發率/100))-(原料用量*袋重比))>0 THEN (生產量*淨重/1000)/(原料用量*(1-(蒸發率/100))-(原料用量*袋重比)) ELSE 0 END  AS '片得料率'");
            SB.AppendFormat(" ,CASE WHEN 類別 IN ('單包') AND 原料用量>0 THEN 生產量/原料用量  ELSE 0 END AS '單包得料率'");
            SB.AppendFormat(" ,CASE WHEN 類別 IN ('kg') AND (原料用量*(1-(蒸發率/100))-(原料用量*袋重比))>0 THEN ((生產量)/(原料用量*(1-(蒸發率/100))+(成品用量/1000)-(原料用量*袋重比))) ELSE 0 END AS 'kg得料率'");
            SB.AppendFormat(" ,(CASE WHEN 領料是否扣袋重 IN ('Y') AND 成品是否扣袋重 IN ('Y') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000)-(袋重比*原料用量))>0 THEN (((生產量*淨重*(1-袋重比)))/1000)/(原料用量*(1-(蒸發率/100))+(成品用量/1000)-(袋重比*原料用量)) ELSE 0 END)+(CASE WHEN 領料是否扣袋重 IN ('Y') AND 成品是否扣袋重 IN ('N') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000))>0 THEN (((生產量*淨重))/1000)/(原料用量*(1-蒸發率)+(成品用量/1000)-(袋重比*原料用量)) ELSE 0 END)+(CASE WHEN 領料是否扣袋重 IN ('N') AND 成品是否扣袋重 IN ('Y') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000)-(袋重比*原料用量))>0 THEN (((生產量*淨重*(1-袋重比)))/1000)/(原料用量*(1-(蒸發率/100))+(成品用量/1000)) ELSE 0 END)+(CASE WHEN 領料是否扣袋重 IN ('N') AND 成品是否扣袋重 IN ('N') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000))>0 THEN (((生產量*淨重)/1000)/(原料用量*(1-(蒸發率/100))+(成品用量/1000))) ELSE 0 END)+(CASE WHEN 類別 IN ('半成品') AND 原料用量>0  AND 成品是否扣袋重 IN ('Y') THEN (生產量-(生產量*袋重比))/(原料用量*(1-蒸發率/100)) ELSE 0 END)+(CASE WHEN 類別 IN ('半成品') AND 原料用量>0  AND 成品是否扣袋重 IN ('N') THEN (生產量)/(原料用量*(1-蒸發率/100)) ELSE 0 END)+(CASE WHEN 類別 IN ('個','試吃') AND 原料用量>0 AND (原料用量*(1-(蒸發率/100)))>0 THEN (生產量*淨重/1000)/(原料用量*(1-(蒸發率/100))) ELSE 0 END)+(CASE WHEN 類別 IN ('片') AND 原料用量>0 AND (原料用量*(1-(蒸發率/100))-(原料用量*袋重比))>0 THEN (生產量*淨重/1000)/(原料用量*(1-(蒸發率/100))-(原料用量*袋重比)) ELSE 0 END)+(CASE WHEN 類別 IN ('單包') AND 原料用量>0 THEN 生產量/原料用量  ELSE 0 END)+(CASE WHEN 類別 IN ('kg') AND (原料用量*(1-(蒸發率/100))-(原料用量*袋重比))>0 THEN ((生產量)/(原料用量*(1-(蒸發率/100))+(成品用量/1000)-(原料用量*袋重比))) ELSE 0 END) AS '得料率'");
            SB.AppendFormat(" ");
            SB.AppendFormat(" FROM(");
            SB.AppendFormat(" SELECT MD002 AS '線別',TA006 AS '品號',TA034 AS '品名',TA001 AS '製令單別',TA002 AS '製令單號',TA007 AS '生產單位',MB114 AS '類別',TA017 AS '生產量',INVMB.UDF07 AS '淨重',INVMB.UDF08 AS '單片重',INVMB.UDF09 AS '袋重',INVMB.UDF06 AS '蒸發率',MB112 AS '成品是否扣袋重',MB113 AS '領料是否扣袋重'");
            SB.AppendFormat(" ,(SELECT ISNULL(SUM(TB005),0) FROM [TK].dbo.MOCTB TB WHERE (TB.TB003 LIKE '1%' OR TB.TB003 LIKE '3%') AND TB.TB001=MOCTA.TA001 AND TB.TB002=MOCTA.TA002)  AS '原料用量'");
            SB.AppendFormat(" ,(SELECT ISNULL(SUM(TB005*MB.UDF07),0) FROM [TK].dbo.MOCTB TB,[TK].dbo.INVMB MB WHERE TB.TB003=MB.MB001 AND TB.TB003 LIKE '4%' AND TB.TB001=MOCTA.TA001 AND TB.TB002=MOCTA.TA002) AS '成品用量'");
            SB.AppendFormat(" ,CASE WHEN INVMB.UDF08>0 AND   INVMB.UDF09>0  THEN 1/(INVMB.UDF08+INVMB.UDF09)*INVMB.UDF09 ELSE 1 END  AS '袋重比'");
            SB.AppendFormat(" FROM [TK].dbo.INVMB,[TK].dbo.MOCTA,[TK].dbo.CMSMD");
            SB.AppendFormat(" WHERE TA006=MB001 AND TA021=MD001");
            SB.AppendFormat(" AND ISNULL(MB114,'')<>''");
            SB.AppendFormat(" AND TA003>='{0}' AND TA003<='{1}'",dateTimePicker4.Value.ToString("yyyyMMdd"), dateTimePicker5.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" ) AS TEMP");
            SB.AppendFormat(" ) AS TEMP2");
            SB.AppendFormat(" WHERE 線別='{0}'",comboBox2.Text);
            SB.AppendFormat(" GROUP BY 線別,SUBSTRING(製令單號,1,8)");
            SB.AppendFormat(" ORDER BY 線別,SUBSTRING(製令單號,1,8)");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");


            return SB;

        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            SEARCHMOCDAILYRECORDNG();
        }


        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            ADDMOCDAILYRECORDNG();
            SETNULL2();

            SEARCHMOCDAILYRECORDNG();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            UPDATEMOCDAILYRECORDNG();
            SETNULL2();

            SEARCHMOCDAILYRECORDNG();
        }
        private void button4_Click(object sender, EventArgs e)
        {
            SETFASTREPORT2();
        }


        #endregion

       
    }
}
