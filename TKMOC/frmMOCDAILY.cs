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
        SqlDataAdapter adapter2 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder2 = new SqlCommandBuilder();
        SqlDataAdapter adapter3 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder3 = new SqlCommandBuilder();
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3= new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        int rownum = 0;

        SqlTransaction tran;
       
        DataSet ds1 = new DataSet();
        int result;

        Report report1 = new Report();
        Report report2 = new Report();
        Report report3 = new Report();

        public frmMOCDAILY()
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
            comboBox10load();
            comboBox11load();

            SEARCHMOCDAILYRECORDNG();
            SEARCHMOCDAILYRECORDNG2();
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

        public void comboBox2load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MD001,MD002 FROM [TK].dbo.CMSMD    WHERE MD002 LIKE '新廠%' AND MD002 IN ('新廠製一組','新廠製二組')  ");
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
            Sequel.AppendFormat(@"SELECT MD001,MD002 FROM [TK].dbo.CMSMD    WHERE MD002 LIKE '新廠%'  AND MD002 IN ('新廠製一組','新廠製二組') ");
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
            Sequel.AppendFormat(@"SELECT MD001,MD002 FROM [TK].dbo.CMSMD    WHERE MD002 LIKE '新廠%' AND MD002 IN ('新廠製三組(手工)')  ");
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
            Sequel.AppendFormat(@"SELECT MD001,MD002 FROM [TK].dbo.CMSMD    WHERE MD002 LIKE '新廠%'  AND MD002 IN ('新廠製三組(手工)') ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MD001", typeof(string));
            dt.Columns.Add("MD002", typeof(string));
            da.Fill(dt);
            comboBox5.DataSource = dt.DefaultView;
            comboBox5.ValueMember = "MD002";
            comboBox5.DisplayMember = "MD002";
            sqlConn.Close();


        }

        public void comboBox6load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT [ID],[NAME]  FROM [TKMOC].[dbo].[MOCDAILYSLOTNAME] ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox6.DataSource = dt.DefaultView;
            comboBox6.ValueMember = "NAME";
            comboBox6.DisplayMember = "NAME";
            sqlConn.Close();


        }

        public void comboBox7load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT  [KIND]   FROM [TKMOC].[dbo].[MOCKIND] ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("KIND", typeof(string));

            da.Fill(dt);
            comboBox7.DataSource = dt.DefaultView;
            comboBox7.ValueMember = "KIND";
            comboBox7.DisplayMember = "KIND";
            sqlConn.Close();


        }

        public void comboBox10load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MD001,MD002 FROM [TK].dbo.CMSMD    WHERE MD002 LIKE '新廠%'  AND MD002 IN ('新廠製二組','新廠製一組') ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MD001", typeof(string));
            dt.Columns.Add("MD002", typeof(string));
            da.Fill(dt);
            comboBox10.DataSource = dt.DefaultView;
            comboBox10.ValueMember = "MD002";
            comboBox10.DisplayMember = "MD002";
            sqlConn.Close();


        }
        public void comboBox11load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MD001,MD002 FROM [TK].dbo.CMSMD    WHERE MD002 LIKE '新廠%'  AND MD002 IN ('新廠製二組','新廠製一組') ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MD001", typeof(string));
            dt.Columns.Add("MD002", typeof(string));
            da.Fill(dt);
            comboBox11.DataSource = dt.DefaultView;
            comboBox11.ValueMember = "MD002";
            comboBox11.DisplayMember = "MD002";
            sqlConn.Close();


        }

        public void comboBox8load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT TA034 ,MD002 FROM [TK].dbo.MOCTA ,[TK].dbo.CMSMD WHERE TA021=MD001  AND MD002='{0}' AND  TA003='{1}' ORDER BY TA034", comboBox11.Text,dateTimePicker15.Value.ToString("yyyyMMdd"));
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("TA034", typeof(string));
            
            da.Fill(dt);
            comboBox8.DataSource = dt.DefaultView;
            comboBox8.ValueMember = "TA034";
            comboBox8.DisplayMember = "TA034";
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
            SB.AppendFormat(" ,(CASE WHEN 領料是否扣袋重 IN ('Y') AND 成品是否扣袋重 IN ('Y') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000)-(袋重比*原料用量))>0 THEN (((生產量*淨重*(1-袋重比)))/1000) ELSE 0 END)+(CASE WHEN 領料是否扣袋重 IN ('Y') AND 成品是否扣袋重 IN ('N') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000))>0 THEN (((生產量*淨重))/1000) ELSE 0 END)+(CASE WHEN 領料是否扣袋重 IN ('N') AND 成品是否扣袋重 IN ('Y') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000)-(袋重比*原料用量))>0 THEN (((生產量*淨重*(1-袋重比)))/1000) ELSE 0 END)+(CASE WHEN 領料是否扣袋重 IN ('N') AND 成品是否扣袋重 IN ('N') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000))>0 THEN ((生產量*淨重)/1000) ELSE 0 END)+(CASE WHEN 類別 IN ('半成品') AND 原料用量>0  AND 成品是否扣袋重 IN ('Y') THEN (生產量-(生產量*袋重比)) ELSE 0 END )+(CASE WHEN 類別 IN ('半成品') AND 原料用量>0  AND 成品是否扣袋重 IN ('N') THEN (生產量) ELSE 0 END)+(CASE WHEN 類別 IN ('個','試吃') AND 原料用量>0 AND (原料用量*(1-(蒸發率/100)))>0 THEN (生產量*淨重/1000) ELSE 0 END)+(CASE WHEN 類別 IN ('片') AND 原料用量>0 AND (原料用量*(1-(蒸發率/100))-(原料用量*袋重比))>0 THEN (生產量*淨重/1000) ELSE 0 END)+(CASE WHEN 類別 IN ('單包') AND 原料用量>0 THEN 生產量  ELSE 0 END)+(CASE WHEN 類別 IN ('kg') AND (原料用量*(1-(蒸發率/100))-(原料用量*袋重比))>0 THEN ((生產量)) ELSE 0 END) AS '分子'");
            SB.AppendFormat(" ,(CASE WHEN 領料是否扣袋重 IN ('Y') AND 成品是否扣袋重 IN ('Y') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000)-(袋重比*原料用量))>0 THEN (原料用量*(1-(蒸發率/100))+(成品用量/1000)-(袋重比*原料用量)) ELSE 0 END)+(CASE WHEN 領料是否扣袋重 IN ('Y') AND 成品是否扣袋重 IN ('N') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000))>0 THEN (原料用量*(1-蒸發率)+(成品用量/1000)-(袋重比*原料用量)) ELSE 0 END)+(CASE WHEN 領料是否扣袋重 IN ('N') AND 成品是否扣袋重 IN ('Y') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000)-(袋重比*原料用量))>0 THEN (原料用量*(1-(蒸發率/100))+(成品用量/1000)) ELSE 0 END)+(CASE WHEN 領料是否扣袋重 IN ('N') AND 成品是否扣袋重 IN ('N') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000))>0 THEN ((原料用量*(1-(蒸發率/100))+(成品用量/1000))) ELSE 0 END)+(CASE WHEN 類別 IN ('半成品') AND 原料用量>0  AND 成品是否扣袋重 IN ('Y') THEN (原料用量*(1-蒸發率/100)) ELSE 0 END)+(CASE WHEN 類別 IN ('半成品') AND 原料用量>0  AND 成品是否扣袋重 IN ('N') THEN (原料用量*(1-蒸發率/100)) ELSE 0 END)+(CASE WHEN 類別 IN ('個','試吃') AND 原料用量>0 AND (原料用量*(1-(蒸發率/100)))>0 THEN (原料用量*(1-(蒸發率/100))) ELSE 0 END)+(CASE WHEN 類別 IN ('片') AND 原料用量>0 AND (原料用量*(1-(蒸發率/100))-(原料用量*袋重比))>0 THEN (原料用量*(1-(蒸發率/100))-(原料用量*袋重比)) ELSE 0 END)+(CASE WHEN 類別 IN ('單包') AND 原料用量>0 THEN 原料用量  ELSE 0 END)+(CASE WHEN 類別 IN ('kg') AND (原料用量*(1-(蒸發率/100))-(原料用量*袋重比))>0 THEN ((原料用量*(1-(蒸發率/100))+(成品用量/1000)-(原料用量*袋重比))) ELSE 0 END) AS '分母'");
            SB.AppendFormat(" ");
            SB.AppendFormat(" FROM(");
            SB.AppendFormat(" SELECT MD002 AS '線別',TA006 AS '品號',TA034 AS '品名',MB003 AS '規格' ,TA001 AS '製令單別',TA002 AS '製令單號',TA007 AS '生產單位',INVMB.UDF03 AS '類別',TA015 AS '預計產量',TA017 AS '生產量',INVMB.UDF07 AS '淨重',INVMB.UDF08 AS '單片重',INVMB.UDF09 AS '袋重',INVMB.UDF06 AS '蒸發率',INVMB.UDF01 AS '成品是否扣袋重',INVMB.UDF02 AS '領料是否扣袋重'");
            SB.AppendFormat(" ,(SELECT ISNULL(SUM(TB005),0) FROM [TK].dbo.MOCTB TB WHERE (TB.TB003 LIKE '1%' OR TB.TB003 LIKE '3%') AND TB.TB001=MOCTA.TA001 AND TB.TB002=MOCTA.TA002)  AS '原料用量'");
            SB.AppendFormat(" ,(SELECT ISNULL(SUM(TB005*MB.UDF07),0) FROM [TK].dbo.MOCTB TB,[TK].dbo.INVMB MB WHERE TB.TB003=MB.MB001 AND TB.TB003 LIKE '4%' AND TB.TB001=MOCTA.TA001 AND TB.TB002=MOCTA.TA002) AS '成品用量'");
            SB.AppendFormat(" ,CASE WHEN INVMB.UDF08>0 AND   INVMB.UDF09>0  THEN 1/(INVMB.UDF08+INVMB.UDF09)*INVMB.UDF09 ELSE 0 END  AS '袋重比'");
            SB.AppendFormat(" FROM [TK].dbo.INVMB,[TK].dbo.MOCTA,[TK].dbo.CMSMD");
            SB.AppendFormat(" WHERE TA006=MB001 AND TA021=MD001");
            SB.AppendFormat(" AND ISNULL(INVMB.UDF03,'')<>''");
            SB.AppendFormat(" AND TA003>='{0}' AND TA003<='{1}'",dateTimePicker1.Value.ToString("yyyyMMdd"),dateTimePicker2.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" ) AS TEMP");
            SB.AppendFormat(" WHERE 線別='{0}'",comboBox1.Text);
            SB.AppendFormat(" ORDER BY 線別,SUBSTRING(製令單號,1,8),品號");
            SB.AppendFormat("  ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");

            return SB;

        }


        public void SETNULL2()
        {
            textBox1.Text = "0";
            textBox2.Text = "0";
            textBox3.Text = "0";
            textBox4.Text = "0";


        }

        public void SETNULL3()
        {
            textBox5.Text = "0";
            textBox6.Text = "0";
            textBox7.Text = "0";
            textBox8.Text = "0";


        }
        public void SETNULL4()
        {
            textBox9.Text = "0";
            textBox10.Text = "0";
            textBox11.Text = "0";
            textBox12.Text = "0";
            textBox13.Text = "0";
            textBox14.Text = "0";
            textBox15.Text = "0";
            textBox16.Text = "0";
            textBox17.Text = "0";
            textBox18.Text = "0";
            textBox19.Text = "0";
            textBox20.Text = "0";


        }

        public void SETNULL5()
        {
            textBox21.Text = "0";
            textBox22.Text = "0";
            textBox23.Text = "0";
            textBox24.Text = "0";
            textBox25.Text = "0";
        }

        public void SETNULL6()
        {
            textBox26.Text = "0";
            textBox27.Text = "0";
            textBox28.Text = "0";
            textBox29.Text = "0";
            textBox30.Text = "0";
            textBox31.Text = "0";
            textBox32.Text = "0";
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

        public void ADDMOCDAILYRECORDNG2()
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
                sbSql.AppendFormat(" ([DATES],[MOCLINE],[NGRECYCLESIDE],[NGSIDE],[NG],[NGRECYCLE])");
                sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}')", dateTimePicker6.Value.ToString("yyyy/MM/dd"), comboBox4.Text, textBox5.Text, textBox6.Text, textBox7.Text, textBox8.Text);
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

        public void UPDATEMOCDAILYRECORDNG2()
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
                sbSql.AppendFormat(" SET [NGRECYCLESIDE]='{0}',[NGSIDE]='{1}',[NG]='{2}',[NGRECYCLE]='{3}'", textBox5.Text, textBox6.Text, textBox7.Text, textBox8.Text);
                sbSql.AppendFormat(" WHERE [DATES]='{0}' AND [MOCLINE]='{1}'", dateTimePicker6.Value.ToString("yyyy/MM/dd"), comboBox4.Text);
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

        public void ADDMOCDAILYRECORDNGMONEY()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[MOCDAILYRECORDNGMONEY] ");
                sbSql.AppendFormat(" ([DATES],[NGCLEAN],[NGBAT],[NGSELECT],[NGSIDE],[NGSIDENG],[NGCOOKNG],[NG1],[NG2],[NGCOOKIES],[NGSIDEMANU],[MGOTHERS],[REMARK]) ");
                sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}') ",dateTimePicker9.Value.ToString("yyyyMMdd"),textBox9.Text, textBox10.Text, textBox11.Text, textBox12.Text, textBox13.Text, textBox14.Text, textBox15.Text, textBox16.Text, textBox17.Text, textBox18.Text, textBox19.Text, textBox20.Text);
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
        public void UPDATEMOCDAILYRECORDNGMONEY()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" UPDATE [TKMOC].[dbo].[MOCDAILYRECORDNGMONEY]");
                sbSql.AppendFormat(" SET [NGCLEAN]='{0}',[NGBAT]='{1}',[NGSELECT]='{2}',[NGSIDE]='{3}',[NGSIDENG]='{4}',[NGCOOKNG]='{5}',[NG1]='{6}',[NG2]='{7}',[NGCOOKIES]='{8}',[NGSIDEMANU]='{9}',[MGOTHERS]='{10}',[REMARK]='{11}'", textBox9.Text, textBox10.Text, textBox11.Text, textBox12.Text, textBox13.Text, textBox14.Text, textBox15.Text, textBox16.Text, textBox17.Text, textBox18.Text, textBox19.Text, textBox20.Text);
                sbSql.AppendFormat(" WHERE [DATES]='{0}'",dateTimePicker9.Value.ToString("yyyyMMdd"));
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

        public void SEARCHMOCDAILYRECORDNG2()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@" SELECT CONVERT(NVARCHAR,[DATES],111) AS '日期',[MOCLINE] AS '線別',[NGRECYCLESIDE] AS '可回收邊料',[NGSIDE] AS '邊料報廢 (kg)',[NG] AS '不良報廢重 (kg)',[NGRECYCLE] AS '回收餅' ");
                sbSql.AppendFormat(@" FROM [TKMOC].[dbo].[MOCDAILYRECORDNG]");
                sbSql.AppendFormat(@" WHERE [MOCLINE]='{0}' AND CONVERT(NVARCHAR,[DATES],112) LIKE '{1}%' ", comboBox4.Text, dateTimePicker6.Value.ToString("yyyyMM"));
                sbSql.AppendFormat(@" ORDER BY CONVERT(NVARCHAR,[DATES],112)");
                sbSql.AppendFormat(@"  ");

                adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                sqlConn.Open();
                ds2.Clear();
                adapter2.Fill(ds2, "ds2");
                sqlConn.Close();


                if (ds2.Tables["ds2"].Rows.Count == 0)
                {

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

            //新廠製二組，每日加總不計成品
            //新廠製一組，每日加總計成品
            if (comboBox2.Text.Equals("新廠製二組"))
            {
                SB.AppendFormat(" SELECT 線別,SUBSTRING(製令單號,1,8) AS '日期'");
                SB.AppendFormat(" ,SUM(分子) AS '入庫淨重'");
                SB.AppendFormat(" ,SUM(分母) AS '製造投入量'");
                SB.AppendFormat(" ,CASE WHEN  SUM(領料扣成品扣的投入量+領料扣成品不扣的投入量+領料不扣成品扣的投入量+領料不扣成品不扣的投入量+半成品得料率成品扣袋重的投入量+半成品得料率成品不扣袋重的投入量+個試吃的投入量+片的投入量+單包的投入量+kg的投入量)>0 THEN SUM(領料扣成品扣的入庫量+領料扣成品不扣的入庫量+領料不扣成品扣的入庫量+領料不扣成品不扣的入庫量+半成品得料率成品扣袋重的入庫量+半成品得料率成品不扣袋重的入庫量+個試吃的入庫量+片的入庫量+單包的入庫量+kg的入庫量)/SUM(領料扣成品扣的投入量+領料扣成品不扣的投入量+領料不扣成品扣的投入量+領料不扣成品不扣的投入量+半成品得料率成品扣袋重的投入量+半成品得料率成品不扣袋重的投入量+個試吃的投入量+片的投入量+單包的投入量+kg的投入量) ELSE 0 END  AS '得料率(%)'");
                SB.AppendFormat(" ,ISNULL((SELECT [NGCOOK] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8)  AND MOCLINE=線別),0) AS '可回收-烘焙不良品 	'");
                SB.AppendFormat(" ,ISNULL((SELECT [NGCOOL] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8)  AND MOCLINE=線別),0) AS '打餅區落地-冷卻不良品'");
                SB.AppendFormat(" ,ISNULL((SELECT [NGPACKF] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8)  AND MOCLINE=線別),0) AS '前端-包裝不良品'");
                SB.AppendFormat(" ,ISNULL((SELECT [NGPACKB] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8)  AND MOCLINE=線別),0) AS '後端落地-包裝不良品'");
                SB.AppendFormat(" ,(ISNULL((SELECT [NGCOOK] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8)  AND MOCLINE=線別) ,0)+ISNULL((SELECT [NGCOOL] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8) AND MOCLINE=線別),0)+ISNULL((SELECT [NGPACKF] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8) AND MOCLINE=線別),0)+ISNULL((SELECT [NGPACKB] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8) AND MOCLINE=線別),0)) AS '不良合計'");
                SB.AppendFormat(" ,CASE WHEN  SUM(領料扣成品扣的投入量+領料扣成品不扣的投入量+領料不扣成品扣的投入量+領料不扣成品不扣的投入量+半成品得料率成品扣袋重的投入量+半成品得料率成品不扣袋重的投入量+個試吃的投入量+片的投入量+單包的投入量+kg的投入量)>0 THEN (ISNULL((SELECT [NGCOOK] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8) AND MOCLINE=線別),0)+ISNULL((SELECT [NGCOOL] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8) AND MOCLINE=線別),0)+ISNULL((SELECT [NGPACKF] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8) AND MOCLINE=線別),0)+ISNULL((SELECT [NGPACKB] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8) AND MOCLINE=線別),0))/SUM(領料扣成品扣的投入量+領料扣成品不扣的投入量+領料不扣成品扣的投入量+領料不扣成品不扣的投入量+半成品得料率成品扣袋重的投入量+半成品得料率成品不扣袋重的投入量+個試吃的投入量+片的投入量+單包的投入量+kg的投入量) ELSE 0 END  AS '報廢率(％)'");
                SB.AppendFormat("  ");
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
                SB.AppendFormat(" ,(CASE WHEN 領料是否扣袋重 IN ('Y') AND 成品是否扣袋重 IN ('Y') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000)-(袋重比*原料用量))>0 THEN (((生產量*淨重*(1-袋重比)))/1000) ELSE 0 END)+(CASE WHEN 領料是否扣袋重 IN ('Y') AND 成品是否扣袋重 IN ('N') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000))>0 THEN (((生產量*淨重))/1000) ELSE 0 END)+(CASE WHEN 領料是否扣袋重 IN ('N') AND 成品是否扣袋重 IN ('Y') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000)-(袋重比*原料用量))>0 THEN (((生產量*淨重*(1-袋重比)))/1000) ELSE 0 END)+(CASE WHEN 領料是否扣袋重 IN ('N') AND 成品是否扣袋重 IN ('N') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000))>0 THEN ((生產量*淨重)/1000) ELSE 0 END)+(CASE WHEN 類別 IN ('半成品') AND 原料用量>0  AND 成品是否扣袋重 IN ('Y') THEN (生產量-(生產量*袋重比)) ELSE 0 END )+(CASE WHEN 類別 IN ('半成品') AND 原料用量>0  AND 成品是否扣袋重 IN ('N') THEN (生產量) ELSE 0 END)+(CASE WHEN 類別 IN ('個','試吃') AND 原料用量>0 AND (原料用量*(1-(蒸發率/100)))>0 THEN (生產量*淨重/1000) ELSE 0 END)+(CASE WHEN 類別 IN ('片') AND 原料用量>0 AND (原料用量*(1-(蒸發率/100))-(原料用量*袋重比))>0 THEN (生產量*淨重/1000) ELSE 0 END)+(CASE WHEN 類別 IN ('單包') AND 原料用量>0 THEN 生產量  ELSE 0 END)+(CASE WHEN 類別 IN ('kg') AND (原料用量*(1-(蒸發率/100))-(原料用量*袋重比))>0 THEN ((生產量)) ELSE 0 END) AS '分子'");
                SB.AppendFormat(" ,(CASE WHEN 領料是否扣袋重 IN ('Y') AND 成品是否扣袋重 IN ('Y') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000)-(袋重比*原料用量))>0 THEN (原料用量*(1-(蒸發率/100))+(成品用量/1000)-(袋重比*原料用量)) ELSE 0 END)+(CASE WHEN 領料是否扣袋重 IN ('Y') AND 成品是否扣袋重 IN ('N') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000))>0 THEN (原料用量*(1-蒸發率)+(成品用量/1000)-(袋重比*原料用量)) ELSE 0 END)+(CASE WHEN 領料是否扣袋重 IN ('N') AND 成品是否扣袋重 IN ('Y') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000)-(袋重比*原料用量))>0 THEN (原料用量*(1-(蒸發率/100))+(成品用量/1000)) ELSE 0 END)+(CASE WHEN 領料是否扣袋重 IN ('N') AND 成品是否扣袋重 IN ('N') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000))>0 THEN ((原料用量*(1-(蒸發率/100))+(成品用量/1000))) ELSE 0 END)+(CASE WHEN 類別 IN ('半成品') AND 原料用量>0  AND 成品是否扣袋重 IN ('Y') THEN (原料用量*(1-蒸發率/100)) ELSE 0 END)+(CASE WHEN 類別 IN ('半成品') AND 原料用量>0  AND 成品是否扣袋重 IN ('N') THEN (原料用量*(1-蒸發率/100)) ELSE 0 END)+(CASE WHEN 類別 IN ('個','試吃') AND 原料用量>0 AND (原料用量*(1-(蒸發率/100)))>0 THEN (原料用量*(1-(蒸發率/100))) ELSE 0 END)+(CASE WHEN 類別 IN ('片') AND 原料用量>0 AND (原料用量*(1-(蒸發率/100))-(原料用量*袋重比))>0 THEN (原料用量*(1-(蒸發率/100))-(原料用量*袋重比)) ELSE 0 END)+(CASE WHEN 類別 IN ('單包') AND 原料用量>0 THEN 原料用量  ELSE 0 END)+(CASE WHEN 類別 IN ('kg') AND (原料用量*(1-(蒸發率/100))-(原料用量*袋重比))>0 THEN ((原料用量*(1-(蒸發率/100))+(成品用量/1000)-(原料用量*袋重比))) ELSE 0 END) AS '分母'");
                SB.AppendFormat(" ");
                SB.AppendFormat(" FROM(");
                SB.AppendFormat(" SELECT MD002 AS '線別',TA006 AS '品號',TA034 AS '品名',TA001 AS '製令單別',TA002 AS '製令單號',TA007 AS '生產單位',INVMB.UDF03 AS '類別',TA017 AS '生產量',INVMB.UDF07 AS '淨重',INVMB.UDF08 AS '單片重',INVMB.UDF09 AS '袋重',INVMB.UDF06 AS '蒸發率',INVMB.UDF01 AS '成品是否扣袋重',INVMB.UDF02 AS '領料是否扣袋重'");
                SB.AppendFormat(" ,(SELECT ISNULL(SUM(TB005),0) FROM [TK].dbo.MOCTB TB WHERE (TB.TB003 LIKE '1%' OR TB.TB003 LIKE '3%') AND TB.TB001=MOCTA.TA001 AND TB.TB002=MOCTA.TA002)  AS '原料用量'");
                SB.AppendFormat(" ,(SELECT ISNULL(SUM(TB005*MB.UDF07),0) FROM [TK].dbo.MOCTB TB,[TK].dbo.INVMB MB WHERE TB.TB003=MB.MB001 AND TB.TB003 LIKE '4%' AND TB.TB001=MOCTA.TA001 AND TB.TB002=MOCTA.TA002) AS '成品用量'");
                SB.AppendFormat(" ,CASE WHEN INVMB.UDF08>0 AND   INVMB.UDF09>0  THEN 1/(INVMB.UDF08+INVMB.UDF09)*INVMB.UDF09 ELSE 1 END  AS '袋重比'");
                SB.AppendFormat(" FROM [TK].dbo.INVMB,[TK].dbo.MOCTA,[TK].dbo.CMSMD");
                SB.AppendFormat(" WHERE TA006=MB001 AND TA021=MD001");
                SB.AppendFormat(" AND ISNULL(INVMB.UDF03,'')<>''");
                SB.AppendFormat(" AND TA003>='{0}' AND TA003<='{1}'", dateTimePicker4.Value.ToString("yyyyMMdd"), dateTimePicker5.Value.ToString("yyyyMMdd"));
                SB.AppendFormat(" AND TA006 LIKE '3%'");
                SB.AppendFormat(" ) AS TEMP");
                SB.AppendFormat(" ) AS TEMP2");
                SB.AppendFormat(" WHERE 線別='{0}'", comboBox2.Text);
                SB.AppendFormat(" GROUP BY 線別,SUBSTRING(製令單號,1,8)");
                SB.AppendFormat(" ORDER BY 線別,SUBSTRING(製令單號,1,8)");
                SB.AppendFormat(" ");
                SB.AppendFormat(" ");

            }
            else if (comboBox2.Text.Equals("新廠製一組"))
            {
                SB.AppendFormat(" SELECT 線別,SUBSTRING(製令單號,1,8) AS '日期'");
                SB.AppendFormat(" ,SUM(分子) AS '入庫淨重'");
                SB.AppendFormat(" ,SUM(分母) AS '製造投入量'");
                SB.AppendFormat(" ,CASE WHEN  SUM(領料扣成品扣的投入量+領料扣成品不扣的投入量+領料不扣成品扣的投入量+領料不扣成品不扣的投入量+半成品得料率成品扣袋重的投入量+半成品得料率成品不扣袋重的投入量+個試吃的投入量+片的投入量+單包的投入量+kg的投入量)>0 THEN SUM(領料扣成品扣的入庫量+領料扣成品不扣的入庫量+領料不扣成品扣的入庫量+領料不扣成品不扣的入庫量+半成品得料率成品扣袋重的入庫量+半成品得料率成品不扣袋重的入庫量+個試吃的入庫量+片的入庫量+單包的入庫量+kg的入庫量)/SUM(領料扣成品扣的投入量+領料扣成品不扣的投入量+領料不扣成品扣的投入量+領料不扣成品不扣的投入量+半成品得料率成品扣袋重的投入量+半成品得料率成品不扣袋重的投入量+個試吃的投入量+片的投入量+單包的投入量+kg的投入量) ELSE 0 END  AS '得料率(%)'");
                SB.AppendFormat(" ,ISNULL((SELECT [NGCOOK] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8)  AND MOCLINE=線別),0) AS '可回收-烘焙不良品 	'");
                SB.AppendFormat(" ,ISNULL((SELECT [NGCOOL] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8)  AND MOCLINE=線別),0) AS '打餅區落地-冷卻不良品'");
                SB.AppendFormat(" ,ISNULL((SELECT [NGPACKF] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8)  AND MOCLINE=線別),0) AS '前端-包裝不良品'");
                SB.AppendFormat(" ,ISNULL((SELECT [NGPACKB] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8)  AND MOCLINE=線別),0) AS '後端落地-包裝不良品'");
                SB.AppendFormat(" ,(ISNULL((SELECT [NGCOOK] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8)  AND MOCLINE=線別) ,0)+ISNULL((SELECT [NGCOOL] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8) AND MOCLINE=線別),0)+ISNULL((SELECT [NGPACKF] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8) AND MOCLINE=線別),0)+ISNULL((SELECT [NGPACKB] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8) AND MOCLINE=線別),0)) AS '不良合計'");
                SB.AppendFormat(" ,CASE WHEN  SUM(領料扣成品扣的投入量+領料扣成品不扣的投入量+領料不扣成品扣的投入量+領料不扣成品不扣的投入量+半成品得料率成品扣袋重的投入量+半成品得料率成品不扣袋重的投入量+個試吃的投入量+片的投入量+單包的投入量+kg的投入量)>0 THEN (ISNULL((SELECT [NGCOOK] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8) AND MOCLINE=線別),0)+ISNULL((SELECT [NGCOOL] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8) AND MOCLINE=線別),0)+ISNULL((SELECT [NGPACKF] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8) AND MOCLINE=線別),0)+ISNULL((SELECT [NGPACKB] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8) AND MOCLINE=線別),0))/SUM(領料扣成品扣的投入量+領料扣成品不扣的投入量+領料不扣成品扣的投入量+領料不扣成品不扣的投入量+半成品得料率成品扣袋重的投入量+半成品得料率成品不扣袋重的投入量+個試吃的投入量+片的投入量+單包的投入量+kg的投入量) ELSE 0 END  AS '報廢率(％)'");
                SB.AppendFormat("  ");
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
                SB.AppendFormat(" ,(CASE WHEN 領料是否扣袋重 IN ('Y') AND 成品是否扣袋重 IN ('Y') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000)-(袋重比*原料用量))>0 THEN (((生產量*淨重*(1-袋重比)))/1000) ELSE 0 END)+(CASE WHEN 領料是否扣袋重 IN ('Y') AND 成品是否扣袋重 IN ('N') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000))>0 THEN (((生產量*淨重))/1000) ELSE 0 END)+(CASE WHEN 領料是否扣袋重 IN ('N') AND 成品是否扣袋重 IN ('Y') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000)-(袋重比*原料用量))>0 THEN (((生產量*淨重*(1-袋重比)))/1000) ELSE 0 END)+(CASE WHEN 領料是否扣袋重 IN ('N') AND 成品是否扣袋重 IN ('N') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000))>0 THEN ((生產量*淨重)/1000) ELSE 0 END)+(CASE WHEN 類別 IN ('半成品') AND 原料用量>0  AND 成品是否扣袋重 IN ('Y') THEN (生產量-(生產量*袋重比)) ELSE 0 END )+(CASE WHEN 類別 IN ('半成品') AND 原料用量>0  AND 成品是否扣袋重 IN ('N') THEN (生產量) ELSE 0 END)+(CASE WHEN 類別 IN ('個','試吃') AND 原料用量>0 AND (原料用量*(1-(蒸發率/100)))>0 THEN (生產量*淨重/1000) ELSE 0 END)+(CASE WHEN 類別 IN ('片') AND 原料用量>0 AND (原料用量*(1-(蒸發率/100))-(原料用量*袋重比))>0 THEN (生產量*淨重/1000) ELSE 0 END)+(CASE WHEN 類別 IN ('單包') AND 原料用量>0 THEN 生產量  ELSE 0 END)+(CASE WHEN 類別 IN ('kg') AND (原料用量*(1-(蒸發率/100))-(原料用量*袋重比))>0 THEN ((生產量)) ELSE 0 END) AS '分子'");
                SB.AppendFormat(" ,(CASE WHEN 領料是否扣袋重 IN ('Y') AND 成品是否扣袋重 IN ('Y') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000)-(袋重比*原料用量))>0 THEN (原料用量*(1-(蒸發率/100))+(成品用量/1000)-(袋重比*原料用量)) ELSE 0 END)+(CASE WHEN 領料是否扣袋重 IN ('Y') AND 成品是否扣袋重 IN ('N') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000))>0 THEN (原料用量*(1-蒸發率)+(成品用量/1000)-(袋重比*原料用量)) ELSE 0 END)+(CASE WHEN 領料是否扣袋重 IN ('N') AND 成品是否扣袋重 IN ('Y') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000)-(袋重比*原料用量))>0 THEN (原料用量*(1-(蒸發率/100))+(成品用量/1000)) ELSE 0 END)+(CASE WHEN 領料是否扣袋重 IN ('N') AND 成品是否扣袋重 IN ('N') AND  類別 NOT IN ('半成品','個','試吃','片','單包','kg') AND (原料用量*(1-(蒸發率/100))+(成品用量/1000))>0 THEN ((原料用量*(1-(蒸發率/100))+(成品用量/1000))) ELSE 0 END)+(CASE WHEN 類別 IN ('半成品') AND 原料用量>0  AND 成品是否扣袋重 IN ('Y') THEN (原料用量*(1-蒸發率/100)) ELSE 0 END)+(CASE WHEN 類別 IN ('半成品') AND 原料用量>0  AND 成品是否扣袋重 IN ('N') THEN (原料用量*(1-蒸發率/100)) ELSE 0 END)+(CASE WHEN 類別 IN ('個','試吃') AND 原料用量>0 AND (原料用量*(1-(蒸發率/100)))>0 THEN (原料用量*(1-(蒸發率/100))) ELSE 0 END)+(CASE WHEN 類別 IN ('片') AND 原料用量>0 AND (原料用量*(1-(蒸發率/100))-(原料用量*袋重比))>0 THEN (原料用量*(1-(蒸發率/100))-(原料用量*袋重比)) ELSE 0 END)+(CASE WHEN 類別 IN ('單包') AND 原料用量>0 THEN 原料用量  ELSE 0 END)+(CASE WHEN 類別 IN ('kg') AND (原料用量*(1-(蒸發率/100))-(原料用量*袋重比))>0 THEN ((原料用量*(1-(蒸發率/100))+(成品用量/1000)-(原料用量*袋重比))) ELSE 0 END) AS '分母'");
                SB.AppendFormat(" ");
                SB.AppendFormat(" FROM(");
                SB.AppendFormat(" SELECT MD002 AS '線別',TA006 AS '品號',TA034 AS '品名',TA001 AS '製令單別',TA002 AS '製令單號',TA007 AS '生產單位',INVMB.UDF03 AS '類別',TA017 AS '生產量',INVMB.UDF07 AS '淨重',INVMB.UDF08 AS '單片重',INVMB.UDF09 AS '袋重',INVMB.UDF06 AS '蒸發率',INVMB.UDF01 AS '成品是否扣袋重',INVMB.UDF02 AS '領料是否扣袋重'");
                SB.AppendFormat(" ,(SELECT ISNULL(SUM(TB005),0) FROM [TK].dbo.MOCTB TB WHERE (TB.TB003 LIKE '1%' OR TB.TB003 LIKE '3%') AND TB.TB001=MOCTA.TA001 AND TB.TB002=MOCTA.TA002)  AS '原料用量'");
                SB.AppendFormat(" ,(SELECT ISNULL(SUM(TB005*MB.UDF07),0) FROM [TK].dbo.MOCTB TB,[TK].dbo.INVMB MB WHERE TB.TB003=MB.MB001 AND TB.TB003 LIKE '4%' AND TB.TB001=MOCTA.TA001 AND TB.TB002=MOCTA.TA002) AS '成品用量'");
                SB.AppendFormat(" ,CASE WHEN INVMB.UDF08>0 AND   INVMB.UDF09>0  THEN 1/(INVMB.UDF08+INVMB.UDF09)*INVMB.UDF09 ELSE 1 END  AS '袋重比'");
                SB.AppendFormat(" FROM [TK].dbo.INVMB,[TK].dbo.MOCTA,[TK].dbo.CMSMD");
                SB.AppendFormat(" WHERE TA006=MB001 AND TA021=MD001");
                SB.AppendFormat(" AND ISNULL(INVMB.UDF03,'')<>''");
                SB.AppendFormat(" AND TA003>='{0}' AND TA003<='{1}'", dateTimePicker4.Value.ToString("yyyyMMdd"), dateTimePicker5.Value.ToString("yyyyMMdd"));
                SB.AppendFormat(" ) AS TEMP");
                SB.AppendFormat(" ) AS TEMP2");
                SB.AppendFormat(" WHERE 線別='{0}'", comboBox2.Text);
                SB.AppendFormat(" GROUP BY 線別,SUBSTRING(製令單號,1,8)");
                SB.AppendFormat(" ORDER BY 線別,SUBSTRING(製令單號,1,8)");
                SB.AppendFormat(" ");
                SB.AppendFormat(" ");
            }
           
           
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");


            return SB;

        }

        public void SETFASTREPORT3()
        {
            StringBuilder SQL1 = new StringBuilder();
            StringBuilder SQL2 = new StringBuilder();

            SQL1 = SETSQL3();

            Report report3 = new Report();
            report3.Load(@"REPORT\生產報表-每日得料率報表-手工.frx");

            report3.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            TableDataSource table = report3.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            report3.Preview = previewControl3;
            report3.Show();
        }

        public StringBuilder SETSQL3()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(" SELECT 線別,SUBSTRING(製令單號,1,8) AS '日期'");
            SB.AppendFormat(" ,SUM(領料扣成品扣的投入量+領料扣成品不扣的投入量+領料不扣成品扣的投入量+領料不扣成品不扣的投入量+半成品得料率成品扣袋重的投入量+半成品得料率成品不扣袋重的投入量+個試吃的投入量+片的投入量+單包的投入量+kg的投入量) AS '應產出量 '");
            SB.AppendFormat(" ,SUM(領料扣成品扣的入庫量+領料扣成品不扣的入庫量+領料不扣成品扣的入庫量+領料不扣成品不扣的入庫量+半成品得料率成品扣袋重的入庫量+半成品得料率成品不扣袋重的入庫量+個試吃的入庫量+片的入庫量+單包的入庫量+kg的入庫量) AS '入庫淨重'");
            SB.AppendFormat(" ,CASE WHEN  SUM(領料扣成品扣的投入量+領料扣成品不扣的投入量+領料不扣成品扣的投入量+領料不扣成品不扣的投入量+半成品得料率成品扣袋重的投入量+半成品得料率成品不扣袋重的投入量+個試吃的投入量+片的投入量+單包的投入量+kg的投入量)>0 THEN SUM(領料扣成品扣的入庫量+領料扣成品不扣的入庫量+領料不扣成品扣的入庫量+領料不扣成品不扣的入庫量+半成品得料率成品扣袋重的入庫量+半成品得料率成品不扣袋重的入庫量+個試吃的入庫量+片的入庫量+單包的入庫量+kg的入庫量)/SUM(領料扣成品扣的投入量+領料扣成品不扣的投入量+領料不扣成品扣的投入量+領料不扣成品不扣的投入量+半成品得料率成品扣袋重的投入量+半成品得料率成品不扣袋重的投入量+個試吃的投入量+片的投入量+單包的投入量+kg的投入量) ELSE 0 END  AS '得料率(%)'");
            SB.AppendFormat(" ,ISNULL((SELECT [NGRECYCLESIDE] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8)  AND MOCLINE=線別),0) AS '可回收邊料'");
            SB.AppendFormat(" ,ISNULL((SELECT [NGSIDE] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8)  AND MOCLINE=線別),0) AS '邊料報廢 (kg)'");
            SB.AppendFormat(" ,ISNULL((SELECT [NG] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8)  AND MOCLINE=線別),0) AS '不良報廢重 (kg)'");
            SB.AppendFormat(" ,ISNULL((SELECT [NGRECYCLE]  FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8)  AND MOCLINE=線別),0) AS '回收餅'");
            SB.AppendFormat(" ,(ISNULL((SELECT [NGRECYCLESIDE] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8)  AND MOCLINE=線別),0))+(ISNULL((SELECT [NGSIDE] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8)  AND MOCLINE=線別),0))+(ISNULL((SELECT [NG] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8)  AND MOCLINE=線別),0))+(ISNULL((SELECT [NGRECYCLE]  FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8)  AND MOCLINE=線別),0)) AS '不良合計'");
            SB.AppendFormat(" ,CASE WHEN  SUM(領料扣成品扣的投入量+領料扣成品不扣的投入量+領料不扣成品扣的投入量+領料不扣成品不扣的投入量+半成品得料率成品扣袋重的投入量+半成品得料率成品不扣袋重的投入量+個試吃的投入量+片的投入量+單包的投入量+kg的投入量)>0 THEN ((ISNULL((SELECT [NGRECYCLESIDE] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8)  AND MOCLINE=線別),0))+(ISNULL((SELECT [NGSIDE] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8)  AND MOCLINE=線別),0))+(ISNULL((SELECT [NG] FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8)  AND MOCLINE=線別),0))+(ISNULL((SELECT [NGRECYCLE]  FROM  [TKMOC].[dbo].[MOCDAILYRECORDNG] WHERE CONVERT(nvarchar,[DATES],112)=SUBSTRING(製令單號,1,8)  AND MOCLINE=線別),0)))/SUM(領料扣成品扣的投入量+領料扣成品不扣的投入量+領料不扣成品扣的投入量+領料不扣成品不扣的投入量+半成品得料率成品扣袋重的投入量+半成品得料率成品不扣袋重的投入量+個試吃的投入量+片的投入量+單包的投入量+kg的投入量) ELSE 0 END  AS '報廢率(％)'");
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
            SB.AppendFormat(" SELECT MD002 AS '線別',TA006 AS '品號',TA034 AS '品名',TA001 AS '製令單別',TA002 AS '製令單號',TA007 AS '生產單位',INVMB.UDF03 AS '類別',TA017 AS '生產量',INVMB.UDF07 AS '淨重',INVMB.UDF08 AS '單片重',INVMB.UDF09 AS '袋重',INVMB.UDF06 AS '蒸發率',MB112 AS '成品是否扣袋重',MB113 AS '領料是否扣袋重'");
            SB.AppendFormat(" ,(SELECT ISNULL(SUM(TB005),0) FROM [TK].dbo.MOCTB TB WHERE (TB.TB003 LIKE '1%' OR TB.TB003 LIKE '3%') AND TB.TB001=MOCTA.TA001 AND TB.TB002=MOCTA.TA002)  AS '原料用量'");
            SB.AppendFormat(" ,(SELECT ISNULL(SUM(TB005*MB.UDF07),0) FROM [TK].dbo.MOCTB TB,[TK].dbo.INVMB MB WHERE TB.TB003=MB.MB001 AND TB.TB003 LIKE '4%' AND TB.TB001=MOCTA.TA001 AND TB.TB002=MOCTA.TA002) AS '成品用量'");
            SB.AppendFormat(" ,CASE WHEN INVMB.UDF08>0 AND   INVMB.UDF09>0  THEN 1/(INVMB.UDF08+INVMB.UDF09)*INVMB.UDF09 ELSE 1 END  AS '袋重比'");
            SB.AppendFormat(" FROM [TK].dbo.INVMB,[TK].dbo.MOCTA,[TK].dbo.CMSMD");
            SB.AppendFormat(" WHERE TA006=MB001 AND TA021=MD001");
            SB.AppendFormat(" AND ISNULL(INVMB.UDF03,'')<>''");
            SB.AppendFormat(" AND TA003>='{0}' AND TA003<='{1}'", dateTimePicker7.Value.ToString("yyyyMMdd"), dateTimePicker8.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" ) AS TEMP");
            SB.AppendFormat(" ) AS TEMP2");
            SB.AppendFormat(" WHERE 線別='{0}'", comboBox5.Text);
            SB.AppendFormat(" GROUP BY 線別,SUBSTRING(製令單號,1,8)");
            SB.AppendFormat(" ORDER BY 線別,SUBSTRING(製令單號,1,8)");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");


            return SB;

        }

        public void SETFASTREPORT4()
        {
            StringBuilder SQL1 = new StringBuilder();
            StringBuilder SQL2 = new StringBuilder();

            SQL1 = SETSQL4();

            Report report4 = new Report();
            report4.Load(@"REPORT\生產報表-每日得料率報表-報廢.frx");

            report4.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            TableDataSource table = report4.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            report4.Preview = previewControl4;
            report4.Show();
        }

        public StringBuilder SETSQL4()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(" SELECT ");
            SB.AppendFormat(" CONVERT(NVARCHAR,[DATES],112) AS '日期',[NGCLEAN] AS '製造前端-打掃',[NGBAT] AS '製造前端-打餅區',[NGSELECT] AS '製造前端-篩選餅乾不良',[NGSIDE] AS '當日-邊料',[NGSIDENG] AS '過期-邊料報廢',[NGCOOKNG] AS '過期-餅麩報廢',[NG1] AS '製造後端-大線',[NG2] AS '製造後端-小線',[NGCOOKIES] AS '手工-廢餅',[NGSIDEMANU] AS '手工-邊料',[MGOTHERS] AS '其他/品保'");
            SB.AppendFormat(" ,([NGCLEAN]+[NGBAT]+[NGSELECT]+[NGSIDE]+[NGSIDENG]+[NGCOOKNG]+[NG1]+[NG2]+[NGCOOKIES]+[NGSIDEMANU]+[MGOTHERS]) AS '小計'");
            SB.AppendFormat(" ,[REMARK] AS '備註'");
            SB.AppendFormat(" ,(SELECT ISNULL(SUM(TB005),0) FROM [TK].dbo.MOCTB TB ,[TK].dbo.MOCTA TA WHERE TA.TA001=TB.TB001 AND TA.TA002=TB.TB002 AND  (TB.TB003 LIKE '1%' OR TB.TB003 LIKE '3%') AND TA.TA021 IN ('02','03') AND TA.TA012= CONVERT(NVARCHAR,[DATES],112))  AS '原料用量'");
            SB.AppendFormat(" ,(SELECT ISNULL(SUM(TB005*MB.UDF07),0) FROM [TK].dbo.MOCTB TB ,[TK].dbo.MOCTA TA,[TK].dbo.INVMB MB   WHERE TA.TA001=TB.TB001 AND TA.TA002=TB.TB002 AND TB.TB003=MB.MB001 AND TB.TB003 LIKE '4%' AND TA.TA021 IN ('02','03') AND TA.TA012= CONVERT(NVARCHAR,[DATES],112)) AS '成品用量'");
            SB.AppendFormat(" FROM [TKMOC].[dbo].[MOCDAILYRECORDNGMONEY]");
            SB.AppendFormat(" WHERE CONVERT(NVARCHAR,[DATES],112)='{0}'  ",dateTimePicker9.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" ORDER BY  CONVERT(NVARCHAR,[DATES],112)");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");

            return SB;

        }

        public void SETFASTREPORT5()
        {
            StringBuilder SQL1 = new StringBuilder();
            StringBuilder SQL2 = new StringBuilder();

            SQL1 = SETSQL5();

            Report report5 = new Report();
            report5.Load(@"REPORT\生產報表-每日得料率報表-報廢.frx");

            report5.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            TableDataSource table = report5.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            report5.Preview = previewControl4;
            report5.Show();
        }

        public StringBuilder SETSQL5()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(" SELECT ");
            SB.AppendFormat(" CONVERT(NVARCHAR,[DATES],112) AS '日期'");
            SB.AppendFormat(" ,[NGCLEAN] AS '製造前端-打掃'");
            SB.AppendFormat(" ,[NGBAT] AS '製造前端-打餅區'");
            SB.AppendFormat(" ,[NGSELECT] AS '製造前端-篩選餅乾不良'");
            SB.AppendFormat(" ,[NGSIDE] AS '當日-邊料'");
            SB.AppendFormat(" ,[NGSIDENG] AS '過期-邊料報廢'");
            SB.AppendFormat(" ,[NGCOOKNG] AS '過期-餅麩報廢'");
            SB.AppendFormat(" ,[NG1] AS '製造後端-大線'");
            SB.AppendFormat(" ,[NG2] AS '製造後端-小線'");
            SB.AppendFormat(" ,[NGCOOKIES] AS '手工-廢餅'");
            SB.AppendFormat(" ,[NGSIDEMANU] AS '手工-邊料'");
            SB.AppendFormat(" ,[MGOTHERS] AS '其他/品保'");
            SB.AppendFormat(" ,([NGCLEAN]+[NGBAT]+[NGSELECT]+[NGSIDE]+[NGSIDENG]+[NGCOOKNG]+[NG1]+[NG2]+[NGCOOKIES]+[NGSIDEMANU]+[MGOTHERS]) AS '小計'");
            SB.AppendFormat(" ,([NGCLEAN]+[NGBAT]+[NGSELECT]+[NGSIDE]+[NG1]+[NG2]+[NGCOOKIES]+[NGSIDEMANU]) AS '報廢總數'");
            SB.AppendFormat(" ,[NGCLEAN] AS '打掃報廢'");
            SB.AppendFormat(" ,([NGCLEAN]+[NGBAT]+[NGSELECT]+[NGSIDE]+[NGSIDENG]+[NGCOOKNG]+[NG1]+[NG2]+[NGCOOKIES]+[NGSIDEMANU]+[MGOTHERS]) AS '乾+濕'");
            SB.AppendFormat(" ,[NGSIDENG] AS '邊料報廢'");
            SB.AppendFormat(" ,[NGCOOKNG] AS '餅麩報廢'");
            SB.AppendFormat(" ,[NGCLEAN]*16.36 AS '打掃報廢-成本'");
            SB.AppendFormat(" ,([NGCLEAN]+[NGBAT]+[NGSELECT]+[NGSIDE]+[NGSIDENG]+[NGCOOKNG]+[NG1]+[NG2]+[NGCOOKIES]+[NGSIDEMANU]+[MGOTHERS])*54.01 AS '乾+濕成本'");
            SB.AppendFormat(" ,[REMARK] AS '備註'");
            SB.AppendFormat(" ,(SELECT ISNULL(SUM(TB005),0) FROM [TK].dbo.MOCTB TB ,[TK].dbo.MOCTA TA WHERE TA.TA001=TB.TB001 AND TA.TA002=TB.TB002 AND  (TB.TB003 LIKE '1%' OR TB.TB003 LIKE '3%') AND TA.TA021 IN ('02','03') AND TA.TA012= CONVERT(NVARCHAR,[DATES],112))  AS '原料用量'");
            SB.AppendFormat(" ,(SELECT ISNULL(SUM(TB005*MB.UDF07),0) FROM [TK].dbo.MOCTB TB ,[TK].dbo.MOCTA TA,[TK].dbo.INVMB MB   WHERE TA.TA001=TB.TB001 AND TA.TA002=TB.TB002 AND TB.TB003=MB.MB001 AND TB.TB003 LIKE '4%' AND TA.TA021 IN ('02','03') AND TA.TA012= CONVERT(NVARCHAR,[DATES],112)) AS '成品用量'");
            SB.AppendFormat(" FROM [TKMOC].[dbo].[MOCDAILYRECORDNGMONEY]");
            SB.AppendFormat(" WHERE CONVERT(NVARCHAR,[DATES],112)>='{0}' AND CONVERT(NVARCHAR,[DATES],112)<='{1}' ", dateTimePicker11.Value.ToString("yyyyMMdd"), dateTimePicker12.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" ORDER BY  CONVERT(NVARCHAR,[DATES],112)");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");

            return SB;

        }

        public void SETFASTREPORT6()
        {
            StringBuilder SQL1 = new StringBuilder();
            StringBuilder SQL2 = new StringBuilder();

            SQL1 = SETSQL6();

            Report report6 = new Report();
            report6.Load(@"REPORT\生產報表-每日得料率報表-桶數報廢.frx");

            report6.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            TableDataSource table = report6.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            report6.Preview = previewControl5;
            report6.Show();
        }

        public StringBuilder SETSQL6()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(" SELECT [PROD] AS '口味',CONVERT(NVARCHAR, [DATES],112) AS '日期',[LASTSLOT] AS '前日庫存',[PRODOUT] AS '當日產出',[PRODIN] AS '當日投入',[NG] AS '當日報廢',[NOWSLOT] AS '當日庫存'");
            SB.AppendFormat(" FROM [TKMOC].[dbo].[MOCDAILYSLOT]");
            SB.AppendFormat(" WHERE CONVERT(NVARCHAR, [DATES],112)='{0}'  ",dateTimePicker10.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" AND [KIND]='{0}'", comboBox7.Text);
            SB.AppendFormat(" AND [PROD] ='{0}'",comboBox6.Text);          
            SB.AppendFormat(" ORDER BY CONVERT(NVARCHAR, [DATES],112)");
            SB.AppendFormat(" ");

            return SB;

        }

        public void SETFASTREPORT7()
        {
            StringBuilder SQL1 = new StringBuilder();
            StringBuilder SQL2 = new StringBuilder();

            SQL1 = SETSQL7();

            Report report7 = new Report();
            report7.Load(@"REPORT\生產報表-每日得料率報表-桶數報廢.frx");

            report7.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            TableDataSource table = report7.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            report7.Preview = previewControl5;
            report7.Show();
        }

        public StringBuilder SETSQL7()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(" SELECT [PROD] AS '口味',CONVERT(NVARCHAR, [DATES],112) AS '日期',[LASTSLOT] AS '前日庫存',[PRODOUT] AS '當日產出',[PRODIN] AS '當日投入',[NG] AS '當日報廢',[NOWSLOT] AS '當日庫存'");
            SB.AppendFormat(" FROM [TKMOC].[dbo].[MOCDAILYSLOT]");
            SB.AppendFormat(" WHERE CONVERT(NVARCHAR, [DATES],112)>='{0}' AND  CONVERT(NVARCHAR, [DATES],112)<='{1}'  ", dateTimePicker13.Value.ToString("yyyyMMdd"), dateTimePicker14.Value.ToString("yyyyMMdd"));
            //SB.AppendFormat(" AND [PROD] ='{0}'", comboBox7.Text);
            SB.AppendFormat(" ORDER BY CONVERT(NVARCHAR, [DATES],112),[PROD]");
            SB.AppendFormat(" ");

            return SB;

        }

        public void SETFASTREPORT8()
        {
            StringBuilder SQL1 = new StringBuilder();
            StringBuilder SQL2 = new StringBuilder();

            SQL1 = SETSQL8();

            Report report8 = new Report();
            report8.Load(@"REPORT\生產報表-每日得料率報表-烘烤.frx");

            report8.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            TableDataSource table = report8.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            report8.Preview = previewControl6;
            report8.Show();
        }

        public StringBuilder SETSQL8()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(" SELECT");
            SB.AppendFormat(" [MOC] AS '線別',CONVERT(NVARCHAR,[DATES],112) AS '日期',[PROD] AS '品名',[SLOT] AS '桶數',[OVERCOOK] AS '烤焦(KG)',[COOKIESTIME] AS '攪餅麩時間(分)',[COOKIESWH] AS '攪餅麩工時',[NOCOOK] AS '未熟(KG)',[REWORK] AS '重工時間',[RECOOK] AS '重烤時間工時'");
            SB.AppendFormat(" FROM [TKMOC].[dbo].[MOCDAILYRECORDCOOK]");
            SB.AppendFormat(" WHERE  CONVERT(NVARCHAR,[DATES],112)='{0}' ",dateTimePicker15.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" AND [MOC]='{0}'", comboBox11.Text);
            SB.AppendFormat(" AND [PROD]='{0}'", comboBox8.Text);
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");

            return SB;

        }

        public void SETFASTREPORT9()
        {
            StringBuilder SQL1 = new StringBuilder();
            StringBuilder SQL2 = new StringBuilder();

            SQL1 = SETSQL9();

            Report report9 = new Report();
            report9.Load(@"REPORT\生產報表-每日得料率報表-烘烤.frx");

            report9.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            TableDataSource table = report9.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            report9.Preview = previewControl6;
            report9.Show();
        }

        public StringBuilder SETSQL9()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(" SELECT");
            SB.AppendFormat(" [MOC] AS '線別',CONVERT(NVARCHAR,[DATES],112) AS '日期',[PROD] AS '品名',[SLOT] AS '桶數',[OVERCOOK] AS '烤焦(KG)',[COOKIESTIME] AS '攪餅麩時間(分)',[COOKIESWH] AS '攪餅麩工時',[NOCOOK] AS '未熟(KG)',[REWORK] AS '重工時間',[RECOOK] AS '重烤時間工時'");
            SB.AppendFormat(" FROM [TKMOC].[dbo].[MOCDAILYRECORDCOOK]");
            SB.AppendFormat(" WHERE  CONVERT(NVARCHAR,[DATES],112)>='{0}' AND  CONVERT(NVARCHAR,[DATES],112)<='{1}'", dateTimePicker16.Value.ToString("yyyyMMdd"), dateTimePicker17.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" AND [MOC]='{0}'", comboBox10.Text);
            SB.AppendFormat(" ORDER BY [PROD]");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");

            return SB;

        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            SEARCHMOCDAILYRECORDNG();
        }

        private void dateTimePicker6_ValueChanged(object sender, EventArgs e)
        {
            SEARCHMOCDAILYRECORDNG2();
        }

        public string GETNOWSLOT()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT TOP 1 [NOWSLOT] ");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCDAILYSLOT] ");
                sbSql.AppendFormat(@"  WHERE [PROD] ='{0}'  AND CONVERT(NVARCHAR, [DATES],112)<='{1}' ", comboBox6.Text,dateTimePicker10.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  ORDER BY CONVERT(NVARCHAR, [DATES],112) DESC");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter3 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder3 = new SqlCommandBuilder(adapter3);
                sqlConn.Open();
                ds3.Clear();
                adapter3.Fill(ds3, "ds3");
                sqlConn.Close();


                if (ds3.Tables["ds3"].Rows.Count == 0)
                {
                    return "0";
                }
                else
                {
                    if (ds3.Tables["ds3"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        return ds3.Tables["ds3"].Rows[0]["NOWSLOT"].ToString();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                    }
                    return "0";
                }

            }
            catch
            {
                return "0";
            }
            finally
            {
                
            }
        }

        public void SETTEXTBOX25()
        {
            if(!string.IsNullOrEmpty(textBox21.Text)&& !string.IsNullOrEmpty(textBox22.Text) && !string.IsNullOrEmpty(textBox23.Text) && !string.IsNullOrEmpty(textBox24.Text))
            {
                textBox25.Text = (Convert.ToDecimal(textBox21.Text) + Convert.ToDecimal(textBox22.Text) - Convert.ToDecimal(textBox23.Text) - Convert.ToDecimal(textBox24.Text)).ToString();
            }
           
        }

        public void SETTEXTBOX22()
        {
            if (!string.IsNullOrEmpty(textBox25.Text) && !string.IsNullOrEmpty(textBox21.Text) && !string.IsNullOrEmpty(textBox23.Text) && !string.IsNullOrEmpty(textBox24.Text))
            {
                textBox22.Text = (Convert.ToDecimal(textBox25.Text)- Convert.ToDecimal(textBox21.Text) + Convert.ToDecimal(textBox23.Text) + Convert.ToDecimal(textBox24.Text)).ToString();
            }

        }


        private void textBox22_TextChanged(object sender, EventArgs e)
        {
            
            //SETTEXTBOX25();
        }

        private void textBox23_TextChanged(object sender, EventArgs e)
        {

            //SETTEXTBOX25();
            SETTEXTBOX22();
        }

        private void textBox24_TextChanged(object sender, EventArgs e)
        {

            //SETTEXTBOX25();
            SETTEXTBOX22();
        }


        public void ADDMOCDAILYSLOT()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat("  INSERT [TKMOC].[dbo].[MOCDAILYSLOT]");
                sbSql.AppendFormat("  ([PROD],[DATES],[KIND],[LASTSLOT],[PRODOUT],[PRODIN],[NG],[NOWSLOT])");
                sbSql.AppendFormat("  VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7 }')", comboBox6.Text, dateTimePicker10.Value.ToString("yyyy/MM/dd"), comboBox7.Text, textBox21.Text, textBox22.Text, textBox23.Text, textBox24.Text, textBox25.Text);
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
        public void UPDATEMOCDAILYSLOT()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" UPDATE [TKMOC].[dbo].[MOCDAILYSLOT]");
                sbSql.AppendFormat(" SET [LASTSLOT]='{0}',[PRODOUT]='{1}',[PRODIN]='{2}',[NG]='{3}',[NOWSLOT]='{4}'",textBox21.Text, textBox22.Text, textBox23.Text, textBox24.Text, textBox25.Text);
                sbSql.AppendFormat(" WHERE [PROD]='{0}' AND CONVERT(NVARCHAR, [DATES],112)='{1}' AND [KIND]='{2}'", comboBox6.Text,dateTimePicker10.Value.ToString("yyyyMMdd"), comboBox7.Text);
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

        public void ADDMOCDAILYRECORDCOOK()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat("  INSERT INTO [TKMOC].[dbo].[MOCDAILYRECORDCOOK]");
                sbSql.AppendFormat("  ([MOC],[DATES],[PROD],[SLOT],[OVERCOOK],[COOKIESTIME],[COOKIESWH],[NOCOOK],[REWORK],[RECOOK])");
                sbSql.AppendFormat("  VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}')", comboBox11.Text, dateTimePicker15.Value.ToString("yyyy/MM/dd"),comboBox8.Text,textBox26.Text, textBox27.Text, textBox28.Text, textBox29.Text, textBox30.Text, textBox31.Text, textBox32.Text);
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
        public void UPDATEMOCDAILYRECORDCOOK()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" UPDATE [TKMOC].[dbo].[MOCDAILYRECORDCOOK]");
                sbSql.AppendFormat(" SET [SLOT]='{0}',[OVERCOOK]='{1}',[COOKIESTIME]='{2}',[COOKIESWH]='{3}',[NOCOOK]='{4}',[REWORK]='{5}',[RECOOK]='{6}'", textBox26.Text, textBox27.Text, textBox28.Text, textBox29.Text, textBox30.Text, textBox31.Text, textBox32.Text);
                sbSql.AppendFormat(" WHERE  CONVERT(VARCHAR,[DATES],112)='{0}' AND [MOC]='{1}' AND [PROD]='{2}'", dateTimePicker15.Value.ToString("yyyyMMdd"), comboBox11.Text, comboBox8.Text);
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

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox21.Text = "0";
            textBox21.Text = GETNOWSLOT().ToString();
            //SETTEXTBOX25();
        }

        private void textBox21_TextChanged(object sender, EventArgs e)
        {
            //SETTEXTBOX25();
            SETTEXTBOX22();
        }

        private void dateTimePicker15_ValueChanged(object sender, EventArgs e)
        {
            comboBox8load();
        }

        private void comboBox11_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox8load();
        }

        private void textBox25_TextChanged(object sender, EventArgs e)
        {
            SETTEXTBOX22();
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
        private void button5_Click(object sender, EventArgs e)
        {
            ADDMOCDAILYRECORDNG2();
            SETNULL3();

            SEARCHMOCDAILYRECORDNG2();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            UPDATEMOCDAILYRECORDNG2();
            SETNULL3();

            SEARCHMOCDAILYRECORDNG2();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            SETFASTREPORT3();
        }
        private void button8_Click(object sender, EventArgs e)
        {
            SETFASTREPORT4();
        }
        private void button9_Click(object sender, EventArgs e)
        {
            ADDMOCDAILYRECORDNGMONEY();
            SETNULL4();

            SETFASTREPORT4();

        }

        private void button10_Click(object sender, EventArgs e)
        {
            UPDATEMOCDAILYRECORDNGMONEY();
            SETNULL4();

            SETFASTREPORT4();
        }
        private void button11_Click(object sender, EventArgs e)
        {
            SETFASTREPORT5();
        }

        private void button14_Click(object sender, EventArgs e)
        {
            SETFASTREPORT6();

            textBox21.Text = "0";
            textBox21.Text = GETNOWSLOT().ToString();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            ADDMOCDAILYSLOT();
            SETNULL5();

            SETFASTREPORT6();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            UPDATEMOCDAILYSLOT();
            SETNULL5();

            SETFASTREPORT6();
        }

        private void button15_Click(object sender, EventArgs e)
        {
            SETFASTREPORT7();
        }


        private void button18_Click(object sender, EventArgs e)
        {
            SETFASTREPORT8();
        }

        private void button16_Click(object sender, EventArgs e)
        {
            ADDMOCDAILYRECORDCOOK();
            SETNULL6();

            SETFASTREPORT8();
        }

        private void button17_Click(object sender, EventArgs e)
        {
            UPDATEMOCDAILYRECORDCOOK();
            SETNULL6();

            SETFASTREPORT8();

        }

        private void button19_Click(object sender, EventArgs e)
        {
            SETFASTREPORT9();
        }





        #endregion

       
    }
}
