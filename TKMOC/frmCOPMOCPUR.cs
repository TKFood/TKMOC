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

namespace TKMOC
{
    public partial class frmCOPMOCPUR : Form
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

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();

        DataTable dt = new DataTable();
        string tablename = null;
        int result;

        string MID;
        string DID;
        string TA001;
        string TA002;
        string TC015;
        string TC001;
        string TC002;
        string TC003;

        public class PURTADATA
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

            public string TA001;
            public string TA002;
            public string TA003;
            public string TA004;
            public string TA005;
            public string TA006;
            public string TA007;
            public string TA008;
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
            public string TA023;
            public string TA024;
            public string TA025;
            public string TA026;
            public string TA027;
            public string TA028;
            public string TA029;
            public string TA030;
            public string TA031;
            public string TA032;
            public string TA033;
            public string TA034;
            public string TA035;
            public string TA036;
            public string TA037;
            public string TA038;
            public string TA039;
            public string TA040;
            public string TA041;
            public string TA042;
            public string TA043;
            public string TA044;
            public string TA045;
            public string TA046;
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

        public class PURTBDATA
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

            public string TB001;
            public string TB002;
            public string TB003;
            public string TB004;
            public string TB005;
            public string TB006;
            public string TB007;
            public string TB008;
            public string TB009;
            public string TB010;
            public string TB011;
            public string TB012;
            public string TB013;
            public string TB014;
            public string TB015;
            public string TB016;
            public string TB017;
            public string TB018;
            public string TB019;
            public string TB020;
            public string TB021;
            public string TB022;
            public string TB023;
            public string TB024;
            public string TB025;
            public string TB026;
            public string TB027;
            public string TB028;
            public string TB029;
            public string TB030;
            public string TB031;
            public string TB032;
            public string TB033;
            public string TB034;
            public string TB035;
            public string TB036;
            public string TB037;
            public string TB038;
            public string TB039;
            public string TB040;
            public string TB041;
            public string TB042;
            public string TB043;
            public string TB044;
            public string TB045;
            public string TB046;
            public string TB047;
            public string TB048;
            public string TB049;
            public string TB050;
            public string TB051;
            public string TB052;
            public string TB053;
            public string TB054;
            public string TB055;
            public string TB056;
            public string TB057;
            public string TB058;
            public string TB059;
            public string TB060;
            public string TB061;
            public string TB062;
            public string TB063;
            public string TB064;
            public string TB065;
            public string TB066;
            public string TB067;
            public string TB068;
            public string TB069;
            public string TB070;
            public string TB071;
            public string TB072;
            public string TB073;
            public string TB074;
            public string TB075;
            public string TB076;
            public string TB077;
            public string TB078;
            public string TB079;
            public string TB080;
            public string TB081;
            public string TB082;
            public string TB083;
            public string TB084;
            public string TB085;
            public string TB086;
            public string TB087;
            public string TB088;
            public string TB089;
            public string TB090;
            public string TB091;
            public string TB092;
            public string TB093;
            public string TB094;
            public string TB095;
            public string TB096;
            public string TB097;
            public string TB098;
            public string TB099;
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

        public frmCOPMOCPUR()
        {
            InitializeComponent();
        }

        #region FUNCTION

        public void SEARCHCOP(DateTime dt1, DateTime dt2)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  SELECT TD013 AS '預交日',TD001 AS '訂單',TD002 AS '訂單號',TD003 AS '序號',TD004 AS '品號',TD005 AS '品名',TD006 AS '規格',(TD008-TD009+TD024-TD025) AS '訂單數量',TD010 AS '訂單單位'");
                sbSql.AppendFormat(@"  ,CONVERT(DECIMAL(18,3),(CASE WHEN MD002=TD010   THEN (TD008-TD009+TD024-TD025)*MD004/MD003 ELSE (TD008-TD009+TD024-TD025) END )) AS '數量'");
                sbSql.AppendFormat(@"  ,MB004 AS '單位',TC015 AS '單頭備註',TD020 AS '單身備註'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.COPTC,[TK].dbo.COPTD");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMD ON MD001=TD004 AND MD002=TD010 ");
                sbSql.AppendFormat(@"  ,[TK].dbo.INVMB");
                sbSql.AppendFormat(@"  WHERE TC001=TD001 AND TC002=TD002");
                sbSql.AppendFormat(@"  AND TD004=MB001");
                //sbSql.AppendFormat(@"  AND (TD004 LIKE '410%')");
                sbSql.AppendFormat(@"  AND (TD008-TD009)>0");
                sbSql.AppendFormat(@"  AND TD013>='{0}' AND TD013<='{1}'", dt1.ToString("yyyyMMdd"), dt2.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  ORDER BY TD013,TD001,TD002,TD004");
                sbSql.AppendFormat(@"  ");
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
                    dataGridView1.DataSource = null;
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
                sqlConn.Close();
            }

        }


        public void SEARCHCOPMOCPUR(string MID, string DID)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();



                sbSql.AppendFormat(@"  SELECT [MID] AS '來源單別',[DID] AS '來源單號',[TA001] AS '採購單',[TA002] AS '採購單號'");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[COPMOCPUR]");
                sbSql.AppendFormat(@"  WHERE [MID]='{0}' AND [DID]='{1}'", MID, DID);
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
                sqlConn.Close();
            }
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            textBox1.Text = null;
            textBox2.Text = null;
            textBox4.Text = null;

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    textBox1.Text = row.Cells["訂單"].Value.ToString();
                    textBox2.Text = row.Cells["訂單號"].Value.ToString();
                    textBox4.Text = row.Cells["序號"].Value.ToString();

                    SEARCHCOPMOCPUR(textBox1.Text, textBox2.Text);
                }
                else
                {
                    textBox1.Text = null;
                    textBox2.Text = null;
                    textBox4.Text = null;
                }
            }
        }

        public string GETMAXTA002(string TA001, string dt)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds2.Clear();

                sbSql.AppendFormat(@"  SELECT ISNULL(MAX(TA002),'0000000000') AS TA002");
                sbSql.AppendFormat(@"  FROM [TK].[dbo].[PURTA]");
                sbSql.AppendFormat(@"  WHERE  TA001='{0}' AND TA003='{1}'", TA001, dt);
                sbSql.AppendFormat(@"  ");

                adapter3 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder3 = new SqlCommandBuilder(adapter3);
                sqlConn.Open();
                ds3.Clear();
                adapter3.Fill(ds3, "ds3");
                sqlConn.Close();


                if (ds3.Tables["ds3"].Rows.Count == 0)
                {
                    return null;
                }
                else
                {
                    if (ds3.Tables["ds3"].Rows.Count >= 1)
                    {
                        TA002 = SETTA002(ds3.Tables["ds3"].Rows[0]["TA002"].ToString(), dt);
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

        public string SETTA002(string TA002, string dt)
        {
            if (TA002.Equals("0000000000"))
            {
                return dt + "001";
            }

            else
            {
                int serno = Convert.ToInt16(TA002.Substring(8, 3));
                serno = serno + 1;
                string temp = serno.ToString();
                temp = temp.PadLeft(3, '0');
                return dt + temp.ToString();
            }

        }

        public void ADDCOPMOCPUR(string MID,string DID,string TA001,string TA002)
        {
            if (!string.IsNullOrEmpty(MID) && !string.IsNullOrEmpty(DID) && !string.IsNullOrEmpty(TA001) && !string.IsNullOrEmpty(TA002))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();

                    sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[COPMOCPUR]");
                    sbSql.AppendFormat(" ([MID],[DID],[TA001],[TA002])");
                    sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}')",MID,DID,TA001,TA002);
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

        public void ADDPURTAPURTB(string MID,string DID,string TA001,string TA002, string TC001, string TC002,string TC003,string TB011,string TB019)
        {
            PURTADATA PURTA = new PURTADATA();
            PURTA = SETPURTA(MID,DID,dateTimePicker3.Value,TA001,TA002);

            if (!string.IsNullOrEmpty(MID) && !string.IsNullOrEmpty(DID) && !string.IsNullOrEmpty(TA001) && !string.IsNullOrEmpty(TA002))
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();


                    sbSql.AppendFormat(" INSERT INTO [TK].[dbo].[PURTA]");
                    sbSql.AppendFormat(" (");
                    sbSql.AppendFormat(" [COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE]");
                    sbSql.AppendFormat(" ,[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count],[DataUser],[DataGroup]");
                    sbSql.AppendFormat(" ,[TA001],[TA002],[TA003],[TA004],[TA005],[TA006],[TA007],[TA008],[TA009],[TA010]");
                    sbSql.AppendFormat(" ,[TA011],[TA012],[TA013],[TA014],[TA015],[TA016],[TA017],[TA018],[TA019],[TA020]");
                    sbSql.AppendFormat(" ,[TA021],[TA022],[TA023],[TA024],[TA025],[TA026],[TA027],[TA028],[TA029],[TA030]");
                    sbSql.AppendFormat(" ,[TA031],[TA032],[TA033],[TA034],[TA035],[TA036],[TA037],[TA038],[TA039],[TA040]");
                    sbSql.AppendFormat(" ,[TA041],[TA042],[TA043],[TA044],[TA045],[TA046]");
                    sbSql.AppendFormat(" ,[UDF01],[UDF02],[UDF03],[UDF04],[UDF05],[UDF06],[UDF07],[UDF08],[UDF09],[UDF10]");
                    sbSql.AppendFormat(" )");
                    sbSql.AppendFormat(" VALUES");
                    sbSql.AppendFormat(" (");
                    sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}',", PURTA.COMPANY, PURTA.CREATOR, PURTA.USR_GROUP, PURTA.CREATE_DATE, PURTA.MODIFIER, PURTA.MODI_DATE, PURTA.FLAG, PURTA.CREATE_TIME, PURTA.MODI_TIME, PURTA.TRANS_TYPE);
                    sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}','{6}',", PURTA.TRANS_NAME, PURTA.sync_date, PURTA.sync_time, PURTA.sync_mark, PURTA.sync_count, PURTA.DataUser, PURTA.DataGroup);
                    sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}','{6}',{7},'{8}','{9}',", PURTA.TA001, PURTA.TA002, PURTA.TA003, PURTA.TA004, PURTA.TA005, PURTA.TA006, PURTA.TA007, PURTA.TA008, PURTA.TA009, PURTA.TA010);
                    sbSql.AppendFormat(" {0},'{1}','{2}','{3}',{4},'{5}',{6},'{7}','{8}',{9},", PURTA.TA011, PURTA.TA012, PURTA.TA013, PURTA.TA014, PURTA.TA015, PURTA.TA016, PURTA.TA017, PURTA.TA018, PURTA.TA019, PURTA.TA020);
                    sbSql.AppendFormat(" '{0}','{1}',{2},{3},'{4}','{5}','{6}','{7}','{8}',{9},", PURTA.TA021, PURTA.TA022, PURTA.TA023, PURTA.TA024, PURTA.TA025, PURTA.TA026, PURTA.TA027, PURTA.TA028, PURTA.TA029, PURTA.TA030);
                    sbSql.AppendFormat(" '{0}',{1},'{2}','{3}','{4}',{5},{6},{7},{8},{9},", PURTA.TA031, PURTA.TA032, PURTA.TA033, PURTA.TA034, PURTA.TA035, PURTA.TA036, PURTA.TA037, PURTA.TA038, PURTA.TA039, PURTA.TA040);
                    sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}',", PURTA.TA041, PURTA.TA042, PURTA.TA043, PURTA.TA044, PURTA.TA045, PURTA.TA046);
                    sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',{5},{6},{7},{8},{9}", PURTA.UDF01, PURTA.UDF02, PURTA.UDF03, PURTA.UDF04, PURTA.UDF05, PURTA.UDF06, PURTA.UDF07, PURTA.UDF08, PURTA.UDF09, PURTA.UDF10);
                    sbSql.AppendFormat(" )");
                    sbSql.AppendFormat(" ");
                    sbSql.AppendFormat(" INSERT INTO [TK].[dbo].[PURTB]");
                    sbSql.AppendFormat(" (");
                    sbSql.AppendFormat(" [COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE]");
                    sbSql.AppendFormat(" ,[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count],[DataUser],[DataGroup]");
                    sbSql.AppendFormat(" ,[TB001],[TB002],[TB003],[TB004],[TB005],[TB006],[TB007],[TB008],[TB009],[TB010]");
                    sbSql.AppendFormat(" ,[TB011],[TB012],[TB013],[TB014],[TB015],[TB016],[TB017],[TB018],[TB019],[TB020]");
                    sbSql.AppendFormat(" ,[TB021],[TB022],[TB023],[TB024],[TB025],[TB026],[TB027],[TB028],[TB029],[TB030]");
                    sbSql.AppendFormat(" ,[TB031],[TB032],[TB033],[TB034],[TB035],[TB036],[TB037],[TB038],[TB039],[TB040]");
                    sbSql.AppendFormat(" ,[TB041],[TB042],[TB043],[TB044],[TB045],[TB046],[TB047],[TB048],[TB049],[TB050]");
                    sbSql.AppendFormat(" ,[TB051],[TB052],[TB053],[TB054],[TB055],[TB056],[TB057],[TB058],[TB059],[TB060]");
                    sbSql.AppendFormat(" ,[TB061],[TB062],[TB063],[TB064],[TB065],[TB066],[TB067],[TB068],[TB069],[TB070]");
                    sbSql.AppendFormat(" ,[TB071],[TB072],[TB073],[TB074],[TB075],[TB076],[TB077],[TB078],[TB079],[TB080]");
                    sbSql.AppendFormat(" ,[TB081],[TB082],[TB083],[TB084],[TB085],[TB086],[TB087],[TB088],[TB089],[TB090]");
                    sbSql.AppendFormat(" ,[TB091],[TB092],[TB093],[TB094],[TB095],[TB096],[TB097],[TB098],[TB099]");
                    sbSql.AppendFormat(" ,[UDF01],[UDF02],[UDF03],[UDF04],[UDF05],[UDF06],[UDF07],[UDF08],[UDF09],[UDF10]");
                    sbSql.AppendFormat(" )");
                    sbSql.AppendFormat(" (");
                    sbSql.AppendFormat(" SELECT ");
                    sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}',", PURTA.COMPANY, PURTA.CREATOR, PURTA.USR_GROUP, PURTA.CREATE_DATE, PURTA.MODIFIER, PURTA.MODI_DATE, PURTA.FLAG, PURTA.CREATE_TIME, PURTA.MODI_TIME, PURTA.TRANS_TYPE);
                    sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}','{6}',", PURTA.TRANS_NAME, PURTA.sync_date, PURTA.sync_time, PURTA.sync_mark, PURTA.sync_count, PURTA.DataUser, PURTA.DataGroup);
                    sbSql.AppendFormat(" '{0}' AS [TB001],'{1}' AS [TB002],RIGHT(REPLICATE('0',4) + CAST(ROW_NUMBER() OVER(ORDER BY TD001,TD002,TD003)  as NVARCHAR),4)  AS [TB003],MB001 AS [TB004],MB002 AS [TB005],MB003 AS [TB006],MB004 AS [TB007],MB017 AS [TB008],(TD008-TD009+TD024-TD025) AS [TB009],MB032 AS [TB010]", PURTA.TA001, PURTA.TA002);
                    sbSql.AppendFormat(" ,'{0}' AS [TB011],TD001+'-'+TD002+'-'+TD003 AS [TB012],'170017' AS [TB013],(TD008-TD009+TD024-TD025) AS [TB014],MB004 AS [TB015],'NTD' AS [TB016],MB049 AS [TB017],(TD008-TD009+TD024-TD025)*MB049 AS [TB018],'{0}' AS [TB019],'N' AS [TB020]", TB011,TB019);
                    sbSql.AppendFormat(" ,'N' AS [TB021],'' AS [TB022],'' AS [TB023],'' AS [TB024],'N' AS [TB025],MA044 AS [TB026],'' AS [TB027],'' AS [TB028],'' AS [TB029],'' AS [TB030]");
                    sbSql.AppendFormat(" ,'' AS [TB031],'N' AS [TB032],'' AS [TB033],'0' AS [TB034],'0' AS [TB035],'' AS [TB036],'' AS [TB037],'' AS [TB038],'N' AS [TB039],MB051 AS [TB040]");
                    sbSql.AppendFormat(" ,(TD008-TD009+TD024-TD025)*MB051 AS [TB041],'' AS [TB042],'' AS [TB043],'' AS [TB044],'' AS [TB045],'' AS [TB046],'' AS [TB047],'' AS [TB048],'0' AS [TB049],'' AS [TB050]");
                    sbSql.AppendFormat(" ,'0' AS [TB051],'0' AS [TB052],'0' AS [TB053],'' AS [TB054],'' AS [TB055],'' AS [TB056],'' AS [TB057],'1' AS [TB058],'' AS [TB059],'' AS [TB060]");
                    sbSql.AppendFormat(" ,'' AS [TB061],'' AS [TB062],'0' AS [TB063],'N' AS [TB064],'1' AS [TB065],'' AS [TB066],'2' AS [TB067],'0' AS [TB068],'0' AS [TB069],'' AS [TB070]");
                    sbSql.AppendFormat(" ,'' AS [TB071],'' AS [TB072],'' AS [TB073],'' AS [TB074],'0' AS [TB075],'' AS [TB076],'0' AS [TB077],'' AS [TB078],'' AS [TB079],'' AS [TB080]");
                    sbSql.AppendFormat(" ,'0' AS [TB081],'0' AS [TB082],'0' AS [TB083],'0' AS [TB084],'0' AS [TB085],'' AS [TB086],'' AS [TB087],'0' AS [TB088],'1' AS [TB089],'0' AS [TB090]");
                    sbSql.AppendFormat(" ,'0' AS [TB091],'0' AS [TB092],'0' AS [TB093],'' AS [TB094],'' AS [TB095],'' AS [TB096],'' AS [TB097],'' AS [TB098],'' AS [TB099]");
                    sbSql.AppendFormat(" ,'' AS [UDF01],'' AS [UDF02],'' AS [UDF03],'' AS [UDF04],'' AS [UDF05],'0' AS [UDF06],'0' AS [UDF07],'0' AS [UDF08],'0' AS [UDF09],'0' AS [UDF10]");
                    sbSql.AppendFormat(" FROM [TK].dbo.COPTC,[TK].dbo.COPTD,[TK].dbo.INVMB");
                    sbSql.AppendFormat(" LEFT JOIN [TK].dbo.PURMA ON MA001=MB032");
                    sbSql.AppendFormat(" WHERE TC001=TD001 AND TC002=TD002");
                    sbSql.AppendFormat(" AND MB001=TD004");
                    sbSql.AppendFormat(" AND TD001='{0}' AND TD002='{1}' AND TD003='{2}'",TC001,TC002,TC003);
                    sbSql.AppendFormat(" )");
                
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

        public PURTADATA SETPURTA(string MID,string DID,DateTime dt,string TA001,string TA002)
        {
            PURTADATA PURTA = new PURTADATA();
            PURTA.COMPANY = "TK";
            PURTA.CREATOR = "140020";
            PURTA.USR_GROUP = "103000";
            PURTA.CREATE_DATE = dt.ToString("yyyyMMdd");
            PURTA.MODIFIER = "120024";
            PURTA.MODI_DATE = dt.ToString("yyyyMMdd");
            PURTA.FLAG = "0";
            PURTA.CREATE_TIME = dt.ToString("HH:mm:dd");
            PURTA.MODI_TIME = dt.ToString("HH:mm:dd");
            PURTA.TRANS_TYPE = "P001";
            PURTA.TRANS_NAME = "PURI05";
            PURTA.sync_date = null;
            PURTA.sync_time = null;
            PURTA.sync_mark = null;
            PURTA.sync_count = "0";
            PURTA.DataUser = null;
            PURTA.DataGroup = "103000";

            PURTA.TA001 = TA001;
            PURTA.TA002 = TA002;
            PURTA.TA003 = dt.ToString("yyyyMMdd");
            PURTA.TA004 = "103500";
            PURTA.TA005 = null;
            PURTA.TA006 = TC015;
            PURTA.TA007 = "N";
            PURTA.TA008 = "0";
            PURTA.TA009 = "9";
            PURTA.TA010 = "20";
            PURTA.TA011 = "0";
            PURTA.TA012 = "140020";
            PURTA.TA013 = dt.ToString("yyyyMMdd");
            PURTA.TA014 = null;
            PURTA.TA015 = "0";
            PURTA.TA016 = "N";
            PURTA.TA017 = "0";
            PURTA.TA018 = null;
            PURTA.TA019 = null;
            PURTA.TA020 = "0";
            PURTA.TA021 = null;
            PURTA.TA022 = null;
            PURTA.TA023 = "0";
            PURTA.TA024 = "0";
            PURTA.TA025 = null;
            PURTA.TA026 = null;
            PURTA.TA027 = null;
            PURTA.TA028 = null;
            PURTA.TA029 = null;
            PURTA.TA030 = "0";
            PURTA.TA031 = null;
            PURTA.TA032 = "0";
            PURTA.TA033 = null;
            PURTA.TA034 = null;
            PURTA.TA035 = null;
            PURTA.TA036 = "0";
            PURTA.TA037 = "0";
            PURTA.TA038 = "0";
            PURTA.TA039 = "0";
            PURTA.TA040 = "0";
            PURTA.TA041 = null;
            PURTA.TA042 = null;
            PURTA.TA043 = null;
            PURTA.TA044 = null;
            PURTA.TA045 = null;
            PURTA.TA046 = null;
            PURTA.UDF01 = null;
            PURTA.UDF02 = null;
            PURTA.UDF03 = null;
            PURTA.UDF04 = null;
            PURTA.UDF05 = null;
            PURTA.UDF06 = "0";
            PURTA.UDF07 = "0";
            PURTA.UDF08 = "0";
            PURTA.UDF09 = "0";
            PURTA.UDF10 = "0";

            return PURTA;

        }


        public void DELETECOPMOCPUR(string MID,string DID)
        {
            if (!string.IsNullOrEmpty(MID) && !string.IsNullOrEmpty(DID) )
            {
                try
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();

                    sbSql.AppendFormat(" DELETE [TKMOC].[dbo].[COPMOCPUR]");
                    sbSql.AppendFormat(" WHERE [MID]='{0}' AND [DID]='{1}'",MID,DID);
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

        public void SEARCHCOPTC(string TC001,string TC002)
        {
            TC015 = null;

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@" SELECT TOP 1 TC015 ");
                sbSql.AppendFormat(@" FROM [TK].dbo.COPTC ");
                sbSql.AppendFormat(@" WHERE TC001='{0}' AND TC002='{1}' ",TC001,TC002);
                sbSql.AppendFormat(@"  ");

                adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder4= new SqlCommandBuilder(adapter4);
                sqlConn.Open();
                ds4.Clear();
                adapter4.Fill(ds4, "ds4");
                sqlConn.Close();


                if (ds4.Tables["ds4"].Rows.Count == 0)
                {
                    TC015 = null;
                    
                }
                else
                {
                    if (ds4.Tables["ds4"].Rows.Count >= 1)
                    {
                        TC015 = ds4.Tables["ds4"].Rows[0]["TC015"].ToString();
                     
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

        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            SEARCHCOP(dateTimePicker1.Value,dateTimePicker2.Value);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SEARCHCOPTC(textBox1.Text,textBox2.Text);

            TA001 = textBox3.Text;
            TA002 = GETMAXTA002(TA001, dateTimePicker3.Value.ToString("yyyyMMdd"));

            TC001 = textBox1.Text;
            TC002 = textBox2.Text;
            TC003 = textBox4.Text;
            //ADDCOPMOCPUR(textBox1.Text, textBox2.Text, TA001, TA002);
            ADDPURTAPURTB(textBox1.Text,textBox2.Text, TA001, TA002, TC001, TC002, TC003,dateTimePicker4.Value.ToString("yyyyMMdd"), dateTimePicker4.Value.ToString("yyyyMMdd"));

            SEARCHCOPMOCPUR(textBox1.Text,textBox2.Text);

            //MessageBox.Show(TA001 + " " + TA002);
        }
        private void button11_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELETECOPMOCPUR(textBox1.Text, textBox2.Text);
                SEARCHCOPMOCPUR(textBox1.Text, textBox2.Text);
            }
        }

        #endregion


    }
}
