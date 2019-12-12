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
    public partial class frmBATCHMOCTAB : Form
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
        SqlDataAdapter adapter3= new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder3 = new SqlCommandBuilder();
        SqlDataAdapter adapter4 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder4 = new SqlCommandBuilder();
        SqlDataAdapter adapter5 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder5= new SqlCommandBuilder();

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();
        DataSet ds5 = new DataSet();

        int result;

        List<ADDITEM> ADDTARGET = new List<ADDITEM>();
        List<ADDITEM> FIND = new List<ADDITEM>();

        string TA001;
        string TA002;
        DateTime DTMOCTAB;

        string BOMVARSION;
        string UNIT;
        decimal BOMBAR;
        string MB002;
        string MB003;
        string IN;

        DataSet dsBOMMC = new DataSet();
        DataSet dsBOMMD = new DataSet();

        public class ADDITEM
        {
            public string MB001;
            public double NUM;
            public string MB068;

        }

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
            public string TA026;
            public string TA027;
            public string TA028;
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

        public frmBATCHMOCTAB()
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
                sbSql.AppendFormat(@"  ,MB068 AS '生產線別' ");
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
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            textBox1.Text = null;
            textBox2.Text = null; 
            textBox3.Text = null;
            textBox4.Text = null;
            textBox5.Text = null;
            textBox6.Text = null;
            textBox7.Text = null;

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    textBox1.Text = row.Cells["訂單"].Value.ToString();
                    textBox2.Text = row.Cells["訂單號"].Value.ToString();
                    textBox3.Text = row.Cells["序號"].Value.ToString();
                    textBox4.Text = row.Cells["品號"].Value.ToString();
                    textBox5.Text = row.Cells["數量"].Value.ToString();
                    textBox6.Text = row.Cells["單頭備註"].Value.ToString();
                    textBox7.Text = row.Cells["生產線別"].Value.ToString();
                }
                else
                {
                    textBox1.Text = null;
                    textBox2.Text = null;
                    textBox3.Text = null;
                    textBox4.Text = null;
                    textBox5.Text = null;
                    textBox6.Text = null;
                    textBox7.Text = null;
                }
            }
        }

        public void GENADDTARGET()
        {
            ADDTARGET.Clear();
            FIND.Clear();
            //ADDTARGET.RemoveAll(it => true);

            ADDTARGET.Add(new ADDITEM { MB001 =textBox4.Text , NUM = Convert.ToDouble(textBox5.Text) ,MB068=textBox7.Text});

            SERACH(ADDTARGET[0].MB001, ADDTARGET[0].NUM, FIND);

            foreach (var find in FIND)
            {
                CHECKBOMMD(find.MB001, find.NUM);
            }

            //foreach (var find in ADDTARGET)
            //{
            //    MessageBox.Show(find.MB001 + " " + find.NUM);
            //}

        }

        public void SERACH(string MB001, double NUM, List<ADDITEM> FIND)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  WITH NODE (MD001,MD003,LAYER,MD004,MC004,PREMC004,MD006,MD007,MD008,USEDNUM ,PREUSEDNUM) AS");
                sbSql.AppendFormat(@"  (");
                sbSql.AppendFormat(@"  SELECT MD001,MD003,0 ,[MD004],[MC004],[MC004] AS PREMC004,[MD006],[MD007],[MD008],CONVERT(DECIMAL(18,4),([MD006]/[MD007]/[MC004]*(1+MD008))),CONVERT(DECIMAL(18,4),1) AS PREUSEDNUM  FROM [TK].[dbo].[VBOMMD]");
                sbSql.AppendFormat(@"  UNION ALL");
                sbSql.AppendFormat(@"  SELECT TB1.MD001,TB2.MD003,TB2.LAYER+1,TB2.MD004,TB2.MC004,TB1.MC004,TB2.MD006,TB2.MD007,TB2.MD008,TB2.USEDNUM,CONVERT(DECIMAL(18,4),(TB1.[MD006]/TB1.[MD007]/TB1.[MC004]*(1+TB1.MD008))) AS PREUSEDNUM FROM [TK].[dbo].[VBOMMD] TB1");
                sbSql.AppendFormat(@"  INNER JOIN NODE TB2");
                sbSql.AppendFormat(@"  ON TB1.MD003 = TB2.MD001");
                sbSql.AppendFormat(@"  )");
                sbSql.AppendFormat(@"  SELECT MD001,MD003,LAYER,MD004,MC004,PREMC004,MD006,MD007,MD008,USEDNUM ,PREUSEDNUM ,USEDNUM*PREUSEDNUM*{0} AS TOTALUSED FROM NODE", NUM);
                sbSql.AppendFormat(@"  WHERE  MD001='{0}'", MB001);
                sbSql.AppendFormat(@"  ORDER BY LAYER ,MD001, MD003");
                sbSql.AppendFormat(@"  ");
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

                }
                else
                {
                    if (ds2.Tables["ds2"].Rows.Count >= 1)
                    {
                        foreach (DataRow od in ds2.Tables["ds2"].Rows)
                        {
                            FIND.Add(new ADDITEM { MB001 = od["MD003"].ToString(), NUM = Convert.ToDouble(od["TOTALUSED"].ToString()) });
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

        public void CHECKBOMMD(string MB001, double NUM)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT MD001,MD003,MB068");
                sbSql.AppendFormat(@"  FROM [TK].dbo.BOMMD,[TK].dbo.INVMB ");
                sbSql.AppendFormat(@"  WHERE MD001=MB001 AND MD001='{0}'", MB001);
                sbSql.AppendFormat(@"  ");
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

                }
                else
                {
                    if (ds3.Tables["ds3"].Rows.Count >= 1)
                    {
                        ADDTARGET.Add(new ADDITEM { MB001 = MB001, NUM = NUM,MB068= ds3.Tables["ds3"].Rows[0]["MB068"].ToString() });

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

        public void GENMOCTAB(DateTime DTMOCTAB,string TA029,string TC001, string TC002, string TC003)
        {
            foreach (var find in ADDTARGET)
            {
                TA001 = "A510";
                TA002 = GETMAXTA002(TA001, DTMOCTAB);

                ADDMOCTATB(TA001, TA002, find.MB001, find.NUM, find.MB068, DTMOCTAB, TA029, TC001, TC002, TC003);
                //MessageBox.Show(find.MB001 + " " + find.NUM + " " + find.MB068 + " "+ TA001+"-"+ TA002);
            }


        }

        public string GETMAXTA002(string TA001,DateTime DTMOCTAB)
        {
            string TA002;

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
                sbSql.AppendFormat(@"  WHERE  TA001='{0}' AND TA003='{1}'", TA001, DTMOCTAB.ToString("yyyyMMdd"));
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
                        TA002 = SETTA002(ds4.Tables["TEMPds4"].Rows[0]["TA002"].ToString(), DTMOCTAB);
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
        public string SETTA002(string TA002,DateTime DTMOCTAB)
        {
            if (TA002.Equals("00000000000"))
            {
                return DTMOCTAB.ToString("yyyyMMdd") + "001";
            }

            else
            {
                int serno = Convert.ToInt16(TA002.Substring(8, 3));
                serno = serno + 1;
                string temp = serno.ToString();
                temp = temp.PadLeft(3, '0');
                return DTMOCTAB.ToString("yyyyMMdd") + temp.ToString();
            }

        }

        public void ADDMOCTATB(string TA001,string TA002,string MB001,double NUM,string MB068,DateTime DTMOCTAB,string TA029, string TC001, string TC002, string TC003)
        {
            MOCTADATA MOCTA = new MOCTADATA();
            MOCTA = SETMOCTA(TA001, TA002, MB001, NUM, MB068, DTMOCTAB,TA029,TC001,TC002,TC003);

            string MOCMB001 = null;
            decimal MOCTA004 = 0; ;
            string MOCTB009 = null;


            const int MaxLength = 100;

            MOCMB001 = MB001;
            MOCTA004 = Convert.ToDecimal(NUM);

            try
            {
                //check TA002=2,TA040=2
                if (MOCTA.TA002.Substring(0, 1).Equals("2") && MOCTA.TA040.Substring(0, 1).Equals("2"))
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
                    sbSql.AppendFormat(" ,[TA019],[TA020],[TA021],[TA022],[TA024],[TA025],[TA026],[TA027],[TA028],[TA029],[TA030],[TA031],[TA034],[TA035]");
                    sbSql.AppendFormat(" ,[TA040],[TA041],[TA043],[TA044],[TA045],[TA046],[TA047],[TA049],[TA050],[TA200]");
                    sbSql.AppendFormat(" ");
                    sbSql.AppendFormat(" )");
                    sbSql.AppendFormat(" VALUES");
                    sbSql.AppendFormat(" ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}',", MOCTA.COMPANY, MOCTA.CREATOR, MOCTA.USR_GROUP, MOCTA.CREATE_DATE, MOCTA.MODIFIER, MOCTA.MODI_DATE, MOCTA.FLAG, MOCTA.CREATE_TIME, MOCTA.MODI_TIME, MOCTA.TRANS_TYPE);
                    sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}',", MOCTA.TRANS_NAME, MOCTA.sync_count, MOCTA.DataGroup, MOCTA.TA001, MOCTA.TA002, MOCTA.TA003, MOCTA.TA004, MOCTA.TA005, MOCTA.TA006, MOCTA.TA007);
                    sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}',", MOCTA.TA009, MOCTA.TA010, MOCTA.TA011, MOCTA.TA012, MOCTA.TA013, MOCTA.TA014, MOCTA.TA015, MOCTA.TA016, MOCTA.TA017, MOCTA.TA018);
                    sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}',N'{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}',", MOCTA.TA019, MOCTA.TA020, MOCTA.TA021, MOCTA.TA022, MOCTA.TA024, MOCTA.TA025, MOCTA.TA026, MOCTA.TA027, MOCTA.TA028, MOCTA.TA029, MOCTA.TA030, MOCTA.TA031, MOCTA.TA034, MOCTA.TA035);
                    sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}'", MOCTA.TA040, MOCTA.TA041, MOCTA.TA043, MOCTA.TA044, MOCTA.TA045, MOCTA.TA046, MOCTA.TA047, MOCTA.TA049, MOCTA.TA050, MOCTA.TA200);
                    sbSql.AppendFormat(" )");
                    sbSql.AppendFormat(" ");
                    sbSql.AppendFormat(" ");
                    //sbSql.AppendFormat(" INSERT INTO [TK].dbo.[MOCTB]");
                    //sbSql.AppendFormat(" ([COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE]");
                    //sbSql.AppendFormat(" ,[TRANS_NAME],[sync_count],[DataGroup],[TB001],[TB002],[TB003],[TB004],[TB005],[TB006],[TB007]");
                    //sbSql.AppendFormat(" ,[TB009],[TB011],[TB012],[TB013],[TB014],[TB018],[TB019],[TB020],[TB022],[TB024]");
                    //sbSql.AppendFormat(" ,[TB025],[TB026],[TB027],[TB029],[TB030],[TB031],[TB501],[TB554],[TB556],[TB560])");
                    //sbSql.AppendFormat(" (SELECT ");
                    //sbSql.AppendFormat(" '{0}' [COMPANY],'{1}' [CREATOR],'{2}' [USR_GROUP],'{3}' [CREATE_DATE],'{4}' [MODIFIER],'{5}' [MODI_DATE],'{6}' [FLAG],'{7}' [CREATE_TIME],'{8}' [MODI_TIME],'{9}' [TRANS_TYPE]", MOCTA.COMPANY, MOCTA.CREATOR, MOCTA.USR_GROUP, MOCTA.CREATE_DATE, MOCTA.MODIFIER, MOCTA.MODI_DATE, MOCTA.FLAG, MOCTA.CREATE_TIME, MOCTA.MODI_TIME, MOCTA.TRANS_TYPE);
                    //sbSql.AppendFormat(" ,'{0}' [TRANS_NAME],{1} [sync_count],'{2}' [DataGroup],'{3}' [TB001],'{4}' [TB002],[BOMMD].MD003 [TB003],ROUND({5}*[BOMMD].MD006/[BOMMD].MD007*(1+[BOMMD].MD008),3) [TB004],0 [TB005],'****' [TB006],[INVMB].MB004  [TB007]", MOCTA.TRANS_NAME, MOCTA.sync_count, MOCTA.DataGroup, MOCTA.TA001, MOCTA.TA002, MOCTA004);
                    //sbSql.AppendFormat(" ,[INVMB].MB017 [TB009],'1' [TB011],[INVMB].MB002 [TB012],[INVMB].MB003 [TB013],[BOMMD].MD001 [TB014],'N' [TB018],'0' [TB019],'0' [TB020],'2' [TB022],'0' [TB024]");
                    //sbSql.AppendFormat(" ,'****' [TB025],'0' [TB026],'1' [TB027],'0' [TB029],'0' [TB030],'0' [TB031],'0' [TB501],'N' [TB554],'0' [TB556],'0' [TB560]");
                    //sbSql.AppendFormat(" FROM [TK].dbo.[BOMMD],[TK].dbo.[INVMB]");
                    //sbSql.AppendFormat(" WHERE [BOMMD].MD003=[INVMB].MB001");
                    //sbSql.AppendFormat(" AND MD001='{0}' AND ISNULL(MD012,'')='' )", MOCMB001);
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


            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        public MOCTADATA SETMOCTA(string TA001, string TA002, string MB001, double NUM, string MB068, DateTime DTMOCTAB,string TA029, string TC001, string TC002, string TC003)
        {
            SEARCHBOMMC(MB001);

            MOCTADATA MOCTA = new MOCTADATA();
            MOCTA.COMPANY = "TK";
            MOCTA.CREATOR = "140020";
            MOCTA.USR_GROUP = "103000";
            //MOCTA.CREATE_DATE = dt1.ToString("yyyyMMdd");
            MOCTA.CREATE_DATE = DateTime.Now.ToString("yyyyMMdd");
            MOCTA.MODIFIER = "140020";
            MOCTA.MODI_DATE = DTMOCTAB.ToString("yyyyMMdd");
            MOCTA.FLAG = "0";
            MOCTA.CREATE_TIME = DateTime.Now.ToString("HH:mm:dd");
            MOCTA.MODI_TIME = DateTime.Now.ToString("HH:mm:dd");
            MOCTA.TRANS_TYPE = "P001";
            MOCTA.TRANS_NAME = "MOCMI02";
            MOCTA.sync_count = "0";
            MOCTA.DataGroup = "103000";
            MOCTA.TA001 = "A510";
            MOCTA.TA002 = TA002;
            MOCTA.TA003 = DTMOCTAB.ToString("yyyyMMdd");
            MOCTA.TA004 = DTMOCTAB.ToString("yyyyMMdd");
            MOCTA.TA005 = BOMVARSION;
            MOCTA.TA006 = MB001;
            MOCTA.TA007 = UNIT;
            MOCTA.TA009 = DTMOCTAB.ToString("yyyyMMdd");
            MOCTA.TA010 = DTMOCTAB.ToString("yyyyMMdd");
            MOCTA.TA011 = "1";
            MOCTA.TA012 = DTMOCTAB.ToString("yyyyMMdd");
            MOCTA.TA013 = "N";           
            MOCTA.TA014 = "";          
            MOCTA.TA015 = NUM.ToString();
            MOCTA.TA016 = "0";
            MOCTA.TA017 = "0";
            MOCTA.TA018 = "0";
            MOCTA.TA019 = "20";
            MOCTA.TA020 = IN;
            MOCTA.TA021 = "02";
            MOCTA.TA022 = "0";
            MOCTA.TA024 = "A510";
            MOCTA.TA025 = TA002;
            MOCTA.TA026 = TC001;
            MOCTA.TA027 = TC002;
            MOCTA.TA028 = TC003;
            MOCTA.TA029 = TA029;
            MOCTA.TA030 = "1";
            MOCTA.TA031 = "0";
            MOCTA.TA034 = MB002;
            MOCTA.TA035 = MB003;
            MOCTA.TA040 = DTMOCTAB.ToString("yyyyMMdd");
            MOCTA.TA041 = "";
            MOCTA.TA043 = "1";
            MOCTA.TA044 = "N";
            MOCTA.TA045 = "0";
            MOCTA.TA046 = "0";
            MOCTA.TA047 = "0";
            MOCTA.TA049 = "0";
            MOCTA.TA050 = "0";
            MOCTA.TA200 = "1";

            return MOCTA;

        }


        public void SEARCHBOMMC(string MB001)
        {
            BOMVARSION = null;
            UNIT = null;
            BOMBAR = 0;
            MB002 = null;
            MB003 = null;

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT ");
                sbSql.AppendFormat(@"  [MC001],[MC002],[MC003],[MC004],[MC005],[MC006],[MC007],[MC008],[MC009],[MC010],[MC011],[MC012],[MC013],[MC014],[MC015],[MC016],[MC017],[MC018],[MC019],[MC020],[MC021],[MC022],[MC023],[MC024],[MC025],[MC026],[MC027]");
                sbSql.AppendFormat(@"  ,INVMB.MB002,INVMB.MB003,INVMB.MB004,INVMB.MB017");
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
                    IN = null;
                    MB002 = null;
                    MB003 = null;
                }
                else
                {
                    if (dsBOMMC.Tables["dsBOMMC"].Rows.Count >= 1)
                    {
                        BOMVARSION = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MC009"].ToString();
                        //UNIT = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MC002"].ToString();
                        UNIT = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MB004"].ToString();
                        BOMBAR = Convert.ToDecimal(dsBOMMC.Tables["dsBOMMC"].Rows[0]["MC004"].ToString());
                        IN= dsBOMMC.Tables["dsBOMMC"].Rows[0]["MB017"].ToString();
                        MB002 = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MB002"].ToString();
                        MB003 = dsBOMMC.Tables["dsBOMMC"].Rows[0]["MB003"].ToString();
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
       

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCHCOP(dateTimePicker1.Value, dateTimePicker2.Value);
        }
        private void button2_Click(object sender, EventArgs e)
        {
            DTMOCTAB = dateTimePicker3.Value;

            GENADDTARGET();
            GENMOCTAB(DTMOCTAB,textBox7.Text,textBox1.Text, textBox2.Text, textBox3.Text);
        }

        #endregion

        
    }
}
