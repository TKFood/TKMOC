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
using TKITDLL;

namespace TKMOC
{
    public partial class frmMOCMANULINEBAKING : Form
    {
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
     
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();
        int result;

        string MANU = "";
        // 宣告一個變數來儲存使用者手動選擇排序的欄位
        string SortedColumn = string.Empty;
        string SortedModel = string.Empty;

        string ID;
        DateTime dt1 ;
        string MB001B;
        string MB002B;
        string MB003B;
        decimal BOX ;
        decimal SUM2;
        string TA001 = "A513";
        string TA002;
        string TA029;
        string TA026 ;
        string TA027 ;
        string TA028;
        string SUBID;
        string SUBBAR2;
        string SUBNUM2;
        string SUBBOX;
        string SUBPACKAGE2;
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
            public string TA033;
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

        public frmMOCMANULINEBAKING()
        {
            InitializeComponent();
        }

        #region FUNCTION
        private void frmMOCMANULINEBAKING_Load(object sender, EventArgs e)
        {
            MANU = "吧台烘焙線";

            comboBox1load();
            comboBox2load();
            comboBox3load();
        }
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage1"])
            {
                //MessageBox.Show("製二線");
                MANU = "吧台烘焙線";
            }
        }
        public void comboBox1load()
        {
            LoadComboBoxData(comboBox1, "SELECT MD001,MD002 FROM [TK].dbo.CMSMD WHERE MD001 IN ('08')  ", "MD002", "MD002");
        }
        public void comboBox2load()
        {
            LoadComboBoxData(comboBox2, "SELECT MD001,MD002 FROM [TK].dbo.CMSMD WHERE MD001 IN ('08')  ", "MD002", "MD002");
        }

        public void comboBox3load()
        {
            LoadComboBoxData(comboBox3, "SELECT MC001 ,MC001+MC002 AS 'MC002' FROM [TK].dbo.CMSMC WHERE MC001 LIKE '2000%'  ORDER BY MC001 ", "MC001", "MC002");
        }
        public void LoadComboBoxData(ComboBox comboBox, string query, string valueMember, string displayMember)
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            using (SqlConnection connection = new SqlConnection(sqlsb.ConnectionString))
            {
                SqlCommand command = new SqlCommand(query, connection);
                connection.Open();

                SqlDataAdapter adapter = new SqlDataAdapter(command);
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);

                comboBox.DataSource = dataTable;
                comboBox.ValueMember = valueMember;
                comboBox.DisplayMember = displayMember;
            }
        }

        public void SEARCHMOCMANULINE_BAKING(string SDATES,string MANU)
        {
            if (MANU.Equals("吧台烘焙線"))
            {
                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"                                      
                                    SELECT 
                                    [MANU] AS '線別'
                                    ,CONVERT(varchar(100),[MANUDATE],112) AS '生產日'
                                    ,[MOCMANULINEBAKING].[MB001] AS '品號'
                                    ,[MOCMANULINEBAKING].[MB002] AS '品名' 
                                    ,[MOCMANULINEBAKING].[MB003] AS '規格'
                                    ,ALLERGEN AS '過敏原'
                                    ,ORI AS '素別'
                                    ,[BAR] AS '桶數'
                                    ,[PACKAGE] AS '包裝數'
                                    ,[NUM] AS '數量'
                                    ,[CLINET] AS '客戶'
                                    ,[OUTDATE] AS '交期'
                                    ,[TA029] AS '備註'
                                    ,[HALFPRO] AS '半成品數量'
                                    ,[COPTD001] AS '訂單單別'
                                    ,[COPTD002] AS '訂單號'
                                    ,[COPTD003] AS '訂單序號'
                                    ,[BOX] AS '箱數'
                                    ,[SERNO]
                                    ,[ID]

                                    FROM [TKMOC].[dbo].[MOCMANULINEBAKING]
                                    LEFT JOIN [TKMOC].[dbo].[ERPINVMB] ON [ERPINVMB].MB001=[MOCMANULINEBAKING].MB001

                                    WHERE [MANU]='{0}' 
                                    AND CONVERT(varchar(100),[MANUDATE],112) LIKE '{1}%'
                                    ORDER BY [MANUDATE],[SERNO]"

                                   , MANU, SDATES);

                sbSql.AppendFormat(@"  ");

                SEARCH_MANULINE(sbSql.ToString(), dataGridView1, SortedColumn, SortedModel);

                ////SET欄位寬度
                //if (dataGridView1.Columns.Contains("規格"))
                //{
                //    // 欄位存在
                //    dataGridView1.Columns["規格"].Width = 30;
                //}

            }
        }

        public void SEARCH_MANULINE(string QUERY, DataGridView DataGridViewNew, string SortedColumn, string SortedModel)
        {
            SqlConnection sqlConn = new SqlConnection();
            SqlDataAdapter SqlDataAdapterNEW = new SqlDataAdapter();
            SqlCommandBuilder SqlCommandBuilderNEW = new SqlCommandBuilder();
            DataSet DataSetNEW = new DataSet();

            DataGridViewNew.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                SqlDataAdapterNEW = new SqlDataAdapter(@"" + sbSql, sqlConn);

                SqlCommandBuilderNEW = new SqlCommandBuilder(SqlDataAdapterNEW);
                sqlConn.Open();
                DataSetNEW.Clear();
                SqlDataAdapterNEW.Fill(DataSetNEW, "DataSetNEW");
                sqlConn.Close();


                DataGridViewNew.DataSource = null;

                if (DataSetNEW.Tables["DataSetNEW"].Rows.Count >= 1)
                {
                    //DataGridViewNew.Rows.Clear();
                    DataGridViewNew.DataSource = DataSetNEW.Tables["DataSetNEW"];
                    DataGridViewNew.AutoResizeColumns();
                    //DataGridViewNew.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                    //DataGridViewNew.CurrentCell = dataGridView1[0, rownum];
                    //dataGridView20SORTNAME
                    //dataGridView20SORTMODE

                    if (!string.IsNullOrEmpty(SortedColumn))
                    {
                        if (SortedModel.Equals("Ascending"))
                        {
                            DataGridViewNew.Sort(DataGridViewNew.Columns["" + SortedColumn + ""], ListSortDirection.Ascending);
                        }
                        else
                        {
                            DataGridViewNew.Sort(DataGridViewNew.Columns["" + SortedColumn + ""], ListSortDirection.Descending);
                        }
                    }

                    //SET欄位寬度
                    if (DataGridViewNew.Columns.Contains("規格"))
                    {
                        // 欄位存在
                        DataGridViewNew.Columns["規格"].Width = 100;
                    }
                    if (DataGridViewNew.Columns.Contains("過敏原"))
                    {
                        // 欄位存在
                        DataGridViewNew.Columns["過敏原"].Width = 30;
                    }
                    if (DataGridViewNew.Columns.Contains("素別"))
                    {
                        // 欄位存在
                        DataGridViewNew.Columns["素別"].Width = 50;
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
        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            SEARCHMB001(textBox7.Text.Trim());

            SEARCHMOCMANULINETEMPDATAS(textBox7.Text.Trim());
        }

        public void SEARCHMB001(string MB001)
        {

            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlConnection sqlConn = new SqlConnection();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            DataSet ds1 = new DataSet();

            if (MANU.Equals("吧台烘焙線"))
            {
                
                try
                {
                    //20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);


                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    
                    sbSql.AppendFormat(@"  
                                        SELECT MB001,MB002,MB003,MC004 ,MB017 
                                        FROM [TK].dbo.INVMB,[TK].dbo.BOMMC
                                        WHERE MB001=MC001
                                        AND MB001='{0}'
                                        ", MB001);

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
                            textBox10.Text = ds1.Tables["ds1"].Rows[0]["MB002"].ToString();
                            textBox11.Text = ds1.Tables["ds1"].Rows[0]["MB003"].ToString();
                            textBox33.Text = ds1.Tables["ds1"].Rows[0]["MC004"].ToString();
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
        }
        public void SEARCHMOCMANULINETEMPDATAS(string MB001)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlConnection sqlConn = new SqlConnection();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            DataSet ds1 = new DataSet();
            DataSet TEMPds = new DataSet();

            decimal SUM21 = 0;
     

            if (MANU.Equals("吧台烘焙線"))
            {
                try
                {
                    //20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);


                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@" SELECT [ID]  FROM [TKMOC].[dbo].[MOCMANULINETEMP] WHERE [MB001]='{0}' AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINEBAKING] )", MB001);
                    sbSql.AppendFormat(@"  ");

                    adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                    sqlConn.Open();
                    ds1.Clear();
                    adapter1.Fill(ds1, "ds1");
                    sqlConn.Close();


                    if (ds1.Tables["ds1"] !=null && ds1.Tables["ds1"].Rows.Count >= 1)
                    {

                        TEMPds.Clear();
                        frmMOCMANULINESubTEMPADD MOCMANULINESubTEMPADD = new frmMOCMANULINESubTEMPADD(MB001, TEMPds);
                        MOCMANULINESubTEMPADD.ShowDialog();

                        TEMPds = MOCMANULINESubTEMPADD.SETDATASET;

                        if (TEMPds.Tables[0].Rows.Count >= 1)
                        {
                            foreach (DataRow dr in TEMPds.Tables[0].Rows)
                            {
                                SUM21 = SUM21 + Convert.ToDecimal(dr["包裝數"].ToString());
                                //SUM2 = SUM2 + Convert.ToDecimal(dr["箱數"].ToString());
                            }
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

        }

        public void ADDMOCMANULINE(
            string MANU,
            string MANUDATE,
            string MB001,
            string MB002,
            string MB003,
            string CLINET,
            string MANUHOUR,
            string BOX,
            string NUM,
            string PACKAGE,
            string OUTDATE,
            string TA029,
            string HALFPRO,
            string COPTD001,
            string COPTD002,
            string COPTD003
            )
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlConnection sqlConn = new SqlConnection();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            DataSet ds1 = new DataSet();
            DataSet TEMPds = new DataSet();

            if (MANU.Equals("吧台烘焙線"))
            {
                Guid NEWGUID = new Guid();
                NEWGUID = Guid.NewGuid();

                try
                {
                    //20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);


                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();


                    sbSql.AppendFormat(@" INSERT INTO [TKMOC].[dbo].[MOCMANULINEBAKING]
                                        ([ID],[MANU],[MANUDATE],[MB001],[MB002],[MB003],[CLINET],[MANUHOUR],[BOX],[NUM],[PACKAGE],[OUTDATE],[TA029],[HALFPRO],[COPTD001],[COPTD002],[COPTD003])
                                        VALUES ('{0}','{1}','{2}','{3}',N'{4}','{5}',N'{6}',N'{7}','{8}','{9}','{10}','{11}',N'{12}','{13}','{14}','{15}','{16}')"
                                        , NEWGUID.ToString()
                                        , MANU
                                        , MANUDATE
                                        , MB001
                                        , MB002
                                        , MB003
                                        , CLINET
                                        , MANUHOUR
                                        , BOX
                                        , NUM
                                        , PACKAGE
                                        , OUTDATE
                                        , TA029
                                        , HALFPRO
                                        , COPTD001
                                        , COPTD002
                                        , COPTD003
                                        );
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


            SEARCHMOCMANULINE_BAKING(dateTimePicker1.Value.ToString("yyyyMMdd"), comboBox1.Text.Trim());
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    textBoxID.Text = row.Cells["ID"].Value.ToString();

                    ID = row.Cells["ID"].Value.ToString();
                    dt1 = Convert.ToDateTime(row.Cells["生產日"].Value.ToString().Substring(0, 4) + "/" + row.Cells["生產日"].Value.ToString().Substring(4, 2) + "/" + row.Cells["生產日"].Value.ToString().Substring(6, 2));
                    MB001B = row.Cells["品號"].Value.ToString();
                    MB002B = row.Cells["品名"].Value.ToString();
                    MB003B = row.Cells["規格"].Value.ToString();
                    BOX = Convert.ToDecimal(row.Cells["包裝數"].Value.ToString());
                    SUM2 = Convert.ToDecimal(row.Cells["包裝數"].Value.ToString());
                    TA029 = row.Cells["備註"].Value.ToString();
                    TA026 = row.Cells["訂單單別"].Value.ToString();
                    TA027 = row.Cells["訂單號"].Value.ToString();
                    TA028 = row.Cells["訂單序號"].Value.ToString();

                    SUBID = row.Cells["ID"].Value.ToString();
                    SUBBAR2 = "";
                    SUBNUM2 = "";
                    SUBBOX = row.Cells["箱數"].Value.ToString();
                    SUBPACKAGE2= row.Cells["包裝數"].Value.ToString();

                    //SEARCHMOCMANULINERESULT();
                    //SEARCHMOCMANULINEMERGERESLUTMOCTA(ID2.ToString());
                    ////SEARCHMOCMANULINECOP();

                }
                else
                {
                    //ID2 = null;
                    //SUBID2 = null;
                    //SUBBAR2 = null;
                    //SUBNUM2 = null;
                    //SUBBOX2 = null;
                    //SUBPACKAGE2 = null;
                    //TA026 = null;
                    //TA027 = null;
                    //TA028 = null;

                }
            }
        }

        public void DELMOCMANULINE(string ID)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlConnection sqlConn = new SqlConnection();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            DataSet ds1 = new DataSet();
            DataSet TEMPds = new DataSet();

            if (MANU.Equals("吧台烘焙線"))
            {
                try
                {
                    //20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);


                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();
                    sbSql.AppendFormat(@"  
                                        DELETE [TKMOC].[dbo].[MOCMANULINEBAKING]
                                        WHERE ID='{0}'"
                                        , ID);
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


            SEARCHMOCMANULINE_BAKING(dateTimePicker1.Value.ToString("yyyyMMdd"), comboBox1.Text.Trim());
        }

        public void CHECKMOCTAB()
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlConnection sqlConn = new SqlConnection();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            DataSet ds1 = new DataSet();
            DataSet TEMPds = new DataSet();
            string CHECKID = null;

            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage1"])
            {
                CHECKID = textBoxID.Text.ToString().Trim();
            }
          
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  
                                    SELECT	MOCTA001,MOCTA002
                                    FROM  [TKMOC].[dbo].[MOCMANULINERESULTBAKING]
                                    WHERE [SID]='{0}'
                                    UNION ALL
                                    SELECT	TA001,TA002
                                    FROM [TK].[dbo].[MOCTA]
                                    WHERE EXISTS (SELECT [MOCTA001],[MOCTA002] FROM [TKMOC].[dbo].[MOCMANULINERESULTBAKING] WHERE [SID]='{0}' AND TA001=MOCTA001 AND TA002=MOCTA002)"
                                    , CHECKID);
                sbSql.AppendFormat(@"  ");


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count == 0)
                {
                    UPDATEMOCMANULINE(CHECKID);
                }
                else
                {
                    MessageBox.Show("ERP跟外中 有製令未刪除，請檢查一下");
                }

            }
            catch
            {

            }
            finally
            {

            }
        }
        public void UPDATEMOCMANULINE(string CHECKID)
        {
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage1"])
            {
                frmMOCMANULINES_BAKING_SUB MOCMANULINE_BAKUING_Sub = new frmMOCMANULINES_BAKING_SUB(CHECKID);
                MOCMANULINE_BAKUING_Sub.ShowDialog();
            }

        }

        public void SEARCHCOPDEFAULT(string TD001, string TD002, string TD003)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlConnection sqlConn = new SqlConnection();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            DataSet ds1 = new DataSet();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  
                                    SELECT TC053,TD004,TD005,TD006,(TD008+TD024) AS TD008,TD010,(TC015+'-'+TD020) TC015 ,TD013
                                    ,(CASE WHEN ISNULL(MD002,'')<>'' THEN (TD008+TD024)*MD004 ELSE (TD008+TD024)  END ) AS NUM
                                    FROM [TK].dbo.INVMB WITH(NOLOCK),[TK].dbo.COPTC WITH(NOLOCK),[TK].dbo.COPTD WITH(NOLOCK)
                                    LEFT JOIN [TK].dbo.INVMD ON MD001=TD004 AND TD010=MD002
                                    WHERE TC001=TD001 AND TC002=TD002
                                    AND MB001=TD004
                                    AND TD001='{0}' AND TD002='{1}' AND TD003='{2}'"
                                    , TD001, TD002, TD003);
                sbSql.AppendFormat(@"  ");


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (MANU.Equals("吧台烘焙線"))
                {
                    if (ds1.Tables["ds1"].Rows.Count == 0)
                    {
                        textBox7.Text = null;
                        textBox10.Text = null;
                        textBox11.Text = null;
                        textBox12.Text = null;
                        textBox53.Text = null;
                        textBox9.Text = null;
                        textBox42.Text = null;
                        textBox43.Text = null;
                        textBox72.Text = null;
                    }
                    else
                    {
                        if (ds1.Tables["ds1"].Rows.Count >= 1)
                        {
                            textBox7.Text = ds1.Tables["ds1"].Rows[0]["TD004"].ToString();
                            textBox10.Text = ds1.Tables["ds1"].Rows[0]["TD005"].ToString();
                            textBox11.Text = ds1.Tables["ds1"].Rows[0]["TD006"].ToString();
                            textBox9.Text = ds1.Tables["ds1"].Rows[0]["TC053"].ToString();
                            textBox53.Text = ds1.Tables["ds1"].Rows[0]["TC015"].ToString();
                            dateTimePicker5.Value = Convert.ToDateTime(ds1.Tables["ds1"].Rows[0]["TD013"].ToString().Substring(0, 4) + "/" + ds1.Tables["ds1"].Rows[0]["TD013"].ToString().Substring(4, 2) + "/" + ds1.Tables["ds1"].Rows[0]["TD013"].ToString().Substring(6, 2));

                            textBox12.Text = ds1.Tables["ds1"].Rows[0]["NUM"].ToString();

                            //if (SUM21 > 0)
                            //{
                            //    textBox12.Text = (SUM21 + Convert.ToDecimal(ds27.Tables["ds27"].Rows[0]["NUM"].ToString())).ToString();

                                //    SUM21 = 0;
                                //}
                                //else
                                //{
                                //    textBox12.Text = ds1.Tables["ds1"].Rows[0]["NUM"].ToString();
                                //}
                        }
                    }
                }
            }
            catch
            { }
            finally
            { }
               
        }

        public void SEARCHCOPDEFAULT2(string TD001, string TD002, string TD003)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlConnection sqlConn = new SqlConnection();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            DataSet ds1 = new DataSet();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();
                //手工*INVMB.UDF08、其他*INVMB.UDF07



                sbSql.AppendFormat(@"  
                                    SELECT TC053,TD004,TD005,TD006,(TD008+TD024) AS TD008,TD010,TC015
                                    ,(CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024)  END ) AS NUM
                                    ,BOMMD.MD003,BOMMD.MD035,BOMMD.MD036,INVMB.UDF07
                                     ,((CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024)  END ))/BOMMC.MC004*BOMMD.MD006 AS 'NUM2'
                                    
                                    FROM [TK].dbo.INVMB WITH(NOLOCK),[TK].dbo.COPTC WITH(NOLOCK),[TK].dbo.COPTD WITH(NOLOCK)
                                    LEFT JOIN [TK].dbo.INVMD WITH(NOLOCK) ON INVMD.MD001=TD004 AND INVMD.MD002=TD010
                                    LEFT JOIN [TK].dbo.BOMMC WITH(NOLOCK) ON BOMMC.MC001=TD004 
                                    LEFT JOIN [TK].dbo.BOMMD WITH(NOLOCK) ON BOMMD.MD001=TD004 
                                    WHERE TC001=TD001 AND TC002=TD002
                                    AND MB001 = TD004
                                    AND(BOMMD.MD003 LIKE '3%' OR BOMMD.MD003 LIKE '4%')
                                    AND TD001 = '{0}' AND TD002 = '{1}' AND TD003 = '{2}'

                                    ", TD001, TD002, TD003);
                sbSql.AppendFormat(@"  ");

                //半成品的舊算法
                //sbSql.AppendFormat(@"  ,((CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024)  END ))/BOMMC.MC004*INVMB.UDF07/1000 AS 'NUM2'");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();



                if (MANU.Equals("吧台烘焙線"))
                {
                    if (ds1.Tables["ds1"].Rows.Count == 0)
                    {
                        textBox7.Text = null;
                        textBox10.Text = null;
                        textBox11.Text = null;
                        textBox12.Text = null;
                        textBox9.Text = null;
                        textBox53.Text = null;
                        textBox42.Text = null;
                        textBox43.Text = null;
                        textBox72.Text = null;
                    }
                    else
                    {
                        if (ds1.Tables["ds1"].Rows.Count >= 1)
                        {
                            textBox7.Text = ds1.Tables["ds1"].Rows[0]["MD003"].ToString();
                            textBox10.Text = ds1.Tables["ds1"].Rows[0]["MD035"].ToString();
                            textBox11.Text = ds1.Tables["ds1"].Rows[0]["MD036"].ToString();
                            textBox12.Text = ds1.Tables["ds1"].Rows[0]["NUM2"].ToString();
                            textBox9.Text = ds1.Tables["ds1"].Rows[0]["TC053"].ToString();
                            textBox53.Text = ds1.Tables["ds1"].Rows[0]["TC015"].ToString();

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

        public void SEARCHCOPDEFAULT3(string TD001, string TD002, string TD003)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlConnection sqlConn = new SqlConnection();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            DataSet ds1 = new DataSet();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();
                //手工*INVMB.UDF08、其他*INVMB.UDF07
                sbSql.AppendFormat(@"  
                                    SELECT TC053,TD004,TD005,TD006,(TD008+TD024) AS TD008,TD010,TC015
                                    ,(CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024)  END ) AS NUM
                                    ,BOMMD.MD003,BOMMD.MD035,BOMMD.MD036,INVMB.UDF07
                                    ,((CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024)  END ))/BOMMC.MC004*BOMMD.MD006 AS 'NUM2'

                                    FROM [TK].dbo.INVMB WITH(NOLOCK),[TK].dbo.COPTC WITH(NOLOCK),[TK].dbo.COPTD WITH(NOLOCK)
                                    LEFT JOIN [TK].dbo.INVMD WITH(NOLOCK) ON INVMD.MD001=TD004 AND INVMD.MD002=TD010
                                    LEFT JOIN [TK].dbo.BOMMC WITH(NOLOCK) ON BOMMC.MC001=TD004 
                                    LEFT JOIN [TK].dbo.BOMMD WITH(NOLOCK) ON BOMMD.MD001=TD004 
                                    WHERE TC001=TD001 AND TC002=TD002
                                    AND MB001=TD004
                                    AND (BOMMD.MD003 LIKE '3%' OR BOMMD.MD003 LIKE '4%') 
                                    AND TD001='{0}' AND TD002='{1}' AND TD003='{2}'
                                    ", TD001, TD002, TD003);
                sbSql.AppendFormat(@"  ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();
                

                if (MANU.Equals("吧台烘焙線"))
                {
                    if (ds1.Tables["ds1"].Rows.Count == 0)
                    {
                        textBox7.Text = null;
                        textBox10.Text = null;
                        textBox11.Text = null;
                        textBox12.Text = null;
                        textBox9.Text = null;
                        textBox53.Text = null;
                        textBox42.Text = null;
                        textBox43.Text = null;
                        textBox72.Text = null;
                    }
                    else
                    {
                        if (ds1.Tables["ds1"].Rows.Count >= 1)
                        {
                            textBox7.Text = ds1.Tables["ds1"].Rows[0]["MD003"].ToString();
                            textBox10.Text = ds1.Tables["ds1"].Rows[0]["MD035"].ToString();
                            textBox11.Text = ds1.Tables["ds1"].Rows[0]["MD036"].ToString();
                            textBox12.Text = ds1.Tables["ds1"].Rows[0]["NUM2"].ToString();
                            textBox9.Text = ds1.Tables["ds1"].Rows[0]["TC053"].ToString();
                            textBox53.Text = null;
                            //textBox53.Text = ds28.Tables["ds28"].Rows[0]["TC015"].ToString();

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

        public string GETMAXTA002(string TA001,DateTime DT)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlConnection sqlConn = new SqlConnection();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            DataSet ds1 = new DataSet();
            string TA002;

            if (MANU.Equals("吧台烘焙線"))
            {

                try
                {
                    //20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();
                    ds1.Clear();

                    sbSql.AppendFormat(@" 
                                        SELECT ISNULL(MAX(TA002),'00000000000') AS TA002
                                        FROM [TK].[dbo].[MOCTA]
                                        WHERE  TA001='{0}' AND TA002 LIKE '%{1}%' 
                                        ", TA001, DT.ToString("yyyyMMdd"));

                    adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                    sqlConn.Open();
                    ds1.Clear();
                    adapter1.Fill(ds1, "ds1");
                    sqlConn.Close();


                    if (ds1.Tables["ds1"].Rows.Count == 0)
                    {
                        return null;
                    }
                    else
                    {
                        if (ds1.Tables["ds1"].Rows.Count >= 1)
                        {
                            TA002 = SETTA002(ds1.Tables["ds1"].Rows[0]["TA002"].ToString(), DT);
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
        public string SETTA002(string TA002, DateTime DT)
        {

            if (MANU.Equals("吧台烘焙線"))
            {
                if (TA002.Equals("00000000000"))
                {
                    return DT.ToString("yyyyMMdd") + "001";
                }

                else
                {
                    int serno = Convert.ToInt16(TA002.Substring(8, 3));
                    serno = serno + 1;
                    string temp = serno.ToString();
                    temp = temp.PadLeft(3, '0');
                    return DT.ToString("yyyyMMdd") + temp.ToString();
                }
            }

            return null;
        }

        public void ADDMOCMANULINERESULT(string ID,string TA001,string TA002)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlConnection sqlConn = new SqlConnection();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            DataSet ds1 = new DataSet();

            if (MANU.Equals("吧台烘焙線"))
            {
                try
                {
                    //20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);


                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();


                    
                    sbSql.AppendFormat(@" 
                                        INSERT INTO[TKMOC].[dbo].[MOCMANULINERESULTBAKING]
                                        ([SID],[MOCTA001],[MOCTA002])
                                        VALUES('{0}', '{1}', '{2}')
                                        ", ID, TA001, TA002);
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

        public void ADDMOCTATB(DateTime DT)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlConnection sqlConn = new SqlConnection();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            DataSet ds1 = new DataSet();

            MOCTADATA MOCTA = new MOCTADATA();
            MOCTA = SETMOCTA(DT);
            string MOCMB001 = null;
            decimal MOCTA004 = 0; ;
            string MOCTB009 = null;


            const int MaxLength = 100;

            if (MANU.Equals("吧台烘焙線"))
            {
                MOCMB001 = MB001B;
                MOCTA004 = BOX;
                MOCTA.TA026 = TA026;
                MOCTA.TA027 = TA027;
                MOCTA.TA028 = TA028;
                //MOCTB009 = textBox78.Text;

            }
           

          
            try
            {
                //check TA002=2,TA040=2
                //[TB004]的計算，如果領用倍數MB041=1且不是201開頭的箱子，就取整數、MB041=1且是201開頭的箱子，就4捨5入到整數、其他就取到小數第3位
                if (MOCTA.TA002.Substring(0, 1).Equals("2") && MOCTA.TA040.Substring(0, 1).Equals("2"))
                {
                    //20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);


                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();

                    sbSql.AppendFormat(@"
                                        INSERT INTO [TK].[dbo].[MOCTA]
                                        ([COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE]
                                        ,[TRANS_NAME],[sync_count],[DataGroup],[TA001],[TA002],[TA003],[TA004],[TA005],[TA006],[TA007]
                                        ,[TA009],[TA010],[TA011],[TA012],[TA013],[TA014],[TA015],[TA016],[TA017],[TA018]
                                        ,[TA019],[TA020],[TA021],[TA022],[TA024],[TA025],[TA029],[TA030],[TA031],[TA034],[TA035]
                                        ,[TA040],[TA041],[TA043],[TA044],[TA045],[TA046],[TA047],[TA049],[TA050],[TA200]
                                        ,[TA026],[TA027],[TA028])
                                        VALUES
                                        ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}',
                                        '{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}','{19}',
                                        '{20}','{21}','{22}','{23}','{24}','{25}','{26}','{27}','{28}','{29}',
                                        '{30}','{31}','{32}','{33}','{34}','{35}',N'{36}','{37}','{38}','{39}',
                                        '{40}','{41}','{42}','{43}','{44}','{45}','{46}','{47}','{48}','{49}',
                                        '{50}','{51}','{52}')
    
                                        INSERT INTO [TK].dbo.[MOCTB]
                                        ([COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE]
                                        ,[TRANS_NAME],[sync_count],[DataGroup],[TB001],[TB002],[TB003],[TB004],[TB005],[TB006],[TB007]
                                        ,[TB009],[TB011],[TB012],[TB013],[TB014],[TB018],[TB019],[TB020],[TB022],[TB024]
                                        ,[TB025],[TB026],[TB027],[TB029],[TB030],[TB031],[TB501],[TB554],[TB556],[TB560])
                                        SELECT 
                                        '{0}' [COMPANY],'{1}' [CREATOR],'{2}' [USR_GROUP],'{3}' [CREATE_DATE],'{4}' [MODIFIER],'{5}' [MODI_DATE],'{6}' [FLAG],'{7}' [CREATE_TIME],'{8}' [MODI_TIME],'{9}' [TRANS_TYPE],
                                        '{10}' [TRANS_NAME],{11} [sync_count],'{12}' [DataGroup],'{13}' [TB001],'{14}' [TB002],[BOMMD].MD003 [TB003],
                                        CASE WHEN MB041=1 AND [BOMMD].MD003 NOT LIKE '201%' THEN CONVERT(decimal(16,4),CEILING({15}*[BOMMD].MD006/[BOMMD].MD007*(1+[BOMMD].MD008))) 
                                             WHEN MB041=1 AND [BOMMD].MD003 LIKE '201%' THEN ROUND({15}*[BOMMD].MD006/[BOMMD].MD007*(1+[BOMMD].MD008),0) 
                                             ELSE ROUND({15}*[BOMMD].MD006/[BOMMD].MD007*(1+[BOMMD].MD008),3) END [TB004],0 [TB005],'****' [TB006],[INVMB].MB004 [TB007],
                                        [INVMB].MB017 [TB009],'1' [TB011],[INVMB].MB002 [TB012],[INVMB].MB003 [TB013],[BOMMD].MD001 [TB014],'N' [TB018],'0' [TB019],'0' [TB020],'2' [TB022],'0' [TB024],
                                        '****' [TB025],'0' [TB026],'1' [TB027],'0' [TB029],'0' [TB030],'0' [TB031],'0' [TB501],'N' [TB554],'0' [TB556],'0' [TB560]
                                        FROM [TK].dbo.[BOMMD],[TK].dbo.[INVMB]
                                        WHERE [BOMMD].MD003=[INVMB].MB001
                                        AND MD001='{16}' AND ISNULL(MD012,'')=''
                                    ",
                                        MOCTA.COMPANY, MOCTA.CREATOR, MOCTA.USR_GROUP, MOCTA.CREATE_DATE, MOCTA.MODIFIER, MOCTA.MODI_DATE, MOCTA.FLAG, MOCTA.CREATE_TIME, MOCTA.MODI_TIME, MOCTA.TRANS_TYPE,
                                        MOCTA.TRANS_NAME, MOCTA.sync_count, MOCTA.DataGroup, MOCTA.TA001, MOCTA.TA002, MOCTA.TA003, MOCTA.TA004, MOCTA.TA005, MOCTA.TA006, MOCTA.TA007,
                                        MOCTA.TA009, MOCTA.TA010, MOCTA.TA011, MOCTA.TA012, MOCTA.TA013, MOCTA.TA014, MOCTA.TA015, MOCTA.TA016, MOCTA.TA017, MOCTA.TA018,
                                        MOCTA.TA019, MOCTA.TA020, MOCTA.TA021, MOCTA.TA022, MOCTA.TA024, MOCTA.TA025, MOCTA.TA029, MOCTA.TA030, MOCTA.TA031, MOCTA.TA034, MOCTA.TA035,
                                        MOCTA.TA040, MOCTA.TA041, MOCTA.TA043, MOCTA.TA044, MOCTA.TA045, MOCTA.TA046, MOCTA.TA047, MOCTA.TA049, MOCTA.TA050, MOCTA.TA200,
                                        MOCTA.TA026, MOCTA.TA027, MOCTA.TA028, MOCTA.TA004, MOCMB001);

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

        public MOCTADATA SETMOCTA(DateTime DT)
        {
            string BOMVARSION="";
            string UNIT = "";
            decimal BOMBAR = 0;
            //硯微墨-烘焙倉
            string IN = "21002";

            if (MANU.Equals("吧台烘焙線"))
            {
                DataTable DATATABLE=SEARCHBOMMC();
                
                if (DATATABLE != null && DATATABLE.Rows.Count>=1)
                {
                    BOMVARSION = ds1.Tables["ds1"].Rows[0]["MC009"].ToString();               
                    UNIT = ds1.Tables["ds1"].Rows[0]["MB004"].ToString();
                    BOMBAR = Convert.ToDecimal(ds1.Tables["ds1"].Rows[0]["MC004"].ToString());
                }

                MOCTADATA MOCTA = new MOCTADATA();
                MOCTA.COMPANY = "TK";
                MOCTA.CREATOR = "140020";
                MOCTA.USR_GROUP = "103000";
                //MOCTA.CREATE_DATE = dt2.ToString("yyyyMMdd");
                MOCTA.CREATE_DATE = DateTime.Now.ToString("yyyyMMdd");
                MOCTA.MODIFIER = "140020";
                MOCTA.MODI_DATE = DT.ToString("yyyyMMdd");
                MOCTA.FLAG = "0";
                MOCTA.CREATE_TIME = DateTime.Now.ToString("HH:mm:dd");
                MOCTA.MODI_TIME = DateTime.Now.ToString("HH:mm:dd");
                MOCTA.TRANS_TYPE = "P001";
                MOCTA.TRANS_NAME = "MOCMI02";
                MOCTA.sync_count = "0";
                MOCTA.DataGroup = "103000";
                MOCTA.TA001 = "A510";
                MOCTA.TA002 = TA002;
                MOCTA.TA003 = DT.ToString("yyyyMMdd");
                MOCTA.TA004 = DT.ToString("yyyyMMdd");
                MOCTA.TA005 = BOMVARSION;
                MOCTA.TA006 = MB001B;
                MOCTA.TA007 = UNIT;
                MOCTA.TA009 = DT.ToString("yyyyMMdd");
                MOCTA.TA010 = DT.ToString("yyyyMMdd");
                MOCTA.TA011 = "1";
                MOCTA.TA012 = DT.ToString("yyyyMMdd");
                MOCTA.TA013 = "N";
                // MOCTA.TA014 = dt2.ToString("yyyyMMdd");
                MOCTA.TA014 = "";
                //MOCTA.TA015 = (BOX * BOMBAR).ToString();
                MOCTA.TA015 = SUM2.ToString();
                MOCTA.TA016 = "0";
                MOCTA.TA017 = "0";
                MOCTA.TA018 = "0";
                MOCTA.TA019 = "20";
                MOCTA.TA020 = IN;
                MOCTA.TA021 = "09";
                MOCTA.TA022 = "0";
                MOCTA.TA024 = "A510";
                MOCTA.TA025 = TA002;
                MOCTA.TA029 = TA029;
                MOCTA.TA030 = "1";
                MOCTA.TA031 = "0";
                MOCTA.TA034 = MB002B;
                MOCTA.TA035 = MB003B;
                MOCTA.TA040 = DT.ToString("yyyyMMdd");
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
            


            return null;

        }

        public DataTable SEARCHBOMMC()
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlConnection sqlConn = new SqlConnection();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            DataSet ds1 = new DataSet();

            string BOMVARSION = null;
            string UNIT = null;
            decimal BOMBAR = 0;
            
            if (MANU.Equals("吧台烘焙線"))
            {
                try
                {
                    //20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);


                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"
                                        SELECT 
                                        [MC001],[MC002],[MC003],[MC004],[MC005],[MC006],[MC007],[MC008],[MC009],[MC010],
                                        [MC011],[MC012],[MC013],[MC014],[MC015],[MC016],[MC017],[MC018],[MC019],[MC020],
                                        [MC021],[MC022],[MC023],[MC024],[MC025],[MC026],[MC027],
                                        INVMB.MB004
                                        FROM [TK].[dbo].[BOMMC]
                                        LEFT JOIN [TK].dbo.[INVMB] ON MB001=MC001
                                        WHERE [MC001]='{0}'", MB001B);

                    sbSql.AppendFormat(@"  ");

                    adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                    sqlConn.Open();
                    ds1.Clear();
                    adapter1.Fill(ds1, "ds1");
                    sqlConn.Close();

                    if (ds1 != null && ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        return ds1.Tables["ds1"];                       

                    }
                    else
                    {
                        return null;
                    }

                }
                catch
                {
                    return null;
                }
                finally
                {

                }
            }
            else
            {
                return null;
            }

        }
        public void SETNULL()
        {
            textBox7.Text = null;
            textBox8.Text = "0";
            textBox9.Text = null;
            textBox10.Text = null;
            textBox11.Text = null;
            textBox12.Text = null;
            textBox13.Text = "0";
            textBox33.Text = "0";
            textBox53.Text = null;
            textBox68.Text = "0";
            textBox42.Text = null;
            textBox43.Text = null;
            textBox72.Text = null;
        }


        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            SEARCHMOCMANULINE_BAKING(dateTimePicker1.Value.ToString("yyyyMMdd"),comboBox1.Text.Trim());
        }
        private void button2_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox7.Text))
            {
                string MANU = comboBox2.Text.ToString().Trim();
                string MANUDATE = dateTimePicker4.Value.ToString("yyyy/MM/dd");
                string MB001 = textBox7.Text.ToString().Trim();
                string MB002 = textBox10.Text.ToString().Trim();
                string MB003 = textBox11.Text.ToString().Trim();
                string CLINET = textBox9.Text.ToString().Trim();
                string MANUHOUR = textBox13.Text.ToString().Trim();
                string BOX = textBox8.Text.ToString().Trim();
                string NUM = textBox12.Text.ToString().Trim();
                string PACKAGE = textBox12.Text.ToString().Trim();
                string OUTDATE = dateTimePicker5.Value.ToString("yyyy/MM/dd");
                string TA029 = textBox53.Text.Replace("'", "");
                string HALFPRO = textBox68.Text.ToString().Trim();
                string COPTD001 = textBox42.Text.ToString().Trim();
                string COPTD002 = textBox43.Text.ToString().Trim();
                string COPTD003 = textBox72.Text.ToString().Trim();

                ADDMOCMANULINE(
                    MANU,
                    MANUDATE,
                    MB001,
                    MB002,
                    MB003,
                    CLINET,
                    MANUHOUR,
                    BOX,
                    NUM,
                    PACKAGE,
                    OUTDATE,
                    TA029,
                    HALFPRO,
                    COPTD001,
                    COPTD002,
                    COPTD003
                    );

                SETNULL();
            }
            else
            {
                MessageBox.Show("品名錯誤");
            }
        }
        private void button5_Click(object sender, EventArgs e)
        {
            SETNULL();

            frmSUBMOCMANULINE SUBfrmSUBMOCMANULINE = new frmSUBMOCMANULINE();
            SUBfrmSUBMOCMANULINE.ShowDialog();
            textBox7.Text = SUBfrmSUBMOCMANULINE.TextBoxMsg;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            frmSUBMOCCOPMA SUBfrmSUBMOCCOPMA = new frmSUBMOCCOPMA();
            SUBfrmSUBMOCCOPMA.ShowDialog();
            textBox9.Text = SUBfrmSUBMOCCOPMA.TextBoxMsg;
        }
        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELMOCMANULINE(textBoxID.Text.ToString().Trim());
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            CHECKMOCTAB();
            SEARCHMOCMANULINE_BAKING(dateTimePicker1.Value.ToString("yyyyMMdd"), comboBox1.Text.Trim());
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox42.Text) & !string.IsNullOrEmpty(textBox43.Text) & !string.IsNullOrEmpty(textBox72.Text))
            {
                SEARCHCOPDEFAULT(textBox42.Text, textBox43.Text, textBox72.Text);
            }
        }
        private void button8_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox42.Text) & !string.IsNullOrEmpty(textBox43.Text) & !string.IsNullOrEmpty(textBox72.Text))
            {
                SEARCHCOPDEFAULT2(textBox42.Text, textBox43.Text, textBox72.Text);
            }
        }
        private void button9_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox42.Text) & !string.IsNullOrEmpty(textBox43.Text) & !string.IsNullOrEmpty(textBox72.Text))
            {
                SEARCHCOPDEFAULT3(textBox42.Text, textBox43.Text, textBox72.Text);
            }
        }
        private void button10_Click(object sender, EventArgs e)
        {
            DateTime DT = new DateTime();

            if (!string.IsNullOrEmpty(TA028))
            {
                //指定日期=生產日
                DT = dt1;
                TA002 = GETMAXTA002(TA001, DT);
                ADDMOCMANULINERESULT(textBoxID.Text.ToString().Trim(), TA001, TA002);
                ADDMOCTATB(DT);

                SEARCHMOCMANULINE_BAKING(dateTimePicker1.Value.ToString("yyyyMMdd"), comboBox1.Text.Trim());

                MessageBox.Show("完成");
            }
            else
            {
                MessageBox.Show("訂單沒有指定");
            }
        }

        #endregion


    }
}
