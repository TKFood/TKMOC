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

        string MANU = "";
        // 宣告一個變數來儲存使用者手動選擇排序的欄位
        string SortedColumn = string.Empty;
        string SortedModel = string.Empty;

        public frmMOCMANULINEBAKING()
        {
            InitializeComponent();
        }

        #region FUNCTION
        private void frmMOCMANULINEBAKING_Load(object sender, EventArgs e)
        {
            MANU = "吧台烘焙線";

            comboBox1load();
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
                                    ,[MOCMANULINE].[MB001] AS '品號'
                                    ,[MOCMANULINE].[MB002] AS '品名' 
                                    ,[MOCMANULINE].[MB003] AS '規格'
                                    ,ALLERGEN AS '過敏原'
                                    ,ORI AS '素別'
                                    ,[BAR] AS '桶數'
                                    ,[NUM] AS '數量'
                                    ,[CLINET] AS '客戶'
                                    ,[OUTDATE] AS '交期'
                                    ,[TA029] AS '備註'
                                    ,[HALFPRO] AS '半成品數量'
                                    ,[COPTD001] AS '訂單單別'
                                    ,[COPTD002] AS '訂單號'
                                    ,[COPTD003] AS '訂單序號'
                                    ,[BOX] AS '箱數'
                                    ,[ID]
                                    FROM [TKMOC].[dbo].[MOCMANULINE]
                                    LEFT JOIN [TKMOC].[dbo].[ERPINVMB] ON [ERPINVMB].MB001=[MOCMANULINE].MB001

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

                    sbSql.AppendFormat(@" SELECT [ID]  FROM [TKMOC].[dbo].[MOCMANULINETEMP] WHERE [MB001]='{0}' AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE] )", MB001);
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

        public void SETNULL()
        {
            textBox7.Text = null;
            textBox8.Text = null;
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

        }


        #endregion

       
    }
}
