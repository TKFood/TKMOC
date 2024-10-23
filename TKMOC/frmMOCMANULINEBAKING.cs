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
            comboBox1load();
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

        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            SEARCHMOCMANULINE_BAKING(dateTimePicker1.Value.ToString("yyyyMMdd"),comboBox1.Text.Trim());
        }
        #endregion

       
    }
}
