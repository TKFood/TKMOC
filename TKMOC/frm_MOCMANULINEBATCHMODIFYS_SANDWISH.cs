using FastReport.DevComponents.DotNetBar.Controls;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using System.Windows.Forms;
using TKITDLL;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace TKMOC
{
    public partial class frm_MOCMANULINEBATCHMODIFYS_SANDWISH : Form
    {
        // 宣告一個變數來儲存使用者手動選擇排序的欄位
        string SortedColumn = string.Empty;
        string SortedModel = string.Empty;

        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();
        int result;
        public frm_MOCMANULINEBATCHMODIFYS_SANDWISH()
        {
            InitializeComponent();
        }

        #region FUNCTION 
        public void SEARCH()
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();

            sbSql.Clear();
            sbSqlQuery.Clear();

            sbSql.AppendFormat(@"   
                                SELECT 
                                [MB001] AS '主品號'
                                ,[MB002] AS '主品名'
                                ,[MODIFY_MB001] AS '連動修改品號'
                                ,[MODIFY_MB002] AS '連動修改品名'
                                FROM [TKMOC].[dbo].[MOCMANULINEBATCHMODIFYS_SANDWISH]

                                    ");
            sbSql.AppendFormat(@"  ");

            SEARCH_DataGridView(sbSql.ToString(), dataGridView1, SortedColumn, SortedModel);

        }

        public void SEARCH_DataGridView(string QUERY, DataGridView DataGridViewNew, string SortedColumn, string SortedModel)
        {
            SqlConnection sqlConn = new SqlConnection();
            SqlDataAdapter SqlDataAdapterNEW = new SqlDataAdapter();
            SqlCommandBuilder SqlCommandBuilderNEW = new SqlCommandBuilder();
            DataSet DataSetNEW = new DataSet();
            StringBuilder sbSql = new StringBuilder();
            DataGridViewNew.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;

            sbSql = new StringBuilder(QUERY);

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

                }

            }
            catch (Exception ex)
            {
                //MessageBox.Show($"發生錯誤：{ex.Message}");
            }
            finally
            {
                sqlConn.Close();
            }
        }
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            textBox5.Text = null;
            textBox6.Text = null;

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    textBox5.Text = row.Cells["主品號"].Value.ToString();
                    textBox6.Text = row.Cells["連動修改品號"].Value.ToString();
                }
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            textBox2.Text = null;

            DataTable DT=FIND_INVMB_MB002(textBox1.Text);
            if(DT!=null && DT.Rows.Count>=1)
            {
                textBox2.Text = DT.Rows[0]["品名"].ToString();
            }
            
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
        }
        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            textBox4.Text = null;

            DataTable DT = FIND_INVMB_MB002(textBox3.Text);
            if (DT != null && DT.Rows.Count >= 1)
            {
                textBox4.Text = DT.Rows[0]["品名"].ToString();
            }
        }
        public DataTable FIND_INVMB_MB002(string MB001)
        {
            SqlConnection sqlConn = new SqlConnection();
            SqlDataAdapter SqlDataAdapterNEW = new SqlDataAdapter();
            SqlCommandBuilder SqlCommandBuilderNEW = new SqlCommandBuilder();
            DataSet DataSetNEW = new DataSet();     

            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();

            sbSql.Clear();
            sbSqlQuery.Clear();
            sbSql.AppendFormat(@"   
                                SELECT 
                                [MB001] AS '品號'
                                ,[MB002] AS '品名'
                                FROM [TK].dbo.INVMB
                                WHERE MB001='{0}'
                                    ", MB001);

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

                if(DataSetNEW.Tables[0]!=null && DataSetNEW.Tables[0].Rows.Count>=1)
                {
                    return DataSetNEW.Tables[0];
                }
                else
                {
                    return null;
                }

            }
            catch (Exception ex)
            {
                return null;
                //MessageBox.Show($"發生錯誤：{ex.Message}");
            }
            finally
            {
            }
        }


        public void ADD_MOCMANULINEBATCHMODIFYS_SANDWISH(string MB001, string MB002, string MODIFY_MB001, string MODIFY_MB002)
        {
            try
            {
                Class1 TKID = new Class1();
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                using (SqlConnection sqlConn = new SqlConnection(sqlsb.ConnectionString))
                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.Connection = sqlConn;
                    cmd.CommandTimeout = 60;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = @"INSERT INTO [TKMOC].[dbo].[MOCMANULINEBATCHMODIFYS_SANDWISH]
                                        (
                                        [MB001]
                                        , [MB002]
                                        , [MODIFY_MB001]
                                        , [MODIFY_MB002]
                                        )
                                        VALUES
                                        (
                                        @MB001
                                        , @MB002
                                        , @MODIFY_MB001
                                        , @MODIFY_MB002
                                        )";

                    cmd.Parameters.AddWithValue("@MB001", MB001 ?? string.Empty);
                    cmd.Parameters.AddWithValue("@MB002", MB002 ?? string.Empty);
                    cmd.Parameters.AddWithValue("@MODIFY_MB001", MODIFY_MB001 ?? string.Empty);
                    cmd.Parameters.AddWithValue("@MODIFY_MB002", MODIFY_MB002 ?? string.Empty);

                    if (tran != null)
                        cmd.Transaction = tran;

                    sqlConn.Open();
                    result = cmd.ExecuteNonQuery();

                    if (result == 0 && tran != null)
                    {
                        tran.Rollback();
                    }
                    else if (tran != null)
                    {
                        tran.Commit();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"發生錯誤：{ex.Message}");
                if (tran != null)
                    tran.Rollback();
            }
        }

        public void DEETE_MOCMANULINEBATCHMODIFYS_SANDWISH(string MB001, string MODIFY_MB001)
        {

            try
            {
                Class1 TKID = new Class1();
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                using (SqlConnection sqlConn = new SqlConnection(sqlsb.ConnectionString))
                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.Connection = sqlConn;
                    cmd.CommandTimeout = 60;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = @"DELETE  [TKMOC].[dbo].[MOCMANULINEBATCHMODIFYS_SANDWISH]
                                        WHERE MB001=@MB001 AND MODIFY_MB001=@MODIFY_MB001";

                    cmd.Parameters.AddWithValue("@MB001", MB001 ?? string.Empty);                    
                    cmd.Parameters.AddWithValue("@MODIFY_MB001", MODIFY_MB001 ?? string.Empty);         

                    if (tran != null)
                        cmd.Transaction = tran;

                    sqlConn.Open();
                    result = cmd.ExecuteNonQuery();

                    if (result == 0 && tran != null)
                    {
                        tran.Rollback();
                    }
                    else if (tran != null)
                    {
                        tran.Commit();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"發生錯誤：{ex.Message}");
                if (tran != null)
                    tran.Rollback();
            }
        }


        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCH();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string mb001 = textBox1.Text;
            string mb002= textBox2.Text;
            string MODIFY_mb001 = textBox3.Text;
            string MODIFY_mb002 = textBox4.Text;

            if (!string.IsNullOrEmpty(mb001) && !string.IsNullOrEmpty(mb002) && !string.IsNullOrEmpty(MODIFY_mb001) && !string.IsNullOrEmpty(MODIFY_mb002) )
            {
                ADD_MOCMANULINEBATCHMODIFYS_SANDWISH(mb001, mb002, MODIFY_mb001, MODIFY_mb002);
                SEARCH();

                MessageBox.Show("完成");

            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string mb001 = textBox5.Text;           
            string MODIFY_mb001 = textBox6.Text;    
            
            DialogResult dialogResult = MessageBox.Show("確定要刪除嗎?", "刪除確認", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                if (!string.IsNullOrEmpty(mb001) && !string.IsNullOrEmpty(MODIFY_mb001))
                {
                    DEETE_MOCMANULINEBATCHMODIFYS_SANDWISH(mb001, MODIFY_mb001);
                    SEARCH();

                    MessageBox.Show("完成");

                }
            }

               
        }

        #endregion

       
    }
}
