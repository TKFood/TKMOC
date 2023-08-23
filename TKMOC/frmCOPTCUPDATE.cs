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
    public partial class frmCOPTCUPDATE : Form
    {
        public frmCOPTCUPDATE()
        {
            InitializeComponent();

            SET_DEFAULT();
            InitializeDataGridView();

        }

        #region FUNCTION
        public void SET_DEFAULT()
        {
            textBox1.Text = DateTime.Now.ToString("yyyyMMdd");
        }

        public void InitializeDataGridView()
        {
            // Create the DataGridViewCheckBoxColumn
            DataGridViewCheckBoxColumn checkBoxColumn = new DataGridViewCheckBoxColumn();
            checkBoxColumn.HeaderText = "勾選";
            checkBoxColumn.Name = "CheckBoxColumn";
            // Add the CheckBoxColumn to the DataGridView
            dataGridView1.Columns.Insert(0, checkBoxColumn);

            
        }
               

        public void SEARCH_COPTC(string TC002)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
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
                var sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();
                
                sbSql.AppendFormat(@"  
                                    SELECT TC027 AS '確認碼',TC001  AS '訂單單別',TC002  AS '訂單單號', TC053 AS '客戶'
                                    FROM [TK].dbo.COPTC
                                    WHERE TC002 LIKE '{0}%'
                                    ", TC002);

                var adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);
                var sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
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
                    {
                        //dataGridView1.Rows.Clear();
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

            }
        }

        public void GW1_CHECKBOX()
        {
            List<string> selectedIDs = new List<string>();

            // Loop through the DataGridView and check if CheckBox is selected
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                // Skip new rows and header row
                if (!row.IsNewRow)
                {
                    DataGridViewCheckBoxCell checkBoxCell = row.Cells["CheckBoxColumn"] as DataGridViewCheckBoxCell;
                    if (checkBoxCell != null && checkBoxCell.Value != null && (bool)checkBoxCell.Value)
                    {
                        string id = row.Cells["訂單單別"].Value.ToString().Trim()+ row.Cells["訂單單號"].Value.ToString().Trim();
                        selectedIDs.Add(id);
                    }
                }
            }

            // Display selected IDs
            //MessageBox.Show("Selected IDs: " + string.Join(", ", selectedIDs));

            // Wrap each selected ID with single quotes
            List<string> wrappedIDs = selectedIDs.Select(id => "'" + id + "'").ToList();
            // Combine wrapped IDs into a single string
            string SLEECTED = string.Join(", ", wrappedIDs);
            // MessageBox.Show(SLEECTED);

            COPTC_UPDATE_TC027(comboBox1.Text.ToString(), SLEECTED);
        }

        public void COPTC_UPDATE_TC027(string TC027, string TC002)
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);
   
            string connectionString = sqlsb.ConnectionString;

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                // Start a new transaction
                SqlTransaction transaction = connection.BeginTransaction();
                try
                {
                    // Execute SQL statements within the transaction
                    using (SqlCommand command = connection.CreateCommand())
                    {
                        command.Transaction = transaction;

                        StringBuilder SQLEXECTE = new StringBuilder();
                        SQLEXECTE.AppendFormat(@"
                                                UPDATE [TK].dbo.COPTC
                                                SET TC027='{0}'
                                                WHERE TC001+TC002 IN ({1})
                                                ", TC027,TC002);

                        command.CommandText = SQLEXECTE.ToString();
                        command.ExecuteNonQuery();
                    }

                    // Commit the transaction if everything succeeded
                    transaction.Commit();                   
                }
                catch (Exception ex)
                {
                    try
                    {
                        transaction.Rollback();                       
                    }
                    catch (Exception rollbackEx)
                    {                        
                    }
                }
            }
        }

        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            SEARCH_COPTC(textBox1.Text.Trim());
        }
        private void button2_Click(object sender, EventArgs e)
        {            
            DialogResult dialogResult = MessageBox.Show("確認?", "確認?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                GW1_CHECKBOX();
                SEARCH_COPTC(textBox1.Text.Trim());
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }


        #endregion


    }
}
