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
using TKITDLL;

namespace TKMOC
{
    public partial class frmSUBMOCCOPMA : Form
    {
        private ComponentResourceManager _ResourceManager = new ComponentResourceManager();
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();


        DataSet dsBOMMC = new DataSet();
        DataSet dsBOMMD = new DataSet();

        DataTable dt = new DataTable();
        string tablename = null;
        int result;


        string MA001;
        int rowindex;

        public frmSUBMOCCOPMA()
        {
            InitializeComponent();
            this.textBox1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBox1_KeyDown);
        }

        #region FUNCTION
        public void SEARCHMOCCOPMA(string MA002)
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                if (string.IsNullOrEmpty(textBox1.Text))
                {
                    sbSql.AppendFormat(@"
                                        SELECT
                                        MA001 AS '代號'
                                        ,MA002   AS '名稱'
                                        FROM [TK].dbo.COPMA
                                        WHERE 1=1
                                        AND MA002 NOT LIKE '%停用%'
                                        AND (MA001 LIKE '2%')
                                        ORDER BY MA002 
                                        ");
                }
                else
                {
                    sbSql.AppendFormat(@"
                                        SELECT
                                        MA001 AS '代號'
                                        ,MA002   AS '名稱'
                                        FROM [TK].dbo.COPMA
                                        WHERE 1=1
                                        AND MA002 NOT LIKE '%停用%'
                                        AND (MA001 LIKE '2%')
                                        AND MA002 LIKE '{0}%'
                                        ORDER BY MA002 
                                       ", MA002);
                }

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

                        dataGridView1.DataSource = ds1.Tables["TEMPds1"];
                        dataGridView1.AutoResizeColumns();
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

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (Keys.Enter == e.KeyCode)
            {
                SEARCHMOCCOPMA(textBox1.Text.ToString().Trim());
            }
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    MA001 = row.Cells["名稱"].Value.ToString();

                }
                else
                {
                    MA001 = null;

                }
            }
        }

        public string TextBoxMsg
        {
            set
            {

            }
            get
            {
                return MA001;
            }
        }

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {    
            if (Keys.Enter == e.KeyCode)
            {
                e.Handled = true;
                                
                this.Close();
            }
        }
        private void dataGridView1_KeyPress(object sender, KeyPressEventArgs e)
        {

        }

        public void SEARCHMOCCOPMA2(string MA002)
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                if (string.IsNullOrEmpty(textBox1.Text))
                {
                    sbSql.AppendFormat(@"  
                                        SELECT
                                        MA001 AS '代號'
                                        ,MA002   AS '名稱'
                                        FROM [TK].dbo.COPMA
                                        WHERE 1=1
                                        AND MA002 NOT LIKE '%停用%'
                                        AND (MA001 LIKE '3%')                                        
                                        ORDER BY MA002 
                                        ");
                }
                else
                {
                    sbSql.AppendFormat(@"  
                                        SELECT
                                        MA001 AS '代號'
                                        ,MA002   AS '名稱'
                                        FROM [TK].dbo.COPMA
                                        WHERE 1=1
                                        AND MA002 NOT LIKE '%停用%'
                                        AND (MA001 LIKE '3%')
                                        AND MA002 LIKE '{0}%'
                                        ORDER BY MA002 "
                                        , MA002);
                }

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

                        dataGridView1.DataSource = ds1.Tables["TEMPds1"];
                        dataGridView1.AutoResizeColumns();
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
            SEARCHMOCCOPMA(textBox1.Text.ToString().Trim());
        }
        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SEARCHMOCCOPMA2(textBox1.Text.ToString().Trim());
        }



        #endregion

    }
}
