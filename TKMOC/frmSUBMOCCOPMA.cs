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

        public frmSUBMOCCOPMA()
        {
            InitializeComponent();
            this.textBox1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBox1_KeyDown);
        }

        #region FUNCTION
        public void SEARCHMOCCOPMA()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                if (string.IsNullOrEmpty(textBox1.Text))
                {
                    sbSql.AppendFormat(@"SELECT [ID] AS '代號',[NAME] AS '名稱' FROM [TKMOC].[dbo].[MOCCOPMA]  ");
                }
                else
                {
                    sbSql.AppendFormat(@"SELECT [ID] AS '代號',[NAME] AS '名稱' FROM [TKMOC].[dbo].[MOCCOPMA] WHERE   [NAME]  LIKE '{0}%'", textBox1.Text);
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
                SEARCHMOCCOPMA();
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
                    MA001 = row.Cells["MA001"].Value.ToString();

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
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCHMOCCOPMA();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        #endregion


    }
}
