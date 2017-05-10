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
    public partial class frmMOCCOPMA : Form
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

        DataTable dt = new DataTable();
        string tablename = null;
        int result;
        string EDITSTATUS;

        public frmMOCCOPMA()
        {
            InitializeComponent();
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


               
                sbSql.AppendFormat(@" SELECT [ID] AS '代號',[NAME] AS '名稱' FROM [TKMOC].[dbo].[MOCCOPMA] ORDER BY [ID] ");

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

        public void SAVEMOCCOPMA()
        {
            if (EDITSTATUS.Equals("ADD"))
            {
                ADDMOCCOPMA();
            }
            else if (EDITSTATUS.Equals("UPDATE"))
            {
                UPDATEMOCCOPMA();
            }
            else
            {
                MessageBox.Show("存檔失敗");
            }
        }

        public void ADDMOCCOPMA()
        {

        }

        public void UPDATEMOCCOPMA()
        {

        }

        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            SEARCHMOCCOPMA();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            EDITSTATUS = "ADD";
            textBox1.Text = null;
            textBox2.Text = null;

            textBox1.ReadOnly = false;
            textBox1.ReadOnly = false;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            EDITSTATUS = "UPDATE";

            textBox1.ReadOnly = false;
            textBox1.ReadOnly = false;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            SAVEMOCCOPMA();

            textBox1.ReadOnly = true;
            textBox1.ReadOnly = true;
        }

        #endregion


    }
}
