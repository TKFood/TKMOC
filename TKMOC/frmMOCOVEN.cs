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

namespace TKMOC
{
    public partial class frmMOCOVEN : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataTable dt = new DataTable();
        DataGridViewRow drEMPLOYEE = new DataGridViewRow();
        string tablename = null;
        string ID;
        int result;
        int rownum = 0;
        Thread TD;

        public frmMOCOVEN()
        {
            InitializeComponent();
            combobox1load();
            combobox2load();
            combobox3load();
            combobox4load();
            combobox5load();
        }

        #region FUNCTION
        public void combobox1load()
        {

            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            String Sequel = "SELECT [ID],[DEPNAME]  FROM [TKMOC].[dbo].[MANUDEP]";
            SqlDataAdapter da = new SqlDataAdapter(Sequel, sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("DEPNAME", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "ID";
            comboBox1.DisplayMember = "DEPNAME";
            sqlConn.Close();

        }
        public void combobox2load()
        {

            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            String Sequel = "SELECT  [ID],[NAME] FROM [TKMOC].[dbo].[MANUEMPLOYEE] ORDER BY ID";
            SqlDataAdapter da = new SqlDataAdapter(Sequel, sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "ID";
            comboBox2.DisplayMember = "NAME";
            sqlConn.Close();

        }

        public void combobox3load()
        {

            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            String Sequel = "SELECT  [ID],[NAME] FROM [TKMOC].[dbo].[MANUEMPLOYEE] ORDER BY ID";
            SqlDataAdapter da = new SqlDataAdapter(Sequel, sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox3.DataSource = dt.DefaultView;
            comboBox3.ValueMember = "ID";
            comboBox3.DisplayMember = "NAME";
            sqlConn.Close();

        }

        public void combobox4load()
        {

            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            String Sequel = "SELECT  [ID],[NAME] FROM [TKMOC].[dbo].[MANUEMPLOYEE] ORDER BY ID";
            SqlDataAdapter da = new SqlDataAdapter(Sequel, sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox4.DataSource = dt.DefaultView;
            comboBox4.ValueMember = "ID";
            comboBox4.DisplayMember = "NAME";
            sqlConn.Close();

        }

        public void combobox5load()
        {

            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            String Sequel = "SELECT  [ID],[NAME] FROM [TKMOC].[dbo].[MANUEMPLOYEE] ORDER BY ID";
            SqlDataAdapter da = new SqlDataAdapter(Sequel, sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox5.DataSource = dt.DefaultView;
            comboBox5.ValueMember = "ID";
            comboBox5.DisplayMember = "NAME";
            sqlConn.Close();

        }

        public void Search()
        {
            StringBuilder Query = new StringBuilder();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();
                sbSql.AppendFormat(@" SELECT  [ID],[OVENDATE],[MANUDEP],[PREHEARTSTART],[PREHEARTEND],[GAS],[FLODPEOPLE1],[FLODPEOPLE2],[MANAGER],[OPERATOR]");
                sbSql.AppendFormat(@" FROM [TKMOC].[dbo].[MOCOVEN] WITH(NOLOCK)");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {
                    //label1.Text = "找不到資料";
                }
                else
                {
                    if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                    {

                       
                        //dataGridView1.Rows.Clear();
                        dataGridView1.DataSource = ds.Tables["TEMPds1"];
                        dataGridView1.AutoResizeColumns();
                        textBoxID.Text = ds.Tables["TEMPds1"].Rows[0]["ID"].ToString();

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

        public void SETADD()
        {
            dateTimePicker1.Enabled = true;
            dateTimePicker2.Enabled = true;
            dateTimePicker3.Enabled = true;
            textBox1.ReadOnly = false;
            comboBox1.Enabled = true;
            comboBox2.Enabled = true;
            comboBox3.Enabled = true;
            comboBox4.Enabled = true;
            comboBox5.Enabled = true;
        }

        public void SETUPDATE()
        {
            dateTimePicker1.Enabled = true;
            dateTimePicker2.Enabled = true;
            dateTimePicker3.Enabled = true;
            textBox1.ReadOnly = false;
            comboBox1.Enabled = true;
            comboBox2.Enabled = true;
            comboBox3.Enabled = true;
            comboBox4.Enabled = true;
            comboBox5.Enabled = true;
        }

        public void SETFINISH()
        {
            dateTimePicker1.Enabled = false;
            dateTimePicker2.Enabled = false;
            dateTimePicker3.Enabled = false;
            textBox1.ReadOnly = true;
            comboBox1.Enabled = false;
            comboBox2.Enabled = false;
            comboBox3.Enabled = false;
            comboBox4.Enabled = false;
            comboBox5.Enabled = false;
        }

        public void UPDATE()
        {
            try
            {
                //add ZWAREWHOUSEPURTH
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.Append(" UPDATE [TKMOC].[dbo].[MOCOVEN] ");
                sbSql.AppendFormat("  SET [OVENDATE]='{1}',[MANUDEP]='{2}',[PREHEARTSTART]='{3}',[PREHEARTEND]='{4}',[GAS]='{5}',[FLODPEOPLE1]='{6}',[FLODPEOPLE2]='{7}',[MANAGER]='{8}',[OPERATOR]='{9}'  WHERE [ID]='{0}' ", textBoxID.Text.ToString(), dateTimePicker1.Value.ToString("yyyy/MM/dd"), comboBox1.SelectedValue.ToString(), dateTimePicker2.Value.ToString("HH:mm"), dateTimePicker3.Value.ToString("HH:mm"), textBox1.Text.ToString(), comboBox2.SelectedValue.ToString(), comboBox3.SelectedValue.ToString(), comboBox4.SelectedValue.ToString(), comboBox5.SelectedValue.ToString());
                sbSql.Append("   ");

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

        public void ADD()
        {
            try
            {
                //add ZWAREWHOUSEPURTH
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.Append(" INSERT INTO [TKMOC].[dbo].[MOCOVEN]  ");
                sbSql.Append(" ([ID],[OVENDATE],[MANUDEP],[PREHEARTSTART],[PREHEARTEND],[GAS],[FLODPEOPLE1],[FLODPEOPLE2],[MANAGER],[OPERATOR])  ");
                sbSql.AppendFormat("  VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}') ", Guid.NewGuid(), dateTimePicker1.Value.ToString("yyyy/MM/dd"),comboBox1.SelectedValue.ToString(), dateTimePicker2.Value.ToString("HH:mm"), dateTimePicker3.Value.ToString("HH:mm"),textBox1.Text.ToString(),comboBox2.SelectedValue.ToString(), comboBox3.SelectedValue.ToString(), comboBox4.SelectedValue.ToString(), comboBox5.SelectedValue.ToString());

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
        #endregion

        #region BUTTOON
        private void button1_Click(object sender, EventArgs e)
        {
            SETADD();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SETUPDATE();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBoxID.Text.ToString()))
            {
                UPDATE();
            }
            else
            {
                ADD();
            }
            Search();
            SETFINISH();
        }
        private void button4_Click(object sender, EventArgs e)
        {
            Search();
        }
        #endregion


    }
}
