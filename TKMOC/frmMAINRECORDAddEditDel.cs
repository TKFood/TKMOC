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
    public partial class frmMAINRECORDAddEditDel : Form
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
        string tablename = null;
        string EDITID;
        int result;
        Thread TD;

        public frmMAINRECORDAddEditDel()
        {
            InitializeComponent();
            combobox1load();
            combobox2load();
            combobox3load();
        }
        public frmMAINRECORDAddEditDel(string ID)
        {
            InitializeComponent();
            combobox1load();
            combobox2load();
            combobox3load();
            if (!string.IsNullOrEmpty(ID))
            {
                EDITID = ID;
                comboBox1.Enabled = false;
                Search(ID);
            }
        }
        #region FUNCTION

        public void combobox1load()
        {

            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            String Sequel = "SELECT  [UNITID],[UNITNAME] FROM [TKMOC].[dbo].[ENDUNIT] WHERE [UNITID]<>'0' ";
            SqlDataAdapter da = new SqlDataAdapter(Sequel, sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("UNITID", typeof(string));
            dt.Columns.Add("UNITNAME", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "UNITID";
            comboBox1.DisplayMember = "UNITNAME";
            sqlConn.Close();

        }
        private void comboBox1_SelectionChangeCommitted(object sender, EventArgs e)
        {
            combobox2load();
        }
        public void combobox2load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT  [ID],[NAME] FROM [TKMOC].[dbo].[ENGEQUIPMENT] ", comboBox1.SelectedValue.ToString());
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "ID";
            comboBox2.DisplayMember = "ID";
            sqlConn.Close();
            SETTEXT();

        }
        public void combobox3load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.Append(@"SELECT [ID],[MALFUNCION]  FROM [TKMOC].[dbo].[MALFUNCION]");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("MALFUNCION", typeof(string));
            da.Fill(dt);
            comboBox3.DataSource = dt.DefaultView;
            comboBox3.ValueMember = "ID";
            comboBox3.DisplayMember = "MALFUNCION";
            sqlConn.Close();
            SETTEXT();

        }
        private void comboBox2_SelectionChangeCommitted(object sender, EventArgs e)
        {
            SETTEXT();
        }
        public void SETTEXT()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);

            sbSql.Clear();
            sbSqlQuery.Clear();

            sbSql.AppendFormat(@" SELECT  [ID],[NAME] FROM [TKMOC].[dbo].[ENGEQUIPMENT] WITH (NOLOCK) WHERE[ID]='{0}' ", comboBox2.SelectedValue.ToString());
            sbSql.Append(@"  ");


            adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

            sqlCmdBuilder = new SqlCommandBuilder(adapter);
            sqlConn.Open();
            ds.Clear();
            adapter.Fill(ds, "TEMPds1");
            sqlConn.Close();


            if (ds.Tables["TEMPds1"].Rows.Count == 0)
            {
                label1.Text = "找不到資料";
            }
            else
            {
                if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    textBox1.Text = ds.Tables["TEMPds1"].Rows[0]["NAME"].ToString();
                }
            }
        }

        public void Search(string ID)
        {
            StringBuilder Query = new StringBuilder();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@" SELECT  [ID] AS '編號',[EQUIPMENTID] AS '財產編號',[EQUIPMENTNAME] AS '設備名稱',[UNIT] AS '使用部門',[ERROR] AS '故障情形',[MAINDATEBEGIN] AS '維修時間起',[MAINDATEEND] AS '維修時間迄',[MAINDATHR] AS '維修時數',[MAINEMP] AS '維修人員',[MALFUNCIONID] AS '故障性質',[MAINSTATUS] AS '維修內容',[MAINUSED] AS '本次更換' ");
                sbSql.Append(@" FROM [TKMOC].[dbo].[MAINRECORD]  WITH (NOLOCK)");
                sbSql.AppendFormat(@" WHERE [ID] ='{0}'", ID);
                sbSql.Append(@" ORDER BY [ID]  ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {
                    label1.Text = "找不到資料";
                }
                else
                {
                    if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        textBox1.Text = ds.Tables["TEMPds1"].Rows[0]["設備名稱"].ToString();
                        textBox2.Text = ds.Tables["TEMPds1"].Rows[0]["故障情形"].ToString();
                        textBox3.Text = ds.Tables["TEMPds1"].Rows[0]["維修時數"].ToString();
                        textBox4.Text = ds.Tables["TEMPds1"].Rows[0]["維修人員"].ToString();
                        textBox5.Text = ds.Tables["TEMPds1"].Rows[0]["維修內容"].ToString();
                        textBox6.Text = ds.Tables["TEMPds1"].Rows[0]["本次更換"].ToString();
                        textBox8.Text = ds.Tables["TEMPds1"].Rows[0]["編號"].ToString();
                        
                        comboBox1.SelectedValue = ds.Tables["TEMPds1"].Rows[0]["使用部門"].ToString();                        
                        comboBox2.SelectedValue = ds.Tables["TEMPds1"].Rows[0]["財產編號"].ToString();
                        comboBox2.Text = ds.Tables["TEMPds1"].Rows[0]["財產編號"].ToString();
                        comboBox3.SelectedValue = ds.Tables["TEMPds1"].Rows[0]["故障性質"].ToString();
                        dateTimePicker1.Value = Convert.ToDateTime(ds.Tables["TEMPds1"].Rows[0]["維修時間起"].ToString());
                        dateTimePicker2.Value = Convert.ToDateTime(ds.Tables["TEMPds1"].Rows[0]["維修時間迄"].ToString());
                        

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
                sbSql.Append(" UPDATE [TKMOC].[dbo].[MAINRECORD] ");
                sbSql.AppendFormat("  SET [EQUIPMENTID]='{1}',[EQUIPMENTNAME]='{2}',[UNIT]='{3}',[ERROR]='{4}',[MAINDATEBEGIN]='{5}',[MAINDATEEND]='{6}',[MAINDATHR]='{7}',[MAINEMP]='{8}',[MALFUNCIONID]='{9}',[MAINSTATUS]='{10}',[MAINUSED]='{11}' WHERE [ID]='{0}' ", textBox8.Text.ToString(),comboBox2.SelectedValue.ToString(),textBox1.Text.ToString(),comboBox1.SelectedValue.ToString(),textBox2.Text.ToString(),dateTimePicker1.Value.ToString("yyyy/MM/dd HH:mm"), dateTimePicker2.Value.ToString("yyyy/MM/dd HH:mm"),textBox3.Text.ToString(),textBox4.Text.ToString(),comboBox3.SelectedValue.ToString(),textBox5.Text.ToString(),textBox6.Text.ToString());
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
                    this.Close();

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
                sbSql.Append(" INSERT INTO [TKMOC].[dbo].[MAINRECORD] ");
                sbSql.Append(" ([ID],[EQUIPMENTID],[EQUIPMENTNAME],[UNIT],[ERROR],[MAINDATEBEGIN],[MAINDATEEND],[MAINDATHR],[MAINEMP],[MALFUNCIONID],[MAINSTATUS],[MAINUSED])  ");
                sbSql.AppendFormat("  VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}') ", Guid.NewGuid(), comboBox2.SelectedValue.ToString(), textBox1.Text.ToString(), comboBox1.SelectedValue.ToString(), textBox2.Text.ToString(), dateTimePicker1.Value.ToString("yyyy/MM/dd HH:mm"), dateTimePicker2.Value.ToString("yyyy/MM/dd HH:mm"), textBox3.Text.ToString(), textBox4.Text.ToString(), comboBox3.SelectedValue.ToString(), textBox5.Text.ToString(), textBox6.Text.ToString());

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
                    this.Close();

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
            if (!string.IsNullOrEmpty(EDITID))
            {
                UPDATE();
            }
            else
            {
                ADD();
            }
        }


        #endregion

       
    }
}
