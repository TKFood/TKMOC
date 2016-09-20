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
    public partial class frmMACHINEMAINWEEK : Form
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

        public frmMACHINEMAINWEEK()
        {
            InitializeComponent();
            combobox1load();
            combobox2load();
        }
        public frmMACHINEMAINWEEK(string ID)
        {
            InitializeComponent();
            combobox1load();
            combobox2load();
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
            Sequel.AppendFormat(@"SELECT  [ID],[NAME] FROM [TKMOC].[dbo].[ENGEQUIPMENT] WHERE [UNIT]='{0}'", comboBox1.SelectedValue.ToString());
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

            sbSql.AppendFormat(@" SELECT  [ID] ,[NAME] FROM [TKMOC].[dbo].[ENGEQUIPMENT] WITH (NOLOCK) WHERE[ID]='{0}' ", comboBox2.SelectedValue.ToString());
            sbSql.Append(@"  ");


            adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

            sqlCmdBuilder = new SqlCommandBuilder(adapter);
            sqlConn.Open();
            ds.Clear();
            adapter.Fill(ds, "TEMPds1");
            sqlConn.Close();


            if (ds.Tables["TEMPds1"].Rows.Count == 0)
            {
               
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

                sbSql.AppendFormat(@" SELECT [ID] AS '編號',[EQUIPMENTID] AS '設備編號',[EQUIPMENTNAME]  AS '設備名稱',[UNIT] AS '使用單位',[MAINDATE]  AS '保養日',[MAINYEAR] AS '保養年',[MAINMONTH] AS '保養月',[MAINWEEK] AS '保養週次',[ISMAIN] AS '是否保養'   ");
                sbSql.Append(@" FROM [TKMOC].[dbo].[MACHINEMAINWEEK] WITH (NOLOCK)");
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
                    
                }
                else
                {
                    if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        textBox1.Text = ds.Tables["TEMPds1"].Rows[0]["設備名稱"].ToString();
                        textBox2.Text = ds.Tables["TEMPds1"].Rows[0]["保養週次"].ToString();
                        textID.Text = ds.Tables["TEMPds1"].Rows[0]["編號"].ToString();
                        comboBox1.SelectedValue = ds.Tables["TEMPds1"].Rows[0]["使用單位"].ToString();
                        comboBox2.SelectedValue = ds.Tables["TEMPds1"].Rows[0]["設備編號"].ToString();
                        comboBox3.SelectedValue = ds.Tables["TEMPds1"].Rows[0]["是否保養"].ToString();
                        comboBox3.Text = ds.Tables["TEMPds1"].Rows[0]["是否保養"].ToString();

                        dateTimePicker1.Value = Convert.ToDateTime(ds.Tables["TEMPds1"].Rows[0]["保養日"].ToString());
                        dateTimePicker2.Value = Convert.ToDateTime(ds.Tables["TEMPds1"].Rows[0]["保養年"].ToString());
                        dateTimePicker3.Value = Convert.ToDateTime(ds.Tables["TEMPds1"].Rows[0]["保養月"].ToString());


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
                sbSql.Append(" UPDATE [TKMOC].[dbo].[MACHINEMAINWEEK]");
                sbSql.AppendFormat("  SET [EQUIPMENTID]='{1}',[EQUIPMENTNAME]='{2}',[UNIT]='{3}',[MAINDATE]='{4}',[MAINYEAR]='{5}',[MAINMONTH]='{6}',[MAINWEEK]='{7}',[ISMAIN]='{8}' WHERE [ID]='{0}'", textID.Text.ToString(), comboBox2.SelectedValue.ToString(), textBox1.Text.ToString(),comboBox1.SelectedValue.ToString(), dateTimePicker1.Value.ToString("yyyy/MM/dd"), dateTimePicker2.Value.ToString("yyyy"), dateTimePicker3.Value.ToString("MM"),textBox2.Text.ToString(),comboBox3.Text.ToString());
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
                sbSql.Append(" INSERT INTO [TKMOC].[dbo].[MACHINEMAINWEEK] ");
                sbSql.Append(" ([ID],[EQUIPMENTID],[EQUIPMENTNAME],[UNIT],[MAINDATE],[MAINYEAR],[MAINMONTH],[MAINWEEK],[ISMAIN])  ");
                sbSql.AppendFormat("  VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}') ", Guid.NewGuid(),  comboBox2.SelectedValue.ToString(), textBox1.Text.ToString(), comboBox1.SelectedValue.ToString(), dateTimePicker1.Value.ToString("yyyy/MM/dd"), dateTimePicker2.Value.ToString("yyyy"), dateTimePicker3.Value.ToString("MM"), textBox2.Text.ToString(), comboBox3.Text.ToString());

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

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            dateTimePicker2.Value = dateTimePicker1.Value;
            dateTimePicker3.Value = dateTimePicker1.Value;
            textBox2.Text = getWeekOfYear(dateTimePicker1.Value).ToString();
        }
        public int getWeekOfYear(DateTime dt)
        {

            System.Globalization.GregorianCalendar getWeek = new  System.Globalization.GregorianCalendar();
            return getWeek.GetWeekOfYear(dt, System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday);//設定星期日為一週的第一天

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
