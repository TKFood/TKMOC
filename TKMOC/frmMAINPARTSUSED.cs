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
    public partial class frmMAINPARTSUSED : Form
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
        int NOWNUM;
        Thread TD;

        public frmMAINPARTSUSED()
        {
            InitializeComponent();
            combobox2load();
        }
        public frmMAINPARTSUSED(string ID)
        {
            InitializeComponent();
            combobox2load();
            if (!string.IsNullOrEmpty(ID))
            {
                EDITID = ID;
                Search(ID);
            }
        }

        #region FUNCTION

        public void combobox2load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT  [ID],[NAME] FROM [TKMOC].[dbo].[ENGEQUIPMENT] ");
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

                sbSql.AppendFormat(@" SELECT  [ID],[EQUIPMENTID] AS '設備編號',[EQUIPMENTNAME]  AS '設備名稱',[PARTSNO]  AS '備品編號',[PARTSNAME] AS '品名',[PARTSSPEC]  AS '規格',CAST([PARTSPRICE] AS  DECIMAL(16,2) )AS '單價',[PARTSFACTORY]  AS '供應商',[TEL] AS '電話',[YEARS] AS '使用壽命',[STOCKNUM] AS '安全庫存',[NOWNUM] AS '現有庫存',[USEDDATE] AS '入庫/領用日',[INUM] AS '入庫數'  ,[USEDNUM] AS '領用數' ");
                sbSql.Append(@" FROM [TKMOC].[dbo].[MAINPARTSUSED] WITH (NOLOCK)");
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
                        textBox2.Text = ds.Tables["TEMPds1"].Rows[0]["備品編號"].ToString();
                        textBox3.Text = ds.Tables["TEMPds1"].Rows[0]["品名"].ToString();
                        textBox4.Text = ds.Tables["TEMPds1"].Rows[0]["規格"].ToString();
                        textBox5.Text = ds.Tables["TEMPds1"].Rows[0]["單價"].ToString();
                        textBox6.Text = ds.Tables["TEMPds1"].Rows[0]["供應商"].ToString();
                        textBox7.Text = ds.Tables["TEMPds1"].Rows[0]["電話"].ToString();
                        textBox8.Text = ds.Tables["TEMPds1"].Rows[0]["使用壽命"].ToString();
                       
                        textID.Text = ds.Tables["TEMPds1"].Rows[0]["ID"].ToString();
                        comboBox2.SelectedValue = ds.Tables["TEMPds1"].Rows[0]["設備編號"].ToString();
                        numericUpDown1.Value = Convert.ToInt32(ds.Tables["TEMPds1"].Rows[0]["安全庫存"].ToString());
                        numericUpDown2.Value = Convert.ToInt32(ds.Tables["TEMPds1"].Rows[0]["現有庫存"].ToString());
                        numericUpDown3.Value = Convert.ToInt32(ds.Tables["TEMPds1"].Rows[0]["入庫數"].ToString());
                        numericUpDown4.Value = Convert.ToInt32(ds.Tables["TEMPds1"].Rows[0]["領用數"].ToString());
                        dateTimePicker1.Value = Convert.ToDateTime(ds.Tables["TEMPds1"].Rows[0]["入庫/領用日"].ToString());

                        NOWNUM= Convert.ToInt32(ds.Tables["TEMPds1"].Rows[0]["現有庫存"].ToString());
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
                sbSql.Append(" UPDATE [TKMOC].[dbo].[MAINPARTSUSED]");
                sbSql.AppendFormat("  SET [EQUIPMENTID]='{1}',[EQUIPMENTNAME]='{2}',[PARTSNO]='{3}',[PARTSNAME]='{4}',[PARTSSPEC]='{5}',[PARTSPRICE]='{6}',[PARTSFACTORY]='{7}',[TEL]='{8}',[YEARS]='{9}',[STOCKNUM]='{10}',[NOWNUM]='{11}',[USEDDATE]='{12}',[INUM] ='{13}',[USEDNUM] ='{14}' WHERE [ID]='{0}' ", textID.Text.ToString(), comboBox2.SelectedValue.ToString(), textBox1.Text.ToString(), textBox2.Text.ToString(), textBox3.Text.ToString(), textBox4.Text.ToString(), textBox5.Text.ToString(), textBox6.Text.ToString(), textBox7.Text.ToString(), textBox8.Text.ToString(), numericUpDown1.Value.ToString(),numericUpDown2.Value.ToString(), dateTimePicker1.Value.ToString("yyyy/MM/dd"), numericUpDown3.Value.ToString(),numericUpDown4.Value.ToString());
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
                sbSql.Append(" INSERT INTO [TKMOC].[dbo].[MAINPARTSUSED]  ");
                sbSql.Append(" ([ID],[EQUIPMENTID],[EQUIPMENTNAME],[PARTSNO],[PARTSNAME],[PARTSSPEC],[PARTSPRICE],[PARTSFACTORY],[TEL],[YEARS],[STOCKNUM],[NOWNUM],[USEDDATE],[INUM],[USEDNUM] )  ");
                sbSql.AppendFormat("  VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}') ", Guid.NewGuid(), comboBox2.SelectedValue.ToString(), textBox1.Text.ToString(), textBox2.Text.ToString(), textBox3.Text.ToString(), textBox4.Text.ToString(), textBox5.Text.ToString(), textBox6.Text.ToString(), textBox7.Text.ToString(), textBox8.Text.ToString(), numericUpDown1.Value.ToString(), numericUpDown2.Value.ToString(), dateTimePicker1.Value.ToString("yyyy/MM/dd"), numericUpDown3.Value.ToString(), numericUpDown4.Value.ToString());

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

        public void CalNowStock()
        {
            int cal = 0;
            cal = NOWNUM +Convert.ToInt32 (numericUpDown3.Value)- Convert.ToInt32(numericUpDown4.Value);
            numericUpDown2.Value = cal;
        }

        private void numericUpDown3_ValueChanged(object sender, EventArgs e)
        {
            CalNowStock();
        }

        private void numericUpDown4_ValueChanged(object sender, EventArgs e)
        {
            CalNowStock();
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
