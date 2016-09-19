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
    public partial class frmMAINPARTS : Form
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

        public frmMAINPARTS()
        {
            InitializeComponent();
            combobox2load();
        }

        public frmMAINPARTS(string ID)
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

                sbSql.AppendFormat(@" SELECT [ID],[EQUIPMENTID] AS '設備編號',[EQUIPMENTNAME] AS '設備名稱',[PARTSNO] AS '備品編號',[PARTSNAME] AS '品名',[PARTSSPEC] AS '規格',CAST([PARTSPRICE] AS  DECIMAL(16,2) )AS '單價',[PARTSFACTORY] AS '供應商',[TEL] AS '電話',[YEARS] AS '使用壽命',[STOCKNUM] AS '安全庫存',[PRETIME] AS '前置時間',[REMARK] AS '備註'");
                sbSql.Append(@" FROM [TKMOC].[dbo].[MAINPARTS] WITH (NOLOCK)");
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
                        textBox9.Text = ds.Tables["TEMPds1"].Rows[0]["前置時間"].ToString();
                        textBox10.Text = ds.Tables["TEMPds1"].Rows[0]["備註"].ToString();
                        textID.Text = ds.Tables["TEMPds1"].Rows[0]["ID"].ToString();
                        comboBox2.SelectedValue = ds.Tables["TEMPds1"].Rows[0]["設備編號"].ToString();
                        numericUpDown1.Value = Convert.ToInt32(ds.Tables["TEMPds1"].Rows[0]["安全庫存"].ToString());


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
                sbSql.Append(" UPDATE [TKMOC].[dbo].[MAINPARTS]");
                sbSql.AppendFormat("  SET [EQUIPMENTID]='{1}',[EQUIPMENTNAME]='{2}',[PARTSNO]='{3}',[PARTSNAME]='{4}',[PARTSSPEC]='{5}',[PARTSPRICE]='{6}',[PARTSFACTORY]='{7}',[TEL]='{8}',[YEARS]='{9}',[STOCKNUM]='{10}',[PRETIME]='{11}',[REMARK]='{12}' WHERE [ID]='{0}' ", textID.Text.ToString(), comboBox2.SelectedValue.ToString(), textBox1.Text.ToString(), textBox2.Text.ToString(), textBox3.Text.ToString(), textBox4.Text.ToString(), textBox5.Text.ToString(), textBox6.Text.ToString(), textBox7.Text.ToString(), textBox8.Text.ToString(),numericUpDown1.Value.ToString(), textBox9.Text.ToString(), textBox10.Text.ToString());
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
                sbSql.Append(" INSERT INTO [TKMOC].[dbo].[MAINPARTS]  ");
                sbSql.Append(" ([ID],[EQUIPMENTID],[EQUIPMENTNAME],[PARTSNO],[PARTSNAME],[PARTSSPEC],[PARTSPRICE],[PARTSFACTORY],[TEL],[YEARS],[STOCKNUM],[PRETIME],[REMARK] )  ");
                //sbSql.AppendFormat("  VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}') ", Guid.NewGuid(), comboBox2.SelectedValue.ToString(), textBox1.Text.ToString(), textBox2.Text.ToString(), textBox3.Text.ToString(), textBox4.Text.ToString(), textBox5.Text.ToString(), textBox6.Text.ToString(), textBox7.Text.ToString(), textBox8.Text.ToString(), numericUpDown1.Value.ToString(), textBox9.Text.ToString(), textBox10.Text.ToString());
                sbSql.AppendFormat("  VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}') ", Guid.NewGuid(), comboBox2.SelectedValue.ToString(), textBox1.Text.ToString(), textBox2.Text.ToString(), textBox3.Text.ToString(), textBox4.Text.ToString(), textBox5.Text.ToString(), textBox6.Text.ToString(), textBox7.Text.ToString(), textBox8.Text.ToString(), numericUpDown1.Value.ToString(), textBox9.Text.ToString(), textBox10.Text.ToString());

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
