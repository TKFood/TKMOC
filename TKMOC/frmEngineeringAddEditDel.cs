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
    public partial class frmEngineeringAddEditDel : Form
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
        string EDITEquipmentID;
        int result;
        Thread TD;

        public frmEngineeringAddEditDel()
        {
            InitializeComponent();
            combobox1load();
        }
        public frmEngineeringAddEditDel(string EquipmentID)
        {
            InitializeComponent();
            combobox1load();
            textBox1.Text = EquipmentID;
            textBox1.ReadOnly = false;

            if (!string.IsNullOrEmpty(EquipmentID))
            {
                EDITEquipmentID = EquipmentID;
                Search(EquipmentID);
                textBox1.ReadOnly = true;
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
        public void Search(string ID)
        {
            StringBuilder Query = new StringBuilder();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();               

                sbSql.Append(@" SELECT [ID] AS '設備編號',[NAME]  AS '設備名稱',[UNIT]  AS '單位',[FACTORY]  AS '廠牌',[TYPE]  AS '型別',[MAINTENANCE]  AS '保養',[CHEKCK]  AS '點檢',[STATUS]  AS '狀況說明'  ");
                sbSql.Append(@" FROM [TKMOC].[dbo].[ENGEQUIPMENT] WITH (NOLOCK)");
                sbSql.AppendFormat(@" WHERE [ID] ='{0}'",ID);
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
                        textBox1.Text = ds.Tables["TEMPds1"].Rows[0]["設備編號"].ToString();
                        textBox2.Text = ds.Tables["TEMPds1"].Rows[0]["設備名稱"].ToString();
                        textBox3.Text = ds.Tables["TEMPds1"].Rows[0]["廠牌"].ToString();
                        textBox4.Text = ds.Tables["TEMPds1"].Rows[0]["型別"].ToString();
                        textBox5.Text = ds.Tables["TEMPds1"].Rows[0]["保養"].ToString();
                        textBox6.Text = ds.Tables["TEMPds1"].Rows[0]["點檢"].ToString();
                        textBox7.Text = ds.Tables["TEMPds1"].Rows[0]["狀況說明"].ToString();
                        comboBox1.SelectedValue = ds.Tables["TEMPds1"].Rows[0]["單位"].ToString();

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
                sbSql.Append(" UPDATE [TKMOC].[dbo].[ENGEQUIPMENT]");
                sbSql.AppendFormat(" SET [NAME]='{1}',[UNIT]='{2}',[FACTORY]='{3}',[TYPE]='{4}',[MAINTENANCE]='{5}',[CHEKCK]='{6}',[STATUS]='{7}' WHERE [ID]='{0}' ",textBox1.Text.ToString(), textBox2.Text.ToString(),comboBox1.SelectedValue.ToString(), textBox3.Text.ToString(), textBox4.Text.ToString(), textBox5.Text.ToString(), textBox6.Text.ToString(), textBox7.Text.ToString());
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
                sbSql.Append(" INSERT INTO [TKMOC].[dbo].[ENGEQUIPMENT]  ");
                sbSql.Append(" ([ID],[NAME],[UNIT],[FACTORY],[TYPE],[MAINTENANCE],[CHEKCK],[STATUS])  ");
                sbSql.AppendFormat("  VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}') ", textBox1.Text.ToString(), textBox2.Text.ToString(), comboBox1.SelectedValue.ToString(), textBox3.Text.ToString(), textBox4.Text.ToString(), textBox5.Text.ToString(), textBox6.Text.ToString(), textBox7.Text.ToString());

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
            if (!string.IsNullOrEmpty(EDITEquipmentID))
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
