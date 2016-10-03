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
    public partial class frmMACHINEDAILYCHECK : Form
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

        public frmMACHINEDAILYCHECK()
        {
            InitializeComponent();
            combobox1load();
            combobox2load();
        }
        public frmMACHINEDAILYCHECK(string ID)
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

                sbSql.AppendFormat(@" SELECT [EQUIPMENTID] AS '設備編號',[EQUIPMENTNAME] AS '設備名稱',[UNIT] AS '使用單位',[MAINDATE] AS '保養日期',[CHECK1] AS '工作前機台內外部的清理消毒',[CHECK2] AS '各部螺絲確實鎖緊',[CHECK3] AS '各操作按鍵鈕正常無異',[CHECK4] AS '機台運行順暢無異常',[CHECK5] AS '各設定確實依作業標準書',[CHECK6] AS '機器運行正常無異聲',[CHECK7] AS '零件使用後確實清潔消毒',[CHECK8] AS '各指示燈確實亮起無異',[CHECK9] AS '各設定溫度時間確實達到',[CHECK10] AS '零件安裝固定完全',[CHECK11] AS '工作後機台內外部清潔消毒',[CHECKOR] ,[ID] ");
                sbSql.Append(@" FROM [TKMOC].[dbo].[MACHINEDAILYCHECK] WITH (NOLOCK)");
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
                        textBox2.Text = ds.Tables["TEMPds1"].Rows[0]["CHECKOR"].ToString();
                        textID.Text= ds.Tables["TEMPds1"].Rows[0]["ID"].ToString();
                        comboBox1.SelectedValue = ds.Tables["TEMPds1"].Rows[0]["使用單位"].ToString();
                        comboBox2.SelectedValue = ds.Tables["TEMPds1"].Rows[0]["設備編號"].ToString();
                        comboBox2.Text = ds.Tables["TEMPds1"].Rows[0]["設備編號"].ToString();
                        comboBox3.Text = ds.Tables["TEMPds1"].Rows[0]["工作前機台內外部的清理消毒"].ToString();
                        comboBox4.Text = ds.Tables["TEMPds1"].Rows[0]["各部螺絲確實鎖緊"].ToString();
                        comboBox5.Text = ds.Tables["TEMPds1"].Rows[0]["各操作按鍵鈕正常無異"].ToString();
                        comboBox6.Text = ds.Tables["TEMPds1"].Rows[0]["機台運行順暢無異常"].ToString();
                        comboBox7.Text = ds.Tables["TEMPds1"].Rows[0]["各設定確實依作業標準書"].ToString();
                        comboBox8.Text = ds.Tables["TEMPds1"].Rows[0]["機器運行正常無異聲"].ToString();
                        comboBox9.Text = ds.Tables["TEMPds1"].Rows[0]["零件使用後確實清潔消毒"].ToString();
                        comboBox10.Text = ds.Tables["TEMPds1"].Rows[0]["各指示燈確實亮起無異"].ToString();
                        comboBox11.Text = ds.Tables["TEMPds1"].Rows[0]["各設定溫度時間確實達到"].ToString();
                        comboBox12.Text = ds.Tables["TEMPds1"].Rows[0]["零件安裝固定完全"].ToString();
                        comboBox13.Text = ds.Tables["TEMPds1"].Rows[0]["工作後機台內外部清潔消毒"].ToString();

                        dateTimePicker1.Value = Convert.ToDateTime(ds.Tables["TEMPds1"].Rows[0]["保養日期"].ToString());
                     

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
                sbSql.Append(" UPDATE [TKMOC].[dbo].[MACHINEDAILYCHECK]");
                sbSql.AppendFormat("  SET [MAINDATE]='{1}',[CHECK1]='{2}',[CHECK2]='{3}',[CHECK3]='{4}',[CHECK4]='{5}',[CHECK5]='{6}',[CHECK6]='{7}',[CHECK7]='{8}',[CHECK8]='{9}',[CHECK9]='{10}',[CHECK10]='{11}',[CHECK11]='{12}',[CHECKOR]='{13}'  WHERE [ID]='{0}' ", textID.Text.ToString(), dateTimePicker1.Value.ToString("yyyy/MM/dd"), comboBox3.Text.ToString(), comboBox4.Text.ToString(), comboBox5.Text.ToString(), comboBox6.Text.ToString(), comboBox7.Text.ToString(), comboBox8.Text.ToString(), comboBox9.Text.ToString(), comboBox10.Text.ToString(), comboBox11.Text.ToString(), comboBox12.Text.ToString(), comboBox13.Text.ToString(),textBox2.Text.ToString());
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
                sbSql.Append(" INSERT INTO [TKMOC].[dbo].[MACHINEDAILYCHECK]  ");
                sbSql.Append(" ([ID],[EQUIPMENTID],[EQUIPMENTNAME],[UNIT],[MAINDATE],[CHECK1],[CHECK2],[CHECK3],[CHECK4],[CHECK5],[CHECK6],[CHECK7],[CHECK8],[CHECK9],[CHECK10],[CHECK11],[CHECKOR])  ");
                sbSql.AppendFormat("  VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}') ", Guid.NewGuid(),comboBox2.SelectedValue.ToString(),textBox1.Text.ToString(),comboBox2.SelectedValue.ToString(), dateTimePicker1.Value.ToString("yyyy/MM/dd"), comboBox3.Text.ToString(), comboBox4.Text.ToString(), comboBox5.Text.ToString(), comboBox6.Text.ToString(), comboBox7.Text.ToString(), comboBox8.Text.ToString(), comboBox9.Text.ToString(), comboBox10.Text.ToString(), comboBox11.Text.ToString(), comboBox12.Text.ToString(), comboBox13.Text.ToString(), textBox2.Text.ToString());

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
