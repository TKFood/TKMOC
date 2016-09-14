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
    public partial class frmMACHINECARD : Form
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

        public frmMACHINECARD()
        {
            InitializeComponent();
            combobox1load();
            combobox2load();
        }
        public frmMACHINECARD(string ID)
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

                sbSql.AppendFormat(@" SELECT  [EQUIPMENTID] AS '設備編號',[EQUIPMENTNAME] AS '設備名稱',[VALUE] AS '價值',[TYPE] AS '型號',[WEIGHT] AS '重量',[MACHINECODE] AS '機械製造號碼',[SIZE] AS '外形尺寸',[FACTORY] AS '製造廠商',[MACHINEID] AS '機器編號',[SELLFACTORY] AS '出售廠商',[MACHYEAR] AS '製造年份',[UNIT] AS '使用單位',[BUYDATE] AS '購入日期',[OWNER] AS '保管人',[STATUS] AS '重要規格'  ,USEWATER AS '用水tom/hr',USEPOWER AS '電力kW',USEAIR AS '空氣m3/min'  ,MANAGER AS '主管' ,CREATOR  AS '建卡人',CREATEDATE  AS '建卡日期' ");
                sbSql.Append(@"FROM [TKMOC].[dbo].[MACHINECARD] WITH (NOLOCK)");
                sbSql.AppendFormat(@" WHERE [EQUIPMENTID] ='{0}'", ID);
                sbSql.Append(@" ORDER BY [EQUIPMENTID]  ");


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
                        textBox2.Text = ds.Tables["TEMPds1"].Rows[0]["價值"].ToString();
                        textBox3.Text = ds.Tables["TEMPds1"].Rows[0]["型號"].ToString();
                        textBox4.Text = ds.Tables["TEMPds1"].Rows[0]["重量"].ToString();
                        textBox5.Text = ds.Tables["TEMPds1"].Rows[0]["機械製造號碼"].ToString();
                        textBox6.Text = ds.Tables["TEMPds1"].Rows[0]["外形尺寸"].ToString();
                        textBox7.Text = ds.Tables["TEMPds1"].Rows[0]["製造廠商"].ToString();
                        textBox8.Text = ds.Tables["TEMPds1"].Rows[0]["機器編號"].ToString();
                        textBox9.Text = ds.Tables["TEMPds1"].Rows[0]["出售廠商"].ToString();
                        textBox10.Text = ds.Tables["TEMPds1"].Rows[0]["製造年份"].ToString();
                        textBox11.Text = ds.Tables["TEMPds1"].Rows[0]["保管人"].ToString();
                        textBox12.Text = ds.Tables["TEMPds1"].Rows[0]["重要規格"].ToString();
                        textBox13.Text = ds.Tables["TEMPds1"].Rows[0]["用水tom/hr"].ToString();
                        textBox14.Text = ds.Tables["TEMPds1"].Rows[0]["電力kW"].ToString();
                        textBox15.Text = ds.Tables["TEMPds1"].Rows[0]["空氣m3/min"].ToString();
                        textBox16.Text = ds.Tables["TEMPds1"].Rows[0]["主管"].ToString();
                        textBox17.Text = ds.Tables["TEMPds1"].Rows[0]["建卡人"].ToString();
                        comboBox1.SelectedValue = ds.Tables["TEMPds1"].Rows[0]["使用單位"].ToString();
                        comboBox2.SelectedValue = ds.Tables["TEMPds1"].Rows[0]["設備編號"].ToString();
                        dateTimePicker1.Value = Convert.ToDateTime(ds.Tables["TEMPds1"].Rows[0]["購入日期"].ToString());
                        dateTimePicker2.Value = Convert.ToDateTime(ds.Tables["TEMPds1"].Rows[0]["建卡日期"].ToString());

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
                sbSql.Append(" UPDATE [TKMOC].[dbo].[MACHINECARD]");
                sbSql.AppendFormat("  SET [EQUIPMENTNAME]='{1}',[VALUE]='{2}',[TYPE]='{3}',[WEIGHT]='{4}',[MACHINECODE]='{5}',[SIZE]='{6}',[FACTORY]='{7}',[MACHINEID]='{8}',[SELLFACTORY]='{9}',[MACHYEAR]='{10}',[UNIT]='{11}',[BUYDATE]='{12}',[OWNER]='{13}',[STATUS]='{14}',[USEWATER]='{15}',[USEPOWER]='{16}',[USEAIR]='{17}',[MANAGER]='{18}',[CREATOR]='{19}',[CREATEDATE]='{20}' WHERE [EQUIPMENTID]='{0}' ", comboBox2.SelectedValue.ToString(),textBox1.Text.ToString(), textBox2.Text.ToString(), textBox3.Text.ToString(), textBox4.Text.ToString(), textBox5.Text.ToString(), textBox6.Text.ToString(), textBox7.Text.ToString(), textBox8.Text.ToString(), textBox9.Text.ToString(), textBox10.Text.ToString(),comboBox1.SelectedValue.ToString(),dateTimePicker1.Value.ToString("yyyy/MM/dd"), textBox11.Text.ToString(), textBox12.Text.ToString(), textBox13.Text.ToString(), textBox14.Text.ToString(), textBox15.Text.ToString(), textBox16.Text.ToString(), textBox17.Text.ToString(),dateTimePicker2.Value.ToString("yyyy/MM/dd"));
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
                sbSql.Append(" INSERT INTO [TKMOC].[dbo].[MACHINECARD]  ");
                sbSql.Append(" ([EQUIPMENTID],[EQUIPMENTNAME],[VALUE],[TYPE],[WEIGHT],[MACHINECODE],[SIZE],[FACTORY],[MACHINEID],[SELLFACTORY],[MACHYEAR],[UNIT],[BUYDATE],[OWNER],[STATUS],[USEWATER],[USEPOWER],[USEAIR],[MANAGER],[CREATOR],[CREATEDATE] )  ");
                sbSql.AppendFormat("  VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}','{19}','{20}') ", comboBox2.SelectedValue.ToString(), textBox1.Text.ToString(), textBox2.Text.ToString(), textBox3.Text.ToString(), textBox4.Text.ToString(), textBox5.Text.ToString(), textBox6.Text.ToString(), textBox7.Text.ToString(), textBox8.Text.ToString(), textBox9.Text.ToString(), textBox10.Text.ToString(), comboBox1.SelectedValue.ToString(), dateTimePicker1.Value.ToString("yyyy/MM/dd"), textBox11.Text.ToString(), textBox12.Text.ToString(), textBox13.Text.ToString(), textBox14.Text.ToString(), textBox15.Text.ToString(), textBox16.Text.ToString(), textBox17.Text.ToString(), dateTimePicker2.Value.ToString("yyyy/MM/dd"));

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
