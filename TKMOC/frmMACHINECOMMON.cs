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
    public partial class frmMACHINECOMMON : Form
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

        public frmMACHINECOMMON()
        {
            InitializeComponent();
        }

        public frmMACHINECOMMON(string ID)
        {
            InitializeComponent();
           
            if (!string.IsNullOrEmpty(ID))
            {
                EDITID = ID;
                Search(ID);
            }
        }

        #region FUNCTION

        public void Search(string ID)
        {
            StringBuilder Query = new StringBuilder();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@" SELECT [EQUIPMENTID] AS '備品編號',[EQUIPMENTNAME] AS '品名',[SPEC] AS '規格'");
                sbSql.AppendFormat(@" ,[PRICES] AS '單價',[SUPPLY] AS '供應商',[TEL] AS '電話'");
                sbSql.AppendFormat(@" ,[USEDTIME] AS '使用壽命',[SAFENUM] AS '安全庫存',[PRETIME] AS '前置時間'");
                sbSql.AppendFormat(@" ,[EQUIPMENT] AS '適用設備'");
                sbSql.AppendFormat(@" ,[COMMENT] AS '備註'");
                sbSql.AppendFormat(@" ,[ID]");
                sbSql.AppendFormat(@" FROM [TKMOC].[dbo].[MACHINECOMMON]");
                sbSql.AppendFormat(@" WHERE [ID]='{0}'",ID);
                sbSql.AppendFormat(@" ");



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
                        textBox1.Text = ds.Tables["TEMPds1"].Rows[0]["備品編號"].ToString();
                        textBox2.Text = ds.Tables["TEMPds1"].Rows[0]["品名"].ToString();
                        textBox3.Text = ds.Tables["TEMPds1"].Rows[0]["規格"].ToString();
                        textBox4.Text = ds.Tables["TEMPds1"].Rows[0]["單價"].ToString();
                        textBox5.Text = ds.Tables["TEMPds1"].Rows[0]["供應商"].ToString();
                        textBox6.Text = ds.Tables["TEMPds1"].Rows[0]["電話"].ToString();
                        textBox7.Text = ds.Tables["TEMPds1"].Rows[0]["使用壽命"].ToString();
                        textBox8.Text = ds.Tables["TEMPds1"].Rows[0]["安全庫存"].ToString();
                        textBox9.Text = ds.Tables["TEMPds1"].Rows[0]["前置時間"].ToString();
                        textBox10.Text = ds.Tables["TEMPds1"].Rows[0]["適用設備"].ToString();
                        textBox11.Text = ds.Tables["TEMPds1"].Rows[0]["備註"].ToString();
                        textBox12.Text = ds.Tables["TEMPds1"].Rows[0]["ID"].ToString();


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

                sbSql.AppendFormat("   UPDATE [TKMOC].[dbo].[MACHINECOMMON]");
                sbSql.AppendFormat("   SET [EQUIPMENTID]='{1}',[EQUIPMENTNAME]='{2}',[SPEC]='{3}',[PRICES]='{4}',[SUPPLY]='{5}',[TEL]='{6}',[USEDTIME]='{7}',[SAFENUM]='{8}',[PRETIME]='{9}',[EQUIPMENT]='{10}',[COMMENT]='{11}' WHERE [ID]='{0}'", EDITID, textBox1.Text, textBox2.Text, textBox3.Text, textBox4.Text, textBox5.Text, textBox6.Text, textBox7.Text, textBox8.Text, textBox9.Text, textBox10.Text, textBox11.Text);
                sbSql.AppendFormat("   ");
                sbSql.AppendFormat("   ");


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
                //add
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat("  INSERT INTO [TKMOC].[dbo].[MACHINECOMMON]");
                sbSql.AppendFormat("  ([ID],[EQUIPMENTID],[EQUIPMENTNAME],[SPEC],[PRICES],[SUPPLY],[TEL],[USEDTIME],[SAFENUM],[PRETIME],[EQUIPMENT],[COMMENT])");
                sbSql.AppendFormat("  VALUES ({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}')", "NEWID()",textBox1.Text, textBox2.Text, textBox3.Text, textBox4.Text, textBox5.Text, textBox6.Text, textBox7.Text, textBox8.Text, textBox9.Text, textBox10.Text, textBox11.Text);
                sbSql.AppendFormat("  ");

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
