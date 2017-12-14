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
    public partial class frmMOCMANULINESub : Form
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

        string EDITID;
        int result;

        public frmMOCMANULINESub()
        {
            InitializeComponent();
        }

        public frmMOCMANULINESub(string ID)
        {
            EDITID = ID;
            InitializeComponent();

            SEARCHMOCMANULINE();
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        #region FUNCTION

        #endregion
        public void SEARCHMOCMANULINE()
        {

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();



                sbSql.AppendFormat(@"  SELECT ");
                sbSql.AppendFormat(@"  [MANU] AS '線別',CONVERT(varchar(100),[MANUDATE],112) AS '生產日',[MB001] AS '品號',[MB002] AS '品名'");
                sbSql.AppendFormat(@"  ,[MB003] AS '規格',ISNULL([BAR],0) AS '桶數',ISNULL([NUM],0) AS '數量',ISNULL([BOX],0)   AS '箱數'   ,ISNULL([PACKAGE],0)  AS '片數',[CLINET] AS '客戶'");
                sbSql.AppendFormat(@"  ,[MC004]");
                sbSql.AppendFormat(@"  ,[ID]");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINE],[TK].[dbo].[BOMMC]");
                sbSql.AppendFormat(@"  WHERE [MB001]=[MC001]");
                sbSql.AppendFormat(@"  AND  [ID]='{0}'", EDITID);
                sbSql.AppendFormat(@"  ");


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
                        textBox1.Text = ds1.Tables["TEMPds1"].Rows[0]["線別"].ToString();
                        textBox2.Text = ds1.Tables["TEMPds1"].Rows[0]["生產日"].ToString();
                        textBox3.Text = ds1.Tables["TEMPds1"].Rows[0]["品號"].ToString();
                        textBox4.Text = ds1.Tables["TEMPds1"].Rows[0]["品名"].ToString();
                        textBox5.Text = ds1.Tables["TEMPds1"].Rows[0]["規格"].ToString();
                        textBox6.Text = ds1.Tables["TEMPds1"].Rows[0]["桶數"].ToString();
                        textBox7.Text = ds1.Tables["TEMPds1"].Rows[0]["數量"].ToString();
                        textBox8.Text = ds1.Tables["TEMPds1"].Rows[0]["箱數"].ToString();
                        textBox9.Text = ds1.Tables["TEMPds1"].Rows[0]["片數"].ToString();
                        textBox10.Text = ds1.Tables["TEMPds1"].Rows[0]["客戶"].ToString();
                        textBox32.Text = ds1.Tables["TEMPds1"].Rows[0]["MC004"].ToString();

                        textBoxID.Text = ds1.Tables["TEMPds1"].Rows[0]["ID"].ToString();
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

        public void UPDATEMOCMANULINE()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" UPDATE [TKMOC].[dbo].[MOCMANULINE] SET [BAR]={0},[NUM]={1},[BOX]={2},[PACKAGE]={3},[CLINET]='{4}'", textBox6.Text, textBox7.Text, textBox8.Text, textBox9.Text, textBox10.Text);
                sbSql.AppendFormat(" WHERE  [ID]='{0}'", textBoxID.Text);
                sbSql.AppendFormat(" ");


                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                    MessageBox.Show("更新失敗");
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
        public void CALPRODUCTDETAIL()
        {
            Decimal num1;
            Decimal num2;
            try
            {
                if (Decimal.TryParse(textBox7.Text, out num1) && Decimal.TryParse(textBox32.Text, out num2))
                {
                    textBox6.Text = Math.Round(Convert.ToDecimal(textBox7.Text) / Convert.ToDecimal(textBox32.Text), 4).ToString();
                }

                if (Decimal.TryParse(textBox9.Text, out num1) && Decimal.TryParse(textBox32.Text, out num2))
                {
                    textBox8.Text = Math.Round(Convert.ToDecimal(textBox9.Text) / Convert.ToDecimal(textBox32.Text), 4).ToString();
                }

            }
            catch
            {
                //MessageBox.Show("請填數字");
            }
            finally
            {

            }

        }
        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            CALPRODUCTDETAIL();
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            CALPRODUCTDETAIL();
        }

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            UPDATEMOCMANULINE();

        }
        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        #endregion

        
    }
}
