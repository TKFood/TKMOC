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
    public partial class frmMOCMANULINESubTEMP : Form
    {
        private ComponentResourceManager _ResourceManager = new ComponentResourceManager();
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        SqlDataAdapter adapter20 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder20 = new SqlCommandBuilder();
        DataSet ds20 = new DataSet();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();

        string EDITID;
        int result;
        int BOXNUMERB;

        public frmMOCMANULINESubTEMP()
        {
            InitializeComponent();
        }

        public frmMOCMANULINESubTEMP(string ID)
        {
            EDITID = ID;
            InitializeComponent();

            comboBox1load();

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
        public void comboBox1load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MD001,MD002 FROM [TK].dbo.CMSMD    WHERE MD002 LIKE '新廠%'   ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MD001", typeof(string));
            dt.Columns.Add("MD002", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "MD002";
            comboBox1.DisplayMember = "MD002";
            sqlConn.Close();


        }
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
                sbSql.AppendFormat(@"  ,[MC004],CONVERT(varchar(100),[OUTDATE],112)  AS '交期',[TA029] AS '備註' ,[HALFPRO] AS '半成品數量'");
                sbSql.AppendFormat(@"  ,[MANUHOUR] AS 生產時間 ");
                sbSql.AppendFormat(@"  ,[COPTD001] AS '訂單單別',[COPTD002] AS '訂單號',[COPTD003] AS '訂單序號'");
                sbSql.AppendFormat(@"  ,[ID]");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINETEMP],[TK].[dbo].[BOMMC]");
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
                        
                        textBox3.Text = ds1.Tables["TEMPds1"].Rows[0]["品號"].ToString();
                        textBox4.Text = ds1.Tables["TEMPds1"].Rows[0]["品名"].ToString();
                        textBox5.Text = ds1.Tables["TEMPds1"].Rows[0]["規格"].ToString();
                        textBox6.Text = ds1.Tables["TEMPds1"].Rows[0]["桶數"].ToString();
                        textBox7.Text = ds1.Tables["TEMPds1"].Rows[0]["數量"].ToString();
                        textBox8.Text = ds1.Tables["TEMPds1"].Rows[0]["片數"].ToString();
                        textBox9.Text = ds1.Tables["TEMPds1"].Rows[0]["箱數"].ToString();
                        textBox10.Text = ds1.Tables["TEMPds1"].Rows[0]["客戶"].ToString();
                        textBox32.Text = ds1.Tables["TEMPds1"].Rows[0]["MC004"].ToString();
                        textBox2.Text = ds1.Tables["TEMPds1"].Rows[0]["備註"].ToString();
                        textBox13.Text = ds1.Tables["TEMPds1"].Rows[0]["生產時間"].ToString();
                        textBox12.Text = ds1.Tables["TEMPds1"].Rows[0]["半成品數量"].ToString();
                        textBox40.Text = ds1.Tables["TEMPds1"].Rows[0]["訂單單別"].ToString();
                        textBox41.Text = ds1.Tables["TEMPds1"].Rows[0]["訂單號"].ToString();
                        textBox42.Text = ds1.Tables["TEMPds1"].Rows[0]["訂單序號"].ToString();

                        comboBox1.Text = ds1.Tables["TEMPds1"].Rows[0]["線別"].ToString();

                        string yy = ds1.Tables["TEMPds1"].Rows[0]["生產日"].ToString().Substring(0, 4);
                        string MM = ds1.Tables["TEMPds1"].Rows[0]["生產日"].ToString().Substring(4, 2);
                        string dd = ds1.Tables["TEMPds1"].Rows[0]["生產日"].ToString().Substring(6, 2);

                        dateTimePicker1.Value = Convert.ToDateTime(yy+"/"+MM+"/"+dd);

                        if(!String.IsNullOrEmpty(ds1.Tables["TEMPds1"].Rows[0]["交期"].ToString()))
                        {
                            string OUTyy = ds1.Tables["TEMPds1"].Rows[0]["交期"].ToString().Substring(0, 4);
                            string OUTMM = ds1.Tables["TEMPds1"].Rows[0]["交期"].ToString().Substring(4, 2);
                            string OUTdd = ds1.Tables["TEMPds1"].Rows[0]["交期"].ToString().Substring(6, 2);

                            dateTimePicker2.Value = Convert.ToDateTime(OUTyy + "/" + OUTMM + "/" + OUTdd);
                        }
                        else
                        {
                            dateTimePicker2.Format = DateTimePickerFormat.Custom;
                            dateTimePicker2.CustomFormat = " ";
                        }

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

                sbSql.AppendFormat(" UPDATE [TKMOC].[dbo].[MOCMANULINETEMP] ");
                sbSql.AppendFormat(" SET [BAR]={0},[NUM]={1},[PACKAGE]={2},[BOX]={3},[CLINET]='{4}',MANUDATE='{5}'", textBox6.Text, textBox7.Text, textBox8.Text, textBox9.Text, textBox10.Text, dateTimePicker1.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(" ,MANU='{0}',[OUTDATE]='{1}',[TA029]=N'{2}',[MANUHOUR]={3},HALFPRO={4},COPTD001='{5}'", textBox1.Text, dateTimePicker2.Value.ToString("yyyyMMdd"), textBox2.Text, textBox13.Text, textBox12.Text, textBox40.Text);
                sbSql.AppendFormat(" ,COPTD002='{0}',COPTD003='{1}'", textBox41.Text, textBox42.Text);
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
                
                if (Decimal.TryParse(textBox8.Text, out num1) && Decimal.TryParse(textBox32.Text, out num2))
                {
                    textBox9.Text = Math.Round(Convert.ToDecimal(textBox8.Text) / Convert.ToDecimal(textBox32.Text) / BOXNUMERB, 4).ToString();
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
            //SEARCHMB001BOX();
            //textBox11.Text = BOXNUMERB.ToString();

            //CALPRODUCTDETAIL();
        }

        public void SEARCHMB001BOX()
        {
           

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT TOP 1 MD001,MD003,MB001,MB002,ISNULL(MD007,1) AS MD007,ISNULL(MD010,1) AS MD010");
                sbSql.AppendFormat(@"  FROM [TK].dbo.BOMMD,[TK].dbo.INVMB");
                sbSql.AppendFormat(@"  WHERE MD003=MB001");
                sbSql.AppendFormat(@"  AND MB002 LIKE '%箱%'");
                sbSql.AppendFormat(@"  AND MD003 LIKE '2%'");
                sbSql.AppendFormat(@"  AND MD001='{0}'", textBox3.Text);
                sbSql.AppendFormat(@"  ");

                adapter20 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder20 = new SqlCommandBuilder(adapter20);
                sqlConn.Open();
                ds20.Clear();
                adapter20.Fill(ds20, "TEMPds20");
                sqlConn.Close();


                if (ds20.Tables["TEMPds20"].Rows.Count == 0)
                {
                    BOXNUMERB = 1;
                }
                else
                {
                    if (ds20.Tables["TEMPds20"].Rows.Count >= 1)
                    {
                        BOXNUMERB = (Convert.ToInt32(ds20.Tables["TEMPds20"].Rows[0]["MD007"].ToString()) / Convert.ToInt32(ds20.Tables["TEMPds20"].Rows[0]["MD010"].ToString()));
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

        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            SEARCHMB001BOX();
            textBox11.Text = BOXNUMERB.ToString();

            CALPRODUCTDETAIL();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox1.Text = comboBox1.Text;
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
