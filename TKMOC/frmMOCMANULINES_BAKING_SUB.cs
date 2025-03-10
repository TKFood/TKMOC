﻿using System;
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
using TKITDLL;

namespace TKMOC
{
    public partial class frmMOCMANULINES_BAKING_SUB : Form
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
        decimal BOXNUMERB;

        public frmMOCMANULINES_BAKING_SUB()
        {
            InitializeComponent();
        }

        public frmMOCMANULINES_BAKING_SUB(string ID)
        {
            EDITID = ID;
            InitializeComponent();

            SEARCH_MOCMANULINEBAKING();
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


        public void SEARCH_MOCMANULINEBAKING()
        {

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();




                sbSql.AppendFormat(@"  
                                     SELECT 
                                     [MANU] AS '線別',CONVERT(varchar(100),[MANUDATE],112) AS '生產日',[MB001] AS '品號',[MB002] AS '品名'
                                     ,[MB003] AS '規格',ISNULL([BAR],0) AS '桶數',ISNULL([NUM],0) AS '數量',ISNULL([BOX],0)   AS '箱數'   ,ISNULL([PACKAGE],0)  AS '包裝數',[CLINET] AS '客戶'
                                     ,[MC004],CONVERT(varchar(100),[OUTDATE],112)  AS '交期',[TA029] AS '備註' ,[HALFPRO] AS '半成品數量'
                                     ,[MANUHOUR] AS 生產時間 
                                     ,[COPTD001] AS '訂單單別',[COPTD002] AS '訂單號',[COPTD003] AS '訂單序號'
                                    ,[MANUPRENUMS] AS '需多投數量做底'
                                     ,[ID]
                                     FROM [TKMOC].[dbo].[MOCMANULINEBAKING],[TK].[dbo].[BOMMC]
                                     WHERE [MB001]=[MC001]
                                     AND  [ID]='{0}'
                                    ", EDITID);


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
                        textBox8.Text = ds1.Tables["TEMPds1"].Rows[0]["箱數"].ToString();
                        textBox9.Text = ds1.Tables["TEMPds1"].Rows[0]["包裝數"].ToString();
                        textBox10.Text = ds1.Tables["TEMPds1"].Rows[0]["客戶"].ToString();
                        textBox32.Text = ds1.Tables["TEMPds1"].Rows[0]["MC004"].ToString();
                        textBox2.Text = ds1.Tables["TEMPds1"].Rows[0]["備註"].ToString();
                        textBox13.Text = ds1.Tables["TEMPds1"].Rows[0]["生產時間"].ToString();
                        textBox12.Text = ds1.Tables["TEMPds1"].Rows[0]["半成品數量"].ToString();
                        textBox40.Text = ds1.Tables["TEMPds1"].Rows[0]["訂單單別"].ToString();
                        textBox41.Text = ds1.Tables["TEMPds1"].Rows[0]["訂單號"].ToString();
                        textBox42.Text = ds1.Tables["TEMPds1"].Rows[0]["訂單序號"].ToString();
                        textBox99.Text = ds1.Tables["TEMPds1"].Rows[0]["需多投數量做底"].ToString();

                        string yy = ds1.Tables["TEMPds1"].Rows[0]["生產日"].ToString().Substring(0, 4);
                        string MM = ds1.Tables["TEMPds1"].Rows[0]["生產日"].ToString().Substring(4, 2);
                        string dd = ds1.Tables["TEMPds1"].Rows[0]["生產日"].ToString().Substring(6, 2);

                        dateTimePicker1.Value = Convert.ToDateTime(yy + "/" + MM + "/" + dd);

                        if (!String.IsNullOrEmpty(ds1.Tables["TEMPds1"].Rows[0]["交期"].ToString()))
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

        public void UPDATE_MOCMANULINEBAKING()
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@" 
                                     UPDATE [TKMOC].[dbo].[MOCMANULINEBAKING] 
                                     SET [BAR]={1},[NUM]={2},[BOX]={3},[PACKAGE]={4},[CLINET]='{5}',MANUDATE='{6}',[OUTDATE]='{7}',[TA029]=N'{8}',[MANUHOUR]={9},HALFPRO={10},COPTD001='{11}',COPTD002='{12}',COPTD003='{13}',MANUPRENUMS='{14}'
                                     WHERE  [ID]='{0}'"
                                     , textBoxID.Text, textBox6.Text, textBox7.Text, textBox8.Text, textBox9.Text, textBox10.Text, dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"), textBox2.Text, textBox13.Text, textBox12.Text, textBox40.Text, textBox41.Text, textBox42.Text, textBox99.Text);
                sbSql.AppendFormat(@" ");


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
                if (Decimal.TryParse(textBox7.Text, out num1) && Decimal.TryParse(textBox99.Text, out num1) && Decimal.TryParse(textBox32.Text, out num2))
                {
                    if (Convert.ToDecimal(textBox7.Text) >= 0 & Convert.ToDecimal(textBox99.Text) >= 0 & Convert.ToDecimal(textBox32.Text) > 0)
                    {
                        textBox6.Text = Math.Round((Convert.ToDecimal(textBox7.Text) + Convert.ToDecimal(textBox99.Text)) / Convert.ToDecimal(textBox32.Text), 4).ToString();
                    }
                    else
                    {
                        textBox6.Text = "0";
                    }

                }

                if (Decimal.TryParse(textBox9.Text, out num1) && Decimal.TryParse(textBox32.Text, out num2))
                {
                    if (Convert.ToDecimal(textBox9.Text) > 0 & Convert.ToDecimal(textBox32.Text) > 0 & BOXNUMERB > 0)
                    {
                        textBox8.Text = Math.Round(Convert.ToDecimal(textBox9.Text) / Convert.ToDecimal(textBox32.Text) / BOXNUMERB, 4).ToString();
                    }
                    else
                    {
                        textBox8.Text = "0";
                    }

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

        private void textBox99_TextChanged(object sender, EventArgs e)
        {
            CALPRODUCTDETAIL();
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            SEARCHMB001BOX();
            textBox11.Text = BOXNUMERB.ToString();

            CALPRODUCTDETAIL();
        }

        public void SEARCHMB001BOX()
        {


            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@" 
                                        SELECT  MD001,MD003,MB001,MB002,ISNULL(MD007,1) AS MD007,ISNULL(MD010,1) AS MD010,ISNULL(MD006,1) AS MD006
                                        FROM [TK].dbo.BOMMD,[TK].dbo.INVMB
                                        WHERE MD003=MB001
                                        AND MB002 LIKE '%箱%'
                                        AND MD003 LIKE '2%'
                                        AND MD001 LIKE '{0}%'
                                    ", textBox3.Text);

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
                        if (Convert.ToDecimal(ds20.Tables["TEMPds20"].Rows[0]["MD007"].ToString()) > 0 & Convert.ToDecimal(ds20.Tables["TEMPds20"].Rows[0]["MD006"].ToString()) > 0)
                        {
                            BOXNUMERB = (Convert.ToDecimal(ds20.Tables["TEMPds20"].Rows[0]["MD007"].ToString()) / Convert.ToDecimal(ds20.Tables["TEMPds20"].Rows[0]["MD006"].ToString()));
                        }
                        else
                        {
                            BOXNUMERB = 1;
                        }

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

        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            UPDATE_MOCMANULINEBAKING();

        }
        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        #endregion


    }
}
