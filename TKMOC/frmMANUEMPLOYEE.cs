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
using TKITDLL;

namespace TKMOC
{
    public partial class frmMANUEMPLOYEE : Form
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
        DataGridViewRow drEMPLOYEE = new DataGridViewRow();
        string tablename = null;
        string ID;
        int result;
        int rownum = 0;
        Thread TD;


        public frmMANUEMPLOYEE()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void Search()
        {
            StringBuilder Query = new StringBuilder();

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

                sbSql.Append(@" SELECT [ID] AS '工號',[NAME] AS '姓名'  FROM [TKMOC].[dbo].[MANUEMPLOYEE] ORDER BY [ID]");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {
                    labelSearch.Text = "找不到資料";
                }
                else
                {
                    if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                    {

                        labelSearch.Text = "有 " + ds.Tables["TEMPds1"].Rows.Count.ToString() + " 筆";
                        //dataGridView1.Rows.Clear();
                        dataGridView1.DataSource = ds.Tables["TEMPds1"];
                        dataGridView1.AutoResizeColumns();
                        dataGridView1.CurrentCell = dataGridView1[0, rownum];
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
                sbSql.Append(" UPDATE [TKMOC].[dbo].[MANUEMPLOYEE] ");
                sbSql.AppendFormat(" SET [NAME]='{1}' WHERE [ID]='{0}' ", textBox1.Text.ToString(), textBox2.Text.ToString());
                sbSql.Append("  ");

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
                    rownum = dataGridView1.CurrentCell.RowIndex;

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

        public void SetUPDATE()
        {
            textBox1.ReadOnly = false;
            textBox2.ReadOnly = false;
           
        }

        public void SetFINISH()
        {
            textBox1.ReadOnly = true;
            textBox2.ReadOnly = true;
          
        }

        public void ADDNEW()
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
                sbSql.Append("INSERT INTO [TKMOC].dbo.MANUEMPLOYEE (ID,NAME) ");
                sbSql.AppendFormat("  SELECT [Code],[CnName]    FROM [HRMDB].[dbo].[Employee] WHERE [DepartmentId] IN ( SELECT [DepartmentId]FROM [HRMDB].[dbo].[Department] WHERE CODE LIKE '103%') AND  Code   COLLATE Chinese_Taiwan_Stroke_BIN NOT IN (SELECT ID FROM [TKMOC].dbo.MANUEMPLOYEE WITH (NOLOCK) ) ");
                sbSql.Append("  ");

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
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count >= 1)
            {
                drEMPLOYEE = dataGridView1.Rows[dataGridView1.SelectedCells[0].RowIndex];
                textBox1.Text = drEMPLOYEE.Cells["工號"].Value.ToString();
                textBox2.Text = drEMPLOYEE.Cells["姓名"].Value.ToString();

            }
        }


        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }

        #endregion

        private void button2_Click(object sender, EventArgs e)
        {
            ADDNEW();
            Search();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SetUPDATE();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            UPDATE();
            SetFINISH();
            Search();
        }


    }
}

