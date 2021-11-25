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
using Calendar.NET;
using TKITDLL;

namespace TKMOC
{
    public partial class frmMOCMANULINECAPACITYCAL : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        StringBuilder sbSqlQuery2 = new StringBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        int result;
        DataSet ds1 = new DataSet();

        public frmMOCMANULINECAPACITYCAL()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void ADDMOCMANULINECAPACITYCAL()
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
                                    DELETE [TKMOC].[dbo].[MOCMANULINECAPACITYCAL]
                                    WHERE CONVERT(NVARCHAR,[MOCDATES],112)>='{0}' AND  CONVERT(NVARCHAR,[MOCDATES],112)<='{1}'

                                    INSERT INTO [TKMOC].[dbo].[MOCMANULINECAPACITYCAL]
                                    ([MOCDATES],[YEARS],[WEEKS],[LINEBIG],[LINESMALL],[LINEMANU])
                                    SELECT CONVERT(NVARCHAR,[MANUDATE],111) [MANUDATE],DATEPART (YEAR,[MANUDATE] ),DATEPART ( WEEK ,[MANUDATE] )
                                    ,(SELECT  ISNULL(SUM([BAR]),0) FROM [TKMOC].[dbo].[MOCMANULINE] MOC WHERE MOC.[MANUDATE]=[MOCMANULINE].[MANUDATE] AND [MANU]='製一線') 'LINEBIG'
                                    ,(SELECT  ISNULL(SUM([BAR]),0) FROM [TKMOC].[dbo].[MOCMANULINE] MOC WHERE MOC.[MANUDATE]=[MOCMANULINE].[MANUDATE] AND [MANU]='製二線') 'LINESMALL'
                                    ,(SELECT  ISNULL(SUM([BAR]),0) FROM [TKMOC].[dbo].[MOCMANULINE] MOC WHERE MOC.[MANUDATE]=[MOCMANULINE].[MANUDATE] AND [MANU]='手工線') 'LINEMANU'
                                    FROM [TKMOC].[dbo].[MOCMANULINE]
                                    WHERE  CONVERT(NVARCHAR,[MANUDATE],112)>='{0}' AND CONVERT(NVARCHAR,[MANUDATE],112)<='{1}'
                                    GROUP BY [MANUDATE]
                                    ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));


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


        public void SEARCHMOCMANULINECAPACITYCAL()
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
                                     CONVERT(NVARCHAR,[MOCDATES],112) AS '預排日'
                                    ,[LINEBIG] AS '大線桶數'
                                    ,[LINESMALL] AS '小線桶數'
                                    ,[LINEBIGCAP] AS '大線產能'
                                    ,[LINESMALLCAP] AS '小線產能'
                                    ,[LINEBIGCAL] AS '大線稼動率'
                                    ,[LINESMALLCAL] AS '小線稼動率'
                                    FROM [TKMOC].[dbo].[MOCMANULINECAPACITYCAL]
                                    WHERE CONVERT(NVARCHAR,[MOCDATES],112)>='{0}' AND  CONVERT(NVARCHAR,[MOCDATES],112)<='{1}'

                                    ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView1.DataSource = ds1.Tables["TEMPds1"];
                        dataGridView1.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                    }
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
        private void button6_Click(object sender, EventArgs e)
        {
            ADDMOCMANULINECAPACITYCAL();
            SEARCHMOCMANULINECAPACITYCAL();

        }
        #endregion

    }
}
