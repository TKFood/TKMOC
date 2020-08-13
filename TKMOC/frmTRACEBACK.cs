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
using Calendar.NET;

namespace TKMOC
{
    public partial class frmTRACEBACK : Form
    {
        private ComponentResourceManager _ResourceManager = new ComponentResourceManager();
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();

        int result;

        public frmTRACEBACK()
        {
            InitializeComponent();
        }

        #region FUNCTION

        public void SEARCHOUT(string MB001,string LOTNO)
        {
            StringBuilder sbSql = new StringBuilder();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();

                sbSql.AppendFormat(@"  SELECT MF001,MF002,'0',MF003,MF004,MF005,MF006,MF010");
                sbSql.AppendFormat(@"  FROM [TK].dbo.INVME WITH (NOLOCK),[TK].dbo.INVMF WITH (NOLOCK)");
                sbSql.AppendFormat(@"  WHERE MF001=ME001 AND MF002=ME002");
                sbSql.AppendFormat(@"  AND MF009 IN ('2','5')");
                sbSql.AppendFormat(@"  AND MF001='{0}' AND MF002='{1}'",MB001,LOTNO);
                sbSql.AppendFormat(@"  ORDER BY MF002,MF003,MF004,MF005");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    ADDTRACEBACKOUT(MB001, LOTNO);
                    ADDTRACEBACKMOC(MB001, LOTNO);
                    ADDTRACEBACKMOCOUTIN(MB001, LOTNO);
                    ADDTRACEBACKINVMF(MB001, LOTNO);
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

        public void ADDTRACEBACKOUT(string MB001, string LOTNO)
        {

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" DELETE [TKMOC].[dbo].[TRACEBACK]");
                sbSql.AppendFormat(" WHERE [MMB001]='{0}' AND [MLOTNO]='{1}'",MB001,LOTNO);
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[TRACEBACK]");
                sbSql.AppendFormat(" ([MMB001],[MLOTNO],[KINDS],[LEVELS],[DATES],[MID],[DID],[SID],[MB001],[MB002],[LOTNO],[NUMS])");
                sbSql.AppendFormat(" SELECT MF001,MF002,'1銷貨','0',MF003,MF004,MF005,MF006,MF001,'',MF002,MF010");
                sbSql.AppendFormat(" FROM [TK].dbo.INVME WITH (NOLOCK),[TK].dbo.INVMF WITH (NOLOCK)");
                sbSql.AppendFormat(" WHERE MF001=ME001 AND MF002=ME002");
                sbSql.AppendFormat(" AND MF009 IN ('2','5')");
                sbSql.AppendFormat(" AND MF001='{0}' AND MF002='{1}'", MB001, LOTNO);
                sbSql.AppendFormat(" ORDER BY INVMF.MF002,MF003,MF004,MF005");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");


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

        public void ADDTRACEBACKMOC(string MB001, string LOTNO)
        {

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
            
                sbSql.AppendFormat(" WITH RTABLES");
                sbSql.AppendFormat(" AS (");
                sbSql.AppendFormat(" SELECT 0 AS LEVELS,[TG001],[TG002],[TG003],[TG004],[TG011],[TG017],[TG014],[TG015],[TE001],[TE002],[TE003],[TE004],[TE005],[TE010]");
                sbSql.AppendFormat(" FROM [TK].[dbo].[VMOCTGMOCTE] WITH (NOLOCK)");
                sbSql.AppendFormat(" WHERE [VMOCTGMOCTE].TG004 ='{0}' AND [VMOCTGMOCTE].TG017 ='{1}' ", MB001, LOTNO);
                sbSql.AppendFormat(" UNION ALL");
                sbSql.AppendFormat(" SELECT LEVELS+1,B.[TG001], B.[TG002], B.[TG003], B.[TG004], B.[TG011], B.[TG017], B.[TG014], B.[TG015], B.[TE001], B.[TE002],B.[TE003], B.[TE004], B.[TE005], B.[TE010]");
                sbSql.AppendFormat(" FROM [TK].[dbo].[VMOCTGMOCTE] B WITH (NOLOCK)");
                sbSql.AppendFormat(" INNER JOIN RTABLES ON RTABLES.[TE004]=B.[TG004] AND RTABLES.[TE010]=B.[TG017]");
                sbSql.AppendFormat(" ) ");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[TRACEBACK]");
                sbSql.AppendFormat(" ([MMB001],[MLOTNO],[KINDS],[LEVELS],[DATES],[MID],[DID],[SID],[MB001],[MB002],[LOTNO],[NUMS])");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" SELECT '{0}','{1}','2生產',LEVELS ",MB001,LOTNO);
                sbSql.AppendFormat(" ,(SELECT TOP 1 TF003 FROM [TK].dbo.MOCTF WHERE TF001=TG001 AND TF002=TG002 ORDER BY TF003)");
                sbSql.AppendFormat(" ,[TG001],[TG002],[TG003],[TG004], '',[TG017],[TG011]");
                sbSql.AppendFormat("  FROM RTABLES");
                sbSql.AppendFormat(" GROUP BY LEVELS,[TG001],[TG002],[TG003],[TG004],[TG017],[TG011]");
                sbSql.AppendFormat(" ORDER BY LEVELS,[TG004]");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");

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

        public void ADDTRACEBACKMOCOUTIN(string MB001, string LOTNO)
        {

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" WITH RTABLES");
                sbSql.AppendFormat(" AS (");
                sbSql.AppendFormat(" SELECT 0 AS LEVELS,[TG001],[TG002],[TG003],[TG004],[TG011],[TG017],[TG014],[TG015],[TE001],[TE002],[TE003],[TE004],[TE005],[TE010]");
                sbSql.AppendFormat(" FROM [TK].[dbo].[VMOCTGMOCTE] WITH (NOLOCK)");
                sbSql.AppendFormat(" WHERE [VMOCTGMOCTE].TG004 ='{0}' AND [VMOCTGMOCTE].TG017 ='{1}' ",MB001,LOTNO);
                sbSql.AppendFormat(" UNION ALL");
                sbSql.AppendFormat(" SELECT LEVELS+1,B.[TG001], B.[TG002], B.[TG003], B.[TG004], B.[TG011], B.[TG017], B.[TG014], B.[TG015], B.[TE001], B.[TE002],B.[TE003], B.[TE004], B.[TE005], B.[TE010]");
                sbSql.AppendFormat(" FROM [TK].[dbo].[VMOCTGMOCTE] B WITH (NOLOCK)");
                sbSql.AppendFormat(" INNER JOIN RTABLES ON RTABLES.[TE004]=B.[TG004] AND RTABLES.[TE010]=B.[TG017]");
                sbSql.AppendFormat(" ) ");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[TRACEBACK]");
                sbSql.AppendFormat(" ([MMB001],[MLOTNO],[KINDS],[LEVELS],[DATES],[MID],[DID],[SID],[MB001],[MB002],[LOTNO],[NUMS])");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" SELECT '{0}','{1}','3領退料',LEVELS", MB001, LOTNO); ;
                sbSql.AppendFormat(" ,(SELECT TOP 1 TC003 FROM [TK].dbo.MOCTC WHERE TC001=TE001  AND TC002=TE002)");
                sbSql.AppendFormat(" ,[TE001],[TE002],[TE003],[TE004],'',[TE010],[TE005] ");
                sbSql.AppendFormat("  FROM RTABLES");
                sbSql.AppendFormat(" GROUP BY LEVELS,[TE001],[TE002],[TE003],[TE004],[TE010],[TE005]");
                sbSql.AppendFormat(" ORDER BY LEVELS,[TE001],[TE002],[TE003],[TE004],[TE010],[TE005]");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");

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

        public void ADDTRACEBACKINVMF(string MB001, string LOTNO)
        {

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[TRACEBACK]");
                sbSql.AppendFormat(" ([MMB001],[MLOTNO],[KINDS],[LEVELS],[DATES],[MID],[DID],[SID],[MB001],[MB002],[LOTNO],[NUMS])");
                sbSql.AppendFormat(" SELECT '{0}','{1}','4入庫','0',MF002,MF004,MF005,MF006,MF001,'',MF002,MF010", MB001, LOTNO);
                sbSql.AppendFormat(" FROM [TK].dbo.INVMF WITH (NOLOCK)");
                sbSql.AppendFormat(" WHERE INVMF.MF009 IN ('1')");
                sbSql.AppendFormat(" AND RTRIM(LTRIM(MF001))+RTRIM(LTRIM(MF002)) IN (SELECT RTRIM(LTRIM([MB001]))+RTRIM(LTRIM([LOTNO])) FROM [TKMOC].[dbo].[TRACEBACK] WHERE MMB001='{0}' AND MLOTNO='{1}')", MB001, LOTNO);
                sbSql.AppendFormat(" AND RTRIM(LTRIM(MF004)) +RTRIM(LTRIM(MF005)) +RTRIM(LTRIM(MF006))  NOT IN (SELECT RTRIM(LTRIM([MID])) +RTRIM(LTRIM([DID])) +RTRIM(LTRIM([SID]))  FROM [TKMOC].[dbo].[TRACEBACK] WHERE MMB001='{0}' AND MLOTNO='{1}')",MB001,LOTNO);
                sbSql.AppendFormat(" ORDER BY INVMF.MF002,MF004,MF005,MF006,MF001");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[TRACEBACK]");
                sbSql.AppendFormat(" ([MMB001],[MLOTNO],[KINDS],[LEVELS],[DATES],[MID],[DID],[SID],[MB001],[MB002],[LOTNO],[NUMS])");
                sbSql.AppendFormat(" SELECT '{0}','{1}','5調整','0',MF002,MF004,MF005,MF006,MF001,'',MF002,MF010", MB001, LOTNO);
                sbSql.AppendFormat(" FROM [TK].dbo.INVMF WITH (NOLOCK)");
                sbSql.AppendFormat(" WHERE INVMF.MF009 IN ('5')");
                sbSql.AppendFormat(" AND RTRIM(LTRIM(MF001))+RTRIM(LTRIM(MF002)) IN (SELECT RTRIM(LTRIM([MB001]))+RTRIM(LTRIM([LOTNO])) FROM [TKMOC].[dbo].[TRACEBACK] WHERE MMB001='{0}' AND MLOTNO='{1}')", MB001, LOTNO);
                sbSql.AppendFormat(" AND RTRIM(LTRIM(MF004)) +RTRIM(LTRIM(MF005)) +RTRIM(LTRIM(MF006))  NOT IN (SELECT RTRIM(LTRIM([MID])) +RTRIM(LTRIM([DID])) +RTRIM(LTRIM([SID]))  FROM [TKMOC].[dbo].[TRACEBACK] WHERE MMB001='{0}' AND MLOTNO='{1}')", MB001, LOTNO);
                sbSql.AppendFormat(" ORDER BY INVMF.MF002,MF004,MF005,MF006,MF001");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[TRACEBACK]");
                sbSql.AppendFormat(" ([MMB001],[MLOTNO],[KINDS],[LEVELS],[DATES],[MID],[DID],[SID],[MB001],[MB002],[LOTNO],[NUMS])");
                sbSql.AppendFormat(" SELECT '{0}','{1}','6其他','0',MF002,MF004,MF005,MF006,MF001,'',MF002,MF010", MB001, LOTNO);
                sbSql.AppendFormat(" FROM [TK].dbo.INVMF WITH (NOLOCK)");
                sbSql.AppendFormat(" WHERE INVMF.MF009 IN ('3') AND MF004 LIKE 'A1%'");
                sbSql.AppendFormat(" AND RTRIM(LTRIM(MF001))+RTRIM(LTRIM(MF002)) IN (SELECT RTRIM(LTRIM([MB001]))+RTRIM(LTRIM([LOTNO])) FROM [TKMOC].[dbo].[TRACEBACK] WHERE MMB001='{0}' AND MLOTNO='{1}')", MB001, LOTNO);
                sbSql.AppendFormat(" AND RTRIM(LTRIM(MF004)) +RTRIM(LTRIM(MF005)) +RTRIM(LTRIM(MF006))  NOT IN (SELECT RTRIM(LTRIM([MID])) +RTRIM(LTRIM([DID])) +RTRIM(LTRIM([SID]))  FROM [TKMOC].[dbo].[TRACEBACK] WHERE MMB001='{0}' AND MLOTNO='{1}')", MB001, LOTNO);
                sbSql.AppendFormat(" ORDER BY INVMF.MF002,MF004,MF005,MF006,MF001");
                sbSql.AppendFormat(" ");

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

        #endregion

        #region BUTTON


        private void button1_Click(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(textBox1.Text)&& !string.IsNullOrEmpty(textBox2.Text))
            {
                if(comboBox1.Text.Trim().Equals("成品逆溯"))
                {
                    SEARCHOUT(textBox1.Text.Trim(), textBox2.Text.Trim());
                }
                else if (comboBox1.Text.Trim().Equals("原料順溯"))
                {
                   
                }


            }
        }

        #endregion
    }
}
