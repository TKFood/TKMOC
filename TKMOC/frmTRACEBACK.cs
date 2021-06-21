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
using FastReport;
using FastReport.Data;

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

                sbSql.AppendFormat(@"  
                                    SELECT MF001,MF002,'0',MF003,MF004,MF005,MF006,MF010
                                    FROM [TK].dbo.INVME WITH (NOLOCK),[TK].dbo.INVMF WITH (NOLOCK)
                                    WHERE MF001=ME001 AND MF002=ME002
                                    AND MF009 IN ('2','5')
                                    AND MF001='{0}' AND MF002='{1}'
                                    ORDER BY MF002,MF003,MF004,MF005
                                    ", MB001, LOTNO);

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

        public void SEARCHOUT2(string MB001, string LOTNO)
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

                sbSql.AppendFormat(@" 
                                SELECT MF001,MF002,'1入庫','0',MF003,MF004,MF005,MF006,MF001,'',MF002,MF010
                                FROM [TK].dbo.INVME WITH (NOLOCK),[TK].dbo.INVMF WITH (NOLOCK),[TK].dbo.CMSMQ WITH (NOLOCK)
                                WHERE MF001=ME001 AND MF002=ME002
                                AND MQ001=MF004
                                AND MQ003='34'
                                AND MF001='{0}' AND MF002='{1}'
                                ORDER BY INVMF.MF002,MF003,MF004,MF005
                                    ", MB001, LOTNO);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    ADDTRACEBACKOUT2(MB001, LOTNO);
                    ADDTRACEBACKMOC2(MB001, LOTNO);
                    ADDTRACEBACKMOCOUTIN2(MB001, LOTNO);
                    ADDTRACEBACKINVMFSALE2(MB001, LOTNO);
                    ADDTRACEBACKINVMF2(MB001, LOTNO);
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

         
                sbSql.AppendFormat(@" 
                                     DELETE [TKMOC].[dbo].[TRACEBACK]            
 
                                     INSERT INTO [TKMOC].[dbo].[TRACEBACK]
                                     ([MMB001],[MLOTNO],[KINDS],[LEVELS],[DATES],[MID],[DID],[SID],[MB001],[MB002],[LOTNO],[NUMS])
                                     SELECT MF001,MF002,'1銷貨','0',MF003,MF004,MF005,MF006,MF001,'',MF002,MF010
                                     FROM [TK].dbo.INVME WITH (NOLOCK),[TK].dbo.INVMF WITH (NOLOCK)
                                     WHERE MF001=ME001 AND MF002=ME002
                                     AND MF009 IN ('2','5')
                                     AND MF001='{0}' AND MF002='{1}'
                                     ORDER BY INVMF.MF002,MF003,MF004,MF005

                                     ", MB001, LOTNO);

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

        public void ADDTRACEBACKOUT2(string MB001, string LOTNO)
        {

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@" 
                    DELETE[TKMOC].[dbo].[TRACEBACK]
                   
                    INSERT INTO [TKMOC].[dbo].[TRACEBACK]
                    ([MMB001],[MLOTNO],[KINDS],[LEVELS],[DATES],[MID],[DID],[SID],[MB001],[MB002],[LOTNO],[NUMS])
                    SELECT MF001,MF002,'1入庫','0',MF003,MF004,MF005,MF006,MF001,'',MF002,MF010
                    FROM [TK].dbo.INVME WITH (NOLOCK),[TK].dbo.INVMF WITH (NOLOCK),[TK].dbo.CMSMQ WITH (NOLOCK)
                    WHERE MF001=ME001 AND MF002=ME002
                    AND MQ001=MF004
                    AND MQ003='34'
                    AND MF001='{0}' AND MF002='{1}'
                    ORDER BY INVMF.MF002,MF003,MF004,MF005
                    ", MB001,LOTNO);


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

                sbSql.AppendFormat(@" 
                                     WITH RTABLES AS 
                                     ( SELECT 0 AS LEVELS,[TG001],[TG002],[TG003],[TG004],[TG011],[TG017],[TG014],[TG015],[TE001],[TE002],[TE003],[TE004],[TE005],[TE010] 
                                     FROM [TK].[dbo].[VMOCTGMOCTE] WITH (NOLOCK) 
                                     WHERE [VMOCTGMOCTE].TG004 ='{0}' AND [VMOCTGMOCTE].TG017 ='{1}'  
                                     UNION ALL 
                                     SELECT LEVELS+1,B.[TG001], B.[TG002], B.[TG003], B.[TG004], B.[TG011], B.[TG017], B.[TG014], B.[TG015], B.[TE001], B.[TE002],B.[TE003], B.[TE004], B.[TE005], B.[TE010] 
                                     FROM [TK].[dbo].[VMOCTGMOCTE] B WITH (NOLOCK) 
                                     INNER JOIN RTABLES ON RTABLES.[TE004]=B.[TG004] AND RTABLES.[TE010]=B.[TG017] )   
 
 
                                    INSERT INTO [TKMOC].[dbo].[TRACEBACK] 
                                    ([MMB001],[MLOTNO],[KINDS],[LEVELS],[DATES],[MID],[DID],[SID],[MB001],[MB002],[LOTNO],[NUMS],[TG014],[TG015]) 
 
                                     SELECT '{0}','{1}','2生產',LEVELS  
                                     ,(SELECT TOP 1 TF003 FROM [TK].dbo.MOCTF WHERE TF001=TG001 AND TF002=TG002 ORDER BY TF003) 
                                     ,[TG001],[TG002],[TG003],[TG004], '',[TG017],[TG011]  ,[TG014],[TG015]
                                     FROM RTABLES 
                                     GROUP BY LEVELS,[TG001],[TG002],[TG003],[TG004],[TG017],[TG011] ,[TG014],[TG015]
                                     ORDER BY LEVELS,[TG004] 

                                    ", MB001, LOTNO);

                //sbSql.AppendFormat(" WITH RTABLES");
                //sbSql.AppendFormat(" AS (");
                //sbSql.AppendFormat(" SELECT 0 AS LEVELS,[TG001],[TG002],[TG003],[TG004],[TG011],[TG017],[TG014],[TG015],[TE001],[TE002],[TE003],[TE004],[TE005],[TE010]");
                //sbSql.AppendFormat(" FROM [TK].[dbo].[VMOCTGMOCTE] WITH (NOLOCK)");
                //sbSql.AppendFormat(" WHERE [VMOCTGMOCTE].TG004 ='{0}' AND [VMOCTGMOCTE].TG017 ='{1}' ", MB001, LOTNO);
                //sbSql.AppendFormat(" UNION ALL");
                //sbSql.AppendFormat(" SELECT LEVELS+1,B.[TG001], B.[TG002], B.[TG003], B.[TG004], B.[TG011], B.[TG017], B.[TG014], B.[TG015], B.[TE001], B.[TE002],B.[TE003], B.[TE004], B.[TE005], B.[TE010]");
                //sbSql.AppendFormat(" FROM [TK].[dbo].[VMOCTGMOCTE] B WITH (NOLOCK)");
                //sbSql.AppendFormat(" INNER JOIN RTABLES ON RTABLES.[TE004]=B.[TG004] AND RTABLES.[TE010]=B.[TG017]");
                //sbSql.AppendFormat(" ) ");
                //sbSql.AppendFormat(" ");
                //sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[TRACEBACK]");
                //sbSql.AppendFormat(" ([MMB001],[MLOTNO],[KINDS],[LEVELS],[DATES],[MID],[DID],[SID],[MB001],[MB002],[LOTNO],[NUMS])");
                //sbSql.AppendFormat(" ");
                //sbSql.AppendFormat(" SELECT '{0}','{1}','2生產',LEVELS ",MB001,LOTNO);
                //sbSql.AppendFormat(" ,(SELECT TOP 1 TF003 FROM [TK].dbo.MOCTF WHERE TF001=TG001 AND TF002=TG002 ORDER BY TF003)");
                //sbSql.AppendFormat(" ,[TG001],[TG002],[TG003],[TG004], '',[TG017],[TG011]");
                //sbSql.AppendFormat("  FROM RTABLES");
                //sbSql.AppendFormat(" GROUP BY LEVELS,[TG001],[TG002],[TG003],[TG004],[TG017],[TG011]");
                //sbSql.AppendFormat(" ORDER BY LEVELS,[TG004]");
                //sbSql.AppendFormat(" ");


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

        public void ADDTRACEBACKMOC2(string MB001, string LOTNO)
        {

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@"                                     
                                INSERT INTO [TKMOC].[dbo].[TRACEBACK]
                                ([MMB001],[MLOTNO],[KINDS],[LEVELS],[DATES],[MID],[DID],[SID],[TG014],[TG015],[MB001],[MB002],[LOTNO],[NUMS])

                                SELECT DISTINCT  MF001,MF002,'2領退料','0',MF003,MF004,MF005,MF006,TG014,TG015,TG004,'',TG017,TG011
                                FROM [TK].dbo.INVME WITH (NOLOCK),[TK].dbo.INVMF WITH (NOLOCK),[TK].dbo.CMSMQ WITH (NOLOCK),[TK].dbo.MOCTE WITH (NOLOCK),[TK].dbo.MOCTG WITH (NOLOCK)
                                WHERE MF001=ME001 AND MF002=ME002
                                AND MQ001=MF004
                                AND TE001=MF004 AND TE002=MF005 AND TE004=MF001 AND TE010=MF002
                                AND TG014=TE011 AND TG015=TE012
                                AND MQ003 IN ('54','56')
                                AND MF001='{0}' AND MF002='{1}'
                                UNION ALL
                                SELECT DISTINCT  MF001,MF002,'2領退料','0',MF003,MF004,MF005,MF006,TG014,TG015,TG004,'',TG017,TG011
                                FROM [TK].dbo.INVME WITH (NOLOCK),[TK].dbo.INVMF WITH (NOLOCK),[TK].dbo.CMSMQ WITH (NOLOCK),[TK].dbo.MOCTE WITH (NOLOCK),[TK].dbo.MOCTG WITH (NOLOCK)
                                WHERE MF001=ME001 AND MF002=ME002
                                AND MQ001=MF004
                                AND TE001=MF004 AND TE002=MF005 AND TE004=MF001 AND TE010=MF002
                                AND TG014=TE011 AND TG015=TE012
                                AND MQ003 IN ('54','56')
                                AND LTRIM(RTRIM(MF001))+LTRIM(RTRIM(MF002)) 
                                IN (
                                SELECT LTRIM(RTRIM(TG004))+LTRIM(RTRIM(TG017))
                                FROM [TK].dbo.INVME WITH (NOLOCK),[TK].dbo.INVMF WITH (NOLOCK),[TK].dbo.CMSMQ WITH (NOLOCK),[TK].dbo.MOCTE WITH (NOLOCK),[TK].dbo.MOCTG WITH (NOLOCK)
                                WHERE MF001=ME001 AND MF002=ME002
                                AND MQ001=MF004
                                AND TE001=MF004 AND TE002=MF005 AND TE004=MF001 AND TE010=MF002
                                AND TG014=TE011 AND TG015=TE012
                                AND MQ003 IN ('54','56')
                                AND MF001='{0}' AND MF002='{1}'
                                GROUP BY LTRIM(RTRIM(TG004))+LTRIM(RTRIM(TG017))
                                )
                                UNION ALL
                                SELECT DISTINCT  MF001,MF002,'2領退料','0',MF003,MF004,MF005,MF006,TG014,TG015,TG004,'',TG017,TG011
                                FROM [TK].dbo.INVME WITH (NOLOCK),[TK].dbo.INVMF WITH (NOLOCK),[TK].dbo.CMSMQ WITH (NOLOCK),[TK].dbo.MOCTE WITH (NOLOCK),[TK].dbo.MOCTG WITH (NOLOCK)
                                WHERE MF001=ME001 AND MF002=ME002
                                AND MQ001=MF004
                                AND TE001=MF004 AND TE002=MF005 AND TE004=MF001 AND TE010=MF002
                                AND TG014=TE011 AND TG015=TE012
                                AND MQ003 IN ('54','56')
                                AND LTRIM(RTRIM(MF001))+LTRIM(RTRIM(MF002)) 
                                IN (
                                SELECT LTRIM(RTRIM(TG004))+LTRIM(RTRIM(TG017))
                                FROM [TK].dbo.INVME WITH (NOLOCK),[TK].dbo.INVMF WITH (NOLOCK),[TK].dbo.CMSMQ WITH (NOLOCK),[TK].dbo.MOCTE WITH (NOLOCK),[TK].dbo.MOCTG WITH (NOLOCK)
                                WHERE MF001=ME001 AND MF002=ME002
                                AND MQ001=MF004
                                AND TE001=MF004 AND TE002=MF005 AND TE004=MF001 AND TE010=MF002
                                AND TG014=TE011 AND TG015=TE012
                                AND MQ003 IN ('54','56')
                                AND LTRIM(RTRIM(MF001))+LTRIM(RTRIM(MF002)) 
                                IN (
                                SELECT LTRIM(RTRIM(TG004))+LTRIM(RTRIM(TG017))
                                FROM [TK].dbo.INVME WITH (NOLOCK),[TK].dbo.INVMF WITH (NOLOCK),[TK].dbo.CMSMQ WITH (NOLOCK),[TK].dbo.MOCTE WITH (NOLOCK),[TK].dbo.MOCTG WITH (NOLOCK)
                                WHERE MF001=ME001 AND MF002=ME002
                                AND MQ001=MF004
                                AND TE001=MF004 AND TE002=MF005 AND TE004=MF001 AND TE010=MF002
                                AND TG014=TE011 AND TG015=TE012
                                AND MQ003 IN ('54','56')
                                AND MF001='{0}' AND MF002='{1}'
                                GROUP BY LTRIM(RTRIM(TG004))+LTRIM(RTRIM(TG017))
                                )
                                GROUP BY LTRIM(RTRIM(TG004))+LTRIM(RTRIM(TG017))
                                )
                                UNION ALL
                                SELECT DISTINCT  MF001,MF002,'2領退料','0',MF003,MF004,MF005,MF006,TG014,TG015,TG004,'',TG017,TG011
                                FROM [TK].dbo.INVME WITH (NOLOCK),[TK].dbo.INVMF WITH (NOLOCK),[TK].dbo.CMSMQ WITH (NOLOCK),[TK].dbo.MOCTE WITH (NOLOCK),[TK].dbo.MOCTG WITH (NOLOCK)
                                WHERE MF001=ME001 AND MF002=ME002
                                AND MQ001=MF004
                                AND TE001=MF004 AND TE002=MF005 AND TE004=MF001 AND TE010=MF002
                                AND TG014=TE011 AND TG015=TE012
                                AND MQ003 IN ('54','56')
                                AND LTRIM(RTRIM(MF001))+LTRIM(RTRIM(MF002)) 
                                IN (
                                SELECT LTRIM(RTRIM(TG004))+LTRIM(RTRIM(TG017))
                                FROM [TK].dbo.INVME WITH (NOLOCK),[TK].dbo.INVMF WITH (NOLOCK),[TK].dbo.CMSMQ WITH (NOLOCK),[TK].dbo.MOCTE WITH (NOLOCK),[TK].dbo.MOCTG WITH (NOLOCK)
                                WHERE MF001=ME001 AND MF002=ME002
                                AND MQ001=MF004
                                AND TE001=MF004 AND TE002=MF005 AND TE004=MF001 AND TE010=MF002
                                AND TG014=TE011 AND TG015=TE012
                                AND MQ003 IN ('54','56')
                                AND LTRIM(RTRIM(MF001))+LTRIM(RTRIM(MF002)) 
                                IN (
                                SELECT LTRIM(RTRIM(TG004))+LTRIM(RTRIM(TG017))
                                FROM [TK].dbo.INVME WITH (NOLOCK),[TK].dbo.INVMF WITH (NOLOCK),[TK].dbo.CMSMQ WITH (NOLOCK),[TK].dbo.MOCTE WITH (NOLOCK),[TK].dbo.MOCTG WITH (NOLOCK)
                                WHERE MF001=ME001 AND MF002=ME002
                                AND MQ001=MF004
                                AND TE001=MF004 AND TE002=MF005 AND TE004=MF001 AND TE010=MF002
                                AND TG014=TE011 AND TG015=TE012
                                AND MQ003 IN ('54','56')
                                AND LTRIM(RTRIM(MF001))+LTRIM(RTRIM(MF002)) 
                                IN (
                                SELECT LTRIM(RTRIM(TG004))+LTRIM(RTRIM(TG017))
                                FROM [TK].dbo.INVME WITH (NOLOCK),[TK].dbo.INVMF WITH (NOLOCK),[TK].dbo.CMSMQ WITH (NOLOCK),[TK].dbo.MOCTE WITH (NOLOCK),[TK].dbo.MOCTG WITH (NOLOCK)
                                WHERE MF001=ME001 AND MF002=ME002
                                AND MQ001=MF004
                                AND TE001=MF004 AND TE002=MF005 AND TE004=MF001 AND TE010=MF002
                                AND TG014=TE011 AND TG015=TE012
                                AND MQ003 IN ('54','56')
                                AND MF001='{0}' AND MF002='{1}'
                                GROUP BY LTRIM(RTRIM(TG004))+LTRIM(RTRIM(TG017))
                                )
                                GROUP BY LTRIM(RTRIM(TG004))+LTRIM(RTRIM(TG017))
                                )
                                GROUP BY LTRIM(RTRIM(TG004))+LTRIM(RTRIM(TG017))
                                )
                                UNION ALL
                                SELECT DISTINCT  MF001,MF002,'2領退料','0',MF003,MF004,MF005,MF006,TG014,TG015,TG004,'',TG017,TG011
                                FROM [TK].dbo.INVME WITH (NOLOCK),[TK].dbo.INVMF WITH (NOLOCK),[TK].dbo.CMSMQ WITH (NOLOCK),[TK].dbo.MOCTE WITH (NOLOCK),[TK].dbo.MOCTG WITH (NOLOCK)
                                WHERE MF001=ME001 AND MF002=ME002
                                AND MQ001=MF004
                                AND TE001=MF004 AND TE002=MF005 AND TE004=MF001 AND TE010=MF002
                                AND TG014=TE011 AND TG015=TE012
                                AND MQ003 IN ('54','56')
                                AND LTRIM(RTRIM(MF001))+LTRIM(RTRIM(MF002)) 
                                IN (
                                SELECT LTRIM(RTRIM(TG004))+LTRIM(RTRIM(TG017))
                                FROM [TK].dbo.INVME WITH (NOLOCK),[TK].dbo.INVMF WITH (NOLOCK),[TK].dbo.CMSMQ WITH (NOLOCK),[TK].dbo.MOCTE WITH (NOLOCK),[TK].dbo.MOCTG WITH (NOLOCK)
                                WHERE MF001=ME001 AND MF002=ME002
                                AND MQ001=MF004
                                AND TE001=MF004 AND TE002=MF005 AND TE004=MF001 AND TE010=MF002
                                AND TG014=TE011 AND TG015=TE012
                                AND MQ003 IN ('54','56')
                                AND LTRIM(RTRIM(MF001))+LTRIM(RTRIM(MF002)) 
                                IN (
                                SELECT LTRIM(RTRIM(TG004))+LTRIM(RTRIM(TG017))
                                FROM [TK].dbo.INVME WITH (NOLOCK),[TK].dbo.INVMF WITH (NOLOCK),[TK].dbo.CMSMQ WITH (NOLOCK),[TK].dbo.MOCTE WITH (NOLOCK),[TK].dbo.MOCTG WITH (NOLOCK)
                                WHERE MF001=ME001 AND MF002=ME002
                                AND MQ001=MF004
                                AND TE001=MF004 AND TE002=MF005 AND TE004=MF001 AND TE010=MF002
                                AND TG014=TE011 AND TG015=TE012
                                AND MQ003 IN ('54','56')
                                AND LTRIM(RTRIM(MF001))+LTRIM(RTRIM(MF002)) 
                                IN (
                                SELECT LTRIM(RTRIM(TG004))+LTRIM(RTRIM(TG017))
                                FROM [TK].dbo.INVME WITH (NOLOCK),[TK].dbo.INVMF WITH (NOLOCK),[TK].dbo.CMSMQ WITH (NOLOCK),[TK].dbo.MOCTE WITH (NOLOCK),[TK].dbo.MOCTG WITH (NOLOCK)
                                WHERE MF001=ME001 AND MF002=ME002
                                AND MQ001=MF004
                                AND TE001=MF004 AND TE002=MF005 AND TE004=MF001 AND TE010=MF002
                                AND TG014=TE011 AND TG015=TE012
                                AND MQ003 IN ('54','56')
                                AND LTRIM(RTRIM(MF001))+LTRIM(RTRIM(MF002)) 
                                IN (
                                SELECT LTRIM(RTRIM(TG004))+LTRIM(RTRIM(TG017))
                                FROM [TK].dbo.INVME WITH (NOLOCK),[TK].dbo.INVMF WITH (NOLOCK),[TK].dbo.CMSMQ WITH (NOLOCK),[TK].dbo.MOCTE WITH (NOLOCK),[TK].dbo.MOCTG WITH (NOLOCK)
                                WHERE MF001=ME001 AND MF002=ME002
                                AND MQ001=MF004
                                AND TE001=MF004 AND TE002=MF005 AND TE004=MF001 AND TE010=MF002
                                AND TG014=TE011 AND TG015=TE012
                                AND MQ003 IN ('54','56')
                                AND MF001='{0}' AND MF002='{1}'
                                GROUP BY LTRIM(RTRIM(TG004))+LTRIM(RTRIM(TG017))
                                )
                                GROUP BY LTRIM(RTRIM(TG004))+LTRIM(RTRIM(TG017))
                                )
                                GROUP BY LTRIM(RTRIM(TG004))+LTRIM(RTRIM(TG017))
                                )
                                GROUP BY LTRIM(RTRIM(TG004))+LTRIM(RTRIM(TG017))
                                )
                                ORDER BY INVMF.MF001,INVMF.MF002,MF004,MF005
                              
                                    ", MB001,LOTNO);

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

        public void ADDTRACEBACKMOCOUTIN2(string MB001, string LOTNO)
        {

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@" 
                                    INSERT INTO [TKMOC].[dbo].[TRACEBACK]
                                    ([MMB001],[MLOTNO],[KINDS],[LEVELS],[DATES],[MID],[DID],[SID],[MB001],[MB002],[LOTNO],[NUMS],TG014,TG015)

                                    SELECT MF001,MF002,'3生產入庫','0',MF003,MF004,MF005,MF006,MF001,'',MF002,MF010,TG014,TG015
                                    FROM [TK].dbo.INVME WITH (NOLOCK),[TK].dbo.INVMF WITH (NOLOCK),[TK].dbo.CMSMQ WITH (NOLOCK),[TK].dbo.MOCTG WITH (NOLOCK)
                                    WHERE MF001=ME001 AND MF002=ME002
                                    AND MQ001=MF004
                                    AND TG001=MF004 AND TG002=MF005
                                    AND MQ003='58'
                                    AND RTRIM(LTRIM(MF001))+RTRIM(LTRIM(MF002)) IN
                                    (
                                    SELECT RTRIM(LTRIM([MB001]))+RTRIM(LTRIM([LOTNO]))
                                    FROM [TKMOC].[dbo].[TRACEBACK]
                                    )
                                    ORDER BY INVMF.MF002,MF003,MF004,MF005

                                    ", MB001,LOTNO);

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
                sbSql.AppendFormat(" SELECT '{0}','{1}','4入庫','0',MF003,MF004,MF005,MF006,MF001,'',MF002,MF010", MB001, LOTNO);
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

        public void ADDTRACEBACKINVMFSALE2(string MB001, string LOTNO)
        {

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@" INSERT INTO [TKMOC].[dbo].[TRACEBACK]
                                    ([MMB001],[MLOTNO],[KINDS],[LEVELS],[DATES],[MID],[DID],[SID],[MB001],[MB002],[LOTNO],[NUMS])

                                    SELECT MF001,MF002,'5銷貨','0',MF003,MF004,MF005,MF006,MF001,'',MF002,MF010
                                    FROM [TK].dbo.INVME WITH (NOLOCK),[TK].dbo.INVMF WITH (NOLOCK),[TK].dbo.CMSMQ WITH (NOLOCK)
                                    WHERE MF001=ME001 AND MF002=ME002
                                    AND MQ001=MF004
                                    AND MF009 IN ('2','5')
                                    AND RTRIM(LTRIM(MF001))+RTRIM(LTRIM(MF002)) IN
                                    (
                                    SELECT RTRIM(LTRIM([MB001]))+RTRIM(LTRIM([LOTNO]))
                                    FROM [TKMOC].[dbo].[TRACEBACK]
                                    )
                                    ORDER BY INVMF.MF002,MF003,MF004,MF005

                                ");

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

        public void ADDTRACEBACKINVMF2(string MB001, string LOTNO)
        {

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@" INSERT INTO [TKMOC].[dbo].[TRACEBACK]
                                    ([MMB001],[MLOTNO],[KINDS],[LEVELS],[DATES],[MID],[DID],[SID],[MB001],[MB002],[LOTNO],[NUMS])

                                    SELECT MF001,MF002,'6其他','0',MF003,MF004,MF005,MF006,MF001,'',MF002,MF010
                                    FROM [TK].dbo.INVME WITH (NOLOCK),[TK].dbo.INVMF WITH (NOLOCK),[TK].dbo.CMSMQ WITH (NOLOCK)
                                    WHERE MF001=ME001 AND MF002=ME002
                                    AND MQ001=MF004
                                    AND MQ003 IN ('13','14','15','16','17')
                                    AND RTRIM(LTRIM(MF001))+RTRIM(LTRIM(MF002)) IN
                                    (
                                    SELECT RTRIM(LTRIM([MB001]))+RTRIM(LTRIM([LOTNO]))
                                    FROM [TKMOC].[dbo].[TRACEBACK]
                                    )
                                    ORDER BY INVMF.MF002,MF003,MF004,MF005

                                ");

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

        public void UPDATETRACEBACK()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@" 
                                    UPDATE [TKMOC].[dbo].[TRACEBACK]
                                    SET [MMB002]=INVMB.MB002
                                    FROM [TK].dbo.INVMB
                                    WHERE [MMB001]=INVMB.MB001

                                    UPDATE [TKMOC].[dbo].[TRACEBACK]
                                    SET [MB002]=INVMB.MB002
                                    FROM [TK].dbo.INVMB
                                    WHERE [TRACEBACK].[MB001]=INVMB.MB001

                                ");

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

        public void SETFASTREPORT()
        {
            StringBuilder SQL = new StringBuilder();
            Report report1 = new Report();

            SQL = SETSQL();

            report1.Load(@"REPORT\追踨表.frx");

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL.ToString();

            report1.Preview = previewControl2;
            report1.Show();
        }

        public StringBuilder SETSQL()
        {
            StringBuilder SB = new StringBuilder();
         
            SB.AppendFormat(@" 
                            SELECT [MMB001] AS '主品號',[MMB002] AS '主品名',[MLOTNO] AS '主批號',[KINDS] AS '類別',[LEVELS] AS '層別',[DATES] AS '日期',[MID] AS '單別',[DID] AS '單號',[SID] AS '序號',[TG014] AS '製令',[TG015] AS '製令號',[MB001] AS '品號',[MB002] AS '品名',[LOTNO] AS '批號',[NUMS] AS '數量'
                            FROM [TKMOC].[dbo].[TRACEBACK]
                            ORDER BY [KINDS],[MMB001],[MLOTNO],[MID],[DID],[SID],[TG014],[TG015]


                            ");

            return SB;

        }

        private void frmTRACEBACK_Load(object sender, EventArgs e)
        {
            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;      //奇數列顏色
            dataGridView2.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;


            //先建立個 CheckBox 欄
            DataGridViewCheckBoxColumn cbCol = new DataGridViewCheckBoxColumn();
            cbCol.Width = 40;   //設定寬度
            cbCol.HeaderText = "　全選";
            cbCol.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;   //置中
            cbCol.TrueValue = true;
            cbCol.FalseValue = false;
            dataGridView1.Columns.Insert(0, cbCol);

            #region 建立全选 CheckBox

            //建立个矩形，等下计算 CheckBox 嵌入 GridView 的位置
            Rectangle rect = dataGridView1.GetCellDisplayRectangle(0, -1, true);
            rect.X = rect.Location.X + rect.Width / 4 - 18;
            rect.Y = rect.Location.Y + (rect.Height / 2 - 9);

            CheckBox cbHeader = new CheckBox();
            cbHeader.Name = "checkboxHeader";
            cbHeader.Size = new Size(18, 18);
            cbHeader.Location = rect.Location;

            //全选要设定的事件
            cbHeader.CheckedChanged += new EventHandler(cbHeader_CheckedChanged);

            //将 CheckBox 加入到 dataGridView1
            dataGridView1.Controls.Add(cbHeader);
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);

            #endregion

            //先建立個 CheckBox 欄
            cbCol = new DataGridViewCheckBoxColumn();
            cbCol.Width = 40;   //設定寬度
            cbCol.HeaderText = "　全選";
            cbCol.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;   //置中
            cbCol.TrueValue = true;
            cbCol.FalseValue = false;
            dataGridView2.Columns.Insert(0, cbCol);

            #region 建立全选 CheckBox

            //建立个矩形，等下计算 CheckBox 嵌入 GridView 的位置
            rect = dataGridView2.GetCellDisplayRectangle(0, -1, true);
            rect.X = rect.Location.X + rect.Width / 4 - 18;
            rect.Y = rect.Location.Y + (rect.Height / 2 - 9);

            cbHeader = new CheckBox();
            cbHeader.Name = "checkboxHeader";
            cbHeader.Size = new Size(18, 18);
            cbHeader.Location = rect.Location;
            //将 CheckBox 加入到 dataGridView2
            dataGridView2.Controls.Add(cbHeader);
            dataGridView2.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);


            #endregion
        }
        private void cbHeader_CheckedChanged(object sender, EventArgs e)
        {
            dataGridView1.EndEdit();

            foreach (DataGridViewRow dr in dataGridView1.Rows)
            {
                dr.Cells[0].Value = ((CheckBox)dataGridView1.Controls.Find("checkboxHeader", true)[0]).Checked;

            }

        }

        public void SEARCHTRACEBACK1(string STATUS)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);
                SqlDataAdapter adapter = new SqlDataAdapter();
                SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
                DataSet ds = new DataSet();

                sbSql.Clear();
               

                sbSql.AppendFormat(@"  
                                    SELECT [MID] AS '銷貨單別',[DID] AS '銷貨單號',[SID] AS '銷貨序號',[MMB001] AS '主品號',[MMB002] AS '主品名',[MLOTNO] AS '主批號',[KINDS] AS '類別',[LEVELS] AS '層別',[DATES] AS '日期',[TG014] AS '製令',[TG015] AS '製令號',[MB001] AS '品號',[MB002] AS '品名',[LOTNO] AS '批號',[NUMS] AS '數量'
                                    FROM [TKMOC].[dbo].[TRACEBACK]
                                    WHERE [KINDS] IN ('1銷貨')
                                    ORDER BY [KINDS],[MMB001],[MLOTNO],[MID],[DID],[SID],[TG014],[TG015]
                                    ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView1.DataSource = ds.Tables["ds"];
                        dataGridView1.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                        dataGridView1.AutoResizeColumns();
                        dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);
                       
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
        
        public void SETFASTREPORT1(string STATUS)
        {
            StringBuilder SQL = new StringBuilder();
            string SELECT = SELECT1();
            Report report1 = new Report();

            if(!string.IsNullOrEmpty(SELECT))
            {
                SQL.AppendFormat(@"  
                                    SELECT 
                                    CONVERT(NVARCHAR,CONVERT(datetime,TG003),111)  AS '銷貨日期'
                                    ,TG001+'-'+TG002 AS '銷貨單號'
                                    ,CONVERT(NVARCHAR,CONVERT(datetime,TG042),111)   AS '單據日期'
                                    ,TG004 AS '客戶代號'
                                    ,TG007 AS '客戶簡稱'
                                    ,TG033 AS '總數量'
                                    ,TG020 AS '單頭備註'
                                    ,TH003 AS '序號'
                                    ,TH004 AS '品號'
                                    ,TH005 AS '品名'
                                    ,TH006 AS '規格'
                                    ,TH007 AS '庫別代號'
                                    ,MC002 AS '庫別名稱'
                                    ,CONVERT(NVARCHAR,CONVERT(datetime,TH106),111)  AS '有效日期'
                                    ,TH008 AS '銷貨數量'
                                    ,TH009 AS '單位'
                                    ,TH014+'-'+TH015+'-'+TH016 AS '訂單單號'
                                    ,TH017 AS '批號'
                                    ,TH018 AS '單身備註'
                                    FROM [TK].dbo.COPTG,[TK].dbo.COPTH,[TK].dbo.CMSMC
                                    WHERE TG001=TH001 AND TG002=TH002
                                    AND MC001=TH007
                                    AND TH001+TH002+TH003 IN ({0})
                                    ", SELECT.ToString());

                report1.Load(@"REPORT\銷貨單明細表.frx");

                report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;
                Table.SelectCommand = SQL.ToString();

                report1.Preview = previewControl3;
                report1.Show();
            }

           
        }

        public string SELECT1()
        {
            StringBuilder ADDSQL = new StringBuilder();

            foreach (DataGridViewRow dgR in this.dataGridView1.Rows)
            {
                try
                {
                    DataGridViewCheckBoxCell cbx = (DataGridViewCheckBoxCell)dgR.Cells[0];
                    if ((bool)cbx.FormattedValue)
                    {
                        ADDSQL.AppendFormat(@" '{0}', ", dgR.Cells["銷貨單別"].Value.ToString().Trim()+ dgR.Cells["銷貨單號"].Value.ToString().Trim()+ dgR.Cells["銷貨序號"].Value.ToString().Trim());
                    }
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(ex.Message);
                }
            }

            ADDSQL.AppendFormat(@" '' ");

            return ADDSQL.ToString();

        }

        public void SEARCHTRACEBACK2(string STATUS)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);
                SqlDataAdapter adapter = new SqlDataAdapter();
                SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
                DataSet ds = new DataSet();

                sbSql.Clear();


                sbSql.AppendFormat(@"  
                                    SELECT [TG014] AS '製令',[TG015] AS '製令單號',[MID] AS '入庫單別',[DID] AS '入庫單號',[SID] AS '入庫序號',[MMB001] AS '主品號',[MMB002] AS '主品名',[MLOTNO] AS '主批號',[KINDS] AS '類別',[LEVELS] AS '層別',[DATES] AS '日期',[MB001] AS '品號',[MB002] AS '品名',[LOTNO] AS '批號',[NUMS] AS '數量'
                                    FROM [TKMOC].[dbo].[TRACEBACK]
                                    WHERE [KINDS] IN ('2生產')
                                    ORDER BY [KINDS],[MMB001],[MLOTNO],[MID],[DID],[SID],[TG014],[TG015]
                                    ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView2.DataSource = ds.Tables["ds"];
                        dataGridView2.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                        dataGridView2.AutoResizeColumns();
                        dataGridView2.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);

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

        public void SETFASTREPORT2(string STATUS)
        {
            StringBuilder SQL = new StringBuilder();
            string SELECT = SELECT2();
            Report report1 = new Report();

            if (!string.IsNullOrEmpty(SELECT))
            {
                SQL.AppendFormat(@"  
                                    SELECT 
                                    CONVERT(NVARCHAR,CONVERT(datetime,TF003),111)  AS '入庫日期'
                                    ,TF001+'-'+TF002 AS '單別-單號'
                                    ,CONVERT(NVARCHAR,CONVERT(datetime,TF012),111)  AS '單據日期'
                                    ,TG004 AS '品號'
                                    ,TG005 AS '品名'
                                    ,TG006 AS '規格'
                                    ,TG011 AS '入庫數量'
                                    ,TG007 AS '單位'
                                    ,TG014+'-'+TG015 AS '製令編號'
                                    ,TG017 AS '批號'
                                    ,TA026+'-'+TA027 AS '訂單單號'
                                    ,TG020 AS '備註'
                                    FROM [TK].dbo.MOCTF, [TK].dbo.MOCTG
                                    LEFT JOIN [TK].dbo.MOCTA ON TA001=TG014 AND TA002=TG015
                                    WHERE TF001=TG001 AND TF002=TG002
                                    AND TG001+TG002+TG003 IN ({0})
                                    ORDER BY TF001,TF002,TG004
                                    ", SELECT.ToString());

                report1.Load(@"REPORT\生產入庫單明細表.frx");

                report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;
                Table.SelectCommand = SQL.ToString();

                report1.Preview = previewControl4;
                report1.Show();
            }


        }

        public string SELECT2()
        {
            StringBuilder ADDSQL = new StringBuilder();

            foreach (DataGridViewRow dgR in this.dataGridView2.Rows)
            {
                try
                {
                    DataGridViewCheckBoxCell cbx = (DataGridViewCheckBoxCell)dgR.Cells[0];
                    if ((bool)cbx.FormattedValue)
                    {
                        ADDSQL.AppendFormat(@" '{0}', ", dgR.Cells["入庫單別"].Value.ToString().Trim() + dgR.Cells["入庫單號"].Value.ToString().Trim() + dgR.Cells["入庫序號"].Value.ToString().Trim());
                    }
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(ex.Message);
                }
            }

            ADDSQL.AppendFormat(@" '' ");

            return ADDSQL.ToString();

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
                    UPDATETRACEBACK();
                }
                else if (comboBox1.Text.Trim().Equals("原料順溯"))
                {
                    SEARCHOUT2(textBox1.Text.Trim(), textBox2.Text.Trim());
                    UPDATETRACEBACK();
                }


            }

            SETFASTREPORT();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            SEARCHTRACEBACK1("1銷貨");
        }
        private void button3_Click(object sender, EventArgs e)
        {
            SETFASTREPORT1("1銷貨");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            SEARCHTRACEBACK2("2生產");
        }
        private void button5_Click(object sender, EventArgs e)
        {
            SETFASTREPORT2("2生產");
        }
        #endregion


    }
}
