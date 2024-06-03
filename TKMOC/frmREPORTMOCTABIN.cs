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
using System.Xml;
using TKITDLL;

namespace TKMOC
{
    public partial class frmREPORTMOCTABIN : Form
    {
        string DBNAME = "UOF";


        private ComponentResourceManager _ResourceManager = new ComponentResourceManager();
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        int result;
        int rownum = 0;
        DataSet ds1 = new DataSet();

        string STATUS = "";

        public Report report1 { get; private set; }

        public frmREPORTMOCTABIN()
        {
            InitializeComponent();
        }

        #region FUNCTION
        private void frmREPORTMOCTABIN_Load(object sender, EventArgs e)
        {
            SETCODE();
        }

        public void SETCODE()
        {
            DataSet ds = new DataSet();
            string yyyyMMdd = dateTimePicker1.Value.ToString("yyyyMMdd");
            string MM = Convert.ToUInt32(yyyyMMdd.Substring(4, 2)).ToString();//除0開頭
            string d1 = yyyyMMdd.Substring(6, 1);
            string d2 = yyyyMMdd.Substring(7, 1);
            string CODE = "";
            string CODE1 = "";
            string CODE2 = "";
            string CODE3 = "";
            string VDATES = "";

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
                                    SELECT [ID],[CODE] FROM [TKMOC].[dbo].[MOCCODE]
                                    ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);


                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    foreach (DataRow od in ds.Tables["ds"].Rows)
                    {

                        if (MM.Equals(od["ID"].ToString()))
                        {
                            CODE1 = od["CODE"].ToString();
                        }
                        if (d1.Equals(od["ID"].ToString()))
                        {
                            CODE2 = od["CODE"].ToString();
                        }
                        if (d2.Equals(od["ID"].ToString()))
                        {
                            CODE3 = od["CODE"].ToString();
                        }
                    }

                    textBox3.Text = CODE1 + CODE2 + CODE3;



                }
            }


            catch
            {

            }
            finally
            {

            }
        }

        public void ADDDELETEMOCLOTNO(string MOCDATES, string LOTNO, string VDATES)
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
                                    DELETE [TKMOC].[dbo].[MOCLOTNO] WHERE [MOCDATES]='{0}'

                                    INSERT INTO  [TKMOC].[dbo].[MOCLOTNO] ( [MOCDATES],[LOTNO],[VDATES])
                                    VALUES ('{0}','{1}','{2}')
                                    ", MOCDATES, LOTNO, VDATES);



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

        public string SEARCHMOCLOTNO(string MOCDATES)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

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
                                    SELECT  [MOCDATES],[LOTNO]
                                    FROM [TKMOC].[dbo].[MOCLOTNO]
                                    WHERE [MOCDATES]='{0}'
                                    ", MOCDATES);


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();

                if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["TEMPds1"].Rows[0]["LOTNO"].ToString().Trim();
                }
                else
                {
                    return null;

                }

            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }
        }


        public void ADDREPORTMOCMANULINE(string LOTNO, string TA003)
        {
            sbSql.Clear();

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


                sbSql.AppendFormat(@" 
                                    INSERT INTO [TKMOC].[dbo].[REPORTMOCMANULINE]
                                    ([ID],[MANULINE],[LOTNO],[TA001],[TA002],[TA003],[TA006],[TA007],[TA015],[TA017],[MB002],[MB003],[PCTS],[SEQ],[ALLERGEN],[COOKIES],[BARS],[BOXS],[VDATES],[COMMENT],[ORI])

                                    SELECT  NEWID(),TA021,'{0}',TA001,TA002,TA003,TA006,TA007,TA015,TA017,TA034,TA035,PCT,UDF01,ALLERGEN,SPEC,BARS,BOXS,DATES,TA029,ORI
                                    FROM 
                                    (
                                     SELECT TA021,TA001,TA002,TA003,TA006,TA007,TA015,TA017,TA034,TA035,[ERPINVMB].[PCT],MOCTA.UDF01,[ERPINVMB].[ALLERGEN] ,[ERPINVMB].[SPEC] 
                                    ,(CONVERT(decimal(16,3),TA015/ISNULL(MC004,1))) AS BARS
                                    ,(CASE WHEN ISNULL(1/[MOCHALFPRODUCTDBOXS].NUMS*[MOCHALFPRODUCTDBOXS].BOXS,0)>0  THEN CONVERT(decimal(16, 3), TA015 / ISNULL(MD007, 1) * ISNULL(1, 1))*(1/[MOCHALFPRODUCTDBOXS].NUMS*[MOCHALFPRODUCTDBOXS].BOXS) ELSE CONVERT(decimal(16, 3), TA015 / ISNULL(MD007, 1) * ISNULL(1, 1)) END ) AS BOXS
                                    ,(CASE WHEN MB198='2' THEN  CONVERT(NVARCHAR,DATEADD(MONTH,MB023,DATEADD(DAY,-1,TA003)),112) ELSE CONVERT(NVARCHAR,DATEADD(DAY,MB023,DATEADD(DAY,-1,TA003)),112) END) AS DATES
                                    ,TA029
                                    ,[ERPINVMB].[ORI]

                                    FROM [TK].dbo.MOCTA
                                    LEFT JOIN [TK].dbo.INVMB ON MB001=TA006
                                    LEFT JOIN [TKMOC].[dbo].[ERPINVMB] ON [ERPINVMB].MB001=TA006
                                    LEFT JOIN [TK].dbo.BOMMC ON MC001=TA006
                                    LEFT JOIN [TK].dbo.BOMMD ON MD035 LIKE '%箱%' AND MD003 LIKE '2%' AND MD007>1 AND MD001=TA006 AND MD003 NOT  IN ( SELECT  [MD003]  FROM [TKMOC].[dbo].[MOCHALFPRODUCTDBOXSLIMITS] )
                                    LEFT JOIN [TKMOC].[dbo].[MOCHALFPRODUCTDBOXS] ON [MOCHALFPRODUCTDBOXS].MB001=TA006
                                    WHERE [TA001]+[TA002] NOT IN (SELECT [TA001]+[TA002] FROM [TKMOC].[dbo].[REPORTMOCMANULINE] WHERE TA003='{1}')
                                    AND TA003='{1}' 
                                    ) AS TEMP

                                    GROUP BY  TA021,TA001,TA002,TA003,TA006,TA007,TA015,TA017,TA034,TA035,PCT,UDF01,ALLERGEN,SPEC,BARS,BOXS,DATES,TA029,ORI
                                    ORDER BY TA003,TA021,TA001,TA002     
                                    ", LOTNO, TA003);


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
        public void SETFASTREPORT2(string SDAY)
        {
            string SQL;
            report1 = new Report();
            report1.Load(@"REPORT\內包生產排程.frx"); 


            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;


            TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;
            SQL = SETFASETSQL2(SDAY);
            Table.SelectCommand = SQL;

            report1.Preview = previewControl1;
            report1.Show();
        }


        public string SETFASETSQL2(string SDAY)
        {
            StringBuilder FASTSQL = new StringBuilder();

            //,CASE WHEN TA006 NOT LIKE '4%' THEN CONVERT(decimal(16,3),TA015/ISNULL(MC004,1)) ELSE 0 END AS '桶數'
            //,CASE WHEN TA006 LIKE '4%' THEN CONVERT(decimal(16, 3), TA015 / ISNULL(MD007, 1) * ISNULL(MD010, 1)) ELSE 0 END AS '箱數'

            FASTSQL.AppendFormat(@"    
                                SELECT 
                                 [ID]
                                ,[REPORTMOCMANULINE].[MANULINE] AS '生產線別'
                                ,[REPORTMOCMANULINE].[LOTNO] AS 'LOTNO'
                                ,[REPORTMOCMANULINE].[TA001] AS '製令別'
                                ,[REPORTMOCMANULINE].[TA002] AS '製令編號'
                                ,CONVERT(NVARCHAR,[REPORTMOCMANULINE].[TA003],112) AS '製令日期'
                                ,[REPORTMOCMANULINE].[TA006] AS '品號'
                                ,[REPORTMOCMANULINE].[TA007] AS '單位'
                                ,[REPORTMOCMANULINE].[TA015] AS '預計產量'
                                ,[REPORTMOCMANULINE].[TA017] AS '實際產出'
                                ,[REPORTMOCMANULINE].[MB002] AS '品名'
                                ,[REPORTMOCMANULINE].[MB003] AS '規格'
                                ,[REPORTMOCMANULINE].[PCTS] AS '比例'
                                ,[REPORTMOCMANULINE].[SEQ] AS '順序'
                                ,[REPORTMOCMANULINE].[ALLERGEN]  AS '過敏原'
                                ,[REPORTMOCMANULINE].[COOKIES] AS '餅體'
                                ,[REPORTMOCMANULINE].[BARS] AS '桶數'
                                ,[REPORTMOCMANULINE].[BOXS] AS '箱數'
                                ,CONVERT(NVARCHAR,[REPORTMOCMANULINE].[VDATES],112) AS '有效日期'
                                ,[REPORTMOCMANULINE].[COMMENT] AS '備註'
                                ,MOCTA.TA026 AS '訂單別'
                                ,MOCTA.TA027 AS '訂單號'
                                ,TC053  AS '客戶'
                                ,[REPORTMOCMANULINE].[ORI] AS '素別'
                                FROM [TKMOC].[dbo].[REPORTMOCMANULINE]
                                LEFT JOIN [TK].dbo.MOCTA ON [REPORTMOCMANULINE].TA001=MOCTA.[TA001] AND [REPORTMOCMANULINE].[TA002]=MOCTA.[TA002]
                                LEFT JOIN [TK].dbo.COPTC ON TC001= TA026 AND TC002=TA027 
                                WHERE CONVERT(NVARCHAR,[REPORTMOCMANULINE].TA003,112)='{0}'   
	                            AND [REPORTMOCMANULINE].[MANULINE] IN ('02','03')
                                ORDER BY [REPORTMOCMANULINE].TA003,[MANULINE],[REPORTMOCMANULINE].TA001,[REPORTMOCMANULINE].TA002   

                                ", SDAY);

            return FASTSQL.ToString();
        }

        #endregion

        #region BUTTON
        private void button7_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox3.Text.Trim()))
            {
                ADDDELETEMOCLOTNO(dateTimePicker1.Value.ToString("yyyyMMdd"), textBox3.Text.Trim(), dateTimePicker5.Value.ToString("yyyyMMdd"));
                textBox3.Text = SEARCHMOCLOTNO(dateTimePicker1.Value.ToString("yyyyMMdd"));

                ADDREPORTMOCMANULINE(textBox3.Text.Trim(), dateTimePicker1.Value.ToString("yyyyMMdd"));

                MessageBox.Show("完成");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT2(dateTimePicker1.Value.ToString("yyyyMMdd"));
        }

        #endregion


    }
}
