using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI;
using NPOI.HPSF;
using NPOI.HSSF;
using NPOI.HSSF.UserModel;
using NPOI.POIFS;
using NPOI.Util;
using NPOI.HSSF.Util;
using NPOI.HSSF.Extractor;
using System.IO;
using System.Data.SqlClient;
using NPOI.SS.UserModel;
using System.Configuration;
using NPOI.XSSF.UserModel;
using FastReport;
using FastReport.Data;
using System.Collections;

namespace TKMOC
{
    public partial class frmPURINVCHECK : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();

        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        SqlDataAdapter adapter2= new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder2 = new SqlCommandBuilder();
        SqlDataAdapter adapter3 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder3 = new SqlCommandBuilder();
        SqlDataAdapter adapter4 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder4 = new SqlCommandBuilder();


        DataSet ds = new DataSet();
        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        int rownum = 0;
        int result;

        ArrayList myAL = new ArrayList();
        string MOCTA001;
        string MOCTA002;
        string MOCTA003;
        string ID;
        string MAXID;
        string MF004 = null;

        int rowIndexDG1 = -1;
        string MD003;

        public class PURTA
        {
            public string COMPANY;
            public string CREATOR;
            public string USR_GROUP;
            public string CREATE_DATE;
            public string MODIFIER;
            public string MODI_DATE;
            public string FLAG;
            public string CREATE_TIME;
            public string MODI_TIME;
            public string TRANS_TYPE;
            public string TRANS_NAME;
            public string sync_date;
            public string sync_time;
            public string sync_mark;
            public string sync_count;
            public string DataUser;
            public string DataGroup;
            public string TA001;
            public string TA002;
            public string TA003;
            public string TA004;
            public string TA005;
            public string TA006;
            public string TA007;
            public string TA008;
            public string TA009;
            public string TA010;
            public string TA011;
            public string TA012;
            public string TA013;
            public string TA014;
            public string TA015;
            public string TA016;
            public string TA017;
            public string TA018;
            public string TA019;
            public string TA020;
            public string TA021;
            public string TA022;
            public string TA023;
            public string TA024;
            public string TA025;
            public string TA026;
            public string TA027;
            public string TA028;
            public string TA029;
            public string TA030;
            public string TA031;
            public string TA032;
            public string TA033;
            public string TA034;
            public string TA035;
            public string TA036;
            public string TA037;
            public string TA038;
            public string TA039;
            public string TA040;
            public string TA041;
            public string TA042;
            public string TA043;
            public string TA044;
            public string TA045;
            public string TA046;
            public string UDF01;
            public string UDF02;
            public string UDF03;
            public string UDF04;
            public string UDF05;
            public string UDF06;
            public string UDF07;
            public string UDF08;
            public string UDF09;
            public string UDF10;
        }

        public class PURTB
        {
            public string COMPANY;
            public string CREATOR;
            public string USR_GROUP;
            public string CREATE_DATE;
            public string MODIFIER;
            public string MODI_DATE;
            public string FLAG;
            public string CREATE_TIME;
            public string MODI_TIME;
            public string TRANS_TYPE;
            public string TRANS_NAME;
            public string sync_date;
            public string sync_time;
            public string sync_mark;
            public string sync_count;
            public string DataUser;
            public string DataGroup;
            public string TB001;
            public string TB002;
            public string TB003;
            public string TB004;
            public string TB005;
            public string TB006;
            public string TB007;
            public string TB008;
            public string TB009;
            public string TB010;
            public string TB011;
            public string TB012;
            public string TB013;
            public string TB014;
            public string TB015;
            public string TB016;
            public string TB017;
            public string TB018;
            public string TB019;
            public string TB020;
            public string TB021;
            public string TB022;
            public string TB023;
            public string TB024;
            public string TB025;
            public string TB026;
            public string TB027;
            public string TB028;
            public string TB029;
            public string TB030;
            public string TB031;
            public string TB032;
            public string TB033;
            public string TB034;
            public string TB035;
            public string TB036;
            public string TB037;
            public string TB038;
            public string TB039;
            public string TB040;
            public string TB041;
            public string TB042;
            public string TB043;
            public string TB044;
            public string TB045;
            public string TB046;
            public string TB047;
            public string TB048;
            public string TB049;
            public string TB050;
            public string TB051;
            public string TB052;
            public string TB053;
            public string TB054;
            public string TB055;
            public string TB056;
            public string TB057;
            public string TB058;
            public string TB059;
            public string TB060;
            public string TB061;
            public string TB062;
            public string TB063;
            public string TB064;
            public string TB065;
            public string TB066;
            public string TB067;
            public string TB068;
            public string TB069;
            public string TB070;
            public string TB071;
            public string TB072;
            public string TB073;
            public string TB074;
            public string TB075;
            public string TB076;
            public string TB077;
            public string TB078;
            public string TB079;
            public string TB080;
            public string TB081;
            public string TB082;
            public string TB083;
            public string TB084;
            public string TB085;
            public string TB086;
            public string TB087;
            public string TB088;
            public string TB089;
            public string TB090;
            public string TB091;
            public string TB092;
            public string TB093;
            public string TB094;
            public string TB095;
            public string TB096;
            public string TB097;
            public string TB098;
            public string TB099;
            public string UDF01;
            public string UDF02;
            public string UDF03;
            public string UDF04;
            public string UDF05;
            public string UDF06;
            public string UDF07;
            public string UDF08;
            public string UDF09;
            public string UDF10;
        }

        public frmPURINVCHECK()
        {
            InitializeComponent();
        }
        #region FUNCTION

        public void SEARCHINVMC()
        {
            DateTime SEARCHDATE2 = DateTime.Now;

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT MC001 AS '品號',MB002 AS '品名',MC002 AS '庫別',MB004 AS '單位',MC004 AS '安全批量',MC005 AS '補貨點'");
                sbSql.AppendFormat(@"  ,ISNULL((SELECT SUM(LA005*LA011) FROM [TK].dbo.INVLA WHERE MC001=LA001 AND LA009=MC002) ,0) AS '目前庫存'");
                sbSql.AppendFormat(@"  ,ISNULL(((SELECT SUM(LA005*LA011) FROM [TK].dbo.INVLA WHERE MC001=LA001 AND LA009=MC002) -MC004),0) AS '庫存差異量'");
                sbSql.AppendFormat(@"  ,(SELECT ISNULL(CONVERT(DECIMAL(16,2),SUM(NUM)),0) FROM [TK].dbo.VPURTDINVMD WHERE  TD004=MC001 AND TD007=TD007 AND TD012>='{0}') AS '總採購量'", SEARCHDATE2.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  ,(SELECT TOP 1 ISNULL(TD012,'')+' 預計到貨:'+CONVERT(nvarchar,CONVERT(DECIMAL(16,2),NUM))  FROM [TK].dbo.VPURTDINVMD WHERE  TD004=MC001 AND TD007=TD007 AND TD012>='{0}') AS '最快採購日'", SEARCHDATE2.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  ,(ISNULL((MC004-(SELECT SUM(LA005*LA011) FROM [TK].dbo.INVLA WHERE MC001=LA001 AND LA009=MC002) ),0)-(SELECT ISNULL(CONVERT(DECIMAL(16,2),SUM(NUM)),0) FROM [TK].dbo.VPURTDINVMD WHERE  TD004=MC001 AND TD007=TD007 AND TD012>='{0}') ) AS '需採購量' ", SEARCHDATE2.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  FROM [TK].dbo.INVMC,[TK].dbo.INVMB");
                sbSql.AppendFormat(@"  WHERE MC001=MB001");
                sbSql.AppendFormat(@"  AND MC002=@MC002 AND MC003='201904制定'");
                sbSql.AppendFormat(@"  ORDER BY (ISNULL((MC004-(SELECT SUM(LA005*LA011) FROM [TK].dbo.INVLA WHERE MC001=LA001 AND LA009=MC002) ),0)-(SELECT ISNULL(CONVERT(DECIMAL(16,2),SUM(NUM)),0) FROM [TK].dbo.VPURTDINVMD WHERE  TD004=MC001 AND TD007=TD007 AND TD012>='{0}') ) DESC,MC001", SEARCHDATE2.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
                adapter.SelectCommand.Parameters.AddWithValue("@MC002", "20004");

                sqlCmdBuilder = new SqlCommandBuilder(adapter);


                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    //建立一筆新的DataRow，並且等於新的dt row
                    DataRow row = ds.Tables["ds"].NewRow();

                    //指定每個欄位要儲存的資料                   
                    row[0] = "本日無資料"; ;

                    //ds
                    ds.Tables["dsINVMC"].Rows.Add(row);
                    
                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        dataGridView1.DataSource = ds.Tables["ds"];
                        dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;//調整寬度(標題+儲存格)

                        //dataGridView1.AutoResizeColumns();
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

        public void SEARCHINVMC2(string SDAY,string EDAY)
        {
            
            DataSet ds = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                
                sbSql.AppendFormat(@"  
                                    SELECT 品號,品名,安全水位,庫存量,採購未交數量,最快採購日,需求量
                                    ,CASE WHEN (庫存量-需求量)<安全水位 THEN '低於水位' ELSE '' END AS '庫存-需求'
                                    
                                    FROM (
                                    SELECT MOCINV.MB001 AS '品號',MOCINV.MB002 AS '品名'
                                    ,CASE WHEN  DATEPART(month,GETDATE()) IN ('1','2','3') THEN NUMS ELSE 
                                    (CASE WHEN  DATEPART(month,GETDATE()) IN ('4','5','6') THEN NUMS2 ELSE 
                                    (CASE WHEN  DATEPART(month,GETDATE()) IN ('7','8','9') THEN NUMS3 ELSE 
                                    (CASE WHEN  DATEPART(month,GETDATE()) IN ('10','11','12') THEN NUMS4 ELSE 0 END)
                                     END)  
                                    END) 
                                    END AS '安全水位'
                                    ,(SELECT SUM(LA005*LA011) FROM [TK].dbo.INVLA WHERE LA009 IN ('20004','20006') AND LA001=MOCINV.MB001) AS '庫存量'
                                    ,(SELECT ISNULL(SUM(TD008 - TD015), 0) FROM[TK].dbo.PURTD WHERE TD004 =MOCINV.MB001 AND TD018 = 'Y' AND TD016 = 'N' AND TD012>='20210226'  AND TD012<='20210305' ) AS '採購未交數量'
                                    ,(SELECT TOP 1 ISNULL(TD012,'')+' 預計到貨:'+CONVERT(nvarchar,CONVERT(DECIMAL(16,2),NUM))  FROM [TK].dbo.VPURTDINVMD WHERE  TD004=MOCINV.MB001 AND TD007=TD007 AND TD012>='20210226') AS '最快採購日'
                                    ,ISNULL(TEMP2.TNUM,0)  AS '需求量'
                                    ,NUMS ,NUMS2 ,NUMS3 ,NUMS4 
                                    FROM [TKMOC].dbo.MOCINV
                                    LEFT JOIN (
	                                    SELECT [MD003],SUM(TNUM) TNUM
                                        FROM(
                                        SELECT[MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                                        , CONVERT(decimal(16, 3), ([PACKAGE] /[MC004] *[MD006] /[MD007] * (1 +[MD008]))) AS TNUM
                                        FROM[TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                                        WHERE[MOCMANULINE].MB001 = MC001
                                        AND MC001 = MD001
                                        AND [MANU] = '新廠包裝線'
                                        AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                                        UNION ALL
                                        SELECT[MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003], MC1.[MC001], MC1.[MC004], CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, CASE WHEN ISNULL(MD2.[MD035], '') = '' THEN MD1.[MD035] ELSE MD2.[MD035] END, MD1.[MD006], MD1.[MD007], MD1.[MD008]
                                        , CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN CONVERT(decimal(16, 3), ([NUM] / MC1.[MC004] * MD1.[MD006] / MD1.[MD007] * (1 + MD1.[MD008]))) ELSE CONVERT(decimal(16, 3), ([NUM] / MC1.[MC004] * MD1.[MD006] / MD1.[MD007] * (1 + MD1.[MD008])) / MC2.[MC004] * MD2.[MD006] / MD2.[MD007] * (1 + MD2.[MD008])) END  AS TNUM
                                        FROM[TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                                        LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001 = MD1.MD003
                                        LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001 = MD2.MD001
                                        WHERE[MOCMANULINE].MB001 = MC1.MC001
                                        AND MC1.MC001 = MD1.MD001
                                        AND [MANU] = '新廠製一組'
                                        AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                                        AND REPLACE(MC1.[MC001], ' ', '') + REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') NOT IN(SELECT REPLACE([MB001], ' ', '') + REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') FROM[TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001 = MD1.MD003 LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001 = MD2.MD001 WHERE[PREMANUUSEDINVMB].MB001 = MC1.MC001 AND MC1.MC001 = MD1.MD001 AND(REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') LIKE '1%' OR REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') LIKE '203%'))
                                        UNION ALL
                                        SELECT[MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003], MC1.[MC001], MC1.[MC004], CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, CASE WHEN ISNULL(MD2.[MD035], '') = '' THEN MD1.[MD035] ELSE MD2.[MD035] END, MD1.[MD006], MD1.[MD007], MD1.[MD008]
                                        , CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN CONVERT(decimal(16, 3), ([NUM] / MC1.[MC004] * MD1.[MD006] / MD1.[MD007] * (1 + MD1.[MD008]))) ELSE CONVERT(decimal(16, 3), ([NUM] / MC1.[MC004] * MD1.[MD006] / MD1.[MD007] * (1 + MD1.[MD008])) / MC2.[MC004] * MD2.[MD006] / MD2.[MD007] * (1 + MD2.[MD008])) END  AS TNUM
                                        FROM[TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                                        LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001 = MD1.MD003
                                        LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001 = MD2.MD001
                                        WHERE[MOCMANULINE].MB001 = MC1.MC001
                                        AND MC1.MC001 = MD1.MD001
                                        AND [MANU] = '新廠製二組'
                                        AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                                        AND REPLACE(MC1.[MC001], ' ', '') + REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') NOT IN(SELECT REPLACE([MB001], ' ', '') + REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') FROM[TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001 = MD1.MD003 LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001 = MD2.MD001 WHERE[PREMANUUSEDINVMB].MB001 = MC1.MC001 AND MC1.MC001 = MD1.MD001 AND(REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') LIKE '1%' OR REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') LIKE '203%'))
                                        UNION ALL
                                        SELECT[MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END , CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END, MD1.[MD006], MD1.[MD007], MD1.[MD008]
                                        , CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END AS TNUM
                                        FROM[TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                                        LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003
                                        LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001= MD2.MD001
                                        WHERE[MOCMANULINE].MB001= MC1.MC001
                                        AND MC1.MC001= MD1.MD001
                                        AND [MANU]= '新廠製三組(手工)'
                                        AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                                        AND REPLACE(MC1.[MC001],' ','')+REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') NOT IN (SELECT REPLACE([MB001],' ','')+REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') FROM[TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003 LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001= MD2.MD001 WHERE[PREMANUUSEDINVMB].MB001= MC1.MC001 AND MC1.MC001= MD1.MD001 AND  (REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') LIKE '1%' OR REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') LIKE '203%' ) )
      
                                        ) AS TEMP  GROUP BY [MD003]
                                    ) TEMP2 ON TEMP2.MD003=MOCINV.MB001
                                    ) TMEP3
                                    ORDER BY TMEP3.品號

                                    ", SDAY, EDAY);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);        

                sqlCmdBuilder = new SqlCommandBuilder(adapter);


                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    //建立一筆新的DataRow，並且等於新的dt row
                    DataRow row = ds.Tables["ds"].NewRow();

                    //指定每個欄位要儲存的資料                   
                    row[0] = "本日無資料"; ;

                    //ds
                    ds.Tables["dsINVMC"].Rows.Add(row);

                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        dataGridView3.DataSource = ds.Tables["ds"];
                        dataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;//調整寬度(標題+儲存格)

                        //根据列表中数据不同，显示不同颜色背景
                        foreach (DataGridViewRow dgRow in dataGridView3.Rows)
                        {
                            //判断
                            if (!string.IsNullOrEmpty(dgRow.Cells[7].Value.ToString()))
                            {
                                //将这行的背景色设置成Pink
                                dgRow.DefaultCellStyle.BackColor = Color.Pink;
                            }
                        }
                        //dataGridView1.AutoResizeColumns();
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


        public void ADDPURTAB()
        {
            myAL.Clear();

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if(Convert.ToDecimal(row.Cells[10].Value.ToString())>0)
                {
                    myAL.Add(Convert.ToDecimal(row.Cells[10].Value.ToString()));
                }
                
            }

            //foreach
            foreach (object num in myAL)
            {
                
            }
        }

        public string GETMAXMOCTA002(string MOCTA001)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds1.Clear();

                sbSql.AppendFormat(@"  SELECT ISNULL(MAX(TA002),'00000000000') AS ID ");
                sbSql.AppendFormat(@"  FROM [TK].[dbo].[PURTA] ");
                //sbSql.AppendFormat(@"  WHERE  TC001='{0}' AND TC003='{1}'", "A542","20170119");
                sbSql.AppendFormat(@"  WHERE [TA001]='{0}' AND TA003='{1}'", MOCTA001, dateTimePicker1.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count == 0)
                {
                    return null;
                }
                else
                {
                    if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        MAXID = SETID(ds1.Tables["ds1"].Rows[0]["ID"].ToString(), dateTimePicker1.Value);
                        return MAXID;

                    }
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

        public string SETID(string MAXID, DateTime dt)
        {
            if (MAXID.Equals("00000000000"))
            {
                return dt.ToString("yyyyMMdd") + "001";
            }

            else
            {
                int serno = Convert.ToInt16(MAXID.Substring(8, 3));
                serno = serno + 1;
                string temp = serno.ToString();
                temp = temp.PadLeft(3, '0');
                return dt.ToString("yyyyMMdd") + temp.ToString();
            }
        }
        public string GETMAXID()
        {

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds2.Clear();

                sbSql.AppendFormat(@"  SELECT ISNULL(MAX(ID),'00000000000') AS ID");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[PURTAB] ");
                //sbSql.AppendFormat(@"  WHERE  TC001='{0}' AND TC003='{1}'", "A542","20170119");
                sbSql.AppendFormat(@"  WHERE [IDDATES]='{0}'", dateTimePicker1.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                sqlConn.Open();
                ds2.Clear();
                adapter2.Fill(ds2, "ds2");
                sqlConn.Close();


                if (ds2.Tables["ds2"].Rows.Count == 0)
                {
                    return null;
                }
                else
                {
                    if (ds2.Tables["ds2"].Rows.Count >= 1)
                    {
                        MAXID = SETID(ds2.Tables["ds2"].Rows[0]["ID"].ToString(), dateTimePicker1.Value);
                        return MAXID;

                    }
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

        public void SETNULL()
        {
            textBox1.Text = null;
            textBox4.Text = null;
            textBox5.Text = null;
        }

        public void SEARCHPURTAB()
        {
            StringBuilder SLQURY = new StringBuilder();
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();

                sbSql.AppendFormat(@"  SELECT [ID] AS '批號',[IDDATES] AS '請購日',[PURTA001] AS '請購單別',[PURTA002] AS '請購單號'");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[PURTAB]");
                sbSql.AppendFormat(@"  WHERE [IDDATES]='{0}'", dateTimePicker1.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  GROUP BY  [ID],[IDDATES],[PURTA001],[PURTA002] ");
                sbSql.AppendFormat(@"  ");

                adapter3 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder3 = new SqlCommandBuilder(adapter3);
                sqlConn.Open();
                ds3.Clear();
                adapter3.Fill(ds3, "ds3");
                sqlConn.Close();


                if (ds3.Tables["ds3"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;

                }
                else
                {
                    if (ds3.Tables["ds3"].Rows.Count >= 1)
                    {
                        dataGridView2.DataSource = ds3.Tables["ds3"];

                        dataGridView2.AutoResizeColumns();
                        dataGridView2.FirstDisplayedScrollingRowIndex = dataGridView2.RowCount - 1;


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

        public void ADDPURTAB(string ID)
        {
            sbSql.Clear();

            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {
                if (dr.Cells[10].Value != null && Convert.ToDecimal(dr.Cells[10].Value)>0)
                {
                    
                    sbSql.AppendFormat(@" INSERT INTO [TKMOC].[dbo].[PURTAB]");
                    sbSql.AppendFormat(@" ([ID],[IDDATES],[MB001],[NUM],[PURTA001],[PURTA002])");
                    sbSql.AppendFormat(@" VALUES ({0},'{1}','{2}','{3}','{4}','{5}')", ID, dateTimePicker1.Value.ToString("yyyyMMdd"), dr.Cells["品號"].Value.ToString(), dr.Cells["需採購量"].Value.ToString(), "", "");
                    sbSql.AppendFormat(@" ");
                }
            }

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

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

        public void ADDERPPURAB()
        {
            PURTA PURTA = new PURTA();
            PURTB PURTB = new PURTB();
            MOCTA003 = dateTimePicker1.Value.ToString("yyyyMMdd");

            PURTA = SETPURTA();
            PURTB = SETPURTB();

            connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);

            sqlConn.Close();
            sqlConn.Open();
            tran = sqlConn.BeginTransaction();

            sbSql.Clear();

            sbSql.AppendFormat(" INSERT INTO [TK].[dbo].[PURTA]");
            sbSql.AppendFormat(" ( [COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER]");
            sbSql.AppendFormat(" ,[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE]");
            sbSql.AppendFormat(" ,[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count]");
            sbSql.AppendFormat(" ,[DataUser],[DataGroup]");
            sbSql.AppendFormat(" ,[TA001],[TA002],[TA003],[TA004],[TA005]");
            sbSql.AppendFormat(" ,[TA006],[TA007],[TA008],[TA009],[TA010]");
            sbSql.AppendFormat(" ,[TA011],[TA012],[TA013],[TA014],[TA015]");
            sbSql.AppendFormat(" ,[TA016],[TA017],[TA018],[TA019],[TA020]");
            sbSql.AppendFormat(" ,[TA021],[TA022],[TA023],[TA024],[TA025]");
            sbSql.AppendFormat(" ,[TA026],[TA027],[TA028],[TA029],[TA030]");
            sbSql.AppendFormat(" ,[TA031],[TA032],[TA033],[TA034],[TA035]");
            sbSql.AppendFormat(" ,[TA036],[TA037],[TA038],[TA039],[TA040]");
            sbSql.AppendFormat(" ,[TA041],[TA042],[TA043],[TA044],[TA045]");
            sbSql.AppendFormat(" ,[TA046],[UDF01],[UDF02],[UDF03],[UDF04]");
            sbSql.AppendFormat(" ,[UDF05],[UDF06],[UDF07],[UDF08],[UDF09]");
            sbSql.AppendFormat(" ,[UDF10]");
            sbSql.AppendFormat(" )");
            sbSql.AppendFormat(" VALUES ");
            sbSql.AppendFormat(" ('{0}','{1}','{2}','{3}','{4}',", PURTA.COMPANY, PURTA.CREATOR, PURTA.USR_GROUP, PURTA.CREATE_DATE, PURTA.MODIFIER);
            sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.MODI_DATE, PURTA.FLAG, PURTA.CREATE_TIME, PURTA.MODI_TIME, PURTA.TRANS_TYPE);
            sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TRANS_NAME, PURTA.sync_date, PURTA.sync_time, PURTA.sync_mark, PURTA.sync_count);
            sbSql.AppendFormat(" '{0}','{1}',", PURTA.DataUser, PURTA.DataGroup);
            sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA001, PURTA.TA002, PURTA.TA003, PURTA.TA004, PURTA.TA005);
            sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA006, PURTA.TA007, PURTA.TA008, PURTA.TA009, PURTA.TA010);
            sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA011, PURTA.TA012, PURTA.TA013, PURTA.TA014, PURTA.TA015);
            sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA016, PURTA.TA017, PURTA.TA018, PURTA.TA019, PURTA.TA020);
            sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA021, PURTA.TA022, PURTA.TA023, PURTA.TA024, PURTA.TA025);
            sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA026, PURTA.TA027, PURTA.TA028, PURTA.TA029, PURTA.TA030);
            sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA031, PURTA.TA032, PURTA.TA033, PURTA.TA034, PURTA.TA035);
            sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA036, PURTA.TA037, PURTA.TA038, PURTA.TA039, PURTA.TA040);
            sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA041, PURTA.TA042, PURTA.TA043, PURTA.TA044, PURTA.TA045);
            sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.TA046, PURTA.UDF01, PURTA.UDF02, PURTA.UDF03, PURTA.UDF04);
            sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", PURTA.UDF05, PURTA.UDF06, PURTA.UDF07, PURTA.UDF08, PURTA.UDF09);
            sbSql.AppendFormat(" '{0}'", PURTA.UDF10);
            sbSql.AppendFormat(" )");
            sbSql.AppendFormat(" ");
            sbSql.AppendFormat(" INSERT INTO [TK].[dbo].[PURTB]");
            sbSql.AppendFormat(" ([COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER]");
            sbSql.AppendFormat(" ,[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE]");
            sbSql.AppendFormat(" ,[TRANS_NAME],[sync_date],[sync_time],[sync_mark],[sync_count]");
            sbSql.AppendFormat(" ,[DataUser],[DataGroup]");
            sbSql.AppendFormat(" ,[TB001],[TB002],[TB003],[TB004],[TB005]");
            sbSql.AppendFormat(" ,[TB006],[TB007],[TB008],[TB009],[TB010]");
            sbSql.AppendFormat(" ,[TB011],[TB012],[TB013],[TB014],[TB015]");
            sbSql.AppendFormat(" ,[TB016],[TB017],[TB018],[TB019],[TB020]");
            sbSql.AppendFormat(" ,[TB021],[TB022],[TB023],[TB024],[TB025]");
            sbSql.AppendFormat(" ,[TB026],[TB027],[TB028],[TB029],[TB030]");
            sbSql.AppendFormat(" ,[TB031],[TB032],[TB033],[TB034],[TB035]");
            sbSql.AppendFormat(" ,[TB036],[TB037],[TB038],[TB039],[TB040]");
            sbSql.AppendFormat(" ,[TB041],[TB042],[TB043],[TB044],[TB045]");
            sbSql.AppendFormat(" ,[TB046],[TB047],[TB048],[TB049],[TB050]");
            sbSql.AppendFormat(" ,[TB051],[TB052],[TB053],[TB054],[TB055]");
            sbSql.AppendFormat(" ,[TB056],[TB057],[TB058],[TB059],[TB060]");
            sbSql.AppendFormat(" ,[TB061],[TB062],[TB063],[TB064],[TB065]");
            sbSql.AppendFormat(" ,[TB066],[TB067],[TB068],[TB069],[TB070]");
            sbSql.AppendFormat(" ,[TB071],[TB072],[TB073],[TB074],[TB075]");
            sbSql.AppendFormat(" ,[TB076],[TB077],[TB078],[TB079],[TB080]");
            sbSql.AppendFormat(" ,[TB081],[TB082],[TB083],[TB084],[TB085]");
            sbSql.AppendFormat(" ,[TB086],[TB087],[TB088],[TB089],[TB090]");
            sbSql.AppendFormat(" ,[TB091],[TB092],[TB093],[TB094],[TB095]");
            sbSql.AppendFormat(" ,[TB096],[TB097],[TB098],[TB099],[UDF01]");
            sbSql.AppendFormat(" ,[UDF02],[UDF03],[UDF04],[UDF05],[UDF06]");
            sbSql.AppendFormat(" ,[UDF07],[UDF08],[UDF09],[UDF10]");
            sbSql.AppendFormat(" )");
            sbSql.AppendFormat(" (SELECT '{0}' [COMPANY],'{1}' [CREATOR],'{2}' [USR_GROUP],'{3}' [CREATE_DATE],'{4}' [MODIFIER],", PURTB.COMPANY, PURTB.CREATOR, PURTB.USR_GROUP, PURTB.CREATE_DATE, PURTB.MODIFIER);
            sbSql.AppendFormat(" '{0}' [MODI_DATE],{1} [FLAG],'{2}' [CREATE_TIME],'{3}' [MODI_TIME],'{4}' [TRANS_TYPE],", PURTB.MODI_DATE, PURTB.FLAG, PURTB.CREATE_TIME, PURTB.MODI_TIME, PURTB.TRANS_TYPE);
            sbSql.AppendFormat(" '{0}' [TRANS_NAME],'{1}' [sync_date],'{2}' [sync_time],'{3}' [sync_mark],{4} [sync_count],", PURTB.TRANS_NAME, PURTB.sync_date, PURTB.sync_time, PURTB.sync_mark, PURTB.sync_count);
            sbSql.AppendFormat(" '{0}' [DataUser],'{1}' [DataGroup],", PURTB.DataUser, PURTB.DataGroup);
            sbSql.AppendFormat(" '{0}' [TB001],'{1}' [TB002],Right('0000' + Cast(ROW_NUMBER() OVER( ORDER BY [PURTAB].MB001)  as varchar),4) AS TB003,[PURTAB].MB001 AS TB004,MB002 AS TB005,", PURTB.TB001, PURTB.TB002);
            sbSql.AppendFormat(" MB003 AS TB006,MB004 AS TB007,MB017 AS TB008,[PURTAB].NUM AS TB009,MB032 AS TB010,");
            sbSql.AppendFormat(" '{0}' [TB011],'{1}' [TB012],'{2}' [TB013],[PURTAB].NUM  [TB014],'{3}' [TB015],", PURTB.TB011, PURTB.TB012, PURTB.TB013, PURTB.TB015);
            sbSql.AppendFormat(" '{0}' [TB016],MB050 AS TB017,ROUND((MB050*[PURTAB].NUM ),0) AS TB018,'{1}' [TB019],'{2}' [TB020],", PURTB.TB016, PURTB.TB019, PURTB.TB020);
            sbSql.AppendFormat(" '{0}' [TB021],'{1}' [TB022],'{2}' [TB023],'{3}' [TB024],'{4}' [TB025],", PURTB.TB021, PURTB.TB022, PURTB.TB023, PURTB.TB024, PURTB.TB025);
            sbSql.AppendFormat(" '{0}' [TB026],'{1}' [TB027],'{2}' [TB028],'{3}' [TB029],'{4}' [TB030],", PURTB.TB026, PURTB.TB027, PURTB.TB028, PURTB.TB029, PURTB.TB030);
            sbSql.AppendFormat(" '{0}' [TB031],'{1}' [TB032],'{2}' [TB033],{3} [TB034],{4} [TB035],", PURTB.TB031, PURTB.TB032, PURTB.TB033, PURTB.TB034, PURTB.TB035);
            sbSql.AppendFormat(" '{0}' [TB036],'{1}' [TB037],'{2}' [TB038],'{3}' [TB039],'{4}' [TB040],", PURTB.TB036, PURTB.TB037, PURTB.TB038, PURTB.TB039, PURTB.TB040);
            sbSql.AppendFormat(" {0} [TB041],'{1}' [TB042],'{2}' [TB043],'{3}' [TB044],'{4}' [TB045],", PURTB.TB041, PURTB.TB042, PURTB.TB043, PURTB.TB044, PURTB.TB045);
            sbSql.AppendFormat(" '{0}' [TB046],'{1}' [TB047],'{2}' [TB048],{3} [TB049],'{4}' [TB050],", PURTB.TB046, PURTB.TB047, PURTB.TB048, PURTB.TB049, PURTB.TB050);
            sbSql.AppendFormat(" {0} [TB051],{1} [TB052],{2} [TB053],'{3}' [TB054],'{4}' [TB055],", PURTB.TB051, PURTB.TB052, PURTB.TB053, PURTB.TB054, PURTB.TB055);
            sbSql.AppendFormat(" '{0}' [TB056],'{1}' [TB057],'{2}' [TB058],'{3}' [TB059],'{4}' [TB060],", PURTB.TB056, PURTB.TB057, PURTB.TB058, PURTB.TB059, PURTB.TB060);
            sbSql.AppendFormat(" '{0}' [TB061],'{1}' [TB062],{2} [TB063],'{3}' [TB064],'{4}' [TB065],", PURTB.TB061, PURTB.TB062, PURTB.TB063, PURTB.TB064, PURTB.TB065);
            sbSql.AppendFormat(" '{0}' [TB066],'{1}' [TB067],{2} [TB068],{3} [TB069],'{4}' [TB070],", PURTB.TB066, PURTB.TB067, PURTB.TB068, PURTB.TB069, PURTB.TB070);
            sbSql.AppendFormat(" '{0}' [TB071],'{1}' [TB072],'{2}' [TB073],'{3}' [TB074],{4} [TB075],", PURTB.TB071, PURTB.TB072, PURTB.TB073, PURTB.TB074, PURTB.TB075);
            sbSql.AppendFormat(" '{0}' [TB076],{1} [TB077],'{2}' [TB078],'{3}' [TB079],'{4}' [TB080],", PURTB.TB076, PURTB.TB077, PURTB.TB078, PURTB.TB079, PURTB.TB080);
            sbSql.AppendFormat(" {0} [TB081],{1} [TB082],{2} [TB083],{3} [TB084],{4} [TB085],", PURTB.TB081, PURTB.TB082, PURTB.TB083, PURTB.TB084, PURTB.TB085);
            sbSql.AppendFormat(" '{0}' [TB086],'{1}' [TB087],{2} [TB088],'{3}' [TB089],{4} [TB090],", PURTB.TB086, PURTB.TB087, PURTB.TB088, PURTB.TB089, PURTB.TB090);
            sbSql.AppendFormat(" {0} [TB091],{1} [TB092],{2} [TB093],'{3}' [TB094],'{4}' [TB095],", PURTB.TB091, PURTB.TB092, PURTB.TB093, PURTB.TB094, PURTB.TB095);
            sbSql.AppendFormat(" '{0}' [TB096],'{1}' [TB097],'{2}' [TB098],'{3}' [TB099],'{4}' [UDF01],", PURTB.TB096, PURTB.TB097, PURTB.TB098, PURTB.TB099, PURTB.UDF01);
            sbSql.AppendFormat(" '{0}' [UDF02],'{1}' [UDF03],'{2}' [UDF04],'{3}' [UDF05],{4} [UDF06],", PURTB.UDF02, PURTB.UDF03, PURTB.UDF04, PURTB.UDF05, PURTB.UDF06);
            sbSql.AppendFormat(" {0} [UDF07],{1}[UDF08],{2} [UDF09],{3} [UDF10]", PURTB.UDF07, PURTB.UDF08, PURTB.UDF09, PURTB.UDF10);
            sbSql.AppendFormat(" FROM [TKMOC].[dbo].[PURTAB],[TK].dbo.[INVMB]");
            sbSql.AppendFormat(" WHERE [PURTAB].MB001=[INVMB].MB001");
            sbSql.AppendFormat(" AND [PURTAB].MB001 LIKE '2%'");
            sbSql.AppendFormat(" AND [PURTAB].[ID]='{0}'", ID);
            sbSql.AppendFormat(" )");

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

                UPDATEPURTA();
            }
        }

        public void UPDATEPURTA()
        {
            if (!string.IsNullOrEmpty(MOCTA001) && !string.IsNullOrEmpty(MOCTA002) && !string.IsNullOrEmpty(ID))
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                //UPDATE TB039='N'
                sbSql.AppendFormat(" UPDATE  [TK].dbo.PURTB SET TB039='N' WHERE ISNULL(TB039,'')=''");
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" UPDATE [TK].dbo.PURTB");
                sbSql.AppendFormat(" SET TB017=(SELECT TOP 1 TN008 FROM [TK].dbo.VPURTLMN WHERE  TM004=TB004 AND TL004=TB010 AND TN007<=TB009 ORDER BY TN008),TB018=ROUND((SELECT TOP 1 TN008 FROM [TK].dbo.VPURTLMN WHERE  TM004=TB004 AND TL004=TB010 AND TN007<=TB009 ORDER BY TN008)*TB009,0)");
                sbSql.AppendFormat(" FROM [TK].dbo.VPURTLMN");
                sbSql.AppendFormat(" WHERE  TL004=TB010 AND TM004=TB004");
                sbSql.AppendFormat(" AND TB001='{0}' AND TB002='{1}'", MOCTA001, MOCTA002);
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" UPDATE  [TK].dbo.PURTA");
                sbSql.AppendFormat(" SET TA011=(SELECT SUM(TB009) FROM [TK].dbo.PURTB WHERE PURTA.TA001=PURTB.TB001 AND  PURTA.TA002=PURTB.TB002)");
                sbSql.AppendFormat(" WHERE TA001='{0}' AND TA002='{1}'", MOCTA001, MOCTA002);
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" UPDATE [TKMOC].[dbo].[PURTAB]");
                sbSql.AppendFormat(" SET [PURTA001]='{0}',[PURTA002]='{1}'", MOCTA001, MOCTA002);
                sbSql.AppendFormat(" WHERE [ID]='{0}'", ID);
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

        }

        public PURTA SETPURTA()
        {
            PURTA PURTA = new PURTA();

            PURTA.COMPANY = "TK";
            PURTA.CREATOR = textBox4.Text;
            PURTA.USR_GROUP = MF004;
            PURTA.CREATE_DATE = DateTime.Now.ToString("yyyyMMdd");
            PURTA.CREATE_DATE = DateTime.Now.ToString("yyyyMMdd");
            PURTA.MODIFIER = textBox4.Text;
            PURTA.MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            PURTA.FLAG = "0";
            PURTA.CREATE_TIME = DateTime.Now.ToString("HH:mm:dd");
            PURTA.MODI_TIME = DateTime.Now.ToString("HH:mm:dd");
            PURTA.TRANS_TYPE = "P001";
            PURTA.TRANS_NAME = "PURI05";
            PURTA.sync_date = null;
            PURTA.sync_time = null;
            PURTA.sync_mark = null;
            PURTA.sync_count = null;
            PURTA.sync_count = "0";
            PURTA.DataUser = null;
            PURTA.DataGroup = null;
            PURTA.DataGroup = textBox4.Text;
            PURTA.TA001 = MOCTA001;
            PURTA.TA002 = MOCTA002;
            PURTA.TA003 = MOCTA003;
            PURTA.TA004 = MF004;
            PURTA.TA005 = ID;
            PURTA.TA006 = null;
            PURTA.TA007 = "N";
            PURTA.TA008 = "0";
            PURTA.TA009 = "9";
            PURTA.TA010 = "20";
            PURTA.TA011 = "0";
            PURTA.TA012 = textBox4.Text;
            PURTA.TA013 = MOCTA003;
            PURTA.TA014 = null;
            PURTA.TA015 = "0";
            PURTA.TA016 = "N";
            PURTA.TA017 = "0";
            PURTA.TA018 = null;
            PURTA.TA019 = null;
            PURTA.TA020 = "0";
            PURTA.TA021 = null;
            PURTA.TA022 = null;
            PURTA.TA023 = "0";
            PURTA.TA024 = "0";
            PURTA.TA025 = null;
            PURTA.TA026 = null;
            PURTA.TA027 = null;
            PURTA.TA028 = null;
            PURTA.TA029 = null;
            PURTA.TA030 = "0";
            PURTA.TA031 = null;
            PURTA.TA032 = "0";
            PURTA.TA033 = null;
            PURTA.TA034 = null;
            PURTA.TA035 = null;
            PURTA.TA036 = "0";
            PURTA.TA037 = "0";
            PURTA.TA038 = "0";
            PURTA.TA039 = "0";
            PURTA.TA040 = "0";
            PURTA.TA041 = null;
            PURTA.TA042 = null;
            PURTA.TA043 = null;
            PURTA.TA044 = null;
            PURTA.TA045 = null;
            PURTA.TA046 = null;
            PURTA.UDF01 = null;
            PURTA.UDF02 = null;
            PURTA.UDF03 = null;
            PURTA.UDF04 = null;
            PURTA.UDF05 = null;
            PURTA.UDF06 = "0";
            PURTA.UDF07 = "0";
            PURTA.UDF08 = "0";
            PURTA.UDF09 = "0";
            PURTA.UDF10 = "0";

            return PURTA;
        }


        public PURTB SETPURTB()
        {
            PURTB PURTB = new PURTB();

            PURTB.COMPANY = "TK";
            PURTB.CREATOR = "120025";
            PURTB.USR_GROUP = "103400";
            PURTB.CREATE_DATE = DateTime.Now.ToString("yyyyMMdd");
            PURTB.MODIFIER = "160115";
            PURTB.MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            PURTB.FLAG = "0";
            PURTB.CREATE_TIME = DateTime.Now.ToString("HH:mm:dd");
            PURTB.MODI_TIME = DateTime.Now.ToString("HH:mm:dd");
            PURTB.TRANS_TYPE = "P001";
            PURTB.TRANS_NAME = "PURI05";
            PURTB.sync_count = "0";
            PURTB.TB001 = MOCTA001;
            PURTB.TB002 = MOCTA002;
            PURTB.TB003 = null;
            PURTB.TB004 = null;
            PURTB.TB005 = null;
            PURTB.TB006 = null;
            PURTB.TB007 = null;
            PURTB.TB008 = null;
            PURTB.TB009 = null;
            PURTB.TB010 = null;
            PURTB.TB011 = MOCTA003;
            PURTB.TB012 = null;
            PURTB.TB013 = null;
            PURTB.TB014 = "0";
            PURTB.TB015 = null;
            PURTB.TB016 = "NTD";
            PURTB.TB017 = null;
            PURTB.TB018 = null;
            PURTB.TB019 = MOCTA003;
            PURTB.TB020 = "N";
            PURTB.TB021 = "N";
            PURTB.TB022 = null;
            PURTB.TB023 = null;
            PURTB.TB024 = null;
            PURTB.TB025 = "N";
            PURTB.TB026 = "2";
            PURTB.TB027 = null;
            PURTB.TB028 = null;
            PURTB.TB029 = null;
            PURTB.TB030 = null;
            PURTB.TB031 = null;
            PURTB.TB032 = "N";
            PURTB.TB033 = null;
            PURTB.TB034 = "0";
            PURTB.TB035 = "0";
            PURTB.TB036 = null;
            PURTB.TB037 = null;
            PURTB.TB038 = null;
            PURTB.TB039 = "N";
            PURTB.TB040 = "0";
            PURTB.TB041 = "0";
            PURTB.TB042 = null;
            PURTB.TB043 = null;
            PURTB.TB044 = null;
            PURTB.TB045 = null;
            PURTB.TB046 = null;
            PURTB.TB047 = null;
            PURTB.TB048 = null;
            PURTB.TB049 = "0";
            PURTB.TB050 = null;
            PURTB.TB051 = "0";
            PURTB.TB052 = "0";
            PURTB.TB053 = "0";
            PURTB.TB054 = null;
            PURTB.TB055 = null;
            PURTB.TB056 = null;
            PURTB.TB057 = null;
            PURTB.TB058 = "1";
            PURTB.TB059 = null;
            PURTB.TB060 = null;
            PURTB.TB061 = null;
            PURTB.TB062 = null;
            PURTB.TB063 = "0";
            PURTB.TB064 = null;
            PURTB.TB065 = null;
            PURTB.TB066 = null;
            PURTB.TB067 = "2";
            PURTB.TB068 = "0";
            PURTB.TB069 = "0";
            PURTB.TB070 = null;
            PURTB.TB071 = null;
            PURTB.TB072 = null;
            PURTB.TB073 = null;
            PURTB.TB074 = null;
            PURTB.TB075 = "0";
            PURTB.TB076 = null;
            PURTB.TB077 = "0";
            PURTB.TB078 = null;
            PURTB.TB079 = null;
            PURTB.TB080 = null;
            PURTB.TB081 = "0";
            PURTB.TB082 = "0";
            PURTB.TB083 = "0";
            PURTB.TB084 = "0";
            PURTB.TB085 = "0";
            PURTB.TB086 = null;
            PURTB.TB087 = null;
            PURTB.TB088 = "0";
            PURTB.TB089 = "1";
            PURTB.TB090 = "0";
            PURTB.TB091 = "0";
            PURTB.TB092 = "0";
            PURTB.TB093 = "0";
            PURTB.TB094 = null;
            PURTB.TB095 = null;
            PURTB.TB096 = null;
            PURTB.TB097 = null;
            PURTB.TB098 = null;
            PURTB.TB099 = null;
            PURTB.UDF01 = null;
            PURTB.UDF02 = null;
            PURTB.UDF03 = null;
            PURTB.UDF04 = null;
            PURTB.UDF05 = null;
            PURTB.UDF06 = "0";
            PURTB.UDF07 = "0";
            PURTB.UDF08 = "0";
            PURTB.UDF09 = "0";
            PURTB.UDF10 = "0";
            return PURTB;
        }

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            //if (dataGridView2.CurrentRow != null)
            //{
            //    int rowindex = dataGridView2.CurrentRow.Index;
            //    if (rowindex >= 0)
            //    {
            //        DataGridViewRow row = dataGridView2.Rows[rowindex];
            //        ID = row.Cells["批號"].Value.ToString();
            //        MOCTA003 = row.Cells["請購日"].Value.ToString();                    
            //    }
            //    else
            //    {

            //    }
            //}
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            ID = textBox1.Text;
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox4.Text))
            {
                MF004 = GETADMMF(textBox4.Text);
            }

        }
        public string GETADMMF(string MF001)
        {
            string MF004;

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds4.Clear();

                sbSql.AppendFormat(@" SELECT MF001,MF004 FROM [TK].dbo.ADMMF WHERE MF001='{0}' ", MF001);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder4 = new SqlCommandBuilder(adapter4);
                sqlConn.Open();
                ds4.Clear();
                adapter4.Fill(ds4, "ds4");
                sqlConn.Close();


                if (ds4.Tables["ds4"].Rows.Count == 0)
                {
                    textBox5.Text = null;
                    return null;
                }
                else
                {
                    if (ds4.Tables["ds4"].Rows.Count >= 1)
                    {
                        MF004 = ds4.Tables["ds4"].Rows[0]["MF004"].ToString();
                        textBox5.Text = MF004;
                        return MF004;

                    }
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

        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            string MD003;
            if (dataGridView3.CurrentRow != null)
            {
                int rowindex = dataGridView3.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView3.Rows[rowindex];
                    MD003 = row.Cells["品號"].Value.ToString().Trim();

                    SEARCHINVPURMOC(MD003, dateTimePicker2.Value.ToString("yyyyMMdd"), dateTimePicker3.Value.ToString("yyyyMMdd"));


                }
                else
                {
                    MD003 = null;
                }
            }
        }

        public void SEARCHINVPURMOC(string MD003,string SDAY,string EDAY)
        {
            DataSet ds2 = new DataSet();
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                //查每日原物料的使用量、進貨量、庫存量
                //另外查詢領料數量
                //sbSql.AppendFormat(@"  ,ISNULL((SELECT SUM(TE005) FROM [TK].dbo.MOCTE WHERE TE011+TE012 IN ( SELECT TA001+TA002 FROM [TK].dbo.MOCTA WHERE TE004=TEMP2.MD003 AND TA026=TEMP2.COPTD001 AND TA027=TEMP2.COPTD002 AND TA028=TEMP2.COPTD003 )),0)*-1 AS '領料數量' ");

                
                sbSql.AppendFormat(@"  
                                      SELECT SUM(TEMP4.TNUM) AS '預計庫存量',TEMP2.ID AS '列數',TEMP2.MANU AS '線別',TEMP2.MANUDATE AS '日期',TEMP2.MD003 AS '品號',TEMP2.MD035 AS '品名',TEMP2.TNUM AS '用量'

                                      ,TEMP2.MB004 AS '單位',TEMP2.MB001 AS '成品',TEMP2.MB002 AS '成品名',TEMP2.PACKAGE AS '成品數',TEMP2.COPTD001 AS '訂單單別',TEMP2.COPTD002 AS '訂單單號',TEMP2.COPTD003 AS '訂單序號' 
                                      FROM (
  
                                      SELECT ROW_NUMBER() OVER (ORDER BY TEMP.MANUDATE) AS ID,MANU,MANUDATE,MD003,MD035,TNUM,MB004,MB001,MB002,PACKAGE,COPTD001,COPTD002,COPTD003
                                      FROM (
  
                                      SELECT [MANU],CONVERT(NVARCHAR,[MANUDATE],112) AS MANUDATE,[MD003],[MD035],CONVERT(decimal(16,3),([PACKAGE]/[MC004]*[MD006]/[MD007]*(1+[MD008])))*-1 AS TNUM,[MB004],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[PACKAGE],[COPTD001],[COPTD002],[COPTD003]
                                      FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                                      LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                                      WHERE [MOCMANULINE].MB001=MC001
                                      AND MC001=MD001
                                      AND [MANU]='新廠包裝線'
                                      AND CONVERT(NVARCHAR,[MANUDATE],112)>='{1}' AND CONVERT(NVARCHAR,[MANUDATE],112)<='{2}'
                                      AND [MD003]='{0}'
                                      UNION 
                                      SELECT [MANU],CONVERT(NVARCHAR,[MANUDATE],112) AS MANUDATE,[MD003],[MD035],CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008])))*-1 AS TNUM,[MB004],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[NUM],[COPTD001],[COPTD002],[COPTD003]
                                      FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                                      LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                                      WHERE [MOCMANULINE].MB001=MC001
                                      AND MC001=MD001
                                      AND [MANU] NOT IN ('新廠包裝線')
                                      AND CONVERT(NVARCHAR,[MANUDATE],112)>='{1}' AND CONVERT(NVARCHAR,[MANUDATE],112)<='{2}'
                                      AND [MD003]='{0}'
                                      UNION
                                      SELECT '1進貨',TD012,TD004,MB002,CONVERT(DECIMAL(14,2),(CASE WHEN ISNULL(MD002,'')<>'' THEN (ISNULL(TD008-TD015,0)*MD004/MD003) ELSE (TD008-TD015) END )) ,MB004,NULL,NULL,NULL,TD001,TD002,TD003
                                      FROM [TK].dbo.INVMB,[TK].dbo.PURTC,[TK].dbo.PURTD 
                                      LEFT JOIN [TK].dbo.INVMD ON MD001=TD004 AND MD002=TD009  
                                      WHERE TC001=TD001 AND TC002=TD002 AND TD004=MB001 AND TD018='Y' AND TD016='N'
                                      AND TD012>='{1}' AND TD012<='{2}'
                                      AND TD004='{0}'
                                      UNION 
                                      SELECT '0庫存' AS MANU,CONVERT(NVARCHAR,GETDATE(),112) AS MANUDATE,LA001 AS MD003,MB002,SUM(LA005*LA011) TNUM, MB004,NULL AS MB001,NULL AS MB002,NULL AS PACKAGE,NULL AS COPTD001,NULL AS COPTD002,NULL AS COPTD002
                                      FROM [TK].dbo.INVLA,[TK].dbo.INVMB
                                      WHERE LA001=MB001
                                      AND  LA009 IN ('20004','20006' )
                                      AND LA001='{0}' 
                                      GROUP BY LA001,MB002,MB004
                                      UNION
                                      SELECT '1手動進出貨',CONVERT(NVARCHAR,INVPURUESD.DATES,112),INVPURUESD.MB001,MB002,NUM ,MB004,NULL,NULL,NULL,NULL,NULL,NULL
                                      FROM [TK].dbo.INVMB,[TKMOC].dbo.INVPURUESD 
                                      WHERE INVMB.MB001=INVPURUESD.MB001
                                      AND INVPURUESD.DATES>='{1}' AND INVPURUESD.DATES<='{2}'
                                      AND INVPURUESD.MB001='{0}'
  
                                      ) AS TEMP 
  
                                      ) AS TEMP2 JOIN 
                                      (SELECT ROW_NUMBER() OVER (ORDER BY TEMP3.MANUDATE) AS ID,MANU,MANUDATE,MD003,MD035,TNUM,MB004,MB001,MB002,PACKAGE,COPTD001,COPTD002,COPTD003
                                      FROM (
                                      SELECT [MANU],CONVERT(NVARCHAR,[MANUDATE],112) AS MANUDATE,[MD003],[MD035],CONVERT(decimal(16,3),([PACKAGE]/[MC004]*[MD006]/[MD007]*(1+[MD008])))*-1 AS TNUM,[MB004],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[PACKAGE],[COPTD001],[COPTD002],[COPTD003]
                                      FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                                      LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                                      WHERE [MOCMANULINE].MB001=MC001
                                      AND MC001=MD001
                                      AND [MANU]='新廠包裝線'
                                      AND CONVERT(NVARCHAR,[MANUDATE],112)>='{1}' AND CONVERT(NVARCHAR,[MANUDATE],112)<='{2}'
                                      AND [MD003]='{0}'
                                      UNION 
                                      SELECT [MANU],CONVERT(NVARCHAR,[MANUDATE],112) AS MANUDATE,[MD003],[MD035],CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008])))*-1 AS TNUM,[MB004],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[NUM],[COPTD001],[COPTD002],[COPTD003]
                                      FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                                      LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                                      WHERE [MOCMANULINE].MB001=MC001
                                      AND MC001=MD001
                                      AND [MANU] NOT IN ('新廠包裝線')
                                      AND CONVERT(NVARCHAR,[MANUDATE],112)>='{1}' AND CONVERT(NVARCHAR,[MANUDATE],112)<='{2}'
                                      AND [MD003]='{0}'
                                      UNION
                                      SELECT '1進貨',TD012,TD004,MB002,CONVERT(DECIMAL(14,2),(CASE WHEN ISNULL(MD002,'')<>'' THEN (ISNULL(TD008-TD015,0)*MD004/MD003) ELSE (TD008-TD015) END )) ,MB004,NULL,NULL,NULL,TD001,TD002,TD003
                                      FROM [TK].dbo.INVMB,[TK].dbo.PURTC,[TK].dbo.PURTD 
                                      LEFT JOIN [TK].dbo.INVMD ON MD001=TD004 AND MD002=TD009 
                                      WHERE TC001=TD001 AND TC002=TD002 AND TD004=MB001 AND TD018='Y' AND TD016='N'
                                      AND TD012>='{1}' AND TD012<='{2}' 
                                      AND TD004='{0}'
                                      UNION 
                                      SELECT '0庫存' AS MANU,CONVERT(NVARCHAR,GETDATE(),112) AS MANUDATE,LA001 AS MD003,MB002,SUM(LA005*LA011) TNUM, MB004,NULL AS MB001,NULL AS MB002,NULL AS PACKAGE,NULL AS COPTD001,NULL AS COPTD002,NULL AS COPTD002
                                      FROM [TK].dbo.INVLA,[TK].dbo.INVMB
                                      WHERE LA001=MB001
                                      AND  LA009 IN ('20004','20006' )  
                                      AND LA001='{0}' 
                                      GROUP BY LA001,MB002,MB004
                                      UNION
                                      SELECT '1手動進出貨',CONVERT(NVARCHAR,INVPURUESD.DATES,112),INVPURUESD.MB001,MB002,NUM ,MB004,NULL,NULL,NULL,NULL,NULL,NULL
                                      FROM [TK].dbo.INVMB,[TKMOC].dbo.INVPURUESD 
                                      WHERE INVMB.MB001=INVPURUESD.MB001
                                      AND INVPURUESD.DATES>='{1}' AND INVPURUESD.DATES<='{2}'
                                      AND INVPURUESD.MB001='{0}'
  
                                      ) AS TEMP3
                                      ) AS TEMP4 ON TEMP2.ID>=TEMP4.ID
                                      GROUP BY TEMP2.ID,TEMP2.MANU,TEMP2.MANUDATE,TEMP2.MD003,TEMP2.MD035,TEMP2.TNUM,TEMP2.MB004,TEMP2.MB001,TEMP2.MB002,TEMP2.PACKAGE,TEMP2.COPTD001,TEMP2.COPTD002,TEMP2.COPTD003
                                      ORDER BY TEMP2.MANUDATE, TEMP2.MANU
  
                                    ", MD003, SDAY, EDAY);

                adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                sqlConn.Open();
                ds2.Clear();
                adapter2.Fill(ds2, "ds2");
                sqlConn.Close();


                if (ds2.Tables["ds2"].Rows.Count == 0)
                {
                    dataGridView4.DataSource = null;
                }
                else
                {
                    if (ds2.Tables["ds2"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView4.DataSource = ds2.Tables["ds2"];
                        dataGridView4.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                        //根据列表中数据不同，显示不同颜色背景
                        foreach (DataGridViewRow dgRow in dataGridView4.Rows)
                        {
                            //判断
                            if (Convert.ToDecimal(dgRow.Cells[5].Value) > 0)
                            {
                                //将这行的背景色设置成Pink
                                dgRow.DefaultCellStyle.BackColor = Color.Pink;
                            }
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
        public void SEARCHDG3(string SEARCHSTRING, int INDEX)
        {
            String searchValue = SEARCHSTRING;
            rowIndexDG1 = INDEX;
            int ROWS = 0;

            for (int i = INDEX; i < dataGridView3.Rows.Count; i++)
            {
                ROWS = i;

                if (dataGridView3.Rows[i].Cells[0].Value.ToString().Contains(searchValue))
                {
                    rowIndexDG1 = i;

                    dataGridView3.CurrentRow.Selected = false;
                    dataGridView3.Rows[i].Selected = true;
                    int index = rowIndexDG1;
                    dataGridView3.FirstDisplayedScrollingRowIndex = index;

                    DataGridViewRow row = dataGridView3.Rows[index];
                    MD003 = row.Cells["品號"].Value.ToString().Trim();

                    SEARCHINVPURMOC(MD003, dateTimePicker2.Value.ToString("yyyyMMdd"), dateTimePicker3.Value.ToString("yyyyMMdd"));

                    break;
                }
                if (dataGridView3.Rows[i].Cells[1].Value.ToString().Contains(searchValue))
                {
                    rowIndexDG1 = i;

                    dataGridView3.CurrentRow.Selected = false;
                    dataGridView3.Rows[i].Selected = true;
                    int index = rowIndexDG1;
                    dataGridView3.FirstDisplayedScrollingRowIndex = index;

                    DataGridViewRow row = dataGridView3.Rows[index];
                    MD003 = row.Cells["品號"].Value.ToString().Trim();

                    SEARCHINVPURMOC(MD003, dateTimePicker2.Value.ToString("yyyyMMdd"), dateTimePicker3.Value.ToString("yyyyMMdd"));

                    break;
                }
            }

            if (ROWS == dataGridView1.Rows.Count - 1)
            {
                if (MessageBox.Show("已查到最後一筆，是否從頭開始?", "已查到最後一筆，是否從頭開始?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    SEARCHDG3(textBox2.Text.Trim(), 0);
                }
                else
                {

                }
            }
        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCHINVMC();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 0 && !string.IsNullOrEmpty(textBox1.Text) && !string.IsNullOrEmpty(textBox4.Text) && !string.IsNullOrEmpty(textBox5.Text))
            {
                ADDPURTAB(textBox1.Text);

                MOCTA001 = "A311";
                MOCTA002 = GETMAXMOCTA002(MOCTA001);

                ADDERPPURAB();

                MessageBox.Show("已完成請購單" + MOCTA001 + " " + MOCTA002);

                SETNULL();
                SEARCHPURTAB();
            }
            else
            {
                MessageBox.Show("1-查詢、2-取新批號、3-填人請人");
            }

        }
        private void button3_Click(object sender, EventArgs e)
        {

        }
        private void button4_Click(object sender, EventArgs e)
        {
            SETNULL();
            SEARCHPURTAB();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            textBox1.Text = GETMAXID();
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            SEARCHINVMC2(dateTimePicker2.Value.ToString("yyyyMMdd"), dateTimePicker3.Value.ToString("yyyyMMdd"));
        }



        private void button6_Click(object sender, EventArgs e)
        {
            if (rowIndexDG1 == -1)
            {
                SEARCHDG3(textBox2.Text.Trim(), 0);
            }
            else
            {
                SEARCHDG3(textBox2.Text.Trim(), rowIndexDG1 + 1);
            }
        }



        #endregion


    }
}
