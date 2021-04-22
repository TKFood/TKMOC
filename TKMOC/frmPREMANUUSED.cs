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
using Calendar.NET;
using Excel = Microsoft.Office.Interop.Excel;
using FastReport;
using FastReport.Data;

namespace TKMOC
{
    public partial class frmPREMANUUSED : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        SqlDataAdapter adapter2 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder2 = new SqlCommandBuilder();
        SqlDataAdapter adapter3 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder3 = new SqlCommandBuilder();
        SqlDataAdapter adapter4 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder4 = new SqlCommandBuilder();
        SqlDataAdapter adapter5 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder5 = new SqlCommandBuilder();

        SqlCommand cmd = new SqlCommand();
        SqlTransaction tran;
        int result;

        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();
        DataSet ds5 = new DataSet();

        string tablename = null;

        string MD003;
        int rowIndexDG1 = -1;
        int rowIndexDG2 = -1;

        public frmPREMANUUSED()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SEARCHMOCMANULINE()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                //過濾指定品號的原料，在TKMOC的PREMANUUSEDINVMB記錄要過濾指定品號，再用品號去找出BOM的原料品號
                //不含水麵只往BOM表找下1層  NOT IN
                //含水麵往BOM表找下2層  NOT IN
                //在預排用日期過濾
                //在少量訂單用ID過濾是否已預排
                if (comboBox3.Text.Equals("不含少量") && comboBox1.Text.Equals("含水麵") && comboBox2.Text.Equals("過濾指定品號的原料"))
                {
                    sbSql = MQUERY1(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                }
                else if (comboBox3.Text.Equals("含少量") && comboBox1.Text.Equals("含水麵") && comboBox2.Text.Equals("過濾指定品號的原料"))
                {
                    sbSql = MQUERY2(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                }
                else if (comboBox3.Text.Equals("不含少量") && comboBox1.Text.Equals("不含水麵") && comboBox2.Text.Equals("過濾指定品號的原料"))
                {
                    sbSql = MQUERY3(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                }
                else if (comboBox3.Text.Equals("含少量") && comboBox1.Text.Equals("不含水麵") && comboBox2.Text.Equals("過濾指定品號的原料"))
                {
                    sbSql = MQUERY4(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                }
                else if (comboBox3.Text.Equals("不含少量") && comboBox1.Text.Equals("含水麵") && comboBox2.Text.Equals("不過濾指定品號的原料"))
                {
                    sbSql = MQUERY5(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                }
                else if (comboBox3.Text.Equals("含少量") && comboBox1.Text.Equals("含水麵") && comboBox2.Text.Equals("不過濾指定品號的原料"))
                {
                    sbSql = MQUERY6(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                }
                else if (comboBox3.Text.Equals("不含少量") && comboBox1.Text.Equals("不含水麵") && comboBox2.Text.Equals("不過濾指定品號的原料"))
                {
                    sbSql = MQUERY7(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                }
                else if (comboBox3.Text.Equals("含少量") && comboBox1.Text.Equals("不含水麵") && comboBox2.Text.Equals("不過濾指定品號的原料"))
                {
                    sbSql = MQUERY8(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                }
                
               

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView1.DataSource = ds1.Tables["ds1"];
                        dataGridView1.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                        //根据列表中数据不同，显示不同颜色背景
                        foreach (DataGridViewRow dgRow in dataGridView1.Rows)
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

        //不含少量+含水麵+過濾指定品號的原料
        public StringBuilder MQUERY1(string SDAY,string EDAY)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@"  
                            SELECT[MD003] AS '品號',[MD035] AS '品名', SUM(TNUM) AS '需求量', MB004 AS '單位'
                            , (SELECT ISNULL(SUM(LA005 * LA011), 0) FROM[TK].dbo.INVLA WHERE  LA001 = MD003 AND LA009 IN('20004', '20005', '20006'))AS '庫存量'
                            ,(SUM(TNUM) - (SELECT ISNULL(SUM(LA005 * LA011), 0) FROM[TK].dbo.INVLA WHERE LA001 = MD003 AND LA009 IN('20004', '20005', '20006') )) AS '差異量'
                            ,(SELECT ISNULL(SUM(TD008 - TD015), 0) FROM[TK].dbo.PURTD WHERE TD004 =[MD003] AND TD018 = 'Y' AND TD016 = 'N' AND TD012>= '{0}' AND TD012<= '{1}') AS '採購未交數量'
                            FROM(
                            SELECT[MANU],[MANUDATE], TEMP.[MB001], TEMP.[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008], TNUM
                            FROM(
                            SELECT[MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            , CONVERT(decimal(16, 3), ([PACKAGE] /[MC004] *[MD006] /[MD007] * (1 +[MD008]))) AS TNUM
                            FROM[TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            WHERE[MOCMANULINE].MB001 = MC001
                            AND MC001 = MD001
                            AND[MANU] = '新廠包裝線'
                            AND[MANUDATE] >= '{0}' AND[MANUDATE] <= '{1}'
                            UNION ALL
                            SELECT[MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003], MC1.[MC001], MC1.[MC004], CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, CASE WHEN ISNULL(MD2.[MD035], '') = '' THEN MD1.[MD035] ELSE MD2.[MD035] END, MD1.[MD006], MD1.[MD007], MD1.[MD008]
                            , CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN CONVERT(decimal(16, 3), ([NUM] / MC1.[MC004] * MD1.[MD006] / MD1.[MD007] * (1 + MD1.[MD008]))) ELSE CONVERT(decimal(16, 3), ([NUM] / MC1.[MC004] * MD1.[MD006] / MD1.[MD007] * (1 + MD1.[MD008])) / MC2.[MC004] * MD2.[MD006] / MD2.[MD007] * (1 + MD2.[MD008])) END  AS TNUM
                            FROM[TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                            LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001 = MD1.MD003
                            LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001 = MD2.MD001
                            WHERE[MOCMANULINE].MB001 = MC1.MC001
                            AND MC1.MC001 = MD1.MD001
                            AND[MANU] = '新廠製一組'
                            AND[MANUDATE] >= '{0}' AND[MANUDATE] <= '{1}'
                            AND REPLACE(MC1.[MC001], ' ', '') + REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') NOT IN(SELECT REPLACE([MB001], ' ', '') + REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') FROM[TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001 = MD1.MD003 LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001 = MD2.MD001 WHERE[PREMANUUSEDINVMB].MB001 = MC1.MC001 AND MC1.MC001 = MD1.MD001 AND(REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') LIKE '1%' OR REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') LIKE '203%'))
                            UNION ALL
                            SELECT[MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003], MC1.[MC001], MC1.[MC004], CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, CASE WHEN ISNULL(MD2.[MD035], '') = '' THEN MD1.[MD035] ELSE MD2.[MD035] END, MD1.[MD006], MD1.[MD007], MD1.[MD008]
                            , CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN CONVERT(decimal(16, 3), ([NUM] / MC1.[MC004] * MD1.[MD006] / MD1.[MD007] * (1 + MD1.[MD008]))) ELSE CONVERT(decimal(16, 3), ([NUM] / MC1.[MC004] * MD1.[MD006] / MD1.[MD007] * (1 + MD1.[MD008])) / MC2.[MC004] * MD2.[MD006] / MD2.[MD007] * (1 + MD2.[MD008])) END  AS TNUM
                            FROM[TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                            LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001 = MD1.MD003
                            LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001 = MD2.MD001
                            WHERE[MOCMANULINE].MB001 = MC1.MC001
                            AND MC1.MC001 = MD1.MD001
                            AND[MANU] = '新廠製二組'
                            AND[MANUDATE] >= '{0}' AND[MANUDATE] <= '{1}'
                            AND REPLACE(MC1.[MC001], ' ', '') + REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') NOT IN(SELECT REPLACE([MB001], ' ', '') + REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') FROM[TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001 = MD1.MD003 LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001 = MD2.MD001 WHERE[PREMANUUSEDINVMB].MB001 = MC1.MC001 AND MC1.MC001 = MD1.MD001 AND(REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') LIKE '1%' OR REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') LIKE '203%'))
                            UNION ALL
                            SELECT[MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END , CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END, MD1.[MD006], MD1.[MD007], MD1.[MD008]

                            , CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END AS TNUM
                            FROM[TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                            LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003
                            LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001= MD2.MD001

                            WHERE[MOCMANULINE].MB001= MC1.MC001

                            AND MC1.MC001= MD1.MD001

                            AND[MANU]= '新廠製三組(手工)'

                            AND[MANUDATE]>='{0}' AND[MANUDATE]<='{1}'

                            AND REPLACE(MC1.[MC001],' ','')+REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') NOT IN (SELECT REPLACE([MB001],' ','')+REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') FROM[TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003 LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001= MD2.MD001 WHERE[PREMANUUSEDINVMB].MB001= MC1.MC001 AND MC1.MC001= MD1.MD001 AND  (REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') LIKE '1%' OR REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') LIKE '203%' ) )
      
                            ) AS TEMP 
                            ) AS TEMP2
                            LEFT JOIN[TK].dbo.INVMB ON INVMB.MB001=TEMP2.[MD003]
                            GROUP BY[MD003],[MD035],[MB004]
                            ORDER BY[MD003],[MD035],[MB004]
     
                            ", SDAY, EDAY);

            return SB;

        }
        //含少量+含水麵+過濾指定品號的原料
        public StringBuilder MQUERY2(string SDAY, string EDAY)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@"  
                            SELECT[MD003] AS '品號',[MD035] AS '品名', SUM(TNUM) AS '需求量', MB004 AS '單位'
                            , (SELECT ISNULL(SUM(LA005 * LA011), 0) FROM[TK].dbo.INVLA WHERE  LA001 = MD003 AND LA009 IN('20004', '20005', '20006'))AS '庫存量'
                            ,(SUM(TNUM) - (SELECT ISNULL(SUM(LA005 * LA011), 0) FROM[TK].dbo.INVLA WHERE LA001 = MD003 AND LA009 IN('20004', '20005', '20006') )) AS '差異量'
                            ,(SELECT ISNULL(SUM(TD008 - TD015), 0) FROM[TK].dbo.PURTD WHERE TD004 =[MD003] AND TD018 = 'Y' AND TD016 = 'N' AND TD012>= '{0}' AND TD012<= '{1}') AS '採購未交數量'
                            FROM(
                            SELECT[MANU],[MANUDATE], TEMP.[MB001], TEMP.[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008], TNUM
                            FROM(
                            SELECT[MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            , CONVERT(decimal(16, 3), ([PACKAGE] /[MC004] *[MD006] /[MD007] * (1 +[MD008]))) AS TNUM
                            FROM[TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            WHERE[MOCMANULINE].MB001 = MC001
                            AND MC001 = MD001
                            AND[MANU] = '新廠包裝線'
                            AND[MANUDATE] >= '{0}' AND[MANUDATE] <= '{1}'
                            UNION ALL
                            SELECT[MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003], MC1.[MC001], MC1.[MC004], CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, CASE WHEN ISNULL(MD2.[MD035], '') = '' THEN MD1.[MD035] ELSE MD2.[MD035] END, MD1.[MD006], MD1.[MD007], MD1.[MD008]
                            , CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN CONVERT(decimal(16, 3), ([NUM] / MC1.[MC004] * MD1.[MD006] / MD1.[MD007] * (1 + MD1.[MD008]))) ELSE CONVERT(decimal(16, 3), ([NUM] / MC1.[MC004] * MD1.[MD006] / MD1.[MD007] * (1 + MD1.[MD008])) / MC2.[MC004] * MD2.[MD006] / MD2.[MD007] * (1 + MD2.[MD008])) END  AS TNUM
                            FROM[TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                            LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001 = MD1.MD003
                            LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001 = MD2.MD001
                            WHERE[MOCMANULINE].MB001 = MC1.MC001
                            AND MC1.MC001 = MD1.MD001
                            AND[MANU] = '新廠製一組'
                            AND[MANUDATE] >= '{0}' AND[MANUDATE] <= '{1}'
                            AND REPLACE(MC1.[MC001], ' ', '') + REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') NOT IN(SELECT REPLACE([MB001], ' ', '') + REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') FROM[TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001 = MD1.MD003 LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001 = MD2.MD001 WHERE[PREMANUUSEDINVMB].MB001 = MC1.MC001 AND MC1.MC001 = MD1.MD001 AND(REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') LIKE '1%' OR REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') LIKE '203%'))
                            UNION ALL
                            SELECT[MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003], MC1.[MC001], MC1.[MC004], CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, CASE WHEN ISNULL(MD2.[MD035], '') = '' THEN MD1.[MD035] ELSE MD2.[MD035] END, MD1.[MD006], MD1.[MD007], MD1.[MD008]
                            , CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN CONVERT(decimal(16, 3), ([NUM] / MC1.[MC004] * MD1.[MD006] / MD1.[MD007] * (1 + MD1.[MD008]))) ELSE CONVERT(decimal(16, 3), ([NUM] / MC1.[MC004] * MD1.[MD006] / MD1.[MD007] * (1 + MD1.[MD008])) / MC2.[MC004] * MD2.[MD006] / MD2.[MD007] * (1 + MD2.[MD008])) END  AS TNUM
                            FROM[TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                            LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001 = MD1.MD003
                            LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001 = MD2.MD001
                            WHERE[MOCMANULINE].MB001 = MC1.MC001
                            AND MC1.MC001 = MD1.MD001
                            AND[MANU] = '新廠製二組'
                            AND[MANUDATE] >= '{0}' AND[MANUDATE] <= '{1}'
                            AND REPLACE(MC1.[MC001], ' ', '') + REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') NOT IN(SELECT REPLACE([MB001], ' ', '') + REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') FROM[TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001 = MD1.MD003 LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001 = MD2.MD001 WHERE[PREMANUUSEDINVMB].MB001 = MC1.MC001 AND MC1.MC001 = MD1.MD001 AND(REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') LIKE '1%' OR REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') LIKE '203%'))
                            UNION ALL
                            SELECT[MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END , CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END, MD1.[MD006], MD1.[MD007], MD1.[MD008]
                            , CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END AS TNUM
                            FROM[TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                            LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003
                            LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001= MD2.MD001
                            WHERE[MOCMANULINE].MB001= MC1.MC001
                            AND MC1.MC001= MD1.MD001
                            AND[MANU]= '新廠製三組(手工)'
                            AND[MANUDATE]>='{0}' AND[MANUDATE]<='{1}'
                            AND REPLACE(MC1.[MC001],' ','')+REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') NOT IN (SELECT REPLACE([MB001],' ','')+REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') FROM[TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003 LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001= MD2.MD001 WHERE[PREMANUUSEDINVMB].MB001= MC1.MC001 AND MC1.MC001= MD1.MD001 AND  (REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') LIKE '1%' OR REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') LIKE '203%' ) )
     
                            UNION ALL
                            SELECT [MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            , CONVERT(decimal(16, 3), ([PACKAGE] /[MC004] *[MD006] /[MD007] * (1 +[MD008]))) AS TNUM
                            FROM[TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            WHERE [MOCMANULINETEMP].MB001 = MC001
                            AND MC001 = MD001
                            AND[MANU] = '新廠包裝線'                           
                            AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE])
                            UNION ALL
                            SELECT[MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003], MC1.[MC001], MC1.[MC004],MD1.[MD003], MD1.[MD035], MD1.[MD006], MD1.[MD007], MD1.[MD008]
                            ,  CONVERT(decimal(16, 3), ([NUM] / MC1.[MC004] * MD1.[MD006] / MD1.[MD007] * (1 + MD1.[MD008]))) AS TNUM
                            FROM[TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                            WHERE [MOCMANULINETEMP].MB001 = MC1.MC001
                            AND MC1.MC001 = MD1.MD001
                            AND[MANU] = '新廠製一組'      
                            AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE])
                            AND REPLACE(MC1.[MC001], ' ', '') + REPLACE(MD1.[MD003], ' ', '') NOT IN (SELECT REPLACE([MB001], ' ', '') + REPLACE(MD1.[MD003], ' ', '') FROM[TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1  WHERE[PREMANUUSEDINVMB].MB001 = MC1.MC001 AND MC1.MC001 = MD1.MD001 AND (MD1.[MD003] LIKE '1%' OR MD1.[MD003] LIKE '203%'))    
                            UNION ALL
                            SELECT[MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003], MC1.[MC001], MC1.[MC004],MD1.[MD003], MD1.[MD035], MD1.[MD006], MD1.[MD007], MD1.[MD008]
                            ,  CONVERT(decimal(16, 3), ([NUM] / MC1.[MC004] * MD1.[MD006] / MD1.[MD007] * (1 + MD1.[MD008]))) AS TNUM
                            FROM[TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                            WHERE [MOCMANULINETEMP].MB001 = MC1.MC001
                            AND MC1.MC001 = MD1.MD001
                            AND[MANU] = '新廠製二組'      
                            AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE])
                            AND REPLACE(MC1.[MC001], ' ', '') + REPLACE(MD1.[MD003], ' ', '') NOT IN (SELECT REPLACE([MB001], ' ', '') + REPLACE(MD1.[MD003], ' ', '') FROM[TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1  WHERE[PREMANUUSEDINVMB].MB001 = MC1.MC001 AND MC1.MC001 = MD1.MD001 AND (MD1.[MD003] LIKE '1%' OR MD1.[MD003] LIKE '203%'))    
                            UNION ALL
                            SELECT[MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003], MC1.[MC001], MC1.[MC004],MD1.[MD003], MD1.[MD035], MD1.[MD006], MD1.[MD007], MD1.[MD008]
                            ,  CONVERT(decimal(16, 3), ([NUM] / MC1.[MC004] * MD1.[MD006] / MD1.[MD007] * (1 + MD1.[MD008]))) AS TNUM
                            FROM[TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                            WHERE [MOCMANULINETEMP].MB001 = MC1.MC001
                            AND MC1.MC001 = MD1.MD001
                            AND[MANU] = '新廠製三組(手工)'      
                            AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE])
                            AND REPLACE(MC1.[MC001], ' ', '') + REPLACE(MD1.[MD003], ' ', '') NOT IN (SELECT REPLACE([MB001], ' ', '') + REPLACE(MD1.[MD003], ' ', '') FROM[TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1  WHERE[PREMANUUSEDINVMB].MB001 = MC1.MC001 AND MC1.MC001 = MD1.MD001 AND (MD1.[MD003] LIKE '1%' OR MD1.[MD003] LIKE '203%'))    

                            ) AS TEMP 
                            ) AS TEMP2
                            LEFT JOIN[TK].dbo.INVMB ON INVMB.MB001=TEMP2.[MD003]
                            GROUP BY[MD003],[MD035],[MB004]
                            ORDER BY[MD003],[MD035],[MB004]
     
                            ", SDAY, EDAY);


            //SB.AppendFormat(@"  
            //                SELECT[MD003] AS '品號',[MD035] AS '品名', SUM(TNUM) AS '需求量', MB004 AS '單位'
            //                , (SELECT ISNULL(SUM(LA005 * LA011), 0) FROM[TK].dbo.INVLA WHERE  LA001 = MD003 AND LA009 IN('20004', '20005', '20006'))AS '庫存量'
            //                ,(SUM(TNUM) - (SELECT ISNULL(SUM(LA005 * LA011), 0) FROM[TK].dbo.INVLA WHERE LA001 = MD003 AND LA009 IN('20004', '20005', '20006') )) AS '差異量'
            //                ,(SELECT ISNULL(SUM(TD008 - TD015), 0) FROM[TK].dbo.PURTD WHERE TD004 =[MD003] AND TD018 = 'Y' AND TD016 = 'N' AND TD012>= '{0}' AND TD012<= '{1}') AS '採購未交數量'
            //                FROM(
            //                SELECT[MANU],[MANUDATE], TEMP.[MB001], TEMP.[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008], TNUM
            //                FROM(
            //                SELECT[MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
            //                , CONVERT(decimal(16, 3), ([PACKAGE] /[MC004] *[MD006] /[MD007] * (1 +[MD008]))) AS TNUM
            //                FROM[TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
            //                WHERE[MOCMANULINE].MB001 = MC001
            //                AND MC001 = MD001
            //                AND[MANU] = '新廠包裝線'
            //                AND[MANUDATE] >= '{0}' AND[MANUDATE] <= '{1}'
            //                UNION ALL
            //                SELECT[MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003], MC1.[MC001], MC1.[MC004], CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, CASE WHEN ISNULL(MD2.[MD035], '') = '' THEN MD1.[MD035] ELSE MD2.[MD035] END, MD1.[MD006], MD1.[MD007], MD1.[MD008]
            //                , CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN CONVERT(decimal(16, 3), ([NUM] / MC1.[MC004] * MD1.[MD006] / MD1.[MD007] * (1 + MD1.[MD008]))) ELSE CONVERT(decimal(16, 3), ([NUM] / MC1.[MC004] * MD1.[MD006] / MD1.[MD007] * (1 + MD1.[MD008])) / MC2.[MC004] * MD2.[MD006] / MD2.[MD007] * (1 + MD2.[MD008])) END  AS TNUM
            //                FROM[TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
            //                LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001 = MD1.MD003
            //                LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001 = MD2.MD001
            //                WHERE[MOCMANULINE].MB001 = MC1.MC001
            //                AND MC1.MC001 = MD1.MD001
            //                AND[MANU] = '新廠製一組'
            //                AND[MANUDATE] >= '{0}' AND[MANUDATE] <= '{1}'
            //                AND REPLACE(MC1.[MC001], ' ', '') + REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') NOT IN(SELECT REPLACE([MB001], ' ', '') + REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') FROM[TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001 = MD1.MD003 LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001 = MD2.MD001 WHERE[PREMANUUSEDINVMB].MB001 = MC1.MC001 AND MC1.MC001 = MD1.MD001 AND(REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') LIKE '1%' OR REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') LIKE '203%'))
            //                UNION ALL
            //                SELECT[MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003], MC1.[MC001], MC1.[MC004], CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, CASE WHEN ISNULL(MD2.[MD035], '') = '' THEN MD1.[MD035] ELSE MD2.[MD035] END, MD1.[MD006], MD1.[MD007], MD1.[MD008]
            //                , CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN CONVERT(decimal(16, 3), ([NUM] / MC1.[MC004] * MD1.[MD006] / MD1.[MD007] * (1 + MD1.[MD008]))) ELSE CONVERT(decimal(16, 3), ([NUM] / MC1.[MC004] * MD1.[MD006] / MD1.[MD007] * (1 + MD1.[MD008])) / MC2.[MC004] * MD2.[MD006] / MD2.[MD007] * (1 + MD2.[MD008])) END  AS TNUM
            //                FROM[TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
            //                LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001 = MD1.MD003
            //                LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001 = MD2.MD001
            //                WHERE[MOCMANULINE].MB001 = MC1.MC001
            //                AND MC1.MC001 = MD1.MD001
            //                AND[MANU] = '新廠製二組'
            //                AND[MANUDATE] >= '{0}' AND[MANUDATE] <= '{1}'
            //                AND REPLACE(MC1.[MC001], ' ', '') + REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') NOT IN(SELECT REPLACE([MB001], ' ', '') + REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') FROM[TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001 = MD1.MD003 LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001 = MD2.MD001 WHERE[PREMANUUSEDINVMB].MB001 = MC1.MC001 AND MC1.MC001 = MD1.MD001 AND(REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') LIKE '1%' OR REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') LIKE '203%'))
            //                UNION ALL
            //                SELECT[MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END , CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END, MD1.[MD006], MD1.[MD007], MD1.[MD008]
            //                , CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END AS TNUM
            //                FROM[TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
            //                LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003
            //                LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001= MD2.MD001
            //                WHERE[MOCMANULINE].MB001= MC1.MC001
            //                AND MC1.MC001= MD1.MD001
            //                AND[MANU]= '新廠製三組(手工)'
            //                AND[MANUDATE]>='{0}' AND[MANUDATE]<='{1}'
            //                AND REPLACE(MC1.[MC001],' ','')+REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') NOT IN (SELECT REPLACE([MB001],' ','')+REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') FROM[TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003 LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001= MD2.MD001 WHERE[PREMANUUSEDINVMB].MB001= MC1.MC001 AND MC1.MC001= MD1.MD001 AND  (REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') LIKE '1%' OR REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') LIKE '203%' ) )

            //                UNION ALL
            //                SELECT [MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
            //                , CONVERT(decimal(16, 3), ([PACKAGE] /[MC004] *[MD006] /[MD007] * (1 +[MD008]))) AS TNUM
            //                FROM[TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC,[TK].dbo.BOMMD
            //                WHERE [MOCMANULINETEMP].MB001 = MC001
            //                AND MC001 = MD001
            //                AND[MANU] = '新廠包裝線'                           
            //                AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE])
            //                UNION ALL
            //                SELECT[MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003], MC1.[MC001], MC1.[MC004], CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, CASE WHEN ISNULL(MD2.[MD035], '') = '' THEN MD1.[MD035] ELSE MD2.[MD035] END, MD1.[MD006], MD1.[MD007], MD1.[MD008]
            //                , CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN CONVERT(decimal(16, 3), ([NUM] / MC1.[MC004] * MD1.[MD006] / MD1.[MD007] * (1 + MD1.[MD008]))) ELSE CONVERT(decimal(16, 3), ([NUM] / MC1.[MC004] * MD1.[MD006] / MD1.[MD007] * (1 + MD1.[MD008])) / MC2.[MC004] * MD2.[MD006] / MD2.[MD007] * (1 + MD2.[MD008])) END  AS TNUM
            //                FROM[TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
            //                LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001 = MD1.MD003
            //                LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001 = MD2.MD001
            //                WHERE [MOCMANULINETEMP].MB001 = MC1.MC001
            //                AND MC1.MC001 = MD1.MD001
            //                AND[MANU] = '新廠製一組'      
            //                AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE])
            //                AND REPLACE(MC1.[MC001], ' ', '') + REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') NOT IN(SELECT REPLACE([MB001], ' ', '') + REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') FROM[TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001 = MD1.MD003 LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001 = MD2.MD001 WHERE[PREMANUUSEDINVMB].MB001 = MC1.MC001 AND MC1.MC001 = MD1.MD001 AND(REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') LIKE '1%' OR REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') LIKE '203%'))                            
            //                UNION ALL
            //                SELECT[MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003], MC1.[MC001], MC1.[MC004], CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, CASE WHEN ISNULL(MD2.[MD035], '') = '' THEN MD1.[MD035] ELSE MD2.[MD035] END, MD1.[MD006], MD1.[MD007], MD1.[MD008]
            //                , CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN CONVERT(decimal(16, 3), ([NUM] / MC1.[MC004] * MD1.[MD006] / MD1.[MD007] * (1 + MD1.[MD008]))) ELSE CONVERT(decimal(16, 3), ([NUM] / MC1.[MC004] * MD1.[MD006] / MD1.[MD007] * (1 + MD1.[MD008])) / MC2.[MC004] * MD2.[MD006] / MD2.[MD007] * (1 + MD2.[MD008])) END  AS TNUM
            //                FROM[TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
            //                LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001 = MD1.MD003
            //                LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001 = MD2.MD001
            //                WHERE [MOCMANULINETEMP].MB001 = MC1.MC001
            //                AND MC1.MC001 = MD1.MD001
            //                AND[MANU] = '新廠製二組'
            //                AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE])
            //                AND REPLACE(MC1.[MC001], ' ', '') + REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') NOT IN(SELECT REPLACE([MB001], ' ', '') + REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') FROM[TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001 = MD1.MD003 LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001 = MD2.MD001 WHERE[PREMANUUSEDINVMB].MB001 = MC1.MC001 AND MC1.MC001 = MD1.MD001 AND(REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') LIKE '1%' OR REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') LIKE '203%'))                            
            //                UNION ALL
            //                SELECT[MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END , CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END, MD1.[MD006], MD1.[MD007], MD1.[MD008]
            //                , CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END AS TNUM
            //                FROM [TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
            //                LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003
            //                LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001= MD2.MD001
            //                WHERE [MOCMANULINETEMP].MB001= MC1.MC001
            //                AND MC1.MC001= MD1.MD001
            //                AND[MANU]= '新廠製三組(手工)'
            //                AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE]) 
            //                AND REPLACE(MC1.[MC001],' ','')+REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') NOT IN (SELECT REPLACE([MB001],' ','')+REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') FROM[TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003 LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001= MD2.MD001 WHERE[PREMANUUSEDINVMB].MB001= MC1.MC001 AND MC1.MC001= MD1.MD001 AND  (REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') LIKE '1%' OR REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') LIKE '203%' ) )                             

            //                ) AS TEMP 
            //                ) AS TEMP2
            //                LEFT JOIN[TK].dbo.INVMB ON INVMB.MB001=TEMP2.[MD003]
            //                GROUP BY[MD003],[MD035],[MB004]
            //                ORDER BY[MD003],[MD035],[MB004]

            //                ", SDAY, EDAY);

            return SB;

        }
        //不含少量+不含水麵+過濾指定品號的原料
        public StringBuilder MQUERY3(string SDAY, string EDAY)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
                            SELECT [MD003] AS '品號',[MD035] AS '品名',SUM(TNUM) AS '需求量',[MB004]  AS '單位'
                            ,(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA WHERE  LA001=MD003 AND LA009 IN ('20004','20005','20006') )AS '庫存量'
                            ,(SUM(TNUM)-(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA WHERE  LA001=MD003 AND LA009 IN ('20004','20005','20006') )) AS '差異量'
                            ,(SELECT ISNULL(SUM(TD008-TD015),0) FROM [TK].dbo.PURTD WHERE TD004=[MD003] AND TD018='Y' AND TD016='N' AND TD012>='{0}' AND TD012<='{1}') AS '採購未交數量'
                            FROM (
                            SELECT [MANU],[MANUDATE],[MB001],[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008],TNUM,[MB004]
                            FROM (
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([PACKAGE]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINE].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠包裝線'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINE].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠製一組'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                            AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND  (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ) )) 
                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINE].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠製二組'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                            AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ))) 
                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINE].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠製三組(手工)'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                            AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND  (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ) )) 
                            ) AS TEMP 
                            ) AS TEMP2 
  
                            GROUP BY [MD003],[MD035],[MB004]
                            ORDER BY [MD003],[MD035],[MB004]
                            ", SDAY, EDAY);

            return SB;

        }
        //含少量+不含水麵+過濾指定品號的原料
        public StringBuilder MQUERY4(string SDAY, string EDAY)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
                            SELECT [MD003] AS '品號',[MD035] AS '品名',SUM(TNUM) AS '需求量',[MB004]  AS '單位'
                            ,(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA WHERE  LA001=MD003 AND LA009 IN ('20004','20005','20006') )AS '庫存量'
                            ,(SUM(TNUM)-(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA WHERE  LA001=MD003 AND LA009 IN ('20004','20005','20006') )) AS '差異量'
                            ,(SELECT ISNULL(SUM(TD008-TD015),0) FROM [TK].dbo.PURTD WHERE TD004=[MD003] AND TD018='Y' AND TD016='N' AND TD012>='{0}' AND TD012<='{1}') AS '採購未交數量'
                            FROM (
                            SELECT [MANU],[MANUDATE],[MB001],[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008],TNUM,[MB004]
                            FROM (
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([PACKAGE]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINE].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠包裝線'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINE].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠製一組'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                            AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND  (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ) )) 
                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINE].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠製二組'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                            AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ))) 
                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINE].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠製三組(手工)'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                            AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND  (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ) )) 

                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([PACKAGE]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINETEMP].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠包裝線'
                            AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE])
                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINETEMP].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠製一組'
                            AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE])
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                            AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND  (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ) )) 
                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINETEMP].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠製二組'
                            AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE])
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                            AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ))) 
                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINETEMP].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠製三組(手工)'
                            AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE])
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                            AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND  (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ) )) 


                            ) AS TEMP 
                            ) AS TEMP2 
  
                            GROUP BY [MD003],[MD035],[MB004]
                            ORDER BY [MD003],[MD035],[MB004]
                             ", SDAY, EDAY);

            //SB.AppendFormat(@" 
            //                SELECT [MD003] AS '品號',[MD035] AS '品名',SUM(TNUM) AS '需求量',[MB004]  AS '單位'
            //                ,(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA WHERE  LA001=MD003 AND LA009 IN ('20004','20005','20006') )AS '庫存量'
            //                ,(SUM(TNUM)-(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA WHERE  LA001=MD003 AND LA009 IN ('20004','20005','20006') )) AS '差異量'
            //                ,(SELECT ISNULL(SUM(TD008-TD015),0) FROM [TK].dbo.PURTD WHERE TD004=[MD003] AND TD018='Y' AND TD016='N' AND TD012>='{0}' AND TD012<='{1}') AS '採購未交數量'
            //                FROM (
            //                SELECT [MANU],[MANUDATE],[MB001],[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008],TNUM,[MB004]
            //                FROM (
            //                SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
            //                ,CONVERT(decimal(16,3),([PACKAGE]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
            //                ,[MB004]
            //                FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
            //                LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
            //                WHERE [MOCMANULINE].MB001=MC001
            //                AND MC001=MD001
            //                AND [MANU]='新廠包裝線'
            //                AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
            //                UNION ALL 
            //                SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
            //                ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
            //                ,[MB004]
            //                FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
            //                LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
            //                WHERE [MOCMANULINE].MB001=MC001
            //                AND MC001=MD001
            //                AND [MANU]='新廠製一組'
            //                AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
            //                AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND  (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ) )) 
            //                UNION ALL 
            //                SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
            //                ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
            //                ,[MB004]
            //                FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
            //                LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
            //                WHERE [MOCMANULINE].MB001=MC001
            //                AND MC001=MD001
            //                AND [MANU]='新廠製二組'
            //                AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
            //                AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ))) 
            //                UNION ALL 
            //                SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
            //                ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
            //                ,[MB004]
            //                FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
            //                LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
            //                WHERE [MOCMANULINE].MB001=MC001
            //                AND MC001=MD001
            //                AND [MANU]='新廠製三組(手工)'
            //                AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
            //                AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND  (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ) )) 

            //                UNION ALL 
            //                SELECT [MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
            //                ,CONVERT(decimal(16,3),([PACKAGE]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
            //                ,[MB004]
            //                FROM [TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC,[TK].dbo.BOMMD
            //                LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
            //                WHERE [MOCMANULINETEMP].MB001=MC001
            //                AND MC001=MD001
            //                AND [MANU]='新廠包裝線'
            //                AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE])
            //                UNION ALL 
            //                SELECT [MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
            //                ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
            //                ,[MB004]
            //                FROM [TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC,[TK].dbo.BOMMD
            //                LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
            //                WHERE [MOCMANULINETEMP].MB001=MC001
            //                AND MC001=MD001
            //                AND [MANU]='新廠製一組'
            //                AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE])
            //                AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND  (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ) )) 
            //                UNION ALL 
            //                SELECT [MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
            //                ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
            //                ,[MB004]
            //                FROM [TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC,[TK].dbo.BOMMD
            //                LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
            //                WHERE [MOCMANULINETEMP].MB001=MC001
            //                AND MC001=MD001
            //                AND [MANU]='新廠製二組'
            //                AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE])
            //                AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ))) 
            //                UNION ALL 
            //                SELECT [MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
            //                ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
            //                ,[MB004]
            //                FROM [TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC,[TK].dbo.BOMMD
            //                LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
            //                WHERE [MOCMANULINETEMP].MB001=MC001
            //                AND MC001=MD001
            //                AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE])
            //                AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
            //                AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND  (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ) )) 


            //                ) AS TEMP 
            //                ) AS TEMP2 

            //                GROUP BY [MD003],[MD035],[MB004]
            //                ORDER BY [MD003],[MD035],[MB004]
            //                 ", SDAY, EDAY);

            return SB;

        }
        //不含少量+含水麵+不過濾指定品號的原料
        public StringBuilder MQUERY5(string SDAY, string EDAY)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
                            SELECT [MD003] AS '品號',[MD035] AS '品名',SUM(TNUM) AS '需求量',MB004 AS '單位'
                            ,(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA WHERE  LA001=MD003 AND LA009 IN ('20004','20005','20006') )AS '庫存量'
                            ,(SUM(TNUM)-(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA WHERE  LA001=MD003 AND LA009 IN ('20004','20005','20006') )) AS '差異量'
                            ,(SELECT ISNULL(SUM(TD008-TD015),0) FROM [TK].dbo.PURTD WHERE TD004=[MD003] AND TD018='Y' AND TD016='N' AND TD012>='{0}' AND TD012<='{1}') AS '採購未交數量'
                            FROM (
                            SELECT [MANU],[MANUDATE],TEMP.[MB001],TEMP.[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008],TNUM
                            FROM (
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([PACKAGE]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            WHERE [MOCMANULINE].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠包裝線'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END,MD1.[MD006],MD1.[MD007],MD1.[MD008]
                            ,CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END  AS TNUM
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                            LEFT JOIN [TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003
                            LEFT JOIN [TK].dbo.BOMMD MD2 ON MC2.MC001=MD2.MD001
                            WHERE [MOCMANULINE].MB001=MC1.MC001
                            AND MC1.MC001=MD1.MD001
                            AND [MANU]='新廠製一組'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                    
                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END,MD1.[MD006],MD1.[MD007],MD1.[MD008]
                            ,CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END  AS TNUM
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                            LEFT JOIN [TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003
                            LEFT JOIN [TK].dbo.BOMMD MD2 ON MC2.MC001=MD2.MD001
                            WHERE [MOCMANULINE].MB001=MC1.MC001
                            AND MC1.MC001=MD1.MD001
                            AND [MANU]='新廠製二組'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                    
                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END,MD1.[MD006],MD1.[MD007],MD1.[MD008]
                            ,CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END  AS TNUM
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                            LEFT JOIN [TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003
                            LEFT JOIN [TK].dbo.BOMMD MD2 ON MC2.MC001=MD2.MD001
                            WHERE [MOCMANULINE].MB001=MC1.MC001
                            AND MC1.MC001=MD1.MD001
                            AND [MANU]='新廠製三組(手工)'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                    
                      
                            ) AS TEMP 
                            ) AS TEMP2
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=TEMP2.[MD003]
                            GROUP BY [MD003],[MD035],[MB004]
                            ORDER BY [MD003],[MD035],[MB004]
                            ", SDAY, EDAY);

            return SB;

        }
        //含少量+含水麵+不過濾指定品號的原料
        public StringBuilder MQUERY6(string SDAY, string EDAY)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
                            SELECT [MD003] AS '品號',[MD035] AS '品名',SUM(TNUM) AS '需求量',MB004 AS '單位'
                            ,(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA WHERE  LA001=MD003 AND LA009 IN ('20004','20005','20006') )AS '庫存量'
                            ,(SUM(TNUM)-(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA WHERE  LA001=MD003 AND LA009 IN ('20004','20005','20006') )) AS '差異量'
                            ,(SELECT ISNULL(SUM(TD008-TD015),0) FROM [TK].dbo.PURTD WHERE TD004=[MD003] AND TD018='Y' AND TD016='N' AND TD012>='{0}' AND TD012<='{1}') AS '採購未交數量'
                            FROM (
                            SELECT [MANU],[MANUDATE],TEMP.[MB001],TEMP.[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008],TNUM
                            FROM (
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([PACKAGE]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            WHERE [MOCMANULINE].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠包裝線'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END,MD1.[MD006],MD1.[MD007],MD1.[MD008]
                            ,CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END  AS TNUM
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                            LEFT JOIN [TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003
                            LEFT JOIN [TK].dbo.BOMMD MD2 ON MC2.MC001=MD2.MD001
                            WHERE [MOCMANULINE].MB001=MC1.MC001
                            AND MC1.MC001=MD1.MD001
                            AND [MANU]='新廠製一組'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'                    
                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END,MD1.[MD006],MD1.[MD007],MD1.[MD008]
                            ,CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END  AS TNUM
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                            LEFT JOIN [TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003
                            LEFT JOIN [TK].dbo.BOMMD MD2 ON MC2.MC001=MD2.MD001
                            WHERE [MOCMANULINE].MB001=MC1.MC001
                            AND MC1.MC001=MD1.MD001
                            AND [MANU]='新廠製二組'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'                    
                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END,MD1.[MD006],MD1.[MD007],MD1.[MD008]
                            ,CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END  AS TNUM
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                            LEFT JOIN [TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003
                            LEFT JOIN [TK].dbo.BOMMD MD2 ON MC2.MC001=MD2.MD001
                            WHERE [MOCMANULINE].MB001=MC1.MC001
                            AND MC1.MC001=MD1.MD001
                            AND [MANU]='新廠製三組(手工)'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                    
                            UNION ALL
                            SELECT [MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([PACKAGE]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            FROM [TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            WHERE [MOCMANULINETEMP].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠包裝線'
                            AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE])
                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],MD1.[MD003] ,MD1.[MD035],MD1.[MD006],MD1.[MD007],MD1.[MD008]
                            ,CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008])))  AS TNUM
                            FROM [TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                            WHERE [MOCMANULINETEMP].MB001=MC1.MC001
                            AND MC1.MC001=MD1.MD001
                            AND [MANU]='新廠製一組'
                            AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE])              
                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],MD1.[MD003] ,MD1.[MD035],MD1.[MD006],MD1.[MD007],MD1.[MD008]
                            ,CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008])))  AS TNUM
                            FROM [TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                            WHERE [MOCMANULINETEMP].MB001=MC1.MC001
                            AND MC1.MC001=MD1.MD001
                            AND [MANU]='新廠製二組'
                            AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE])                 
                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],MD1.[MD003] ,MD1.[MD035],MD1.[MD006],MD1.[MD007],MD1.[MD008]
                            ,CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008])))  AS TNUM
                            FROM [TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                            WHERE [MOCMANULINETEMP].MB001=MC1.MC001
                            AND MC1.MC001=MD1.MD001
                            AND [MANU]='新廠製三組(手工)'
                            AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE])
                      
                            ) AS TEMP 
                            ) AS TEMP2
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=TEMP2.[MD003]
                            GROUP BY [MD003],[MD035],[MB004]
                            ORDER BY [MD003],[MD035],[MB004]
                            ", SDAY, EDAY);


            //SB.AppendFormat(@" 
            //                SELECT [MD003] AS '品號',[MD035] AS '品名',SUM(TNUM) AS '需求量',MB004 AS '單位'
            //                ,(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA WHERE  LA001=MD003 AND LA009 IN ('20004','20005','20006') )AS '庫存量'
            //                ,(SUM(TNUM)-(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA WHERE  LA001=MD003 AND LA009 IN ('20004','20005','20006') )) AS '差異量'
            //                ,(SELECT ISNULL(SUM(TD008-TD015),0) FROM [TK].dbo.PURTD WHERE TD004=[MD003] AND TD018='Y' AND TD016='N' AND TD012>='{0}' AND TD012<='{1}') AS '採購未交數量'
            //                FROM (
            //                SELECT [MANU],[MANUDATE],TEMP.[MB001],TEMP.[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008],TNUM
            //                FROM (
            //                SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
            //                ,CONVERT(decimal(16,3),([PACKAGE]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
            //                FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
            //                WHERE [MOCMANULINE].MB001=MC001
            //                AND MC001=MD001
            //                AND [MANU]='新廠包裝線'
            //                AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
            //                UNION ALL 
            //                SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END,MD1.[MD006],MD1.[MD007],MD1.[MD008]
            //                ,CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END  AS TNUM
            //                FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
            //                LEFT JOIN [TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003
            //                LEFT JOIN [TK].dbo.BOMMD MD2 ON MC2.MC001=MD2.MD001
            //                WHERE [MOCMANULINE].MB001=MC1.MC001
            //                AND MC1.MC001=MD1.MD001
            //                AND [MANU]='新廠製一組'
            //                AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'                    
            //                UNION ALL 
            //                SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END,MD1.[MD006],MD1.[MD007],MD1.[MD008]
            //                ,CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END  AS TNUM
            //                FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
            //                LEFT JOIN [TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003
            //                LEFT JOIN [TK].dbo.BOMMD MD2 ON MC2.MC001=MD2.MD001
            //                WHERE [MOCMANULINE].MB001=MC1.MC001
            //                AND MC1.MC001=MD1.MD001
            //                AND [MANU]='新廠製二組'
            //                AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'                    
            //                UNION ALL 
            //                SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END,MD1.[MD006],MD1.[MD007],MD1.[MD008]
            //                ,CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END  AS TNUM
            //                FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
            //                LEFT JOIN [TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003
            //                LEFT JOIN [TK].dbo.BOMMD MD2 ON MC2.MC001=MD2.MD001
            //                WHERE [MOCMANULINE].MB001=MC1.MC001
            //                AND MC1.MC001=MD1.MD001
            //                AND [MANU]='新廠製三組(手工)'
            //                AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'

            //                UNION ALL
            //                SELECT [MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
            //                ,CONVERT(decimal(16,3),([PACKAGE]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
            //                FROM [TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC,[TK].dbo.BOMMD
            //                WHERE [MOCMANULINETEMP].MB001=MC001
            //                AND MC001=MD001
            //                AND [MANU]='新廠包裝線'
            //                AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE])
            //                UNION ALL 
            //                SELECT [MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END,MD1.[MD006],MD1.[MD007],MD1.[MD008]
            //                ,CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END  AS TNUM
            //                FROM [TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
            //                LEFT JOIN [TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003
            //                LEFT JOIN [TK].dbo.BOMMD MD2 ON MC2.MC001=MD2.MD001
            //                WHERE [MOCMANULINETEMP].MB001=MC1.MC001
            //                AND MC1.MC001=MD1.MD001
            //                AND [MANU]='新廠製一組'
            //                AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE])              
            //                UNION ALL 
            //                SELECT [MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END,MD1.[MD006],MD1.[MD007],MD1.[MD008]
            //                ,CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END  AS TNUM
            //                FROM [TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
            //                LEFT JOIN [TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003
            //                LEFT JOIN [TK].dbo.BOMMD MD2 ON MC2.MC001=MD2.MD001
            //                WHERE [MOCMANULINETEMP].MB001=MC1.MC001
            //                AND MC1.MC001=MD1.MD001
            //                AND [MANU]='新廠製二組'
            //                AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE])                 
            //                UNION ALL 
            //                SELECT [MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END,MD1.[MD006],MD1.[MD007],MD1.[MD008]
            //                ,CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END  AS TNUM
            //                FROM [TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
            //                LEFT JOIN [TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003
            //                LEFT JOIN [TK].dbo.BOMMD MD2 ON MC2.MC001=MD2.MD001
            //                WHERE [MOCMANULINETEMP].MB001=MC1.MC001
            //                AND MC1.MC001=MD1.MD001
            //                AND [MANU]='新廠製三組(手工)'
            //                AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE])

            //                ) AS TEMP 
            //                ) AS TEMP2
            //                LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=TEMP2.[MD003]
            //                GROUP BY [MD003],[MD035],[MB004]
            //                ORDER BY [MD003],[MD035],[MB004]
            //                ", SDAY, EDAY);

            return SB;

        }
        //不含少量+不含水麵+過濾指定品號的原料
        public StringBuilder MQUERY7(string SDAY, string EDAY)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@"  
                            SELECT [MD003] AS '品號',[MD035] AS '品名',SUM(TNUM) AS '需求量',[MB004]  AS '單位'
                            ,(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA WHERE  LA001=MD003 AND LA009 IN ('20004','20005','20006') )AS '庫存量'
                            ,(SUM(TNUM)-(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA WHERE  LA001=MD003 AND LA009 IN ('20004','20005','20006') )) AS '差異量'
                            ,(SELECT ISNULL(SUM(TD008-TD015),0) FROM [TK].dbo.PURTD WHERE TD004=[MD003] AND TD018='Y' AND TD016='N' AND TD012>='{0}' AND TD012<='{1}') AS '採購未交數量'
                            FROM (
                            SELECT [MANU],[MANUDATE],[MB001],[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008],TNUM,[MB004]
                            FROM (
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([PACKAGE]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINE].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠包裝線'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'

                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINE].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠製一組'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                            AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND  (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ) )) 
                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINE].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠製二組'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                            AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND  (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ) )) 
                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINE].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠製三組(手工)'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                            AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND  (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ) )) 
                            ) AS TEMP 
                            ) AS TEMP2 
  
                            GROUP BY [MD003],[MD035],[MB004]
                            ORDER BY [MD003],[MD035],[MB004]

                            ", SDAY, EDAY);


            return SB;

        }
        //不含少量+不含水麵+不過濾指定品號的原料
        public StringBuilder MQUERY8(string SDAY, string EDAY)
        {
            StringBuilder SB = new StringBuilder();
            SB.AppendFormat(@"  
                            SELECT [MD003] AS '品號',[MD035] AS '品名',SUM(TNUM) AS '需求量',[MB004]  AS '單位'
                            ,(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA WHERE  LA001=MD003 AND LA009 IN ('20004','20005','20006') )AS '庫存量'
                            ,(SUM(TNUM)-(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA WHERE  LA001=MD003 AND LA009 IN ('20004','20005','20006') )) AS '差異量'
                            ,(SELECT ISNULL(SUM(TD008-TD015),0) FROM [TK].dbo.PURTD WHERE TD004=[MD003] AND TD018='Y' AND TD016='N' AND TD012>='{0}' AND TD012<='{1}') AS '採購未交數量'
                            FROM (
                            SELECT [MANU],[MANUDATE],[MB001],[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008],TNUM,[MB004]
                            FROM (
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([PACKAGE]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINE].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠包裝線'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'

                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINE].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠製一組'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                            AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND  (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ) )) 
                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINE].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠製二組'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                            AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND  (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ) )) 
                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINE].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠製三組(手工)'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                            AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND  (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ) )) 

                            UNION ALL
                            SELECT [MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([PACKAGE]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINETEMP].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠包裝線'
                            AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE])
                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINETEMP].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠製一組'
                            AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE])
                            AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND  (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ) )) 
                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINETEMP].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠製二組'
                            AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE])
                            AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND  (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ) )) 
                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINETEMP].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠製三組(手工)'
                            AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE])
                            AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND  (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ) )) 
                            ) AS TEMP 
                            ) AS TEMP2   
                            GROUP BY [MD003],[MD035],[MB004]
                            ORDER BY [MD003],[MD035],[MB004]

                            ", SDAY, EDAY);



            return SB;

        }
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    MD003 = row.Cells["品號"].Value.ToString().Trim();

                    SEARCHMOCMANULINE2(MD003);


                }
                else
                {
                    MD003 = null;
                }
            }
        }

        public void SEARCHMOCMANULINE2(string MD003)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();
                if (comboBox3.Text.Equals("不含少量") && comboBox1.Text.Equals("含水麵") && comboBox2.Text.Equals("過濾指定品號的原料"))
                {
                    sbSql = DQUERY1(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"), MD003);
                }
                else if (comboBox3.Text.Equals("含少量") && comboBox1.Text.Equals("含水麵") && comboBox2.Text.Equals("過濾指定品號的原料"))
                {
                    sbSql = DQUERY2(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"), MD003);
                }
                else if (comboBox3.Text.Equals("不含少量") && comboBox1.Text.Equals("不含水麵") && comboBox2.Text.Equals("過濾指定品號的原料"))
                {
                    sbSql = DQUERY3(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"), MD003);
                }
                else if (comboBox3.Text.Equals("含少量") && comboBox1.Text.Equals("不含水麵") && comboBox2.Text.Equals("過濾指定品號的原料"))
                {
                    sbSql = DQUERY4(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"), MD003);
                }
                else if (comboBox3.Text.Equals("不含少量") && comboBox1.Text.Equals("含水麵") && comboBox2.Text.Equals("不過濾指定品號的原料"))
                {
                    sbSql = DQUERY5(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"), MD003);
                }
                else if (comboBox3.Text.Equals("含少量") && comboBox1.Text.Equals("含水麵") && comboBox2.Text.Equals("不過濾指定品號的原料"))
                {
                    sbSql = DQUERY6(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"), MD003);
                }
                else if (comboBox3.Text.Equals("不含少量") && comboBox1.Text.Equals("不含水麵") && comboBox2.Text.Equals("不過濾指定品號的原料"))
                {
                    sbSql = DQUERY7(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"), MD003);
                }
                else if (comboBox3.Text.Equals("含少量") && comboBox1.Text.Equals("不含水麵") && comboBox2.Text.Equals("不過濾指定品號的原料"))
                {
                    sbSql = DQUERY8(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"), MD003);
                }

                //if (comboBox1.Text.Equals("不含水麵") && comboBox2.Text.Equals("過濾指定品號的原料"))
                //{
                //    sbSql.AppendFormat(@"  
                //                        SELECT [MANU] AS '線別',CONVERT(nvarchar,[MANUDATE],112) AS '生產日',[MD003] AS '組件',[MD035] AS '組件名',TNUM AS '需求量',[MB004] AS '單位'
                //                        ,[MB001] AS '成品',[MB002] AS '成品名',[COPTD001] AS '訂單單別',[COPTD002] AS '訂單單號',[COPTD003] AS '訂單序號',[BAR] AS '桶數',[NUM] AS '數量',[PACKAGE] AS '包裝數',[MC001] AS '主件',[MC004] AS '批量',[MD006] AS '分子',[MD007] AS '分母',[MD008] AS '損秏率'
                //                        FROM (
                //                        SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                //                        ,CONVERT(decimal(16,3),([PACKAGE]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                //                        ,[MB004]
                //                        FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                //                        LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                //                        WHERE [MOCMANULINE].MB001=MC001
                //                        AND MC001=MD001
                //                        AND [MANU]='新廠包裝線'
                //                        AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                //                        UNION ALL 
                //                        SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                //                        ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                //                        ,[MB004]
                //                        FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                //                        LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                //                        WHERE [MOCMANULINE].MB001=MC001
                //                        AND MC001=MD001
                //                        AND [MANU]='新廠製一組'
                //                        AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                //                        AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND  (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ) )) 
                //                        UNION ALL 
                //                        SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                //                        ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                //                        ,[MB004]
                //                        FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                //                        LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                //                        WHERE [MOCMANULINE].MB001=MC001
                //                        AND MC001=MD001
                //                        AND [MANU]='新廠製二組'
                //                        AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                //                        AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND  (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ) )) 
                //                        UNION ALL 
                //                        SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                //                        ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                //                        ,[MB004]
                //                        FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                //                        LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                //                        WHERE [MOCMANULINE].MB001=MC001
                //                        AND MC001=MD001
                //                        AND [MANU]='新廠製三組(手工)'
                //                        AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                //                        AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND  (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ) )) 
                //                        AND MD2.[MD003]='{2}'
                //                        ) AS TEMP 
                //                        WHERE [MD003]='{3}'
                //                        ORDER BY [MANU] ,[MANUDATE],[MD003]
                //                        ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"), MD003);

                //}
                //else if (comboBox1.Text.Equals("不含水麵") && comboBox2.Text.Equals("不過濾指定品號的原料"))
                //{
                //    sbSql.AppendFormat(@"  
                //                        SELECT [MANU] AS '線別',CONVERT(nvarchar,[MANUDATE],112) AS '生產日',[MD003] AS '組件',[MD035] AS '組件名',TNUM AS '需求量',[MB004] AS '單位'
                //                        ,[MB001] AS '成品',[MB002] AS '成品名',[COPTD001] AS '訂單單別',[COPTD002] AS '訂單單號',[COPTD003] AS '訂單序號',[BAR] AS '桶數',[NUM] AS '數量',[PACKAGE] AS '包裝數',[MC001] AS '主件',[MC004] AS '批量',[MD006] AS '分子',[MD007] AS '分母',[MD008] AS '損秏率'
                //                        FROM (
                //                        SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                //                        ,CONVERT(decimal(16,3),([PACKAGE]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                //                        ,[MB004]
                //                        FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                //                        LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                //                        WHERE [MOCMANULINE].MB001=MC001
                //                        AND MC001=MD001
                //                        AND [MANU]='新廠包裝線'
                //                        AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                //                        UNION ALL 
                //                        SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                //                        ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                //                        ,[MB004]
                //                        FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                //                        LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                //                        WHERE [MOCMANULINE].MB001=MC001
                //                        AND MC001=MD001
                //                        AND [MANU]='新廠製一組'
                //                        AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'

                //                        UNION ALL 
                //                        SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                //                        ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                //                        ,[MB004]
                //                        FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                //                        LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                //                        WHERE [MOCMANULINE].MB001=MC001
                //                        AND MC001=MD001
                //                        AND [MANU]='新廠製二組'
                //                        AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'

                //                        UNION ALL 
                //                        SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                //                        ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                //                        ,[MB004]
                //                        FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                //                        LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                //                        WHERE [MOCMANULINE].MB001=MC001
                //                        AND MC001=MD001
                //                        AND [MANU]='新廠製三組(手工)'
                //                        AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'

                //                        ) AS TEMP 
                //                        WHERE [MD003]='{2}'
                //                        ORDER BY [MANU] ,[MANUDATE],[MD003]

                //                        ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"), MD003);

                //}
                //else if (comboBox1.Text.Equals("含水麵") && comboBox2.Text.Equals("過濾指定品號的原料"))
                //{
                //    sbSql.AppendFormat(@" 
                //                        SELECT [MANU] AS '線別',CONVERT(nvarchar,[MANUDATE],112) AS '生產日',[MD003] AS '組件',[MD035] AS '組件名',TNUM AS '需求量'
                //                        ,[MB001] AS '成品',[MB002] AS '成品名',[COPTD001] AS '訂單單別',[COPTD002] AS '訂單單號',[COPTD003] AS '訂單序號',[BAR] AS '桶數',[NUM] AS '數量',[PACKAGE] AS '包裝數',[MC001] AS '主件',[MC004] AS '批量',[MD006] AS '分子',[MD007] AS '分母',[MD008] AS '損秏率'
                //                        FROM (
                //                        SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                //                        ,CONVERT(decimal(16,3),([PACKAGE]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                //                        FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                //                        WHERE [MOCMANULINE].MB001=MC001
                //                        AND MC001=MD001
                //                        AND [MANU]='新廠包裝線'
                //                        AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                //                        UNION ALL 
                //                        SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END,MD1.[MD006],MD1.[MD007],MD1.[MD008]
                //                        ,CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END  AS TNUM
                //                        FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                //                        LEFT JOIN [TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003
                //                        LEFT JOIN [TK].dbo.BOMMD MD2 ON MC2.MC001=MD2.MD001
                //                        WHERE [MOCMANULINE].MB001=MC1.MC001
                //                        AND MC1.MC001=MD1.MD001
                //                        AND [MANU]='新廠製一組'
                //                        AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                //                        AND REPLACE(MC1.[MC001],' ','')+REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') NOT IN (SELECT REPLACE([MB001],' ','')+REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 LEFT JOIN [TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003 LEFT JOIN [TK].dbo.BOMMD MD2 ON MC2.MC001=MD2.MD001 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND  (REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') LIKE '1%' OR REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') LIKE '203%' ) )
                //                        UNION ALL 
                //                        SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END,MD1.[MD006],MD1.[MD007],MD1.[MD008]
                //                        ,CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END  AS TNUM
                //                        FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                //                        LEFT JOIN [TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003
                //                        LEFT JOIN [TK].dbo.BOMMD MD2 ON MC2.MC001=MD2.MD001
                //                        WHERE [MOCMANULINE].MB001=MC1.MC001
                //                        AND MC1.MC001=MD1.MD001
                //                        AND [MANU]='新廠製二組'
                //                        AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                //                        AND REPLACE(MC1.[MC001],' ','')+REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') NOT IN (SELECT REPLACE([MB001],' ','')+REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 LEFT JOIN [TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003 LEFT JOIN [TK].dbo.BOMMD MD2 ON MC2.MC001=MD2.MD001 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND  (REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') LIKE '1%' OR REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') LIKE '203%' ) )
                //                        UNION ALL 
                //                        SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END,MD1.[MD006],MD1.[MD007],MD1.[MD008]
                //                        ,CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END  AS TNUM
                //                        FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                //                        LEFT JOIN [TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003
                //                        LEFT JOIN [TK].dbo.BOMMD MD2 ON MC2.MC001=MD2.MD001
                //                        WHERE [MOCMANULINE].MB001=MC1.MC001
                //                        AND MC1.MC001=MD1.MD001
                //                        AND [MANU]='新廠製三組(手工)'
                //                        AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                //                        AND REPLACE(MC1.[MC001],' ','')+REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') NOT IN (SELECT REPLACE([MB001],' ','')+REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 LEFT JOIN [TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003 LEFT JOIN [TK].dbo.BOMMD MD2 ON MC2.MC001=MD2.MD001 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND  (REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') LIKE '1%' OR REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') LIKE '203%' ) )
                //                        AND MD2.[MD003]='{2}'
                //                        ) AS TEMP 
                //                        WHERE [MD003]='{2}'
                //                        ORDER BY [MANU] ,[MANUDATE],[MD003]
                //                        ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"), MD003);
                //}
                //else if (comboBox1.Text.Equals("含水麵") && comboBox2.Text.Equals("不過濾指定品號的原料"))
                //{
                //    sbSql.AppendFormat(@"  
                //                        SELECT [MANU] AS '線別',CONVERT(nvarchar,[MANUDATE],112) AS '生產日',[MD003] AS '組件',[MD035] AS '組件名',TNUM AS '需求量'
                //                        ,[MB001] AS '成品',[MB002] AS '成品名',[COPTD001] AS '訂單單別',[COPTD002] AS '訂單單號',[COPTD003] AS '訂單序號',[BAR] AS '桶數',[NUM] AS '數量',[PACKAGE] AS '包裝數',[MC001] AS '主件',[MC004] AS '批量',[MD006] AS '分子',[MD007] AS '分母',[MD008] AS '損秏率'
                //                        FROM (
                //                        SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                //                        ,CONVERT(decimal(16,3),([PACKAGE]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                //                        FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                //                        WHERE [MOCMANULINE].MB001=MC001
                //                        AND MC001=MD001
                //                        AND [MANU]='新廠包裝線'
                //                        AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'

                //                        UNION ALL 
                //                        SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END,MD1.[MD006],MD1.[MD007],MD1.[MD008]
                //                        ,CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END  AS TNUM
                //                        FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                //                        LEFT JOIN [TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003
                //                        LEFT JOIN [TK].dbo.BOMMD MD2 ON MC2.MC001=MD2.MD001
                //                        WHERE [MOCMANULINE].MB001=MC1.MC001
                //                        AND MC1.MC001=MD1.MD001
                //                        AND [MANU]='新廠製一組'
                //                        AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'

                //                        UNION ALL 
                //                        SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END,MD1.[MD006],MD1.[MD007],MD1.[MD008]
                //                        ,CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END  AS TNUM
                //                        FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                //                        LEFT JOIN [TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003
                //                        LEFT JOIN [TK].dbo.BOMMD MD2 ON MC2.MC001=MD2.MD001
                //                        WHERE [MOCMANULINE].MB001=MC1.MC001
                //                        AND MC1.MC001=MD1.MD001
                //                        AND [MANU]='新廠製二組'
                //                        AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'

                //                        UNION ALL 
                //                        SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END,MD1.[MD006],MD1.[MD007],MD1.[MD008]
                //                        ,CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END  AS TNUM
                //                        FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                //                        LEFT JOIN [TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003
                //                        LEFT JOIN [TK].dbo.BOMMD MD2 ON MC2.MC001=MD2.MD001
                //                        WHERE [MOCMANULINE].MB001=MC1.MC001
                //                        AND MC1.MC001=MD1.MD001
                //                        AND [MANU]='新廠製三組(手工)'
                //                        AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                //                        ) AS TEMP 
                //                        WHERE [MD003]='{2}'
                //                        ORDER BY [MANU] ,[MANUDATE],[MD003]
                //                        ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"), MD003);
                //}


                adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                sqlConn.Open();
                ds2.Clear();
                adapter2.Fill(ds2, "ds2");
                sqlConn.Close();


                if (ds2.Tables["ds2"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else
                {
                    if (ds2.Tables["ds2"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView2.DataSource = ds2.Tables["ds2"];
                        dataGridView2.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

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

        //不含少量+含水麵+過濾指定品號的原料
        public StringBuilder DQUERY1(string SDAY, string EDAY,string MD003)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@"  
                            SELECT [MANU] AS '線別',CONVERT(nvarchar,[MANUDATE],112) AS '生產日',[MD003] AS '組件',[MD035] AS '組件名',TNUM AS '需求量',[MB004] AS '單位'
                            ,TEMP2.[MB001] AS '成品',TEMP2.[MB002] AS '成品名',[COPTD001] AS '訂單單別',[COPTD002] AS '訂單單號',[COPTD003] AS '訂單序號',[BAR] AS '桶數',[NUM] AS '數量',[PACKAGE] AS '包裝數',[MC001] AS '主件',[MC004] AS '批量',[MD006] AS '分子',[MD007] AS '分母',[MD008] AS '損秏率'
                            FROM(
                            SELECT[MANU],[MANUDATE], TEMP.[MB001], TEMP.[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008], TNUM
                            FROM(
                            SELECT[MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            , CONVERT(decimal(16, 3), ([PACKAGE] /[MC004] *[MD006] /[MD007] * (1 +[MD008]))) AS TNUM
                            FROM[TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            WHERE[MOCMANULINE].MB001 = MC001
                            AND MC001 = MD001
                            AND[MANU] = '新廠包裝線'
                            AND[MANUDATE] >= '{0}' AND[MANUDATE] <= '{1}'
                            UNION ALL
                            SELECT[MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003], MC1.[MC001], MC1.[MC004], CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, CASE WHEN ISNULL(MD2.[MD035], '') = '' THEN MD1.[MD035] ELSE MD2.[MD035] END, MD1.[MD006], MD1.[MD007], MD1.[MD008]
                            , CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN CONVERT(decimal(16, 3), ([NUM] / MC1.[MC004] * MD1.[MD006] / MD1.[MD007] * (1 + MD1.[MD008]))) ELSE CONVERT(decimal(16, 3), ([NUM] / MC1.[MC004] * MD1.[MD006] / MD1.[MD007] * (1 + MD1.[MD008])) / MC2.[MC004] * MD2.[MD006] / MD2.[MD007] * (1 + MD2.[MD008])) END  AS TNUM
                            FROM[TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                            LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001 = MD1.MD003
                            LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001 = MD2.MD001
                            WHERE[MOCMANULINE].MB001 = MC1.MC001
                            AND MC1.MC001 = MD1.MD001
                            AND[MANU] = '新廠製一組'
                            AND[MANUDATE] >= '{0}' AND[MANUDATE] <= '{1}'
                            AND REPLACE(MC1.[MC001], ' ', '') + REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') NOT IN(SELECT REPLACE([MB001], ' ', '') + REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') FROM[TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001 = MD1.MD003 LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001 = MD2.MD001 WHERE[PREMANUUSEDINVMB].MB001 = MC1.MC001 AND MC1.MC001 = MD1.MD001 AND(REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') LIKE '1%' OR REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') LIKE '203%'))
                            UNION ALL
                            SELECT[MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003], MC1.[MC001], MC1.[MC004], CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, CASE WHEN ISNULL(MD2.[MD035], '') = '' THEN MD1.[MD035] ELSE MD2.[MD035] END, MD1.[MD006], MD1.[MD007], MD1.[MD008]
                            , CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN CONVERT(decimal(16, 3), ([NUM] / MC1.[MC004] * MD1.[MD006] / MD1.[MD007] * (1 + MD1.[MD008]))) ELSE CONVERT(decimal(16, 3), ([NUM] / MC1.[MC004] * MD1.[MD006] / MD1.[MD007] * (1 + MD1.[MD008])) / MC2.[MC004] * MD2.[MD006] / MD2.[MD007] * (1 + MD2.[MD008])) END  AS TNUM
                            FROM[TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                            LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001 = MD1.MD003
                            LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001 = MD2.MD001
                            WHERE[MOCMANULINE].MB001 = MC1.MC001
                            AND MC1.MC001 = MD1.MD001
                            AND[MANU] = '新廠製二組'
                            AND[MANUDATE] >= '{0}' AND[MANUDATE] <= '{1}'
                            AND REPLACE(MC1.[MC001], ' ', '') + REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') NOT IN(SELECT REPLACE([MB001], ' ', '') + REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') FROM[TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001 = MD1.MD003 LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001 = MD2.MD001 WHERE[PREMANUUSEDINVMB].MB001 = MC1.MC001 AND MC1.MC001 = MD1.MD001 AND(REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') LIKE '1%' OR REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') LIKE '203%'))
                            UNION ALL
                            SELECT[MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END , CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END, MD1.[MD006], MD1.[MD007], MD1.[MD008]

                            , CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END AS TNUM
                            FROM[TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                            LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003
                            LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001= MD2.MD001

                            WHERE[MOCMANULINE].MB001= MC1.MC001

                            AND MC1.MC001= MD1.MD001

                            AND[MANU]= '新廠製三組(手工)'

                            AND[MANUDATE]>='{0}' AND[MANUDATE]<='{1}'

                            AND REPLACE(MC1.[MC001],' ','')+REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') NOT IN (SELECT REPLACE([MB001],' ','')+REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') FROM[TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003 LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001= MD2.MD001 WHERE[PREMANUUSEDINVMB].MB001= MC1.MC001 AND MC1.MC001= MD1.MD001 AND  (REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') LIKE '1%' OR REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') LIKE '203%' ) )
                            AND (MD2.[MD003]='{2}' OR MD1.[MD003]='{2}' )
                            ) AS TEMP 
                            ) AS TEMP2
                            LEFT JOIN[TK].dbo.INVMB ON INVMB.MB001=TEMP2.[MD003]
                            WHERE [MD003]='{2}'
     
                            ", SDAY, EDAY, MD003);

            return SB;

        }
        //含少量+含水麵+過濾指定品號的原料
        public StringBuilder DQUERY2(string SDAY, string EDAY, string MD003)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@"  
                            SELECT [MANU] AS '線別',CONVERT(nvarchar,[MANUDATE],112) AS '生產日',[MD003] AS '組件',[MD035] AS '組件名',TNUM AS '需求量',[MB004] AS '單位'
                            ,TEMP2.[MB001] AS '成品',TEMP2.[MB002] AS '成品名',[COPTD001] AS '訂單單別',[COPTD002] AS '訂單單號',[COPTD003] AS '訂單序號',[BAR] AS '桶數',[NUM] AS '數量',[PACKAGE] AS '包裝數',[MC001] AS '主件',[MC004] AS '批量',[MD006] AS '分子',[MD007] AS '分母',[MD008] AS '損秏率'
                            FROM(
                            SELECT[MANU],[MANUDATE], TEMP.[MB001], TEMP.[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008], TNUM
                            FROM(
                            SELECT[MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            , CONVERT(decimal(16, 3), ([PACKAGE] /[MC004] *[MD006] /[MD007] * (1 +[MD008]))) AS TNUM
                            FROM[TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            WHERE[MOCMANULINE].MB001 = MC001
                            AND MC001 = MD001
                            AND[MANU] = '新廠包裝線'
                            AND[MANUDATE] >= '{0}' AND[MANUDATE] <= '{1}'
                            UNION ALL
                            SELECT[MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003], MC1.[MC001], MC1.[MC004], CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, CASE WHEN ISNULL(MD2.[MD035], '') = '' THEN MD1.[MD035] ELSE MD2.[MD035] END, MD1.[MD006], MD1.[MD007], MD1.[MD008]
                            , CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN CONVERT(decimal(16, 3), ([NUM] / MC1.[MC004] * MD1.[MD006] / MD1.[MD007] * (1 + MD1.[MD008]))) ELSE CONVERT(decimal(16, 3), ([NUM] / MC1.[MC004] * MD1.[MD006] / MD1.[MD007] * (1 + MD1.[MD008])) / MC2.[MC004] * MD2.[MD006] / MD2.[MD007] * (1 + MD2.[MD008])) END  AS TNUM
                            FROM[TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                            LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001 = MD1.MD003
                            LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001 = MD2.MD001
                            WHERE[MOCMANULINE].MB001 = MC1.MC001
                            AND MC1.MC001 = MD1.MD001
                            AND[MANU] = '新廠製一組'
                            AND[MANUDATE] >= '{0}' AND[MANUDATE] <= '{1}'
                            AND REPLACE(MC1.[MC001], ' ', '') + REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') NOT IN(SELECT REPLACE([MB001], ' ', '') + REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') FROM[TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001 = MD1.MD003 LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001 = MD2.MD001 WHERE[PREMANUUSEDINVMB].MB001 = MC1.MC001 AND MC1.MC001 = MD1.MD001 AND(REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') LIKE '1%' OR REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') LIKE '203%'))
                            UNION ALL
                            SELECT[MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003], MC1.[MC001], MC1.[MC004], CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, CASE WHEN ISNULL(MD2.[MD035], '') = '' THEN MD1.[MD035] ELSE MD2.[MD035] END, MD1.[MD006], MD1.[MD007], MD1.[MD008]
                            , CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN CONVERT(decimal(16, 3), ([NUM] / MC1.[MC004] * MD1.[MD006] / MD1.[MD007] * (1 + MD1.[MD008]))) ELSE CONVERT(decimal(16, 3), ([NUM] / MC1.[MC004] * MD1.[MD006] / MD1.[MD007] * (1 + MD1.[MD008])) / MC2.[MC004] * MD2.[MD006] / MD2.[MD007] * (1 + MD2.[MD008])) END  AS TNUM
                            FROM[TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                            LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001 = MD1.MD003
                            LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001 = MD2.MD001
                            WHERE[MOCMANULINE].MB001 = MC1.MC001
                            AND MC1.MC001 = MD1.MD001
                            AND[MANU] = '新廠製二組'
                            AND[MANUDATE] >= '{0}' AND[MANUDATE] <= '{1}'
                            AND REPLACE(MC1.[MC001], ' ', '') + REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') NOT IN(SELECT REPLACE([MB001], ' ', '') + REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') FROM[TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001 = MD1.MD003 LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001 = MD2.MD001 WHERE[PREMANUUSEDINVMB].MB001 = MC1.MC001 AND MC1.MC001 = MD1.MD001 AND(REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') LIKE '1%' OR REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') LIKE '203%'))
                            UNION ALL
                            SELECT[MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END , CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END, MD1.[MD006], MD1.[MD007], MD1.[MD008]
                            , CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END AS TNUM
                            FROM[TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                            LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003
                            LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001= MD2.MD001
                            WHERE[MOCMANULINE].MB001= MC1.MC001
                            AND MC1.MC001= MD1.MD001
                            AND[MANU]= '新廠製三組(手工)'
                            AND[MANUDATE]>='{0}' AND[MANUDATE]<='{1}'
                            AND REPLACE(MC1.[MC001],' ','')+REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') NOT IN (SELECT REPLACE([MB001],' ','')+REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') FROM[TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003 LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001= MD2.MD001 WHERE[PREMANUUSEDINVMB].MB001= MC1.MC001 AND MC1.MC001= MD1.MD001 AND  (REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') LIKE '1%' OR REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') LIKE '203%' ) )
     
                            UNION ALL
                            SELECT '少量'+[MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            , CONVERT(decimal(16, 3), ([PACKAGE] /[MC004] *[MD006] /[MD007] * (1 +[MD008]))) AS TNUM
                            FROM[TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            WHERE [MOCMANULINETEMP].MB001 = MC001
                            AND MC001 = MD001
                            AND[MANU] = '新廠包裝線'                           
                            AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE])
                            UNION ALL
                            SELECT '少量'+[MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003], MC1.[MC001], MC1.[MC004], CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, CASE WHEN ISNULL(MD2.[MD035], '') = '' THEN MD1.[MD035] ELSE MD2.[MD035] END, MD1.[MD006], MD1.[MD007], MD1.[MD008]
                            , CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN CONVERT(decimal(16, 3), ([NUM] / MC1.[MC004] * MD1.[MD006] / MD1.[MD007] * (1 + MD1.[MD008]))) ELSE CONVERT(decimal(16, 3), ([NUM] / MC1.[MC004] * MD1.[MD006] / MD1.[MD007] * (1 + MD1.[MD008])) / MC2.[MC004] * MD2.[MD006] / MD2.[MD007] * (1 + MD2.[MD008])) END  AS TNUM
                            FROM[TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                            LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001 = MD1.MD003
                            LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001 = MD2.MD001
                            WHERE [MOCMANULINETEMP].MB001 = MC1.MC001
                            AND MC1.MC001 = MD1.MD001
                            AND[MANU] = '新廠製一組'      
                            AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE])
                            AND REPLACE(MC1.[MC001], ' ', '') + REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') NOT IN(SELECT REPLACE([MB001], ' ', '') + REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') FROM[TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001 = MD1.MD003 LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001 = MD2.MD001 WHERE[PREMANUUSEDINVMB].MB001 = MC1.MC001 AND MC1.MC001 = MD1.MD001 AND(REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') LIKE '1%' OR REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') LIKE '203%'))                            
                            UNION ALL
                            SELECT '少量'+[MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003], MC1.[MC001], MC1.[MC004], CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, CASE WHEN ISNULL(MD2.[MD035], '') = '' THEN MD1.[MD035] ELSE MD2.[MD035] END, MD1.[MD006], MD1.[MD007], MD1.[MD008]
                            , CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN CONVERT(decimal(16, 3), ([NUM] / MC1.[MC004] * MD1.[MD006] / MD1.[MD007] * (1 + MD1.[MD008]))) ELSE CONVERT(decimal(16, 3), ([NUM] / MC1.[MC004] * MD1.[MD006] / MD1.[MD007] * (1 + MD1.[MD008])) / MC2.[MC004] * MD2.[MD006] / MD2.[MD007] * (1 + MD2.[MD008])) END  AS TNUM
                            FROM[TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                            LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001 = MD1.MD003
                            LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001 = MD2.MD001
                            WHERE [MOCMANULINETEMP].MB001 = MC1.MC001
                            AND MC1.MC001 = MD1.MD001
                            AND[MANU] = '新廠製二組'
                            AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE])
                            AND REPLACE(MC1.[MC001], ' ', '') + REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') NOT IN(SELECT REPLACE([MB001], ' ', '') + REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') FROM[TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001 = MD1.MD003 LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001 = MD2.MD001 WHERE[PREMANUUSEDINVMB].MB001 = MC1.MC001 AND MC1.MC001 = MD1.MD001 AND(REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') LIKE '1%' OR REPLACE(CASE WHEN ISNULL(MD2.[MD003], '') = '' THEN MD1.[MD003] ELSE MD2.[MD003] END, ' ', '') LIKE '203%'))                            
                            UNION ALL
                            SELECT '少量'+[MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END , CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END, MD1.[MD006], MD1.[MD007], MD1.[MD008]
                            , CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END AS TNUM
                            FROM [TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                            LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003
                            LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001= MD2.MD001
                            WHERE [MOCMANULINETEMP].MB001= MC1.MC001
                            AND MC1.MC001= MD1.MD001
                            AND[MANU]= '新廠製三組(手工)'
                            AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE]) 
                            AND REPLACE(MC1.[MC001],' ','')+REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') NOT IN (SELECT REPLACE([MB001],' ','')+REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') FROM[TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 LEFT JOIN[TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003 LEFT JOIN[TK].dbo.BOMMD MD2 ON MC2.MC001= MD2.MD001 WHERE[PREMANUUSEDINVMB].MB001= MC1.MC001 AND MC1.MC001= MD1.MD001 AND  (REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') LIKE '1%' OR REPLACE(CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,' ','') LIKE '203%' ) )                             
                             AND (MD2.[MD003]='{2}' OR MD1.[MD003]='{2}' )
                            ) AS TEMP 
                            ) AS TEMP2

                            LEFT JOIN[TK].dbo.INVMB ON INVMB.MB001=TEMP2.[MD003]
                            WHERE [MD003]='{2}'
     
                            ", SDAY, EDAY, MD003);

            return SB;

        }
        //不含少量+不含水麵+過濾指定品號的原料
        public StringBuilder DQUERY3(string SDAY, string EDAY, string MD003)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
                            SELECT [MANU] AS '線別',CONVERT(nvarchar,[MANUDATE],112) AS '生產日',[MD003] AS '組件',[MD035] AS '組件名',TNUM AS '需求量',[MB004] AS '單位'
                            ,TEMP2.[MB001] AS '成品',TEMP2.[MB002] AS '成品名',[COPTD001] AS '訂單單別',[COPTD002] AS '訂單單號',[COPTD003] AS '訂單序號',[BAR] AS '桶數',[NUM] AS '數量',[PACKAGE] AS '包裝數',[MC001] AS '主件',[MC004] AS '批量',[MD006] AS '分子',[MD007] AS '分母',[MD008] AS '損秏率'
                             FROM (
                            SELECT [MANU],[MANUDATE],[MB001],[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008],TNUM,[MB004]
                            FROM (
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([PACKAGE]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINE].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠包裝線'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINE].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠製一組'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                            AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND  (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ) )) 
                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINE].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠製二組'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                            AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ))) 
                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINE].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠製三組(手工)'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                            AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND  (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ) )) 
                            ) AS TEMP 
                            ) AS TEMP2 
                            WHERE [MD003]='{2}'
                            ", SDAY, EDAY, MD003);

            return SB;

        }
        //含少量+不含水麵+過濾指定品號的原料
        public StringBuilder DQUERY4(string SDAY, string EDAY, string MD003)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
                            SELECT [MANU] AS '線別',CONVERT(nvarchar,[MANUDATE],112) AS '生產日',[MD003] AS '組件',[MD035] AS '組件名',TNUM AS '需求量',[MB004] AS '單位'
                            ,TEMP2.[MB001] AS '成品',TEMP2.[MB002] AS '成品名',[COPTD001] AS '訂單單別',[COPTD002] AS '訂單單號',[COPTD003] AS '訂單序號',[BAR] AS '桶數',[NUM] AS '數量',[PACKAGE] AS '包裝數',[MC001] AS '主件',[MC004] AS '批量',[MD006] AS '分子',[MD007] AS '分母',[MD008] AS '損秏率'
                             FROM (
                            SELECT [MANU],[MANUDATE],[MB001],[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008],TNUM,[MB004]
                            FROM (
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([PACKAGE]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINE].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠包裝線'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINE].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠製一組'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                            AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND  (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ) )) 
                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINE].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠製二組'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                            AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ))) 
                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINE].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠製三組(手工)'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                            AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND  (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ) )) 

                            UNION ALL 
                            SELECT '少量'+[MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([PACKAGE]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINETEMP].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠包裝線'
                            AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE])
                            UNION ALL 
                            SELECT '少量'+[MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINETEMP].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠製一組'
                            AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE])
                            AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND  (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ) )) 
                            UNION ALL 
                            SELECT '少量'+[MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINETEMP].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠製二組'
                            AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE])
                            AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ))) 
                            UNION ALL 
                            SELECT '少量'+[MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINETEMP].MB001=MC001
                            AND MC001=MD001
                            AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE])
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                            AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND  (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ) )) 


                            ) AS TEMP 
                            ) AS TEMP2 
                            WHERE [MD003]='{2}'
                             ", SDAY, EDAY, MD003);

            return SB;

        }
        //不含少量+含水麵+不過濾指定品號的原料
        public StringBuilder DQUERY5(string SDAY, string EDAY, string MD003)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
                            SELECT [MANU] AS '線別',CONVERT(nvarchar,[MANUDATE],112) AS '生產日',[MD003] AS '組件',[MD035] AS '組件名',TNUM AS '需求量',[MB004] AS '單位'
                            ,TEMP2.[MB001] AS '成品',TEMP2.[MB002] AS '成品名',[COPTD001] AS '訂單單別',[COPTD002] AS '訂單單號',[COPTD003] AS '訂單序號',[BAR] AS '桶數',[NUM] AS '數量',[PACKAGE] AS '包裝數',[MC001] AS '主件',[MC004] AS '批量',[MD006] AS '分子',[MD007] AS '分母',[MD008] AS '損秏率'
                             FROM (
                            SELECT [MANU],[MANUDATE],TEMP.[MB001],TEMP.[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008],TNUM
                            FROM (
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([PACKAGE]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            WHERE [MOCMANULINE].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠包裝線'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END,MD1.[MD006],MD1.[MD007],MD1.[MD008]
                            ,CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END  AS TNUM
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                            LEFT JOIN [TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003
                            LEFT JOIN [TK].dbo.BOMMD MD2 ON MC2.MC001=MD2.MD001
                            WHERE [MOCMANULINE].MB001=MC1.MC001
                            AND MC1.MC001=MD1.MD001
                            AND [MANU]='新廠製一組'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                    
                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END,MD1.[MD006],MD1.[MD007],MD1.[MD008]
                            ,CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END  AS TNUM
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                            LEFT JOIN [TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003
                            LEFT JOIN [TK].dbo.BOMMD MD2 ON MC2.MC001=MD2.MD001
                            WHERE [MOCMANULINE].MB001=MC1.MC001
                            AND MC1.MC001=MD1.MD001
                            AND [MANU]='新廠製二組'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                    
                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END,MD1.[MD006],MD1.[MD007],MD1.[MD008]
                            ,CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END  AS TNUM
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                            LEFT JOIN [TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003
                            LEFT JOIN [TK].dbo.BOMMD MD2 ON MC2.MC001=MD2.MD001
                            WHERE [MOCMANULINE].MB001=MC1.MC001
                            AND MC1.MC001=MD1.MD001
                            AND [MANU]='新廠製三組(手工)'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                    
                      
                            ) AS TEMP 
                            ) AS TEMP2

                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=TEMP2.[MD003]
                            WHERE [MD003]='{2}'
                            ", SDAY, EDAY, MD003);

            return SB;

        }
        //含少量+含水麵+不過濾指定品號的原料
        public StringBuilder DQUERY6(string SDAY, string EDAY, string MD003)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
                            SELECT [MANU] AS '線別',CONVERT(nvarchar,[MANUDATE],112) AS '生產日',[MD003] AS '組件',[MD035] AS '組件名',TNUM AS '需求量',[MB004] AS '單位'
                            ,TEMP2.[MB001] AS '成品',TEMP2.[MB002] AS '成品名',[COPTD001] AS '訂單單別',[COPTD002] AS '訂單單號',[COPTD003] AS '訂單序號',[BAR] AS '桶數',[NUM] AS '數量',[PACKAGE] AS '包裝數',[MC001] AS '主件',[MC004] AS '批量',[MD006] AS '分子',[MD007] AS '分母',[MD008] AS '損秏率'
                             FROM (
                            SELECT [MANU],[MANUDATE],TEMP.[MB001],TEMP.[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008],TNUM
                            FROM (
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([PACKAGE]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            WHERE [MOCMANULINE].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠包裝線'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END,MD1.[MD006],MD1.[MD007],MD1.[MD008]
                            ,CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END  AS TNUM
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                            LEFT JOIN [TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003
                            LEFT JOIN [TK].dbo.BOMMD MD2 ON MC2.MC001=MD2.MD001
                            WHERE [MOCMANULINE].MB001=MC1.MC001
                            AND MC1.MC001=MD1.MD001
                            AND [MANU]='新廠製一組'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'                    
                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END,MD1.[MD006],MD1.[MD007],MD1.[MD008]
                            ,CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END  AS TNUM
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                            LEFT JOIN [TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003
                            LEFT JOIN [TK].dbo.BOMMD MD2 ON MC2.MC001=MD2.MD001
                            WHERE [MOCMANULINE].MB001=MC1.MC001
                            AND MC1.MC001=MD1.MD001
                            AND [MANU]='新廠製二組'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'                    
                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END,MD1.[MD006],MD1.[MD007],MD1.[MD008]
                            ,CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END  AS TNUM
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                            LEFT JOIN [TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003
                            LEFT JOIN [TK].dbo.BOMMD MD2 ON MC2.MC001=MD2.MD001
                            WHERE [MOCMANULINE].MB001=MC1.MC001
                            AND MC1.MC001=MD1.MD001
                            AND [MANU]='新廠製三組(手工)'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                    
                            UNION ALL
                            SELECT '少量'+[MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([PACKAGE]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            FROM [TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            WHERE [MOCMANULINETEMP].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠包裝線'
                            AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE])
                            UNION ALL 
                            SELECT '少量'+[MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END,MD1.[MD006],MD1.[MD007],MD1.[MD008]
                            ,CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END  AS TNUM
                            FROM [TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                            LEFT JOIN [TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003
                            LEFT JOIN [TK].dbo.BOMMD MD2 ON MC2.MC001=MD2.MD001
                            WHERE [MOCMANULINETEMP].MB001=MC1.MC001
                            AND MC1.MC001=MD1.MD001
                            AND [MANU]='新廠製一組'
                            AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE])              
                            UNION ALL 
                            SELECT '少量'+[MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END,MD1.[MD006],MD1.[MD007],MD1.[MD008]
                            ,CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END  AS TNUM
                            FROM [TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                            LEFT JOIN [TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003
                            LEFT JOIN [TK].dbo.BOMMD MD2 ON MC2.MC001=MD2.MD001
                            WHERE [MOCMANULINETEMP].MB001=MC1.MC001
                            AND MC1.MC001=MD1.MD001
                            AND [MANU]='新廠製二組'
                            AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE])                 
                            UNION ALL 
                            SELECT '少量'+[MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],MC1.[MC001],MC1.[MC004],CASE WHEN ISNULL(MD2.[MD003],'')='' THEN MD1.[MD003] ELSE MD2.[MD003] END ,CASE WHEN ISNULL(MD2.[MD035],'')='' THEN MD1.[MD035] ELSE MD2.[MD035] END,MD1.[MD006],MD1.[MD007],MD1.[MD008]
                            ,CASE WHEN ISNULL(MD2.[MD003],'')='' THEN CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))) ELSE CONVERT(decimal(16,3),([NUM]/MC1.[MC004]*MD1.[MD006]/MD1.[MD007]*(1+MD1.[MD008]))/MC2.[MC004]*MD2.[MD006]/MD2.[MD007]*(1+MD2.[MD008]) ) END  AS TNUM
                            FROM [TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1
                            LEFT JOIN [TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003
                            LEFT JOIN [TK].dbo.BOMMD MD2 ON MC2.MC001=MD2.MD001
                            WHERE [MOCMANULINETEMP].MB001=MC1.MC001
                            AND MC1.MC001=MD1.MD001
                            AND [MANU]='新廠製三組(手工)'
                            AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE])
                      
                            ) AS TEMP 
                            ) AS TEMP2
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=TEMP2.[MD003]
                            WHERE [MD003]='{2}'
                            ", SDAY, EDAY, MD003);

            return SB;

        }
        //不含少量+不含水麵+過濾指定品號的原料
        public StringBuilder DQUERY7(string SDAY, string EDAY, string MD003)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@"  
                             SELECT [MANU] AS '線別',CONVERT(nvarchar,[MANUDATE],112) AS '生產日',[MD003] AS '組件',[MD035] AS '組件名',TNUM AS '需求量',[MB004] AS '單位'
                            ,TEMP2.[MB001] AS '成品',TEMP2.[MB002] AS '成品名',[COPTD001] AS '訂單單別',[COPTD002] AS '訂單單號',[COPTD003] AS '訂單序號',[BAR] AS '桶數',[NUM] AS '數量',[PACKAGE] AS '包裝數',[MC001] AS '主件',[MC004] AS '批量',[MD006] AS '分子',[MD007] AS '分母',[MD008] AS '損秏率'
                             FROM (
                            SELECT [MANU],[MANUDATE],[MB001],[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008],TNUM,[MB004]
                            FROM (
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([PACKAGE]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINE].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠包裝線'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'

                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINE].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠製一組'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                            AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND  (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ) )) 
                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINE].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠製二組'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                            AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND  (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ) )) 
                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINE].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠製三組(手工)'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                            AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND  (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ) )) 
                            ) AS TEMP 
                            ) AS TEMP2 
                            WHERE [MD003]='{2}'
                            ", SDAY, EDAY, MD003);


            return SB;

        }
        //含少量+不含水麵+不過濾指定品號的原料
        public StringBuilder DQUERY8(string SDAY, string EDAY, string MD003)
        {
            StringBuilder SB = new StringBuilder();
            SB.AppendFormat(@"  
                            SELECT [MANU] AS '線別',CONVERT(nvarchar,[MANUDATE],112) AS '生產日',[MD003] AS '組件',[MD035] AS '組件名',TNUM AS '需求量',[MB004] AS '單位'
                            ,TEMP2.[MB001] AS '成品',TEMP2.[MB002] AS '成品名',[COPTD001] AS '訂單單別',[COPTD002] AS '訂單單號',[COPTD003] AS '訂單序號',[BAR] AS '桶數',[NUM] AS '數量',[PACKAGE] AS '包裝數',[MC001] AS '主件',[MC004] AS '批量',[MD006] AS '分子',[MD007] AS '分母',[MD008] AS '損秏率'
                             FROM (
                            SELECT [MANU],[MANUDATE],[MB001],[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008],TNUM,[MB004]
                            FROM (
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([PACKAGE]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINE].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠包裝線'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'

                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINE].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠製一組'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                            AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND  (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ) )) 
                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINE].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠製二組'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                            AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND  (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ) )) 
                            UNION ALL 
                            SELECT [MANU],[MANUDATE],[MOCMANULINE].[MB001],[MOCMANULINE].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINE],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINE].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠製三組(手工)'
                            AND [MANUDATE]>='{0}' AND [MANUDATE]<='{1}'
                            AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND  (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ) )) 

                            UNION ALL
                            SELECT '少量'+[MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([PACKAGE]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINETEMP].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠包裝線'
                            AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE])
                            UNION ALL 
                            SELECT '少量'+[MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINETEMP].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠製一組'
                            AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE])
                            AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND  (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ) )) 
                            UNION ALL 
                            SELECT '少量'+[MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINETEMP].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠製二組'
                            AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE])
                            AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND  (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ) )) 
                            UNION ALL 
                            SELECT '少量'+[MANU],[MANUDATE],[MOCMANULINETEMP].[MB001],[MOCMANULINETEMP].[MB002],[BAR],[NUM],[PACKAGE],[COPTD001],[COPTD002],[COPTD003],[MC001],[MC004],[MD003],[MD035],[MD006],[MD007],[MD008]
                            ,CONVERT(decimal(16,3),([NUM]/[MC004]*[MD006]/[MD007]*(1+[MD008]))) AS TNUM
                            ,[MB004]
                            FROM [TKMOC].dbo.[MOCMANULINETEMP],[TK].dbo.BOMMC,[TK].dbo.BOMMD
                            LEFT JOIN [TK].dbo.INVMB ON INVMB.MB001=MD003
                            WHERE [MOCMANULINETEMP].MB001=MC001
                            AND MC001=MD001
                            AND [MANU]='新廠製三組(手工)'
                            AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE])
                            AND REPLACE([MC001],' ','')+REPLACE([MD003] ,' ','') NOT IN ( SELECT REPLACE([MB001],' ','')+REPLACE(MD1.[MD003] ,' ','')  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB],[TK].dbo.BOMMC MC1,[TK].dbo.BOMMD MD1 WHERE [PREMANUUSEDINVMB].MB001=MC1.MC001 AND MC1.MC001=MD1.MD001 AND  (REPLACE(MD1.[MD003] ,' ','')  LIKE '1%'  OR (REPLACE(MD1.[MD003] ,' ','')  LIKE '203%' ) )) 
                            ) AS TEMP 
                            ) AS TEMP2   
                            WHERE [MD003]='{2}'
                            ", SDAY, EDAY, MD003);



            return SB;

        }

        public void SEARCHDG1(string SEARCHSTRING, int INDEX)
        {
            String searchValue = SEARCHSTRING;
            rowIndexDG1 = INDEX;
            int ROWS = 0;

            for (int i = INDEX; i < dataGridView1.Rows.Count; i++)
            {
                ROWS = i;

                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Contains(searchValue))
                {
                    rowIndexDG1 = i;

                    dataGridView1.CurrentRow.Selected = false;
                    dataGridView1.Rows[i].Selected = true;
                    int index = rowIndexDG1;
                    dataGridView1.FirstDisplayedScrollingRowIndex = index;

                    break;
                }
                if (dataGridView1.Rows[i].Cells[1].Value.ToString().Contains(searchValue))
                {
                    rowIndexDG1 = i;

                    dataGridView1.CurrentRow.Selected = false;
                    dataGridView1.Rows[i].Selected = true;
                    int index = rowIndexDG1;
                    dataGridView1.FirstDisplayedScrollingRowIndex = index;

                    break;
                }
            }

            if (ROWS == dataGridView1.Rows.Count - 1)
            {
                if (MessageBox.Show("已查到最後一筆，是否從頭開始?", "已查到最後一筆，是否從頭開始?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    SEARCHDG1(textBox1.Text.Trim(), 0);
                }
                else
                {

                }
            }
        }

        public void SEARCHPREMANUUSEDINVMB()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [MB001] AS '品號' ,[MB002] AS '品名'");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[PREMANUUSEDINVMB]");
                sbSql.AppendFormat(@"  ORDER BY [MB001]");
                sbSql.AppendFormat(@"  ");


                adapter3 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder3 = new SqlCommandBuilder(adapter3);
                sqlConn.Open();
                ds3.Clear();
                adapter3.Fill(ds3, "ds3");
                sqlConn.Close();


                if (ds3.Tables["ds3"].Rows.Count == 0)
                {
                    dataGridView3.DataSource = null;
                }
                else
                {
                    if (ds3.Tables["ds3"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView3.DataSource = ds3.Tables["ds3"];
                        dataGridView3.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

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

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            SEARCHINVMB(textBox2.Text.Trim());
        }

        public void SEARCHINVMB(string MB001)
        {
            textBox3.Text = null;

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds4.Clear();

                sbSql.AppendFormat(@"  SELECT MB002,MB004,MB068 ");
                sbSql.AppendFormat(@"  FROM [TK].dbo.INVMB");
                sbSql.AppendFormat(@"  WHERE MB001='{0}'", MB001);
                sbSql.AppendFormat(@"  ");

                adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder4 = new SqlCommandBuilder(adapter4);
                sqlConn.Open();
                ds4.Clear();
                adapter4.Fill(ds4, "ds4");
                sqlConn.Close();


                if (ds4.Tables["ds4"].Rows.Count == 0)
                {

                }
                else
                {
                    if (ds4.Tables["ds4"].Rows.Count >= 1)
                    {
                        textBox3.Text = ds4.Tables["ds4"].Rows[0]["MB002"].ToString();
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

        public void ADDPREMANUUSEDINVMB(string MB001,string MB002)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[PREMANUUSEDINVMB]");
                sbSql.AppendFormat(" ([MB001],[MB002])");
                sbSql.AppendFormat(" VALUES('{0}','{1}')",MB001,MB002);
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
        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            textBox4.Text = null;
            textBox5.Text = null;

            if (dataGridView3.CurrentRow != null)
            {
                int rowindex = dataGridView3.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView3.Rows[rowindex];
                    textBox4.Text = row.Cells["品號"].Value.ToString().Trim();
                    textBox5.Text = row.Cells["品名"].Value.ToString().Trim();

                }
                else
                {
                    MD003 = null;
                }
            }
        }
        public void DELETEPREMANUUSEDINVMB(string MB001)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" DELETE [TKMOC].[dbo].[PREMANUUSEDINVMB]");
                sbSql.AppendFormat(" WHERE [MB001]='{0}'", MB001);          
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
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.CurrentCell.ColumnIndex.Equals(6) && e.RowIndex != -1)
            {
                if (dataGridView1.CurrentCell != null && dataGridView1.CurrentCell.Value != null)
                {
                    //MessageBox.Show(dataGridView1.CurrentCell.Value.ToString());
                    SEARCHPUR(MD003,dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                }

            }
        }

        public void SEARCHPUR(string MD003,string SDay,string EDay)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT TD012 AS '預交日',TD004 AS '品號',TD005 AS '品名',TD006 AS '規格',ISNULL((TD008-TD015),0) AS '採購未交數量',TD009 AS '採購單位'");
                sbSql.AppendFormat(@"  ,CONVERT(DECIMAL(14,2),(CASE WHEN ISNULL(MD002,'')<>'' THEN (ISNULL((TD008-TD015),0)*MD004/MD003) ELSE TD008 END )) AS '數量'");
                sbSql.AppendFormat(@"  ,MB004 AS '單位',TD001 AS '採購單別',TD002 AS '採購單號',TD003 AS '序號',MD003,MD004");
                sbSql.AppendFormat(@"  FROM [TK].dbo.INVMB,[TK].dbo.PURTC,[TK].dbo.PURTD");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMD ON MD001=TD004 AND MD002=TD009");
                sbSql.AppendFormat(@"  WHERE TC001=TD001 AND TC002=TD002");
                sbSql.AppendFormat(@"  AND TD004=MB001");
                sbSql.AppendFormat(@"  AND TD018='Y' AND TD016='N'");
                sbSql.AppendFormat(@"  AND TD012>='{0}' AND TD012<='{1}'",SDay,EDay);
                sbSql.AppendFormat(@"  AND TD004='{0}'",MD003);
                sbSql.AppendFormat(@"  ORDER BY TD012");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");


                adapter5 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder5 = new SqlCommandBuilder(adapter5);
                sqlConn.Open();
                ds5.Clear();
                adapter5.Fill(ds5, "ds5");
                sqlConn.Close();


                if (ds5.Tables["ds5"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else
                {
                    if (ds5.Tables["ds5"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView2.DataSource = ds5.Tables["ds5"];
                        dataGridView2.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

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

        public void RESETTKMOCBOMMD()
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
                                    DELETE [TKMOC].dbo.TKMOCBOMMD

                                    INSERT INTO [TKMOC].dbo.TKMOCBOMMD
                                    ([MD001],[MD003],[WATERNUMS],[OILNUMS],[OILCAL],[WATERCAL])
                                    SELECT RTRIM(LTRIM(MD001)) MD001,RTRIM(LTRIM(MD003)) MD003,WATERNUMS,OILNUMS,TEMP4.OILCAL,TEMP4.WATERCAL
                                    FROM (
                                    -- 4 TEMP4 前  找出油酥的顆數
                                    SELECT TEMP3.MD001,TEMP3.MD003,TEMP3.SUMMD006,TEMP3.WATERNUMS,TEMP3.OILCAL,TEMP3.WATERCAL
                                    ,((SUMMD006*OILCAL)/((SELECT TOP 1 [MOCSEPECIALCAL].[OILNUMS] FROM [TKMOC].[dbo].[MOCSEPECIALCAL] WHERE [MOCSEPECIALCAL].MD003 IN (SELECT MD003 FROM [TK].dbo.BOMMD MD WHERE MD.MD001=TEMP3.MD001)) )) AS 'OILNUMS'
                                    FROM 
                                    (
                                    -- 3 TEMP3 前  找出油酥的總重跟比率
                                    SELECT BOMMD.MD001,BOMMD.MD003,WATERCAL
                                    ,(SELECT SUM(MD.MD006) FROM[TK].dbo.BOMMD  MD WHERE  MD.MD003 LIKE '1%' AND MD.MD003 NOT IN ('101001009') AND MD.MD001= BOMMD.MD001 ) AS 'SUMMD006'
                                    ,(SELECT 66/MD.MD006 FROM [TK].dbo.BOMMD MD WHERE  MD.MD003 LIKE '1%' AND MD.MD003='101001002' AND MD.MD001=BOMMD.MD001 ) AS 'OILCAL'
                                    ,TEMP2.WATERNUMS
                                    FROM [TK].dbo.BOMMD ,(

                                    --2 TEMP2 前 找出水麵顆數
                                    SELECT MD003,WATERCAL
                                    ,((TEMP.MD006*(WATERCAL))/((SELECT TOP 1 [MOCSEPECIALCAL].[WATERNUMS] FROM [TKMOC].[dbo].[MOCSEPECIALCAL] WHERE [MOCSEPECIALCAL].MD003=TEMP.MD003 ) )) AS 'WATERNUMS'
                                    FROM (
                                    --1 TEMP  前 先找出水麵的總重跟比率
                                    SELECT BOMMD.MD001 AS MD003,SUM(BOMMD.MD006) AS MD006
                                    ,(SELECT 66/MD.MD006 FROM [TKMOC].[dbo].[MOCSEPECIALCAL],[TK].dbo.BOMMD MD WHERE [MOCSEPECIALCAL].MD003=MD.MD001 AND MD.MD003 LIKE '1%' AND [MOCSEPECIALCAL].[MD003]=BOMMD.MD001 AND MD.MD003='101001001'  ) AS 'WATERCAL'
                                    FROM [TK].dbo.BOMMD
                                    WHERE  BOMMD.MD003 LIKE '1%'
                                    AND BOMMD.MD003 NOT IN ('101001009')
                                    AND BOMMD.MD001 IN (SELECT MD003 FROM  [TKMOC].[dbo].[MOCSEPECIALCAL])
                                    GROUP BY BOMMD.MD001

                                    ) AS TEMP

                                    ) AS TEMP2
                                    WHERE BOMMD.MD003=TEMP2.MD003

                                    ) AS TEMP3

                                    ) AS TEMP4
                                    ORDER BY MD003,MD001
                                    
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

        public void SEARCHMOCMANULINESPECIAL()
        {
            DataSet ds1 = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();
                
                sbSql.AppendFormat(@"  
                                    SELECT MD003  AS '品號' ,INVMB.MB002 AS '品名',SUM(MDSUM) '需求量', MB004 AS '單位'
                                    FROM (
                                    SELECT MB001,MB002,MD1MD003,MD2MD003,BAR,MD1SUM,MD2SUM,MANUDATE
                                    ,(CASE WHEN ISNULL(MD2MD003,'')='' THEN MD1MD003 ELSE MD2MD003 END) AS 'MD003'
                                    ,(CASE WHEN ISNULL(MD2MD003,'')='' THEN MD1SUM ELSE MD2SUM END) AS 'MDSUM'
                                    FROM(
                                    SELECT 
                                    MOCMANULINE.MB001,MOCMANULINE.MB002,MOCMANULINE.BAR,MOCMANULINE.MANUDATE,MOCMANULINE.OUTDATE
                                    ,TKMOCBOMMD.WATERNUMS,TKMOCBOMMD.OILNUMS,TKMOCBOMMD.OILCAL,TKMOCBOMMD.WATERCAL
                                    ,MC1.MC004 AS MC1MC004
                                    ,MD1.MD001 AS MD1MD001,MD1.MD003 AS MD1MD003,MD1.MD006 AS MD1MD006,MD1.MD007 AS MD1MD007,MD1.MD008 AS MD1MD008
                                    ,ISNULL(MC2.MC004,0) AS MC2MC004
                                    ,MD2.MD001 AS MD2MD001,MD2.MD003 AS MD2MD003,ISNULL(MD2.MD006,0) AS MD2MD006,ISNULL(MD2.MD007,0) AS MD2MD007,ISNULL(MD2.MD008,0) AS MDMD008
                                    ,(MOCMANULINE.BAR*MD1.MD006/MD1.MD007*(1+MD1.MD008)*TKMOCBOMMD.OILCAL) AS 'MD1SUM'
                                    ,ISNULL((MOCMANULINE.BAR*MD2.MD006/MD2.MD007*(1+MD2.MD008)*TKMOCBOMMD.OILNUMS/TKMOCBOMMD.WATERNUMS)*TKMOCBOMMD.WATERCAL,0) AS 'MD2SUM'
                                    FROM [TKMOC].dbo.MOCMANULINE,[TKMOC].dbo.TKMOCBOMMD
                                    LEFT JOIN [TK].dbo.BOMMC MC1 ON MC1.MC001=TKMOCBOMMD.MD001
                                    LEFT JOIN [TK].dbo.BOMMD MD1 ON MD1.MD001=MC1.MC001
                                    LEFT JOIN [TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003
                                    LEFT JOIN [TK].dbo.BOMMD MD2 ON MD2.MD001=MC2.MC001
                                    WHERE MOCMANULINE.MANU IN ('新廠製一組' ,'新廠製二組')
                                    AND MOCMANULINE.MB001 IN (SELECT [MD001] FROM [TKMOC].[dbo].[TKMOCBOMMD])
                                    AND MOCMANULINE.MB001=TKMOCBOMMD.MD001
                                    AND (MD1.MD003 LIKE '1%' OR MD1.MD003 LIKE '3%' )
                                    AND (MD2.MD003 LIKE '1%' OR MD2.MD003 LIKE '3%' OR ISNULL(MD2.MD003,'')='')
                                    AND MOCMANULINE.MANUDATE>='{0}' AND MOCMANULINE.MANUDATE<='{1}'
                                    --AND MOCMANULINE.MB001='3010101601'
                                    --AND MOCMANULINE.BAR='0.8083'
                                    ) AS TEMP
                                    ) AS TEMP2
                                    LEFT JOIN [TK].dbo.INVMB ON TEMP2.MD003=INVMB.MB001
                                    WHERE INVMB.MB002 NOT LIKE '%餅麩%' 
                                    AND INVMB.MB002 NOT LIKE '%回收料%' 
                                    GROUP BY MD003,INVMB.MB002,MB004
                                    ORDER BY MD003,INVMB.MB002,MB004

                                    ",dateTimePicker3.Value.ToString("yyyyMMdd"), dateTimePicker4.Value.ToString("yyyyMMdd"));


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count == 0)
                {
                    dataGridView4.DataSource = null;
                }
                else
                {
                    if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView4.DataSource = ds1.Tables["ds1"];
                        dataGridView4.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

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

        public void SEARCHMOCMANULINESPECIAL2(string MD003)
        {
            DataSet ds1 = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  
                                    SELECT MB001 AS '品號',MB002 AS '品名',BAR  AS '桶數',CONVERT(NVARCHAR,MANUDATE,112)  AS '生產日'
                                    ,(CASE WHEN ISNULL(MD2MD003,'')='' THEN MD1SUM ELSE MD2SUM END) AS 'MDSUM'                                    
                                    ,(CASE WHEN ISNULL(MD2MD003,'')='' THEN MD1MD003 ELSE MD2MD003 END) AS 'MD003'                                    
                                    ,MD1SUM,MD2SUM,MD1MD003,MD2MD003

                                    FROM(
                                    SELECT 
                                    MOCMANULINE.MB001,MOCMANULINE.MB002,MOCMANULINE.BAR,MOCMANULINE.MANUDATE,MOCMANULINE.OUTDATE
                                    ,TKMOCBOMMD.WATERNUMS,TKMOCBOMMD.OILNUMS,TKMOCBOMMD.OILCAL,TKMOCBOMMD.WATERCAL
                                    ,MC1.MC004 AS MC1MC004
                                    ,MD1.MD001 AS MD1MD001,MD1.MD003 AS MD1MD003,MD1.MD006 AS MD1MD006,MD1.MD007 AS MD1MD007,MD1.MD008 AS MD1MD008
                                    ,ISNULL(MC2.MC004,0) AS MC2MC004
                                    ,MD2.MD001 AS MD2MD001,MD2.MD003 AS MD2MD003,ISNULL(MD2.MD006,0) AS MD2MD006,ISNULL(MD2.MD007,0) AS MD2MD007,ISNULL(MD2.MD008,0) AS MDMD008
                                    ,(MOCMANULINE.BAR*MD1.MD006/MD1.MD007*(1+MD1.MD008)*TKMOCBOMMD.OILCAL) AS 'MD1SUM'
                                    ,ISNULL((MOCMANULINE.BAR*MD2.MD006/MD2.MD007*(1+MD2.MD008)*TKMOCBOMMD.OILNUMS/TKMOCBOMMD.WATERNUMS)*TKMOCBOMMD.WATERCAL,0) AS 'MD2SUM'
                                    FROM [TKMOC].dbo.MOCMANULINE,[TKMOC].dbo.TKMOCBOMMD
                                    LEFT JOIN [TK].dbo.BOMMC MC1 ON MC1.MC001=TKMOCBOMMD.MD001
                                    LEFT JOIN [TK].dbo.BOMMD MD1 ON MD1.MD001=MC1.MC001
                                    LEFT JOIN [TK].dbo.BOMMC MC2 ON MC2.MC001=MD1.MD003
                                    LEFT JOIN [TK].dbo.BOMMD MD2 ON MD2.MD001=MC2.MC001
                                    WHERE MOCMANULINE.MANU IN ('新廠製一組' ,'新廠製二組')
                                    AND MOCMANULINE.MB001 IN (SELECT [MD001] FROM [TKMOC].[dbo].[TKMOCBOMMD])
                                    AND MOCMANULINE.MB001=TKMOCBOMMD.MD001
                                    AND (MD1.MD003 LIKE '1%' OR MD1.MD003 LIKE '3%' )
                                    AND (MD2.MD003 LIKE '1%' OR MD2.MD003 LIKE '3%' OR ISNULL(MD2.MD003,'')='')
                                    AND MOCMANULINE.MANUDATE>='{0}' AND MOCMANULINE.MANUDATE<='{1}'
                                    --AND MOCMANULINE.MB001='3010101601'
                                    --AND MOCMANULINE.BAR='0.8083'
                                    ) AS TEMP
                                    WHERE (TEMP.MD1MD003='{2}' OR TEMP.MD2MD003='{2}')  

                                    ", dateTimePicker3.Value.ToString("yyyyMMdd"), dateTimePicker4.Value.ToString("yyyyMMdd"),MD003);


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count == 0)
                {
                    dataGridView5.DataSource = null;
                }
                else
                {
                    if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView5.DataSource = ds1.Tables["ds1"];
                        dataGridView5.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

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
        private void dataGridView4_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView4.CurrentRow != null)
            {
                int rowindex = dataGridView4.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView4.Rows[rowindex];
                    MD003 = row.Cells["品號"].Value.ToString().Trim();

                    SEARCHMOCMANULINESPECIAL2(MD003);


                }
                else
                {
                    MD003 = null;
                }
            }
        }


        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            SEARCHMOCMANULINE();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            
            if (rowIndexDG1 == -1)
            {
                SEARCHDG1(textBox1.Text.Trim(), 0);
            }
            else
            {
                SEARCHDG1(textBox1.Text.Trim(), rowIndexDG1 + 1);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            SEARCHPREMANUUSEDINVMB();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            ADDPREMANUUSEDINVMB(textBox2.Text.Trim(),textBox3.Text.Trim());
            SEARCHPREMANUUSEDINVMB();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELETEPREMANUUSEDINVMB(textBox4.Text.Trim());
                SEARCHPREMANUUSEDINVMB();

            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }


        private void button6_Click(object sender, EventArgs e)
        {
            RESETTKMOCBOMMD();
            SEARCHMOCMANULINESPECIAL();

        }

        private void button7_Click(object sender, EventArgs e)
        {

        }

        private void button8_Click(object sender, EventArgs e)
        {

        }


        #endregion

       
    }
}
