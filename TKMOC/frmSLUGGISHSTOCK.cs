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
    public partial class frmSLUGGISHSTOCK : Form
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

        SqlCommand cmd = new SqlCommand();
        SqlTransaction tran;
        int result;

        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();

        string tablename = null;

        public frmSLUGGISHSTOCK()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SEARCH(string SDay)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  
                                    SELECT 品號,品名,批號,庫存量,單位,在倉日期,有效天數,業務
                                    
                                    FROM (
                                    SELECT   LA001 AS '品號' ,INVMB.MB002 AS '品名',INVMB.MB003 AS '規格',LA016 AS '批號'
                                    ,CONVERT(DECIMAL(16,3),SUM(LA005*LA011)) AS '庫存量',INVMB.MB004 AS '單位'
                                    ,DATEDIFF(DAY,LA016,'{0}') AS '在倉日期old' 
                                    ,(CASE WHEN DATEDIFF(DAY,LA016,'{0}')>=0 THEN DATEDIFF(DAY,LA016,'{0}') ELSE (CASE WHEN DATEDIFF(DAY,LA016,'{0}')<0 THEN  (CASE WHEN MB198='2' THEN DATEDIFF(DAY,DATEADD(month, -1*MB023, LA016 ),'{0}') END ) END ) END) AS '在倉日期' 
                                    ,(CASE WHEN MB198='2' THEN DATEDIFF(DAY,'{0}',DATEADD(month, MB023, '{0}' )) END)-(CASE WHEN DATEDIFF(DAY,LA016,'{0}')>=0 THEN DATEDIFF(DAY,LA016,'{0}') ELSE (CASE WHEN DATEDIFF(DAY,LA016,'{0}')<0 THEN  (CASE WHEN MB198='2' THEN DATEDIFF(DAY,DATEADD(month, -1*MB023, LA016 ),'{0}') END ) END ) END)  AS '有效天數'
                                    ,(SELECT TOP 1 TC006+' '+MV002 FROM [TK].dbo.COPTC,[TK].dbo.CMSMV WHERE TC006=MV001 AND  TC001+TC002 IN (SELECT TOP 1 TA026+TA027 FROM [TK].dbo.MOCTA WHERE TA001+TA002 IN (SELECT TOP 1 TG014+TG015 FROM [TK].dbo.MOCTG WHERE TG004=LA001 AND TG017=LA016))) AS '業務'
                                    FROM [TK].dbo.INVLA WITH (NOLOCK) 
                                    LEFT JOIN  [TK].dbo.INVMB WITH (NOLOCK) ON MB001=LA001  
                                    WHERE  (LA009='20005') 
                                    GROUP BY  LA001,LA009,MB002,MB003,LA016,MB023,MB198,MB004
                                    HAVING SUM(LA005*LA011)<>0 
                                    ) AS TEMP
                                    ORDER BY 在倉日期 DESC  

                                    ", SDay);

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

                        ////根据列表中数据不同，显示不同颜色背景
                        //foreach (DataGridViewRow dgRow in dataGridView1.Rows)
                        //{
                        //    ////判断
                        //    //if (Convert.ToDecimal(dgRow.Cells[5].Value) > 0)
                        //    //{
                        //    //    //将这行的背景色设置成Pink
                        //    //    dgRow.DefaultCellStyle.BackColor = Color.Pink;

                        //    //}
                        //}

                        dataGridView1.Columns["品號"].Width = 100;
                        dataGridView1.Columns["品名"].Width = 220;



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
            SEARCH(DateTime.Now.ToString("yyyyMMdd"));
        }


        #endregion
    }
}
