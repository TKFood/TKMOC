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
    public partial class frmREPORTINVPURUESDBREAD : Form
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

        string MD003;
        int rowIndexDG1 = -1;
        int rowIndexDG2 = -1;


        public frmREPORTINVPURUESDBREAD()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SEARCHMOCMANULINE(string SDay, string EDay)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT TB003 AS '品號',TB012  AS '品名',SUM(TB004)  AS '需領用量'");
                sbSql.AppendFormat(@"  ,(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA WHERE LA009 IN ('20004','20006') AND  LA001=TB003) AS '原料倉庫存'");
                sbSql.AppendFormat(@"  ,((SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA WHERE LA009 IN ('20004','20006') AND  LA001=TB003)-SUM(TB004)) AS '差異量'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.MOCTA,[TK].dbo.MOCTB");
                sbSql.AppendFormat(@"  WHERE TA001=TB001 AND TA002=TB002");
                sbSql.AppendFormat(@"  AND TA021='04'");
                sbSql.AppendFormat(@"  AND TA003>='{0}' AND TA003<='{0}'", SDay, EDay);
                sbSql.AppendFormat(@"  GROUP BY TB003,TB012");
                sbSql.AppendFormat(@"  ORDER BY TB003,TB012");
                sbSql.AppendFormat(@"  ");

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


        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCHMOCMANULINE(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
        }

        #endregion

    }
}
