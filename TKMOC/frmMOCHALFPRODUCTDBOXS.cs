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
    public partial class frmMOCHALFPRODUCTDBOXS : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();


        SqlCommand cmd = new SqlCommand();
        SqlTransaction tran;
        int result;

        DataSet ds1 = new DataSet();

        public frmMOCHALFPRODUCTDBOXS()
        {
            InitializeComponent();
        }


        #region FUNCTION
        public void SEARCHMOCHALFPRODUCTDBOXS()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();                

                sbSql.AppendFormat(@"  
                                SELECT [MOCHALFPRODUCTDBOXS].[MB001] AS '品號',[MB002] AS '品名',[NUMS] AS '箱重',[BOXS] AS '箱數'
                                FROM [TKMOC].[dbo].[MOCHALFPRODUCTDBOXS],[TK].dbo.[INVMB]
                                WHERE [MOCHALFPRODUCTDBOXS].[MB001]=[INVMB].[MB001]
                                 ");

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
            SEARCHMOCHALFPRODUCTDBOXS();
        }

        #endregion
    }
}
