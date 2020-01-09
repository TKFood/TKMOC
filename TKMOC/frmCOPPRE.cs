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
    public partial class frmCOPPRE : Form
    {
        private ComponentResourceManager _ResourceManager = new ComponentResourceManager();
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();
       

        int result;

        public frmCOPPRE()
        {
            InitializeComponent();
        }


        #region FUNCTION

        public void PRESCHEDULE()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [PREORDER].[ORDERNO],[PREORDER].[MB001],[PREORDER].[AMOUNT],[PREORDER].[PRIORITYS],[PREINVMBMANU].MANU,[PREINVMBMANU].TIMES");
                sbSql.AppendFormat(@"  ,CONVERT(INT,ROUND([PREORDER].[AMOUNT]/[PREINVMBMANU].TIMES,0)) AS HRS");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[PREORDER],[TKMOC].[dbo].[PREINVMBMANU]");
                sbSql.AppendFormat(@"  WHERE [PREORDER].MB001=[PREINVMBMANU].MB001");
                sbSql.AppendFormat(@"  ORDER BY [PREINVMBMANU].MANU,[PREORDER].[PRIORITYS] DESC,[ORDERNO]");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

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
        private void button1_Click(object sender, EventArgs e)
        {
            PRESCHEDULE();
        }
        #endregion
    }
}
