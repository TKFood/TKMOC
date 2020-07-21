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

namespace TKMOC
{
    public partial class frmMOCMANULINESubTEMPADD : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        DataSet ds1 = new DataSet();


        public frmMOCMANULINESubTEMPADD()
        {
            InitializeComponent();
        }

        public frmMOCMANULINESubTEMPADD(string MB001,DataSet Fds)
        {
            InitializeComponent();

            textBox1.Text = MB001;
            
        }

        #region FUNCTION
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            SERACH(textBox1.Text.Trim());
        }

        public void SERACH(string MB001)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  SELECT ");
                sbSql.AppendFormat(@"  [MANU] AS '線別',CONVERT(varchar(100),[MANUDATE],112) AS '生產日',[MB001] AS '品號',[MB002] AS '品名' ");
                sbSql.AppendFormat(@"  ,[MB003] AS '規格',[BAR] AS '桶數',[NUM] AS '數量',[CLINET] AS '客戶',[OUTDATE] AS '交期',[TA029] AS '備註',[HALFPRO] AS '半成品數量'");
                sbSql.AppendFormat(@"  ,[COPTD001] AS '訂單單別',[COPTD002] AS '訂單號',[COPTD003] AS '訂單序號'");
                sbSql.AppendFormat(@"  ,[ID]");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINETEMP]");
                sbSql.AppendFormat(@"  WHERE  [MB001]='{0}'",MB001);
                sbSql.AppendFormat(@"  ORDER BY [SERNO]");
                sbSql.AppendFormat(@"  ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count == 0)
                {

                }
                else
                {
                    if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        ////dataGridView1.Rows.Clear();
                        //dataGridView5.DataSource = ds7.Tables["TEMPds7"];
                        //dataGridView5.AutoResizeColumns();
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

        public DataSet SETDATASET
        {
            //set
            //{
            //    textBox1.Text = value;
            //}
            get
            {
                return ds1;
            }
        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        #endregion


    }
}
