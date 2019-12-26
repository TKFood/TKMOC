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
    public partial class frmREPORTGEN : Form
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

        public Report report1 { get; private set; }

        public frmREPORTGEN()
        {
            InitializeComponent();
        }


        #region FUNCTION
        public void SEARCHMOCTA(DateTime dt,DateTime dt2)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@" SELECT TA001 AS '製令',TA002 AS '製令號',TA034 AS '品名' ,TA017 AS '已生產量',TA007 AS '單位',TA035  AS '規格',TA029 '備註'");
                sbSql.AppendFormat(@" FROM [TK].dbo.MOCTA ");
                sbSql.AppendFormat(@" WHERE TA003>='{0}' AND TA003<='{1}' ", dt.ToString("yyyyMMdd"), dt2.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@" AND TA001 IN ('A510','A511','A512','A521','A522') ");
                sbSql.AppendFormat(@" ORDER BY TA001,TA002,TA034 ");
                sbSql.AppendFormat(@"   ");
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

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            textBox1.Text = null;
            textBox2.Text = null;
            textBox3.Text = null;

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    textBox1.Text = row.Cells["製令"].Value.ToString();
                    textBox2.Text = row.Cells["製令號"].Value.ToString();
                    textBox3.Text = row.Cells["備註"].Value.ToString();
                }
                else
                {
                    textBox1.Text = null;
                    textBox2.Text = null;
                    textBox3.Text = null;

                }
            }
        }

        public void SETFASTREPORT(string TA001,string TA002, string COMMENT, string NUM)
        {
            string SQL;
            report1 = new Report();
            report1.Load(@"REPORT\生產入庫單.frx");

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;

            TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;
            SQL = SETFASETSQL(TA001, TA002);
            Table.SelectCommand = SQL;

            report1.SetParameterValue("P1", COMMENT);
            report1.SetParameterValue("P2", NUM);

            report1.Preview = previewControl1;
            report1.Show();
        }


        public string SETFASETSQL(string TA001, string TA002)
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();

           
            FASTSQL.AppendFormat(@" SELECT TA001 AS '製令',TA002 AS '製令號',SUBSTRING(TA002,1,4) AS '年',SUBSTRING(TA002,5,2) AS '月',SUBSTRING(TA002,7,2) AS '日',TA034 AS '品名',MB003 AS '規格'");
            FASTSQL.AppendFormat(@"  ,MB003 AS '規格',TA017 AS '已生產量' ");
            FASTSQL.AppendFormat(@"  FROM [TK].dbo.MOCTA");
            FASTSQL.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMB ON MB001=TA006");
            FASTSQL.AppendFormat(@"  WHERE TA001='{0}' AND TA002='{1}'",TA001,TA002);
            FASTSQL.AppendFormat(@"     ");

            return FASTSQL.ToString();
        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCHMOCTA(dateTimePicker1.Value, dateTimePicker2.Value);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(textBox1.Text,textBox2.Text, textBox3.Text, textBox4.Text);
        }

        #endregion


    }
}
