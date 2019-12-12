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
    public partial class frmBATCHMOCTAB : Form
    {
        private ComponentResourceManager _ResourceManager = new ComponentResourceManager();
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        SqlDataAdapter adapter2 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder2 = new SqlCommandBuilder();
        SqlDataAdapter adapter3= new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder3 = new SqlCommandBuilder();

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();

        List<ADDITEM> ADDTARGET = new List<ADDITEM>();
        List<ADDITEM> FIND = new List<ADDITEM>();

        public class ADDITEM
        {
            public string MB001;
            public double NUM;

        }

        public frmBATCHMOCTAB()
        {
            InitializeComponent();
        }

        



        #region FUNCTION
        public void SEARCHCOP(DateTime dt1, DateTime dt2)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  SELECT TD013 AS '預交日',TD001 AS '訂單',TD002 AS '訂單號',TD003 AS '序號',TD004 AS '品號',TD005 AS '品名',TD006 AS '規格',(TD008-TD009+TD024-TD025) AS '訂單數量',TD010 AS '訂單單位'");
                sbSql.AppendFormat(@"  ,CONVERT(DECIMAL(18,3),(CASE WHEN MD002=TD010   THEN (TD008-TD009+TD024-TD025)*MD004/MD003 ELSE (TD008-TD009+TD024-TD025) END )) AS '數量'");
                sbSql.AppendFormat(@"  ,MB004 AS '單位',TC015 AS '單頭備註',TD020 AS '單身備註'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.COPTC,[TK].dbo.COPTD");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMD ON MD001=TD004 AND MD002=TD010 ");
                sbSql.AppendFormat(@"  ,[TK].dbo.INVMB");
                sbSql.AppendFormat(@"  WHERE TC001=TD001 AND TC002=TD002");
                sbSql.AppendFormat(@"  AND TD004=MB001");
                //sbSql.AppendFormat(@"  AND (TD004 LIKE '410%')");
                sbSql.AppendFormat(@"  AND (TD008-TD009)>0");
                sbSql.AppendFormat(@"  AND TD013>='{0}' AND TD013<='{1}'", dt1.ToString("yyyyMMdd"), dt2.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  ORDER BY TD013,TD001,TD002,TD004");
                sbSql.AppendFormat(@"  ");
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
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            textBox1.Text = null;
            textBox2.Text = null; 
            textBox3.Text = null;
            textBox4.Text = null;
            textBox5.Text = null;
            textBox6.Text = null;

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    textBox1.Text = row.Cells["訂單"].Value.ToString();
                    textBox2.Text = row.Cells["訂單號"].Value.ToString();
                    textBox3.Text = row.Cells["序號"].Value.ToString();
                    textBox4.Text = row.Cells["品號"].Value.ToString();
                    textBox5.Text = row.Cells["數量"].Value.ToString();
                    textBox6.Text = row.Cells["單頭備註"].Value.ToString();
                }
                else
                {
                    textBox1.Text = null;
                    textBox2.Text = null;
                    textBox3.Text = null;
                    textBox4.Text = null;
                    textBox5.Text = null;
                    textBox6.Text = null;
                }
            }
        }

        public void GENADDTARGET()
        {
            ADDTARGET.Clear();
            FIND.Clear();
            //ADDTARGET.RemoveAll(it => true);

            ADDTARGET.Add(new ADDITEM { MB001 =textBox4.Text , NUM = Convert.ToDouble(textBox5.Text) });

            SERACH(ADDTARGET[0].MB001, ADDTARGET[0].NUM, FIND);

            foreach (var find in FIND)
            {
                CHECKBOMMD(find.MB001, find.NUM);
            }

            foreach (var find in ADDTARGET)
            {
                MessageBox.Show(find.MB001 + " " + find.NUM);
            }

        }

        public void SERACH(string MB001, double NUM, List<ADDITEM> FIND)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  WITH NODE (MD001,MD003,LAYER,MD004,MC004,PREMC004,MD006,MD007,MD008,USEDNUM ,PREUSEDNUM) AS");
                sbSql.AppendFormat(@"  (");
                sbSql.AppendFormat(@"  SELECT MD001,MD003,0 ,[MD004],[MC004],[MC004] AS PREMC004,[MD006],[MD007],[MD008],CONVERT(DECIMAL(18,4),([MD006]/[MD007]/[MC004]*(1+MD008))),CONVERT(DECIMAL(18,4),1) AS PREUSEDNUM  FROM [TK].[dbo].[VBOMMD]");
                sbSql.AppendFormat(@"  UNION ALL");
                sbSql.AppendFormat(@"  SELECT TB1.MD001,TB2.MD003,TB2.LAYER+1,TB2.MD004,TB2.MC004,TB1.MC004,TB2.MD006,TB2.MD007,TB2.MD008,TB2.USEDNUM,CONVERT(DECIMAL(18,4),(TB1.[MD006]/TB1.[MD007]/TB1.[MC004]*(1+TB1.MD008))) AS PREUSEDNUM FROM [TK].[dbo].[VBOMMD] TB1");
                sbSql.AppendFormat(@"  INNER JOIN NODE TB2");
                sbSql.AppendFormat(@"  ON TB1.MD003 = TB2.MD001");
                sbSql.AppendFormat(@"  )");
                sbSql.AppendFormat(@"  SELECT MD001,MD003,LAYER,MD004,MC004,PREMC004,MD006,MD007,MD008,USEDNUM ,PREUSEDNUM ,USEDNUM*PREUSEDNUM*{0} AS TOTALUSED FROM NODE", NUM);
                sbSql.AppendFormat(@"  WHERE  MD001='{0}'", MB001);
                sbSql.AppendFormat(@"  ORDER BY LAYER ,MD001, MD003");
                sbSql.AppendFormat(@"  ");
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

                }
                else
                {
                    if (ds2.Tables["ds2"].Rows.Count >= 1)
                    {
                        foreach (DataRow od in ds2.Tables["ds2"].Rows)
                        {
                            FIND.Add(new ADDITEM { MB001 = od["MD003"].ToString(), NUM = Convert.ToDouble(od["TOTALUSED"].ToString()) });
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

        public void CHECKBOMMD(string MB001, double NUM)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT MD001,MD003");
                sbSql.AppendFormat(@"  FROM [TK].dbo.BOMMD");
                sbSql.AppendFormat(@"  WHERE MD001='{0}'", MB001);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter3 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder3 = new SqlCommandBuilder(adapter3);
                sqlConn.Open();
                ds3.Clear();
                adapter3.Fill(ds3, "ds3");
                sqlConn.Close();


                if (ds3.Tables["ds3"].Rows.Count == 0)
                {

                }
                else
                {
                    if (ds3.Tables["ds3"].Rows.Count >= 1)
                    {
                        ADDTARGET.Add(new ADDITEM { MB001 = MB001, NUM = NUM });

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
            SEARCHCOP(dateTimePicker1.Value, dateTimePicker2.Value);
        }
        private void button2_Click(object sender, EventArgs e)
        {
            GENADDTARGET();
        }

        #endregion

        
    }
}
