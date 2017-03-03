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

namespace TKMOC
{
    public partial class frmINV : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();
        DataSet ds5 = new DataSet();
        DataSet ds6 = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        int rownum = 0;

        string MB001;
        string CMSMC;
        string DATESTRING;

        public frmINV()
        {
            InitializeComponent();
            dateTimePicker1.Value = DateTime.Now;
        }

        private void frmINV_Load(object sender, EventArgs e)
        {
            comboBox1load();
        }
        #region FUNCTION
        public void comboBox1load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MC001,MC002 FROM [TK].dbo.CMSMC WHERE MC001 LIKE '2%' ORDER BY MC001  ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MC001", typeof(string));
            dt.Columns.Add("MC002", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "MC001";
            comboBox1.DisplayMember = "MC001";
            sqlConn.Close();
            

        }
        public void Search()
        {
            try
            {
                sbSql.Clear();
                sbSql = SETsbSql();

                if (!string.IsNullOrEmpty(sbSql.ToString()))
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
                    sqlCmdBuilder = new SqlCommandBuilder(adapter);

                    sqlConn.Open();
                    ds.Clear();
                    adapter.Fill(ds, tablename);
                    sqlConn.Close();

                    if (ds.Tables[tablename].Rows.Count == 0)
                    {

                    }
                    else
                    {
                        dataGridView1.DataSource = ds.Tables[tablename];
                        dataGridView1.AutoResizeColumns();
                        dataGridView1.CurrentCell = dataGridView1.Rows[rownum].Cells[0];

                    }
                }
                else
                {

                }



            }
            catch
            {

            }
            finally
            {

            }

        }

        public StringBuilder SETsbSql()
        {
            StringBuilder STR = new StringBuilder();


            STR.AppendFormat(@"  SELECT LA001 AS '品號',MB002 AS '品名',MB003 AS '規格',SUM(LA005*LA011) AS '庫存數量' ");
            STR.AppendFormat(@"  ,(Select ISNULL(SUM(TB009),0) from [TK].dbo.PURTA AS PURTA, [TK].dbo.PURTB AS PURTB LEFT JOIN [TK].dbo.CMSMC AS CMSMC ON MC001=TB008 Where TA001=TB001 and TA002=TB002 and   TB025='Y' and TB021='N' and (TB022='' or TB022 IS NULL)  and TB004=LA001 AND MC005='Y' AND TB039='N' and TB019<='{0}'  and TB008='{1}' ) AS '預計請購量'", dateTimePicker1.Value.ToString("yyyyMMdd"), comboBox1.Text.ToString());
            STR.AppendFormat(@"  ,(select ISNULL(SUM(TD008),0) from[TK].dbo.PURTD PURTD INNER JOIN [TK].dbo.PURTC PURTC ON TC001=TD001 AND TC002=TD002 and TD018='Y' and TD016='N' AND TD008>TD015 INNER JOIN [TK].dbo.CMSMC CMSMC ON TD007=MC001 where TD004=LA001   and MC005='Y' and TD012<='{0}'   and (TD007 in ('{1}') ) ) AS '預計進貨量'", dateTimePicker1.Value.ToString("yyyyMMdd"), comboBox1.Text.ToString());
            STR.AppendFormat(@"  ,(SELECT MOUNT=ISNULL(SUM(TA015),0)-ISNULL(SUM(TA017),0)-ISNULL(SUM(TA018),0)  FROM [TK].dbo.MOCTA AS MOCTA LEFT JOIN [TK].dbo.CMSMC AS CMSMC ON TA020=MC001 LEFT JOIN [TK].dbo.CMSMQ AS CMSMQ ON TA001=MQ001 WHERE TA006=LA001 AND MC005='Y' AND TA013='Y' AND (TA011<>'Y' AND TA011<>'y')  AND MQ014='Y'  AND TA010<='{0}' AND (TA020 in ('{1}') ) ) AS '預計生產量'", dateTimePicker1.Value.ToString("yyyyMMdd"), comboBox1.Text.ToString());
            STR.AppendFormat(@"  ,(SELECT MOUNT=ISNULL(SUM(TB004),0)-ISNULL(SUM(TB005),0) FROM [TK].dbo.MOCTB AS MOCTB LEFT JOIN [TK].dbo.CMSMC AS CMSMC ON TB009=MC001,TK..MOCTA AS MOCTA LEFT JOIN [TK].dbo.CMSMQ AS CMSMQ ON TA001=MQ001 WHERE TA001=TB001 AND TA002=TB002 AND TB003=LA001 AND MC005='Y' AND TB018='Y' AND (TA011<>'Y' AND TA011<>'y')  AND MQ014='Y' AND (TB011='1' OR TB011='2')  AND TB004>TB005  AND TB015<='{0}' AND (TB009 in ('{1}') ) )  AS '預計領料量'", dateTimePicker1.Value.ToString("yyyyMMdd"), comboBox1.Text.ToString());
            STR.AppendFormat(@"  ,(select sum(TD008) from [TK].dbo.COPTD COPTD left join [TK].dbo.CMSMC CMSMC on TD007=MC001 where TD008+TD024+TD050>TD009+TD025+TD051 and  TD004=LA001 and TD021='Y' and TD016='N'  and TD013<='{0}'   and (TD007 in ('{1}') )  and MC005='Y') AS '預計銷貨量' ", dateTimePicker1.Value.ToString("yyyyMMdd"), comboBox1.Text.ToString());
            STR.AppendFormat(@"  FROM [TK].dbo.INVLA WITH (NOLOCK),[TK].dbo.INVMB WITH (NOLOCK)");
            STR.AppendFormat(@"  WHERE LA001=MB001");
            STR.AppendFormat(@"  AND LA001 LIKE '{0}%'",textBox1.Text.ToString());
            STR.AppendFormat(@"  AND LA009='{0}'",comboBox1.Text.ToString());
            STR.AppendFormat(@"  AND LA004<='{0}'",dateTimePicker1.Value.ToString("yyyyMMdd"));
            STR.AppendFormat(@"  GROUP BY LA001,MB002,MB003");
            STR.AppendFormat(@"  ");
            STR.AppendFormat(@"  ");

            tablename = "TEMPds1";
            return STR;
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count >= 1)
            {
                MB001 = dataGridView1.CurrentRow.Cells["品號"].Value.ToString(); ;
                CMSMC = comboBox1.Text.ToString();
                DATESTRING = dateTimePicker1.Value.ToString("yyyyMMdd");

                SearchPURTB();
                SearchPURTD();
                SearchMOCTA();
                SearchMOCTB();
                SearchCOPTD();
            }

        }

        public void SearchPURTB()
        {
            try
            {
                StringBuilder SQLSTING = new StringBuilder();

              
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);
                SQLSTING.AppendFormat(" Select SUM(TB009) AS '預計請購',TB001  AS '單別',TB002  AS '單號' from [TK].dbo.PURTA AS PURTA, [TK].dbo.PURTB AS PURTB");
                SQLSTING.AppendFormat(" LEFT JOIN [TK].dbo.CMSMC AS CMSMC ON MC001=TB008");
                SQLSTING.AppendFormat(" Where TA001=TB001 and TA002=TB002  ");
                SQLSTING.AppendFormat(" and TB025='Y' and TB021='N' and (TB022='' or TB022 IS NULL)");
                SQLSTING.AppendFormat(" and TB004='{0}' AND MC005='Y'",MB001);
                SQLSTING.AppendFormat(" AND TB039='N'");
                SQLSTING.AppendFormat(" and TB019<='{0}' ",DATESTRING);
                SQLSTING.AppendFormat(" and TB008='{0}'",CMSMC);
                SQLSTING.AppendFormat(" GROUP BY TB001,TB002 ");
                SQLSTING.AppendFormat(" ");

                adapter = new SqlDataAdapter(SQLSTING.ToString(), sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds2.Clear();
                adapter.Fill(ds2, "TEMP2");
                sqlConn.Close();

                if (ds2.Tables["TEMP2"].Rows.Count == 0)
                {

                }
                else
                {
                    dataGridView2.DataSource = ds2.Tables["TEMP2"];
                    dataGridView2.AutoResizeColumns();
                    dataGridView2.CurrentCell = dataGridView2.Rows[rownum].Cells[0];

                }
             



            }
            catch
            {

            }
            finally
            {

            }

        }
        public void SearchPURTD()
        {
            try
            {
                StringBuilder SQLSTING = new StringBuilder();

              
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);
           
                SQLSTING.AppendFormat(" select SUM(TD008) AS '預計進貨',TD001  AS '單別',TD002  AS '單號'");
                SQLSTING.AppendFormat(" from[TK].dbo.PURTD PURTD");
                SQLSTING.AppendFormat(" INNER JOIN [TK].dbo.PURTC PURTC ON TC001=TD001 AND TC002=TD002 and TD018='Y' and TD016='N' AND TD008>TD015");
                SQLSTING.AppendFormat(" INNER JOIN [TK].dbo.CMSMC CMSMC ON TD007=MC001");
                SQLSTING.AppendFormat(" where TD004='{0}'  ",MB001);
                SQLSTING.AppendFormat(" and MC005='Y'");
                SQLSTING.AppendFormat(" and TD012<='{0}'  ",DATESTRING);
                SQLSTING.AppendFormat(" and (TD007 in ('{0}') ) ",CMSMC);
                SQLSTING.AppendFormat(" GROUP BY TD001,TD002");
                SQLSTING.AppendFormat(" ");
                adapter = new SqlDataAdapter(SQLSTING.ToString(), sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds3.Clear();
                adapter.Fill(ds3, "TEMP3");
                sqlConn.Close();

                if (ds3.Tables["TEMP3"].Rows.Count == 0)
                {

                }
                else
                {
                    dataGridView3.DataSource = ds3.Tables["TEMP3"];
                    dataGridView3.AutoResizeColumns();
                    dataGridView3.CurrentCell = dataGridView3.Rows[rownum].Cells[0];

                }
          



            }
            catch
            {

            }
            finally
            {

            }

        }
        public void SearchMOCTA()
        {
            try
            {
                StringBuilder SQLSTING = new StringBuilder();


                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);
                SQLSTING.AppendFormat(" SELECT (ISNULL(SUM(TA015),0)-ISNULL(SUM(TA017),0)-ISNULL(SUM(TA018),0)) AS '預計生產量' ,TA001 AS '單別',TA002 AS '單號'");
                SQLSTING.AppendFormat(" FROM [TK].dbo.MOCTA AS MOCTA");
                SQLSTING.AppendFormat(" LEFT JOIN [TK].dbo.CMSMC AS CMSMC ON TA020=MC001 ");
                SQLSTING.AppendFormat(" LEFT JOIN [TK].dbo.CMSMQ AS CMSMQ ON TA001=MQ001");
                SQLSTING.AppendFormat(" WHERE TA006='{0}'", MB001);
                SQLSTING.AppendFormat(" AND MC005='Y' AND TA013='Y' AND (TA011<>'Y' AND TA011<>'y') ");
                SQLSTING.AppendFormat(" AND MQ014='Y' ");
                SQLSTING.AppendFormat(" AND TA010<='{0}'", DATESTRING);
                SQLSTING.AppendFormat(" AND (TA020 in ('{0}') )", CMSMC);
                SQLSTING.AppendFormat("GROUP BY TA001,TA002 ");
                SQLSTING.AppendFormat(" ");
                SQLSTING.AppendFormat(" ");

                adapter = new SqlDataAdapter(SQLSTING.ToString(), sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds4.Clear();
                adapter.Fill(ds4, "TEMP4");
                sqlConn.Close();

                if (ds4.Tables["TEMP4"].Rows.Count == 0)
                {

                }
                else
                {
                    dataGridView4.DataSource = ds4.Tables["TEMP4"];
                    dataGridView4.AutoResizeColumns();
                    dataGridView4.CurrentCell = dataGridView4.Rows[rownum].Cells[0];

                }




            }
            catch
            {

            }
            finally
            {

            }
        }



        public void SearchMOCTB()
        {
            try
            {
                StringBuilder SQLSTING = new StringBuilder();


                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);
         
                SQLSTING.AppendFormat(" SELECT (ISNULL(SUM(TB004),0)-ISNULL(SUM(TB005),0)) AS '預計顉料量', TB001 AS '單別',TB002 AS '單號'");
                SQLSTING.AppendFormat(" FROM [TK].dbo.MOCTB AS MOCTB");
                SQLSTING.AppendFormat(" LEFT JOIN [TK].dbo.CMSMC AS CMSMC ON TB009=MC001,TK..MOCTA AS MOCTA");
                SQLSTING.AppendFormat(" LEFT JOIN [TK].dbo.CMSMQ AS CMSMQ ON TA001=MQ001");
                SQLSTING.AppendFormat(" WHERE TA001=TB001 AND TA002=TB002 AND TB003='{0}'",MB001);
                SQLSTING.AppendFormat(" AND MC005='Y' AND TB018='Y' AND (TA011<>'Y' AND TA011<>'y') ");
                SQLSTING.AppendFormat(" AND MQ014='Y' AND (TB011='1' OR TB011='2') ");
                SQLSTING.AppendFormat(" AND TB004>TB005 ");
                SQLSTING.AppendFormat(" AND TB015<='{0}'",DATESTRING);
                SQLSTING.AppendFormat(" AND (TB009 in ('{0}') ) ",CMSMC);
                SQLSTING.AppendFormat(" GROUP BY TB001,TB002");
                SQLSTING.AppendFormat(" ");
                SQLSTING.AppendFormat(" ");
                SQLSTING.AppendFormat(" ");

                adapter = new SqlDataAdapter(SQLSTING.ToString(), sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds5.Clear();
                adapter.Fill(ds5, "TEMP5");
                sqlConn.Close();

                if (ds5.Tables["TEMP5"].Rows.Count == 0)
                {

                }
                else
                {
                    dataGridView5.DataSource = ds5.Tables["TEMP5"];
                    dataGridView5.AutoResizeColumns();
                    dataGridView5.CurrentCell = dataGridView5.Rows[rownum].Cells[0];

                }




            }
            catch
            {

            }
            finally
            {

            }

        }
        public void SearchCOPTD()
        {
            try
            {
                StringBuilder SQLSTING = new StringBuilder();


                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

           
                SQLSTING.AppendFormat(" select sum(TD008) AS '預計銷貨量',TD001 AS '單別',TD002 AS '單號'");
                SQLSTING.AppendFormat(" from [TK].dbo.COPTD COPTD");
                SQLSTING.AppendFormat(" left join [TK].dbo.CMSMC CMSMC on TD007=MC001");
                SQLSTING.AppendFormat(" where TD008+TD024+TD050>TD009+TD025+TD051   ");
                SQLSTING.AppendFormat(" and TD004='{0}' and TD021='Y' and TD016='N'",MB001);
                SQLSTING.AppendFormat(" and TD013<='{0}'  ",DATESTRING);
                SQLSTING.AppendFormat(" and (TD007 in ('{0}') )",CMSMC);
                SQLSTING.AppendFormat(" and MC005='Y'");
                SQLSTING.AppendFormat(" GROUP BY TD001,TD002");
                SQLSTING.AppendFormat(" ");
                SQLSTING.AppendFormat(" ");
                SQLSTING.AppendFormat(" ");

                adapter = new SqlDataAdapter(SQLSTING.ToString(), sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds6.Clear();
                adapter.Fill(ds6, "TEMP6");
                sqlConn.Close();

                if (ds6.Tables["TEMP6"].Rows.Count == 0)
                {

                }
                else
                {
                    dataGridView6.DataSource = ds6.Tables["TEMP6"];
                    dataGridView6.AutoResizeColumns();
                    dataGridView6.CurrentCell = dataGridView6.Rows[rownum].Cells[0];

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

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }

       
    }
}
