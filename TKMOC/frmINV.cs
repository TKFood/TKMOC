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
        DataTable dt = new DataTable();
        string tablename = null;
        int rownum = 0;

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
            STR.AppendFormat(@"  ,(Select ISNULL(SUM(TB009),0) from [TK].dbo.PURTA AS PURTA, [TK].dbo.PURTB AS PURTB LEFT JOIN [TK].dbo.CMSMC AS CMSMC ON MC001=TB008 Where TA001=TB001 and TA002=TB002 and   TB025='Y' and TB021='N' and (TB022='' or TB022 IS NULL)  and TB004=LA001 AND MC005='Y' AND TB039='N' and TB019<='{0}'  and TB008='{1}' ) AS '預計請購量'", comboBox1.Text.ToString(), dateTimePicker1.Value.ToString("yyyyMMdd"));
            STR.AppendFormat(@"  ,(select ISNULL(SUM(TD008),0) from[TK].dbo.PURTD PURTD INNER JOIN [TK].dbo.PURTC PURTC ON TC001=TD001 AND TC002=TD002 and TD018='Y' and TD016='N' AND TD008>TD015 INNER JOIN [TK].dbo.CMSMC CMSMC ON TD007=MC001 where TD004=LA001   and MC005='Y' and TD012<='{0}'   and (TD007 in ('{1}') ) ) AS '預計進貨量'", comboBox1.Text.ToString(), dateTimePicker1.Value.ToString("yyyyMMdd"));
            STR.AppendFormat(@"  ,(SELECT MOUNT=ISNULL(SUM(TA015),0)-ISNULL(SUM(TA017),0)-ISNULL(SUM(TA018),0)  FROM [TK].dbo.MOCTA AS MOCTA LEFT JOIN [TK].dbo.CMSMC AS CMSMC ON TA020=MC001 LEFT JOIN [TK].dbo.CMSMQ AS CMSMQ ON TA001=MQ001 WHERE TA006=LA001 AND MC005='Y' AND TA013='Y' AND (TA011<>'Y' AND TA011<>'y')  AND MQ014='Y'  AND TA010<='{0}' AND (TA020 in ('{1}') ) ) AS '預計生產量'", comboBox1.Text.ToString(), dateTimePicker1.Value.ToString("yyyyMMdd"));
            STR.AppendFormat(@"  ,(SELECT MOUNT=ISNULL(SUM(TB004),0)-ISNULL(SUM(TB005),0) FROM [TK].dbo.MOCTB AS MOCTB LEFT JOIN [TK].dbo.CMSMC AS CMSMC ON TB009=MC001,TK..MOCTA AS MOCTA LEFT JOIN [TK].dbo.CMSMQ AS CMSMQ ON TA001=MQ001 WHERE TA001=TB001 AND TA002=TB002 AND TB003=LA001 AND MC005='Y' AND TB018='Y' AND (TA011<>'Y' AND TA011<>'y')  AND MQ014='Y' AND (TB011='1' OR TB011='2')  AND TB004>TB005  AND TB015<='{0}' AND (TB009 in ('{1}') ) )  AS '預計領料量'", comboBox1.Text.ToString(), dateTimePicker1.Value.ToString("yyyyMMdd"));
            STR.AppendFormat(@"  ,(select sum(TD008) from [TK].dbo.COPTD COPTD left join [TK].dbo.CMSMC CMSMC on TD007=MC001 where TD008+TD024+TD050>TD009+TD025+TD051 and  TD004=LA001 and TD021='Y' and TD016='N'  and TD013<='{0}'   and (TD007 in ('{1}') )  and MC005='Y') AS '預計銷貨量' ", comboBox1.Text.ToString(), dateTimePicker1.Value.ToString("yyyyMMdd"));
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
        #endregion

        #region BUTTON

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }
    }
}
