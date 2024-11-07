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
using FastReport;
using FastReport.Data;
using TKITDLL;

namespace TKMOC
{
    public partial class frmREPROTMOCTABBAKE : Form
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
        Report report1 = new Report();

        public frmREPROTMOCTABBAKE()
        {
            InitializeComponent();

            comboBox1load();
            ADD_DATAGRID_CHECKED();
        }

        #region FUNCTION
        public void ADD_DATAGRID_CHECKED()
        {
            DataGridViewColumn dgvc = new DataGridViewCheckBoxColumn();
            dgvc.Width = 60;
            dgvc.Name = "選取";

            //新增到DataGridView內的第0欄
            this.dataGridView1.Columns.Insert(0, dgvc);
        }
        public void comboBox1load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT [KIND],[PARAID],[PARANAME] FROM [TKMOC].[dbo].[TBPARA] WHERE [KIND]='BAKE'  ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("PARAID", typeof(string));
            dt.Columns.Add("PARANAME", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "PARAID";
            comboBox1.DisplayMember = "PARAID";
            sqlConn.Close();


        }

        public void SERACH(string TA001, string TA003)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"    
                                   SELECT 
                                    TA001 AS '製令',TA002 AS '製令單',TA003 AS '生產日',TA006 AS '生產品號',MB1.MB002 AS '生產品名',TA015 AS '生產量',TA007 AS '生產單位'

                                    ,(YEAR(TA003)-1911) AS 'YEARS',MONTH(TA003) AS 'MONTHS',DAY(TA003) AS 'DAYS'
                                    FROM [TK].dbo.MOCTA
                                    LEFT JOIN [TK].dbo.INVMB MB1 ON MB1.MB001=TA006

                                    WHERE TA001='{0}'
                                    AND TA003='{1}'
                                    ORDER BY TA001,TA002,TA006

                                    ", TA001, TA003);

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

        public string ADD_QUERY_TA001TA002()
        {
            DataRow row;
            StringBuilder QUERY_TA001TA002 = new StringBuilder();

            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {
                if (dr.Cells[0].Value != null && (bool)dr.Cells[0].Value)
                {
                    QUERY_TA001TA002.AppendFormat(@" '{0}'," , dr.Cells[1].Value.ToString()+ dr.Cells[2].Value.ToString());

                    //MessageBox.Show(dr.Cells[1].Value.ToString() + dr.Cells[2].Value.ToString());
                } 
            } 
             
            QUERY_TA001TA002.AppendFormat(@" ''");

            return QUERY_TA001TA002.ToString();
        }


        public void SETFASTREPORT(string QUERY_TA001TA002)
        {
            StringBuilder SQL1 = new StringBuilder();
            StringBuilder SQL2 = new StringBuilder();

            SQL1 = SETSQL1(QUERY_TA001TA002);
             
            Report report1 = new Report();
            report1.Load(@"REPORT\原物料添加表-烘焙V2.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;


            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL1(string QUERY_TA001TA002)
        {
            StringBuilder SB = new StringBuilder();

           
            SB.AppendFormat(@" 
                            SELECT 
                            TA001 AS '製令',TA002 AS '製令單',TA003 AS '生產日',TA006 AS '生產品號',MB1.MB002 AS '生產品名',TA015 AS '生產量',TA007 AS '生產單位'
                            ,TB003 AS '原/物料品號',MB2.MB002 AS '原/物料品名',TB004 AS '需領料數量',TB007 AS '領料單位'
                            ,(YEAR(TA003)-1911) AS 'YEARS',MONTH(TA003) AS 'MONTHS',DAY(TA003) AS 'DAYS'
                            ,(CASE WHEN TB007 IN ('KG','kg','kG','Kg') THEN TB004*1000 ELSE TB004 END ) AS 'NEW需領料數量'
                            ,(CASE WHEN TB007 IN ('KG','kg','kG','Kg') THEN 'g' ELSE TB007 END ) AS 'NEW領料單位'

                            FROM [TK].dbo.MOCTA
                            LEFT JOIN [TK].dbo.INVMB MB1 ON MB1.MB001=TA006
                            ,[TK].dbo.MOCTB
                            LEFT JOIN [TK].dbo.INVMB MB2 ON MB2.MB001=TB003
                            WHERE TA001=TB001 AND TA002=TB002
                            AND TA001+TA002 IN ({0})
                            ORDER BY TA001,TA002,TA006,TB003

                            ", QUERY_TA001TA002);

            return SB;

        }

        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            SERACH(comboBox1.Text.ToString(), dateTimePicker1.Value.ToString("yyyyMMdd"));

            //SETFASTREPORT(comboBox1.Text.ToString(),dateTimePicker1.Value.ToString("yyyyMMdd"));
        }
        private void button2_Click(object sender, EventArgs e)
        {
            string QUERY_TA001TA002 = ADD_QUERY_TA001TA002();

            if (!string.IsNullOrEmpty(QUERY_TA001TA002))
            {
                SETFASTREPORT(QUERY_TA001TA002);
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {

        }
        #endregion


    }
}
