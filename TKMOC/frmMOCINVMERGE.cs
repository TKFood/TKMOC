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
using System.Text.RegularExpressions;
using FastReport;
using FastReport.Data;
using TKITDLL;

namespace TKMOC
{
    public partial class frmMOCINVMERGE : Form
    {
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
        DataSet ds2 = new DataSet();
        string tablename = null;
        int rownum = 0;
        DataGridViewRow row;
        string SALSESID = null;
        int result;
        string query;

        public Report report1 { get; private set; }

        public frmMOCINVMERGE()
        {
            InitializeComponent();

            combobox2load();

            DataGridViewColumn dgvc = new DataGridViewCheckBoxColumn();
            dgvc.Width = 60;
            dgvc.Name = "選取";

            //新增到DataGridView內的第0欄
            this.dataGridView1.Columns.Insert(0, dgvc);
        }


        #region FUNCTION
        public void combobox2load()
        {

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            String Sequel = "SELECT MD002,MD001 FROM [TK].dbo.CMSMD WHERE MD003 IN ('20') ORDER BY MD001";
            SqlDataAdapter da = new SqlDataAdapter(Sequel, sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MD001", typeof(string));
            dt.Columns.Add("MD002", typeof(string));
            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "MD001";
            comboBox2.DisplayMember = "MD002";
            sqlConn.Close();



        }
        public void SERACH()
        {
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

                sbSql.AppendFormat(@"  SELECT TC001 AS '領料單',TC002 AS '單號',TC005  AS '線別' ");
                sbSql.AppendFormat(@"  FROM [TK].dbo.MOCTC");
                sbSql.AppendFormat(@"  WHERE TC003>='{0}' AND TC003<='{1}'",dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND TC005='{0}'",comboBox2.SelectedValue.ToString());
                sbSql.AppendFormat(@"  ORDER BY TC001,TC002,TC005");
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

            }
        }

        public void ADDDATA()
        {
            DataTable dt = new DataTable();
           
            // Declare DataColumn and DataRow variables.
            DataColumn column;
            DataRow row;
            DataView view;

            // Create new DataColumn, set DataType, ColumnName and add to DataTable.    
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "單別";
            dt.Columns.Add(column);

            // Create second column.
            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName = "單號";
            dt.Columns.Add(column);

            dt.Clear();

            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {
                if (dr.Cells[0].Value != null && (bool)dr.Cells[0].Value)
                {
                    //MessageBox.Show(dr.Cells[2].Value.ToString());

                    //建立一筆新的DataRow，並且等於新的dt row
                    row = dt.NewRow();

                    //指定每個欄位要儲存的資料
                    row["單別"] = dr.Cells[1].Value.ToString();
                    row["單號"] = dr.Cells[2].Value.ToString();

                    //新增資料至DataTable的dt內
                    dt.Rows.Add(row);
                }
            }


            if (dt.Rows.Count == 0)
            {
                dataGridView2.DataSource = null;
            }
            else if(dt.Rows.Count >=1)
            {
                dataGridView2.DataSource = dt;
            }

           
        }

        public void SETREPORT()
        {
            if(dataGridView2.Rows.Count>=1)
            {
                query = null;

                foreach (DataGridViewRow dr in this.dataGridView2.Rows)
                {
                    //MessageBox.Show(dr.Cells[0].Value.ToString()+ dr.Cells[1].Value.ToString());
                    query = query +"'" +dr.Cells[0].Value.ToString() + dr.Cells[1].Value.ToString() + "',";
                }

                query = query + "''";
            }

            string SQL;
            report1 = new Report();
            report1.Load(@"REPORT\tkmoc合併領料.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;


            TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;
            SQL = SETFASETSQL();
            Table.SelectCommand = SQL;
            report1.Preview = previewControl1;
            report1.Show();
        }

        public string SETFASETSQL()
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();

            if (comboBox1.Text.ToString().Equals("原料"))
            {
                FASTSQL.AppendFormat(@" SELECT MD002,TE004,TE017 ,TE011,TE012,SUM(TE005)  AS TE005,TE010 ");
                FASTSQL.AppendFormat(@" FROM [TK].dbo.CMSMD, [TK].dbo.MOCTC,[TK].dbo.MOCTE");
                FASTSQL.AppendFormat(@" WHERE MD002 LIKE '新%' ");
                FASTSQL.AppendFormat(@" AND MD001=TC005 ");
                FASTSQL.AppendFormat(@" AND TC001=TE001 AND TC002=TE002 ");
                FASTSQL.AppendFormat(@" AND ((TE004 LIKE '1%' ) OR (TE004 LIKE '3010000%' AND LEN(TE004)=10))   ");
                FASTSQL.AppendFormat(@" AND TC001+TC002 IN ({0})", query.ToString());
                FASTSQL.AppendFormat(@" GROUP BY MD002,TE004,TE017 ,TE011,TE012,TE010 ");
                FASTSQL.AppendFormat(@" ORDER BY MD002,TE004,TE017 ,TE011,TE012,TE010 ");

               
            }
            else if (comboBox1.Text.ToString().Equals("物料"))
            {
                FASTSQL.AppendFormat(@" SELECT MD002,TE004,TE017 ,TE011,TE012,SUM(TE005)  AS TE005,TE010 ");
                FASTSQL.AppendFormat(@" FROM [TK].dbo.CMSMD, [TK].dbo.MOCTC,[TK].dbo.MOCTE ");
                FASTSQL.AppendFormat(@" WHERE MD002 LIKE '新%' ");
                FASTSQL.AppendFormat(@" AND MD001=TC005 ");
                FASTSQL.AppendFormat(@" AND TC001=TE001 AND TC002=TE002 ");
                FASTSQL.AppendFormat(@" AND TE004 LIKE '2%' ");
                FASTSQL.AppendFormat(@" AND TC001+TC002 IN ({0})", query.ToString());
                FASTSQL.AppendFormat(@" GROUP BY MD002,TE004,TE017 ,TE011,TE012,TE010 ");
                FASTSQL.AppendFormat(@" ORDER BY MD002,TE004,TE017 ,TE011,TE012,TE010 ");
            }
            else if (comboBox1.Text.ToString().Equals("原料+物料"))
            {
                FASTSQL.AppendFormat(@" SELECT MD002,TE004,TE017 ,TE011,TE012,SUM(TE005) AS TE005,TE010 ");
                FASTSQL.AppendFormat(@" FROM [TK].dbo.CMSMD, [TK].dbo.MOCTC,[TK].dbo.MOCTE");
                FASTSQL.AppendFormat(@" WHERE MD002 LIKE '新%' ");
                FASTSQL.AppendFormat(@" AND MD001=TC005 ");
                FASTSQL.AppendFormat(@" AND TC001=TE001 AND TC002=TE002 ");
                FASTSQL.AppendFormat(@" AND (TE004 LIKE '1%' OR TE004 LIKE '2%' OR (TE004 LIKE '3010000%' AND LEN(TE004)=10))");
                FASTSQL.AppendFormat(@" AND TC001+TC002 IN ({0})", query.ToString());
                FASTSQL.AppendFormat(@" GROUP BY MD002,TE004,TE017 ,TE011,TE012,TE010 ");
                FASTSQL.AppendFormat(@" ORDER BY MD002,TE004,TE017 ,TE011,TE012,TE010 ");


            }


            FASTSQL.AppendFormat(@"  ");

            return FASTSQL.ToString();
        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView2.DataSource = null;

            SERACH();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            SETREPORT();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ADDDATA();
        }


        #endregion

       
    }
}
