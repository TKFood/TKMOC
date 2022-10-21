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
    public partial class frmNEWMOCTCMOCTD : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlCommand cmd = new SqlCommand();
       


        DataTable dt = new DataTable();
        SqlTransaction tran;
        int result;

        string tablename = null;
        int rownum = 0;


        public frmNEWMOCTCMOCTD()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SEARCH(string SDATES,string EDATES)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();

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
                                    MQ002 AS '單據'
                                    ,TC001 AS '單別'
                                    ,TC002 AS '單號'
                                    ,TC003 AS '日期'
                                    FROM [TK].dbo.MOCTC,[TK].dbo.CMSMQ
                                    WHERE TC001=MQ001
                                    AND TC003>='{0}' AND TC003<='{1}'
                                    ORDER BY TC001,TC002
                                    ", SDATES,EDATES);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds");
                sqlConn.Close();


                if (ds.Tables["TEMPds"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds.Tables["TEMPds"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView1.DataSource = ds.Tables["TEMPds"];
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

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    textBox1.Text = row.Cells["單別"].Value.ToString().Trim();
                    textBox2.Text = row.Cells["單號"].Value.ToString().Trim();

                    SEARCHMOCTE(textBox1.Text, textBox2.Text);

                }
                else
                {
                    SEARCHMOCTE("", "");

                }
            }

        }

        public void SEARCHMOCTE(string TE001,string TE002)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();

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
                                    TE003 AS '序號'
                                    ,TE004 AS '品號'
                                    ,TE017 AS '材料品名'
                                    ,TE005 AS '領退料數量'
                                    ,TE006 AS '單位'
                                    ,TE010 AS '批號'
                                    ,TE011 AS '製令'
                                    ,TE012 AS '製令號'
                                    FROM [TK].dbo.MOCTE
                                    WHERE TE001='{0}' AND TE002='{1}'
                                    ORDER BY TE003
                                    ", TE001, TE002);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds");
                sqlConn.Close();


                if (ds.Tables["TEMPds"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else
                {
                    if (ds.Tables["TEMPds"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView2.DataSource = ds.Tables["TEMPds"];
                        dataGridView2.AutoResizeColumns();
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

        public void ADDMOCTCMOCTDMOCTE(string TC001,string TC002)
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


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@" 

                                    ");



                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消


                }
                else
                {
                    tran.Commit();      //執行交易                    

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

        public string GETMAXTC002(string TC001,string TC003)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();
            string TC002;

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


                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds.Clear();
                                
                sbSql.AppendFormat(@"  
                                    SELECT ISNULL(MAX(TC002),'00000000000') AS TC002
                                    FROM [TK].[dbo].[MOCTC] 
                                    WHERE  TC001='{0}' AND TC003='{1}'
                                    ",TC001,TC003);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    return null;
                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        TC002 = SETTC002(ds.Tables["ds"].Rows[0]["TC002"].ToString(),dateTimePicker3.Value);
                        return TC002;

                    }
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }

        }
        public string SETTC002(string TC002,DateTime SDT)
        {
            if (TC002.Equals("00000000000"))
            {
                return SDT.ToString("yyyyMMdd") + "001";
            }

            else
            {
                int serno = Convert.ToInt16(TC002.Substring(8, 3));
                serno = serno + 1;
                string temp = serno.ToString();
                temp = temp.PadLeft(3, '0');
                return SDT.ToString("yyyyMMdd") + temp.ToString();
            }
        }


        #endregion

        #region BUTTON
        private void button3_Click(object sender, EventArgs e)
        {
            SEARCH(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
        }


        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要新增嗎?", "要新增嗎?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                string TC001 = textBox1.Text;
                string TC002 = GETMAXTC002(TC001, dateTimePicker3.Value.ToString("yyyyMMdd"));

                textBox3.Text = TC001;
                textBox4.Text = TC002;
                //ADDMOCTCMOCTDMOCTE();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }
        #endregion


    }
}
