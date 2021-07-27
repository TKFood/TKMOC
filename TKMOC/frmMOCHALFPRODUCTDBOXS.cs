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
        DataSet ds1 = new DataSet();

        SqlCommand cmd = new SqlCommand();
        SqlTransaction tran;
        int result;

     

        string STATUS = "";

        public frmMOCHALFPRODUCTDBOXS()
        {
            InitializeComponent();
        }


        #region FUNCTION
        public void SEARCHMOCHALFPRODUCTDBOXS()
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();                

                if(string.IsNullOrEmpty(textBox1.Text))
                {
                    sbSql.AppendFormat(@"  
                                SELECT [MOCHALFPRODUCTDBOXS].[MB001] AS '品號',[MB002] AS '品名',[NUMS] AS '箱重',[BOXS] AS '箱數'
                                FROM [TKMOC].[dbo].[MOCHALFPRODUCTDBOXS],[TK].dbo.[INVMB]
                                WHERE [MOCHALFPRODUCTDBOXS].[MB001]=[INVMB].[MB001]
                                 ");
                }
                else if (!string.IsNullOrEmpty(textBox1.Text))
                {
                    sbSql.AppendFormat(@"  
                                SELECT [MOCHALFPRODUCTDBOXS].[MB001] AS '品號',[MB002] AS '品名',[NUMS] AS '箱重',[BOXS] AS '箱數'
                                FROM [TKMOC].[dbo].[MOCHALFPRODUCTDBOXS],[TK].dbo.[INVMB]
                                WHERE [MOCHALFPRODUCTDBOXS].[MB001]=[INVMB].[MB001]
                                AND ([MOCHALFPRODUCTDBOXS].[MB001] LIKE '%{0}%' OR [MB002] LIKE '%{0}%')
                                 ", textBox1.Text);
                }
               

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


        public void SETNULL1()
        {
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";

            textBox2.ReadOnly = false;
            //textBox3.ReadOnly = false;
            textBox4.ReadOnly = false;
        }

        public void SETNULL2()
        {
            textBox2.ReadOnly = true;
            //textBox3.ReadOnly = true;
            textBox4.ReadOnly = true;
        }
        public void SETNULL3()
        {
            textBox2.ReadOnly = false;
            //textBox3.ReadOnly = false;
            textBox4.ReadOnly = false;
        }

        public void ADDMOCHALFPRODUCTDBOXS(string MB001, string NUMS, string BOXS)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                
                sbSql.AppendFormat(@" 
                                    INSERT INTO[TKMOC].[dbo].[MOCHALFPRODUCTDBOXS]
                                    ([MB001],[NUMS],[BOXS])
                                    VALUES('{0}',{1},{2})
                                        ", MB001, NUMS,1);


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

        public void UPDATEMOCHALFPRODUCTDBOXS(string MB001, string NUMS)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@" 
                                   UPDATE [TKMOC].[dbo].[MOCHALFPRODUCTDBOXS]
                                    SET [NUMS]={1}
                                    WHERE [MB001]='{0}'
                                        ", MB001, NUMS);


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

        public void DELETEMOCHALFPRODUCTDBOXS(string MB001)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@" 
                                   DELETE [TKMOC].[dbo].[MOCHALFPRODUCTDBOXS]
                                    WHERE [MB001]='{0}'
                                        ", MB001);


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

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            textBox3.Text = SERCHINVMB(textBox2.Text.Trim());
        }

        public string SERCHINVMB(string MB001)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();

                sbSql.AppendFormat(@"  
                                    SELECT MB002 FROM [TK].dbo.INVMB WHERE MB001='{0}'
                                    ",MB001);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["ds1"].Rows[0]["MB002"].ToString().Trim();
                }
                else
                {
                    return "";
                }

            }
            catch
            {
                return "";
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
                    textBox2.Text = row.Cells["品號"].Value.ToString().Trim();
                    textBox4.Text = row.Cells["箱重"].Value.ToString().Trim();
                    textBox2.Text = row.Cells["品號"].Value.ToString().Trim();
                    textBox4.Text = row.Cells["箱重"].Value.ToString().Trim();




                }
                else
                {

                    
                }
            }
        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCHMOCHALFPRODUCTDBOXS();
        }
        




        private void button2_Click(object sender, EventArgs e)
        {
            SETNULL1();
            STATUS = "ADD";

        
        }

        private void button4_Click(object sender, EventArgs e)
        {
            SETNULL3();
            STATUS = "UPDATE";

           
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SETNULL2();

            if(STATUS.Equals("ADD"))
            {
                ADDMOCHALFPRODUCTDBOXS(textBox2.Text.Trim(), textBox4.Text.Trim(),"1");
            }
            else if (STATUS.Equals("UPDATE"))
            {
                UPDATEMOCHALFPRODUCTDBOXS(textBox2.Text.Trim(), textBox4.Text.Trim());
            }

            STATUS = "";
            SEARCHMOCHALFPRODUCTDBOXS();

        }

        private void button5_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除嗎?", "要刪除嗎?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELETEMOCHALFPRODUCTDBOXS(textBox2.Text.Trim());

            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }


            STATUS = "";
            SEARCHMOCHALFPRODUCTDBOXS();
        }


        #endregion

       
    }
}
