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

namespace TKMOC
{
    public partial class frmMOCMANULINECOP : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlDataAdapter adapter2 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder2 = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();

        int result;
        string ID;
        string DELID;
 

        public frmMOCMANULINECOP()
        {
            InitializeComponent();
        }
        public frmMOCMANULINECOP(string SUBID, string SUBBAR, string SUBNUM, string SUBBOX, string SUBPACKAGE)
        {
            InitializeComponent();

            ID = SUBID;
            SEARCHMOCMANULINE();
            SEARCHMOCMANULINECOP();
        }

        #region FUNCTION
        public void SEARCHMOCMANULINE()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  SELECT  [ID],[SERNO],[MANU],[MANUDATE],[MB001],[MB002]");
                sbSql.AppendFormat(@"  ,[MB003],[BAR],[NUM],[CLINET],[MANUHOUR],[BOX]");
                sbSql.AppendFormat(@"  ,[PACKAGE],[OUTDATE]");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINE]");
                sbSql.AppendFormat(@"  WHERE ID='{0}'", ID);
                sbSql.AppendFormat(@"  ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds1.Clear();
                adapter.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                {

                }
                else
                {
                    if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        SETVALUES();
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

        public void SETVALUES()
        {
            textBox1.Text = ds1.Tables["TEMPds1"].Rows[0]["MB002"].ToString();
            textBox2.Text = ds1.Tables["TEMPds1"].Rows[0]["BAR"].ToString();
            textBox3.Text = ds1.Tables["TEMPds1"].Rows[0]["NUM"].ToString();
            textBox4.Text = ds1.Tables["TEMPds1"].Rows[0]["BOX"].ToString();
            textBox5.Text = ds1.Tables["TEMPds1"].Rows[0]["PACKAGE"].ToString();
            textBox6.Text = ds1.Tables["TEMPds1"].Rows[0]["MB003"].ToString();
            textBox7.Text = ds1.Tables["TEMPds1"].Rows[0]["MANU"].ToString();
        }


        public void SEARCHMOCMANULINECOP()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [MANU] AS '組別',[TC001] AS '訂單單別',[TC002] AS '訂單單號',[SID] AS '來源',[ID]");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINECOP]");
                sbSql.AppendFormat(@"  WHERE [SID]='{0}'",ID);
                sbSql.AppendFormat(@"  ORDER BY [MANU],[TC001],[TC002]");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");


                adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                sqlConn.Open();
                ds2.Clear();
                adapter2.Fill(ds2, "TEMPds2");
                sqlConn.Close();


                if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                {

                }
                else
                {
                    if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView1.DataSource = ds2.Tables["TEMPds2"];
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

        public void ADDMOCMANULINECOP()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                if (!string.IsNullOrEmpty(textBox20.Text) && !string.IsNullOrEmpty(textBox21.Text))
                {
                 
                    sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[MOCMANULINECOP]");
                    sbSql.AppendFormat(" ([ID],[SID],[MANU],[TC001],[TC002])");
                    sbSql.AppendFormat(" VALUES({0},'{1}','{2}','{3}','{4}')", "NEWID()",ID,textBox7.Text,textBox20.Text,textBox21.Text);
                    sbSql.AppendFormat(" ");
                }



                sbSql.AppendFormat(" ");

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
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            DELID = null;

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    DELID = row.Cells["ID"].Value.ToString();
                }
                else
                {
                    DELID = null;

                }
            }
        }

        public void DELMOCMANULINECOP()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                if (!string.IsNullOrEmpty(DELID) )
                {

                    sbSql.AppendFormat(" DELETE [TKMOC].[dbo].[MOCMANULINECOP]");
                    sbSql.AppendFormat(" WHERE ID='{0}'",DELID);
                    sbSql.AppendFormat(" ");
                }



                sbSql.AppendFormat(" ");

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
        public void SETNULL()
        {
            textBox20.Text = null;
            textBox21.Text = null;
        }
        #endregion

        #region BUTTON
        private void button2_Click(object sender, EventArgs e)
        {
            ADDMOCMANULINECOP();
            SETNULL();
            SEARCHMOCMANULINECOP();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELMOCMANULINECOP();
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
            SEARCHMOCMANULINECOP();
        }


        #endregion

       
    }
}
