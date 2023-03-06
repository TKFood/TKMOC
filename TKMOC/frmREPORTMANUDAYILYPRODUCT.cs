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
using TKITDLL;

namespace TKMOC
{
    public partial class frmREPORTMANUDAYILYPRODUCT : Form
    {
        public frmREPORTMANUDAYILYPRODUCT()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SEARCH(string MANUDATE)
        {
            SqlConnection sqlConn = new SqlConnection();
            SqlCommand sqlComm = new SqlCommand();
           
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();

            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            DataSet ds1 = new DataSet();

            DateTime DT = Convert.ToDateTime(MANUDATE);

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
                                    CONVERT(NVARCHAR,[MANUDATE],112) AS '生產日'
                                    ,[MANU1PUR]  AS '製一組產能桶數'
                                    ,[MANU1ACT] AS '製一組預排桶數'
                                    ,[MANU2PUR] AS '製二組產能桶數'
                                    ,[MANU2ACT] AS '製二組預排桶數'
                                    ,[MANU3PUR] AS '手工產能'
                                    ,[MANU3ACT] AS '手工預排'
                                    ,[MANU4PUR] AS '外包產能'
                                    ,[MANU4ACT] AS '外包預排'
                                    FROM [TKMOC].[dbo].[MANUDAYILYPRODUCT]
                                    WHERE CONVERT(NVARCHAR,[MANUDATE],112) LIKE '{0}%'
                                    ORDER BY CONVERT(NVARCHAR,[MANUDATE],112)

                                    ", DT.ToString("yyyyMM"));

            
                    

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
            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];

                    textBox1.Text = row.Cells["生產日"].Value.ToString();
                    textBox2.Text = row.Cells["製一組產能桶數"].Value.ToString();
                    textBox3.Text = row.Cells["製一組預排桶數"].Value.ToString();
                    textBox4.Text = row.Cells["製二組產能桶數"].Value.ToString();
                    textBox5.Text = row.Cells["製二組預排桶數"].Value.ToString();
                    textBox6.Text = row.Cells["手工產能"].Value.ToString();
                    textBox7.Text = row.Cells["手工預排"].Value.ToString();
                    textBox8.Text = row.Cells["外包產能"].Value.ToString();
                    textBox9.Text = row.Cells["外包預排"].Value.ToString();



                }
                else
                {
                   
                }
            }
        }
        #endregion

        #region BUTTON
        private void button3_Click(object sender, EventArgs e)
        {
            SEARCH(dateTimePicker1.Value.ToString("yyyy/MM/dd"));
        }
        #endregion

      
    }
}
