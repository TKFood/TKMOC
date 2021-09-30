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
using TKITDLL;

namespace TKMOC
{
    public partial class frmMOCMANULINESubTEMPADD : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        SqlDataAdapter adapter2 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder2 = new SqlCommandBuilder();
        DataSet ds1 = new DataSet();
        DataSet TEMPds = new DataSet();


        public frmMOCMANULINESubTEMPADD()
        {
            InitializeComponent();
        }

        public frmMOCMANULINESubTEMPADD(string MB001,DataSet Fds)
        {
            InitializeComponent();

            textBox1.Text = MB001;
            
        }

        #region FUNCTION
        private void frmMOCMANULINESubTEMPADD_Load(object sender, EventArgs e)
        {
            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;      //奇數列顏色

            //先建立個 CheckBox 欄
            DataGridViewCheckBoxColumn cbCol = new DataGridViewCheckBoxColumn();
            cbCol.Width = 50;   //設定寬度
            cbCol.HeaderText = "　選擇";
            cbCol.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;   //置中
            cbCol.TrueValue = true;
            cbCol.FalseValue = false;
            dataGridView1.Columns.Insert(0, cbCol);
        }

        public DataSet SETDATASET
        {
            //set
            //{
            //    textBox1.Text = value;
            //}
            get
            {
                return TEMPds;
            }
        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            SERACH(textBox1.Text.Trim());
        }

        public void SERACH(string MB001)
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


                sbSql.AppendFormat(@"  SELECT ");
                sbSql.AppendFormat(@"  [MANU] AS '線別',CONVERT(varchar(100),[MANUDATE],112) AS '生產日',[MB001] AS '品號',[MB002] AS '品名' ");
                sbSql.AppendFormat(@"  ,[MB003] AS '規格',[BAR] AS '桶數',[NUM] AS '數量',[PACKAGE] AS '包裝數',[BOX] AS '箱數',[CLINET] AS '客戶',[OUTDATE] AS '交期',[TA029] AS '備註',[HALFPRO] AS '半成品數量'");
                sbSql.AppendFormat(@"  ,[COPTD001] AS '訂單單別',[COPTD002] AS '訂單號',[COPTD003] AS '訂單序號'");
                sbSql.AppendFormat(@"  ,[ID]");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINETEMP]");
                sbSql.AppendFormat(@"  WHERE  [MB001]='{0}'",MB001);
                sbSql.AppendFormat(@"  AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINE] )");
                sbSql.AppendFormat(@"  ORDER BY [SERNO]");
                sbSql.AppendFormat(@"  ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count == 0)
                {

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
                sqlConn.Close();
            }
        }

       

        public void SERACHQUERY(string ID)
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


                sbSql.AppendFormat(@"  SELECT ");
                sbSql.AppendFormat(@"  [MANU] AS '線別',CONVERT(varchar(100),[MANUDATE],112) AS '生產日',[MB001] AS '品號',[MB002] AS '品名' ");
                sbSql.AppendFormat(@"  ,[MB003] AS '規格',[BAR] AS '桶數',[NUM] AS '數量',[PACKAGE] AS '包裝數',[BOX] AS '箱數',[CLINET] AS '客戶',[OUTDATE] AS '交期',[TA029] AS '備註',[HALFPRO] AS '半成品數量'");
                sbSql.AppendFormat(@"  ,[COPTD001] AS '訂單單別',[COPTD002] AS '訂單號',[COPTD003] AS '訂單序號'");
                sbSql.AppendFormat(@"  ,[ID]");
                sbSql.AppendFormat(@"  FROM [TKMOC].[dbo].[MOCMANULINETEMP]");
                sbSql.AppendFormat(@"  WHERE  [ID] IN  ({0})", ID);
                sbSql.AppendFormat(@"  ORDER BY [SERNO]");
                sbSql.AppendFormat(@"  ");

                adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                sqlConn.Open();
                TEMPds.Clear();
                adapter2.Fill(TEMPds, "TEMPds");
                sqlConn.Close();


                if (TEMPds.Tables["TEMPds"].Rows.Count == 0)
                {

                }
                else
                {
                    if (TEMPds.Tables["TEMPds"].Rows.Count >= 1)
                    {
                        ////dataGridView1.Rows.Clear();
                        //dataGridView1.DataSource = TEMPds.Tables["TEMPds"];
                        //dataGridView1.AutoResizeColumns();
                        ////dataGridView1.CurrentCell = dataGridView1[0, rownum];


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
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            StringBuilder ID = new StringBuilder();

            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {
                if (dr.Cells[0].Value != null && (bool)dr.Cells[0].Value)
                {
                    ID.AppendFormat(@"'{0}',", dr.Cells["ID"].Value.ToString());

                }
            }

            ID.AppendFormat(@"'d22acdff-fee6-40f4-92cd-acce2a353749' ");

            SERACHQUERY(ID.ToString());
            //MessageBox.Show(ID.ToString());

            this.Close();
        }

        #endregion

       
    }
}
