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
    public partial class frmUPDATEMOCTA : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        SqlDataAdapter adapter2 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder2 = new SqlCommandBuilder();
       
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();
       
        DataTable dt = new DataTable();
        SqlTransaction tran;
        int result;

        string tablename = null;
        int rownum = 0;

        string TA033;



        public frmUPDATEMOCTA()
        {
            InitializeComponent();
        }

        #region FUNCTION
        private void frmUPDATEMOCTA_Load(object sender, EventArgs e)
        {
            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;      //奇數列顏色

            //先建立個 CheckBox 欄
            DataGridViewCheckBoxColumn cbCol = new DataGridViewCheckBoxColumn();
            cbCol.Width = 50;   //設定寬度
            cbCol.HeaderText = "選擇";
            cbCol.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;   //置中
            cbCol.TrueValue = true;
            cbCol.FalseValue = false;
            dataGridView1.Columns.Insert(0, cbCol);

            #region 建立全选 CheckBox

            ////建立个矩形，等下计算 CheckBox 嵌入 GridView 的位置
            //Rectangle rect = dataGridView3.GetCellDisplayRectangle(0, -1, true);
            //rect.X = rect.Location.X + rect.Width / 4 - 18;
            //rect.Y = rect.Location.Y + (rect.Height / 2 - 9);

            //CheckBox cbHeader = new CheckBox();
            //cbHeader.Name = "checkboxHeader";
            //cbHeader.Size = new Size(18, 18);
            //cbHeader.Location = rect.Location;

            ////全选要设定的事件
            //cbHeader.CheckedChanged += new EventHandler(cbHeader_CheckedChanged);

            ////将 CheckBox 加入到 dataGridView
            //dataGridView3.Controls.Add(cbHeader);


            #endregion
        }

        public void SEARCH()
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

                sbSql.AppendFormat(@"  SELECT TA001 AS '製令',TA002 AS '單號',TA003 AS '生產日',TA006 AS '品號',TA034 AS '品名',TA015 AS '生產量',TA033 AS '計劃批號'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.MOCTA");
                sbSql.AppendFormat(@"  WHERE TA003>='{0}' AND TA003<='{1}' ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND TA001='A510'");
                sbSql.AppendFormat(@"  ORDER BY TA001,TA002,TA003");
                sbSql.AppendFormat(@"  ");

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

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            SETNULL();


            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];

                    textBox1.Text = row.Cells["製令"].Value.ToString().Trim();
                    textBox2.Text = row.Cells["單號"].Value.ToString().Trim();
                    textBox3.Text = row.Cells["計劃批號"].Value.ToString().Trim();
                  

                }
                else
                {
                    SETNULL();

                }
            }
        }

        public void SETNULL()
        {
            textBox1.Text = null;
            textBox2.Text = null;
            textBox3.Text = null;
            textBox4.Text = null;

        }

        public void UPDATETA033()
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


                sbSql.AppendFormat(" UPDATE [TK].dbo.MOCTA SET TA033='{0}'", textBox4.Text);
                sbSql.AppendFormat(" WHERE TA001='{0}' AND TA002='{1}'", textBox1.Text, textBox2.Text);
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
                    MessageBox.Show("完成");
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

        public void CHANGEMULTI()
        {
            StringBuilder TA001TA002 = new StringBuilder();
            TA001TA002.Clear();

            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {
                if (dr.Cells[0].Value != null && (bool)dr.Cells[0].Value)
                {
                    if(!string.IsNullOrEmpty(dr.Cells["製令"].Value.ToString())&& !string.IsNullOrEmpty(dr.Cells["單號"].Value.ToString()))
                    {
                        TA001TA002.AppendFormat(@" '{0}{1}',", dr.Cells["製令"].Value.ToString().Trim(), dr.Cells["單號"].Value.ToString().Trim());
                     
                    }
                }
             
            }

            TA001TA002.AppendFormat(@" 'A311'");

            UPDATEMULTITA033(TA001TA002.ToString());
            //MessageBox.Show(TA001TA002.ToString());

        }

        public void UPDATEMULTITA033(string TA001TA002)
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


                sbSql.AppendFormat(" UPDATE [TK].dbo.MOCTA SET TA033='{0}'", textBox4.Text);
                sbSql.AppendFormat(" WHERE TA001+TA002 IN ({0})", TA001TA002);
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
                    MessageBox.Show("完成");
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

        public string GETMAXMOCTATA033(DateTime dt)
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


                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds1.Clear();

                sbSql.AppendFormat(@" 
                            SELECT ISNULL(MAX(TA033),'00000000000') AS TA033
                            FROM [TK].[dbo].[MOCTA] 
                            WHERE [TA033] LIKE '{0}%'
                            ", dt.ToString("yyyyMMdd"));

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count == 0)
                {
                    return null;
                }
                else
                {
                    if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        TA033 = SETMAXNO(dt, ds1.Tables["ds1"].Rows[0]["TA033"].ToString());
                        return TA033;

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

        public string SETMAXNO(DateTime dt, string NO)
        {
            if (NO.Equals("00000000000"))
            {
                return dt.ToString("yyyyMMdd") + "001";
            }

            else
            {
                int serno = Convert.ToInt16(NO.Substring(8, 3));
                serno = serno + 1;
                string temp = serno.ToString();
                temp = temp.PadLeft(3, '0');
                return dt.ToString("yyyyMMdd") + temp.ToString();
            }


        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCH();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            CHANGEMULTI();
            SEARCH();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            textBox4.Text = GETMAXMOCTATA033(DateTime.Now);
        }

        #endregion


    }
}
