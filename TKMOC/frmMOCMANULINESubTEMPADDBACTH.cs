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

namespace TKMOC
{
    public partial class frmMOCMANULINESubTEMPADDBACTH : Form
    {
        private ComponentResourceManager _ResourceManager = new ComponentResourceManager();
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        StringBuilder sbSqlQuery2 = new StringBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();
        int result;

        public frmMOCMANULINESubTEMPADDBACTH()
        {
            InitializeComponent();

            comboBox1load();
        }

        #region FUNCTION
        private void frmMOCMANULINESubTEMPADDBACTH_Load(object sender, EventArgs e)
        {
            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;      //奇數列顏色
            

            //先建立個 CheckBox 欄
            DataGridViewCheckBoxColumn cbCol = new DataGridViewCheckBoxColumn();
            cbCol.Width = 120;   //設定寬度
            cbCol.HeaderText = "　選擇";
            cbCol.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;   //置中
            cbCol.TrueValue = true;
            cbCol.FalseValue = false;
            dataGridView1.Columns.Insert(0, cbCol);

           
            //建立个矩形，等下计算 CheckBox 嵌入 GridView 的位置
            Rectangle rect = dataGridView1.GetCellDisplayRectangle(0, -1, true);
            rect.X = rect.Location.X + rect.Width / 4 - 18;
            rect.Y = rect.Location.Y + (rect.Height / 2 - 9);

            CheckBox cbHeader = new CheckBox();
            cbHeader.Name = "checkboxHeader";
            cbHeader.Size = new Size(18, 18);
            cbHeader.Location = rect.Location;

            //全选要设定的事件
            cbHeader.CheckedChanged += new EventHandler(cbHeader_CheckedChanged);

            //将 CheckBox 加入到 dataGridView
            dataGridView1.Controls.Add(cbHeader);

        }

        private void cbHeader_CheckedChanged(object sender, EventArgs e)
        {
            dataGridView1.EndEdit();

            foreach (DataGridViewRow dr in dataGridView1.Rows)
            {
                dr.Cells[0].Value = ((CheckBox)dataGridView1.Controls.Find("checkboxHeader", true)[0]).Checked;

            }

        }

        public void comboBox1load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT MD001,MD002 FROM [TK].dbo.CMSMD    WHERE MD002 LIKE '新廠%'   ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MD001", typeof(string));
            dt.Columns.Add("MD002", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "MD002";
            comboBox1.DisplayMember = "MD002";
            sqlConn.Close();


        }
        public void SEARCHCOPTD(string TD001,string TD002)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            DataSet ds1 = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@" 
                                    SELECT 
                                    TD001 AS '訂單',TD002 AS '訂單號',TD003 AS '訂單序號',TD013 AS '生產日',TD004 AS '品號'
                                    ,TD005 AS '品名',TD006 AS '規格',(CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024)  END ) AS '數量',(TD008+TD024) AS '箱數',(CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024)  END ) AS '包裝數'
                                    ,(TD008+TD024) AS '桶數',TC053 AS '客戶',TD013 AS '預交日',0 AS '工時',(TC015+'-'+TD020) '備註'
                                    ,0 AS '半成品','' AS TID,'' AS TCOPTD001,'' AS TCOPTD002,'' AS TCOPTD003
                                    FROM [TK].dbo.COPTC,[TK].dbo.COPTD
                                    LEFT JOIN [TK].dbo.INVMD ON INVMD.MD001=TD004 AND TD010=INVMD.MD002
                                    LEFT JOIN [TK].dbo.BOMMD ON BOMMD.MD003 LIKE '201%' AND BOMMD.MD007>1 AND BOMMD.MD001=TD004
                                    LEFT JOIN [TK].dbo.BOMMC ON MC001=TD004
                                    WHERE TC001=TD001 AND TC002=TD002 
                                    AND TD001='{0}' AND TD002='{1}'
                                    AND TD001+TD002+TD003 NOT IN (SELECT COPTD001+COPTD002+COPTD003 FROM [TKMOC].dbo.MOCMANULINETEMP)"
                                 , TD001, TD002);
               

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                {

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

        public void ADDMOCMANULINETEMP(
            Guid ID,  string MANU, string MANUDATE, string MB001
            , string MB002, string MB003, string BAR, string NUM, string CLINET
            , string MANUHOUR, string BOX, string PACKAGE, string OUTDATE, string TA029
            , string HALFPRO, string COPTD001, string COPTD002, string COPTD003, string TID
            , string TCOPTD001, string TCOPTD002, string TCOPTD003
            )
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[MOCMANULINETEMP]");
                sbSql.AppendFormat(" (");
                sbSql.AppendFormat(" [ID],[MANU],[MANUDATE],[MB001]");
                sbSql.AppendFormat(" ,[MB002],[MB003],[BAR],[NUM],[CLINET]");
                sbSql.AppendFormat(" ,[MANUHOUR],[BOX],[PACKAGE],[OUTDATE],[TA029]");
                sbSql.AppendFormat(" ,[HALFPRO],[COPTD001],[COPTD002],[COPTD003]");
                sbSql.AppendFormat(" ,[TCOPTD001],[TCOPTD002],[TCOPTD003]");
                sbSql.AppendFormat(" )");
                sbSql.AppendFormat(" VALUES");
                sbSql.AppendFormat(" (");
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}',", ID.ToString(), MANU, MANUDATE, MB001);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", MB002, MB003, BAR, NUM, CLINET);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}','{4}',", MANUHOUR, BOX, PACKAGE, OUTDATE, TA029);
                sbSql.AppendFormat(" '{0}','{1}','{2}','{3}',", HALFPRO, COPTD001, COPTD002, COPTD003);
                sbSql.AppendFormat(" '{0}','{1}','{2}'", TCOPTD001, TCOPTD002, TCOPTD003);
                sbSql.AppendFormat(" )");
                sbSql.AppendFormat(" ");
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
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCHCOPTD(textBox1.Text.Trim(),textBox2.Text.Trim());
        }

        private void button2_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {
                if (dr.Cells[0].Value != null && (bool)dr.Cells[0].Value)
                {
                    Guid NEWGUID = new Guid();
                    NEWGUID = Guid.NewGuid();
                    string SERNO = dr.Cells[""].Value.ToString().Trim();
                    string MANU = comboBox1.Text.Trim();
                    string MANUDATE = dateTimePicker1.Value.ToString("yyyyMMdd");
                    string MB001 = dr.Cells["品號"].Value.ToString().Trim();
                    string MB002 = dr.Cells["品名"].Value.ToString().Trim();
                    string MB003 = dr.Cells["規格"].Value.ToString().Trim();
                    
                    string CLINET = dr.Cells["客戶"].Value.ToString().Trim();
                    string MANUHOUR = dr.Cells["工時"].Value.ToString().Trim();
                   
                    string OUTDATE = dr.Cells["預交日"].Value.ToString().Trim();
                    string TA029 = dr.Cells["備註"].Value.ToString().Trim();
                    string HALFPRO = dr.Cells["半成品"].Value.ToString().Trim();
                    string COPTD001 = dr.Cells["訂單"].Value.ToString().Trim();
                    string COPTD002 = dr.Cells["訂單號"].Value.ToString().Trim();
                    string COPTD003 = dr.Cells["訂單序號"].Value.ToString().Trim();
                    string TID = dr.Cells["TID"].Value.ToString().Trim();
                    string TCOPTD001 = dr.Cells["TCOPTD001"].Value.ToString().Trim();
                    string TCOPTD002 = dr.Cells["TCOPTD002"].Value.ToString().Trim();
                    string TCOPTD003 = dr.Cells["TCOPTD003"].Value.ToString().Trim();

                    string BAR = "0";
                    string NUM = "0";
                    string BOX = "0";
                    string PACKAGE = "0";

                    if (comboBox1.Text.Equals("新廠包裝線"))
                    {
                        BAR = "0";
                        NUM = dr.Cells["數量"].Value.ToString().Trim();
                        BOX = dr.Cells["箱數"].Value.ToString().Trim();
                        PACKAGE = dr.Cells["包裝數"].Value.ToString().Trim();
                    }
                    else
                    {
                        BAR = dr.Cells["桶數"].Value.ToString().Trim();
                        NUM = dr.Cells["數量"].Value.ToString().Trim();
                        BOX = "0";
                        PACKAGE = "0";
                    }

                    ADDMOCMANULINETEMP(NEWGUID,  MANU, MANUDATE, MB001, MB002, MB003, BAR, NUM, CLINET, MANUHOUR, BOX, PACKAGE, OUTDATE, TA029, HALFPRO, COPTD001, COPTD002, COPTD003, TID, TCOPTD001, TCOPTD002, TCOPTD003);
                }
            }

            MessageBox.Show("完成");
            this.Close();
        }
        #endregion


    }
}
