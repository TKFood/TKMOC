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
        SqlTransaction tran;
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        int rownum = 0;
        int result;
        Report report1 = new Report();

        //找出生產說明用的品號
        string MAINMB001 = "";

        public frmREPROTMOCTABBAKE()
        {
            InitializeComponent();

            comboBox1load();
            ADD_DATAGRID_CHECKED();
        }

        private void frmREPROTMOCTABBAKE_Load(object sender, EventArgs e)
        {
            //先建立個 CheckBox 欄
            DataGridViewCheckBoxColumn cbCol = new DataGridViewCheckBoxColumn();
            cbCol.Width = 50;   //設定寬度
            cbCol.HeaderText = "　選擇";
            cbCol.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;   //置中
            cbCol.TrueValue = true;
            cbCol.FalseValue = false;
            dataGridView2.Columns.Insert(0, cbCol);
                      

            //建立个矩形，等下计算 CheckBox 嵌入 GridView 的位置
            Rectangle rect = dataGridView2.GetCellDisplayRectangle(0,  -1, true);
            rect.X = rect.Location.X + rect.Width / 8 - 1;
            rect.Y = rect.Location.Y + (rect.Height / 4 - 1);

            CheckBox cbHeader = new CheckBox();
            cbHeader.Name = "checkboxHeader";
            cbHeader.Size = new Size(12, 12);
            cbHeader.Location = rect.Location;

            //全选要设定的事件
            cbHeader.CheckedChanged += new EventHandler(cbHeader_CheckedChanged);

            //将 CheckBox 加入到 dataGridView
            dataGridView2.Controls.Add(cbHeader);

        }

        private void cbHeader_CheckedChanged(object sender, EventArgs e)
        {
            dataGridView1.EndEdit();

            foreach (DataGridViewRow dr in dataGridView1.Rows)
            {
                dr.Cells[0].Value = ((CheckBox)dataGridView1.Controls.Find("checkboxHeader", true)[0]).Checked;

            }

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

        public void Search()
        {
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
                                    TA001 AS '製令',TA002 AS '單號',TA006 AS '品號',TA034 AS '品名',TA015 AS '生產量',TA003 AS '生產日',TA035 AS '規格',MC004 AS '標準批量',(TA015/MC004)  AS '桶數'
                                    ,ISNULL([NUMS],0) AS '每桶量'
                                    ,(CASE WHEN [NUMS]>0 AND TA015>0 THEN CONVERT(decimal(16,4),TA015/[NUMS]) ELSE 1 END ) AS '新桶數'
                                    FROM [TK].dbo.MOCTA,[TK].dbo.BOMMC
                                    LEFT JOIN [TKMOC].[dbo].[REPORTMOCBOMBAKINGBUCKETS] ON [REPORTMOCBOMBAKINGBUCKETS].MB001=BOMMC.MC001
                                    WHERE TA006=MC001
                                    AND (TA006 LIKE '3%' OR TA006 LIKE '4%')
                                    AND TA021 IN ('08')
                                    AND TA003='{0}'
                                    ORDER BY TA001,TA002 
                                     
                                    ", dateTimePicker1.Value.ToString("yyyyMMdd"));

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();

                if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                {

                    dataGridView2.DataSource = ds.Tables["TEMPds1"];

                    dataGridView2.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);
                    dataGridView2.DefaultCellStyle.Font = new Font("Tahoma", 10);
                    dataGridView2.Columns["製令"].Width = 60;
                    dataGridView2.Columns["單號"].Width = 100;
                    dataGridView2.Columns["品號"].Width = 100;
                    dataGridView2.Columns["品名"].Width = 120;
                }
                else
                {
                    dataGridView2.DataSource = null;
                }


            }
            catch
            {

            }
            finally
            {

            }
        }
        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];
                    textBox1.Text = row.Cells["製令"].Value.ToString().Trim();
                    textBox2.Text = row.Cells["單號"].Value.ToString().Trim();
                    textBox3.Text = row.Cells["新桶數"].Value.ToString().Trim();
                    textBox4.Text = row.Cells["品號"].Value.ToString().Trim();
                    textBox5.Text = row.Cells["每桶量"].Value.ToString().Trim();

                    MAINMB001 = row.Cells["品號"].Value.ToString().Trim();

                }
                else
                {
                    textBox1.Text = "";
                    textBox2.Text = "";
                    textBox3.Text = "";
                    textBox4.Text = "";
                    textBox5.Text = "";

                }
            }
        }

        public DataTable CHECK_BOMMD(int COUNTS, List<string> MD001)
        {
            //MessageBox.Show(COUNTS.ToString());
            int SETCOUNT = 1;
            DataSet ds = new DataSet();
            StringBuilder SQL = new StringBuilder();

            if (COUNTS >= 1)
            {
                foreach (string MD001STR in MD001)
                {
                    if (SETCOUNT == 1)
                    {
                        SQL.AppendFormat(@"
                                        SELECT MD003, MD006, MD007,COUNT(MD003) COUNTS
                                        FROM (
                                            SELECT MD001, MD003, MD006, MD007
                                            FROM [TK].dbo.BOMMD
                                            WHERE MD003 LIKE '1%' AND MD001 IN ('{0}')
                                    ", MD001STR);
                    }
                    else
                    {
                        SQL.AppendFormat(@"
                                            UNION ALL
                                            SELECT MD001, MD003, MD006, MD007
                                            FROM [TK].dbo.BOMMD
                                            WHERE MD003 LIKE '1%' AND MD001 IN ('{0}')
                                    ", MD001STR);
                    }

                    SETCOUNT = SETCOUNT + 1;
                }

                SQL.AppendFormat(@"                                           
                                    ) AS CombinedData
                                    WHERE MD003 NOT IN (SELECT [MD003]  FROM [TKMOC].[dbo].[REPORTMOCBOMNOSET])
                                    GROUP BY MD003, MD006, MD007
                
                                    HAVING COUNT(MD003)<{0}
                                    ", COUNTS);
            }

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

                sbSql = SQL;

                sbSql.AppendFormat(@"                                   
                                     
                                    ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();

                if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    return ds.Tables["TEMPds1"];
                }
                else
                {
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

        public void SETREPORT_BAKING_MERGE(string TA001, string TA002, float BUCKETSORI, string LINK_TA001TA002, string LINK_TA006, string LINK_TA034, string MAINMB001,decimal BUCKETSNUMS)
        {
            bool CHECKFLOOR = IsIntegerFloor(BUCKETSORI);


            if (BUCKETSORI > 0)
            {
                if (CHECKFLOOR == true)
                {
                    ADD_REPORTMOCBOMBAKING(TA001, TA002, BUCKETSORI.ToString(), BUCKETSNUMS);
                    //MessageBox.Show(CHECKFLOOR  + BUCKETSORI.ToString());
                }
                else
                {
                    ADD_REPORTMOCBOMBAKING_ODD(TA001, TA002, BUCKETSORI.ToString(), BUCKETSNUMS);
                    //MessageBox.Show(CHECKFLOOR  + BUCKETSORI.ToString());
                }


            }


            StringBuilder SQL = new StringBuilder();
            StringBuilder SQL2B = new StringBuilder();

            SQL = SETSQL2(LINK_TA001TA002, LINK_TA006, LINK_TA034);
            SQL2B = SETSQL2B(MAINMB001);

            report1 = new Report();
            report1.Load(@"REPORT\烘培原料添加表V1.frx");

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
            table.SelectCommand = SQL.ToString();
            TableDataSource table1 = report1.GetDataSource("Table1") as TableDataSource;
            table1.SelectCommand = SQL2B.ToString();

            report1.Preview = previewControl2;
            report1.Show();
        }

        public static bool IsIntegerFloor(float f)
        {
            return f == Math.Floor(f);
        }

        /// <summary>
        /// 剛好滿桶數，沒有未滿桶
        /// </summary>
        /// <param name="TA001"></param>
        /// <param name="TA002"></param>
        /// <param name="BUCKETS"></param>
        public void ADD_REPORTMOCBOMBAKING(string TA001, string TA002, string BUCKETS,decimal BUCKETSNUMS)
        {
            float BUCKETSFLOAT = float.Parse(BUCKETS);
            int COUNTS = Convert.ToInt32(Math.Ceiling(BUCKETSFLOAT));

            //MessageBox.Show(COUNTS.ToString());


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

                sbSql.AppendFormat(@" DELETE [TKMOC].[dbo].[REPORTMOCBOMBAKING]");
                sbSql.AppendFormat(@" ");

                for (int i = 1; i <= COUNTS; i++)
                {
                    sbSql.AppendFormat(@"
                                            INSERT INTO [TKMOC].[dbo].[REPORTMOCBOMBAKING]
                                            ([TA001],[TA002],[TA006],[TA034],[BOXS],[MD003],[MB002],[MD006])
                                            SELECT TA001,TA002,TA006,TA034,{2},MD003,MB002,MD006*{3}
                                            FROM [TK].dbo.MOCTA,[TK].dbo.BOMMD,[TK].dbo.INVMB
                                            WHERE TA006=MD001
                                            AND MD003=MB001
                                           
                                            AND TA001='{0}' AND TA002='{1}'
                                            ORDER BY MD003

                                           ", TA001, TA002, i, BUCKETSNUMS);
                }

           

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

        /// <summary>
        /// 未滿桶，第1桶是滿的、第2桶是未滿、其他滿桶
        /// </summary>
        /// <param name="TA001"></param>
        /// <param name="TA002"></param>
        /// <param name="BUCKETS"></param>
        public void ADD_REPORTMOCBOMBAKING_ODD(string TA001, string TA002, string BUCKETS,decimal BUCKETSNUMS)
        {
            float BUCKETSFLOAT = float.Parse(BUCKETS);
            int COUNTS = Convert.ToInt32(Math.Ceiling(BUCKETSFLOAT));
            decimal BUCKETSSMAILL = Convert.ToDecimal(BUCKETSFLOAT - (COUNTS - 1));

            //處理負數
            //BUCKETSFLOAT>0 && BUCKETSFLOAT<1，只有1未滿桶
            //BUCKETSFLOAT>1正常

            if (BUCKETSFLOAT > 0 && BUCKETSFLOAT < 1)
            {
                BUCKETSSMAILL = Convert.ToDecimal(BUCKETSFLOAT);
                COUNTS = 0;
            }
            else if (BUCKETSFLOAT > 1)
            {
                COUNTS = COUNTS;
                BUCKETSSMAILL = BUCKETSSMAILL;
            }


            //MessageBox.Show(BUCKETSFLOAT.ToString()+" "+ COUNTS.ToString()+" "+ BUCKETSSMAILL.ToString());

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

                sbSql.AppendFormat(@" DELETE [TKMOC].[dbo].[REPORTMOCBOMBAKING]");
                sbSql.AppendFormat(@"  ");

                if (COUNTS == 0)
                {
                    sbSql.AppendFormat(@"       
                                            INSERT INTO [TKMOC].[dbo].[REPORTMOCBOMBAKING]
                                            ([TA001],[TA002],[TA006],[TA034],[BOXS],[MD003],[MB002],[MD006])                                            
                                            SELECT TA001,TA002,TA006,TA034,{2},MD003,MB002,CONVERT(DECIMAL(16,3),MD006*{3}*{4})
                                            FROM [TK].dbo.MOCTA,[TK].dbo.BOMMD,[TK].dbo.INVMB
                                            WHERE TA006=MD001
                                            AND MD003=MB001
                                            
                                            AND TA001='{0}' AND TA002='{1}'
                                            ORDER BY MD003

                                           ", TA001, TA002, 1, BUCKETSSMAILL, BUCKETSNUMS);
                }
                else if (COUNTS >= 1)
                {
                    sbSql.AppendFormat(@"       
                                            INSERT INTO [TKMOC].[dbo].[REPORTMOCBOMBAKING]
                                            ([TA001],[TA002],[TA006],[TA034],[BOXS],[MD003],[MB002],[MD006])
                                            SELECT TA001,TA002,TA006,TA034,{2},MD003,MB002,MD006*{3}
                                            FROM [TK].dbo.MOCTA,[TK].dbo.BOMMD,[TK].dbo.INVMB
                                            WHERE TA006=MD001
                                            AND MD003=MB001
                                          
                                            AND TA001='{0}' AND TA002='{1}'
                                            ORDER BY MD003

                                           ", TA001, TA002, 1, BUCKETSNUMS);
                    sbSql.AppendFormat(@"       
                                            INSERT INTO [TKMOC].[dbo].[REPORTMOCBOMBAKING]
                                            ([TA001],[TA002],[TA006],[TA034],[BOXS],[MD003],[MB002],[MD006])                                            
                                            SELECT TA001,TA002,TA006,TA034,{2},MD003,MB002,CONVERT(DECIMAL(16,3),MD006*{3}*{4})
                                            FROM [TK].dbo.MOCTA,[TK].dbo.BOMMD,[TK].dbo.INVMB
                                            WHERE TA006=MD001
                                            AND MD003=MB001
                                            
                                            AND TA001='{0}' AND TA002='{1}'
                                            ORDER BY MD003

                                           ", TA001, TA002, 2, BUCKETSSMAILL, BUCKETSNUMS);

                    for (int i = 3; i <= COUNTS; i++)
                    {
                        sbSql.AppendFormat(@"
                                            INSERT INTO [TKMOC].[dbo].[REPORTMOCBOMBAKING]
                                            ([TA001],[TA002],[TA006],[TA034],[BOXS],[MD003],[MB002],[MD006])
                                            SELECT TA001,TA002,TA006,TA034,{2},MD003,MB002,MD006*{3}
                                            FROM [TK].dbo.MOCTA,[TK].dbo.BOMMD,[TK].dbo.INVMB
                                            WHERE TA006=MD001
                                            AND MD003=MB001
                                           
                                            AND TA001='{0}' AND TA002='{1}'
                                            ORDER BY MD003

                                           ", TA001, TA002, i, BUCKETSNUMS);
                    }
                }
                
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
        public StringBuilder SETSQL2(string LINK_TA001TA002, string TA006, string TA034)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
                                
                            SELECT 
                            '{0}' AS '製令'
                            ,'第'+CONVERT(nvarchar,[BOXS])+'桶' AS '桶數'
                            ,'{1}' AS '成品'
                            ,'{2}' AS '成品名'
                            ,[MD003] AS '品號'
                            ,[MB002] AS '品名'
                            ,[MD006] AS '重量'
                            ,'' AS '複核'
                            ,'' AS '油酥'
                            ,'' AS '檢查麵粉袋的麵粉線頭'
                            ,(SELECT TOP 1 TE010 FROM [TK].dbo.MOCTE WHERE TE011=[TA001] AND TE012=[TA002] AND TE004=[MD003]) AS 'A製造  B有效'
                            ,'' AS '外觀:攪拌均勻度、軟硬度'
                            ,'' AS '攪拌時間  始'
                            ,'' AS '攪拌時間  終'
                            ,'' AS '投 料 人'
                            ,'' AS '對 點 人'
                            ,'' AS '單位幹部'
                            ,'' AS '品質判定'
                            ,'' AS '換線清潔檢查'
                            ,BOMMC.UDF01 AS 'BOM備註(邊料'
                            ,BOMMC.UDF02 AS 'BOM備註(餅麩'
                            ,BOMMC.UDF06 AS '單顆重'
                            ,(SELECT SUM([MD006]) FROM [TKMOC].[dbo].[REPORTMOCBOMBAKING] RE WHERE [MD003] NOT  IN ('101001009','3010000111') AND RE.[BOXS]=[REPORTMOCBOMBAKING].[BOXS]) AS '每桶重'
                            ,(SELECT SUM([MD006]) FROM [TKMOC].[dbo].[REPORTMOCBOMBAKING] WHERE [MD003] NOT  IN ('101001009','3010000111')  ) AS '總重'
                            ,CASE WHEN BOMMC.UDF06=0 THEN 1 ELSE BOMMC.UDF06 END 
                            ,'顆數:'+CONVERT(nvarchar,((SELECT SUM([MD006]) FROM [TKMOC].[dbo].[REPORTMOCBOMBAKING] RE WHERE [MD003] NOT  IN ('101001009','3010000111') AND RE.[BOXS]=[REPORTMOCBOMBAKING].[BOXS])/(CASE WHEN BOMMC.UDF06=0 THEN 1 ELSE BOMMC.UDF06 END))) AS '每桶顆數'
                            ,((SELECT SUM([MD006]) FROM [TKMOC].[dbo].[REPORTMOCBOMBAKING] WHERE [MD003] NOT  IN ('101001009','3010000111') )/(CASE WHEN BOMMC.UDF06=0 THEN 1 ELSE BOMMC.UDF06 END)) AS '總顆數'
                            ,'配比 '+ BOMMC.UDF03 AS '配方比'

                            FROM [TKMOC].[dbo].[REPORTMOCBOMBAKING]
                            LEFT JOIN [TK].dbo.BOMMC ON MC001=TA006
                            WHERE 1=1 
                            ORDER BY [TA001],[TA002],[BOXS],[MD003]
     

                            ", LINK_TA001TA002, TA006, TA034);



            return SB;
        }

        public StringBuilder SETSQL2B(string MAINMB001)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
                                
                           SELECT TOP 1
                            [PROCESSING]
                            FROM [TKMOC].[dbo].[REPORTMOCBOMPROCESS]
                            WHERE MB001 LIKE '%{0}%'
     

                            ", MAINMB001);



            return SB;
        }

        public void SETREPORT_BAKING(string TA001, string TA002, string BUCKETS, string MAINMB001,decimal BUCKETSNUMS)
        {
            float BUCKETSORI = float.Parse(BUCKETS);
            bool CHECKFLOOR = IsIntegerFloor(BUCKETSORI);


            if (!string.IsNullOrEmpty(BUCKETS) && BUCKETSORI > 0)
            {
                if (CHECKFLOOR == true)
                {
                    ADD_REPORTMOCBOMBAKING(TA001, TA002, BUCKETS, BUCKETSNUMS);
                    //MessageBox.Show(CHECKFLOOR  + BUCKETSORI.ToString());
                }
                else
                {
                    ADD_REPORTMOCBOMBAKING_ODD(TA001, TA002, BUCKETS, BUCKETSNUMS);
                    //MessageBox.Show(CHECKFLOOR  + BUCKETSORI.ToString());
                }


            }


            StringBuilder SQL = new StringBuilder();
            StringBuilder SQL1B = new StringBuilder();

            SQL = SETSQL();
            SQL1B = SETSQL1B(MAINMB001);

            report1 = new Report();
            report1.Load(@"REPORT\烘培原料添加表V1.frx");

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
            table.SelectCommand = SQL.ToString();
            TableDataSource table1 = report1.GetDataSource("Table1") as TableDataSource;
            table1.SelectCommand = SQL1B.ToString();

            report1.Preview = previewControl2;
            report1.Show();
        }

        public StringBuilder SETSQL()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
                                
                            SELECT 
                            [TA001]+[TA002] AS '製令'
                            ,'第'+CONVERT(nvarchar,[BOXS])+'桶' AS '桶數'
                            ,TA006 AS '成品'
                            ,TA034 AS '成品名'
                            ,[MD003] AS '品號'
                            ,[MB002] AS '品名'
                            ,[MD006] AS '重量'
                            ,'' AS '複核'
                            ,'' AS '油酥'
                            ,'' AS '檢查麵粉袋的麵粉線頭'
                            ,(SELECT TOP 1 TE010 FROM [TK].dbo.MOCTE WHERE TE011=[TA001] AND TE012=[TA002] AND TE004=[MD003]) AS 'A製造  B有效'
                            ,'' AS '外觀:攪拌均勻度、軟硬度'
                            ,'' AS '攪拌時間  始'
                            ,'' AS '攪拌時間  終'
                            ,'' AS '投 料 人'
                            ,'' AS '對 點 人'
                            ,'' AS '單位幹部'
                            ,'' AS '品質判定'
                            ,'' AS '換線清潔檢查'
                            ,BOMMC.UDF01 AS 'BOM備註(邊料'
                            ,BOMMC.UDF02 AS 'BOM備註(餅麩'
                            ,BOMMC.UDF06 AS '單顆重'
                            ,(SELECT SUM([MD006]) FROM [TKMOC].[dbo].[REPORTMOCBOMBAKING] RE WHERE [MD003] NOT  IN ('101001009','3010000111') AND RE.[BOXS]=[REPORTMOCBOMBAKING].[BOXS]) AS '每桶重'
                            ,(SELECT SUM([MD006]) FROM [TKMOC].[dbo].[REPORTMOCBOMBAKING] WHERE [MD003] NOT  IN ('101001009','3010000111')  ) AS '總重'
                            ,CASE WHEN BOMMC.UDF06=0 THEN 1 ELSE BOMMC.UDF06 END 
                            ,'顆數:'+CONVERT(nvarchar,((SELECT SUM([MD006]) FROM [TKMOC].[dbo].[REPORTMOCBOMBAKING] RE WHERE [MD003] NOT  IN ('101001009','3010000111') AND RE.[BOXS]=[REPORTMOCBOMBAKING].[BOXS])/(CASE WHEN BOMMC.UDF06=0 THEN 1 ELSE BOMMC.UDF06 END))) AS '每桶顆數'
                            ,((SELECT SUM([MD006]) FROM [TKMOC].[dbo].[REPORTMOCBOMBAKING] WHERE [MD003] NOT  IN ('101001009','3010000111') )/(CASE WHEN BOMMC.UDF06=0 THEN 1 ELSE BOMMC.UDF06 END)) AS '總顆數'
                            ,'配比 '+ BOMMC.UDF03 AS '配方比'                            

                            FROM [TKMOC].[dbo].[REPORTMOCBOMBAKING]
                            LEFT JOIN [TK].dbo.BOMMC ON MC001=TA006
                            WHERE 1=1
                            ORDER BY [TA001],[TA002],[BOXS],[MD003]
     

                            ");



            return SB;
        }
        public StringBuilder SETSQL1B(string MAINMB001)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
                                
                            SELECT TOP 1
                            [PROCESSING]
                            FROM [TKMOC].[dbo].[REPORTMOCBOMPROCESS]
                            WHERE MB001 LIKE '%{0}%'

                            ", MAINMB001);



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
            Search();
        }
        private void button4_Click(object sender, EventArgs e)
        {
            //SETREPORT(textBox1.Text.Trim(), textBox2.Text.Trim(),textBox3.Text.Trim());
            int COUNTS = 0;
            List<string> MD001 = new List<string>();
            DataTable DT = null;
            string MESS = "";

            string CHECKED = "N";
            string TA001 = "";
            string TA002 = "";
            string LINK_TA001TA002 = "";
            string LINK_TA006 = "";
            string LINK_TA034 = "";
            string TEMP = "";
            float BUCKETS = 0;
            decimal BUCKETSNUMS = 1;

            //預設每桶量
            BUCKETSNUMS = Convert.ToDecimal(textBox5.Text);
            if (BUCKETSNUMS > 0)
            {
                BUCKETSNUMS = BUCKETSNUMS;
            }
            else
            {
                BUCKETSNUMS = 1;
            }


            //
            foreach (DataGridViewRow dr in this.dataGridView2.Rows)
            {
                try
                {
                    if (dr.Cells[0].Value != null && (bool)dr.Cells[0].Value)
                    {
                        COUNTS = COUNTS + 1;
                        MD001.Add(dr.Cells[3].Value.ToString());
                    }
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(ex.Message);
                }
            }

            //CHECK  原料需單身品號元件一致、組成用量一致
            DT = CHECK_BOMMD(COUNTS, MD001);

            foreach (DataGridViewRow dr in this.dataGridView2.Rows)
            {
                try
                {
                    if (dr.Cells[0].Value != null && (bool)dr.Cells[0].Value)
                    {
                        CHECKED = "Y";

                        TA001 = dr.Cells["製令"].Value.ToString();
                        TA002 = dr.Cells["單號"].Value.ToString();
                        LINK_TA001TA002 = LINK_TA001TA002 + TA001 + TA002 + "*";
                        LINK_TA006 = LINK_TA034 + dr.Cells["品號"].Value.ToString() + "*";
                        LINK_TA034 = LINK_TA034 + dr.Cells["品名"].Value.ToString() + "*";
                        BUCKETS = BUCKETS + float.Parse(dr.Cells["桶數"].Value.ToString());
                        BUCKETS = (float)Math.Round(BUCKETS, 3);

                        

                        MAINMB001 = dr.Cells["品號"].Value.ToString();
                    }
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(ex.Message);
                }
            }

            if (CHECKED.Equals("Y"))
            {
                if (DT == null)
                {
                    SETREPORT_BAKING_MERGE(TA001, TA002, BUCKETS, LINK_TA001TA002, LINK_TA006, LINK_TA034, MAINMB001, BUCKETSNUMS);
                }
                else
                {
                    MESS = "原料需單身品號元件不一致 或 組成用量不一致，不能合併\n";
                    foreach (DataRow ROW in DT.Rows)
                    {
                        // 每一行都是一個 DataRow                       
                        MESS = MESS + "品號:" + ROW["MD003"].ToString();
                        MESS = MESS + "用量:" + ROW["MD006"].ToString();
                        MESS = MESS + "底數:" + ROW["MD007"].ToString();

                        MESS = MESS + "\n";
                    }
                    MessageBox.Show(MESS.ToString());
                }

            }
            else if (CHECKED.Equals("N"))
            {
                SETREPORT_BAKING(textBox1.Text.Trim(), textBox2.Text.Trim(), textBox3.Text.Trim(), MAINMB001, BUCKETSNUMS);
            }
        }


        #endregion

       
    }
}
