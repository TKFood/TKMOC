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
using TKITDLL;

namespace TKMOC
{
    public partial class frmREPORTMOCBOM : Form
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
        SqlDataAdapter adapter3 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder3 = new SqlCommandBuilder();
        SqlCommand cmd = new SqlCommand();
     
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        int rownum = 0;

        SqlTransaction tran;

        DataSet ds1 = new DataSet();
        int result;

        Report report1 = new Report();

        public frmREPORTMOCBOM()
        {
            InitializeComponent();
        }

        private void frmREPORTMOCBOM_Load(object sender, EventArgs e)
        {
            //先建立個 CheckBox 欄
            DataGridViewCheckBoxColumn cbCol = new DataGridViewCheckBoxColumn();
            cbCol.Width = 50;   //設定寬度
            cbCol.HeaderText = "　選擇";
            cbCol.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;   //置中
            cbCol.TrueValue = true;
            cbCol.FalseValue = false;
            dataGridView1.Columns.Insert(0, cbCol);

            #region 建立全选 CheckBox

            //建立个矩形，等下计算 CheckBox 嵌入 GridView 的位置
            Rectangle rect = dataGridView1.GetCellDisplayRectangle(0, -1, true);
            rect.X = rect.Location.X + rect.Width / 8- 1;
            rect.Y = rect.Location.Y + (rect.Height / 4 - 1);

            CheckBox cbHeader = new CheckBox();
            cbHeader.Name = "checkboxHeader";
            cbHeader.Size = new Size(12, 12);
            cbHeader.Location = rect.Location;

            //全选要设定的事件
            cbHeader.CheckedChanged += new EventHandler(cbHeader_CheckedChanged);

            //将 CheckBox 加入到 dataGridView
            dataGridView1.Controls.Add(cbHeader);


            #endregion


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
                                    SELECT TA001 AS '製令',TA002 AS '單號',TA006 AS '品號',TA034 AS '品名',TA015 AS '生產量',TA003 AS '生產日',TA035 AS '規格',MC004 AS '標準批量',(TA015/MC004)  AS '桶數'
                                    FROM [TK].dbo.MOCTA,[TK].dbo.BOMMC
                                    WHERE TA006=MC001
                                    AND (TA006 LIKE '3%' OR TA006 LIKE '4%')
                                    AND TA021 IN ('02','03','04')
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

                    dataGridView1.DataSource = ds.Tables["TEMPds1"];

                    dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);
                    dataGridView1.DefaultCellStyle.Font = new Font("Tahoma", 10);
                    dataGridView1.Columns["製令"].Width = 60;
                    dataGridView1.Columns["單號"].Width = 100;
                    dataGridView1.Columns["品號"].Width = 100;
                    dataGridView1.Columns["品名"].Width = 120;
                }
                else
                {
                    dataGridView1.DataSource = null;
                }
              

            }
            catch
            {

            }
            finally
            {

            }
        }

        public void SearchOIL()
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
                                    SELECT MD002 AS '線別',TA006 AS '品號',TA034 AS '品名',SUM(TA015) AS '總生產量',TA003 AS '生產日',TA035 AS '規格',MC004 AS '標準批量',(SUM(TA015)/MC004)  AS '桶數',TA021 '線別代號'
                                    FROM [TK].dbo.MOCTA,[TK].dbo.BOMMC,[TK].dbo.CMSMD
                                    WHERE TA006=MC001
                                    AND TA021=MD001
                                    AND (TA006 LIKE '3%' OR TA006 LIKE '4%')
                                    AND TA021 IN ('02','03','04')
                                    AND TA003='{0}'
                                    GROUP BY MD002,TA006,TA034,TA015,TA003,TA035,MC004,TA021
                                    ORDER BY MD002,TA006,TA034,TA015,TA003,TA035,MC004,TA021
                                     
                                    ", dateTimePicker2.Value.ToString("yyyyMMdd"));

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
                    dataGridView2.Columns["線別"].Width = 100;
                    dataGridView2.Columns["品號"].Width = 100;
                    dataGridView1.Columns["品名"].Width = 120;
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

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    textBox1.Text = row.Cells["製令"].Value.ToString().Trim();
                    textBox2.Text = row.Cells["單號"].Value.ToString().Trim();
                    textBox3.Text = row.Cells["桶數"].Value.ToString().Trim();



                }
                else
                {
                    textBox1.Text = "";
                    textBox2.Text = "";
                    textBox3.Text = "";

                }
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
                    textBox4.Text = row.Cells["生產日"].Value.ToString().Trim();
                    textBox5.Text = row.Cells["品號"].Value.ToString().Trim();
                    textBox6.Text = row.Cells["桶數"].Value.ToString().Trim();
                    textBox7.Text = row.Cells["線別代號"].Value.ToString().Trim();
                    
                }
                else
                {
                    textBox4.Text = "";
                    textBox5.Text = "";
                    textBox6.Text = "";
                    textBox7.Text = "";

                }
            }

        }

        public void SETREPORT(string TA001,string TA002,string BUCKETS)
        {
            float BUCKETSORI = float.Parse(BUCKETS);
            bool CHECKFLOOR = IsIntegerFloor(BUCKETSORI);


            if (!string.IsNullOrEmpty(BUCKETS)&& BUCKETSORI > 0)
            {
                if(CHECKFLOOR==true)
                {
                    ADDTOREPORTMOCBOM(TA001, TA002, BUCKETS);
                    //MessageBox.Show(CHECKFLOOR  + BUCKETSORI.ToString());
                }
                else
                {
                    ADDTOREPORTMOCBOMODD(TA001, TA002, BUCKETS);
                    //MessageBox.Show(CHECKFLOOR  + BUCKETSORI.ToString());
                }
                
                
            }


            StringBuilder SQL = new StringBuilder();

            SQL = SETSQL();

            report1 = new Report();
            report1.Load(@"REPORT\油酥原料添加表V6.frx");

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

            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
                                
                            SELECT [ID]
                            ,[TA001]+[TA002] AS '製令'
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
                            ,(SELECT SUM([MD006]) FROM [TKMOC].[dbo].[REPORTMOCBOM] RE WHERE [MD003] NOT  IN ('101001009','3010000111') AND RE.[BOXS]=[REPORTMOCBOM].[BOXS]) AS '每桶重'
                            ,(SELECT SUM([MD006]) FROM [TKMOC].[dbo].[REPORTMOCBOM] WHERE [MD003] NOT  IN ('101001009','3010000111')  ) AS '總重'
                            ,CASE WHEN BOMMC.UDF06=0 THEN 1 ELSE BOMMC.UDF06 END 
                            ,'顆數:'+CONVERT(nvarchar,((SELECT SUM([MD006]) FROM [TKMOC].[dbo].[REPORTMOCBOM] RE WHERE [MD003] NOT  IN ('101001009','3010000111') AND RE.[BOXS]=[REPORTMOCBOM].[BOXS])/(CASE WHEN BOMMC.UDF06=0 THEN 1 ELSE BOMMC.UDF06 END))) AS '每桶顆數'
                            ,((SELECT SUM([MD006]) FROM [TKMOC].[dbo].[REPORTMOCBOM] WHERE [MD003] NOT  IN ('101001009','3010000111') )/(CASE WHEN BOMMC.UDF06=0 THEN 1 ELSE BOMMC.UDF06 END)) AS '總顆數'
                            ,'配比 '+ BOMMC.UDF03 AS '配方比'                            

                            FROM [TKMOC].[dbo].[REPORTMOCBOM]
                            LEFT JOIN [TK].dbo.BOMMC ON MC001=TA006
                            WHERE [MD003] NOT  IN ('101001009','3010000111')   
                            ORDER BY [TA001],[TA002],[BOXS],[MD003]
     

                            ");



            return SB;
        }

        public void SETREPORT2(string TA001,string TA002,float BUCKETSORI,string LINK_TA001TA002,string LINK_TA006, string LINK_TA034)
        {
            bool CHECKFLOOR = IsIntegerFloor(BUCKETSORI);


            if (BUCKETSORI > 0)
            {
                if (CHECKFLOOR == true)
                {
                    ADDTOREPORTMOCBOM(TA001, TA002, BUCKETSORI.ToString());
                    //MessageBox.Show(CHECKFLOOR  + BUCKETSORI.ToString());
                }
                else
                {
                    ADDTOREPORTMOCBOMODD(TA001, TA002, BUCKETSORI.ToString());
                    //MessageBox.Show(CHECKFLOOR  + BUCKETSORI.ToString());
                }


            }


            StringBuilder SQL = new StringBuilder();

            SQL = SETSQL2(LINK_TA001TA002, LINK_TA006, LINK_TA034);

            report1 = new Report();
            report1.Load(@"REPORT\油酥原料添加表V6.frx");

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

            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL2(string LINK_TA001TA002, string TA006, string TA034)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
                                
                            SELECT [ID]
                            ,'{0}' AS '製令'
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
                            ,(SELECT SUM([MD006]) FROM [TKMOC].[dbo].[REPORTMOCBOM] RE WHERE [MD003] NOT  IN ('101001009','3010000111') AND RE.[BOXS]=[REPORTMOCBOM].[BOXS]) AS '每桶重'
                            ,(SELECT SUM([MD006]) FROM [TKMOC].[dbo].[REPORTMOCBOM] WHERE [MD003] NOT  IN ('101001009','3010000111')  ) AS '總重'
                            ,CASE WHEN BOMMC.UDF06=0 THEN 1 ELSE BOMMC.UDF06 END 
                            ,'顆數:'+CONVERT(nvarchar,((SELECT SUM([MD006]) FROM [TKMOC].[dbo].[REPORTMOCBOM] RE WHERE [MD003] NOT  IN ('101001009','3010000111') AND RE.[BOXS]=[REPORTMOCBOM].[BOXS])/(CASE WHEN BOMMC.UDF06=0 THEN 1 ELSE BOMMC.UDF06 END))) AS '每桶顆數'
                            ,((SELECT SUM([MD006]) FROM [TKMOC].[dbo].[REPORTMOCBOM] WHERE [MD003] NOT  IN ('101001009','3010000111') )/(CASE WHEN BOMMC.UDF06=0 THEN 1 ELSE BOMMC.UDF06 END)) AS '總顆數'
                            ,'配比 '+ BOMMC.UDF03 AS '配方比'

                            FROM [TKMOC].[dbo].[REPORTMOCBOM]
                            LEFT JOIN [TK].dbo.BOMMC ON MC001=TA006
                            WHERE [MD003] NOT  IN ('101001009','3010000111')   
                            ORDER BY [TA001],[TA002],[BOXS],[MD003]
     

                            ", LINK_TA001TA002,TA006, TA034);



            return SB;
        }
        /// <summary>
        /// 剛好滿桶數，沒有未滿桶
        /// </summary>
        /// <param name="TA001"></param>
        /// <param name="TA002"></param>
        /// <param name="BUCKETS"></param>
        public void ADDTOREPORTMOCBOM(string TA001, string TA002, string BUCKETS)
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

                sbSql.AppendFormat(@" DELETE [TKMOC].[dbo].[REPORTMOCBOM]");
                sbSql.AppendFormat(@" ");

                for (int i = 1; i <= COUNTS; i++)
                    {
                        sbSql.AppendFormat(@"
                                            INSERT INTO [TKMOC].[dbo].[REPORTMOCBOM]
                                            ([TA001],[TA002],[TA006],[TA034],[BOXS],[MD003],[MB002],[MD006])
                                            SELECT TA001,TA002,TA006,TA034,{2},MD003,MB002,MD006
                                            FROM [TK].dbo.MOCTA,[TK].dbo.BOMMD,[TK].dbo.INVMB
                                            WHERE TA006=MD001
                                            AND MD003=MB001
                                            AND (MD003 LIKE '1%' OR MD003 LIKE '301%')
                                            AND MD003 NOT LIKE '301400%'
                                            AND MB002 NOT LIKE '%水麵%'
                                            AND TA001='{0}' AND TA002='{1}'
                                            ORDER BY MD003

                                           ", TA001,TA002,i);
                    }

                ////新增5個空白格子
                //sbSql.AppendFormat(@"
                //                            INSERT INTO [TKMOC].[dbo].[REPORTMOCBOM]
                //                            ([TA001],[TA002],[TA006],[TA034],[BOXS],[MD003],[MB002],[MD006])
                //                            SELECT TA001,TA002,TA006,TA034,{2},MD003,MB002,0
                //                            FROM [TK].dbo.MOCTA,[TK].dbo.BOMMD,[TK].dbo.INVMB
                //                            WHERE TA006=MD001
                //                            AND MD003=MB001
                //                            AND (MD003 LIKE '1%' OR MD003 LIKE '301%')
                //                            AND MD003 NOT LIKE '301400%'
                //                            AND MB002 NOT LIKE '%水麵%'
                //                            AND TA001='{0}' AND TA002='{1}'
                //                            ORDER BY MD003

                //                           ", TA001, TA002, COUNTS+1);

                //sbSql.AppendFormat(@"
                //                            INSERT INTO [TKMOC].[dbo].[REPORTMOCBOM]
                //                            ([TA001],[TA002],[TA006],[TA034],[BOXS],[MD003],[MB002],[MD006])
                //                            SELECT TA001,TA002,TA006,TA034,{2},MD003,MB002,0
                //                            FROM [TK].dbo.MOCTA,[TK].dbo.BOMMD,[TK].dbo.INVMB
                //                            WHERE TA006=MD001
                //                            AND MD003=MB001
                //                            AND (MD003 LIKE '1%' OR MD003 LIKE '301%')
                //                            AND MD003 NOT LIKE '301400%'
                //                            AND MB002 NOT LIKE '%水麵%'
                //                            AND TA001='{0}' AND TA002='{1}'
                //                            ORDER BY MD003

                //                           ", TA001, TA002, COUNTS + 2);

                //sbSql.AppendFormat(@"
                //                            INSERT INTO [TKMOC].[dbo].[REPORTMOCBOM]
                //                            ([TA001],[TA002],[TA006],[TA034],[BOXS],[MD003],[MB002],[MD006])
                //                            SELECT TA001,TA002,TA006,TA034,{2},MD003,MB002,0
                //                            FROM [TK].dbo.MOCTA,[TK].dbo.BOMMD,[TK].dbo.INVMB
                //                            WHERE TA006=MD001
                //                            AND MD003=MB001
                //                            AND (MD003 LIKE '1%' OR MD003 LIKE '301%')
                //                            AND MD003 NOT LIKE '301400%'
                //                            AND MB002 NOT LIKE '%水麵%'
                //                            AND TA001='{0}' AND TA002='{1}'
                //                            ORDER BY MD003

                //                           ", TA001, TA002, COUNTS + 3);

                //sbSql.AppendFormat(@"
                //                            INSERT INTO [TKMOC].[dbo].[REPORTMOCBOM]
                //                            ([TA001],[TA002],[TA006],[TA034],[BOXS],[MD003],[MB002],[MD006])
                //                            SELECT TA001,TA002,TA006,TA034,{2},MD003,MB002,0
                //                            FROM [TK].dbo.MOCTA,[TK].dbo.BOMMD,[TK].dbo.INVMB
                //                            WHERE TA006=MD001
                //                            AND MD003=MB001
                //                            AND (MD003 LIKE '1%' OR MD003 LIKE '301%')
                //                            AND MD003 NOT LIKE '301400%'
                //                            AND MB002 NOT LIKE '%水麵%'
                //                            AND TA001='{0}' AND TA002='{1}'
                //                            ORDER BY MD003

                //                           ", TA001, TA002, COUNTS + 4);

                //sbSql.AppendFormat(@"
                //                            INSERT INTO [TKMOC].[dbo].[REPORTMOCBOM]
                //                            ([TA001],[TA002],[TA006],[TA034],[BOXS],[MD003],[MB002],[MD006])
                //                            SELECT TA001,TA002,TA006,TA034,{2},MD003,MB002,0
                //                            FROM [TK].dbo.MOCTA,[TK].dbo.BOMMD,[TK].dbo.INVMB
                //                            WHERE TA006=MD001
                //                            AND MD003=MB001
                //                            AND (MD003 LIKE '1%' OR MD003 LIKE '301%')
                //                            AND MD003 NOT LIKE '301400%'
                //                            AND MB002 NOT LIKE '%水麵%'
                //                            AND TA001='{0}' AND TA002='{1}'
                //                            ORDER BY MD003

                //                           ", TA001, TA002, COUNTS + 5);


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
        public void ADDTOREPORTMOCBOMODD(string TA001, string TA002, string BUCKETS)
        {
            float BUCKETSFLOAT = float.Parse(BUCKETS);
            int COUNTS = Convert.ToInt32(Math.Ceiling(BUCKETSFLOAT));
            decimal BUCKETSSMAILL = Convert.ToDecimal(BUCKETSFLOAT- (COUNTS-1));
              
            //處理負數
            //BUCKETSFLOAT>0 && BUCKETSFLOAT<1，只有1未滿桶
            //BUCKETSFLOAT>1正常

            if (BUCKETSFLOAT>0 && BUCKETSFLOAT<1)
            {
                BUCKETSSMAILL =Convert.ToDecimal(BUCKETSFLOAT);
                COUNTS = 0;
            }
            else if(BUCKETSFLOAT>1)
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

                sbSql.AppendFormat(@" DELETE [TKMOC].[dbo].[REPORTMOCBOM]");
                sbSql.AppendFormat(@"  ");

                if (COUNTS==0)
                {
                    sbSql.AppendFormat(@"       
                                            INSERT INTO [TKMOC].[dbo].[REPORTMOCBOM]
                                            ([TA001],[TA002],[TA006],[TA034],[BOXS],[MD003],[MB002],[MD006])                                            
                                            SELECT TA001,TA002,TA006,TA034,{2},MD003,MB002,CONVERT(DECIMAL(16,3),MD006*{3})
                                            FROM [TK].dbo.MOCTA,[TK].dbo.BOMMD,[TK].dbo.INVMB
                                            WHERE TA006=MD001
                                            AND MD003=MB001
                                            AND (MD003 LIKE '1%' OR MD003 LIKE '301%')
                                            AND MB002 NOT LIKE '%水麵%'
                                            AND TA001='{0}' AND TA002='{1}'
                                            ORDER BY MD003

                                           ", TA001, TA002, 1, BUCKETSSMAILL);
                }
                else if(COUNTS>=1)
                {
                    sbSql.AppendFormat(@"       
                                            INSERT INTO [TKMOC].[dbo].[REPORTMOCBOM]
                                            ([TA001],[TA002],[TA006],[TA034],[BOXS],[MD003],[MB002],[MD006])
                                            SELECT TA001,TA002,TA006,TA034,{2},MD003,MB002,MD006
                                            FROM [TK].dbo.MOCTA,[TK].dbo.BOMMD,[TK].dbo.INVMB
                                            WHERE TA006=MD001
                                            AND MD003=MB001
                                            AND (MD003 LIKE '1%' OR MD003 LIKE '301%')
                                            AND MB002 NOT LIKE '%水麵%'
                                            AND TA001='{0}' AND TA002='{1}'
                                            ORDER BY MD003

                                           ", TA001, TA002, 1);
                    sbSql.AppendFormat(@"       
                                            INSERT INTO [TKMOC].[dbo].[REPORTMOCBOM]
                                            ([TA001],[TA002],[TA006],[TA034],[BOXS],[MD003],[MB002],[MD006])                                            
                                            SELECT TA001,TA002,TA006,TA034,{2},MD003,MB002,CONVERT(DECIMAL(16,3),MD006*{3})
                                            FROM [TK].dbo.MOCTA,[TK].dbo.BOMMD,[TK].dbo.INVMB
                                            WHERE TA006=MD001
                                            AND MD003=MB001
                                            AND (MD003 LIKE '1%' OR MD003 LIKE '301%')
                                            AND MB002 NOT LIKE '%水麵%'
                                            AND TA001='{0}' AND TA002='{1}'
                                            ORDER BY MD003

                                           ", TA001, TA002, 2, BUCKETSSMAILL);

                    for (int i = 3; i <= COUNTS; i++)
                    {
                        sbSql.AppendFormat(@"
                                            INSERT INTO [TKMOC].[dbo].[REPORTMOCBOM]
                                            ([TA001],[TA002],[TA006],[TA034],[BOXS],[MD003],[MB002],[MD006])
                                            SELECT TA001,TA002,TA006,TA034,{2},MD003,MB002,MD006
                                            FROM [TK].dbo.MOCTA,[TK].dbo.BOMMD,[TK].dbo.INVMB
                                            WHERE TA006=MD001
                                            AND MD003=MB001
                                            AND (MD003 LIKE '1%' OR MD003 LIKE '301%')
                                            AND MB002 NOT LIKE '%水麵%'
                                            AND TA001='{0}' AND TA002='{1}'
                                            ORDER BY MD003

                                           ", TA001, TA002, i);
                    }
                }


                ////新增5個空白格子
                //sbSql.AppendFormat(@"
                //                            INSERT INTO [TKMOC].[dbo].[REPORTMOCBOM]
                //                            ([TA001],[TA002],[TA006],[TA034],[BOXS],[MD003],[MB002],[MD006])
                //                            SELECT TA001,TA002,TA006,TA034,{2},MD003,MB002,0
                //                            FROM [TK].dbo.MOCTA,[TK].dbo.BOMMD,[TK].dbo.INVMB
                //                            WHERE TA006=MD001
                //                            AND MD003=MB001
                //                            AND (MD003 LIKE '1%' OR MD003 LIKE '301%')
                //                            AND MD003 NOT LIKE '301400%'
                //                            AND MB002 NOT LIKE '%水麵%'
                //                            AND TA001='{0}' AND TA002='{1}'
                //                            ORDER BY MD003

                //                           ", TA001, TA002, COUNTS + 1);

                //sbSql.AppendFormat(@"
                //                            INSERT INTO [TKMOC].[dbo].[REPORTMOCBOM]
                //                            ([TA001],[TA002],[TA006],[TA034],[BOXS],[MD003],[MB002],[MD006])
                //                            SELECT TA001,TA002,TA006,TA034,{2},MD003,MB002,0
                //                            FROM [TK].dbo.MOCTA,[TK].dbo.BOMMD,[TK].dbo.INVMB
                //                            WHERE TA006=MD001
                //                            AND MD003=MB001
                //                            AND (MD003 LIKE '1%' OR MD003 LIKE '301%')
                //                            AND MD003 NOT LIKE '301400%'
                //                            AND MB002 NOT LIKE '%水麵%'
                //                            AND TA001='{0}' AND TA002='{1}'
                //                            ORDER BY MD003

                //                           ", TA001, TA002, COUNTS + 2);

                //sbSql.AppendFormat(@"
                //                            INSERT INTO [TKMOC].[dbo].[REPORTMOCBOM]
                //                            ([TA001],[TA002],[TA006],[TA034],[BOXS],[MD003],[MB002],[MD006])
                //                            SELECT TA001,TA002,TA006,TA034,{2},MD003,MB002,0
                //                            FROM [TK].dbo.MOCTA,[TK].dbo.BOMMD,[TK].dbo.INVMB
                //                            WHERE TA006=MD001
                //                            AND MD003=MB001
                //                            AND (MD003 LIKE '1%' OR MD003 LIKE '301%')
                //                            AND MD003 NOT LIKE '301400%'
                //                            AND MB002 NOT LIKE '%水麵%'
                //                            AND TA001='{0}' AND TA002='{1}'
                //                            ORDER BY MD003

                //                           ", TA001, TA002, COUNTS + 3);

                //sbSql.AppendFormat(@"
                //                            INSERT INTO [TKMOC].[dbo].[REPORTMOCBOM]
                //                            ([TA001],[TA002],[TA006],[TA034],[BOXS],[MD003],[MB002],[MD006])
                //                            SELECT TA001,TA002,TA006,TA034,{2},MD003,MB002,0
                //                            FROM [TK].dbo.MOCTA,[TK].dbo.BOMMD,[TK].dbo.INVMB
                //                            WHERE TA006=MD001
                //                            AND MD003=MB001
                //                            AND (MD003 LIKE '1%' OR MD003 LIKE '301%')
                //                            AND MD003 NOT LIKE '301400%'
                //                            AND MB002 NOT LIKE '%水麵%'
                //                            AND TA001='{0}' AND TA002='{1}'
                //                            ORDER BY MD003

                //                           ", TA001, TA002, COUNTS + 4);

                //sbSql.AppendFormat(@"
                //                            INSERT INTO [TKMOC].[dbo].[REPORTMOCBOM]
                //                            ([TA001],[TA002],[TA006],[TA034],[BOXS],[MD003],[MB002],[MD006])
                //                            SELECT TA001,TA002,TA006,TA034,{2},MD003,MB002,0
                //                            FROM [TK].dbo.MOCTA,[TK].dbo.BOMMD,[TK].dbo.INVMB
                //                            WHERE TA006=MD001
                //                            AND MD003=MB001
                //                            AND (MD003 LIKE '1%' OR MD003 LIKE '301%')
                //                            AND MD003 NOT LIKE '301400%'
                //                            AND MB002 NOT LIKE '%水麵%'
                //                            AND TA001='{0}' AND TA002='{1}'
                //                            ORDER BY MD003

                //                           ", TA001, TA002, COUNTS + 5);

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


        public void SETREPORTOIL(string TA003, string TA006, string BUCKETS, string TA021)
        {
            float BUCKETSORI = float.Parse(BUCKETS);
            bool CHECKFLOOR = IsIntegerFloor(BUCKETSORI);


            if (!string.IsNullOrEmpty(BUCKETS) && BUCKETSORI > 0)
            {
                if (CHECKFLOOR == true)
                {
                    ADDTOREPORTMOCBOMOIL(TA003, TA006, BUCKETS, TA021);
                    //MessageBox.Show(CHECKFLOOR  + BUCKETSORI.ToString());
                }
                else
                {
                    ADDTOREPORTMOCBOMODDOIL(TA003, TA006, BUCKETS, TA021);
                    //MessageBox.Show(CHECKFLOOR  + BUCKETSORI.ToString());
                }


            }


            StringBuilder SQL = new StringBuilder();

            SQL = SETSQLOIL();

            report1 = new Report();
            report1.Load(@"REPORT\油酥原料添加表-合併V3.frx");

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

            report1.Preview = previewControl2;
            report1.Show();
        }

        public StringBuilder SETSQLOIL()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@"
                             SELECT [ID]
                            ,[TA003] AS '日期'
                            ,[TA021] AS '線別'
                            ,'第'+CONVERT(nvarchar,[BOXS])+'桶' AS '桶數'
                            ,TA006 AS '成品'
                            ,TA034 AS '成品名'
                            ,[MD003] AS '品號'
                            ,[MB002] AS '品名'
                            ,[MD006] AS '重量'
                            ,'' AS '複核'
                            ,'' AS '油酥'
                            ,'' AS '檢查麵粉袋的麵粉線頭'
                            ,'' AS 'A製造  B有效'
                            ,'' AS '外觀:攪拌均勻度、軟硬度'
                            ,'' AS '攪拌時間  始'
                            ,'' AS '攪拌時間  終'
                            ,'' AS '投 料 人'
                            ,'' AS '對 點 人'
                            ,'' AS '單位幹部'
                            ,'' AS '品質判定'
                            ,'' AS '換線清潔檢查'
                            ,'配比 '+ BOMMC.UDF03 AS '配方比'
                            FROM [TKMOC].[dbo].[REPORTMOCBOMOIL]
                            LEFT JOIN [TK].dbo.BOMMC ON MC001=TA006

                            WHERE [MD003] NOT  IN ('101001009','3010000111')   
                            ORDER BY [TA021],[BOXS],[MD003]
     

                            ");



            return SB;
        }
        /// <summary>
        /// 剛好滿桶數，沒有未滿桶
        /// </summary>
        /// <param name="TA001"></param>
        /// <param name="TA002"></param>
        /// <param name="BUCKETS"></param>
        public void ADDTOREPORTMOCBOMOIL(string TA003, string TA006, string BUCKETS, string TA021)
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

                sbSql.AppendFormat(@" DELETE [TKMOC].[dbo].[REPORTMOCBOMOIL]");
                sbSql.AppendFormat(@" ");

                for (int i = 1; i <= COUNTS; i++)
                {
                    sbSql.AppendFormat(@"
                                            INSERT INTO [TKMOC].[dbo].[REPORTMOCBOMOIL]
                                            ([TA003],[TA021],[TA006],[TA034],[BOXS],[MD003],[MB002],[MD006])
                                            SELECT TA003,TA021,TA006,TA034,'{3}',MD003,MB002,MD006
                                            FROM [TK].dbo.MOCTA,[TK].dbo.BOMMD,[TK].dbo.INVMB
                                            WHERE TA006=MD001
                                            AND MD003=MB001
                                            AND (MD003 LIKE '1%' OR MD003 LIKE '301%')
                                            AND MD003 NOT LIKE '301400%'
                                            AND TA003='{0}' AND TA006='{1}' AND TA021='{2}' 
                                            ORDER BY MD003

                                           ", TA003, TA006,  TA021, i);
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
        public void ADDTOREPORTMOCBOMODDOIL(string TA003, string TA006, string BUCKETS, string TA021)
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

                sbSql.AppendFormat(@" DELETE [TKMOC].[dbo].[REPORTMOCBOMOIL]");
                sbSql.AppendFormat(@" ");

                if (COUNTS == 0)
                {
                    sbSql.AppendFormat(@"       
                                            INSERT INTO [TKMOC].[dbo].[REPORTMOCBOMOIL]
                                            ([TA003],[TA021],[TA006],[TA034],[BOXS],[MD003],[MB002],[MD006])
                                            SELECT TA003,TA021,TA006,TA034,'{3}',MD003,MB002,MD006*{4}
                                            FROM [TK].dbo.MOCTA,[TK].dbo.BOMMD,[TK].dbo.INVMB
                                            WHERE TA006=MD001
                                            AND MD003=MB001
                                            AND (MD003 LIKE '1%' OR MD003 LIKE '301%')
                                            AND MD003 NOT LIKE '301400%'
                                            AND TA003='{0}' AND TA006='{1}' AND TA021='{2}' 
                                            ORDER BY MD003

                                           ", TA003, TA006, TA021, 1, BUCKETSSMAILL);
                }
                else if (COUNTS >= 1)
                {
                    sbSql.AppendFormat(@"       
                                            INSERT INTO [TKMOC].[dbo].[REPORTMOCBOMOIL]
                                            ([TA003],[TA021],[TA006],[TA034],[BOXS],[MD003],[MB002],[MD006])
                                            SELECT TA003,TA021,TA006,TA034,'{3}',MD003,MB002,MD006
                                            FROM [TK].dbo.MOCTA,[TK].dbo.BOMMD,[TK].dbo.INVMB
                                            WHERE TA006=MD001
                                            AND MD003=MB001
                                            AND (MD003 LIKE '1%' OR MD003 LIKE '301%')
                                            AND MD003 NOT LIKE '301400%'
                                            AND TA003='{0}' AND TA006='{1}' AND TA021='{2}' 
                                            ORDER BY MD003

                                           ", TA003, TA006, TA021, 1);

                    sbSql.AppendFormat(@"       
                                            INSERT INTO [TKMOC].[dbo].[REPORTMOCBOMOIL]
                                            ([TA003],[TA021],[TA006],[TA034],[BOXS],[MD003],[MB002],[MD006])
                                            SELECT TA003,TA021,TA006,TA034,'{3}',MD003,MB002,MD006*{4}
                                            FROM [TK].dbo.MOCTA,[TK].dbo.BOMMD,[TK].dbo.INVMB
                                            WHERE TA006=MD001
                                            AND MD003=MB001
                                            AND (MD003 LIKE '1%' OR MD003 LIKE '301%')
                                            AND MD003 NOT LIKE '301400%'
                                            AND TA003='{0}' AND TA006='{1}' AND TA021='{2}' 
                                            ORDER BY MD003

                                           ", TA003, TA006, TA021, 2, BUCKETSSMAILL);

                    for (int i = 3; i <= COUNTS; i++)
                    {
                        sbSql.AppendFormat(@"
                                             INSERT INTO [TKMOC].[dbo].[REPORTMOCBOMOIL]
                                            ([TA003],[TA021],[TA006],[TA034],[BOXS],[MD003],[MB002],[MD006])
                                            SELECT TA003,TA021,TA006,TA034,'{3}',MD003,MB002,MD006
                                            FROM [TK].dbo.MOCTA,[TK].dbo.BOMMD,[TK].dbo.INVMB
                                            WHERE TA006=MD001
                                            AND MD003=MB001
                                            AND (MD003 LIKE '1%' OR MD003 LIKE '301%')
                                            AND MD003 NOT LIKE '301400%'
                                            AND TA003='{0}' AND TA006='{1}' AND TA021='{2}' 
                                            ORDER BY MD003

                                           ", TA003, TA006, TA021, i);
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


        public static bool IsIntegerFloor(float f)
        {
            return f == Math.Floor(f);
        }

        public DataTable CHECK_BOMMD(int COUNTS, List<string> MD001)
        {
            //MessageBox.Show(COUNTS.ToString());
            int SETCOUNT = 1;
            DataSet ds = new DataSet();
            StringBuilder SQL = new StringBuilder();

            if(COUNTS>=1)
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

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }
        private void button2_Click(object sender, EventArgs e)
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

            //
            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
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

            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
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
                    }
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(ex.Message);
                }
            }

            if (CHECKED.Equals("Y"))
            {
                if(DT==null)
                {
                    SETREPORT2(TA001, TA002, BUCKETS, LINK_TA001TA002, LINK_TA006, LINK_TA034);
                }
                else
                {
                    MESS = "原料需單身品號元件不一致 或 組成用量不一致，不能合併\n";
                    foreach (DataRow ROW in DT.Rows)
                    {
                        // 每一行都是一個 DataRow                       
                        MESS = MESS+"品號:"+ROW["MD003"].ToString();
                        MESS=  MESS+"用量:"+ROW["MD006"].ToString();
                        MESS = MESS+"底數:"+ROW["MD007"].ToString();

                        MESS = MESS + "\n";
                    }
                    MessageBox.Show(MESS.ToString());
                }
               
            }
            else if (CHECKED.Equals("N"))
            {
                SETREPORT(textBox1.Text.Trim(), textBox2.Text.Trim(), textBox3.Text.Trim());
            }

        }
        private void button3_Click(object sender, EventArgs e)
        {
            SearchOIL();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //油酥原料添加表-合併
            SETREPORTOIL(textBox4.Text.Trim(), textBox5.Text.Trim(), textBox6.Text.Trim(), textBox7.Text.Trim());
        }



        #endregion

      
    }
}
