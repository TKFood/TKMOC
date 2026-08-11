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
    public partial class frmMOCCOMMENTS : Form
    {
        public frmMOCCOMMENTS()
        {
            InitializeComponent();
        }


        #region FUNCTION
        public void SEARCH_DG1(string SDATE, string EDATE)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                SqlConnection sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  
                                    SELECT 
                                    MOCTA.[TA001] AS '製令單別'
                                    ,MOCTA.[TA002] AS '製令單號'
                                    ,MOCTA.[TA003] AS '製令日期'
                                    ,MOCTA.[TA006] AS '品號'
                                    ,MOCTA.[TA034] AS '品名'
                                    ,MOCTA.[TA015] AS '預計產量'
                                    ,MOCTA.[TA017] AS '已生產量'
                                    ,MOCTA.[TA007] AS '單位'
                                    ,[MOCREASON] AS '差異說明'

                                    FROM [TK].dbo.MOCTA
                                    LEFT JOIN [TKMOC].[dbo].[MOCCOMMENTS] ON MOCTA.TA001=MOCCOMMENTS.TA001 AND MOCTA.TA002=MOCCOMMENTS.TA002
                                    WHERE MOCTA.TA001='A513'
                                    AND MOCTA.[TA003]>='{0}' AND MOCTA.[TA003]<='{1}'

                                    ORDER BY MOCTA.[TA001] ,MOCTA.[TA002] 

                             ", SDATE, EDATE);
                SqlDataAdapter da = new SqlDataAdapter(@"" + sbSql, sqlConn);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;

                dataGridView1.Columns["差異說明"].Width = 400;


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
            SETTEXTBOX();

            if (dataGridView1.CurrentRow != null)
            {
                DataRow DR= ((DataRowView)dataGridView1.CurrentRow.DataBoundItem).Row;
                textBox1.Text = DR["製令單別"].ToString();
                textBox2.Text = DR["製令單號"].ToString();
                textBox3.Text = DR["差異說明"].ToString();
            }
        }

        public void ADD_UPDATE_MOCCOMMENTS(
            string TA001,
            string TA002,
            string MOCREASON
        )
        {
            // 將 MOCTA 資料表直接在 USING 區塊內關聯，一次完成 MERGE
            string sql = @"
                            MERGE INTO [TKMOC].[dbo].[MOCCOMMENTS] AS Target
                            USING (
                                SELECT 
                                    @TA001 AS TA001,
                                    @TA002 AS TA002,
                                    @MOCREASON AS MOCREASON,
                                    m.TA003, m.TA006, m.TA007, m.TA034, m.TA015, m.TA017
                                FROM (SELECT @TA001 AS TA001, @TA002 AS TA002) AS Param
                                LEFT JOIN [TK].[dbo].[MOCTA] m 
                                       ON m.TA001 = Param.TA001 AND m.TA002 = Param.TA002
                            ) AS Source
                            ON Target.[TA001] = Source.[TA001] AND Target.[TA002] = Source.[TA002]
        
                            -- 存在：更新 MOCREASON 以及 MOCTA 的最新資料
                            WHEN MATCHED THEN
                                UPDATE SET 
                                    Target.[MOCREASON] = Source.[MOCREASON],
                                    Target.[TA003]     = Source.[TA003],
                                    Target.[TA006]     = Source.[TA006],
                                    Target.[TA007]     = Source.[TA007],
                                    Target.[TA034]     = Source.[TA034],
                                    Target.[TA015]     = Source.[TA015],
                                    Target.[TA017]     = Source.[TA017]

                            -- 不存在：新增資料（包含 MOCTA 欄位）
                            WHEN NOT MATCHED THEN
                                INSERT ([TA001], [TA002], [MOCREASON], [TA003], [TA006], [TA007], [TA034], [TA015], [TA017])
                                VALUES (Source.[TA001], Source.[TA002], Source.[MOCREASON], Source.[TA003], Source.[TA006], Source.[TA007], Source.[TA034], Source.[TA015], Source.[TA017]);
                        ";

            try
            {
                Class1 TKID = new Class1();
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                using (SqlConnection sqlConn = new SqlConnection(sqlsb.ConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand(sql, sqlConn))
                    {
                        cmd.Parameters.AddWithValue("@TA001", (object)TA001 ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@TA002", (object)TA002 ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@MOCREASON", (object)MOCREASON ?? DBNull.Value);

                        sqlConn.Open();
                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("ADD_UPDATE_MOCCOMMENTS 執行失敗: " + ex.Message, ex);
            }
        }

        /// <summary>
        /// 搜尋並定位回指定資料列
        /// </summary>
        private void RestoreSelectedRow(string ta001, string ta002)
        {
            if (string.IsNullOrEmpty(ta001) || string.IsNullOrEmpty(ta002)) return;

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                // 找到與剛才 Key 相同的資料列
                if (row.Cells["製令單別"].Value?.ToString() == ta001 &&
                    row.Cells["製令單號"].Value?.ToString() == ta002)
                {
                    // 取消目前選取狀態
                    dataGridView1.ClearSelection();

                    // 設定游標焦點到該行的第一個儲存格 (自動選取該行)
                    dataGridView1.CurrentCell = row.Cells[0];
                    row.Selected = true;

                    // 讓捲軸自動滾動到該行位置（讓該行顯示在畫面中）
                    dataGridView1.FirstDisplayedScrollingRowIndex = row.Index;
                    break;
                }
            }
        }

        public void SETFASTREPORT(string SDATES, string EDATES)
        {
            

            StringBuilder SQL1 = new StringBuilder();
            StringBuilder SQL2 = new StringBuilder();

            SQL1 = SETSQL1(SDATES, EDATES);

            Report report1 = new Report();
            report1.Load(@"REPORT\入庫差異表.frx");

            //20210902密      
            Class1 TKID = new Class1();
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;


            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL1(string SDATES, string EDATES)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
                            SELECT 
                            MOCTA.[TA001] AS '製令單別'
                            ,MOCTA.[TA002] AS '製令單號'
                            ,MOCTA.[TA003] AS '製令日期'
                            ,MOCTA.[TA006] AS '品號'
                            ,MOCTA.[TA034] AS '品名'
                            ,MOCTA.[TA015] AS '預計產量'
                            ,MOCTA.[TA017] AS '已生產量'
                            ,MOCTA.[TA007] AS '單位'
                            ,[MOCREASON] AS '差異說明'

                            FROM [TK].dbo.MOCTA
                            LEFT JOIN [TKMOC].[dbo].[MOCCOMMENTS] ON MOCTA.TA001=MOCCOMMENTS.TA001 AND MOCTA.TA002=MOCCOMMENTS.TA002
                            WHERE MOCTA.TA001='A513'
                            AND MOCTA.[TA003]>='{0}' AND MOCTA.[TA003]<='{1}'

                            ORDER BY MOCTA.[TA001] ,MOCTA.[TA002] 

                            ", SDATES, EDATES);

            return SB;

        }


        public void SETTEXTBOX()
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            string SDATE = dateTimePicker1.Value.ToString("yyyyMMdd");
            string EDATE = dateTimePicker2.Value.ToString("yyyyMMdd");
            SEARCH_DG1(SDATE, EDATE);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string TA001 = textBox1.Text;
            string TA002 = textBox2.Text;
            string MOCREASON = textBox3.Text;
            ADD_UPDATE_MOCCOMMENTS(TA001, TA002, MOCREASON);

            string SDATE = dateTimePicker1.Value.ToString("yyyyMMdd");
            string EDATE = dateTimePicker2.Value.ToString("yyyyMMdd");
            SEARCH_DG1(SDATE, EDATE);

            // 4. 定位回到剛才更新的那一筆  
            RestoreSelectedRow(TA001, TA002);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string SDATE = dateTimePicker3.Value.ToString("yyyyMMdd");
            string EDATE = dateTimePicker4.Value.ToString("yyyyMMdd");
            SETFASTREPORT(SDATE, EDATE);
        }
        #endregion


    }
}
