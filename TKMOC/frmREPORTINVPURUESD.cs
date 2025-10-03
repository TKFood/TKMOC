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
    public partial class frmREPORTINVPURUESD : Form
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

        SqlCommand cmd = new SqlCommand();
        SqlTransaction tran;
        int result;

        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();

        string tablename = null;

        string MD003;
        int rowIndexDG1 = -1;
        int rowIndexDG2 = -1;


        public frmREPORTINVPURUESD()
        {
            InitializeComponent();
        }
        private void frmREPORTINVPURUESD_Load(object sender, EventArgs e)
        {
            SHOW_TBALERTMESSAGES();
        }

        #region FUNCTION
        public void SEARCHMOCMANULINE(string SDay,string EDay)
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

                sbSql.AppendFormat(@"  
                                    WITH INVLA_SUMMARY AS (
                                    -- 在這裡一次性計算所有品號的 20019 倉庫庫存總和
                                    SELECT 
                                        LA001, 
                                        SUM(LA005 * LA011) AS Sum_20019
                                    FROM [TK].dbo.INVLA WITH (NOLOCK)
                                    WHERE 
                                        LA009 = '20019' 
                                        AND LA016 <> '********************'
                                    GROUP BY 
                                        LA001
                                )
                                -- 接下來的主查詢使用 JOIN 來連接這個彙總結果
                                SELECT
                                    T3.MD003 AS '品號',
                                    T3.MD035 AS '品名',
                                    ISNULL(T_SUM.Sum_20019, 0) AS '20019外倉',
                                    (
                                        -- 批號查詢仍須使用相關子查詢或 APPLY (因為 FOR XML PATH 難以用標準 JOIN 實現)
                                        SELECT CAST(LA016 AS NVARCHAR) + ',' 
                                        FROM [TK].dbo.INVLA AS INVLA_INNER WITH (NOLOCK)
                                        WHERE 
                                            INVLA_INNER.LA001 = T3.MD003  -- 依賴主查詢的 MD003
                                            AND INVLA_INNER.LA009 = '20006' 
                                            AND ISNULL(INVLA_INNER.LA016,'') <> '' 
                                            AND INVLA_INNER.LA016 <> '********************' 
                                        GROUP BY 
                                            INVLA_INNER.LA001, INVLA_INNER.LA016 
                                        HAVING 
                                            ISNULL(SUM(INVLA_INNER.LA005*INVLA_INNER.LA011), 0) > 0 
                                        FOR XML PATH('')
                                    ) AS '20006倉批號'
                                FROM 
                                    [TKMOC].dbo.[MOCMANULINE] AS T1 WITH(NOLOCK)
                                INNER JOIN 
                                    [TK].dbo.BOMMC AS T2 WITH(NOLOCK) ON T1.MB001 = T2.MC001
                                INNER JOIN 
                                    [TK].dbo.BOMMD AS T3 WITH(NOLOCK) ON T2.MC001 = T3.MD001
                                LEFT JOIN 
                                    [TK].dbo.INVMB AS T4 WITH(NOLOCK) ON T4.MB001 = T3.MD003
                                -- 將預先計算好的庫存總和連接回來
                                LEFT JOIN 
                                    INVLA_SUMMARY AS T_SUM ON T_SUM.LA001 = T3.MD003
                                WHERE 
                                    T1.[MANUDATE] >= '{0}' 
                                    AND T1.[MANUDATE] < ='{1}'
                                GROUP BY 
                                    T3.MD003, T3.MD035, T_SUM.Sum_20019  -- 彙總值也需加入 GROUP BY
                                ORDER BY 
                                    T3.MD003, T3.MD035;
                                    ", SDay, EDay);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);
                // 設定 SqlDataAdapter 的 CommandTimeout
                adapter1.SelectCommand.CommandTimeout = 120;  // 這裡設定 Timeout 為 2 分鐘，您可以根據需要調整


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

                        //根据列表中数据不同，显示不同颜色背景
                        foreach (DataGridViewRow dgRow in dataGridView1.Rows)
                        {
                            ////判断
                            //if (Convert.ToDecimal(dgRow.Cells[5].Value) > 0)
                            //{
                            //    //将这行的背景色设置成Pink
                            //    dgRow.DefaultCellStyle.BackColor = Color.Pink;
                        
                            //}
                        }

                        //dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 10);
                        //dataGridView1.DefaultCellStyle.Font = new Font("Tahoma", 11);
                        dataGridView1.Columns["品號"].Width = 100;
                        dataGridView1.Columns["品名"].Width = 220;
                        dataGridView1.Columns["20019外倉"].Width = 60;

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
                    MD003 = row.Cells["品號"].Value.ToString().Trim();

                    SEARCHINVPURMOC(MD003,dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));


                }
                else
                {
                    MD003 = null;
                }
            }
            
        }

        public void SEARCHINVPURMOC(string MD003,string SDay,string EDay)
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


                //查每日原物料的使用量、進貨量、庫存量
                //另外查詢領料數量

                sbSql.AppendFormat(@"   
                                    WITH
                                    -- 1. 統一定義基礎數據 (BASE_DATA)
                                    -- 結合所有四個 UNION 區塊的資料，並計算出 TNUM
                                    BASE_DATA AS (
                                        -- 區塊 1 & 2: MOCMANULINE 數據 (MANU = '包裝線' / MANU NOT IN ('包裝線'))
                                        SELECT
                                            [MANU],
                                            CONVERT(NVARCHAR, [MANUDATE], 112) AS MANUDATE,
                                            [MD003],
                                            [MD035],
                                            -- 複雜的 TNUM 計算邏輯 (保持不變)
		                                    --MOCTB1.TB001,MOCTB1.TB002,MOCTB1.TB003,MOCTB1.TB004,MOCTB1.TB005,
                                            --[NUM] ,([NUM] / [MC004] * [MD006] / [MD007] * (1 + [MD008])) * -1,
		                                    CONVERT(DECIMAL(16, 3), 
                                                CASE
                                                    -- 處理 MOCTB1 (MOCMANULINE.MOCTA001/002)
                                                    WHEN ISNULL(MOCTB1.TB003, '') <> '' AND ISNULL(MOCTB1.TB004 - MOCTB1.TB005, 0) <= 0 THEN 0
                                                    WHEN ISNULL(MOCTB1.TB003, '') <> '' AND ISNULL(MOCTB1.TB004 - MOCTB1.TB005, 0) > 0 THEN ISNULL((MOCTB1.TB004 - MOCTB1.TB005) * -1, 0)
                                                    -- 處理 MOCTB2 (MOCMANULINEMERGE.NO -> MOCTA.TA033)
                                                    WHEN ISNULL(MOCTB2.TB003, '') <> '' AND ISNULL(MOCTB2.TB004 - MOCTB2.TB005, 0) <= 0 THEN 0
                                                    WHEN ISNULL(MOCTB2.TB003, '') <> '' AND ISNULL(MOCTB2.TB004 - MOCTB2.TB005, 0) > 0 THEN 
                                                        (ISNULL(MOCTB2.TB004 - MOCTB2.TB005, 0) / ISNULL((SELECT COUNT(NO) FROM [TKMOC].[dbo].[MOCMANULINEMERGE] WHERE [MOCMANULINEMERGE].NO = MOCTA.TA033), 1)) * -1
                                                    -- 預設計算邏輯
                                                    WHEN [MANU] = '包裝線' THEN ([PACKAGE] / [MC004] * [MD006] / [MD007] * (1 + [MD008])) * -1
                                                    ELSE ([NUM] / [MC004] * [MD006] / [MD007] * (1 + [MD008])) * -1
                                                END
                                            ) AS TNUM,
                                            [MB004],
                                            [MOCMANULINE].[MB001],
                                            [MOCMANULINE].[MB002],
                                            CASE WHEN [MANU] = '包裝線' THEN [PACKAGE] ELSE [NUM] END AS PACKAGE, -- 根據線別區分 PACKAGE 或 NUM
                                            [MOCMANULINE].BAR AS BAR,
                                            [COPTD001],
                                            [COPTD002],
                                            [COPTD003]
                                        FROM [TKMOC].dbo.[MOCMANULINE] WITH(NOLOCK)
                                        LEFT JOIN [TKMOC].dbo.[MOCMANULINERESULT] WITH(NOLOCK) ON [MOCMANULINERESULT].SID = [MOCMANULINE].ID
                                        JOIN [TK].dbo.BOMMC WITH(NOLOCK) ON [MOCMANULINE].[MB001] = BOMMC.MC001
                                        JOIN [TK].dbo.BOMMD WITH(NOLOCK) ON BOMMC.MC001 = BOMMD.MD001 AND [MOCMANULINE].MB001 = BOMMD.MD001
                                        LEFT JOIN [TK].dbo.MOCTB MOCTB1 WITH(NOLOCK) ON [MOCTA001] = MOCTB1.TB001 AND [MOCTA002] = MOCTB1.TB002 AND MOCTB1.TB003 = MD003
                                        LEFT JOIN [TK].dbo.INVMB WITH(NOLOCK) ON INVMB.MB001 = MD003
                                        LEFT JOIN [TKMOC].dbo.[MOCMANULINEMERGE] WITH(NOLOCK) ON [MOCMANULINEMERGE].SID = [MOCMANULINE].ID
                                        LEFT JOIN [TK].dbo.MOCTA WITH(NOLOCK) ON TA033 = [MOCMANULINEMERGE].NO
                                        LEFT JOIN [TK].dbo.MOCTB MOCTB2 WITH(NOLOCK) ON MOCTA.TA001 = MOCTB2.TB001 AND MOCTA.TA002 = MOCTB2.TB002 AND MOCTB2.TB003 = MD003
                                        WHERE
                                            ([MOCMANULINE].MB001 = BOMMC.MC001)
                                            AND (
                                                (ISNULL([MOCMANULINEMERGE].NO, '') <> '' AND ISNULL(MOCTA.TA001, '') <> '')
                                                OR (ISNULL([MOCMANULINEMERGE].NO, '') = '')
                                            )
                                            AND CONVERT(NVARCHAR, [MANUDATE], 112) BETWEEN '{0}' AND '{1}'
                                            AND [MD003] = '{2}'

                                        UNION ALL

                                        -- 區塊 3: PURTD 進貨數據
                                        SELECT
                                            '1進貨' AS MANU,
                                            TD012 AS MANUDATE,
                                            TD004 AS MD003,
                                            MB002 AS MD035,
                                            CONVERT(DECIMAL(14, 3), 
                                                CASE 
                                                    WHEN ISNULL(MD002, '') <> '' THEN (ISNULL(TD008 - TD015, 0) * MD004 / MD003)
                                                    ELSE (TD008 - TD015) 
                                                END
                                            ) AS TNUM,
                                            MB004,
                                            NULL AS MB001,
                                            NULL AS MB002,
                                            NULL AS PACKAGE,
                                            0 AS BAR,
                                            TD001 AS COPTD001,
                                            TD002 AS COPTD002,
                                            TD003 AS COPTD003
                                        FROM [TK].dbo.PURTD WITH(NOLOCK)
                                        JOIN [TK].dbo.PURTC WITH(NOLOCK) ON TC001 = TD001 AND TC002 = TD002 AND TC014 = 'Y'
                                        JOIN [TK].dbo.INVMB WITH(NOLOCK) ON TD004 = MB001
                                        LEFT JOIN [TK].dbo.INVMD WITH(NOLOCK) ON MD001 = TD004 AND MD002 = TD009
                                        WHERE
                                            TD018 = 'Y' AND TD016 = 'N'
                                            AND TD012 BETWEEN '{0}' AND '{1}'
                                            AND TD004 = '{2}'
                                            AND TD007 IN (SELECT [TD007] FROM [TKMOC].dbo.[REPORTINVPURUESD_TD007])

                                        UNION ALL

                                        -- 區塊 4: INVLA 庫存數據 (僅選取第一個日期)
                                        SELECT
                                            '0庫存' AS MANU,
                                            CONVERT(NVARCHAR, GETDATE(), 112) AS MANUDATE, -- 庫存日期使用當前日期
                                            LA001 AS MD003,
                                            MB002 AS MD035,
                                            SUM(LA005 * LA011) AS TNUM,
                                            MB004,
                                            NULL AS MB001,
                                            NULL AS MB002,
                                            NULL AS PACKAGE,
                                            0 AS BAR,
                                            NULL AS COPTD001,
                                            NULL AS COPTD002,
                                            NULL AS COPTD003
                                        FROM [TK].dbo.INVLA WITH(NOLOCK)
                                        JOIN [TK].dbo.INVMB WITH(NOLOCK) ON LA001 = MB001
                                        WHERE
                                            LA009 IN (SELECT [TD007] FROM [TKMOC].dbo.[REPORTINVPURUESD_TD007])
                                            AND LA001 = '{2}'
                                        GROUP BY LA001, MB002, MB004

                                        UNION ALL

                                        -- 區塊 5: INVPURUESD 手動進出貨數據
                                        SELECT
                                            '1手動進出貨' AS MANU,
                                            CONVERT(NVARCHAR, INVPURUESD.DATES, 112) AS MANUDATE,
                                            INVPURUESD.MB001 AS MD003,
                                            INVMB.MB002 AS MD035,
                                            INVPURUESD.NUM AS TNUM,
                                            INVMB.MB004,
                                            NULL AS MB001,
                                            NULL AS MB002,
                                            NULL AS PACKAGE,
                                            0 AS BAR,
                                            NULL AS COPTD001,
                                            NULL AS COPTD002,
                                            NULL AS COPTD003
                                        FROM [TKMOC].dbo.INVPURUESD WITH(NOLOCK)
                                        JOIN [TK].dbo.INVMB WITH(NOLOCK) ON INVMB.MB001 = INVPURUESD.MB001
                                        WHERE
                                            INVPURUESD.DATES BETWEEN '{0}' AND '{1}'
                                            AND INVPURUESD.MB001 = '{2}'
                                    ),
                                    -- 2. 應用 Row_Number 以建立排序 ID (與原邏輯一致)
                                    ORDERED_DATA AS (
                                        SELECT
                                            ROW_NUMBER() OVER (ORDER BY MANUDATE, MANU) AS ID, -- 排序邏輯與原本 TEMP2/TEMP4 的 ROW_NUMBER 相同
                                            MANU, MANUDATE, MD003, MD035, TNUM, MB004, MB001, MB002, PACKAGE, BAR, COPTD001, COPTD002, COPTD003
                                        FROM BASE_DATA
                                    )
                                    -- 3. 最終結果集：計算移動總和 (Running Total)
                                    SELECT
                                        -- 替換：使用 CROSS APPLY 計算累積總和
                                        Total.CumulativeSum AS '預計庫存量',
                                        A.ID AS '列數',
                                        A.MANU AS '線別',
                                        A.MANUDATE AS '日期',
                                        A.MD003 AS '品號',
                                        A.MD035 AS '品名',
                                        A.TNUM AS '用量',
                                        A.MB004 AS '單位',
                                        A.MB001 AS '成品',
                                        A.MB002 AS '成品名',
                                        A.PACKAGE AS '成品數',
                                        A.BAR AS '桶數',
                                        A.COPTD001 AS '訂單單別',
                                        A.COPTD002 AS '訂單單號',
                                        A.COPTD003 AS '訂單序號'
                                    FROM
                                        ORDERED_DATA AS A
                                    -- 使用 CROSS APPLY 模擬 SUM() OVER (ORDER BY ID)
                                    CROSS APPLY (
                                        SELECT
                                            SUM(B.TNUM) AS CumulativeSum
                                        FROM
                                            ORDERED_DATA AS B
                                        -- 關鍵：利用 ID 欄位確保累加到當前行 (A.ID)
                                        WHERE
                                            B.ID <= A.ID
                                    ) AS Total
                                    ORDER BY
                                        A.MANUDATE, A.MANU;

           
                                    ", SDay, EDay, MD003);


                adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
             
                sqlConn.Open();                

                ds2.Clear();
                adapter2.Fill(ds2, "ds2");
                //CommandTimeout
                adapter2.SelectCommand.CommandTimeout = 240;

                sqlConn.Close();


                if (ds2.Tables["ds2"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else
                {
                    if (ds2.Tables["ds2"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView2.DataSource = ds2.Tables["ds2"];
                        dataGridView2.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                        //根据列表中数据不同，显示不同颜色背景
                        foreach (DataGridViewRow dgRow in dataGridView2.Rows)
                        {
                            //判断
                            if (Convert.ToDecimal(dgRow.Cells[5].Value) > 0)
                            {
                                //将这行的背景色设置成Pink
                                dgRow.DefaultCellStyle.BackColor = Color.Pink;
                            }
                        }

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

        public void SEARCHDG1(string SEARCHSTRING, int INDEX)
        {
            String searchValue = SEARCHSTRING;
            rowIndexDG1 = INDEX;
            int ROWS = 0;

            for (int i = INDEX; i < dataGridView1.Rows.Count; i++)
            {
                ROWS = i;

                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Contains(searchValue))
                {
                    rowIndexDG1 = i;

                    dataGridView1.CurrentRow.Selected = false;
                    dataGridView1.Rows[i].Selected = true;
                    int index = rowIndexDG1;

                    DataGridViewRow row = dataGridView1.Rows[index];
                    MD003 = row.Cells["品號"].Value.ToString().Trim();

                    SEARCHINVPURMOC(MD003, dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));

                    dataGridView1.FirstDisplayedScrollingRowIndex = index;

                    break;
                }
                if (dataGridView1.Rows[i].Cells[1].Value.ToString().Contains(searchValue))
                {
                    rowIndexDG1 = i;

                    dataGridView1.CurrentRow.Selected = false;
                    dataGridView1.Rows[i].Selected = true;
                    int index = rowIndexDG1;

                    DataGridViewRow row = dataGridView1.Rows[index];
                    MD003 = row.Cells["品號"].Value.ToString().Trim();

                    SEARCHINVPURMOC(MD003, dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));

                    dataGridView1.FirstDisplayedScrollingRowIndex = index;

                    break;
                }
            }

            if (ROWS == dataGridView1.Rows.Count - 1)
            {

                if (MessageBox.Show("已查到最後一筆，是否從頭開始?", "已查到最後一筆，是否從頭開始?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    SEARCHDG1(textBox1.Text.Trim(), 0);
                }
                else
                {

                }
            }
        }

        public void ADDINVPURUESD(string KIND,string DATES,string MB001,string NUM)
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


                //sbSql.AppendFormat(" DELETE [TKMOC].[dbo].[INVPURUESD]");
                sbSql.AppendFormat(" INSERT INTO [TKMOC].[dbo].[INVPURUESD]");
                sbSql.AppendFormat(" ([KIND],[DATES],[MB001],[NUM])");
                sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}')",KIND,DATES,MB001,NUM);
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
        
        public void DELINVPURUESD(string MB001)
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


                sbSql.AppendFormat(" DELETE [TKMOC].[dbo].[INVPURUESD]");               
                sbSql.AppendFormat(" WHERE MB001='{0}'",MB001);
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

        public void SEARCHCOPTD(string TD002)
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
                                    SELECT TD004 AS '品號',TD005  AS '品名',TD006  AS '規格'
                                    FROM [TK].dbo.COPTD
                                    WHERE TD002='{0}'
                                    ORDER BY TD004
                                    ", TD002);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count == 0)
                {
                    dataGridView3.DataSource = null;
                }
                else
                {
                    if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView3.DataSource = ds1.Tables["ds1"];
                        dataGridView3.AutoResizeColumns();
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

        public void SEARCHBOM(string MD001)
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

                //找下層bom
                sbSql.AppendFormat(@"
                                   SELECT MD003  AS '品號',MD035 AS '品名',MD001 AS '成品號'
                                    FROM [TK].dbo.BOMMD
                                    WHERE  MD001='{0}'
                                    ORDER BY  MD001,MD003
                                    ", MD001);

                //找所有階層的bom
                //sbSql.AppendFormat(@"
                //                    ;WITH BOMOrder AS (

                //                    SELECT MD001, MD003 ,MD006,MD007,MD008 ,1 AS BOMLevel,CONVERT(DECIMAL(16,5),1*MD006/MD007) AS NUM,MC004
                //                    FROM [TK].dbo.VBOMMD 
                //                    UNION ALL	
                //                    SELECT A.MD001, B.MD003,B.MD006,B.MD007,B.MD008, (B.BOMLevel + 1) AS BOMLevel,CONVERT(DECIMAL(16,5),B.NUM*A.MD006/A.MD007) AS NUM,B.MC004
                //                    FROM [TK].dbo.VBOMMD A
                //                    INNER JOIN BOMOrder B ON A.MD003 = B.MD001
                //                    )
                //                    SELECT  MD003 AS '品號' ,MB002  AS '品名',MD001  AS '成品號'
                //                    FROM BOMOrder,[TK].dbo.INVMB
                //                    WHERE BOMLevel<=5
                //                    AND MD003=MB001
                //                    AND MD001='{0}'
                //                    GROUP BY MD001, MD003 ,MB002
                //                    ORDER BY MD003
                //                    ", MD001);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count == 0)
                {
                    dataGridView4.DataSource = null;
                }
                else
                {
                    if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView4.DataSource = ds1.Tables["ds1"];
                        dataGridView4.AutoResizeColumns();
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

        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            textBox4.Text = null;

            if (dataGridView3.CurrentRow != null)
            {
                int rowindex = dataGridView3.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView3.Rows[rowindex];
                    MD003 = row.Cells["品號"].Value.ToString().Trim();

                    textBox4.Text = MD003;
                }
                else
                {
                    textBox4.Text = null;
                }
            }
        }

        private void dataGridView4_SelectionChanged(object sender, EventArgs e)
        {

            textBox1.Text = null;

            if (dataGridView4.CurrentRow != null)
            {
                int rowindex = dataGridView4.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView4.Rows[rowindex];
                    MD003 = row.Cells["品號"].Value.ToString().Trim();

                    textBox1.Text = MD003;
                }
                else
                {
                    textBox1.Text = null;
                }
            }
        }
        private void dataGridView3_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox1.Text = null;
            SEARCHBOM(textBox4.Text.Trim());
        }

        private void dataGridView4_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            SEARCHDG1(textBox1.Text.Trim(), 0);
        }

        public void SHOW_TBALERTMESSAGES()
        {
            DataTable DT = FIND_TBALERTMESSAGES();
            string message = "";

            if (DT!=null && DT.Rows.Count>=1)
            {
                foreach(DataRow DR in DT.Rows)
                {
                     message = message+ DR["MESSAGES"].ToString() + Environment.NewLine;
                }
            }

            MessageBox.Show(message);
        }

        public DataTable FIND_TBALERTMESSAGES()
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

                //找下層bom
                sbSql.AppendFormat(@"
                                   SELECT 
                                    [ID]
                                    ,[MESSAGES]
                                    ,[ISCLOSES]
                                    FROM [TKMOC].[dbo].[TBALERTMESSAGES]
                                    WHERE [ISCLOSES] IN ('Y')
                                    ORDER BY [ID]
                                    ");

               
                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();


                if(ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["ds1"];
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

            }
        }


        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            SEARCHMOCMANULINE(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
        }
        private void button2_Click(object sender, EventArgs e)
        {
            if (rowIndexDG1 == -1)
            {
                SEARCHDG1(textBox1.Text.Trim(), 0);
            }
            else
            {
                SEARCHDG1(textBox1.Text.Trim(), rowIndexDG1 + 1);
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox2.Text)&& !string.IsNullOrEmpty(MD003))
            {
                ADDINVPURUESD("1手動進貨", dateTimePicker3.Value.ToString("yyyy/MM/dd"),MD003, textBox2.Text);
            }

            SEARCHINVPURMOC(MD003, dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
        }
        private void button4_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show(MD003+" 要刪除了?", MD003+ " 要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELINVPURUESD(MD003);
                SEARCHINVPURMOC(MD003, dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }
        private void button5_Click(object sender, EventArgs e)
        {
            textBox4.Text = null;
            SEARCHCOPTD(textBox3.Text.Trim());
        }

        private void button6_Click(object sender, EventArgs e)
        {
            textBox1.Text = null;
            SEARCHBOM(textBox4.Text.Trim());
        }





        #endregion

       
    }
}
