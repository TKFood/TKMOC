using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
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
        //手動進出貨的類別，寫入與查詢顯示統一使用同一個值
        const string KIND_MANUAL = "1手動進出貨";

        string MD003;
        int rowIndexDG1 = -1;
        bool isLoadingDataGrid1 = false;


        public frmREPORTINVPURUESD()
        {
            InitializeComponent();
        }
        private void frmREPORTINVPURUESD_Load(object sender, EventArgs e)
        {
            SHOW_TBALERTMESSAGES();
        }

        #region FUNCTION

        /// <summary>
        /// 建立資料庫連線 (帳密解密)
        /// </summary>
        private SqlConnection CreateConnection()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            return new SqlConnection(sqlsb.ConnectionString);
        }

        /// <summary>
        /// 參數化查詢，回傳 DataTable
        /// </summary>
        private DataTable QueryTable(string sql, int timeout, params SqlParameter[] parameters)
        {
            using (SqlConnection conn = CreateConnection())
            using (SqlCommand command = new SqlCommand(sql, conn))
            using (SqlDataAdapter adapter = new SqlDataAdapter(command))
            {
                command.CommandTimeout = timeout;
                if (parameters != null)
                {
                    command.Parameters.AddRange(parameters);
                }

                DataTable dt = new DataTable();
                adapter.Fill(dt);
                return dt;
            }
        }

        /// <summary>
        /// 參數化異動 (含交易)，回傳影響筆數
        /// </summary>
        private int ExecuteNonQueryTran(string sql, params SqlParameter[] parameters)
        {
            using (SqlConnection conn = CreateConnection())
            {
                conn.Open();
                using (SqlTransaction tran = conn.BeginTransaction())
                using (SqlCommand command = new SqlCommand(sql, conn, tran))
                {
                    command.CommandTimeout = 60;
                    if (parameters != null)
                    {
                        command.Parameters.AddRange(parameters);
                    }

                    try
                    {
                        int result = command.ExecuteNonQuery();
                        if (result == 0)
                        {
                            tran.Rollback();    //交易取消
                        }
                        else
                        {
                            tran.Commit();      //執行交易
                        }
                        return result;
                    }
                    catch
                    {
                        tran.Rollback();
                        throw;
                    }
                }
            }
        }

        private void ShowError(string where, Exception ex)
        {
            MessageBox.Show(where + " 發生錯誤：" + Environment.NewLine + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private static DateTime ToDate(string yyyyMMdd)
        {
            return DateTime.ParseExact(yyyyMMdd, "yyyyMMdd", CultureInfo.InvariantCulture);
        }

        public void SEARCHMOCMANULINE(string SDay, string EDay)
        {
            try
            {
                string sql = @"
                                SET NOCOUNT ON;

                                -- 1. 先找出區間內排程用到的原料 (只展一層 BOM)
                                SELECT T3.MD003, T3.MD035
                                INTO #MAT
                                FROM [TKMOC].dbo.[MOCMANULINE] AS T1 WITH(NOLOCK)
                                INNER JOIN [TK].dbo.BOMMC AS T2 WITH(NOLOCK) ON T1.MB001 = T2.MC001
                                INNER JOIN [TK].dbo.BOMMD AS T3 WITH(NOLOCK) ON T2.MC001 = T3.MD001
                                WHERE T1.[MANUDATE] >= @SDate
                                AND T1.[MANUDATE] < DATEADD(DAY, 1, @EDate)
                                GROUP BY T3.MD003, T3.MD035;

                                SELECT DISTINCT MD003 INTO #ITEM FROM #MAT;
                                CREATE CLUSTERED INDEX IX_ITEM ON #ITEM (MD003);

                                -- 本次要讀取的倉別：20028、20006 (左表顯示用) + 20004 (總庫存量用)
                                SELECT CAST('20028' AS NVARCHAR(10)) AS LA009 INTO #WH
                                UNION SELECT '20006' UNION SELECT '20004';
                                CREATE CLUSTERED INDEX IX_WH ON #WH (LA009);

                                -- 2. 只針對這些原料、這些倉別掃一次 INVLA (改用 JOIN 取代 OR + IN 子查詢，可使用索引)
                                SELECT LA.LA001, LA.LA009, LA.LA016, SUM(LA.LA005 * LA.LA011) AS QTY
                                INTO #LA
                                FROM [TK].dbo.INVLA AS LA WITH(NOLOCK)
                                INNER JOIN #ITEM AS I ON I.MD003 = LA.LA001
                                INNER JOIN #WH AS W ON W.LA009 = LA.LA009
                                GROUP BY LA.LA001, LA.LA009, LA.LA016;
                                CREATE CLUSTERED INDEX IX_LA ON #LA (LA001, LA009);

                                -- 3. 組出結果 (庫存與批號條件與原本相同)
                                SELECT
                                    M.MD003 AS '品號',
                                    M.MD035 AS '品名',
                                    (
                                        -- 總庫存量 (限制庫別 20006、20004)，數量 = SUM(LA005 * LA011)
                                        -- 改讀已彙總的 #LA，不再逐筆回查 INVLA
                                        SELECT ISNULL(SUM(L.QTY), 0)
                                        FROM #LA AS L
                                        WHERE L.LA001 = M.MD003
                                        AND L.LA009 IN ('20006', '20004')
                                    ) AS '總庫存量',
                                    ISNULL(S.Sum_20028, 0) AS '3F--倉庫',
                                    (
                                        SELECT CAST(L.LA016 AS NVARCHAR) + ','
                                        FROM #LA AS L
                                        WHERE L.LA001 = M.MD003
                                        AND L.LA009 = '20006'
                                        AND ISNULL(L.LA016, '') <> ''
                                        AND L.LA016 <> '********************'
                                        AND ISNULL(L.QTY, 0) > 0
                                        ORDER BY L.LA016
                                        FOR XML PATH('')
                                    ) AS '20006倉批號'
                                FROM #MAT AS M
                                LEFT JOIN (
                                    SELECT
                                        LA001,
                                        SUM(QTY) AS Sum_20028
                                    FROM #LA
                                    WHERE LA009 = '20028'
                                    AND LA016 <> '********************'
                                    GROUP BY LA001
                                ) AS S ON S.LA001 = M.MD003
                                ORDER BY M.MD003, M.MD035;

                                DROP TABLE #LA;
                                DROP TABLE #WH;
                                DROP TABLE #ITEM;
                                DROP TABLE #MAT;
                                ";

                DataTable dt = QueryTable(sql, 300,
                    new SqlParameter("@SDate", SqlDbType.DateTime) { Value = ToDate(SDay) },
                    new SqlParameter("@EDate", SqlDbType.DateTime) { Value = ToDate(EDay) },
                    new SqlParameter("@SDay", SqlDbType.NVarChar, 8) { Value = SDay },
                    new SqlParameter("@EDay", SqlDbType.NVarChar, 8) { Value = EDay });

                isLoadingDataGrid1 = true;  // 設置標誌
                try
                {
                    if (dt.Rows.Count == 0)
                    {
                        dataGridView1.DataSource = null;
                    }
                    else
                    {
                        dataGridView1.DataSource = dt;
                        dataGridView1.AutoResizeColumns();

                        dataGridView1.Columns["品號"].Width = 100;
                        dataGridView1.Columns["品名"].Width = 220;
                        dataGridView1.Columns["總庫存量"].Width = 80;
                        dataGridView1.Columns["3F--倉庫"].Width = 70;
                    }
                }
                finally
                {
                    isLoadingDataGrid1 = false; // 清除標誌
                }
            }
            catch (Exception ex)
            {
                ShowError("查詢原料清單", ex);
            }
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            // 如果正在加載數據，則跳過此事件
            if (isLoadingDataGrid1)
                return;

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    MD003 = row.Cells["品號"].Value.ToString().Trim();

                    MESSAGESHOW MSGSHOW = new MESSAGESHOW();
                    // 鎖定控制項
                    this.Enabled = false;
                    // 顯示跳出視窗
                    MSGSHOW.Show();

                    try
                    {
                        SEARCHINVPURMOC(MD003, dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                    }
                    finally
                    {
                        // 關閉跳出視窗
                        MSGSHOW.Close();
                        // 解除鎖定
                        this.Enabled = true;
                    }
                }
                else
                {
                    MD003 = null;
                }
            }
        }

        public void SEARCHINVPURMOC(string MD003, string SDay, string EDay)
        {
            if (string.IsNullOrEmpty(MD003))
            {
                return;
            }

            try
            {
                //查每日原物料的使用量、進貨量、庫存量
                //預計庫存量 (累計) 改在 C# 端計算，SQL Server 2008 R2 不支援 SUM() OVER (ORDER BY)
                string sql = @"
                                WITH
                                BASE_DATA AS (
                                    -- 區塊 1 & 2: MOCMANULINE 生產用量 (負數)
                                    SELECT
                                        [MANU],
                                        CONVERT(NVARCHAR(8), [MANUDATE], 112) AS MANUDATE,
                                        [MD003],
                                        [MD035],
                                        CONVERT(DECIMAL(16, 3),
                                            CASE
                                                -- 處理 MOCTB1 (MOCMANULINE.MOCTA001/002)
                                                WHEN ISNULL(MOCTB1.TB003, '') <> '' AND ISNULL(MOCTB1.TB004 - MOCTB1.TB005, 0) <= 0 THEN 0
                                                WHEN ISNULL(MOCTB1.TB003, '') <> '' AND ISNULL(MOCTB1.TB004 - MOCTB1.TB005, 0) > 0 THEN ISNULL((MOCTB1.TB004 - MOCTB1.TB005) * -1, 0)
                                                -- 處理 MOCTB2 (MOCMANULINEMERGE.NO -> MOCTA.TA033)
                                                WHEN ISNULL(MOCTB2.TB003, '') <> '' AND ISNULL(MOCTB2.TB004 - MOCTB2.TB005, 0) <= 0 THEN 0
                                                WHEN ISNULL(MOCTB2.TB003, '') <> '' AND ISNULL(MOCTB2.TB004 - MOCTB2.TB005, 0) > 0 THEN
                                                    (ISNULL(MOCTB2.TB004 - MOCTB2.TB005, 0) / ISNULL(MC.CNT, 1)) * -1
                                                -- 預設計算邏輯 (BOM 理論用量)
                                                WHEN [MANU] = N'包裝線' THEN ([PACKAGE] / [MC004] * [MD006] / [MD007] * (1 + [MD008])) * -1
                                                ELSE ([NUM] / [MC004] * [MD006] / [MD007] * (1 + [MD008])) * -1
                                            END
                                        ) AS TNUM,
                                        [MB004],
                                        [MOCMANULINE].[MB001],
                                        [MOCMANULINE].[MB002],
                                        CASE WHEN [MANU] = N'包裝線' THEN [PACKAGE] ELSE [NUM] END AS PACKAGE,
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
                                    -- 合併工單筆數 (取代逐筆 COUNT 子查詢)
                                    LEFT JOIN (
                                        SELECT NO, COUNT(NO) AS CNT
                                        FROM [TKMOC].[dbo].[MOCMANULINEMERGE] WITH(NOLOCK)
                                        GROUP BY NO
                                    ) AS MC ON MC.NO = MOCTA.TA033
                                    WHERE
                                        ([MOCMANULINE].MB001 = BOMMC.MC001)
                                        AND (
                                            (ISNULL([MOCMANULINEMERGE].NO, '') <> '' AND ISNULL(MOCTA.TA001, '') <> '')
                                            OR (ISNULL([MOCMANULINEMERGE].NO, '') = '')
                                        )
                                        -- 改為可用索引的日期條件
                                        AND [MANUDATE] >= @SDate
                                        AND [MANUDATE] < DATEADD(DAY, 1, @EDate)
                                        AND [MD003] = @MD003
                                        -- 先由 BOM 找出有用到此原料的成品，縮小排程範圍
                                        AND [MOCMANULINE].MB001 IN (SELECT MD001 FROM [TK].dbo.BOMMD WITH(NOLOCK) WHERE MD003 = @MD003)

                                    UNION ALL

                                    -- 區塊 3: PURTD 進貨數據
                                    SELECT
                                        N'1進貨' AS MANU,
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
                                        AND TD012 BETWEEN @SDay AND @EDay
                                        AND TD004 = @MD003
                                        AND TD007 IN (SELECT [TD007] FROM [TKMOC].dbo.[REPORTINVPURUESD_TD007])

                                    UNION ALL

                                    -- 區塊 4: INVLA 現有庫存，日期放在查詢起日，讓累計從現有庫存開始
                                    SELECT
                                        N'0庫存' AS MANU,
                                        @SDay AS MANUDATE,
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
                                        AND LA001 = @MD003
                                    GROUP BY LA001, MB002, MB004

                                    UNION ALL

                                    -- 區塊 5: INVPURUESD 手動進出貨數據
                                    SELECT
                                        N'1手動進出貨' AS MANU,
                                        CONVERT(NVARCHAR(8), INVPURUESD.DATES, 112) AS MANUDATE,
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
                                        INVPURUESD.DATES >= @SDate
                                        AND INVPURUESD.DATES < DATEADD(DAY, 1, @EDate)
                                        AND INVPURUESD.MB001 = @MD003
                                ),
                                -- 同一天的排序：0 期初庫存 → 1 領用/使用 (含手動出貨) → 2 進貨 (含手動進貨)
                                SORTED_DATA AS (
                                    SELECT
                                        *,
                                        CASE
                                            WHEN MANU = N'0庫存' THEN 0
                                            WHEN MANU = N'1進貨' THEN 2
                                            WHEN MANU = N'1手動進出貨' AND TNUM >= 0 THEN 2
                                            ELSE 1
                                        END AS SORTKEY
                                    FROM BASE_DATA
                                )
                                SELECT
                                    CONVERT(DECIMAL(18, 3), 0) AS '預計庫存量',
                                    ROW_NUMBER() OVER (ORDER BY MANUDATE, SORTKEY, MANU, MB001, COPTD001, COPTD002, COPTD003) AS '列數',
                                    MANU AS '線別',
                                    MANUDATE AS '日期',
                                    MD003 AS '品號',
                                    MD035 AS '品名',
                                    TNUM AS '用量',
                                    MB004 AS '單位',
                                    MB001 AS '成品',
                                    MB002 AS '成品名',
                                    PACKAGE AS '成品數',
                                    BAR AS '桶數',
                                    COPTD001 AS '訂單單別',
                                    COPTD002 AS '訂單單號',
                                    COPTD003 AS '訂單序號'
                                FROM SORTED_DATA
                                ORDER BY MANUDATE, SORTKEY, MANU, MB001, COPTD001, COPTD002, COPTD003;
                                ";

                DataTable dt = QueryTable(sql, 300,
                    new SqlParameter("@SDate", SqlDbType.DateTime) { Value = ToDate(SDay) },
                    new SqlParameter("@EDate", SqlDbType.DateTime) { Value = ToDate(EDay) },
                    new SqlParameter("@SDay", SqlDbType.NVarChar, 8) { Value = SDay },
                    new SqlParameter("@EDay", SqlDbType.NVarChar, 8) { Value = EDay },
                    new SqlParameter("@MD003", SqlDbType.NVarChar, 40) { Value = MD003 });

                if (dt.Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                    return;
                }

                //計算預計庫存量 (累計)，資料已依 日期 → 庫存/領用/進貨 排序
                dt.Columns["預計庫存量"].ReadOnly = false;
                decimal runningTotal = 0;
                foreach (DataRow dr in dt.Rows)
                {
                    if (dr["用量"] != DBNull.Value)
                    {
                        runningTotal += Convert.ToDecimal(dr["用量"]);
                    }
                    dr["預計庫存量"] = runningTotal;
                }

                dataGridView2.DataSource = dt;
                dataGridView2.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                ShowError("查詢每日明細", ex);
            }
        }

        public void SEARCHDG1(string SEARCHSTRING, int INDEX)
        {
            String searchValue = SEARCHSTRING;
            rowIndexDG1 = INDEX;
            bool found = false;

            if (string.IsNullOrEmpty(searchValue))
            {
                return;
            }

            for (int i = INDEX; i < dataGridView1.Rows.Count; i++)
            {
                string code = Convert.ToString(dataGridView1.Rows[i].Cells[0].Value);
                string name = Convert.ToString(dataGridView1.Rows[i].Cells[1].Value);

                if (code.Contains(searchValue) || name.Contains(searchValue))
                {
                    rowIndexDG1 = i;
                    found = true;

                    if (dataGridView1.CurrentRow != null)
                    {
                        dataGridView1.CurrentRow.Selected = false;
                    }
                    dataGridView1.Rows[i].Selected = true;

                    MD003 = code.Trim();

                    SEARCHINVPURMOC(MD003, dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));

                    dataGridView1.FirstDisplayedScrollingRowIndex = i;

                    break;
                }
            }

            if (!found && dataGridView1.Rows.Count > 0)
            {
                if (INDEX == 0)
                {
                    MessageBox.Show("查無 " + searchValue);
                }
                else if (MessageBox.Show("已查到最後一筆，是否從頭開始?", "已查到最後一筆，是否從頭開始?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    SEARCHDG1(searchValue, 0);
                }
            }
        }

        public void ADDINVPURUESD(string KIND, DateTime DATES, string MB001, decimal NUM)
        {
            try
            {
                string sql = @"
                                INSERT INTO [TKMOC].[dbo].[INVPURUESD]
                                ([KIND],[DATES],[MB001],[NUM])
                                VALUES (@KIND, @DATES, @MB001, @NUM)
                                ";

                ExecuteNonQueryTran(sql,
                    new SqlParameter("@KIND", SqlDbType.NVarChar, 50) { Value = KIND },
                    new SqlParameter("@DATES", SqlDbType.DateTime) { Value = DATES.Date },
                    new SqlParameter("@MB001", SqlDbType.NVarChar, 40) { Value = MB001 },
                    new SqlParameter("@NUM", SqlDbType.Decimal) { Precision = 18, Scale = 3, Value = NUM });
            }
            catch (Exception ex)
            {
                ShowError("新增手動進出貨", ex);
            }
        }

        /// <summary>
        /// 只刪除選到的那一筆手動進出貨 (以 品號 + 日期 + 數量 識別)
        /// </summary>
        public int DELINVPURUESD(string MB001, string DATES, decimal NUM)
        {
            try
            {
                string sql = @"
                                DELETE TOP (1) FROM [TKMOC].[dbo].[INVPURUESD]
                                WHERE MB001 = @MB001
                                AND CONVERT(NVARCHAR(8), DATES, 112) = @DATES
                                AND NUM = @NUM
                                ";

                return ExecuteNonQueryTran(sql,
                    new SqlParameter("@MB001", SqlDbType.NVarChar, 40) { Value = MB001 },
                    new SqlParameter("@DATES", SqlDbType.NVarChar, 8) { Value = DATES },
                    new SqlParameter("@NUM", SqlDbType.Decimal) { Precision = 18, Scale = 3, Value = NUM });
            }
            catch (Exception ex)
            {
                ShowError("刪除手動進出貨", ex);
                return 0;
            }
        }

        public void SEARCHCOPTD(string TD002)
        {
            try
            {
                string sql = @"
                                SELECT TD004 AS '品號',TD005  AS '品名',TD006  AS '規格'
                                FROM [TK].dbo.COPTD
                                WHERE TD002 = @TD002
                                ORDER BY TD004
                                ";

                DataTable dt = QueryTable(sql, 60,
                    new SqlParameter("@TD002", SqlDbType.NVarChar, 20) { Value = TD002 });

                if (dt.Rows.Count == 0)
                {
                    dataGridView3.DataSource = null;
                }
                else
                {
                    dataGridView3.DataSource = dt;
                    dataGridView3.AutoResizeColumns();
                }
            }
            catch (Exception ex)
            {
                ShowError("查詢訂單", ex);
            }
        }

        public void SEARCHBOM(string MD001)
        {
            try
            {
                //找下層bom
                string sql = @"
                                SELECT MD003  AS '品號',MD035 AS '品名',MD001 AS '成品號'
                                FROM [TK].dbo.BOMMD
                                WHERE MD001 = @MD001
                                ORDER BY MD001,MD003
                                ";

                DataTable dt = QueryTable(sql, 60,
                    new SqlParameter("@MD001", SqlDbType.NVarChar, 40) { Value = MD001 });

                if (dt.Rows.Count == 0)
                {
                    dataGridView4.DataSource = null;
                }
                else
                {
                    dataGridView4.DataSource = dt;
                    dataGridView4.AutoResizeColumns();
                }
            }
            catch (Exception ex)
            {
                ShowError("查詢BOM", ex);
            }
        }

        // 注意：dataGridView3 / dataGridView4 只更新文字框，不可改到 MD003 (目前選取的原料)
        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            textBox4.Text = null;

            if (dataGridView3.CurrentRow != null && dataGridView3.CurrentRow.Index >= 0)
            {
                textBox4.Text = Convert.ToString(dataGridView3.CurrentRow.Cells["品號"].Value).Trim();
            }
        }

        private void dataGridView4_SelectionChanged(object sender, EventArgs e)
        {
            textBox1.Text = null;

            if (dataGridView4.CurrentRow != null && dataGridView4.CurrentRow.Index >= 0)
            {
                textBox1.Text = Convert.ToString(dataGridView4.CurrentRow.Cells["品號"].Value).Trim();
            }
        }
        private void dataGridView3_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox1.Text = null;
            SEARCHBOM(textBox4.Text.Trim());
        }

        private void dataGridView4_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            isLoadingDataGrid1 = true;  // 避免 SelectionChanged 重複查詢
            try
            {
                SEARCHDG1(textBox1.Text.Trim(), 0);
            }
            finally
            {
                isLoadingDataGrid1 = false;
            }
        }

        public void SHOW_TBALERTMESSAGES()
        {
            DataTable DT = FIND_TBALERTMESSAGES();
            StringBuilder message = new StringBuilder();

            if (DT != null && DT.Rows.Count >= 1)
            {
                foreach (DataRow DR in DT.Rows)
                {
                    message.AppendLine(DR["MESSAGES"].ToString());
                }
            }

            //沒有公告就不跳空白視窗
            if (message.Length > 0)
            {
                MessageBox.Show(message.ToString());
            }
        }

        public DataTable FIND_TBALERTMESSAGES()
        {
            try
            {
                string sql = @"
                                SELECT
                                [ID]
                                ,[MESSAGES]
                                ,[ISCLOSES]
                                FROM [TKMOC].[dbo].[TBALERTMESSAGES]
                                WHERE [ISCLOSES] IN ('Y')
                                ORDER BY [ID]
                                ";

                DataTable dt = QueryTable(sql, 60);
                return dt.Rows.Count >= 1 ? dt : null;
            }
            catch
            {
                //公告讀取失敗不影響主功能
                return null;
            }
        }


        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            rowIndexDG1 = -1;
            SEARCHMOCMANULINE(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
        }
        private void button2_Click(object sender, EventArgs e)
        {
            isLoadingDataGrid1 = true;  // 設置標誌
            try
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
            finally
            {
                isLoadingDataGrid1 = false;  // 清除標誌
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(MD003))
            {
                MessageBox.Show("請先在左邊選擇原料");
                return;
            }

            decimal num;
            if (!decimal.TryParse(textBox2.Text.Trim(), out num) || num == 0)
            {
                MessageBox.Show("數量請輸入數字 (進貨為正數，出貨為負數)");
                return;
            }

            ADDINVPURUESD(KIND_MANUAL, dateTimePicker3.Value, MD003, num);

            SEARCHINVPURMOC(MD003, dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
        }
        private void button4_Click(object sender, EventArgs e)
        {
            DataGridViewRow row = dataGridView2.CurrentRow;

            if (row == null || Convert.ToString(row.Cells["線別"].Value) != KIND_MANUAL)
            {
                MessageBox.Show("請先在明細中選擇一筆「" + KIND_MANUAL + "」的資料");
                return;
            }

            string mb001 = Convert.ToString(row.Cells["品號"].Value).Trim();
            string dates = Convert.ToString(row.Cells["日期"].Value).Trim();
            decimal num = Convert.ToDecimal(row.Cells["用量"].Value);

            string confirm = string.Format("要刪除這筆手動進出貨嗎?{0}品號：{1}{0}日期：{2}{0}數量：{3}", Environment.NewLine, mb001, dates, num);
            if (MessageBox.Show(confirm, "刪除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                int result = DELINVPURUESD(mb001, dates, num);
                if (result == 0)
                {
                    MessageBox.Show("找不到符合的資料，未刪除");
                }
                SEARCHINVPURMOC(mb001, dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
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
