using System;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.Windows.Forms;
using FastReport;
using FastReport.Data;

namespace TKMOC
{
    /// <summary>
    /// 維護 frmREPORTMOCBOM 報表用的「生產說明」資料
    /// 資料表：[TKMOC].[dbo].[REPORTMOCBOMPROCESS] (MB001 品號, MB002 品名, PROCESSING 生產說明)
    /// </summary>
    public partial class frmREPORTMOCBOMPROCESS : Form
    {
        //目前在 DataGridView 選取的品號（修改、刪除時作為條件）
        string SELECTED_MB001 = "";

        public frmREPORTMOCBOMPROCESS()
        {
            InitializeComponent();
        }

        private void frmREPORTMOCBOMPROCESS_Load(object sender, EventArgs e)
        {
            SEARCH(textBoxSearch.Text.Trim(), null);
        }

        #region FUNCTION

        /// <summary>
        /// 查詢資料，keyword 可為品號或品名的部分文字；selectMB001 不為空時，查詢後自動選到該品號
        /// </summary>
        public void SEARCH(string keyword, string selectMB001)
        {
            try
            {
                string sql = @"
                            SELECT
                            [MB001] AS '品號'
                            ,[MB002] AS '品名'
                            ,[PROCESSING] AS '生產說明'
                            FROM [TKMOC].[dbo].[REPORTMOCBOMPROCESS]
                            WHERE (@KEYWORD = '' OR [MB001] LIKE '%' + @KEYWORD + '%' OR [MB002] LIKE '%' + @KEYWORD + '%')
                            ORDER BY [MB001]
                            ";

                DataTable dt = Db.Query(sql, new SqlParameter("@KEYWORD", keyword ?? ""));

                dataGridView1.DataSource = dt;

                if (dataGridView1.Columns.Count > 0)
                {
                    dataGridView1.Columns["品號"].Width = 160;
                    dataGridView1.Columns["品名"].Width = 300;
                    dataGridView1.Columns["生產說明"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                }

                labelCount.Text = "共 " + dt.Rows.Count + " 筆";

                if (!string.IsNullOrEmpty(selectMB001))
                {
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (row.Cells["品號"].Value != null && row.Cells["品號"].Value.ToString() == selectMB001)
                        {
                            dataGridView1.CurrentCell = row.Cells["品號"];
                            break;
                        }
                    }
                }

                if (dt.Rows.Count == 0)
                {
                    CLEAR_TEXTBOX();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("查詢失敗：" + ex.Message);
            }
        }

        /// <summary>
        /// 直接更新 DataGridView 目前選取那一筆的內容（不重新查詢，不會跳回第1筆）
        /// </summary>
        public void UPDATE_CURRENT_ROW(string MB001, string MB002, string PROCESSING)
        {
            if (dataGridView1.CurrentRow == null || !(dataGridView1.CurrentRow.DataBoundItem is DataRowView))
            {
                SEARCH(textBoxSearch.Text.Trim(), MB001);
                return;
            }

            DataRow DR = ((DataRowView)dataGridView1.CurrentRow.DataBoundItem).Row;
            DR["品號"] = MB001;
            DR["品名"] = MB002;
            DR["生產說明"] = PROCESSING;
            DR.AcceptChanges();

            SELECTED_MB001 = MB001;
        }

        public void CLEAR_TEXTBOX()
        {
            SELECTED_MB001 = "";
            textBoxMB001.Text = "";
            textBoxMB002.Text = "";
            textBoxPROCESSING.Text = "";
        }

        /// <summary>
        /// 檢查品號是否已存在
        /// </summary>
        public bool CHECK_EXISTS(string MB001)
        {
            string sql = @"
                            SELECT COUNT(*) AS NUMS
                            FROM [TKMOC].[dbo].[REPORTMOCBOMPROCESS]
                            WHERE [MB001] = @MB001
                            ";
            DataTable dt = Db.Query(sql, new SqlParameter("@MB001", MB001));
            return dt.Rows.Count > 0 && Convert.ToInt32(dt.Rows[0]["NUMS"]) > 0;
        }

        /// <summary>
        /// 從 ERP 品號檔 [TK].dbo.INVMB 帶出品名
        /// </summary>
        public string FIND_MB002(string MB001)
        {
            try
            {
                string sql = @"
                            SELECT TOP 1 [MB002]
                            FROM [TK].dbo.INVMB
                            WHERE [MB001] = @MB001
                            ";
                DataTable dt = Db.Query(sql, new SqlParameter("@MB001", MB001));
                if (dt.Rows.Count > 0)
                {
                    return dt.Rows[0]["MB002"].ToString().Trim();
                }
            }
            catch
            {
            }
            return "";
        }

        /// <summary>
        /// 計算 [TK].dbo.INVMB 中尚未存在 REPORTMOCBOMPROCESS 的品號筆數
        /// </summary>
        public int COUNT_ERP_INVMB_NOT_EXISTS(string keyword)
        {
            string sql = @"
                            SELECT COUNT(*) AS NUMS
                            FROM [TK].dbo.INVMB
                            WHERE (INVMB.[MB001] LIKE '3%' OR INVMB.[MB001] LIKE '4%')
                            AND NOT EXISTS (SELECT 1 FROM [TKMOC].[dbo].[REPORTMOCBOMPROCESS] P WHERE P.[MB001] = INVMB.[MB001])
                            AND (@KEYWORD = '' OR INVMB.[MB001] LIKE '%' + @KEYWORD + '%' OR INVMB.[MB002] LIKE '%' + @KEYWORD + '%')
                            ";
            DataTable dt = Db.Query(sql, new SqlParameter("@KEYWORD", keyword ?? ""));
            return dt.Rows.Count > 0 ? Convert.ToInt32(dt.Rows[0]["NUMS"]) : 0;
        }

        /// <summary>
        /// 載入ERP品號：把 [TK].dbo.INVMB 的 MB001、MB002 新增到 REPORTMOCBOMPROCESS（只新增3、4開頭且還不存在的品號，生產說明先空白）
        /// </summary>
        public int ADD_ERP_INVMB(string keyword)
        {
            string sql = @"
                            INSERT INTO [TKMOC].[dbo].[REPORTMOCBOMPROCESS]
                            ([MB001], [MB002], [PROCESSING])
                            SELECT RTRIM(INVMB.[MB001]), RTRIM(INVMB.[MB002]), ''
                            FROM [TK].dbo.INVMB
                            WHERE (INVMB.[MB001] LIKE '3%' OR INVMB.[MB001] LIKE '4%')
                            AND NOT EXISTS (SELECT 1 FROM [TKMOC].[dbo].[REPORTMOCBOMPROCESS] P WHERE P.[MB001] = INVMB.[MB001])
                            AND (@KEYWORD = '' OR INVMB.[MB001] LIKE '%' + @KEYWORD + '%' OR INVMB.[MB002] LIKE '%' + @KEYWORD + '%')
                            ";

            return Db.ExecuteInTransaction(sql, 300, new SqlParameter("@KEYWORD", keyword ?? ""));
        }

        public void ADD_REPORTMOCBOMPROCESS(string MB001, string MB002, string PROCESSING)
        {
            string sql = @"
                            INSERT INTO [TKMOC].[dbo].[REPORTMOCBOMPROCESS]
                            ([MB001], [MB002], [PROCESSING])
                            VALUES
                            (@MB001, @MB002, @PROCESSING)
                            ";

            int rows = Db.ExecuteInTransaction(sql, 60,
                new SqlParameter("@MB001", MB001),
                new SqlParameter("@MB002", MB002),
                new SqlParameter("@PROCESSING", PROCESSING));

            MessageBox.Show(rows > 0 ? "新增完成" : "新增失敗");
        }

        public int UPDATE_REPORTMOCBOMPROCESS(string ORI_MB001, string MB001, string MB002, string PROCESSING)
        {
            string sql = @"
                            UPDATE [TKMOC].[dbo].[REPORTMOCBOMPROCESS]
                            SET [MB001] = @MB001
                            ,[MB002] = @MB002
                            ,[PROCESSING] = @PROCESSING
                            WHERE [MB001] = @ORI_MB001
                            ";

            int rows = Db.ExecuteInTransaction(sql, 60,
                new SqlParameter("@ORI_MB001", ORI_MB001),
                new SqlParameter("@MB001", MB001),
                new SqlParameter("@MB002", MB002),
                new SqlParameter("@PROCESSING", PROCESSING));

            MessageBox.Show(rows > 0 ? "修改完成" : "修改失敗，找不到品號：" + ORI_MB001);

            return rows;
        }

        public void DELETE_REPORTMOCBOMPROCESS(string MB001)
        {
            string sql = @"
                            DELETE [TKMOC].[dbo].[REPORTMOCBOMPROCESS]
                            WHERE [MB001] = @MB001
                            ";

            int rows = Db.ExecuteInTransaction(sql, 60, new SqlParameter("@MB001", MB001));

            MessageBox.Show(rows > 0 ? "刪除完成" : "刪除失敗，找不到品號：" + MB001);
        }

        /// <summary>
        /// 報表分頁：組查詢 SQL
        /// </summary>
        public string SETSQL_REPORT(string keyword)
        {
            //防止單引號造成 SQL 錯誤
            keyword = (keyword ?? "").Replace("'", "''");

            StringBuilder SB = new StringBuilder();
            SB.AppendFormat(@"
                            SELECT 
                            [MB001] AS '品號'
                            ,[MB002] AS '品名'
                            ,[PROCESSING] AS '生產說明'
                            FROM [TKMOC].[dbo].[REPORTMOCBOMPROCESS]
                            WHERE ('{0}' = '' OR [MB001] LIKE '%{0}%' OR [MB002] LIKE '%{0}%')
                            ORDER BY [MB001]
                            ", keyword);

            return SB.ToString();
        }

        /// <summary>
        /// 報表分頁：產生 FastReport 報表
        /// </summary>
        public void SETFASTREPORT(string keyword)
        {
            try
            {
                Report report1 = new Report();
                report1.Load(@"REPORT\生產說明維護.frx");

                report1.Dictionary.Connections[0].ConnectionString = Db.ConnectionString;

                TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
                table.SelectCommand = SETSQL_REPORT(keyword);

                report1.Preview = previewControl1;
                report1.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show("報表產生失敗：" + ex.Message);
            }
        }

        #endregion

        #region BUTTON / EVENT

        //報表分頁-查詢
        private void button3_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(textBox4.Text.Trim());
        }


        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null && dataGridView1.CurrentRow.DataBoundItem is DataRowView)
            {
                DataRow DR = ((DataRowView)dataGridView1.CurrentRow.DataBoundItem).Row;
                SELECTED_MB001 = DR["品號"].ToString();
                textBoxMB001.Text = DR["品號"].ToString();
                textBoxMB002.Text = DR["品名"].ToString();
                textBoxPROCESSING.Text = DR["生產說明"].ToString();
            }
        }

        //輸入品號離開時，若品名空白，自動從 ERP 帶出品名
        private void textBoxMB001_Leave(object sender, EventArgs e)
        {
            string MB001 = textBoxMB001.Text.Trim();
            if (!string.IsNullOrEmpty(MB001) && string.IsNullOrEmpty(textBoxMB002.Text.Trim()))
            {
                textBoxMB002.Text = FIND_MB002(MB001);
            }
        }

        private void textBoxSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SEARCH(textBoxSearch.Text.Trim(), null);
                e.SuppressKeyPress = true;
            }
        }

        //查詢
        private void buttonSearch_Click(object sender, EventArgs e)
        {
            SEARCH(textBoxSearch.Text.Trim(), null);
        }

        //清除，準備新增
        private void buttonClear_Click(object sender, EventArgs e)
        {
            dataGridView1.ClearSelection();
            CLEAR_TEXTBOX();
            textBoxMB001.Focus();
        }

        //載入ERP品號
        private void buttonAdd_Click(object sender, EventArgs e)
        {
            //載入ERP品號：從 [TK].dbo.INVMB 帶入尚未存在 [TKMOC].[dbo].[REPORTMOCBOMPROCESS] 的品號、品名
            //只載入3、4開頭的品號；上方「品號/品名」查詢欄有輸入時，只載入符合的品號
            string keyword = textBoxSearch.Text.Trim();

            try
            {
                int counts = COUNT_ERP_INVMB_NOT_EXISTS(keyword);
                if (counts == 0)
                {
                    MessageBox.Show("沒有需要載入的ERP品號");
                    return;
                }

                string msg = string.IsNullOrEmpty(keyword)
                    ? "共有 " + counts + " 筆ERP品號尚未建立，確定全部載入？"
                    : "符合「" + keyword + "」的ERP品號共有 " + counts + " 筆尚未建立，確定載入？";

                if (MessageBox.Show(msg, "載入ERP品號", MessageBoxButtons.YesNo) != DialogResult.Yes)
                {
                    return;
                }

                int rows = ADD_ERP_INVMB(keyword);
                MessageBox.Show("已載入 " + rows + " 筆ERP品號");

                SEARCH(keyword, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("載入ERP品號失敗：" + ex.Message);
            }
        }

        //修改
        private void buttonUpdate_Click(object sender, EventArgs e)
        {
            string MB001 = textBoxMB001.Text.Trim();
            string MB002 = textBoxMB002.Text.Trim();
            string PROCESSING = textBoxPROCESSING.Text.Trim();

            if (string.IsNullOrEmpty(SELECTED_MB001))
            {
                MessageBox.Show("請先在上方選擇要修改的資料");
                return;
            }
            if (string.IsNullOrEmpty(MB001) || string.IsNullOrEmpty(PROCESSING))
            {
                MessageBox.Show("品號、生產說明 不可空白");
                return;
            }

            try
            {
                //若有修改品號，新品號不可與其他資料重複
                if (MB001 != SELECTED_MB001 && CHECK_EXISTS(MB001))
                {
                    MessageBox.Show("品號 " + MB001 + " 已存在，不可改成重複的品號");
                    return;
                }

                int rows = UPDATE_REPORTMOCBOMPROCESS(SELECTED_MB001, MB001, MB002, PROCESSING);

                //修改成功後不重新查詢，直接更新目前這一筆，畫面停留在原本的資料與捲動位置
                if (rows > 0)
                {
                    UPDATE_CURRENT_ROW(MB001, MB002, PROCESSING);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("修改失敗：" + ex.Message);
            }
        }

        //刪除
        private void buttonDelete_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(SELECTED_MB001))
            {
                MessageBox.Show("請先在上方選擇要刪除的資料");
                return;
            }

            DialogResult dialogResult = MessageBox.Show("確定要刪除品號 " + SELECTED_MB001 + " 的生產說明？", "刪除確認", MessageBoxButtons.YesNo);
            if (dialogResult != DialogResult.Yes)
            {
                return;
            }

            try
            {
                DELETE_REPORTMOCBOMPROCESS(SELECTED_MB001);
                CLEAR_TEXTBOX();
                SEARCH(textBoxSearch.Text.Trim(), null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("刪除失敗：" + ex.Message);
            }
        }

        #endregion
    }
}
