using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FastReport;
using FastReport.Data;

namespace TKMOC
{
    /// <summary>
    /// 油酥原料添加表
    /// 依製令：[TKMOC].[dbo].[REPORTMOCBOM]
    /// 依成品：[TKMOC].[dbo].[REPORTMOCBOMOIL]
    /// 生產說明：[TKMOC].[dbo].[REPORTMOCBOMPROCESS]（由 frmREPORTMOCBOMPROCESS 維護）
    /// </summary>
    public partial class frmREPORTMOCBOM : Form
    {
        //報表檔
        const string REPORT_FILE_MOC = @"REPORT\油酥原料添加表V7.frx";
        const string REPORT_FILE_OIL = @"REPORT\油酥原料添加表-合併V3.frx";

        static readonly Font HEADER_FONT = new Font("Tahoma", 9);
        static readonly Font CELL_FONT = new Font("Tahoma", 10);

        Report report1 = new Report();

        //全選用的表頭 CheckBox
        CheckBox cbHeader;

        //找出生產說明用的品號
        string MAINMB001 = "";

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
            rect.X = rect.Location.X + rect.Width / 8 - 1;
            rect.Y = rect.Location.Y + (rect.Height / 4 - 1);

            cbHeader = new CheckBox();
            cbHeader.Name = "checkboxHeader";
            cbHeader.Size = new Size(12, 12);
            cbHeader.Location = rect.Location;

            //全选要设定的事件
            cbHeader.CheckedChanged += cbHeader_CheckedChanged;

            //将 CheckBox 加入到 dataGridView
            dataGridView1.Controls.Add(cbHeader);

            #endregion
        }

        private void cbHeader_CheckedChanged(object sender, EventArgs e)
        {
            dataGridView1.EndEdit();

            bool isChecked = cbHeader.Checked;
            foreach (DataGridViewRow dr in dataGridView1.Rows)
            {
                dr.Cells[0].Value = isChecked;
            }
        }

        #region FUNCTION

        /// <summary>
        /// 依製令：查詢生產日的製令
        /// </summary>
        public void Search()
        {
            try
            {
                //MC004=0 時以 NULL 處理，避免除以0讓整個查詢失敗
                string sql = @"
                            SELECT TA001 AS '製令',TA002 AS '單號',TA006 AS '品號',TA034 AS '品名',TA015 AS '生產量',TA003 AS '生產日',TA035 AS '規格',MC004 AS '標準批量',(TA015/NULLIF(MC004,0)) AS '桶數'
                            FROM [TK].dbo.MOCTA
                            INNER JOIN [TK].dbo.BOMMC ON TA006=MC001
                            WHERE (TA006 LIKE '3%' OR TA006 LIKE '4%')
                            AND TA021 IN ('02','03','04')
                            AND TA003=@TA003
                            ORDER BY TA001,TA002
                            ";

                DataTable dt = Db.Query(sql, new SqlParameter("@TA003", dateTimePicker1.Value.ToString("yyyyMMdd")));

                if (dt.Rows.Count >= 1)
                {
                    dataGridView1.DataSource = dt;

                    dataGridView1.ColumnHeadersDefaultCellStyle.Font = HEADER_FONT;
                    dataGridView1.DefaultCellStyle.Font = CELL_FONT;
                    dataGridView1.Columns["製令"].Width = 60;
                    dataGridView1.Columns["單號"].Width = 100;
                    dataGridView1.Columns["品號"].Width = 100;
                    dataGridView1.Columns["品名"].Width = 120;
                }
                else
                {
                    dataGridView1.DataSource = null;
                    textBox1.Text = "";
                    textBox2.Text = "";
                    textBox3.Text = "";
                    MAINMB001 = "";
                }

                cbHeader.Checked = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show("查詢失敗：" + ex.Message);
            }
        }

        /// <summary>
        /// 依成品：查詢生產日的成品(依線別)
        /// </summary>
        public void SearchOIL()
        {
            try
            {
                string sql = @"
                            SELECT MD002 AS '線別',TA006 AS '品號',TA034 AS '品名',SUM(TA015) AS '總生產量',TA003 AS '生產日',TA035 AS '規格',MC004 AS '標準批量',(SUM(TA015)/NULLIF(MC004,0)) AS '桶數',TA021 '線別代號'
                            FROM [TK].dbo.MOCTA
                            INNER JOIN [TK].dbo.BOMMC ON TA006=MC001
                            INNER JOIN [TK].dbo.CMSMD ON TA021=MD001
                            WHERE (TA006 LIKE '3%' OR TA006 LIKE '4%')
                            AND TA021 IN ('02','03','04')
                            AND TA003=@TA003
                            GROUP BY MD002,TA006,TA034,TA015,TA003,TA035,MC004,TA021
                            ORDER BY MD002,TA006,TA034,TA015,TA003,TA035,MC004,TA021
                            ";

                DataTable dt = Db.Query(sql, new SqlParameter("@TA003", dateTimePicker2.Value.ToString("yyyyMMdd")));

                if (dt.Rows.Count >= 1)
                {
                    dataGridView2.DataSource = dt;

                    dataGridView2.ColumnHeadersDefaultCellStyle.Font = HEADER_FONT;
                    dataGridView2.DefaultCellStyle.Font = CELL_FONT;
                    dataGridView2.Columns["線別"].Width = 100;
                    dataGridView2.Columns["品號"].Width = 100;
                    dataGridView2.Columns["品名"].Width = 120;
                }
                else
                {
                    dataGridView2.DataSource = null;
                    textBox4.Text = "";
                    textBox5.Text = "";
                    textBox6.Text = "";
                    textBox7.Text = "";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("查詢失敗：" + ex.Message);
            }
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null && dataGridView1.CurrentRow.Index >= 0)
            {
                DataGridViewRow row = dataGridView1.CurrentRow;
                textBox1.Text = CellText(row, "製令");
                textBox2.Text = CellText(row, "單號");
                textBox3.Text = CellText(row, "桶數");

                MAINMB001 = CellText(row, "品號");
            }
            else
            {
                textBox1.Text = "";
                textBox2.Text = "";
                textBox3.Text = "";
            }
        }

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView2.CurrentRow != null && dataGridView2.CurrentRow.Index >= 0)
            {
                DataGridViewRow row = dataGridView2.CurrentRow;
                textBox4.Text = CellText(row, "生產日");
                textBox5.Text = CellText(row, "品號");
                textBox6.Text = CellText(row, "桶數");
                textBox7.Text = CellText(row, "線別代號");
            }
            else
            {
                textBox4.Text = "";
                textBox5.Text = "";
                textBox6.Text = "";
                textBox7.Text = "";
            }
        }

        /// <summary>
        /// 安全取得儲存格文字（null / DBNull 回傳空字串）
        /// </summary>
        static string CellText(DataGridViewRow row, string columnName)
        {
            return Convert.ToString(row.Cells[columnName].Value).Trim();
        }

        /// <summary>
        /// 桶數文字轉數字；空白、非數字、小於等於0 回傳 false
        /// </summary>
        static bool TryParseBuckets(string text, out decimal buckets)
        {
            buckets = 0;
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }
            return decimal.TryParse(text.Trim(), NumberStyles.Number, CultureInfo.CurrentCulture, out buckets) && buckets > 0;
        }

        #region 桶數拆分

        /// <summary>
        /// 一桶的內容：第幾桶、用量倍率(1=滿桶)
        /// </summary>
        class BOX
        {
            public int BOXS;
            public decimal RATIO;
            public BOX(int boxs, decimal ratio) { BOXS = boxs; RATIO = ratio; }
        }

        /// <summary>
        /// 依桶數拆分每一桶的用量倍率
        /// 剛好滿桶：第1~N桶都是滿桶
        /// 未滿桶：小於1桶 → 只有第1桶(未滿)；大於1桶 → 第1桶滿、第2桶未滿、其他滿桶
        /// </summary>
        static List<BOX> BUILD_BOXES(decimal buckets)
        {
            List<BOX> boxes = new List<BOX>();
            if (buckets <= 0)
            {
                return boxes;
            }

            int counts = (int)Math.Ceiling(buckets);

            if (buckets == Math.Floor(buckets))
            {
                for (int i = 1; i <= counts; i++)
                {
                    boxes.Add(new BOX(i, 1m));
                }
            }
            else if (buckets < 1)
            {
                boxes.Add(new BOX(1, buckets));
            }
            else
            {
                decimal small = buckets - (counts - 1);
                boxes.Add(new BOX(1, 1m));
                boxes.Add(new BOX(2, small));
                for (int i = 3; i <= counts; i++)
                {
                    boxes.Add(new BOX(i, 1m));
                }
            }

            return boxes;
        }

        /// <summary>
        /// 把桶數轉成 SQL 的 VALUES 表：(VALUES (1,1),(2,0.3),...) AS B(BOXS,RATIO)
        /// 內容都是程式產生的數字，以 InvariantCulture 輸出，沒有注入風險
        /// </summary>
        static string BOXES_VALUES_SQL(List<BOX> boxes)
        {
            return "(VALUES " + string.Join(",", boxes.Select(b =>
                string.Format(CultureInfo.InvariantCulture, "({0},CONVERT(DECIMAL(18,6),{1}))", b.BOXS, b.RATIO))) + ") AS B(BOXS,RATIO)";
        }

        #endregion

        /// <summary>
        /// 依製令：產生 [REPORTMOCBOM]，一次 INSERT 全部桶數
        /// 滿桶：MD006/MD007；未滿桶：CONVERT(DECIMAL(16,3),MD006/MD007*倍率)
        /// </summary>
        public void ADDTOREPORTMOCBOM(string TA001, string TA002, decimal BUCKETS)
        {
            List<BOX> boxes = BUILD_BOXES(BUCKETS);
            if (boxes.Count == 0)
            {
                return;
            }

            bool isFull = BUCKETS == Math.Floor(BUCKETS);

            //組合SQL條件
            string SQL_Condition = BuildLikeCondition(FIND_REPORTMOCBOMMD003());

            //註：原程式「剛好滿桶」才排除 301400%，「未滿桶」沒有排除，這裡保留原本行為
            string EXCLUDE_301400 = isFull ? "AND MD003 NOT LIKE '301400%'" : "";

            string sql = string.Format(@"
                            DELETE [TKMOC].[dbo].[REPORTMOCBOM]

                            INSERT INTO [TKMOC].[dbo].[REPORTMOCBOM]
                            ([TA001],[TA002],[TA006],[TA034],[BOXS],[MD003],[MB002],[MD006])
                            SELECT TA001,TA002,TA006,TA034,B.BOXS,MD003,MB002
                                ,CASE WHEN B.RATIO=1 THEN MD006/MD007 ELSE CONVERT(DECIMAL(16,3),MD006/MD007*B.RATIO) END
                            FROM [TK].dbo.MOCTA
                            INNER JOIN [TK].dbo.BOMMD ON TA006=MD001
                            INNER JOIN [TK].dbo.INVMB ON MD003=MB001
                            CROSS JOIN {0}
                            WHERE 1=1
                            {1}
                            {2}
                            AND MB002 NOT LIKE '%水麵%'
                            AND TA001=@TA001 AND TA002=@TA002
                            ORDER BY B.BOXS,MD003
                            ", BOXES_VALUES_SQL(boxes), SQL_Condition, EXCLUDE_301400);

            Db.ExecuteInTransaction(sql, 60,
                new SqlParameter("@TA001", TA001),
                new SqlParameter("@TA002", TA002));
        }

        /// <summary>
        /// 依成品：產生 [REPORTMOCBOMOIL]，一次 INSERT 全部桶數
        /// </summary>
        public void ADDTOREPORTMOCBOMOIL(string TA003, string TA006, decimal BUCKETS, string TA021)
        {
            List<BOX> boxes = BUILD_BOXES(BUCKETS);
            if (boxes.Count == 0)
            {
                return;
            }

            string sql = string.Format(@"
                            DELETE [TKMOC].[dbo].[REPORTMOCBOMOIL]

                            INSERT INTO [TKMOC].[dbo].[REPORTMOCBOMOIL]
                            ([TA003],[TA021],[TA006],[TA034],[BOXS],[MD003],[MB002],[MD006])
                            SELECT TA003,TA021,TA006,TA034,B.BOXS,MD003,MB002
                                ,CASE WHEN B.RATIO=1 THEN MD006/MD007 ELSE MD006/MD007*B.RATIO END
                            FROM [TK].dbo.MOCTA
                            INNER JOIN [TK].dbo.BOMMD ON TA006=MD001
                            INNER JOIN [TK].dbo.INVMB ON MD003=MB001
                            CROSS JOIN {0}
                            WHERE (MD003 LIKE '1%' OR MD003 LIKE '301%')
                            AND MD003 NOT LIKE '301400%'
                            AND TA003=@TA003 AND TA006=@TA006 AND TA021=@TA021
                            ORDER BY B.BOXS,MD003
                            ", BOXES_VALUES_SQL(boxes));

            Db.ExecuteInTransaction(sql, 60,
                new SqlParameter("@TA003", TA003),
                new SqlParameter("@TA006", TA006),
                new SqlParameter("@TA021", TA021));
        }

        /// <summary>
        /// 依製令-單一製令報表
        /// </summary>
        public void SETREPORT(string TA001, string TA002, string BUCKETS, string MAINMB001)
        {
            decimal BUCKETSORI;
            if (TryParseBuckets(BUCKETS, out BUCKETSORI))
            {
                ADDTOREPORTMOCBOM(TA001, TA002, BUCKETSORI);
            }

            SHOWREPORT(REPORT_FILE_MOC, previewControl1, SETSQL(null, null, null), SETSQLPROCESS(MAINMB001));
        }

        /// <summary>
        /// 依製令-多張製令合併報表
        /// </summary>
        public void SETREPORT2(string TA001, string TA002, decimal BUCKETSORI, string LINK_TA001TA002, string LINK_TA006, string LINK_TA034, string MAINMB001)
        {
            if (BUCKETSORI > 0)
            {
                ADDTOREPORTMOCBOM(TA001, TA002, BUCKETSORI);
            }

            SHOWREPORT(REPORT_FILE_MOC, previewControl1, SETSQL(LINK_TA001TA002, LINK_TA006, LINK_TA034), SETSQLPROCESS(MAINMB001));
        }

        /// <summary>
        /// 依成品-油酥原料添加表-合併
        /// </summary>
        public void SETREPORTOIL(string TA003, string TA006, string BUCKETS, string TA021)
        {
            decimal BUCKETSORI;
            if (TryParseBuckets(BUCKETS, out BUCKETSORI))
            {
                ADDTOREPORTMOCBOMOIL(TA003, TA006, BUCKETSORI, TA021);
            }

            SHOWREPORT(REPORT_FILE_OIL, previewControl2, SETSQLOIL(), null);
        }

        /// <summary>
        /// 載入 FastReport 報表並顯示在指定的 PreviewControl
        /// </summary>
        void SHOWREPORT(string reportFile, FastReport.Preview.PreviewControl preview, string sqlTable, string sqlTable1)
        {
            report1 = new Report();
            report1.Load(reportFile);
            report1.Dictionary.Connections[0].ConnectionString = Db.ConnectionString;

            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = sqlTable;

            if (sqlTable1 != null)
            {
                TableDataSource table1 = report1.GetDataSource("Table1") as TableDataSource;
                table1.SelectCommand = sqlTable1;
            }

            report1.Preview = preview;
            report1.Show();
        }

        /// <summary>
        /// SQL 字串常值跳脫（FastReport 的 SelectCommand 無法帶參數）
        /// </summary>
        static string SqlLiteral(string value)
        {
            return "N'" + (value ?? "").Replace("'", "''") + "'";
        }

        /// <summary>
        /// 依製令報表 SQL
        /// LINK_TA001TA002 為 null 時，製令/成品/成品名 取資料表內容；否則顯示合併後的文字
        /// </summary>
        public string SETSQL(string LINK_TA001TA002, string LINK_TA006, string LINK_TA034)
        {
            bool isLink = LINK_TA001TA002 != null;
            string COL_MOC = isLink ? SqlLiteral(LINK_TA001TA002) : "[TA001]+[TA002]";
            string COL_TA006 = isLink ? SqlLiteral(LINK_TA006) : "TA006";
            string COL_TA034 = isLink ? SqlLiteral(LINK_TA034) : "TA034";

            return string.Format(@"
                            SELECT [ID]
                            ,{0} AS '製令'
                            ,'第'+CONVERT(nvarchar,[BOXS])+'桶' AS '桶數'
                            ,{1} AS '成品'
                            ,{2} AS '成品名'
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
                            WHERE [MD003] NOT  IN (SELECT [MB001]  FROM [TKMOC].[dbo].[LIMITSREPORTMOCBOM]  )
                            ORDER BY [TA001],[TA002],[BOXS],[MD003]
                            ", COL_MOC, COL_TA006, COL_TA034);
        }

        /// <summary>
        /// 生產說明 SQL：完全相同的品號優先；品號空白時不帶任何說明
        /// </summary>
        public string SETSQLPROCESS(string MAINMB001)
        {
            string MB001 = (MAINMB001 ?? "").Trim().Replace("'", "''");

            return string.Format(@"
                            SELECT TOP 1
                            [PROCESSING]
                            FROM [TKMOC].[dbo].[REPORTMOCBOMPROCESS]
                            WHERE N'{0}'<>''
                            AND MB001 LIKE N'%{0}%'
                            ORDER BY CASE WHEN MB001=N'{0}' THEN 0 ELSE 1 END,MB001
                            ", MB001);
        }

        public string SETSQLOIL()
        {
            return @"
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
                            ";
        }

        /// <summary>
        /// 檢查多張製令的原料(1開頭)是否單身品號元件一致、組成用量一致
        /// 回傳不一致的元件；全部一致回傳 null
        /// </summary>
        public DataTable CHECK_BOMMD(List<string> MD001)
        {
            if (MD001 == null || MD001.Count == 0)
            {
                return null;
            }

            List<SqlParameter> parameters = new List<SqlParameter>();
            List<string> unions = new List<string>();
            for (int i = 0; i < MD001.Count; i++)
            {
                string p = "@MD001_" + i;
                parameters.Add(new SqlParameter(p, MD001[i]));
                //用 UNION ALL：同一成品選兩張製令時也要算兩次
                unions.Add("SELECT MD001, MD003, MD006, MD007 FROM [TK].dbo.BOMMD WHERE MD003 LIKE '1%' AND MD001 = " + p);
            }
            parameters.Add(new SqlParameter("@COUNTS", MD001.Count));

            string sql = @"
                            SELECT MD003, MD006, MD007, COUNT(MD003) COUNTS
                            FROM (
                            " + string.Join("\r\n UNION ALL \r\n", unions) + @"
                            ) AS CombinedData
                            WHERE MD003 NOT IN (SELECT [MD003] FROM [TKMOC].[dbo].[REPORTMOCBOMNOSET])
                            GROUP BY MD003, MD006, MD007
                            HAVING COUNT(MD003) < @COUNTS
                            ";

            DataTable dt = Db.Query(sql, parameters.ToArray());
            return dt.Rows.Count >= 1 ? dt : null;
        }

        public DataTable FIND_REPORTMOCBOMMD003()
        {
            DataTable dt = Db.Query(@"SELECT [MD003] FROM [TKMOC].[dbo].[REPORTMOCBOMMD003]");
            return dt.Rows.Count >= 1 ? dt : null;
        }

        /// <summary>
        /// 組合 AND (MD003 LIKE 'x%' OR MD003 LIKE 'y%' ...)
        /// </summary>
        public string BuildLikeCondition(DataTable dt)
        {
            if (dt == null || dt.Rows.Count == 0)
            {
                return string.Empty;
            }

            List<string> conditions = new List<string>();
            foreach (DataRow row in dt.Rows)
            {
                string value = Convert.ToString(row["MD003"]).Trim();
                if (!string.IsNullOrEmpty(value))
                {
                    conditions.Add(string.Format("MD003 LIKE '{0}%'", value.Replace("'", "''"))); // 防止 SQL injection
                }
            }

            return conditions.Count == 0 ? string.Empty : "AND (" + string.Join(" OR ", conditions) + ")";
        }

        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //先結束編輯，剛勾選的 CheckBox 才會生效
            dataGridView1.EndEdit();

            List<DataGridViewRow> checkedRows = dataGridView1.Rows.Cast<DataGridViewRow>()
                .Where(dr => dr.Cells[0].Value is bool && (bool)dr.Cells[0].Value)
                .ToList();

            Cursor = Cursors.WaitCursor;
            try
            {
                if (checkedRows.Count == 0)
                {
                    //沒有勾選：用目前選取的製令
                    if (string.IsNullOrEmpty(textBox1.Text.Trim()))
                    {
                        MessageBox.Show("請先查詢並選擇製令");
                        return;
                    }
                    SETREPORT(textBox1.Text.Trim(), textBox2.Text.Trim(), textBox3.Text.Trim(), MAINMB001);
                    return;
                }

                //多張製令合併
                List<string> MD001 = checkedRows.Select(dr => CellText(dr, "品號")).ToList();

                //CHECK  原料需單身品號元件一致、組成用量一致
                DataTable DT = CHECK_BOMMD(MD001);
                if (DT != null)
                {
                    StringBuilder MESS = new StringBuilder("原料需單身品號元件不一致 或 組成用量不一致，不能合併\n");
                    foreach (DataRow ROW in DT.Rows)
                    {
                        MESS.AppendFormat("品號:{0}用量:{1}底數:{2}\n", ROW["MD003"], ROW["MD006"], ROW["MD007"]);
                    }
                    MessageBox.Show(MESS.ToString());
                    return;
                }

                StringBuilder LINK_TA001TA002 = new StringBuilder();
                StringBuilder LINK_TA006 = new StringBuilder();
                StringBuilder LINK_TA034 = new StringBuilder();
                decimal BUCKETS = 0;
                string TA001 = "";
                string TA002 = "";

                foreach (DataGridViewRow dr in checkedRows)
                {
                    TA001 = CellText(dr, "製令");
                    TA002 = CellText(dr, "單號");
                    LINK_TA001TA002.Append(TA001).Append(TA002).Append("*");
                    LINK_TA006.Append(CellText(dr, "品號")).Append("*");
                    LINK_TA034.Append(CellText(dr, "品名")).Append("*");

                    decimal b;
                    if (decimal.TryParse(CellText(dr, "桶數"), out b))
                    {
                        BUCKETS = Math.Round(BUCKETS + b, 3);
                    }

                    MAINMB001 = CellText(dr, "品號");
                }

                //BOM 已確認一致，以最後一張製令的 BOM 乘上合計桶數
                SETREPORT2(TA001, TA002, BUCKETS, LINK_TA001TA002.ToString(), LINK_TA006.ToString(), LINK_TA034.ToString(), MAINMB001);
            }
            catch (Exception ex)
            {
                MessageBox.Show("報表產生失敗：" + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SearchOIL();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //油酥原料添加表-合併
            if (string.IsNullOrEmpty(textBox5.Text.Trim()))
            {
                MessageBox.Show("請先查詢並選擇成品");
                return;
            }

            Cursor = Cursors.WaitCursor;
            try
            {
                SETREPORTOIL(textBox4.Text.Trim(), textBox5.Text.Trim(), textBox6.Text.Trim(), textBox7.Text.Trim());
            }
            catch (Exception ex)
            {
                MessageBox.Show("報表產生失敗：" + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        #endregion

        #region 依製令(製程排序)：REPORTMOCBOMMD009（含 BOMMD.MD009 製程順序）
        //與「依製令」相同流程，但寫入 [TKMOC].[dbo].[REPORTMOCBOMMD009]，報表依 MD009 排序
        //報表檔：REPORT\油酥原料添加表V7-製程排序.frx；建表：CREATE_REPORTMOCBOMMD009.sql
        //不使用、不修改原本的 [REPORTMOCBOM]

        const string REPORT_FILE_MD009 = @"REPORT\油酥原料添加表V7-製程排序.frx";

        Report report3 = new Report();

        //全選用的表頭 CheckBox（第一次查詢時才建立）
        CheckBox cbHeaderMD009;

        //找出生產說明用的品號
        string MAINMB001_MD009 = "";

        #region FUNCTION-MD009

        /// <summary>
        /// 第一次使用時建立勾選欄與全選 CheckBox（不動原本的 frmREPORTMOCBOM_Load）
        /// </summary>
        void INIT_DATAGRIDVIEW3()
        {
            if (cbHeaderMD009 != null)
            {
                return;
            }

            DataGridViewCheckBoxColumn cbCol = new DataGridViewCheckBoxColumn();
            cbCol.Width = 50;
            cbCol.HeaderText = "　選擇";
            cbCol.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cbCol.TrueValue = true;
            cbCol.FalseValue = false;
            dataGridView3.Columns.Insert(0, cbCol);

            Rectangle rect = dataGridView3.GetCellDisplayRectangle(0, -1, true);
            rect.X = rect.Location.X + rect.Width / 8 - 1;
            rect.Y = rect.Location.Y + (rect.Height / 4 - 1);

            cbHeaderMD009 = new CheckBox();
            cbHeaderMD009.Name = "checkboxHeaderMD009";
            cbHeaderMD009.Size = new Size(12, 12);
            cbHeaderMD009.Location = rect.Location;
            cbHeaderMD009.CheckedChanged += cbHeaderMD009_CheckedChanged;
            dataGridView3.Controls.Add(cbHeaderMD009);
        }

        private void cbHeaderMD009_CheckedChanged(object sender, EventArgs e)
        {
            dataGridView3.EndEdit();

            bool isChecked = cbHeaderMD009.Checked;
            foreach (DataGridViewRow dr in dataGridView3.Rows)
            {
                dr.Cells[0].Value = isChecked;
            }
        }

        /// <summary>
        /// 依製令(製程排序)：查詢生產日的製令
        /// </summary>
        public void SearchMD009()
        {
            INIT_DATAGRIDVIEW3();

            try
            {
                string sql = @"
                            SELECT TA001 AS '製令',TA002 AS '單號',TA006 AS '品號',TA034 AS '品名',TA015 AS '生產量',TA003 AS '生產日',TA035 AS '規格',MC004 AS '標準批量',(TA015/NULLIF(MC004,0)) AS '桶數'
                            FROM [TK].dbo.MOCTA
                            INNER JOIN [TK].dbo.BOMMC ON TA006=MC001
                            WHERE (TA006 LIKE '3%' OR TA006 LIKE '4%')
                            AND TA021 IN ('02','03','04')
                            AND TA003=@TA003
                            ORDER BY TA001,TA002
                            ";

                DataTable dt = Db.Query(sql, new SqlParameter("@TA003", dateTimePicker3.Value.ToString("yyyyMMdd")));

                if (dt.Rows.Count >= 1)
                {
                    dataGridView3.DataSource = dt;

                    dataGridView3.ColumnHeadersDefaultCellStyle.Font = HEADER_FONT;
                    dataGridView3.DefaultCellStyle.Font = CELL_FONT;
                    dataGridView3.Columns["製令"].Width = 60;
                    dataGridView3.Columns["單號"].Width = 100;
                    dataGridView3.Columns["品號"].Width = 100;
                    dataGridView3.Columns["品名"].Width = 120;
                }
                else
                {
                    dataGridView3.DataSource = null;
                    textBox8.Text = "";
                    textBox9.Text = "";
                    textBox10.Text = "";
                    MAINMB001_MD009 = "";
                }

                cbHeaderMD009.Checked = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show("查詢失敗：" + ex.Message);
            }
        }

        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView3.CurrentRow != null && dataGridView3.CurrentRow.Index >= 0)
            {
                DataGridViewRow row = dataGridView3.CurrentRow;
                textBox8.Text = CellText(row, "製令");
                textBox9.Text = CellText(row, "單號");
                textBox10.Text = CellText(row, "桶數");

                MAINMB001_MD009 = CellText(row, "品號");
            }
            else
            {
                textBox8.Text = "";
                textBox9.Text = "";
                textBox10.Text = "";
            }
        }

        /// <summary>
        /// 產生 [REPORTMOCBOMMD009]（含 MD009 製程代號），桶數拆分規則與 ADDTOREPORTMOCBOM 相同
        /// </summary>
        public int ADDTOREPORTMOCBOMMD009(string TA001, string TA002, decimal BUCKETS)
        {
            List<BOX> boxes = BUILD_BOXES(BUCKETS);
            if (boxes.Count == 0)
            {
                return 0;
            }

            bool isFull = BUCKETS == Math.Floor(BUCKETS);

            //組合SQL條件（與原本依製令相同）
            string SQL_Condition = BuildLikeCondition(FIND_REPORTMOCBOMMD003());

            //與原本依製令相同：「剛好滿桶」才排除 301400%
            string EXCLUDE_301400 = isFull ? "AND MD003 NOT LIKE '301400%'" : "";

            string sql = string.Format(@"
                            DELETE [TKMOC].[dbo].[REPORTMOCBOMMD009]

                            INSERT INTO [TKMOC].[dbo].[REPORTMOCBOMMD009]
                            ([TA001],[TA002],[TA006],[TA034],[BOXS],[MD003],[MB002],[MD006],[MD009])
                            SELECT TA001,TA002,TA006,TA034,B.BOXS,MD003,MB002
                                ,CASE WHEN B.RATIO=1 THEN MD006/MD007 ELSE CONVERT(DECIMAL(16,3),MD006/MD007*B.RATIO) END
                                ,RTRIM(ISNULL(MD009,''))
                            FROM [TK].dbo.MOCTA
                            INNER JOIN [TK].dbo.BOMMD ON TA006=MD001
                            INNER JOIN [TK].dbo.INVMB ON MD003=MB001
                            CROSS JOIN {0}
                            WHERE 1=1
                            {1}
                            {2}
                            AND MB002 NOT LIKE '%水麵%'
                            AND TA001=@TA001 AND TA002=@TA002
                            ORDER BY B.BOXS,MD009,MD003
                            ", BOXES_VALUES_SQL(boxes), SQL_Condition, EXCLUDE_301400);

            return Db.ExecuteInTransaction(sql, 60,
                new SqlParameter("@TA001", TA001),
                new SqlParameter("@TA002", TA002));
        }

        /// <summary>
        /// 報表 SQL：對應 油酥原料添加表V7-製程排序.frx（[MD009] AS '製程順序'，排除 REPORTMOCBOMMDNOTIN 的品號），依 MD009 排序
        /// LINK_TA001TA002 為 null 時，製令/成品/成品名 取資料表內容；否則顯示合併後的文字
        /// </summary>
        public string SETSQLMD009(string LINK_TA001TA002, string LINK_TA006, string LINK_TA034)
        {
            bool isLink = LINK_TA001TA002 != null;
            string COL_MOC = isLink ? SqlLiteral(LINK_TA001TA002) : "[TA001]+[TA002]";
            string COL_TA006 = isLink ? SqlLiteral(LINK_TA006) : "TA006";
            string COL_TA034 = isLink ? SqlLiteral(LINK_TA034) : "TA034";

            return string.Format(@"
                            SELECT [ID]
                            ,[MD009] AS '製程順序'
                            ,{0} AS '製令'
                            ,'第'+CONVERT(nvarchar,[BOXS])+'桶' AS '桶數'
                            ,{1} AS '成品'
                            ,{2} AS '成品名'
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
                            ,(SELECT SUM([MD006]) FROM [TKMOC].[dbo].[REPORTMOCBOMMD009] RE WHERE [MD003] NOT  IN ('101001009','3010000111') AND RE.[BOXS]=[REPORTMOCBOMMD009].[BOXS]) AS '每桶重'
                            ,(SELECT SUM([MD006]) FROM [TKMOC].[dbo].[REPORTMOCBOMMD009] WHERE [MD003] NOT  IN ('101001009','3010000111')  ) AS '總重'
                            ,CASE WHEN BOMMC.UDF06=0 THEN 1 ELSE BOMMC.UDF06 END
                            ,'顆數:'+CONVERT(nvarchar,((SELECT SUM([MD006]) FROM [TKMOC].[dbo].[REPORTMOCBOMMD009] RE WHERE [MD003] NOT  IN ('101001009','3010000111') AND RE.[BOXS]=[REPORTMOCBOMMD009].[BOXS])/(CASE WHEN BOMMC.UDF06=0 THEN 1 ELSE BOMMC.UDF06 END))) AS '每桶顆數'
                            ,((SELECT SUM([MD006]) FROM [TKMOC].[dbo].[REPORTMOCBOMMD009] WHERE [MD003] NOT  IN ('101001009','3010000111') )/(CASE WHEN BOMMC.UDF06=0 THEN 1 ELSE BOMMC.UDF06 END)) AS '總顆數'
                            ,'配方 '+ BOMMC.UDF03 AS '配方比'

                            FROM [TKMOC].[dbo].[REPORTMOCBOMMD009]
                            LEFT JOIN [TK].dbo.BOMMC ON MC001=TA006
                            WHERE NOT EXISTS (SELECT [MD003] FROM [TKMOC].[dbo].[REPORTMOCBOMMDNOTIN] WHERE [REPORTMOCBOMMDNOTIN].[MD003]=[REPORTMOCBOMMD009].[MD003])
                            ORDER BY [TA001],[TA002],[BOXS],[MD009],[MD003]
                            ", COL_MOC, COL_TA006, COL_TA034);
        }

        /// <summary>
        /// 單一製令報表
        /// </summary>
        public void SETREPORTMD009(string TA001, string TA002, string BUCKETS, string MAINMB001)
        {
            decimal BUCKETSORI;
            if (!TryParseBuckets(BUCKETS, out BUCKETSORI))
            {
                MessageBox.Show("桶數不正確（" + BUCKETS + "），未產生資料");
                return;
            }

            ADDTOREPORTMOCBOMMD009(TA001, TA002, BUCKETSORI);

            SHOWREPORTMD009(SETSQLMD009(null, null, null), SETSQLPROCESS(MAINMB001));
        }

        /// <summary>
        /// 多張製令合併報表
        /// </summary>
        public void SETREPORT2MD009(string TA001, string TA002, decimal BUCKETSORI, string LINK_TA001TA002, string LINK_TA006, string LINK_TA034, string MAINMB001)
        {
            if (BUCKETSORI <= 0)
            {
                MessageBox.Show("合計桶數不正確（" + BUCKETSORI + "），未產生資料");
                return;
            }

            ADDTOREPORTMOCBOMMD009(TA001, TA002, BUCKETSORI);

            SHOWREPORTMD009(SETSQLMD009(LINK_TA001TA002, LINK_TA006, LINK_TA034), SETSQLPROCESS(MAINMB001));
        }

        /// <summary>
        /// 新報表檔 油酥原料添加表V7-製程排序.frx：列標題 製程順序 → 品號 → 品名 → A製造B有效
        /// </summary>
        void SHOWREPORTMD009(string sqlTable, string sqlTable1)
        {
            report3 = new Report();
            report3.Load(REPORT_FILE_MD009);
            report3.Dictionary.Connections[0].ConnectionString = Db.ConnectionString;

            TableDataSource table = report3.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = sqlTable;

            TableDataSource table1 = report3.GetDataSource("Table1") as TableDataSource;
            table1.SelectCommand = sqlTable1;

            report3.Preview = previewControl3;
            report3.Show();
        }

        #endregion

        #region BUTTON-MD009

        //依製令(製程排序)-查詢
        private void button5_Click(object sender, EventArgs e)
        {
            SearchMD009();
        }

        //依製令(製程排序)-報表
        private void button6_Click(object sender, EventArgs e)
        {
            INIT_DATAGRIDVIEW3();

            //先結束編輯，剛勾選的 CheckBox 才會生效
            dataGridView3.EndEdit();

            List<DataGridViewRow> checkedRows = dataGridView3.Rows.Cast<DataGridViewRow>()
                .Where(dr => dr.Cells[0].Value is bool && (bool)dr.Cells[0].Value)
                .ToList();

            Cursor = Cursors.WaitCursor;
            try
            {
                if (checkedRows.Count == 0)
                {
                    //沒有勾選：用目前選取的製令
                    if (string.IsNullOrEmpty(textBox8.Text.Trim()))
                    {
                        MessageBox.Show("請先查詢並選擇製令");
                        return;
                    }
                    SETREPORTMD009(textBox8.Text.Trim(), textBox9.Text.Trim(), textBox10.Text.Trim(), MAINMB001_MD009);
                    return;
                }

                //多張製令合併
                List<string> MD001 = checkedRows.Select(dr => CellText(dr, "品號")).ToList();

                //CHECK  原料需單身品號元件一致、組成用量一致
                DataTable DT = CHECK_BOMMD(MD001);
                if (DT != null)
                {
                    StringBuilder MESS = new StringBuilder("原料需單身品號元件不一致 或 組成用量不一致，不能合併\n");
                    foreach (DataRow ROW in DT.Rows)
                    {
                        MESS.AppendFormat("品號:{0}用量:{1}底數:{2}\n", ROW["MD003"], ROW["MD006"], ROW["MD007"]);
                    }
                    MessageBox.Show(MESS.ToString());
                    return;
                }

                StringBuilder LINK_TA001TA002 = new StringBuilder();
                StringBuilder LINK_TA006 = new StringBuilder();
                StringBuilder LINK_TA034 = new StringBuilder();
                decimal BUCKETS = 0;
                string TA001 = "";
                string TA002 = "";

                foreach (DataGridViewRow dr in checkedRows)
                {
                    TA001 = CellText(dr, "製令");
                    TA002 = CellText(dr, "單號");
                    LINK_TA001TA002.Append(TA001).Append(TA002).Append("*");
                    LINK_TA006.Append(CellText(dr, "品號")).Append("*");
                    LINK_TA034.Append(CellText(dr, "品名")).Append("*");

                    decimal b;
                    if (decimal.TryParse(CellText(dr, "桶數"), out b))
                    {
                        BUCKETS = Math.Round(BUCKETS + b, 3);
                    }

                    MAINMB001_MD009 = CellText(dr, "品號");
                }

                //BOM 已確認一致，以最後一張製令的 BOM 乘上合計桶數
                SETREPORT2MD009(TA001, TA002, BUCKETS, LINK_TA001TA002.ToString(), LINK_TA006.ToString(), LINK_TA034.ToString(), MAINMB001_MD009);
            }
            catch (Exception ex)
            {
                MessageBox.Show("報表產生失敗：" + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        #endregion

        #endregion
    }
}
