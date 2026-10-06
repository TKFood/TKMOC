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
    public partial class frmREPORTYIELDRATE : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        int rownum = 0;
        Report report1 = new Report();

        public frmREPORTYIELDRATE()
        {
            InitializeComponent();
        }
        private void frmREPORTYIELDRATE_Load(object sender, EventArgs e)
        {
            comboBox1load();
        }
        #region FUNCTION
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
            Sequel.AppendFormat(@"SELECT MD001,MD002 FROM [TK].dbo.CMSMD    WHERE MD003 IN ('20') ORDER BY  MD001 ");

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
        public void SETFASTREPORT()
        {
            StringBuilder SQL1 = new StringBuilder();
            StringBuilder SQL2 = new StringBuilder();

            SQL1 = SETSQL1();

            Report report1 = new Report();
            report1.Load(@"REPORT\得料率報表.frx");

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

        public StringBuilder SETSQL1()
        {
            StringBuilder SB = new StringBuilder();

            //原料用量，品號1開頭或3開頭，單位是KG，直接加總
            //原料用量，3開頭，單位不是KG，就把投入量*淨重/1000，加總

            SB.AppendFormat(@"
                           WITH TEMP AS (
                            SELECT 
                                MD.MD002 AS [線別],
                                TA.TA006 AS [品號],
                                TA.TA034 AS [品名],
                                MB.MB003 AS [規格],
                                TA.TA001 AS [製令單別],
                                TA.TA002 AS [製令單號],
                                TA.TA007 AS [生產單位],
                                TA.TA015 AS [預計產量],
                                TA.TA017 AS [生產量],
                                MB.UDF03 AS [類別],
                                MB.UDF07 AS [淨重],
                                MB.UDF08 AS [單片重],
                                MB.UDF09 AS [袋重],
                                ISNULL(MB.UDF06, 0) AS [蒸發率],
                                MB.UDF01 AS [成品是否扣袋重],
                                MB.UDF02 AS [領料是否扣袋重],
        
                                -- 原料用量
                                (
                                    SELECT ISNULL(SUM(TB005), 0)
                                    FROM [TK].dbo.MOCTB TB WITH (NOLOCK)
                                    WHERE (TB.TB003 LIKE '1%' OR TB.TB003 LIKE '3%')
                                        AND TB.TB007 IN ('kg', 'KG', 'Kg', 'kG', '公升')
                                        AND TB.TB001 = TA.TA001 AND TB.TB002 = TA.TA002
                                ) + (
                                    SELECT ISNULL(SUM(TB.TB005 * BMB.UDF07 / 1000.0), 0)
                                    FROM [TK].dbo.MOCTB TB WITH (NOLOCK)
                                    JOIN [TK].dbo.INVMB BMB WITH (NOLOCK) ON BMB.MB001 = TB.TB003
                                    WHERE TB.TB003 LIKE '3%' 
                                        AND TB.TB007 NOT IN ('kg', 'KG', 'Kg', 'kG', '公升')
                                        AND TB.TB001 = TA.TA001 AND TB.TB002 = TA.TA002
                                ) AS [原料用量],
        
                                -- 成品用量
                                (
                                    SELECT ISNULL(SUM(TB.TB005 * BMB.UDF07), 0)
                                    FROM [TK].dbo.MOCTB TB WITH (NOLOCK)
                                    JOIN [TK].dbo.INVMB BMB WITH (NOLOCK) ON BMB.MB001 = TB.TB003
                                    WHERE TB.TB003 LIKE '4%'
                                        AND TB.TB001 = TA.TA001 AND TB.TB002 = TA.TA002
                                ) AS [成品用量],
        
                                -- 袋重比
                                CASE 
                                    WHEN MB.UDF08 > 0 AND MB.UDF09 > 0 
                                        THEN (1.0 / (MB.UDF08 + MB.UDF09)) * MB.UDF09
                                    ELSE 0 
                                END AS [袋重比],
        
                                -- 桶數
                                CONVERT(DECIMAL(16, 3), (
                                    TA.TA015 / ISNULL((
                                        SELECT TOP 1 MC004 
                                        FROM [TK].dbo.BOMMC WITH (NOLOCK) 
                                        WHERE MC001 = MB.MB001
                                    ), 1)
                                )) AS [桶數]
                            FROM [TK].dbo.MOCTA TA WITH (NOLOCK)
                            JOIN [TK].dbo.INVMB MB WITH (NOLOCK) ON TA.TA006 = MB.MB001
                            LEFT JOIN [TK].dbo.CMSMD MD WITH (NOLOCK) ON TA.TA021 = MD.MD001
                            WHERE TA.TA013 IN ('Y', 'N')
                                AND TA.TA003 >= '{0}'
                                AND TA.TA003 <= '{1}'
                        ),
                        CALC_BASE AS (
                            SELECT 
                                *,
                                (成品用量 / 1000.0) AS 成品用量_K,
                                -- 通用成品類別過濾條件
                                CASE WHEN 類別 NOT IN ('半成品', '個', '試吃', '片', '單包', 'kg') THEN 1 ELSE 0 END AS [is_standard_category],
                                -- 計算基礎量項
                                (原料用量 * (1.0 - (蒸發率 / 100.0))) AS [蒸發後原料],
                                (原料用量 * 袋重比) AS [領料袋重],
                                ((生產量 * 淨重) / 1000.0) AS [成品總淨重_K]
                            FROM TEMP
                        )
                        SELECT 
                            線別,
                            SUBSTRING(製令單號, 1, 8) AS [日期],
                            品號, 品名, 規格, 製令單別, 製令單號, 生產單位, 預計產量, 生產量, 桶數, 淨重, 單片重, 袋重, 袋重比, 蒸發率, 原料用量,
                            成品用量_K AS 成品用量,
                            類別, 領料是否扣袋重, 成品是否扣袋重,

                            -- 各分類獨立得料率指標
                            CASE WHEN is_standard_category = 1 AND 領料是否扣袋重 = 'Y' AND 成品是否扣袋重 = 'Y' AND (蒸發後原料 + 成品用量_K - 領料袋重) > 0 
                                    THEN (成品總淨重_K * (1.0 - 袋重比)) / (蒸發後原料 + 成品用量_K - 領料袋重) ELSE 0 END AS [領料扣成品扣的得料率],
         
                            CASE WHEN is_standard_category = 1 AND 領料是否扣袋重 = 'Y' AND 成品是否扣袋重 = 'N' AND (蒸發後原料 + 成品用量_K) > 0 
                                    THEN 成品總淨重_K / (蒸發後原料 + 成品用量_K - 領料袋重) ELSE 0 END AS [領料扣成品不扣的得料率],
         
                            CASE WHEN is_standard_category = 1 AND 領料是否扣袋重 = 'N' AND 成品是否扣袋重 = 'Y' AND (蒸發後原料 + 成品用量_K - 領料袋重) > 0 
                                    THEN (成品總淨重_K * (1.0 - 袋重比)) / (蒸發後原料 + 成品用量_K) ELSE 0 END AS [領料不扣成品扣的得料率],
         
                            CASE WHEN is_standard_category = 1 AND 領料是否扣袋重 = 'N' AND 成品是否扣袋重 = 'N' AND (蒸發後原料 + 成品用量_K) > 0 
                                    THEN 成品總淨重_K / (蒸發後原料 + 成品用量_K) ELSE 0 END AS [領料不扣成品不扣的得料率],

                            CASE WHEN 類別 = '半成品' AND 原料用量 > 0 AND 成品是否扣袋重 = 'Y' THEN (生產量 * (1.0 - 袋重比)) / 蒸發後原料 ELSE 0 END AS [半成品得料率成品扣袋重],
                            CASE WHEN 類別 = '半成品' AND 原料用量 > 0 AND 成品是否扣袋重 = 'N' THEN 生產量 / 蒸發後原料 ELSE 0 END AS [半成品得料率成品不扣袋重],
                            CASE WHEN 類別 IN ('個', '試吃') AND 蒸發後原料 > 0 THEN 成品總淨重_K / 蒸發後原料 ELSE 0 END AS [個試吃得料率],
                            CASE WHEN 類別 = '片' AND (蒸發後原料 - 領料袋重) > 0 THEN 成品總淨重_K / (蒸發後原料 - 領料袋重) ELSE 0 END AS [片得料率],
                            CASE WHEN 類別 = '單包' AND 原料用量 > 0 THEN 生產量 / 原料用量 ELSE 0 END AS [單包得料率],
                            CASE WHEN 類別 = 'kg' AND (蒸發後原料 - 領料袋重) > 0 THEN 生產量 / (蒸發後原料 + 成品用量_K - 領料袋重) ELSE 0 END AS [kg得料率],
                            CASE WHEN 類別 = '過篩後' AND 原料用量 > 0 THEN 生產量 / 原料用量 ELSE 0 END AS [過篩後得料率],

                            -- 算式中的「分子」總合
                            VAL_NUMERATOR AS [分子],
    
                            -- 算式中的「分母」總合
                            VAL_DENOMINATOR AS [分母],

                            -- 得料率（分子 / 分母，使用 NULLIF 避免 Divide by Zero Error）
                            ISNULL(VAL_NUMERATOR / NULLIF(VAL_DENOMINATOR, 0), 0) AS [得料率]

                        FROM CALC_BASE
                        CROSS APPLY (
                            -- 利用 CROSS APPLY 在單列中進行「分子」與「分母」加總計算，不需重複寫 10 次 CASE WHEN
                            SELECT 
                                -- 計算分子
                                (CASE WHEN is_standard_category = 1 AND 領料是否扣袋重 = 'Y' AND 成品是否扣袋重 = 'Y' AND (蒸發後原料 + 成品用量_K - 領料袋重) > 0 THEN 成品總淨重_K * (1.0 - 袋重比) ELSE 0 END) +
                                (CASE WHEN is_standard_category = 1 AND 領料是否扣袋重 = 'Y' AND 成品是否扣袋重 = 'N' AND (蒸發後原料 + 成品用量_K) > 0 THEN 成品總淨重_K ELSE 0 END) +
                                (CASE WHEN is_standard_category = 1 AND 領料是否扣袋重 = 'N' AND 成品是否扣袋重 = 'Y' AND (蒸發後原料 + 成品用量_K - 領料袋重) > 0 THEN 成品總淨重_K * (1.0 - 袋重比) ELSE 0 END) +
                                (CASE WHEN is_standard_category = 1 AND 領料是否扣袋重 = 'N' AND 成品是否扣袋重 = 'N' AND (蒸發後原料 + 成品用量_K) > 0 THEN 成品總淨重_K ELSE 0 END) +
                                (CASE WHEN 類別 = '半成品' AND 原料用量 > 0 AND 成品是否扣袋重 = 'Y' THEN 生產量 * (1.0 - 袋重比) ELSE 0 END) +
                                (CASE WHEN 類別 = '半成品' AND 原料用量 > 0 AND 成品是否扣袋重 = 'N' THEN 生產量 ELSE 0 END) +
                                (CASE WHEN 類別 IN ('個', '試吃') AND 蒸發後原料 > 0 THEN 成品總淨重_K ELSE 0 END) +
                                (CASE WHEN 類別 = '片' AND (蒸發後原料 - 領料袋重) > 0 THEN 成品總淨重_K ELSE 0 END) +
                                (CASE WHEN 類別 = '單包' AND 原料用量 > 0 THEN 生產量 ELSE 0 END) +
                                (CASE WHEN 類別 = 'kg' AND (蒸發後原料 - 領料袋重) > 0 THEN 生產量 ELSE 0 END) +
                                (CASE WHEN 類別 = '過篩後' AND 原料用量 > 0 THEN 生產量 ELSE 0 END) AS VAL_NUMERATOR,

                                -- 計算分母
                                (CASE WHEN is_standard_category = 1 AND 領料是否扣袋重 = 'Y' AND 成品是否扣袋重 = 'Y' AND (蒸發後原料 + 成品用量_K - 領料袋重) > 0 THEN (蒸發後原料 + 成品用量_K - 領料袋重) ELSE 0 END) +
                                (CASE WHEN is_standard_category = 1 AND 領料是否扣袋重 = 'Y' AND 成品是否扣袋重 = 'N' AND (蒸發後原料 + 成品用量_K) > 0 THEN (蒸發後原料 + 成品用量_K - 領料袋重) ELSE 0 END) +
                                (CASE WHEN is_standard_category = 1 AND 領料是否扣袋重 = 'N' AND 成品是否扣袋重 = 'Y' AND (蒸發後原料 + 成品用量_K - 領料袋重) > 0 THEN (蒸發後原料 + 成品用量_K) ELSE 0 END) +
                                (CASE WHEN is_standard_category = 1 AND 領料是否扣袋重 = 'N' AND 成品是否扣袋重 = 'N' AND (蒸發後原料 + 成品用量_K) > 0 THEN (蒸發後原料 + 成品用量_K) ELSE 0 END) +
                                (CASE WHEN 類別 = '半成品' AND 原料用量 > 0 THEN 蒸發後原料 ELSE 0 END) +
                                (CASE WHEN 類別 IN ('個', '試吃') AND 蒸發後原料 > 0 THEN 蒸發後原料 ELSE 0 END) +
                                (CASE WHEN 類別 = '片' AND (蒸發後原料 - 領料袋重) > 0 THEN (蒸發後原料 - 領料袋重) ELSE 0 END) +
                                (CASE WHEN 類別 = '單包' AND 原料用量 > 0 THEN 原料用量 ELSE 0 END) +
                                (CASE WHEN 類別 = 'kg' AND (蒸發後原料 - 領料袋重) > 0 THEN (蒸發後原料 + 成品用量_K - 領料袋重) ELSE 0 END) +
                                (CASE WHEN 類別 = '過篩後' AND 原料用量 > 0 THEN 原料用量 ELSE 0 END) AS VAL_DENOMINATOR
                        ) CA
                        WHERE 線別 = '{2}'
                        ORDER BY 線別, SUBSTRING(製令單號, 1, 8), 品號;


                                ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"), comboBox1.Text);


            return SB;

        }

      
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }
        #endregion

       
    }
}
