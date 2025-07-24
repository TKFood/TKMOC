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
using TKITDLL;

namespace TKMOC
{
    public partial class frmMOCMANULINEBAKING : Form
    {
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
     
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();
        int result;

        string MANU = "";
        // 宣告一個變數來儲存使用者手動選擇排序的欄位
        string SortedColumn = string.Empty;
        string SortedModel = string.Empty;

        string ID;
        DateTime dt1 ;
        string MB001B;
        string MB002B;
        string MB003B;
        decimal BOX ;
        decimal SUM21;
        string TA001 = "A513";
        string TA002;
        string TA020;
        string TA026 ;
        string TA027 ;
        string TA028;
        string TA029;
        string SUBID;
        string SUBBAR2;
        string SUBNUM2;
        string SUBBOX;
        string SUBPACKAGE2;
        string DELID;
        string DELMOCTA001B;
        string DELMOCTA002B;
        string TF001 = "";
        string TF002 = "";
        string TF003 = "";
        string TF104 = "";
        DateTime dt_DV4 = new DateTime();
        string ID_DV4 = null;
        string SUBID_DV4 = null;
        decimal SUBBAR_DV4 = 0;
        decimal SUBNUM_DV4 = 0;
        decimal SUM_DV4 = 0;
        decimal SUBBOX_DV4 = 0;
        decimal SUBPACKAGE_DV4 = 0;
        decimal BOX_DV4 = 0;
        string TA026_DV4 = null;
        string TA027_DV4 = null;
        string TA028_DV4 = null;
        string MB001_DV4 = null;
        string MB002_DV4 = null;
        string MB003_DV4 = null;
        string TA029_DV4 = null;
        string DELID_DV6 = null;
        string DELMOCTA001_DV6 = null;
        string DELMOCTA002_DV6 = null;

        public class MOCTADATA
        {
            public string COMPANY;
            public string CREATOR;
            public string USR_GROUP;
            public string CREATE_DATE;
            public string MODIFIER;
            public string MODI_DATE;
            public string FLAG;
            public string CREATE_TIME;
            public string MODI_TIME;
            public string TRANS_TYPE;
            public string TRANS_NAME;
            public string sync_count;
            public string DataGroup;
            public string TA001;
            public string TA002;
            public string TA003;
            public string TA004;
            public string TA005;
            public string TA006;
            public string TA007;
            public string TA009;
            public string TA010;
            public string TA011;
            public string TA012;
            public string TA013;
            public string TA014;
            public string TA015;
            public string TA016;
            public string TA017;
            public string TA018;
            public string TA019;
            public string TA020;
            public string TA021;
            public string TA022;
            public string TA024;
            public string TA025;
            public string TA026;
            public string TA027;
            public string TA028;
            public string TA029;
            public string TA030;
            public string TA031;
            public string TA033;
            public string TA034;
            public string TA035;
            public string TA040;
            public string TA041;
            public string TA043;
            public string TA044;
            public string TA045;
            public string TA046;
            public string TA047;
            public string TA049;
            public string TA050;
            public string TA200;
        }

        public frmMOCMANULINEBAKING()
        {
            InitializeComponent();

           
        }

        #region FUNCTION
        private void frmMOCMANULINEBAKING_Load(object sender, EventArgs e)
        {
            MANU = "烘焙生產線";

            DV_CheckBox();

            comboBox1load();
            comboBox2load();
            comboBox3load();
            comboBox21load();
            comboBox23load();
            comboBox24load();
            comboBox25load();
            comboBox26load();
            comboBox28load();

            comboBox4load();
            comboBox5load();
            comboBox7load();


        }

        public void DV_CheckBox()
        {     
            //先建立個 CheckBox 欄
            DataGridViewCheckBoxColumn cbCol28 = new DataGridViewCheckBoxColumn();
            cbCol28.Width = 120;   //設定寬度
            cbCol28.HeaderText = "　選擇";
            cbCol28.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;   //置中
            cbCol28.TrueValue = true;
            cbCol28.FalseValue = false;
            dataGridView28.Columns.Insert(0, cbCol28);

            //region 建立全选 CheckBox

            //建立个矩形，等下计算 CheckBox 嵌入 GridView 的位置
            Rectangle rect = dataGridView28.GetCellDisplayRectangle(0, -1, true);
            rect.X = rect.Location.X + rect.Width / 4 - 18;
            rect.Y = rect.Location.Y + (rect.Height / 2 - 9);

            CheckBox cbHeader = new CheckBox();
            cbHeader.Name = "checkboxHeader";
            cbHeader.Size = new Size(18, 18);
            cbHeader.Location = rect.Location;

            ////全选要设定的事件
            cbHeader.CheckedChanged += new EventHandler(cbHeader_CheckedChanged28);

            //将 CheckBox 加入到 dataGridView
            dataGridView28.Controls.Add(cbHeader);

        }
        private void cbHeader_CheckedChanged28(object sender, EventArgs e)
        {
            dataGridView28.EndEdit();

            foreach (DataGridViewRow dr in dataGridView28.Rows)
            {
                dr.Cells[0].Value = ((CheckBox)dataGridView28.Controls.Find("checkboxHeader", true)[0]).Checked;

            }

        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage1"])
            {
                MANU = "烘焙生產線";
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage4"])
            {               
                MANU = "烘焙包裝線";
            }


        }
        public void comboBox21load()
        {
            LoadComboBoxData(comboBox21, "SELECT [ID],[LAYERS] FROM [TKMOC].[dbo].[MOCMANULINELAYERS] ORDER BY [ID] ", "ID", "LAYERS");
        }
        public void comboBox1load()
        {
            LoadComboBoxData(comboBox1, "SELECT MD001,MD002 FROM [TK].dbo.CMSMD WHERE MD001 IN ('08')  ", "MD002", "MD002");
        }
        public void comboBox2load()
        {
            LoadComboBoxData(comboBox2, "SELECT MD001,MD002 FROM [TK].dbo.CMSMD WHERE MD001 IN ('08')  ", "MD002", "MD002");
        }
     
        public void comboBox3load()
        {
            LoadComboBoxData(comboBox3, "SELECT MC001 ,MC001+MC002 AS 'MC002' FROM [TK].dbo.CMSMC WHERE MC001 LIKE '21%'  ORDER BY MC001 ", "MC001", "MC002");
        }
        public void comboBox4load()
        {
            LoadComboBoxData(comboBox4, "SELECT MD001,MD002 FROM [TK].dbo.CMSMD WHERE MD001 IN ('12')  ", "MD002", "MD002");
        }
        public void comboBox5load()
        {
            LoadComboBoxData(comboBox5, "SELECT MD001,MD002 FROM [TK].dbo.CMSMD WHERE MD001 IN ('12')  ", "MD002", "MD002");
        }
        public void comboBox7load()
        {
            LoadComboBoxData(comboBox7, "SELECT MC001 ,MC001+MC002 AS 'MC002' FROM [TK].dbo.CMSMC WHERE MC001 LIKE '21%'  ORDER BY MC001 ", "MC001", "MC002");
        }
        public void comboBox23load()
        {
            LoadComboBoxData(comboBox23, "SELECT '未核單' AS 'STATUS' UNION ALL SELECT '已核單' AS 'STATUS' ", "STATUS", "STATUS");
        }
        public void comboBox24load()
        {
            LoadComboBoxData(comboBox24, "SELECT 'Y' AS 'STATUS' UNION ALL SELECT 'N' AS 'STATUS' ", "STATUS", "STATUS");
        }
        public void comboBox25load()
        {
            LoadComboBoxData(comboBox25, "SELECT MD001,MD002 FROM [TK].dbo.CMSMD    WHERE ( MD002 LIKE '烘焙生產線%'  ) ", "MD002", "MD002");
        }
        public void comboBox26load()
        {
            LoadComboBoxData(comboBox26, "SELECT '未核單' AS 'STATUS' UNION ALL SELECT '已核單' AS 'STATUS' ", "STATUS", "STATUS");
        }
        public void comboBox28load()
        {
            LoadComboBoxData(comboBox28, "SELECT MD001,MD002 FROM [TK].dbo.CMSMD    WHERE ( MD002 LIKE '烘焙生產線%'  )  ", "MD002", "MD002");
        }
        public void LoadComboBoxData(ComboBox comboBox, string query, string valueMember, string displayMember)
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            using (SqlConnection connection = new SqlConnection(sqlsb.ConnectionString))
            {
                SqlCommand command = new SqlCommand(query, connection);
                connection.Open();

                SqlDataAdapter adapter = new SqlDataAdapter(command);
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);

                comboBox.DataSource = dataTable;
                comboBox.ValueMember = valueMember;
                comboBox.DisplayMember = displayMember;
            }
        }

        public void SEARCHMOCMANULINE_BAKING(string SDATES, string MANU)
        {
            sbSql.Clear();
            sbSqlQuery.Clear();

            if (MANU.Equals("烘焙生產線"))
            {
                sbSql.Append(BuildSqlQuery(MANU, SDATES));
                SEARCH_MANULINE(sbSql.ToString(), dataGridView1, SortedColumn, SortedModel);
            }
            else if (MANU.Equals("烘焙包裝線"))
            {
                sbSql.Append(BuildSqlQuery(MANU, SDATES));
                SEARCH_MANULINE(sbSql.ToString(), dataGridView4, SortedColumn, SortedModel);
            }
        }

        /// <summary>
        /// 組合查詢 SQL（共用區塊）
        /// </summary>
        private string BuildSqlQuery(string manu, string sdates)
        {
            return string.Format(@"
                                    SELECT 
                                        [MANU] AS '線別',
                                        CONVERT(varchar(100), [MANUDATE], 112) AS '生產日',
                                        [MOCMANULINEBAKING].[MB001] AS '品號',
                                        [MOCMANULINEBAKING].[MB002] AS '品名',
                                        [MOCMANULINEBAKING].[MB003] AS '規格',
                                        ALLERGEN AS '過敏原',
                                        ORI AS '素別',
                                        [BAR] AS '桶數',
                                        [PACKAGE] AS '包裝數',
                                        [NUM] AS '數量',
                                        [CLINET] AS '客戶',
                                        [OUTDATE] AS '交期',
                                        [TA029] AS '備註',
                                        [HALFPRO] AS '半成品數量',
                                        [COPTD001] AS '訂單單別',
                                        [COPTD002] AS '訂單號',
                                        [COPTD003] AS '訂單序號',
                                        [BOX] AS '箱數',
                                        [SERNO],
                                        [ID]
                                    FROM [TKMOC].[dbo].[MOCMANULINEBAKING]
                                    LEFT JOIN [TKMOC].[dbo].[ERPINVMB] 
                                        ON [ERPINVMB].MB001 = [MOCMANULINEBAKING].MB001
                                    WHERE [MANU] = '{0}' 
                                        AND CONVERT(varchar(100), [MANUDATE], 112) LIKE '{1}%'
                                    ORDER BY [MANUDATE], [SERNO]"
                                            , manu, sdates);
        }

        public void SEARCH_MANULINE(string QUERY, DataGridView dataGridViewNew, string sortedColumn, string sortedModel)
        {
            var dataSet = new DataSet();
            dataGridViewNew.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;

            try
            {
                // 解密連線字串
                var TKID = new Class1();
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);
                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                using (var conn = new SqlConnection(sqlsb.ConnectionString))
                using (var adapter = new SqlDataAdapter(QUERY, conn))
                {
                    conn.Open();
                    dataSet.Clear();
                    adapter.Fill(dataSet, "Result");
                }

                if (dataSet.Tables["Result"].Rows.Count > 0)
                {
                    dataGridViewNew.DataSource = null;
                    dataGridViewNew.DataSource = dataSet.Tables["Result"];
                    dataGridViewNew.AutoResizeColumns();

                    // 排序處理
                    if (!string.IsNullOrEmpty(sortedColumn) && dataGridViewNew.Columns.Contains(sortedColumn))
                    {
                        var direction = sortedModel == "Descending"
                            ? ListSortDirection.Descending
                            : ListSortDirection.Ascending;

                        dataGridViewNew.Sort(dataGridViewNew.Columns[sortedColumn], direction);
                    }

                    // 欄位寬度微調
                    if (dataGridViewNew.Columns.Contains("規格"))
                        dataGridViewNew.Columns["規格"].Width = 100;
                    if (dataGridViewNew.Columns.Contains("過敏原"))
                        dataGridViewNew.Columns["過敏原"].Width = 30;
                    if (dataGridViewNew.Columns.Contains("素別"))
                        dataGridViewNew.Columns["素別"].Width = 50;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"資料查詢錯誤：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            string MB001 = textBox7.Text.Trim();

            SEARCHMB001(MB001);
            SEARCHMOCMANULINETEMPDATAS(MB001);
        }

        public void SEARCHMB001(string MB001)
        {
            DataTable dt = new DataTable();

            try
            {
                Class1 tkid = new Class1();
                var connStr = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                var sqlsb = new SqlConnectionStringBuilder(connStr);

                sqlsb.Password = tkid.Decryption(sqlsb.Password);
                sqlsb.UserID = tkid.Decryption(sqlsb.UserID);

                string sql = $@"
                                SELECT MB001, MB002, MB003, MC004, MB017 
                                FROM [TK].dbo.INVMB, [TK].dbo.BOMMC
                                WHERE MB001 = MC001
                                AND MB001 = @MB001";

                using (SqlConnection conn = new SqlConnection(sqlsb.ConnectionString))
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                {
                    cmd.Parameters.AddWithValue("@MB001", MB001);                    
                    conn.Open();
                    adapter.Fill(dt);
                
                }

                if (dt != null && dt.Rows.Count > 0)
                {
                    DataRow row = dt.Rows[0];

                    if (MANU == "烘焙生產線")
                    {
                        textBox10.Text = row["MB002"].ToString();
                        textBox11.Text = row["MB003"].ToString();
                        textBox33.Text = row["MC004"].ToString();
                    }
                    else if (MANU == "烘焙包裝線")
                    {
                        textBox16.Text = row["MB002"].ToString();
                        textBox17.Text = row["MB003"].ToString();
                        textBox14.Text = row["MC004"].ToString();
                    }
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show("查詢資料時發生錯誤：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
               
            }
        }
        public void SEARCHMOCMANULINETEMPDATAS(string MB001)
        {
            StringBuilder sbSql = new StringBuilder();
            DataSet ds1 = new DataSet();
            DataSet TEMPds = new DataSet();

            if (MANU.Equals("烘焙生產線") || MANU.Equals("烘焙包裝線"))
            {
                try
                {
                    // 解密資料庫連線字串
                    Class1 TKID = new Class1();
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    // 組合 SQL 語法
                    sbSql.Clear();
                    sbSql.AppendFormat(@"
                                        SELECT [ID]  
                                        FROM [TKMOC].[dbo].[MOCMANULINETEMP] 
                                        WHERE [MB001]='{0}' 
                                        AND [ID] NOT IN (SELECT [ID] FROM [TKMOC].[dbo].[MOCMANULINEBAKING])", MB001);

                    using (SqlConnection conn = new SqlConnection(sqlsb.ConnectionString))
                    using (SqlCommand cmd = new SqlCommand(sbSql.ToString(), conn))
                    using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                    {
                        conn.Open();
                        ds1.Clear();
                        adapter.Fill(ds1, "ds1");
                    }

                    if (ds1.Tables["ds1"] != null && ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        TEMPds.Clear();

                        // 顯示新增畫面
                        frmMOCMANULINESubTEMPADD MOCMANULINESubTEMPADD = new frmMOCMANULINESubTEMPADD(MB001, TEMPds);
                        MOCMANULINESubTEMPADD.ShowDialog();
                        TEMPds = MOCMANULINESubTEMPADD.SETDATASET;

                        if (TEMPds.Tables[0].Rows.Count >= 1)
                        {
                            foreach (DataRow dr in TEMPds.Tables[0].Rows)
                            {
                                SUM21 += Convert.ToDecimal(dr["數量"].ToString());
                                // SUM2 += Convert.ToDecimal(dr["箱數"].ToString());
                            }
                        }
                    }
                }
                catch
                {
                    // 可補上 Log 或錯誤提示
                }
            }
        }



        public void ADDMOCMANULINE(
                         string MANU,
                         string MANUDATE,
                         string MB001,
                         string MB002,
                         string MB003,
                         string CLINET,
                         string MANUHOUR,
                         string BOX,
                         string NUM,
                         string PACKAGE,
                         string OUTDATE,
                         string TA029,
                         string HALFPRO,
                         string COPTD001,
                         string COPTD002,
                         string COPTD003,
                         string BAR)
        {
            Guid newGuid = Guid.NewGuid();

            try
            {
                // 解密資料庫連線
                Class1 TKID = new Class1();
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);
                 //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                using (SqlConnection sqlConn = new SqlConnection(sqlsb.ConnectionString))
                {
                    sqlConn.Open();
                    using (SqlTransaction tran = sqlConn.BeginTransaction())
                    using (SqlCommand cmd = sqlConn.CreateCommand())
                    {
                        cmd.Transaction = tran;
                        cmd.CommandTimeout = 60;
                        cmd.CommandText = @"
                                            INSERT INTO [TKMOC].[dbo].[MOCMANULINEBAKING]
                                            ([ID],[MANU],[MANUDATE],[MB001],[MB002],[MB003],[CLINET],[MANUHOUR],[BOX],[NUM],[PACKAGE],[OUTDATE],[TA029],[HALFPRO],[COPTD001],[COPTD002],[COPTD003],[BAR])
                                            VALUES 
                                            (@ID,@MANU,@MANUDATE,@MB001,@MB002,@MB003,@CLINET,@MANUHOUR,@BOX,@NUM,@PACKAGE,@OUTDATE,@TA029,@HALFPRO,@COPTD001,@COPTD002,@COPTD003,@BAR)";

                        cmd.Parameters.AddWithValue("@ID", newGuid.ToString());
                        cmd.Parameters.AddWithValue("@MANU", MANU);
                        cmd.Parameters.AddWithValue("@MANUDATE", MANUDATE);
                        cmd.Parameters.AddWithValue("@MB001", MB001);
                        cmd.Parameters.AddWithValue("@MB002", MB002);
                        cmd.Parameters.AddWithValue("@MB003", MB003);
                        cmd.Parameters.AddWithValue("@CLINET", CLINET);
                        cmd.Parameters.AddWithValue("@MANUHOUR", MANUHOUR);
                        cmd.Parameters.AddWithValue("@BOX", BOX);
                        cmd.Parameters.AddWithValue("@NUM", NUM);
                        cmd.Parameters.AddWithValue("@PACKAGE", PACKAGE);
                        cmd.Parameters.AddWithValue("@OUTDATE", OUTDATE);
                        cmd.Parameters.AddWithValue("@TA029", TA029);
                        cmd.Parameters.AddWithValue("@HALFPRO", HALFPRO);
                        cmd.Parameters.AddWithValue("@COPTD001", COPTD001);
                        cmd.Parameters.AddWithValue("@COPTD002", COPTD002);
                        cmd.Parameters.AddWithValue("@COPTD003", COPTD003);
                        cmd.Parameters.AddWithValue("@BAR", BAR);

                        int result = cmd.ExecuteNonQuery();

                        if (result > 0)
                            tran.Commit();
                        else
                            tran.Rollback();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("新增失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            SEARCHMOCMANULINE_BAKING(dateTimePicker1.Value.ToString("yyyyMMdd"), comboBox1.Text.Trim());
        }


        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    textBoxID.Text = row.Cells["ID"].Value.ToString();

                    ID = row.Cells["ID"].Value.ToString();
                    dt1 = Convert.ToDateTime(row.Cells["生產日"].Value.ToString().Substring(0, 4) + "/" + row.Cells["生產日"].Value.ToString().Substring(4, 2) + "/" + row.Cells["生產日"].Value.ToString().Substring(6, 2));
                    MB001B = row.Cells["品號"].Value.ToString();
                    MB002B = row.Cells["品名"].Value.ToString();
                    MB003B = row.Cells["規格"].Value.ToString();
                    BOX = Convert.ToDecimal(row.Cells["包裝數"].Value.ToString());
                    SUM21 = Convert.ToDecimal(row.Cells["包裝數"].Value.ToString());
                    TA029 = row.Cells["備註"].Value.ToString();
                    TA026 = row.Cells["訂單單別"].Value.ToString();
                    TA027 = row.Cells["訂單號"].Value.ToString();
                    TA028 = row.Cells["訂單序號"].Value.ToString();

                    SUBID = row.Cells["ID"].Value.ToString();
                    SUBBAR2 = "";
                    SUBNUM2 = "";
                    SUBBOX = row.Cells["箱數"].Value.ToString();
                    SUBPACKAGE2= row.Cells["數量"].Value.ToString();

                    SEARCH_MOCMANULINERESULTBAKING(ID);
                    //SEARCHMOCMANULINEMERGERESLUTMOCTA(ID2.ToString());
                    ////SEARCHMOCMANULINECOP();

                }
                else
                {
                    ID = null;                   
                    SUBBAR2 = null;
                    SUBNUM2 = null;
                    SUBBOX = null;
                    SUBPACKAGE2 = null;
                    TA026 = null;
                    TA027 = null;
                    TA028 = null;

                }
            }
        }

        private void dataGridView4_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView4.CurrentRow != null)
            {
                int rowindex = dataGridView4.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView4.Rows[rowindex];
                    textBoxID2.Text = row.Cells["ID"].Value.ToString();

                    ID_DV4 = row.Cells["ID"].Value.ToString();
                    dt_DV4 = Convert.ToDateTime(row.Cells["生產日"].Value.ToString().Substring(0, 4) + "/" + row.Cells["生產日"].Value.ToString().Substring(4, 2) + "/" + row.Cells["生產日"].Value.ToString().Substring(6, 2));
                    MB001_DV4 = row.Cells["品號"].Value.ToString();
                    MB002_DV4 = row.Cells["品名"].Value.ToString();
                    MB003_DV4 = row.Cells["規格"].Value.ToString();
                    BOX_DV4 = Convert.ToDecimal(row.Cells["包裝數"].Value.ToString());
                    SUM_DV4 = Convert.ToDecimal(row.Cells["包裝數"].Value.ToString());
                    TA029_DV4 = row.Cells["備註"].Value.ToString();
                    TA026_DV4 = row.Cells["訂單單別"].Value.ToString();
                    TA027_DV4 = row.Cells["訂單號"].Value.ToString();
                    TA028_DV4 = row.Cells["訂單序號"].Value.ToString();

                    SUBID_DV4 = row.Cells["ID"].Value.ToString();
                    SUBBAR_DV4 = 0;
                    SUBNUM_DV4 = 0;
                    SUBBOX_DV4 = Convert.ToDecimal(row.Cells["箱數"].Value.ToString());
                    SUBPACKAGE_DV4 = Convert.ToDecimal(row.Cells["數量"].Value.ToString());

                    SEARCH_MOCMANULINERESULTBAKING_DV4(ID_DV4);
                    //SEARCHMOCMANULINEMERGERESLUTMOCTA(ID2.ToString());
                    ////SEARCHMOCMANULINECOP();

                }
                else
                {
                    ID_DV4 = null;
                    SUBBAR_DV4 = 0;
                    SUM_DV4 = 0;
                    SUBBOX_DV4 = 0;
                    SUBPACKAGE_DV4 = 0;
                    TA026_DV4 = null;
                    TA027_DV4 = null;
                    TA028_DV4 = null;

                }
            }
        }

        public void DELMOCMANULINE(string ID)
        {
            try
            {
                // 解密連線字串
                Class1 TKID = new Class1();
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                using (SqlConnection conn = new SqlConnection(sqlsb.ConnectionString))
                {
                    conn.Open();
                    using (SqlTransaction tran = conn.BeginTransaction())
                    using (SqlCommand cmd = new SqlCommand())
                    {
                        cmd.Connection = conn;
                        cmd.Transaction = tran;
                        cmd.CommandTimeout = 60;
                        cmd.CommandText = @"
                                            DELETE FROM [TKMOC].[dbo].[MOCMANULINEBAKING]
                                            WHERE ID = @ID";

                        cmd.Parameters.AddWithValue("@ID", ID);

                        // 除錯：印出參數化 SQL 的樣貌
                        string debugSql = cmd.CommandText.Replace("@ID", $"'{ID}'");
                        Console.WriteLine("執行的 SQL：\n" + debugSql);

                        int result = cmd.ExecuteNonQuery();

                        if (result == 0)
                        {
                            tran.Rollback();
                            MessageBox.Show("刪除失敗，未找到符合的資料");
                        }
                        else
                        {
                            tran.Commit();
                            // 可以加成功訊息或記錄
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("刪除失敗：" + ex.Message);
            }

            // 重新查詢畫面資料
            SEARCHMOCMANULINE_BAKING(dateTimePicker1.Value.ToString("yyyyMMdd"), comboBox1.Text.Trim());
        }


        public void CHECKMOCTAB()
        {
            string CHECKID = null;

            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage1"])
            {
                CHECKID = textBoxID.Text.Trim();
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage4"])
            {
                CHECKID = textBoxID2.Text.Trim();
            }

            if (string.IsNullOrEmpty(CHECKID))
            {
                MessageBox.Show("請輸入 CHECKID");
                return;
            }

            try
            {
                // 解密連線字串
                Class1 TKID = new Class1();
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                using (SqlConnection conn = new SqlConnection(sqlsb.ConnectionString))
                {
                    string sql = @"
                                SELECT 
                                MOCTA001, MOCTA002
                                FROM [TKMOC].[dbo].[MOCMANULINERESULTBAKING]
                                WHERE [SID] = @SID
                                UNION ALL
                                SELECT TA001, TA002
                                FROM [TK].[dbo].[MOCTA]
                                WHERE EXISTS (
                                    SELECT 1
                                    FROM [TKMOC].[dbo].[MOCMANULINERESULTBAKING]
                                    WHERE [SID] = @SID
                                    AND TA001 = MOCTA001 AND TA002 = MOCTA002
                                )";

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@SID", CHECKID);
                        using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                        {
                            DataTable dt = new DataTable();
                            adapter.Fill(dt);

                            if (dt.Rows.Count == 0)
                            {
                                UPDATEMOCMANULINE(CHECKID);
                            }
                            else
                            {
                                MessageBox.Show("ERP 跟外中有製令未刪除，請檢查一下");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("檢查時發生錯誤：" + ex.Message);
            }
        }

        public void UPDATEMOCMANULINE(string CHECKID)
        {
            if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage1"])
            {
                frmMOCMANULINES_BAKING_SUB MOCMANULINE_BAKUING_Sub = new frmMOCMANULINES_BAKING_SUB(CHECKID);
                MOCMANULINE_BAKUING_Sub.ShowDialog();
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage4"])
            {
                frmMOCMANULINES_BAKING_SUB MOCMANULINE_BAKUING_Sub = new frmMOCMANULINES_BAKING_SUB(CHECKID);
                MOCMANULINE_BAKUING_Sub.ShowDialog();
            }

        }

        public void SEARCHCOPDEFAULT(string TD001, string TD002, string TD003)
        {
            try
            {
                // 解密連線
                var TKID = new Class1();
                var sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);
                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                using (SqlConnection conn = new SqlConnection(sqlsb.ConnectionString))
                {
                    string sql = @"
                                    SELECT 
                                        TC053, TD004, TD005, TD006, (TD008 + TD024) AS TD008, TD010, 
                                        (TC015 + '-' + TD020) AS TC015, TD013,
                                        (CASE WHEN ISNULL(MD002, '') <> '' THEN (TD008 + TD024) * MD004 ELSE (TD008 + TD024) END) AS NUM
                                    FROM [TK].dbo.INVMB WITH(NOLOCK), [TK].dbo.COPTC WITH(NOLOCK), [TK].dbo.COPTD WITH(NOLOCK)
                                    LEFT JOIN [TK].dbo.INVMD ON MD001 = TD004 AND TD010 = MD002
                                    WHERE TC001 = TD001 AND TC002 = TD002 AND MB001 = TD004
                                    AND TD001 = @TD001 AND TD002 = @TD002 AND TD003 = @TD003";

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@TD001", TD001);
                        cmd.Parameters.AddWithValue("@TD002", TD002);
                        cmd.Parameters.AddWithValue("@TD003", TD003);

                        using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                        {
                            DataSet ds = new DataSet();
                            adapter.Fill(ds, "data");

                            var rows = ds.Tables["data"].Rows;
                            if (rows.Count == 0)
                            {
                                ClearTextBoxes();
                                return;
                            }

                            DataRow row = rows[0];
                            string td004 = row["TD004"].ToString();
                            string td005 = row["TD005"].ToString();
                            string td006 = row["TD006"].ToString();
                            string tc053 = row["TC053"].ToString();
                            string tc015 = row["TC015"].ToString();
                            string td013 = row["TD013"].ToString();
                            string num = row["NUM"].ToString();

                            DateTime dt = ParseDateFromString(td013);

                            if (MANU == "烘焙生產線")
                            {
                                textBox7.Text = td004;
                                textBox10.Text = td005;
                                textBox11.Text = td006;
                                textBox9.Text = tc053;
                                textBox53.Text = tc015;
                                textBox12.Text = num;
                                dateTimePicker5.Value = dt;
                            }
                            else if (MANU == "烘焙包裝線")
                            {
                                textBox4.Text = td004;
                                textBox16.Text = td005;
                                textBox17.Text = td006;
                                textBox5.Text = tc053;
                                textBox18.Text = tc015;
                                textBox15.Text = num;
                                dateTimePicker6.Value = dt;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("錯誤：" + ex.Message);
            }
        }

        private void ClearTextBoxes()
        {
            if (MANU == "烘焙生產線")
            {
                textBox7.Clear();
                textBox10.Clear();
                textBox11.Clear();
                textBox12.Clear();
                textBox53.Clear();
                textBox9.Clear();
                textBox42.Clear();
                textBox43.Clear();
                textBox72.Clear();
            }
            else if (MANU == "烘焙包裝線")
            {
                textBox4.Clear();
                textBox16.Clear();
                textBox17.Clear();
                textBox12.Clear();
                textBox18.Clear();
                textBox5.Clear();
                textBox3.Clear();
                textBox19.Clear();
                textBox20.Clear();
            }
        }

        private DateTime ParseDateFromString(string yyyymmdd)
        {
            if (yyyymmdd.Length >= 8)
            {
                string year = yyyymmdd.Substring(0, 4);
                string month = yyyymmdd.Substring(4, 2);
                string day = yyyymmdd.Substring(6, 2);
                return new DateTime(int.Parse(year), int.Parse(month), int.Parse(day));
            }
            return DateTime.Today;
        }


        public void SEARCHCOPDEFAULT2(string TD001, string TD002, string TD003)
        {
            try
            {
                Class1 TKID = new Class1();
                var sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                using (SqlConnection sqlConn = new SqlConnection(sqlsb.ConnectionString))
                {
                    string sql = @"
                                SELECT 
                                    TC053, TD004, TD005, TD006, (TD008 + TD024) AS TD008, TD010, TC015,
                                    (CASE WHEN ISNULL(INVMD.MD002, '') <> '' THEN (TD008 + TD024) * INVMD.MD004 ELSE (TD008 + TD024) END) AS NUM,
                                    BOMMD.MD003, BOMMD.MD035, BOMMD.MD036, INVMB.UDF07,
                                    ((CASE WHEN ISNULL(INVMD.MD002, '') <> '' THEN (TD008 + TD024) * INVMD.MD004 ELSE (TD008 + TD024) END)) 
                                        / BOMMC.MC004 * BOMMD.MD006 AS NUM2
                                FROM [TK].dbo.INVMB WITH(NOLOCK)
                                INNER JOIN [TK].dbo.COPTD WITH(NOLOCK) ON MB001 = TD004
                                INNER JOIN [TK].dbo.COPTC WITH(NOLOCK) ON TC001 = TD001 AND TC002 = TD002
                                LEFT JOIN [TK].dbo.INVMD WITH(NOLOCK) ON INVMD.MD001 = TD004 AND INVMD.MD002 = TD010
                                LEFT JOIN [TK].dbo.BOMMC WITH(NOLOCK) ON BOMMC.MC001 = TD004
                                LEFT JOIN [TK].dbo.BOMMD WITH(NOLOCK) ON BOMMD.MD001 = TD004
                                WHERE (BOMMD.MD003 LIKE '3%' OR BOMMD.MD003 LIKE '4%')
                                    AND TD001 = @TD001 AND TD002 = @TD002 AND TD003 = @TD003";

                    using (SqlCommand cmd = new SqlCommand(sql, sqlConn))
                    {
                        cmd.Parameters.AddWithValue("@TD001", TD001);
                        cmd.Parameters.AddWithValue("@TD002", TD002);
                        cmd.Parameters.AddWithValue("@TD003", TD003);

                        using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                        {
                            DataSet ds = new DataSet();
                            sqlConn.Open();
                            adapter.Fill(ds, "ds1");
                            sqlConn.Close();

                            if (ds.Tables["ds1"].Rows.Count == 0)
                            {
                                ClearTextBoxes2();
                                return;
                            }

                            DataRow row = ds.Tables["ds1"].Rows[0];
                            if (MANU == "烘焙生產線")
                            {
                                textBox7.Text = row["MD003"].ToString();
                                textBox10.Text = row["MD035"].ToString();
                                textBox11.Text = row["MD036"].ToString();
                                textBox12.Text = row["NUM2"].ToString();
                                textBox9.Text = row["TC053"].ToString();
                                textBox53.Text = row["TC015"].ToString();
                            }
                            else if (MANU == "烘焙包裝線")
                            {
                                textBox4.Text = row["MD003"].ToString();
                                textBox16.Text = row["MD035"].ToString();
                                textBox17.Text = row["MD036"].ToString();
                                textBox15.Text = row["NUM2"].ToString();
                                textBox5.Text = row["TC053"].ToString();
                                textBox18.Text = row["TC015"].ToString();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("錯誤: " + ex.Message);
            }
        }

        private void ClearTextBoxes2()
        {
            if (MANU == "烘焙生產線")
            {
                textBox7.Clear();
                textBox10.Clear();
                textBox11.Clear();
                textBox12.Clear();
                textBox9.Clear();
                textBox53.Clear();
                textBox42.Clear();
                textBox43.Clear();
                textBox72.Clear();
            }
            else if (MANU == "烘焙包裝線")
            {
                textBox4.Clear();
                textBox16.Clear();
                textBox17.Clear();
                textBox15.Clear();
                textBox5.Clear();
                textBox18.Clear();
                textBox3.Clear();
                textBox19.Clear();
                textBox20.Clear();
            }
        }


        public void SEARCHCOPDEFAULT3(string TD001, string TD002, string TD003)
        {
            try
            {
                Class1 TKID = new Class1();
                var sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                using (SqlConnection sqlConn = new SqlConnection(sqlsb.ConnectionString))
                {
                    string sql = @"
                                    SELECT 
                                        TC053, TD004, TD005, TD006, (TD008 + TD024) AS TD008, TD010, TC015,
                                        CASE WHEN ISNULL(INVMD.MD002, '') <> '' 
                                            THEN (TD008 + TD024) * INVMD.MD004 
                                            ELSE (TD008 + TD024)  
                                        END AS NUM,
                                        BOMMD.MD003, BOMMD.MD035, BOMMD.MD036, INVMB.UDF07,
                                        (CASE WHEN ISNULL(INVMD.MD002, '') <> '' 
                                            THEN (TD008 + TD024) * INVMD.MD004 
                                            ELSE (TD008 + TD024)  
                                        END) / BOMMC.MC004 * BOMMD.MD006 AS NUM2
                                    FROM [TK].dbo.INVMB WITH(NOLOCK)
                                    INNER JOIN [TK].dbo.COPTD WITH(NOLOCK) ON MB001 = TD004
                                    INNER JOIN [TK].dbo.COPTC WITH(NOLOCK) ON TC001 = TD001 AND TC002 = TD002
                                    LEFT JOIN [TK].dbo.INVMD WITH(NOLOCK) ON INVMD.MD001 = TD004 AND INVMD.MD002 = TD010
                                    LEFT JOIN [TK].dbo.BOMMC WITH(NOLOCK) ON BOMMC.MC001 = TD004
                                    LEFT JOIN [TK].dbo.BOMMD WITH(NOLOCK) ON BOMMD.MD001 = TD004
                                    WHERE (BOMMD.MD003 LIKE '3%' OR BOMMD.MD003 LIKE '4%')
                                      AND TD001 = @TD001 AND TD002 = @TD002 AND TD003 = @TD003";

                    using (SqlCommand cmd = new SqlCommand(sql, sqlConn))
                    {
                        cmd.Parameters.AddWithValue("@TD001", TD001);
                        cmd.Parameters.AddWithValue("@TD002", TD002);
                        cmd.Parameters.AddWithValue("@TD003", TD003);

                        using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                        {
                            DataSet ds = new DataSet();
                            sqlConn.Open();
                            adapter.Fill(ds, "ds1");
                            sqlConn.Close();

                            if (ds.Tables["ds1"].Rows.Count == 0)
                            {
                                ClearTextBoxes3();
                                return;
                            }

                            DataRow row = ds.Tables["ds1"].Rows[0];

                            if (MANU == "烘焙生產線")
                            {
                                textBox7.Text = row["MD003"].ToString();
                                textBox10.Text = row["MD035"].ToString();
                                textBox11.Text = row["MD036"].ToString();
                                textBox12.Text = row["NUM2"].ToString();
                                textBox9.Text = row["TC053"].ToString();
                                textBox53.Text = string.Empty;  // 你註解的是 null，建議空字串代替
                            }
                            else if (MANU == "烘焙包裝線")
                            {
                                textBox4.Text = row["MD003"].ToString();
                                textBox16.Text = row["MD035"].ToString();
                                textBox17.Text = row["MD036"].ToString();
                                textBox15.Text = row["NUM2"].ToString();
                                textBox5.Text = row["TC053"].ToString();
                                textBox53.Text = string.Empty;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("錯誤: " + ex.Message);
            }
        }

        private void ClearTextBoxes3()
        {
            if (MANU == "烘焙生產線")
            {
                textBox7.Clear();
                textBox10.Clear();
                textBox11.Clear();
                textBox12.Clear();
                textBox9.Clear();
                textBox53.Clear();
                textBox42.Clear();
                textBox43.Clear();
                textBox72.Clear();
            }
            else if (MANU == "烘焙包裝線")
            {
                textBox4.Clear();
                textBox16.Clear();
                textBox17.Clear();
                textBox15.Clear();
                textBox5.Clear();
                textBox18.Clear();
                textBox3.Clear();
                textBox19.Clear();
                textBox20.Clear();
            }
        }


        public string GETMAXTA002(string TA001,DateTime DT)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlConnection sqlConn = new SqlConnection();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            DataSet ds1 = new DataSet();
            string TA002;

            if (MANU.Equals("烘焙生產線"))
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
                    ds1.Clear();

                    sbSql.AppendFormat(@" 
                                        SELECT ISNULL(MAX(TA002),'00000000000') AS TA002
                                        FROM [TK].[dbo].[MOCTA]
                                        WHERE  TA001='{0}' AND TA002 LIKE '%{1}%' 
                                        ", TA001, DT.ToString("yyyyMMdd"));

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
                            TA002 = SETTA002(ds1.Tables["ds1"].Rows[0]["TA002"].ToString(), DT);
                            return TA002;

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
            else if (MANU.Equals("烘焙包裝線"))
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
                    ds1.Clear();

                    sbSql.AppendFormat(@" 
                                        SELECT ISNULL(MAX(TA002),'00000000000') AS TA002
                                        FROM [TK].[dbo].[MOCTA]
                                        WHERE  TA001='{0}' AND TA002 LIKE '%{1}%' 
                                        ", TA001, DT.ToString("yyyyMMdd"));

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
                            TA002 = SETTA002(ds1.Tables["ds1"].Rows[0]["TA002"].ToString(), DT);
                            return TA002;

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



            return null;

        }
        public string SETTA002(string TA002, DateTime DT)
        {
            if (TA002.Equals("00000000000"))
            {
                return DT.ToString("yyyyMMdd") + "001";
            }

            else
            {
                int serno = Convert.ToInt16(TA002.Substring(8, 3));
                serno = serno + 1;
                string temp = serno.ToString();
                temp = temp.PadLeft(3, '0');
                return DT.ToString("yyyyMMdd") + temp.ToString();
            }

            //if (MANU.Equals("烘焙生產線"))
            //{
                
            //}

            return null;
        }

        public void ADDMOCMANULINERESULT(string ID, string TA001, string TA002)
        {
            try
            {
                if (MANU == "烘焙生產線" || MANU == "烘焙包裝線")
                {
                    Class1 TKID = new Class1();//解密類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    using (SqlConnection sqlConn = new SqlConnection(sqlsb.ConnectionString))
                    {
                        sqlConn.Open();
                        using (SqlTransaction tran = sqlConn.BeginTransaction())
                        {
                            using (SqlCommand cmd = sqlConn.CreateCommand())
                            {
                                cmd.Transaction = tran;
                                cmd.CommandTimeout = 60;
                                cmd.CommandText = @"
                                                    INSERT INTO [TKMOC].[dbo].[MOCMANULINERESULTBAKING]
                                                    ([SID], [MOCTA001], [MOCTA002])
                                                    VALUES (@ID, @TA001, @TA002)";

                                cmd.Parameters.AddWithValue("@ID", ID);
                                cmd.Parameters.AddWithValue("@TA001", TA001);
                                cmd.Parameters.AddWithValue("@TA002", TA002);

                                int result = cmd.ExecuteNonQuery();

                                if (result == 0)
                                {
                                    tran.Rollback();
                                }
                                else
                                {
                                    tran.Commit();
                                }
                            }
                        }
                    }
                }
                else
                {
                    // MANU 不符合條件，跳出或其他處理
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("新增製令結果失敗: " + ex.Message);
            }
        }


        public void ADDMOCTATB(string TA001,string TA002,string TA020, DateTime DT,string MC001)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
           
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();          

            MOCTADATA MOCTA = new MOCTADATA();
            if (MANU.Equals("烘焙生產線"))
            {
                MOCTA = SETMOCTA(TA001, TA002, DT, MC001);
            }
            else if (MANU.Equals("烘焙包裝線"))
            {
                MOCTA = SETMOCTA(TA001, TA002, DT, MC001);
            }
            string MOCMB001 = null;
            decimal MOCTA015 = Convert.ToDecimal(MOCTA.TA015);
            string MOCTB009 = null;

            const int MaxLength = 100;

            if (MANU.Equals("烘焙生產線"))
            {
                MOCMB001 = MB001B;
                
                MOCTA.TA001 = TA001;
                MOCTA.TA002 = TA002;
                MOCTA.TA006 = MB001B;
                MOCTA.TA020 = TA020;
                MOCTA.TA026 = TA026;
                MOCTA.TA027 = TA027;
                MOCTA.TA028 = TA028;
                MOCTA.TA021 = "08";
                //MOCTB009 = textBox78.Text;

            }
            else  if (MANU.Equals("烘焙包裝線"))
            {
                MOCMB001 = MB001_DV4;

                MOCTA.TA001 = TA001;
                MOCTA.TA002 = TA002;
                MOCTA.TA006 = MB001_DV4;
                MOCTA.TA020 = TA020;
                MOCTA.TA026 = TA026_DV4;
                MOCTA.TA027 = TA027_DV4;
                MOCTA.TA028 = TA028_DV4;
                MOCTA.TA021 = "12";
                //MOCTB009 = textBox78.Text;
            }

            try
            {
                // 只處理 TA002 開頭是2 且 TA040 開頭是2 的情況
                if (MOCTA.TA002.StartsWith("2") && MOCTA.TA040.StartsWith("2"))
                {
                    Class1 TKID = new Class1();
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    // 解密帳密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    using (SqlConnection sqlConn = new SqlConnection(sqlsb.ConnectionString))
                    {
                        sqlConn.Open();

                        using (SqlTransaction tran = sqlConn.BeginTransaction())
                        using (SqlCommand cmd = sqlConn.CreateCommand())
                        {
                            cmd.Transaction = tran;
                            cmd.CommandTimeout = 60;

                            sbSql.Clear();

                            // INSERT INTO MOCTA
                            sbSql.AppendFormat(@"
                                INSERT INTO [TK].[dbo].[MOCTA]
                                ([COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE],
                                 [TRANS_NAME],[sync_count],[DataGroup],[TA001],[TA002],[TA003],[TA004],[TA005],[TA006],[TA007],
                                 [TA009],[TA010],[TA011],[TA012],[TA013],[TA014],[TA015],[TA016],[TA017],[TA018],
                                 [TA019],[TA020],[TA021],[TA022],[TA024],[TA025],[TA029],[TA030],[TA031],[TA034],[TA035],
                                 [TA040],[TA041],[TA043],[TA044],[TA045],[TA046],[TA047],[TA049],[TA050],[TA200],
                                 [TA026],[TA027],[TA028])
                                VALUES
                                ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}',
                                 '{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}','{19}',
                                 '{20}','{21}','{22}','{23}','{24}','{25}','{26}','{27}','{28}','{29}',
                                 '{30}','{31}','{32}','{33}','{34}','{35}',N'{36}','{37}','{38}','{39}','{40}',
                                 '{41}','{42}','{43}','{44}','{45}','{46}','{47}','{48}','{49}','{50}',
                                 '{51}','{52}','{53}')",
                                MOCTA.COMPANY, MOCTA.CREATOR, MOCTA.USR_GROUP, MOCTA.CREATE_DATE, MOCTA.MODIFIER, MOCTA.MODI_DATE, MOCTA.FLAG, MOCTA.CREATE_TIME, MOCTA.MODI_TIME, MOCTA.TRANS_TYPE,
                                MOCTA.TRANS_NAME, MOCTA.sync_count, MOCTA.DataGroup, MOCTA.TA001, MOCTA.TA002, MOCTA.TA003, MOCTA.TA004, MOCTA.TA005, MOCTA.TA006, MOCTA.TA007,
                                MOCTA.TA009, MOCTA.TA010, MOCTA.TA011, MOCTA.TA012, MOCTA.TA013, MOCTA.TA014, MOCTA.TA015, MOCTA.TA016, MOCTA.TA017, MOCTA.TA018,
                                MOCTA.TA019, MOCTA.TA020, MOCTA.TA021, MOCTA.TA022, MOCTA.TA024, MOCTA.TA025, MOCTA.TA029, MOCTA.TA030, MOCTA.TA031, MOCTA.TA034, MOCTA.TA035,
                                MOCTA.TA040, MOCTA.TA041, MOCTA.TA043, MOCTA.TA044, MOCTA.TA045, MOCTA.TA046, MOCTA.TA047, MOCTA.TA049, MOCTA.TA050, MOCTA.TA200,
                                MOCTA.TA026, MOCTA.TA027, MOCTA.TA028);

                            // INSERT INTO MOCTB
                            sbSql.AppendFormat(@"
                                INSERT INTO [TK].dbo.[MOCTB]
                                ([COMPANY],[CREATOR],[USR_GROUP],[CREATE_DATE],[MODIFIER],[MODI_DATE],[FLAG],[CREATE_TIME],[MODI_TIME],[TRANS_TYPE],
                                 [TRANS_NAME],[sync_count],[DataGroup],[TB001],[TB002],[TB003],[TB004],[TB005],[TB006],[TB007],
                                 [TB009],[TB011],[TB012],[TB013],[TB014],[TB018],[TB019],[TB020],[TB022],[TB024],
                                 [TB025],[TB026],[TB027],[TB029],[TB030],[TB031],[TB501],[TB554],[TB556],[TB560])
                                (SELECT 
                                 '{0}' [COMPANY],'{1}' [CREATOR],'{2}' [USR_GROUP],'{3}' [CREATE_DATE],'{4}' [MODIFIER],'{5}' [MODI_DATE],'{6}' [FLAG],'{7}' [CREATE_TIME],'{8}' [MODI_TIME],'{9}' [TRANS_TYPE],
                                 '{10}' [TRANS_NAME],{11} [sync_count],'{12}' [DataGroup],'{13}' [TB001],'{14}' [TB002],[BOMMD].MD003 [TB003],
                                 CASE 
                                    WHEN MB041=1 AND [BOMMD].MD003 NOT LIKE '201%' 
                                        THEN CONVERT(decimal(16,4),CEILING({15}/[BOMMC].MC004*[BOMMD].MD006/[BOMMD].MD007*(1+[BOMMD].MD008))) 
                                    WHEN MB041=1 AND [BOMMD].MD003 LIKE '201%' 
                                        THEN ROUND({15}/[BOMMC].MC004*[BOMMD].MD006/[BOMMD].MD007*(1+[BOMMD].MD008),0) 
                                    ELSE ROUND({15}/[BOMMC].MC004*[BOMMD].MD006/[BOMMD].MD007*(1+[BOMMD].MD008),3) 
                                 END [TB004], 
                                 0 [TB005],MD009 [TB006],[INVMB].MB004 [TB007],
                                 [INVMB].MB017 [TB009],'1' [TB011],[INVMB].MB002 [TB012],[INVMB].MB003 [TB013],[BOMMD].MD001 [TB014],'N' [TB018],'0' [TB019],'0' [TB020],'2' [TB022],'0' [TB024],
                                 '****' [TB025],'0' [TB026],'1' [TB027],'0' [TB029],'0' [TB030],'0' [TB031],'0' [TB501],'N' [TB554],'0' [TB556],'0' [TB560]
                                FROM [TK].dbo.[BOMMD], [TK].dbo.[INVMB],[TK].dbo.[BOMMC]
                                WHERE [BOMMD].MD003 = [INVMB].MB001
                                and [BOMMD].MD001 =[BOMMC].MC001
                                AND MD001 = '{16}' AND ISNULL(MD012, '') = ''
                                )",
                                MOCTA.COMPANY, MOCTA.CREATOR, MOCTA.USR_GROUP, MOCTA.CREATE_DATE, MOCTA.MODIFIER, MOCTA.MODI_DATE, MOCTA.FLAG, MOCTA.CREATE_TIME, MOCTA.MODI_TIME, MOCTA.TRANS_TYPE,
                                MOCTA.TRANS_NAME, MOCTA.sync_count, MOCTA.DataGroup, MOCTA.TA001, MOCTA.TA002, MOCTA015,
                                MOCMB001);

                            cmd.CommandText = sbSql.ToString();

                            result = cmd.ExecuteNonQuery();

                            if (result == 0)
                            {
                                tran.Rollback();
                                MessageBox.Show("新增資料失敗，交易已回滾。");
                            }
                            else
                            {
                                tran.Commit();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("執行發生錯誤: " + ex.Message);
            }
        }

        public MOCTADATA SETMOCTA(string TA001,string TA002, DateTime DT,string MC001)
        {
            string BOMVARSION="";
            string UNIT = "";
            decimal BOMBAR = 0;
            //硯微墨-烘焙倉
            string IN = "21002";

            if (MANU.Equals("烘焙生產線"))
            {
                DataTable DATATABLE=SEARCHBOMMC(MC001);
                
                if (DATATABLE != null && DATATABLE.Rows.Count>=1)
                {
                    BOMVARSION = DATATABLE.Rows[0]["MC009"].ToString();               
                    UNIT = DATATABLE.Rows[0]["MB004"].ToString();
                    BOMBAR = Convert.ToDecimal(DATATABLE.Rows[0]["MC004"].ToString());
                }

                MOCTADATA MOCTA = new MOCTADATA();
                MOCTA.COMPANY = "TK";
                MOCTA.CREATOR = "140020";
                MOCTA.USR_GROUP = "103000";
                //MOCTA.CREATE_DATE = dt2.ToString("yyyyMMdd");
                MOCTA.CREATE_DATE = DateTime.Now.ToString("yyyyMMdd");
                MOCTA.MODIFIER = "140020";
                MOCTA.MODI_DATE = DT.ToString("yyyyMMdd");
                MOCTA.FLAG = "0";
                MOCTA.CREATE_TIME = DateTime.Now.ToString("HH:mm:dd");
                MOCTA.MODI_TIME = DateTime.Now.ToString("HH:mm:dd");
                MOCTA.TRANS_TYPE = "P001";
                MOCTA.TRANS_NAME = "MOCMI02";
                MOCTA.sync_count = "0";
                MOCTA.DataGroup = "103000";
                MOCTA.TA001 = TA001;
                MOCTA.TA002 = TA002;
                MOCTA.TA003 = DT.ToString("yyyyMMdd");
                MOCTA.TA004 = DT.ToString("yyyyMMdd");
                MOCTA.TA005 = BOMVARSION;
                MOCTA.TA006 = MB001B;
                MOCTA.TA007 = UNIT;
                MOCTA.TA009 = DT.ToString("yyyyMMdd");
                MOCTA.TA010 = DT.ToString("yyyyMMdd");
                MOCTA.TA011 = "1";
                MOCTA.TA012 = DT.ToString("yyyyMMdd");
                MOCTA.TA013 = "N";
                // MOCTA.TA014 = dt2.ToString("yyyyMMdd");
                MOCTA.TA014 = "";
                //MOCTA.TA015 = (BOX * BOMBAR).ToString();
                MOCTA.TA015 = SUBPACKAGE2.ToString();
                MOCTA.TA016 = "0";
                MOCTA.TA017 = "0";
                MOCTA.TA018 = "0";
                MOCTA.TA019 = "20";
                MOCTA.TA020 = IN;
                MOCTA.TA021 = "08";
                MOCTA.TA022 = "0";
                MOCTA.TA024 = TA001;
                MOCTA.TA025 = TA002;
                MOCTA.TA029 = TA029;
                MOCTA.TA030 = "1";
                MOCTA.TA031 = "0";
                MOCTA.TA034 = MB002B;
                MOCTA.TA035 = MB003B;
                MOCTA.TA040 = DT.ToString("yyyyMMdd");
                MOCTA.TA041 = "";
                MOCTA.TA043 = "1";
                MOCTA.TA044 = "N";
                MOCTA.TA045 = "0";
                MOCTA.TA046 = "0";
                MOCTA.TA047 = "0";
                MOCTA.TA049 = "0";
                MOCTA.TA050 = "0";
                MOCTA.TA200 = "1";


                return MOCTA;
            }
            else if (MANU.Equals("烘焙包裝線"))
            {
                DataTable DATATABLE = SEARCHBOMMC(MC001);

                if (DATATABLE != null && DATATABLE.Rows.Count >= 1)
                {
                    BOMVARSION = DATATABLE.Rows[0]["MC009"].ToString();
                    UNIT = DATATABLE.Rows[0]["MB004"].ToString();
                    BOMBAR = Convert.ToDecimal(DATATABLE.Rows[0]["MC004"].ToString());
                }

                MOCTADATA MOCTA = new MOCTADATA();
                MOCTA.COMPANY = "TK";  
                MOCTA.CREATOR = "140020";
                MOCTA.USR_GROUP = "103000";
                //MOCTA.CREATE_DATE = dt2.ToString("yyyyMMdd");
                MOCTA.CREATE_DATE = DateTime.Now.ToString("yyyyMMdd");
                MOCTA.MODIFIER = "140020";
                MOCTA.MODI_DATE = DT.ToString("yyyyMMdd");
                MOCTA.FLAG = "0";
                MOCTA.CREATE_TIME = DateTime.Now.ToString("HH:mm:dd");
                MOCTA.MODI_TIME = DateTime.Now.ToString("HH:mm:dd");
                MOCTA.TRANS_TYPE = "P001";
                MOCTA.TRANS_NAME = "MOCMI02";
                MOCTA.sync_count = "0";
                MOCTA.DataGroup = "103000";
                MOCTA.TA001 = TA001;
                MOCTA.TA002 = TA002; 
                MOCTA.TA003 = DT.ToString("yyyyMMdd");
                MOCTA.TA004 = DT.ToString("yyyyMMdd");
                MOCTA.TA005 = BOMVARSION;
                MOCTA.TA006 = MB001_DV4;
                MOCTA.TA007 = UNIT;
                MOCTA.TA009 = DT.ToString("yyyyMMdd");
                MOCTA.TA010 = DT.ToString("yyyyMMdd");
                MOCTA.TA011 = "1";
                MOCTA.TA012 = DT.ToString("yyyyMMdd");
                MOCTA.TA013 = "N";
                // MOCTA.TA014 = dt2.ToString("yyyyMMdd");
                MOCTA.TA014 = "";
                //MOCTA.TA015 = (BOX * BOMBAR).ToString();
                MOCTA.TA015 = SUBPACKAGE_DV4.ToString();
                MOCTA.TA016 = "0";
                MOCTA.TA017 = "0";
                MOCTA.TA018 = "0";
                MOCTA.TA019 = "20";
                MOCTA.TA020 = IN;
                MOCTA.TA021 = "08";
                MOCTA.TA022 = "0";
                MOCTA.TA024 = TA001;
                MOCTA.TA025 = TA002;
                MOCTA.TA029 = TA029_DV4;
                MOCTA.TA030 = "1";
                MOCTA.TA031 = "0";
                MOCTA.TA034 = MB002_DV4;
                MOCTA.TA035 = MB003_DV4;
                MOCTA.TA040 = DT.ToString("yyyyMMdd");
                MOCTA.TA041 = "";
                MOCTA.TA043 = "1";
                MOCTA.TA044 = "N";
                MOCTA.TA045 = "0";
                MOCTA.TA046 = "0";
                MOCTA.TA047 = "0";
                MOCTA.TA049 = "0";
                MOCTA.TA050 = "0";
                MOCTA.TA200 = "1";


                return MOCTA;
            }



            return null;

        }

        public DataTable SEARCHBOMMC(string MC001)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlConnection sqlConn = new SqlConnection();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            DataSet ds1 = new DataSet();

            string BOMVARSION = null;
            string UNIT = null;
            decimal BOMBAR = 0;
            
            if (MANU.Equals("烘焙生產線"))
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
                                        SELECT 
                                        [MC001],[MC002],[MC003],[MC004],[MC005],[MC006],[MC007],[MC008],[MC009],[MC010],
                                        [MC011],[MC012],[MC013],[MC014],[MC015],[MC016],[MC017],[MC018],[MC019],[MC020],
                                        [MC021],[MC022],[MC023],[MC024],[MC025],[MC026],[MC027],
                                        INVMB.MB004
                                        FROM [TK].[dbo].[BOMMC]
                                        LEFT JOIN [TK].dbo.[INVMB] ON MB001=MC001
                                        WHERE [MC001]='{0}'", MC001);

                    sbSql.AppendFormat(@"  ");

                    adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                    sqlConn.Open();
                    ds1.Clear();
                    adapter1.Fill(ds1, "ds1");
                    sqlConn.Close();

                    if (ds1 != null && ds1.Tables["ds1"].Rows.Count >= 1)
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
            else if (MANU.Equals("烘焙包裝線"))
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
                                        SELECT 
                                        [MC001],[MC002],[MC003],[MC004],[MC005],[MC006],[MC007],[MC008],[MC009],[MC010],
                                        [MC011],[MC012],[MC013],[MC014],[MC015],[MC016],[MC017],[MC018],[MC019],[MC020],
                                        [MC021],[MC022],[MC023],[MC024],[MC025],[MC026],[MC027],
                                        INVMB.MB004
                                        FROM [TK].[dbo].[BOMMC]
                                        LEFT JOIN [TK].dbo.[INVMB] ON MB001=MC001
                                        WHERE [MC001]='{0}'", MC001);

                    sbSql.AppendFormat(@"  ");

                    adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                    sqlConn.Open();
                    ds1.Clear();
                    adapter1.Fill(ds1, "ds1");
                    sqlConn.Close();

                    if (ds1 != null && ds1.Tables["ds1"].Rows.Count >= 1)
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
            else
            {
                return null;
            }

        }

        public void SEARCH_MOCMANULINERESULTBAKING(string ID)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlConnection sqlConn = new SqlConnection();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            DataSet ds1 = new DataSet();

            if (MANU.Equals("烘焙生產線"))
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
                                        SELECT  [MOCTA001] AS '製令',[MOCTA002]  AS '單號',[SID]
                                        FROM [TKMOC].[dbo].[MOCMANULINERESULTBAKING]
                                        WHERE [SID]='{0}'"
                                        , ID);
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");

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

                            dataGridView3.DataSource = ds1.Tables["ds1"];
                            dataGridView3.AutoResizeColumns();
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
        }
      
        public void SEARCH_MOCMANULINERESULTBAKING_DV4(string ID)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlConnection sqlConn = new SqlConnection();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            DataSet ds1 = new DataSet();

            if (MANU.Equals("烘焙包裝線"))
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
                                        SELECT  [MOCTA001] AS '製令',[MOCTA002]  AS '單號',[SID]
                                        FROM [TKMOC].[dbo].[MOCMANULINERESULTBAKING]
                                        WHERE [SID]='{0}'"
                                        , ID);
                    sbSql.AppendFormat(@"  ");
                    sbSql.AppendFormat(@"  ");

                    adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                    sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                    sqlConn.Open();
                    ds1.Clear();
                    adapter1.Fill(ds1, "ds1");
                    sqlConn.Close();

                    dataGridView6.DataSource = null;

                    if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {

                        dataGridView6.DataSource = ds1.Tables["ds1"];
                        dataGridView6.AutoResizeColumns();
                    }

                }
                catch
                {

                }
                finally
                {

                }
            }
        }
        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView3.CurrentRow != null)
            {
                int rowindex = dataGridView3.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView3.Rows[rowindex];
                    DELID = row.Cells["SID"].Value.ToString();
                    DELMOCTA001B = row.Cells["製令"].Value.ToString();
                    DELMOCTA002B = row.Cells["單號"].Value.ToString();



                }
                else
                {
                    DELID = null;

                }
            }
        }

        public void DELTE_MOCMANULINERESULTBAKING(string DELID, string DEL_TA001, string DEL_TA002)
        {
            if ((MANU.Equals("烘焙生產線") || MANU.Equals("烘焙包裝線")))
            {
                try
                {
                    // 解密連線字串
                    Class1 TKID = new Class1();
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    using (SqlConnection sqlConn = new SqlConnection(sqlsb.ConnectionString))
                    {
                        sqlConn.Open();
                        using (SqlTransaction tran = sqlConn.BeginTransaction())
                        {
                            try
                            {
                                string sql = @"
                                            DELETE FROM [TKMOC].[dbo].[MOCMANULINERESULTBAKING]
                                            WHERE SID = @SID AND MOCTA001 = @TA001 AND MOCTA002 = @TA002";

                                using (SqlCommand cmd = new SqlCommand(sql, sqlConn, tran))
                                {
                                    cmd.CommandTimeout = 60;
                                    cmd.Parameters.AddWithValue("@SID", DELID);
                                    cmd.Parameters.AddWithValue("@TA001", DEL_TA001);
                                    cmd.Parameters.AddWithValue("@TA002", DEL_TA002);

                                    int result = cmd.ExecuteNonQuery();

                                    if (result > 0)
                                    {
                                        tran.Commit();
                                    }
                                    else
                                    {
                                        tran.Rollback();
                                    }
                                }
                            }
                            catch
                            {
                                tran.Rollback();
                                throw; // 建議拋出例外或記錄錯誤
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"刪除失敗：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }            
        }


        private void textBox12_TextChanged(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(textBox12.Text.ToString()) && !string.IsNullOrEmpty(textBox33.Text.ToString()))
            {
                CAL_BAR(textBox12.Text.ToString(), textBox33.Text.ToString());
            }
           
        }

        private void textBox33_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox12.Text.ToString()) && !string.IsNullOrEmpty(textBox33.Text.ToString()))
            {
                CAL_BAR(textBox12.Text.ToString(), textBox33.Text.ToString());
            }
        }
        public void CAL_BAR(string NUMS,string BOMS)
        {
            decimal COUNT_NUMS = Convert.ToDecimal(NUMS);
            decimal COUNT_BOMS = Convert.ToDecimal(BOMS);
            if(COUNT_NUMS>0 & COUNT_BOMS>0)
            {
                textBox1.Text = Math.Round(COUNT_NUMS / COUNT_BOMS).ToString();
            }

            
        }
        public void CAL_BAR2(string NUMS, string BOMS)
        {
            decimal COUNT_NUMS = Convert.ToDecimal(NUMS);
            decimal COUNT_BOMS = Convert.ToDecimal(BOMS);
            if (COUNT_NUMS > 0 & COUNT_BOMS > 0)
            {
                textBox21.Text = Math.Round(COUNT_NUMS / COUNT_BOMS).ToString();
            }


        }

        public void ADD_MOCMANULINEBAKING_BATCH(string KINDS,string MANU,string MANUDATE, string TD001, string TD002, string TD003)
        {
            SqlConnection sqlConn = new SqlConnection();

            if (KINDS.Equals("成品"))
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


                    sbSql.AppendFormat(@"  
                                        INSERT INTO [TKMOC].[dbo].[MOCMANULINEBAKING]
                                        ([ID],[MANU],[MANUDATE],[MB001],[MB002],[MB003],[BAR],[NUM],[CLINET],[MANUHOUR],[BOX],[PACKAGE],[OUTDATE],[TA029],[HALFPRO],[COPTD001],[COPTD002],[COPTD003])
                                        SELECT ID,'{0}','{1}',TD004 [MB001],TD005 [MB002],TD006 [MB003],0 [BAR],NUM [NUM],TC053 [CLINET],0 [MANUHOUR],(NUM/MD007) [BOX],NUM [PACKAGE],TD013 [OUTDATE],TC015 [TA029],0 [HALFPRO],TD001 [COPTD001] ,TD002 [TCOPTD002], TD003 [TCOPTD003]
                                        FROM 
                                        (
                                        SELECT NEWID() AS ID,TD001,TD002,TD003,TC053,TD004,TD005,TD006,(TD008+TD024) AS TD008,TD010,TC015,TD013,(CASE WHEN ISNULL(MD002,'')<>'' THEN (TD008+TD024)*MD004 ELSE (TD008+TD024)  END ) AS NUM
                                        ,ISNULL((SELECT TOP 1 MD007 FROM [TK].dbo.BOMMD MD WHERE (MD.MD003 LIKE '201%') AND MD.MD001=TD004),1) AS MD007
                                        ,ISNULL((SELECT TOP 1 MC004 FROM [TK].dbo.BOMMC MC WHERE MC.MC001=TD004),1) AS MC004
                                        FROM [TK].dbo.INVMB WITH(NOLOCK),[TK].dbo.COPTC WITH(NOLOCK),[TK].dbo.COPTD WITH(NOLOCK)
                                        LEFT JOIN [TK].dbo.INVMD ON MD001=TD004 AND TD010=MD002
                                        WHERE TC001=TD001 AND TC002=TD002
                                        AND MB001=TD004
                                        AND TD001='{2}' AND TD002='{3}' AND TD003='{4}'
                                        ) AS TEMP 
                                        ", MANU, MANUDATE, TD001, TD002, TD003);

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

                        MessageBox.Show("成功");
                    }

                }
                catch (Exception ex)
                {
                }
                finally
                {
                    sqlConn.Close();
                }
            }
            else if (KINDS.Equals("第一層"))
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


                    sbSql.AppendFormat(@"  
                                        INSERT INTO [TKMOC].[dbo].[MOCMANULINEBAKING]
                                        ([ID],[MANU],[MANUDATE],[MB001],[MB002],[MB003],[BAR],[NUM],[CLINET],[MANUHOUR],[BOX],[PACKAGE],[OUTDATE],[TA029],[HALFPRO],[COPTD001],[COPTD002],[COPTD003])
                                        SELECT ID,'{0}','{1}',MD003 [MB001],MD035 [MB002],MD036 [MB003],(CASE WHEN MD003 LIKE '4%' THEN 0 ELSE CONVERT(DECIMAL(16,4),(BOMNUMS/MC004))  END ) [BAR],BOMNUMS [NUM],TC053 [CLINET],0 [MANUHOUR],(CASE WHEN MD003 LIKE '4%' THEN CONVERT(DECIMAL(16,4),(BOMNUMS/MD007B)) ELSE 0  END) [BOX],BOMNUMS [PACKAGE],TD013 [OUTDATE],TC015 [TA029],0 [HALFPRO],TD001 [COPTD001] ,TD002 [TCOPTD002], TD003 [TCOPTD003]
                                        FROM 
                                        (
                                        SELECT  NEWID() AS ID,TD001,TD002,TD003,TC053,TD004,TD005,TD006,(TD008+TD024) AS TD008,TD010,TC015,TD013,(CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024)  END ) AS NUM
                                        ,BOMMD.MD003,BOMMD.MD035,BOMMD.MD036,BOMMD.MD006,BOMMD.MD007
                                        ,((CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024)  END )*BOMMD.MD006/BOMMD.MD007/BOMMC.MC004) AS BOMNUMS
                                        ,ISNULL((SELECT TOP 1 MD007 FROM [TK].dbo.BOMMD MD WHERE (MD.MD003 LIKE '201%') AND MD.MD001=BOMMD.MD003),1) AS MD007B
                                        ,ISNULL((SELECT TOP 1 MC004 FROM [TK].dbo.BOMMC MC WHERE MC.MC001=BOMMD.MD003),1) AS MC004
                                        FROM [TK].dbo.INVMB WITH(NOLOCK),[TK].dbo.COPTC WITH(NOLOCK),[TK].dbo.COPTD WITH(NOLOCK)
                                        LEFT JOIN [TK].dbo.INVMD ON INVMD.MD001=TD004 AND TD010=MD002
                                        LEFT JOIN [TK].dbo.BOMMC ON BOMMC.MC001=TD004
                                        LEFT JOIN [TK].dbo.BOMMD ON BOMMD.MD001=TD004
                                        WHERE TC001=TD001 AND TC002=TD002
                                        AND MB001=TD004
                                        AND TD001='{2}' AND TD002='{3}' AND TD003='{4}'
                                        AND (BOMMD.MD003 LIKE '3%' OR BOMMD.MD003 LIKE '4%')
                                        ) AS TEMP
                                        ", MANU, MANUDATE, TD001, TD002, TD003);

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

                        MessageBox.Show("成功");
                    }

                }
                catch (Exception ex)
                {
                }
                finally
                {
                    sqlConn.Close();
                }
            }
            else if (KINDS.Equals("第二層"))
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


                    sbSql.AppendFormat(@"  
                                        INSERT INTO [TKMOC].[dbo].[MOCMANULINEBAKING]
                                        ([ID],[MANU],[MANUDATE],[MB001],[MB002],[MB003],[BAR],[NUM],[CLINET],[MANUHOUR],[BOX],[PACKAGE],[OUTDATE],[TA029],[HALFPRO],[COPTD001],[COPTD002],[COPTD003])
                                        SELECT ID,'{0}','{1}',MD003B [MB001],MD035B [MB002],MD036B [MB003],CONVERT(DECIMAL(16,4),(BOMNUMS2/MC004C)) [BAR],BOMNUMS2 [NUM],TC053 [CLINET],0 [MANUHOUR],0 [BOX],BOMNUMS2 [PACKAGE],TD013 [OUTDATE],TC015 [TA029],0 [HALFPRO],TD001 [COPTD001] ,TD002 [TCOPTD002], TD003 [TCOPTD003]
                                        FROM (
                                        SELECT NEWID() AS ID,TD001,TD002,TD003,TC053,TD004,TD005,TD006,(TD008+TD024) AS TD008,TD010,TC015,TD013,(CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024)  END ) AS NUM
                                        ,BOMMD.MD003,BOMMD.MD035,BOMMD.MD036,BOMMD.MD006,BOMMD.MD007
                                        ,((CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024)  END )*BOMMD.MD006/BOMMD.MD007/BOMMC.MC004) AS BOMNUMS
                                        ,BOMMD2.MD003 MD003B,BOMMD2.MD035 MD035B,BOMMD2.MD036 MD036B,BOMMD2.MD006 MD006B,BOMMD2.MD007 MD007B
                                        ,(((CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024)  END )*BOMMD.MD006/BOMMD.MD007/BOMMC.MC004)*BOMMD2.MD006/BOMMD2.MD007/BOMMC2.MC004)AS BOMNUMS2
                                        ,ISNULL((SELECT TOP 1 MD007 FROM [TK].dbo.BOMMD MD WHERE (MD.MD003 LIKE '201%') AND MD.MD001=BOMMD2.MD003),1) AS MD007C
                                        ,ISNULL((SELECT TOP 1 MC004 FROM [TK].dbo.BOMMC MC WHERE MC.MC001=BOMMD2.MD003),1) AS MC004C
                                        FROM [TK].dbo.INVMB WITH(NOLOCK),[TK].dbo.COPTC WITH(NOLOCK),[TK].dbo.COPTD WITH(NOLOCK)
                                        LEFT JOIN [TK].dbo.INVMD ON INVMD.MD001=TD004 AND TD010=MD002
                                        LEFT JOIN [TK].dbo.BOMMC ON BOMMC.MC001=TD004
                                        LEFT JOIN [TK].dbo.BOMMD ON BOMMD.MD001=TD004
                                        LEFT JOIN [TK].dbo.BOMMC BOMMC2 ON BOMMC2.MC001=BOMMD.MD003
                                        LEFT JOIN [TK].dbo.BOMMD BOMMD2 ON BOMMD2.MD001=BOMMD.MD003
                                        WHERE TC001=TD001 AND TC002=TD002
                                        AND MB001=TD004
                                        AND TD001='{2}' AND TD002='{3}' AND TD003='{4}'
                                        AND (BOMMD.MD003 LIKE '3%' OR BOMMD.MD003 LIKE '4%')
                                        AND (BOMMD2.MD003 LIKE '3%' OR BOMMD2.MD003 LIKE '4%')
                                        ) AS TEMP
                                        ", MANU, MANUDATE, TD001, TD002, TD003);

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

                        MessageBox.Show("成功");
                    }

                }
                catch (Exception ex)
                {
                }
                finally
                {
                    sqlConn.Close();
                }
            }
            else if (KINDS.Equals("第三層"))
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


                    sbSql.AppendFormat(@"  
                                        INSERT INTO [TKMOC].[dbo].[MOCMANULINEBAKING]
                                        ([ID],[MANU],[MANUDATE],[MB001],[MB002],[MB003],[BAR],[NUM],[CLINET],[MANUHOUR],[BOX],[PACKAGE],[OUTDATE],[TA029],[HALFPRO],[COPTD001],[COPTD002],[COPTD003])
                                        
                                        SELECT ID,'{0}','{1}',MD003C [MB001],MD035C [MB002],MD036C [MB003],CONVERT(DECIMAL(16,4),(BOMNUMS3/MC004MD003C)) [BAR],BOMNUMS3 [NUM],TC053 [CLINET],0 [MANUHOUR],0 [BOX],BOMNUMS3 [PACKAGE],TD013 [OUTDATE],TC015 [TA029],0 [HALFPRO],TD001 [COPTD001] ,TD002 [TCOPTD002], TD003 [TCOPTD003]
                                        FROM (
                                        SELECT NEWID() AS ID,TD001,TD002,TD003,TC053,TD004,TD005,TD006,(TD008+TD024) AS TD008,TD010,TC015,TD013,(CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024)  END ) AS NUM
                                        ,BOMMD.MD003,BOMMD.MD035,BOMMD.MD036,BOMMD.MD006,BOMMD.MD007
                                        ,((CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024)  END )*BOMMD.MD006/BOMMD.MD007/BOMMC.MC004) AS BOMNUMS

                                        ,BOMMD2.MD003 MD003B,BOMMD2.MD035 MD035B,BOMMD2.MD036 MD036B,BOMMD2.MD006 MD006B,BOMMD2.MD007 MD007B
                                        ,(((CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024)  END )*BOMMD.MD006/BOMMD.MD007/BOMMC.MC004)*BOMMD2.MD006/BOMMD2.MD007/BOMMC2.MC004)AS BOMNUMS2
                                        ,ISNULL((SELECT TOP 1 MD007 FROM [TK].dbo.BOMMD MD WHERE (MD.MD003 LIKE '201%') AND MD.MD001=BOMMD2.MD003),1) AS MD007C
                                        ,ISNULL((SELECT TOP 1 MC004 FROM [TK].dbo.BOMMC MC WHERE MC.MC001=BOMMD2.MD003),1) AS MC004C

                                        ,BOMMD3.MD003 MD003C,BOMMD3.MD035 MD035C,BOMMD3.MD036 MD036C,BOMMD3.MD006 MD006C,BOMMD3.MD007 MD007C3
                                        ,((((CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024)  END )*BOMMD.MD006/BOMMD.MD007/BOMMC.MC004)*BOMMD2.MD006/BOMMD2.MD007/BOMMC2.MC004)*BOMMD3.MD006/BOMMD3.MD007/BOMMC3.MC004)AS BOMNUMS3
                                        ,ISNULL((SELECT TOP 1 MD007 FROM [TK].dbo.BOMMD MD WHERE (MD.MD003 LIKE '201%') AND MD.MD001=BOMMD3.MD003),1) AS MD007BOXC
                                        ,ISNULL((SELECT TOP 1 MC004 FROM [TK].dbo.BOMMC MC WHERE MC.MC001=BOMMD3.MD003),1) AS MC004MD003C

                                        FROM [TK].dbo.INVMB WITH(NOLOCK),[TK].dbo.COPTC WITH(NOLOCK),[TK].dbo.COPTD WITH(NOLOCK)
                                        LEFT JOIN [TK].dbo.INVMD ON INVMD.MD001=TD004 AND TD010=MD002
                                        LEFT JOIN [TK].dbo.BOMMC ON BOMMC.MC001=TD004
                                        LEFT JOIN [TK].dbo.BOMMD ON BOMMD.MD001=TD004
                                        LEFT JOIN [TK].dbo.BOMMC BOMMC2 ON BOMMC2.MC001=BOMMD.MD003
                                        LEFT JOIN [TK].dbo.BOMMD BOMMD2 ON BOMMD2.MD001=BOMMD.MD003
                                        LEFT JOIN [TK].dbo.BOMMC BOMMC3 ON BOMMC3.MC001=BOMMD2.MD003
                                        LEFT JOIN [TK].dbo.BOMMD BOMMD3 ON BOMMD3.MD001=BOMMD2.MD003
                                        WHERE TC001=TD001 AND TC002=TD002
                                        AND MB001=TD004
                                        AND TD001='{2}' AND TD002='{3}' AND TD003='{4}'
                                        AND (BOMMD.MD003 LIKE '3%' OR BOMMD.MD003 LIKE '4%')
                                        AND (BOMMD2.MD003 LIKE '3%' OR BOMMD2.MD003 LIKE '4%')
                                        AND (BOMMD3.MD003 LIKE '3%' OR BOMMD3.MD003 LIKE '4%')
                                        ) AS TEMP
                                        ", MANU, MANUDATE, TD001, TD002, TD003);

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

                        MessageBox.Show("成功");
                    }

                }
                catch (Exception ex)
                {
                }
                finally
                {
                    sqlConn.Close();
                }
            }
            else if (KINDS.Equals("第四層"))
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


                    sbSql.AppendFormat(@"  
                                        INSERT INTO [TKMOC].[dbo].[MOCMANULINEBAKING]
                                        ([ID],[MANU],[MANUDATE],[MB001],[MB002],[MB003],[BAR],[NUM],[CLINET],[MANUHOUR],[BOX],[PACKAGE],[OUTDATE],[TA029],[HALFPRO],[COPTD001],[COPTD002],[COPTD003])
                                       
                                        SELECT ID,'{0}','{1}',MD003D [MB001],MD035D [MB002],MD036D [MB003],CONVERT(DECIMAL(16,4),(BOMNUMS4/MC004E)) [BAR],BOMNUMS4 [NUM],TC053 [CLINET],0 [MANUHOUR],0 [BOX],BOMNUMS4 [PACKAGE],TD013 [OUTDATE],TC015 [TA029],0 [HALFPRO],TD001 [COPTD001], TD002 [COPTD002], TD003 [COPTD003]
                                        FROM (
                                            SELECT NEWID() AS ID, TD001, TD002, TD003, TC053, TD004, TD005, TD006, (TD008+TD024) AS TD008, TD010, TC015, TD013,
                                                   (CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024) END) AS NUM,
                                                   BOMMD.MD003, BOMMD.MD035, BOMMD.MD036, BOMMD.MD006, BOMMD.MD007,
                                                   ((CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024) END)*BOMMD.MD006/BOMMD.MD007/BOMMC.MC004) AS BOMNUMS,
                                                   BOMMD2.MD003 MD003B, BOMMD2.MD035 MD035B, BOMMD2.MD036 MD036B, BOMMD2.MD006 MD006B, BOMMD2.MD007 MD007B,
                                                   (((CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024) END)*BOMMD.MD006/BOMMD.MD007/BOMMC.MC004)*BOMMD2.MD006/BOMMD2.MD007/BOMMC2.MC004) AS BOMNUMS2,
                                                   BOMMD3.MD003 MD003C, BOMMD3.MD035 MD035C, BOMMD3.MD036 MD036C, BOMMD3.MD006 MD006C, BOMMD3.MD007 MD007C,
                                                   ((((CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024) END)*BOMMD.MD006/BOMMD.MD007/BOMMC.MC004)*BOMMD2.MD006/BOMMD2.MD007/BOMMC2.MC004)*BOMMD3.MD006/BOMMD3.MD007/BOMMC3.MC004) AS BOMNUMS3,
                                                   BOMMD4.MD003 MD003D, BOMMD4.MD035 MD035D, BOMMD4.MD036 MD036D, BOMMD4.MD006 MD006D, BOMMD4.MD007 MD007D,
                                                   (((((CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024) END)*BOMMD.MD006/BOMMD.MD007/BOMMC.MC004)*BOMMD2.MD006/BOMMD2.MD007/BOMMC2.MC004)*BOMMD3.MD006/BOMMD3.MD007/BOMMC3.MC004)*BOMMD4.MD006/BOMMD4.MD007/BOMMC4.MC004) AS BOMNUMS4,
                                                   ISNULL((SELECT TOP 1 MD007 FROM [TK].dbo.BOMMD MD WHERE MD.MD003 LIKE '201%' AND MD.MD001=BOMMD4.MD003),1) AS MD007E,
                                                   ISNULL((SELECT TOP 1 MC004 FROM [TK].dbo.BOMMC MC WHERE MC.MC001=BOMMD4.MD003),1) AS MC004E
                                            FROM [TK].dbo.INVMB WITH(NOLOCK), [TK].dbo.COPTC WITH(NOLOCK), [TK].dbo.COPTD WITH(NOLOCK)
                                            LEFT JOIN [TK].dbo.INVMD ON INVMD.MD001=TD004 AND TD010=MD002
                                            LEFT JOIN [TK].dbo.BOMMC ON BOMMC.MC001=TD004
                                            LEFT JOIN [TK].dbo.BOMMD ON BOMMD.MD001=TD004
                                            LEFT JOIN [TK].dbo.BOMMC BOMMC2 ON BOMMC2.MC001=BOMMD.MD003
                                            LEFT JOIN [TK].dbo.BOMMD BOMMD2 ON BOMMD2.MD001=BOMMD.MD003
                                            LEFT JOIN [TK].dbo.BOMMC BOMMC3 ON BOMMC3.MC001=BOMMD2.MD003
                                            LEFT JOIN [TK].dbo.BOMMD BOMMD3 ON BOMMD3.MD001=BOMMD2.MD003
                                            LEFT JOIN [TK].dbo.BOMMC BOMMC4 ON BOMMC4.MC001=BOMMD3.MD003
                                            LEFT JOIN [TK].dbo.BOMMD BOMMD4 ON BOMMD4.MD001=BOMMD3.MD003
                                            WHERE TC001=TD001 AND TC002=TD002
                                              AND MB001=TD004
                                              AND TD001='{2}' AND TD002='{3}' AND TD003='{4}'
                                              AND (BOMMD.MD003 LIKE '3%' OR BOMMD.MD003 LIKE '4%')
                                              AND (BOMMD2.MD003 LIKE '3%' OR BOMMD2.MD003 LIKE '4%')
                                              AND (BOMMD3.MD003 LIKE '3%' OR BOMMD3.MD003 LIKE '4%')
                                              AND (BOMMD4.MD003 LIKE '3%' OR BOMMD4.MD003 LIKE '4%')
                                        ) AS TEMP;

                                        ", MANU, MANUDATE, TD001, TD002, TD003);

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

                        MessageBox.Show("成功");
                    }

                }
                catch (Exception ex)
                {
                }
                finally
                {
                    sqlConn.Close();
                }
            }
            else if (KINDS.Equals("第五層"))
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


                    sbSql.AppendFormat(@"  
                                        INSERT INTO [TKMOC].[dbo].[MOCMANULINEBAKING]
                                        ([ID],[MANU],[MANUDATE],[MB001],[MB002],[MB003],[BAR],[NUM],[CLINET],[MANUHOUR],[BOX],[PACKAGE],[OUTDATE],[TA029],[HALFPRO],[COPTD001],[COPTD002],[COPTD003])
                                       
                                        SELECT ID,'{0}','{1}',MD003E [MB001],MD035E [MB002],MD036E [MB003],CONVERT(DECIMAL(16,4),(BOMNUMS5/MC004F)) [BAR],BOMNUMS5 [NUM],TC053 [CLINET],0 [MANUHOUR],0 [BOX],BOMNUMS5 [PACKAGE],TD013 [OUTDATE],TC015 [TA029],0 [HALFPRO],TD001 [COPTD001], TD002 [COPTD002], TD003 [COPTD003]
                                        FROM (
                                            SELECT NEWID() AS ID, TD001, TD002, TD003, TC053, TD004, TD005, TD006, (TD008+TD024) AS TD008, TD010, TC015, TD013,
                                                   (CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024) END) AS NUM,
                                                   BOMMD.MD003, BOMMD.MD035, BOMMD.MD036, BOMMD.MD006, BOMMD.MD007,
                                                   ((CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024) END)*BOMMD.MD006/BOMMD.MD007/BOMMC.MC004) AS BOMNUMS,
                                                   BOMMD2.MD003 MD003B, BOMMD2.MD035 MD035B, BOMMD2.MD036 MD036B, BOMMD2.MD006 MD006B, BOMMD2.MD007 MD007B,
                                                   (((CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024) END)*BOMMD.MD006/BOMMD.MD007/BOMMC.MC004)*BOMMD2.MD006/BOMMD2.MD007/BOMMC2.MC004) AS BOMNUMS2,
                                                   BOMMD3.MD003 MD003C, BOMMD3.MD035 MD035C, BOMMD3.MD036 MD036C, BOMMD3.MD006 MD006C, BOMMD3.MD007 MD007C,
                                                   ((((CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024) END)*BOMMD.MD006/BOMMD.MD007/BOMMC.MC004)*BOMMD2.MD006/BOMMD2.MD007/BOMMC2.MC004)*BOMMD3.MD006/BOMMD3.MD007/BOMMC3.MC004) AS BOMNUMS3,
                                                   BOMMD4.MD003 MD003D, BOMMD4.MD035 MD035D, BOMMD4.MD036 MD036D, BOMMD4.MD006 MD006D, BOMMD4.MD007 MD007D,
                                                   (((((CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024) END)*BOMMD.MD006/BOMMD.MD007/BOMMC.MC004)*BOMMD2.MD006/BOMMD2.MD007/BOMMC2.MC004)*BOMMD3.MD006/BOMMD3.MD007/BOMMC3.MC004)*BOMMD4.MD006/BOMMD4.MD007/BOMMC4.MC004) AS BOMNUMS4,
                                                   BOMMD5.MD003 MD003E, BOMMD5.MD035 MD035E, BOMMD5.MD036 MD036E, BOMMD5.MD006 MD006E, BOMMD5.MD007 MD007E,
                                                   ((((((CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024) END)*BOMMD.MD006/BOMMD.MD007/BOMMC.MC004)*BOMMD2.MD006/BOMMD2.MD007/BOMMC2.MC004)*BOMMD3.MD006/BOMMD3.MD007/BOMMC3.MC004)*BOMMD4.MD006/BOMMD4.MD007/BOMMC4.MC004)*BOMMD5.MD006/BOMMD5.MD007/BOMMC5.MC004) AS BOMNUMS5,
                                                   ISNULL((SELECT TOP 1 MD007 FROM [TK].dbo.BOMMD MD WHERE MD.MD003 LIKE '201%' AND MD.MD001=BOMMD5.MD003),1) AS MD007F,
                                                   ISNULL((SELECT TOP 1 MC004 FROM [TK].dbo.BOMMC MC WHERE MC.MC001=BOMMD5.MD003),1) AS MC004F
                                            FROM [TK].dbo.INVMB WITH(NOLOCK), [TK].dbo.COPTC WITH(NOLOCK), [TK].dbo.COPTD WITH(NOLOCK)
                                            LEFT JOIN [TK].dbo.INVMD ON INVMD.MD001=TD004 AND TD010=MD002
                                            LEFT JOIN [TK].dbo.BOMMC ON BOMMC.MC001=TD004
                                            LEFT JOIN [TK].dbo.BOMMD ON BOMMD.MD001=TD004
                                            LEFT JOIN [TK].dbo.BOMMC BOMMC2 ON BOMMC2.MC001=BOMMD.MD003
                                            LEFT JOIN [TK].dbo.BOMMD BOMMD2 ON BOMMD2.MD001=BOMMD.MD003
                                            LEFT JOIN [TK].dbo.BOMMC BOMMC3 ON BOMMC3.MC001=BOMMD2.MD003
                                            LEFT JOIN [TK].dbo.BOMMD BOMMD3 ON BOMMD3.MD001=BOMMD2.MD003
                                            LEFT JOIN [TK].dbo.BOMMC BOMMC4 ON BOMMC4.MC001=BOMMD3.MD003
                                            LEFT JOIN [TK].dbo.BOMMD BOMMD4 ON BOMMD4.MD001=BOMMD3.MD003
                                            LEFT JOIN [TK].dbo.BOMMC BOMMC5 ON BOMMC5.MC001=BOMMD4.MD003
                                            LEFT JOIN [TK].dbo.BOMMD BOMMD5 ON BOMMD5.MD001=BOMMD4.MD003
                                            WHERE TC001=TD001 AND TC002=TD002
                                              AND MB001=TD004
                                              AND TD001='{2}' AND TD002='{3}' AND TD003='{4}'
                                              AND (BOMMD.MD003 LIKE '3%' OR BOMMD.MD003 LIKE '4%')
                                              AND (BOMMD2.MD003 LIKE '3%' OR BOMMD2.MD003 LIKE '4%')
                                              AND (BOMMD3.MD003 LIKE '3%' OR BOMMD3.MD003 LIKE '4%')
                                              AND (BOMMD4.MD003 LIKE '3%' OR BOMMD4.MD003 LIKE '4%')
                                              AND (BOMMD5.MD003 LIKE '3%' OR BOMMD5.MD003 LIKE '4%')
                                        ) AS TEMP;

                                        ", MANU, MANUDATE, TD001, TD002, TD003);

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

                        MessageBox.Show("成功");
                    }

                }
                catch (Exception ex)
                {
                }
                finally
                {
                    sqlConn.Close();
                }
            }
            else if (KINDS.Equals("第六層"))
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


                    sbSql.AppendFormat(@"  
                                      INSERT INTO [TKMOC].[dbo].[MOCMANULINEBAKING]
                                      ([ID],[MANU],[MANUDATE],[MB001],[MB002],[MB003],[BAR],[NUM],[CLINET],[MANUHOUR],[BOX],[PACKAGE],[OUTDATE],[TA029],[HALFPRO],[COPTD001],[COPTD002],[COPTD003])
                                    
                                        SELECT ID,'{0}','{1}',MD003F [MB001],MD035F [MB002],MD036F [MB003],CONVERT(DECIMAL(16,4),(BOMNUMS6/MC004G)) [BAR],BOMNUMS6 [NUM],TC053 [CLINET],0 [MANUHOUR],0 [BOX],BOMNUMS6 [PACKAGE],TD013 [OUTDATE],TC015 [TA029],0 [HALFPRO],TD001 [COPTD001], TD002 [COPTD002], TD003 [COPTD003]
                                        FROM (
                                            SELECT NEWID() AS ID, TD001, TD002, TD003, TC053, TD004, TD005, TD006, (TD008+TD024) AS TD008, TD010, TC015, TD013,
                                                    (CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024) END) AS NUM,
                                                    BOMMD.MD003, BOMMD.MD035, BOMMD.MD036, BOMMD.MD006, BOMMD.MD007,
                                                    ((CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024) END)*BOMMD.MD006/BOMMD.MD007/BOMMC.MC004) AS BOMNUMS,
                                                    BOMMD2.MD003 MD003B, BOMMD2.MD035 MD035B, BOMMD2.MD036 MD036B, BOMMD2.MD006 MD006B, BOMMD2.MD007 MD007B,
                                                    (((CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024) END)*BOMMD.MD006/BOMMD.MD007/BOMMC.MC004)*BOMMD2.MD006/BOMMD2.MD007/BOMMC2.MC004) AS BOMNUMS2,
                                                    BOMMD3.MD003 MD003C, BOMMD3.MD035 MD035C, BOMMD3.MD036 MD036C, BOMMD3.MD006 MD006C, BOMMD3.MD007 MD007C,
                                                    ((((CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024) END)*BOMMD.MD006/BOMMD.MD007/BOMMC.MC004)*BOMMD2.MD006/BOMMD2.MD007/BOMMC2.MC004)*BOMMD3.MD006/BOMMD3.MD007/BOMMC3.MC004) AS BOMNUMS3,
                                                    BOMMD4.MD003 MD003D, BOMMD4.MD035 MD035D, BOMMD4.MD036 MD036D, BOMMD4.MD006 MD006D, BOMMD4.MD007 MD007D,
                                                    (((((CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024) END)*BOMMD.MD006/BOMMD.MD007/BOMMC.MC004)*BOMMD2.MD006/BOMMD2.MD007/BOMMC2.MC004)*BOMMD3.MD006/BOMMD3.MD007/BOMMC3.MC004)*BOMMD4.MD006/BOMMD4.MD007/BOMMC4.MC004) AS BOMNUMS4,
                                                    BOMMD5.MD003 MD003E, BOMMD5.MD035 MD035E, BOMMD5.MD036 MD036E, BOMMD5.MD006 MD006E, BOMMD5.MD007 MD007E,
                                                    ((((((CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024) END)*BOMMD.MD006/BOMMD.MD007/BOMMC.MC004)*BOMMD2.MD006/BOMMD2.MD007/BOMMC2.MC004)*BOMMD3.MD006/BOMMD3.MD007/BOMMC3.MC004)*BOMMD4.MD006/BOMMD4.MD007/BOMMC4.MC004)*BOMMD5.MD006/BOMMD5.MD007/BOMMC5.MC004) AS BOMNUMS5,
                                                    BOMMD6.MD003 MD003F, BOMMD6.MD035 MD035F, BOMMD6.MD036 MD036F, BOMMD6.MD006 MD006F, BOMMD6.MD007 MD007F,
                                                    (((((((CASE WHEN ISNULL(INVMD.MD002,'')<>'' THEN (TD008+TD024)*INVMD.MD004 ELSE (TD008+TD024) END)*BOMMD.MD006/BOMMD.MD007/BOMMC.MC004)*BOMMD2.MD006/BOMMD2.MD007/BOMMC2.MC004)*BOMMD3.MD006/BOMMD3.MD007/BOMMC3.MC004)*BOMMD4.MD006/BOMMD4.MD007/BOMMC4.MC004)*BOMMD5.MD006/BOMMD5.MD007/BOMMC5.MC004)*BOMMD6.MD006/BOMMD6.MD007/BOMMC6.MC004) AS BOMNUMS6,
                                                    ISNULL((SELECT TOP 1 MD007 FROM [TK].dbo.BOMMD MD WHERE MD.MD003 LIKE '201%' AND MD.MD001=BOMMD6.MD003),1) AS MD007G,
                                                    ISNULL((SELECT TOP 1 MC004 FROM [TK].dbo.BOMMC MC WHERE MC.MC001=BOMMD6.MD003),1) AS MC004G
                                            FROM [TK].dbo.INVMB WITH(NOLOCK), [TK].dbo.COPTC WITH(NOLOCK), [TK].dbo.COPTD WITH(NOLOCK)
                                            LEFT JOIN [TK].dbo.INVMD ON INVMD.MD001=TD004 AND TD010


                                        ", MANU, MANUDATE, TD001, TD002, TD003);

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

                        MessageBox.Show("成功");
                    }

                }
                catch (Exception ex)
                {
                }
                finally
                {
                    sqlConn.Close();
                }
            }
        }

        public void SEARCHTBCOPTDCHECK(string YYYYMM, string TD021, string UDF01, string TD002)
        {
            SqlConnection sqlConn = new SqlConnection();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();
            StringBuilder QUERYS = new StringBuilder();
            StringBuilder QUERYS2 = new StringBuilder();


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
                QUERYS.Clear();
                

                //日期
                if (!string.IsNullOrEmpty(YYYYMM))
                {
                    QUERYS.AppendFormat(@" AND TD002 LIKE '{0}%'", YYYYMM.Trim());

                }

                //核單
                if (!string.IsNullOrEmpty(TD021))
                {
                    if (TD021.Equals("未核單"))
                    {
                        QUERYS.AppendFormat(@" AND TD021='N'");
                    }
                    else if (TD021.Equals("已核單"))
                    {
                        QUERYS.AppendFormat(@"  AND TD021='Y'");
                    }
                }


                //是否生產
                if (!string.IsNullOrEmpty(UDF01))
                {
                    if (UDF01.Equals("Y"))
                    {
                        QUERYS.AppendFormat(@" AND COPTD.UDF01 IN ('Y','y') ");
                    }
                    else if (UDF01.Equals("N"))
                    {
                        QUERYS.AppendFormat(@" AND COPTD.UDF01 NOT IN ('Y','y')  ");
                    }
                }

                //訂單單號
                if (!string.IsNullOrEmpty(TD002))
                {
                    QUERYS.AppendFormat(@" AND TD002 LIKE '{0}%'", TD002.Trim());

                }

                //限定烘培品
                DataTable DT = SEARCH_MOCMANULINEMB001LIKES();
                if (DT != null && DT.Rows.Count >= 1)
                {
                    QUERYS2.AppendFormat(" AND (");
                    for (int i = 0; i < DT.Rows.Count; i++)
                    {
                        if (i > 0) // Add OR after the first condition
                        {
                            QUERYS2.AppendFormat(" OR ");
                        }
                        QUERYS2.AppendFormat("TD004 LIKE '{0}%'", DT.Rows[i]["MB001"].ToString());
                    }
                    QUERYS2.AppendFormat(")");
                }
                else
                {
                    // No additional SQL clause required
                    QUERYS2.AppendFormat("");
                }

                sbSql.Clear();
                sbSql.AppendFormat(@"  
                                    SELECT  
                                    COPTD.UDF01 AS '是否生產'
                                    ,TC053 AS '客戶',TD001 AS '單別',TD002 AS '單號',TD003 AS '序號',TD004 AS '品號',TD005 AS '品名',TD008 AS '訂單數量',TD024 AS '贈品量',TD009	AS '已交數量',TD025	AS '贈品已交',TD010	 AS '單位',TD013 AS '預交日',TC015 AS '單頭備註',TD020 AS '單身備註'

                                    ,(SELECT TOP 1 ISNULL([MOCCHECKDATES],'') FROM [TKBUSINESS].[dbo].[TBCOPTDCHECK] WHERE [TBCOPTDCHECK].TD001=COPTD.TD001 AND [TBCOPTDCHECK].TD002=COPTD.TD002 AND [TBCOPTDCHECK].TD003=COPTD.TD003  ORDER BY ID DESC) AS '生管更新日期'
                                    ,(SELECT TOP 1 [MOCCHECKS] FROM [TKBUSINESS].[dbo].[TBCOPTDCHECK] WHERE [TBCOPTDCHECK].TD001=COPTD.TD001 AND [TBCOPTDCHECK].TD002=COPTD.TD002 AND [TBCOPTDCHECK].TD003=COPTD.TD003  ORDER BY ID DESC) AS '生管核準'
                                    ,(SELECT TOP 1 [MOCCHECKSCOMMENTS] FROM [TKBUSINESS].[dbo].[TBCOPTDCHECK] WHERE [TBCOPTDCHECK].TD001=COPTD.TD001 AND [TBCOPTDCHECK].TD002=COPTD.TD002 AND [TBCOPTDCHECK].TD003=COPTD.TD003  ORDER BY ID DESC) AS '生管備註'
                                    ,'' AS '生管備註填寫'

                                    ,(SELECT TOP 1 [PURCHECKDATES] FROM [TKBUSINESS].[dbo].[TBCOPTDCHECK] WHERE [TBCOPTDCHECK].TD001=COPTD.TD001 AND [TBCOPTDCHECK].TD002=COPTD.TD002 AND [TBCOPTDCHECK].TD003=COPTD.TD003  ORDER BY ID DESC) AS '採購更新日期'
                                    ,(SELECT TOP 1 [PURCHECKS] FROM [TKBUSINESS].[dbo].[TBCOPTDCHECK] WHERE [TBCOPTDCHECK].TD001=COPTD.TD001 AND [TBCOPTDCHECK].TD002=COPTD.TD002 AND [TBCOPTDCHECK].TD003=COPTD.TD003  ORDER BY ID DESC) AS '採購核準'
                                    ,(SELECT TOP 1 [PURCHECKSCOMMENTS] FROM [TKBUSINESS].[dbo].[TBCOPTDCHECK] WHERE [TBCOPTDCHECK].TD001=COPTD.TD001 AND [TBCOPTDCHECK].TD002=COPTD.TD002 AND [TBCOPTDCHECK].TD003=COPTD.TD003  ORDER BY ID DESC) AS '採購備註'
                                    ,(SELECT TOP 1 [SALESCHECKDATES] FROM [TKBUSINESS].[dbo].[TBCOPTDCHECK] WHERE [TBCOPTDCHECK].TD001=COPTD.TD001 AND [TBCOPTDCHECK].TD002=COPTD.TD002 AND [TBCOPTDCHECK].TD003=COPTD.TD003  ORDER BY ID DESC) AS '業務更新日期'
                                    ,(SELECT TOP 1 [SALESCHECKSCOMMENTS] FROM [TKBUSINESS].[dbo].[TBCOPTDCHECK] WHERE [TBCOPTDCHECK].TD001=COPTD.TD001 AND [TBCOPTDCHECK].TD002=COPTD.TD002 AND [TBCOPTDCHECK].TD003=COPTD.TD003  ORDER BY ID DESC) AS '業務備註'

                                    
                                    FROM [TK].dbo.COPTC,[TK].dbo.COPTD
                                    WHERE TC001=TD001 AND TC002=TD002
                                    AND 1=1
                                
                                    {0}
                                    {1}
                                    ORDER BY TD002,TD001,TD003
                                    ", QUERYS.ToString(), QUERYS2.ToString());




                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                {
                    dataGridView28.DataSource = null;
                }
                else
                {
                    if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView28.DataSource = ds1.Tables["TEMPds1"];
                        dataGridView28.AutoResizeColumns();

                        NEWdataGridView28ComboBoxColumn();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                        dataGridView28.DefaultCellStyle.WrapMode = DataGridViewTriState.True;

                        dataGridView28.AutoResizeColumns();
                        dataGridView28.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCellsExceptHeaders;
                        dataGridView28.Columns["生管備註填寫"].Width = 200;
                        dataGridView28.Columns["生管備註填寫"].DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                        dataGridView28.Columns["生管備註填寫"].DefaultCellStyle.BackColor = Color.LightPink;

                        //設定欄位順序
                        dataGridView28.Columns["生管備註填寫"].DisplayIndex = 18;
                        dataGridView28.Columns["生管核準填寫"].DisplayIndex = 19;

                        if (!string.IsNullOrEmpty(SortedColumn))
                        {
                            if (SortedModel.Equals("Ascending"))
                            {
                                dataGridView28.Sort(dataGridView28.Columns["" + SortedColumn + ""], ListSortDirection.Ascending);
                            }
                            else
                            {
                                dataGridView28.Sort(dataGridView28.Columns["" + SortedColumn + ""], ListSortDirection.Descending);
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
                sqlConn.Close();
            }
        }

        public DataTable SEARCH_MOCMANULINEMB001LIKES()
        {
            SqlConnection sqlConn = new SqlConnection();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();
            StringBuilder QUERYS = new StringBuilder();
    

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
                QUERYS.Clear();
               
                sbSql.AppendFormat(@"  
                                    SELECT [MB001]
                                    FROM [TKMOC].[dbo].[MOCMANULINEMB001LIKES]
                                    ");
                

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1 != null && ds1.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["TEMPds1"];
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

        //新增datagrid的ComboBoxColumn
        public void NEWdataGridView28ComboBoxColumn()
        {
            SqlConnection sqlConn = new SqlConnection();

            //欄位是否存在
            bool yesOrNo = dataGridView28.Columns.Contains("生管核準填寫");

            //不存在欄位=生管核準填寫
            if (yesOrNo == false)
            {
                DataGridViewComboBoxColumn dgvCmb = new DataGridViewComboBoxColumn();
                dgvCmb.HeaderText = "生管核準填寫";
                dgvCmb.Name = "生管核準填寫";

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                StringBuilder Sequel = new StringBuilder();
                Sequel.AppendFormat(@"SELECT 'N' AS 'STATUS' UNION ALL SELECT 'Y' AS 'STATUS'  ");
                SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
                DataTable dt = new DataTable();
                sqlConn.Open();

                dt.Columns.Add("STATUS", typeof(string));

                da.Fill(dt);

                sqlConn.Close();

                //新增combox的item
                dgvCmb.DataSource = dt;
                dgvCmb.DisplayMember = "STATUS";
                dgvCmb.ValueMember = "STATUS";

                ////新增combox的item
                //dgvCmb.Items.Add("N");
                //dgvCmb.Items.Add("Y");
                //新增預設值
                dgvCmb.DefaultCellStyle.NullValue = "N";


                //欄位的表頭名稱
                dgvCmb.Name = "生管核準填寫";


                //加入到datagrid
                dataGridView28.Columns.Add(dgvCmb);
            }

        }

        public void CHECKdataGridView28()
        {
           
            string TD001 = null;
            string TD002 = null;
            string TD003 = null;
            string MOCCHECKSCOMMENTS = null;
            string MOCCHECKS = "N";

            dataGridView28.EndEdit();

            foreach (DataGridViewRow row in dataGridView28.Rows)
            {
                TD001 = row.Cells["單別"].Value.ToString();
                TD002 = row.Cells["單號"].Value.ToString();
                TD003 = row.Cells["序號"].Value.ToString();
                MOCCHECKSCOMMENTS = row.Cells["生管備註填寫"].Value.ToString();

                var cell = row.Cells["生管核準填寫"].Value;
                if (cell != null && cell.ToString().Equals("Y"))
                {
                    MOCCHECKS = cell.ToString();
                }
                else
                {
                    MOCCHECKS = "N";
                }

                if (!string.IsNullOrEmpty(MOCCHECKSCOMMENTS))
                {
                    ADDTBCOPTDCHECKMOC(TD001, TD002, TD003, null, MOCCHECKS, MOCCHECKSCOMMENTS);
                    //MessageBox.Show(TD001+ TD002+ TD003+ MOCCHECKSCOMMENTS+ MOCCHECKS);
                }
            }
        }

        public void ADDTBCOPTDCHECKMOC(string TD001,
                              string TD002,
                              string TD003,
                              string MOCCHECKDATES,
                              string MOCCHECKS,
                              string MOCCHECKSCOMMENTS
                             )
        {
            SqlConnection sqlConn = new SqlConnection();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            MOCCHECKDATES = DateTime.Now.ToString("yyyyMMdd HH:mm:ss");

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


                sbSql.AppendFormat(@" 
                                    INSERT INTO [TKBUSINESS].[dbo].[TBCOPTDCHECK]
                                    ([TD001]
                                    ,[TD002]
                                    ,[TD003]
                                    ,[TD004]
                                    ,[TD005]
                                    ,[TD008]
                                    ,[TD009]
                                    ,[TD010]
                                    ,[TD011]
                                    ,[TD012]
                                    ,[TD013]
                                    ,[TD024]
                                    ,[TD025]
                                    ,[TC015]
                                    ,[TD020]
                                    ,[MOCCHECKDATES]
                                    ,[MOCCHECKS]
                                    ,[MOCCHECKSCOMMENTS]
                                    ,[PURCHECKDATES]
                                    ,[PURCHECKS]
                                    ,[PURCHECKSCOMMENTS]
                                    ,[SALESCHECKDATES]
                                    ,[SALESCHECKSCOMMENTS]
             
                                    )
                                    SELECT 
                                    [TD001]
                                    ,[TD002]
                                    ,[TD003]
                                    ,[TD004]
                                    ,[TD005]
                                    ,[TD008]
                                    ,[TD009]
                                    ,[TD010]
                                    ,[TD011]
                                    ,[TD012]
                                    ,[TD013]
                                    ,[TD024]
                                    ,[TD025]
                                    ,[TC015]
                                    ,[TD020]
                                    ,'{3}' AS [MOCCHECKDATES]
                                    ,'{4}' AS [MOCCHECKS]
                                    ,'{5}' AS [MOCCHECKSCOMMENTS]
                                    ,(SELECT TOP 1 [PURCHECKDATES] FROM [TKBUSINESS].[dbo].[TBCOPTDCHECK] WHERE [TBCOPTDCHECK].TD001=COPTD.TD001 AND [TBCOPTDCHECK].TD002=COPTD.TD002 AND [TBCOPTDCHECK].TD003=COPTD.TD003 ORDER BY ID DESC) AS [PURCHECKDATES]
                                    ,(SELECT TOP 1 [PURCHECKS] FROM [TKBUSINESS].[dbo].[TBCOPTDCHECK] WHERE [TBCOPTDCHECK].TD001=COPTD.TD001 AND [TBCOPTDCHECK].TD002=COPTD.TD002 AND [TBCOPTDCHECK].TD003=COPTD.TD003 ORDER BY ID DESC) AS [PURCHECKS]
                                    ,(SELECT TOP 1 [PURCHECKSCOMMENTS] FROM [TKBUSINESS].[dbo].[TBCOPTDCHECK] WHERE [TBCOPTDCHECK].TD001=COPTD.TD001 AND [TBCOPTDCHECK].TD002=COPTD.TD002 AND [TBCOPTDCHECK].TD003=COPTD.TD003 ORDER BY ID DESC) AS [PURCHECKSCOMMENTS]
                                    ,(SELECT TOP 1 [SALESCHECKDATES] FROM [TKBUSINESS].[dbo].[TBCOPTDCHECK] WHERE [TBCOPTDCHECK].TD001=COPTD.TD001 AND [TBCOPTDCHECK].TD002=COPTD.TD002 AND [TBCOPTDCHECK].TD003=COPTD.TD003 ORDER BY ID DESC) AS [SALESCHECKDATES]
                                    ,(SELECT TOP 1 [SALESCHECKSCOMMENTS] FROM [TKBUSINESS].[dbo].[TBCOPTDCHECK] WHERE [TBCOPTDCHECK].TD001=COPTD.TD001 AND [TBCOPTDCHECK].TD002=COPTD.TD002 AND [TBCOPTDCHECK].TD003=COPTD.TD003 ORDER BY ID DESC) AS [SALESCHECKSCOMMENTS]
                                    FROM [TK].dbo.COPTD,[TK].dbo.COPTC
                                    WHERE TC001=TD001 AND TC002=TD002
                                    AND TD001='{0}' AND TD002='{1}' AND TD003='{2}'

                                    ", TD001, TD002, TD003, MOCCHECKDATES, MOCCHECKS, MOCCHECKSCOMMENTS);


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

                    //MessageBox.Show("完成");
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
        public void ADD_MOCMANULINEBAKING()
        {
            DataTable COPTCTD = new DataTable();

            Guid ID = new Guid();
            string MANU = null;
            string MANUDATE = null;
            string MB001 = null;
            string MB002 = null;
            string MB003 = null;
            string BAR = null;
            string NUM = null;
            string CLINET = null;
            string TA029 = null;
            string OUTDATE = null;
            string HALFPRO = null;
            string COPTD001 = null;
            string COPTD002 = null;
            string COPTD003 = null;
            string BOX = null;
            string PACKAGE = null;




            if (dataGridView28.Rows.Count > 0)
            {
                foreach (DataGridViewRow dr in this.dataGridView28.Rows)
                {
                    if (dr.Cells[0].Value != null && (bool)dr.Cells[0].Value)
                    {
                        //找出訂單明細、桶數、箱數
                        COPTCTD = SEARCHCOPTCTDDATA(dr.Cells["單別"].Value.ToString().Trim(), dr.Cells["單號"].Value.ToString().Trim(), dr.Cells["序號"].Value.ToString().Trim(), dr.Cells["品號"].Value.ToString().Trim());

                        if (COPTCTD.Rows.Count > 0)
                        {
                            MANU = comboBox25.SelectedValue.ToString().Trim();
                            MANUDATE = dateTimePicker29.Value.ToString("yyyy/MM/dd");
                            MB001 = COPTCTD.Rows[0]["TD004"].ToString();
                            MB002 = COPTCTD.Rows[0]["TD005"].ToString();
                            MB003 = COPTCTD.Rows[0]["TD006"].ToString();
                            BAR = COPTCTD.Rows[0]["BARS"].ToString();
                            NUM = COPTCTD.Rows[0]["TD008"].ToString();
                            CLINET = COPTCTD.Rows[0]["TC053"].ToString();
                            TA029 = COPTCTD.Rows[0]["TC015"].ToString();
                            OUTDATE = COPTCTD.Rows[0]["TD013"].ToString().Substring(0, 4) + "/" + COPTCTD.Rows[0]["TD013"].ToString().Substring(4, 2) + "/" + COPTCTD.Rows[0]["TD013"].ToString().Substring(6, 2);
                            HALFPRO = "0";
                            COPTD001 = COPTCTD.Rows[0]["TD001"].ToString();
                            COPTD002 = COPTCTD.Rows[0]["TD002"].ToString();
                            COPTD003 = COPTCTD.Rows[0]["TD003"].ToString();
                            if (string.IsNullOrEmpty(COPTCTD.Rows[0]["BOXS"].ToString()))
                            {
                                BOX = "0";
                            }
                            else
                            {
                                BOX = COPTCTD.Rows[0]["BOXS"].ToString();
                            }


                            PACKAGE = COPTCTD.Rows[0]["TD008"].ToString();
                        }


                        if (comboBox25.SelectedValue.Equals("烘焙生產線"))
                        {
                            ADDNEWTOTKMOCMOCMANULINE(ID, MANU, MANUDATE, MB001, MB002, MB003, BAR, NUM, CLINET, TA029, OUTDATE, HALFPRO, COPTD001, COPTD002, COPTD003, BOX, PACKAGE);
                        }                        
                        else
                        {

                        }
                    }
                    
                }
            }



            // MessageBox.Show(comboBox25.SelectedValue.ToString());
        }

        public DataTable SEARCHCOPTCTDDATA(string TD001, string TD002, string TD003, string TD004)
        {
            SqlConnection sqlConn = new SqlConnection();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();
            StringBuilder QUERYS = new StringBuilder();

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
                QUERYS.Clear();
              

                // ,(CASE WHEN ISNULL(MC004,0)>0 THEN CONVERT(decimal(16,4),((TD008+TD024)/MC004)) END) AS BOXS
                sbSql.AppendFormat(@"  
                                    SELECT TD001,TD002,TD003,TC053,TD004,TD005,TD006,(TD008+TD024) AS TD008,TD010,(TC015+'-'+TD020) TC015 ,TD013
                                    ,(CASE WHEN ISNULL(MD002,'')<>'' THEN (TD008+TD024)*MD004 ELSE (TD008+TD024)  END ) AS NUM
                                    ,MC004,MB017

                                    ,CASE WHEN ISNULL(MC004,0)>0 THEN CONVERT(decimal(16,4),((TD008+TD024)/MC004)) END AS BARS
                                    ,(CASE WHEN ISNULL(MD002,'')<>'' THEN (TD008+TD024)*MD004 ELSE (TD008+TD024)  END ) AS NUMS
                                    ,(CASE WHEN ISNULL(TEMP.MD007,0)>0 THEN CONVERT(decimal(16,4),((TD008+TD024)/TEMP.MD007)) END) AS BOXS
                                    ,TEMP.MD007

                                    FROM [TK].dbo.INVMB WITH(NOLOCK),[TK].dbo.COPTC WITH(NOLOCK),[TK].dbo.COPTD WITH(NOLOCK)
                                    LEFT JOIN [TK].dbo.INVMD ON MD001=TD004 AND TD010=MD002
                                    LEFT JOIN [TK].dbo.BOMMC ON TD004=MC001
                                    LEFT JOIN [TKMOC].[dbo].[MOCHALFPRODUCTDBOXS] ON TD004=[MOCHALFPRODUCTDBOXS].[MB001]
                                    LEFT JOIN
                                    (
                                    SELECT TOP 1 MD001,MD003,MB001,MB002,ISNULL(MD007,1) AS MD007,ISNULL(MD010,1) AS MD010
                                    FROM [TK].dbo.BOMMD,[TK].dbo.INVMB
                                    WHERE MD003=MB001
                                    AND MB002 LIKE '%箱%'
                                    AND MD003 LIKE '2%'
                                    AND MD001='{3}'
                                    ) AS TEMP ON TEMP.MD001=COPTD.TD004

                                    WHERE TC001=TD001 AND TC002=TD002
                                    AND INVMB.MB001=TD004
                                    AND TD001='{0}' AND TD002='{1}' AND TD003='{2}'

                                    ", TD001, TD002, TD003, TD004);




                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count > 0)
                {
                    return ds1.Tables["TEMPds1"];
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

        public void ADDNEWTOTKMOCMOCMANULINE(
                                            Guid ID,
                                            string MANU,
                                            string MANUDATE,
                                            string MB001,
                                            string MB002,
                                            string MB003,
                                            string BAR,
                                            string NUM,
                                            string CLINET,
                                            string TA029,
                                            string OUTDATE,
                                            string HALFPRO,
                                            string COPTD001,
                                            string COPTD002,
                                            string COPTD003,
                                            string BOX,
                                            string PACKAGE
                                            )
        {
            Guid NEWGUID = new Guid();
            ID = Guid.NewGuid();

            SqlConnection sqlConn = new SqlConnection();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();
            StringBuilder QUERYS = new StringBuilder();

            if (MANU.Equals("烘焙生產線"))
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

                    sbSql.AppendFormat(@" 
                                        INSERT INTO [TKMOC].[dbo].[MOCMANULINEBAKING]
                                        ([ID],[MANU],[MANUDATE],[MB001],[MB002],[MB003],[BAR],[NUM],[CLINET],[TA029],[OUTDATE],[HALFPRO],[COPTD001],[COPTD002],[COPTD003],[BOX],[PACKAGE])
                                        VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}',N'{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}')
                                        ", ID, MANU, MANUDATE, MB001, MB002, MB003, BAR, NUM, CLINET, TA029, OUTDATE, HALFPRO, COPTD001, COPTD002, COPTD003, BOX, PACKAGE);


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
                        //UPDATEMOCMANULINETEMP(NEWGUID, TEMPds);

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

        }

        public void SEARCHTBCOPTFCHECK(string YYYY, string TF019, string TF002)
        {
            SqlConnection sqlConn = new SqlConnection();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();
            StringBuilder QUERYS = new StringBuilder();

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
                QUERYS.Clear();
                

                //日期
                if (!string.IsNullOrEmpty(YYYY))
                {
                    QUERYS.AppendFormat(@" AND TF002 LIKE '{0}%' ", YYYY.ToString().Trim());

                }

                //核單
                if (!string.IsNullOrEmpty(TF019))
                {
                    if (TF019.Equals("未核單"))
                    {
                        QUERYS.AppendFormat(@" AND TE029='N' ");
                    }
                    else if (TF019.Equals("已核單"))
                    {
                        QUERYS.AppendFormat(@"  AND TE029='Y' ");
                    }
                }
                

                //訂單單號
                if (!string.IsNullOrEmpty(TF002))
                {
                    QUERYS.AppendFormat(@" AND TF002 LIKE '{0}%'", TF002.ToString().Trim());

                }

                sbSql.AppendFormat(@"  
                                    SELECT 
                                    COPTF.UDF01 AS '是否生產'
                                    ,TC053 AS '客戶',TF001 AS '單別',TF002 AS '單號',TF104 AS '原序號',TF004 AS '新序號'
                                    ,TF003 AS '變更版次',TF005 AS '品號',TF006 AS '品名',TF009 AS '新訂單數量'
                                    ,TF020 AS '新贈品量',TF010 AS '單位',TF015 AS '預交日',TE050 AS '單頭備註',TE006 AS '單頭變更原因'
                                    ,TF032 AS '單身備註',TF018 AS '單身變更原因'
                                  
                                    ,(SELECT TOP 1 ISNULL(MOCCHECKDATES,'') FROM [TKBUSINESS].[dbo].[TBCOPTFCHECK] WHERE TBCOPTFCHECK.TF001=COPTF.TF001 AND TBCOPTFCHECK.TF002=COPTF.TF002 AND TBCOPTFCHECK.TF003=COPTF.TF003 AND  TBCOPTFCHECK.TF004=COPTF.TF004 ORDER BY ID DESC) AS '生管更新日期'
                                    ,(SELECT TOP 1 ISNULL(MOCCHECKS,'') FROM [TKBUSINESS].[dbo].[TBCOPTFCHECK] WHERE TBCOPTFCHECK.TF001=COPTF.TF001 AND TBCOPTFCHECK.TF002=COPTF.TF002 AND TBCOPTFCHECK.TF003=COPTF.TF003 AND  TBCOPTFCHECK.TF004=COPTF.TF004 ORDER BY ID DESC) AS '生管核準'
                                    ,(SELECT TOP 1 ISNULL(MOCCHECKSCOMMENTS,'') FROM [TKBUSINESS].[dbo].[TBCOPTFCHECK] WHERE TBCOPTFCHECK.TF001=COPTF.TF001 AND TBCOPTFCHECK.TF002=COPTF.TF002 AND TBCOPTFCHECK.TF003=COPTF.TF003 AND  TBCOPTFCHECK.TF004=COPTF.TF004 ORDER BY ID DESC) AS '生管備註'
                                    ,'' AS '生管備註填寫'

                                    ,(SELECT TOP 1 ISNULL(PURCHECKDATES,'') FROM [TKBUSINESS].[dbo].[TBCOPTFCHECK] WHERE TBCOPTFCHECK.TF001=COPTF.TF001 AND TBCOPTFCHECK.TF002=COPTF.TF002 AND TBCOPTFCHECK.TF003=COPTF.TF003 AND  TBCOPTFCHECK.TF004=COPTF.TF004 ORDER BY ID DESC) AS '採購更新日期'
                                    ,(SELECT TOP 1 ISNULL(PURCHECKS,'') FROM [TKBUSINESS].[dbo].[TBCOPTFCHECK] WHERE TBCOPTFCHECK.TF001=COPTF.TF001 AND TBCOPTFCHECK.TF002=COPTF.TF002 AND TBCOPTFCHECK.TF003=COPTF.TF003 AND  TBCOPTFCHECK.TF004=COPTF.TF004 ORDER BY ID DESC) AS '採購核準'
                                    ,(SELECT TOP 1 ISNULL(PURCHECKSCOMMENTS,'') FROM [TKBUSINESS].[dbo].[TBCOPTFCHECK] WHERE TBCOPTFCHECK.TF001=COPTF.TF001 AND TBCOPTFCHECK.TF002=COPTF.TF002 AND TBCOPTFCHECK.TF003=COPTF.TF003 AND  TBCOPTFCHECK.TF004=COPTF.TF004 ORDER BY ID DESC) AS '採購備註'
                                    ,(SELECT TOP 1 ISNULL(SALESCHECKDATES,'') FROM [TKBUSINESS].[dbo].[TBCOPTFCHECK] WHERE TBCOPTFCHECK.TF001=COPTF.TF001 AND TBCOPTFCHECK.TF002=COPTF.TF002 AND TBCOPTFCHECK.TF003=COPTF.TF003 AND  TBCOPTFCHECK.TF004=COPTF.TF004 ORDER BY ID DESC) AS '業務更新日期'
                                    ,(SELECT TOP 1 ISNULL(SALESCHECKSCOMMENTS,'') FROM [TKBUSINESS].[dbo].[TBCOPTFCHECK] WHERE TBCOPTFCHECK.TF001=COPTF.TF001 AND TBCOPTFCHECK.TF002=COPTF.TF002 AND TBCOPTFCHECK.TF003=COPTF.TF003 AND  TBCOPTFCHECK.TF004=COPTF.TF004 ORDER BY ID DESC) AS '業務備註'
                                    ,TE001,TE002,TE003,TF001

                                    FROM [TK].dbo.COPTE,[TK].dbo.COPTF
                                    LEFT JOIN [TK].dbo.COPTC ON TC001=TF001 AND TC002=TF002
                                    LEFT JOIN [TK].dbo.COPTD ON TD001=TF001 AND TD002=TF002 AND TD003=TF104
                                    WHERE TE001=TF001 AND TE002=TF002 AND TE003=TF003
                                    AND 1=1
                                    {0}


                                    ", QUERYS.ToString());





                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                {
                    dataGridView29.DataSource = null;
                }
                else
                {
                    if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView29.DataSource = ds1.Tables["TEMPds1"];
                        dataGridView29.AutoResizeColumns();

                        NEWdataGridView29ComboBoxColumn();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                        dataGridView29.DefaultCellStyle.WrapMode = DataGridViewTriState.True;

                        dataGridView29.AutoResizeColumns();
                        dataGridView29.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCellsExceptHeaders;
                        dataGridView29.Columns["生管備註填寫"].Width = 200;
                        dataGridView29.Columns["生管備註填寫"].DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                        dataGridView29.Columns["生管備註填寫"].DefaultCellStyle.BackColor = Color.LightPink;

                        //設定欄位順序
                        dataGridView29.Columns["生管備註填寫"].DisplayIndex = 20;
                        dataGridView29.Columns["生管核準填寫"].DisplayIndex = 21;

                        if (!string.IsNullOrEmpty(SortedColumn))
                        {
                            if (SortedModel.Equals("Ascending"))
                            {
                                dataGridView29.Sort(dataGridView29.Columns["" + SortedColumn + ""], ListSortDirection.Ascending);
                            }
                            else
                            {
                                dataGridView29.Sort(dataGridView29.Columns["" + SortedColumn + ""], ListSortDirection.Descending);
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
                sqlConn.Close();
            }
        }
        public void NEWdataGridView29ComboBoxColumn()
        {
            SqlConnection sqlConn = new SqlConnection();
            //欄位是否存在
            bool yesOrNo = dataGridView29.Columns.Contains("生管核準填寫");

            //不存在欄位=生管核準填寫
            if (yesOrNo == false)
            {
                DataGridViewComboBoxColumn dgvCmb = new DataGridViewComboBoxColumn();
                dgvCmb.HeaderText = "生管核準填寫";
                dgvCmb.Name = "生管核準填寫";

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                StringBuilder Sequel = new StringBuilder();
                Sequel.AppendFormat(@"SELECT 'N' AS 'STATUS' UNION ALL SELECT 'Y' AS 'STATUS'  ");
                SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
                DataTable dt = new DataTable();
                sqlConn.Open();

                dt.Columns.Add("STATUS", typeof(string));

                da.Fill(dt);

                sqlConn.Close();

                //新增combox的item
                dgvCmb.DataSource = dt;
                dgvCmb.DisplayMember = "STATUS";
                dgvCmb.ValueMember = "STATUS";

                ////新增combox的item
                //dgvCmb.Items.Add("N");
                //dgvCmb.Items.Add("Y");
                //新增預設值
                dgvCmb.DefaultCellStyle.NullValue = "N";


                //欄位的表頭名稱
                dgvCmb.Name = "生管核準填寫";


                //加入到datagrid
                dataGridView29.Columns.Add(dgvCmb);
            }

        }

        public void CHECKdataGridView29()
        {
            string TF001 = null;
            string TF002 = null;
            string TF003 = null;
            string TF004 = null;
            string MOCCHECKSCOMMENTS = null;
            string MOCCHECKS = "N";

            dataGridView29.EndEdit();

            foreach (DataGridViewRow row in dataGridView29.Rows)
            {
                TF001 = row.Cells["單別"].Value.ToString();
                TF002 = row.Cells["單號"].Value.ToString();
                TF003 = row.Cells["變更版次"].Value.ToString();
                TF004 = row.Cells["新序號"].Value.ToString();
                MOCCHECKSCOMMENTS = row.Cells["生管備註填寫"].Value.ToString();

                var cell = row.Cells["生管核準填寫"].Value;
                if (cell != null && cell.ToString().Equals("Y"))
                {
                    MOCCHECKS = cell.ToString();
                }
                else
                {
                    MOCCHECKS = "N";
                }

                if (!string.IsNullOrEmpty(MOCCHECKSCOMMENTS))
                {
                    ADDTBCOPTFCHECKMOC(TF001, TF002, TF003, TF004, null, MOCCHECKS, MOCCHECKSCOMMENTS);
                    //MessageBox.Show(TD001+ TD002+ TD003+ MOCCHECKSCOMMENTS+ MOCCHECKS);
                }
            }
        }

        public void ADDTBCOPTFCHECKMOC(string TF001,
                              string TF002,
                              string TF003,
                              string TF004,
                              string MOCCHECKDATES,
                              string MOCCHECKS,
                              string MOCCHECKSCOMMENTS
                             )
        {
            SqlConnection sqlConn = new SqlConnection();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            MOCCHECKDATES = DateTime.Now.ToString("yyyyMMdd HH:mm:ss");

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


                sbSql.AppendFormat(@" 
                                    INSERT INTO [TKBUSINESS].[dbo].[TBCOPTFCHECK]
                                    (
                                    [TF001]
                                    ,[TF002]
                                    ,[TF003]
                                    ,[TF004]
                                    ,[TF005]
                                    ,[TF006]
                                    ,[TF007]
                                    ,[TF009]
                                    ,[TF010]
                                    ,[TF013]
                                    ,[TF014]
                                    ,[TF015]
                                    ,[TF018]
                                    ,[TF032]
                                    ,[TF045]
                                    ,[TF104]
                                    ,[TE006]
                                    ,[TE050]
                                    ,[MOCCHECKDATES]
                                    ,[MOCCHECKS]
                                    ,[MOCCHECKSCOMMENTS]
                                    ,[PURCHECKDATES]
                                    ,[PURCHECKS]
                                    ,[PURCHECKSCOMMENTS]
                                    ,[SALESCHECKDATES]
                                    ,[SALESCHECKSCOMMENTS]
                                    )

                                    SELECT
                                    [TF001]
                                    ,[TF002]
                                    ,[TF003]
                                    ,[TF004]
                                    ,[TF005]
                                    ,[TF006]
                                    ,[TF007]
                                    ,[TF009]
                                    ,[TF010]
                                    ,[TF013]
                                    ,[TF014]
                                    ,[TF015]
                                    ,[TF018]
                                    ,[TF032]
                                    ,[TF045]
                                    ,[TF104]
                                    ,[TE006]
                                    ,[TE050]
                                    ,'{4}' AS 'MOCCHECKDATES'
                                    ,'{5}' AS 'MOCCHECKS'
                                    ,'{6}' AS 'MOCCHECKSCOMMENTS'
                                    ,(SELECT TOP 1 ISNULL(PURCHECKDATES,'') FROM [TKBUSINESS].[dbo].[TBCOPTFCHECK] WHERE TBCOPTFCHECK.TF001=COPTF.TF001 AND TBCOPTFCHECK.TF002=COPTF.TF002 AND TBCOPTFCHECK.TF003=COPTF.TF003 AND  TBCOPTFCHECK.TF004=COPTF.TF004 ORDER BY ID DESC) AS 'PURCHECKDATES'
                                    ,(SELECT TOP 1 ISNULL(PURCHECKS,'') FROM [TKBUSINESS].[dbo].[TBCOPTFCHECK] WHERE TBCOPTFCHECK.TF001=COPTF.TF001 AND TBCOPTFCHECK.TF002=COPTF.TF002 AND TBCOPTFCHECK.TF003=COPTF.TF003 AND  TBCOPTFCHECK.TF004=COPTF.TF004 ORDER BY ID DESC) AS 'PURCHECKS'
                                    ,(SELECT TOP 1 ISNULL(PURCHECKSCOMMENTS,'') FROM [TKBUSINESS].[dbo].[TBCOPTFCHECK] WHERE TBCOPTFCHECK.TF001=COPTF.TF001 AND TBCOPTFCHECK.TF002=COPTF.TF002 AND TBCOPTFCHECK.TF003=COPTF.TF003 AND  TBCOPTFCHECK.TF004=COPTF.TF004 ORDER BY ID DESC) AS 'PURCHECKSCOMMENTS'
                                    ,(SELECT TOP 1 ISNULL(SALESCHECKDATES,'') FROM [TKBUSINESS].[dbo].[TBCOPTFCHECK] WHERE TBCOPTFCHECK.TF001=COPTF.TF001 AND TBCOPTFCHECK.TF002=COPTF.TF002 AND TBCOPTFCHECK.TF003=COPTF.TF003 AND  TBCOPTFCHECK.TF004=COPTF.TF004 ORDER BY ID DESC) AS 'SALESCHECKDATES'
                                    ,(SELECT TOP 1 ISNULL(SALESCHECKSCOMMENTS,'') FROM [TKBUSINESS].[dbo].[TBCOPTFCHECK] WHERE TBCOPTFCHECK.TF001=COPTF.TF001 AND TBCOPTFCHECK.TF002=COPTF.TF002 AND TBCOPTFCHECK.TF003=COPTF.TF003 AND  TBCOPTFCHECK.TF004=COPTF.TF004 ORDER BY ID DESC) AS 'SALESCHECKSCOMMENTS'
                                    FROM [TK].dbo.COPTE,[TK].dbo.COPTF
                                    WHERE TE001=TF001 AND TE002=TF002 AND TE003=TF003
                                    AND TF001='{0}' AND TF002='{1}' AND TF003='{2}' AND TF004='{3}'

                                    ", TF001, TF002, TF003, TF004, MOCCHECKDATES, MOCCHECKS, MOCCHECKSCOMMENTS);


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

                    //MessageBox.Show("完成");
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

        private void dataGridView29_SelectionChanged(object sender, EventArgs e)
        {
            
            if (dataGridView29.CurrentRow != null)
            {
                int rowindex = dataGridView29.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView29.Rows[rowindex];
                    TF001 = row.Cells["單別"].Value.ToString();
                    TF002 = row.Cells["單號"].Value.ToString();
                    TF003 = row.Cells["變更版次"].Value.ToString();
                    TF104 = row.Cells["原序號"].Value.ToString();


                }
                else
                {
                    TF001 = null;
                    TF002 = null;
                    TF003 = null;
                    TF104 = null;


                }
            }
        }

        public void ADDTOTKMOCMOCMANULINECOPTECOPTF(string TF001, string TF002, string TF003, string TF104)
        {
            DataTable COPTETF = new DataTable();

            Guid ID = new Guid();
            string MANU = null;
            string MANUDATE = null;
            string MB001 = null;
            string MB002 = null;
            string MB003 = null;
            string BAR = null;
            string NUM = null;
            string CLINET = null;
            string TA029 = null;
            string OUTDATE = null;
            string HALFPRO = null;
            string COPTD001 = null;
            string COPTD002 = null;
            string COPTD003 = null;
            string BOX = null;
            string PACKAGE = null;


            if (!string.IsNullOrEmpty(TF001))
            {

                //找出訂單變更的明細、桶數、箱數
                COPTETF = SEARCHCOPTETFDATA(TF001, TF002, TF003, TF104);


                if (COPTETF.Rows.Count > 0)
                {
                    MANU = comboBox28.SelectedValue.ToString().Trim();
                    MANUDATE = dateTimePicker31.Value.ToString("yyyy/MM/dd");
                    MB001 = COPTETF.Rows[0]["TF005"].ToString();
                    MB002 = COPTETF.Rows[0]["TF006"].ToString();
                    MB003 = COPTETF.Rows[0]["TF007"].ToString();
                    BAR = COPTETF.Rows[0]["BARS"].ToString();
                    NUM = COPTETF.Rows[0]["NUM"].ToString();
                    CLINET = COPTETF.Rows[0]["TE055"].ToString();
                    TA029 = COPTETF.Rows[0]["TE006"].ToString();
                    OUTDATE = COPTETF.Rows[0]["TF015"].ToString().Substring(0, 4) + "/" + COPTETF.Rows[0]["TF015"].ToString().Substring(4, 2) + "/" + COPTETF.Rows[0]["TF015"].ToString().Substring(6, 2);
                    HALFPRO = "0";
                    COPTD001 = COPTETF.Rows[0]["TF001"].ToString();
                    COPTD002 = COPTETF.Rows[0]["TF002"].ToString();
                    COPTD003 = COPTETF.Rows[0]["TF104"].ToString();
                    BOX = COPTETF.Rows[0]["BOXS"].ToString();
                    PACKAGE = COPTETF.Rows[0]["NUM"].ToString();
                }


                if (comboBox28.SelectedValue.Equals("烘焙生產線"))
                {
                    ADDNEWTOTKMOCMOCMANULINE(ID, MANU, MANUDATE, MB001, MB002, MB003, BAR, NUM, CLINET, TA029, OUTDATE, HALFPRO, COPTD001, COPTD002, COPTD003, BOX, PACKAGE);

                }
                
            }

         

        }

        public DataTable SEARCHCOPTETFDATA(string TF001, string TF002, string TF003, string TF104)
        {
            SqlConnection sqlConn = new SqlConnection();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();
            StringBuilder QUERYS = new StringBuilder();

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
                QUERYS.Clear();
             

                sbSql.AppendFormat(@"  
                                   SELECT TF001,TF002,TF003,TF004,TE055,TF005,TF006,TF007,(TF009+TF020) AS TF009,TF010,(TE006+'-'+TE050+'-'+TF018+'-'+TF032) TE006 ,TF015,TF104
                                    ,(CASE WHEN ISNULL(MD002,'')<>'' THEN (TF009+TF020)*MD004 ELSE (TF009+TF020)  END ) AS NUM
                                    ,MC004,MB017

                                    ,CASE WHEN ISNULL(MC004,0)>0 THEN CONVERT(decimal(16,4),((TF009+TF020)/MC004)) END AS BARS
                                    ,(CASE WHEN ISNULL([NUMS],0)<>0 THEN [NUMS] ELSE 1  END ) AS NUMS
                                    ,(CASE WHEN ISNULL([BOXS],0)<>0 THEN [BOXS] ELSE 1  END ) AS BOXS

                                    FROM [TK].dbo.INVMB WITH(NOLOCK),[TK].dbo.COPTE WITH(NOLOCK),[TK].dbo.COPTF WITH(NOLOCK)
                                    LEFT JOIN [TK].dbo.INVMD ON MD001=TF005 AND TF010=MD002
                                    LEFT JOIN [TK].dbo.BOMMC ON TF005=MC001
                                    LEFT JOIN [TKMOC].[dbo].[MOCHALFPRODUCTDBOXS] ON TF005=[MOCHALFPRODUCTDBOXS].[MB001]
                                    WHERE TE001=TF001 AND TE002=TF002 AND TE003=TF003
                                    AND INVMB.MB001=TF005
                                    AND TF001='{0}' AND TF002='{1}' AND TF003='{2}' AND TF104='{3}'

                                    ", TF001, TF002, TF003, TF104);




                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count > 0)
                {
                    return ds1.Tables["TEMPds1"];
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

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            string MB001 = textBox4.Text.Trim();

            SEARCHMB001(MB001);
            SEARCHMOCMANULINETEMPDATAS(MB001);
        }


       

        private void textBox15_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox15.Text.ToString()) && !string.IsNullOrEmpty(textBox14.Text.ToString()))
            {
                CAL_BAR2(textBox15.Text.ToString(), textBox14.Text.ToString());
            }
        }

        private void dataGridView6_SelectionChanged(object sender, EventArgs e)
        {
            DataGridView DG = new DataGridView();
            DG = dataGridView6;

            if (DG.CurrentRow != null)
            {
                int rowindex = DG.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = DG.Rows[rowindex];
                    DELID_DV6 = row.Cells["SID"].Value.ToString();
                    DELMOCTA001_DV6 = row.Cells["製令"].Value.ToString();
                    DELMOCTA002_DV6 = row.Cells["單號"].Value.ToString();



                }
                else
                {
                    DELID_DV6 = null;

                }
            }
        }

        public void SETNULL()
        {
            textBox7.Text = null;
            textBox8.Text = "0";
            textBox1.Text = "0";
            textBox9.Text = null;
            textBox10.Text = null;
            textBox11.Text = null;
            textBox12.Text = null;
            textBox13.Text = "0";
            textBox33.Text = "0";
            textBox53.Text = null;
            textBox68.Text = "0";
            textBox42.Text = null;
            textBox43.Text = null;
            textBox72.Text = null;

            textBox4.Text = null;
            textBox6.Text = "0";
            textBox21.Text = "0";
            textBox5.Text = null;
            textBox16.Text = null;
            textBox17.Text = null;
            textBox18.Text = null;
            textBox23.Text = "0";
            textBox14.Text = "0";
            textBox18.Text = null;
            textBox23.Text = "0";
            textBox3.Text = null;
            textBox19.Text = null;
            textBox20.Text = null;
        }
        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            string SDATES = dateTimePicker1.Value.ToString("yyyyMMdd");
            string MANU= comboBox1.Text.Trim();
            SEARCHMOCMANULINE_BAKING(SDATES, MANU);
        }
        private void button2_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox7.Text))
            {
                MessageBox.Show("請輸入品號！");
                return;
            }

            // 整理輸入資料
            string MANU = comboBox2.Text.Trim();
            string MANUDATE = dateTimePicker4.Value.ToString("yyyy/MM/dd");
            string MB001 = textBox7.Text.Trim();
            string MB002 = textBox10.Text.Trim();
            string MB003 = textBox11.Text.Trim();
            string CLINET = textBox9.Text.Trim();
            string MANUHOUR = textBox13.Text.Trim();
            string BOX = textBox8.Text.Trim();
            string BAR = textBox1.Text.Trim();
            string NUM = textBox12.Text.Trim();
            string PACKAGE = textBox12.Text.Trim(); // 若 PACKAGE = NUM，有點怪，可確認需求
            string OUTDATE = dateTimePicker5.Value.ToString("yyyy/MM/dd");
            string TA029 = textBox53.Text.Replace("'", ""); // 防 SQL 注入
            string HALFPRO = textBox68.Text.Trim();
            string COPTD001 = textBox42.Text.Trim();
            string COPTD002 = textBox43.Text.Trim();
            string COPTD003 = textBox72.Text.Trim();

            // 呼叫新增方法
            ADDMOCMANULINE(
                MANU,
                MANUDATE,
                MB001,
                MB002,
                MB003,
                CLINET,
                MANUHOUR,
                BOX,
                NUM,
                PACKAGE,
                OUTDATE,
                TA029,
                HALFPRO,
                COPTD001,
                COPTD002,
                COPTD003,
                BAR
            );

            // 清空欄位與重新查詢
            SETNULL();

            string SDATES = dateTimePicker1.Value.ToString("yyyyMMdd");
            string MANUS = comboBox1.Text.Trim();
            SEARCHMOCMANULINE_BAKING(SDATES, MANUS);
        }
        private void button5_Click(object sender, EventArgs e)
        {
            SETNULL();

            frmSUBMOCMANULINE SUBfrmSUBMOCMANULINE = new frmSUBMOCMANULINE();
            SUBfrmSUBMOCMANULINE.ShowDialog();
            textBox7.Text = SUBfrmSUBMOCMANULINE.TextBoxMsg;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            frmSUBMOCCOPMA SUBfrmSUBMOCCOPMA = new frmSUBMOCCOPMA();
            SUBfrmSUBMOCCOPMA.ShowDialog();
            textBox9.Text = SUBfrmSUBMOCCOPMA.TextBoxMsg;
        }
        private void button3_Click(object sender, EventArgs e)
        {
            string SDATES = dateTimePicker1.Value.ToString("yyyyMMdd");
            string MANU = comboBox1.Text.Trim();

            if (string.IsNullOrWhiteSpace(textBoxID.Text))
            {
                MessageBox.Show("請選擇要刪除的資料！");
                return;
            }

            DialogResult result = MessageBox.Show("確定要刪除這筆資料嗎？", "刪除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

            if (result == DialogResult.Yes)
            {
                DELMOCMANULINE(textBoxID.Text.Trim());
                SEARCHMOCMANULINE_BAKING(SDATES, MANU);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string SDATES = dateTimePicker1.Value.ToString("yyyyMMdd");
            string MANU = comboBox1.Text.Trim();

            CHECKMOCTAB();
            SEARCHMOCMANULINE_BAKING(SDATES, MANU);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            string TD001 = textBox42.Text;
            string TD002 = textBox43.Text;
            string TD003 = textBox72.Text;

            if (!string.IsNullOrEmpty(TD001) & !string.IsNullOrEmpty(TD002) & !string.IsNullOrEmpty(TD003))
            {
                SEARCHCOPDEFAULT(TD001, TD002, TD003);
            }
        }
        private void button8_Click(object sender, EventArgs e)
        {
            string TD001 = textBox42.Text;
            string TD002 = textBox43.Text;
            string TD003 = textBox72.Text;

            if (!string.IsNullOrEmpty(TD001) & !string.IsNullOrEmpty(TD002) & !string.IsNullOrEmpty(TD003))
            {
                SEARCHCOPDEFAULT2(TD001, TD002, TD003);
            }
        }
        private void button9_Click(object sender, EventArgs e)
        {
            string TD001 = textBox42.Text;
            string TD002 = textBox43.Text;
            string TD003 = textBox72.Text;

            if (!string.IsNullOrEmpty(TD001) & !string.IsNullOrEmpty(TD002) & !string.IsNullOrEmpty(TD003))
            {
                SEARCHCOPDEFAULT3(TD001, TD002, TD003); 
            }
        }
        private void button10_Click(object sender, EventArgs e)
        {
            string ID = textBoxID.Text.ToString().Trim();
            string TA001 = "A513";
            string TA002 = "";
            string TA020 = "";

            DateTime DT = new DateTime(); 
            if (MANU.Equals("烘焙生產線"))
            {                
                TA020 = comboBox3.SelectedValue.ToString().Trim();
            }

            if (!string.IsNullOrEmpty(TA028))
            {
                //指定日期=生產日
                DT = dt1;
                TA002 = GETMAXTA002(TA001, DT);
                if(!string.IsNullOrEmpty(TA002))
                {
                    ADDMOCMANULINERESULT(ID, TA001, TA002);
                    string MC001 = MB001B;

                    ADDMOCTATB(TA001, TA002, TA020, DT, MC001);

                    string SDATES = dateTimePicker1.Value.ToString("yyyyMMdd");
                    string MANU = comboBox1.Text.Trim();
                    SEARCHMOCMANULINE_BAKING(SDATES, MANU);

                    MessageBox.Show("完成");
                }                
               
            }
            else
            {
                MessageBox.Show("訂單沒有指定");
            }
        }
        private void button22_Click(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(DELID))
            {
                DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    DELTE_MOCMANULINERESULTBAKING(DELID, DELMOCTA001B, DELMOCTA002B);

                    string SDATES = dateTimePicker1.Value.ToString("yyyyMMdd");
                    string MANU = comboBox1.Text.Trim();
                    SEARCHMOCMANULINE_BAKING(SDATES, MANU);


                }
                else if (dialogResult == DialogResult.No)
                {
                    //do something else
                }
            }
            
        }

        private void button79_Click(object sender, EventArgs e)
        {
            ADD_MOCMANULINEBAKING_BATCH(comboBox21.Text,comboBox2.Text,dateTimePicker4.Value.ToString("yyyyMMdd"),textBox42.Text.Trim(), textBox43.Text.Trim(), textBox72.Text.Trim());

            string SDATES = dateTimePicker1.Value.ToString("yyyyMMdd");
            string MANU = comboBox1.Text.Trim();
            SEARCHMOCMANULINE_BAKING(SDATES, MANU);
        }

        private void button91_Click(object sender, EventArgs e)
        {
            SEARCHTBCOPTDCHECK(dateTimePicker28.Value.ToString("yyyyMM"), comboBox23.SelectedValue.ToString(), comboBox24.SelectedValue.ToString(), textBox97.Text.Trim());
        }

        private void button90_Click(object sender, EventArgs e)
        {
            CHECKdataGridView28();

            SEARCHTBCOPTDCHECK(dateTimePicker28.Value.ToString("yyyyMM"), comboBox23.SelectedValue.ToString(), comboBox24.SelectedValue.ToString(), textBox97.Text.Trim());
            MessageBox.Show("完成");
        }
        private void button92_Click(object sender, EventArgs e)
        {
            ADD_MOCMANULINEBAKING();
            MessageBox.Show("完成");
        }
        private void button94_Click(object sender, EventArgs e)
        {
            SEARCHTBCOPTFCHECK(dateTimePicker30.Value.ToString("yyyy"), comboBox26.SelectedValue.ToString(), textBox98.Text.Trim());
        }
        private void button95_Click(object sender, EventArgs e)
        {
            CHECKdataGridView29();

            SEARCHTBCOPTFCHECK(dateTimePicker30.Value.ToString("yyyy"), comboBox26.SelectedValue.ToString(), textBox98.Text.Trim());
            MessageBox.Show("完成");
        }
        private void button96_Click(object sender, EventArgs e)
        {
            ADDTOTKMOCMOCMANULINECOPTECOPTF(TF001, TF002, TF003, TF104);
            MessageBox.Show("完成");
        }

        private void button11_Click(object sender, EventArgs e)
        {
            string SDATES = dateTimePicker2.Value.ToString("yyyyMMdd");
            string MANU = comboBox4.Text.Trim();
            SEARCHMOCMANULINE_BAKING(SDATES, MANU);          
        }
        private void button13_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox4.Text))
            {
                string MANU = comboBox5.Text.ToString().Trim();
                string MANUDATE = dateTimePicker3.Value.ToString("yyyy/MM/dd");
                string MB001 = textBox4.Text.ToString().Trim();
                string MB002 = textBox16.Text.ToString().Trim();
                string MB003 = textBox17.Text.ToString().Trim();
                string CLINET = textBox5.Text.ToString().Trim();
                string MANUHOUR = textBox22.Text.ToString().Trim();
                string BOX = textBox6.Text.ToString().Trim();
                string BAR = textBox21.Text.ToString().Trim();
                string NUM = textBox15.Text.ToString().Trim();
                string PACKAGE = textBox15.Text.ToString().Trim();
                string OUTDATE = dateTimePicker6.Value.ToString("yyyy/MM/dd");
                string TA029 = textBox18.Text.Replace("'", "");
                string HALFPRO = textBox23.Text.ToString().Trim();
                string COPTD001 = textBox3.Text.ToString().Trim();
                string COPTD002 = textBox19.Text.ToString().Trim();
                string COPTD003 = textBox20.Text.ToString().Trim();

                ADDMOCMANULINE(
                    MANU,
                    MANUDATE,
                    MB001,
                    MB002,
                    MB003,
                    CLINET,
                    MANUHOUR,
                    BOX,
                    NUM,
                    PACKAGE,
                    OUTDATE,
                    TA029,
                    HALFPRO,
                    COPTD001,
                    COPTD002,
                    COPTD003,
                    BAR
                    );

                SETNULL();

                string SDATES = dateTimePicker2.Value.ToString("yyyyMMdd");
                string MANUS= comboBox4.Text.Trim();
                SEARCHMOCMANULINE_BAKING(SDATES, MANUS);
            }
            else
            {
                MessageBox.Show("錯誤");
            }
        }
        private void button15_Click(object sender, EventArgs e)
        {

            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELMOCMANULINE(textBoxID2.Text.ToString().Trim());

                string SDATES = dateTimePicker2.Value.ToString("yyyyMMdd");
                string MANU = comboBox4.Text.Trim();
                SEARCHMOCMANULINE_BAKING(SDATES, MANU);
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }
        private void button16_Click(object sender, EventArgs e)
        {
            CHECKMOCTAB();
            string SDATES = dateTimePicker2.Value.ToString("yyyyMMdd");
            string MANU = comboBox4.Text.Trim();
            SEARCHMOCMANULINE_BAKING(SDATES, MANU);
        }
        private void button17_Click(object sender, EventArgs e)
        {
            string TD001 = textBox3.Text;
            string TD002 = textBox19.Text;
            string TD003 = textBox20.Text;

            if (!string.IsNullOrEmpty(TD001) & !string.IsNullOrEmpty(TD002) & !string.IsNullOrEmpty(TD003))
            {
                SEARCHCOPDEFAULT(TD001, TD002, TD003);
            }
        }
        private void button18_Click(object sender, EventArgs e)
        {
            string TD001 = textBox3.Text;
            string TD002 = textBox19.Text;
            string TD003 = textBox20.Text;

            if (!string.IsNullOrEmpty(TD001) & !string.IsNullOrEmpty(TD002) & !string.IsNullOrEmpty(TD003))
            {
                SEARCHCOPDEFAULT2(TD001, TD002, TD003);
            }
        }

        private void button19_Click(object sender, EventArgs e)
        {
            string TD001 = textBox3.Text;
            string TD002 = textBox19.Text;
            string TD003 = textBox20.Text;

            if (!string.IsNullOrEmpty(TD001) & !string.IsNullOrEmpty(TD002) & !string.IsNullOrEmpty(TD003))
            {
                SEARCHCOPDEFAULT3(TD001, TD002, TD003);
            }
        }
        private void button14_Click(object sender, EventArgs e)
        {
            SETNULL();

            frmSUBMOCMANULINE SUBfrmSUBMOCMANULINE = new frmSUBMOCMANULINE();
            SUBfrmSUBMOCMANULINE.ShowDialog();
            textBox4.Text = SUBfrmSUBMOCMANULINE.TextBoxMsg;
        }

        private void button12_Click(object sender, EventArgs e)
        {
            frmSUBMOCCOPMA SUBfrmSUBMOCCOPMA = new frmSUBMOCCOPMA();
            SUBfrmSUBMOCCOPMA.ShowDialog();
            textBox5.Text = SUBfrmSUBMOCCOPMA.TextBoxMsg;
        }

        private void button23_Click(object sender, EventArgs e)
        {
            DateTime DT = new DateTime();
            if (MANU.Equals("烘焙包裝線"))
            {
                TA001 = "A513";
                TA020 = comboBox7.SelectedValue.ToString().Trim();
            }

            if (!string.IsNullOrEmpty(TA028_DV4))
            {
                //指定日期=生產日
                DT = dateTimePicker3.Value; ;
                TA002 = GETMAXTA002(TA001, DT);
                ADDMOCMANULINERESULT(textBoxID2.Text.ToString().Trim(), TA001, TA002);

                string MC001 = MB001_DV4;
                ADDMOCTATB(TA001, TA002, TA020, DT, MC001);

                string SDATES = dateTimePicker3.Value.ToString("yyyyMMdd");
                string MANU = comboBox5.Text.Trim();
                SEARCHMOCMANULINE_BAKING(SDATES, MANU);
                

                MessageBox.Show("完成");
            }
            else
            {
                MessageBox.Show("訂單沒有指定");
            }
        }

        private void button21_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(DELID_DV6))
            {
                DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {                   
                    DELTE_MOCMANULINERESULTBAKING(DELID_DV6, DELMOCTA001_DV6, DELMOCTA002_DV6);

                    string SDATES = dateTimePicker3.Value.ToString("yyyyMMdd");
                    string MANU = comboBox5.Text.Trim();
                    SEARCHMOCMANULINE_BAKING(SDATES, MANU);

                }
                else if (dialogResult == DialogResult.No)
                {
                    //do something else
                }
            }
        }



        #endregion

      
    }
}
