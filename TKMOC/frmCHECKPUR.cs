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
    public partial class frmCHECKPUR : Form
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

        public frmCHECKPUR()
        {
            InitializeComponent();
        }


        #region FUNCTION
        public void Search()
        {
            DataSet ds = new DataSet();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  
                                   
                                    SELECT 
                                    表單
                                    ,ERP單別
                                    ,ERP單號
                                    ,目前簽核人員
                                    ,(TG013+TA006) AS 'ERP是否已結案'
                                    FROM 
                                    (
                                    SELECT 
                                    DOC_NBR AS '表單',
                                    TG001_Value AS 'ERP單別',
                                    TG002_Value AS 'ERP單號',
                                    STUFF(
                                            (
                                                SELECT 
                                                    ' ' + TB_EB_USER.NAME
                                                FROM 
                                                    [UOF].dbo.TB_WKF_TASK_NODE
                                                JOIN 
                                                    [UOF].dbo.TB_EB_USER ON TB_EB_USER.USER_GUID = TB_WKF_TASK_NODE.ORIGINAL_SIGNER
                                                WHERE 
                                                    TB_WKF_TASK_NODE.NODE_STATUS = 1
                                                    AND TB_WKF_TASK_NODE.TASK_ID = TEMP.TASK_ID
                                                FOR XML PATH(''), TYPE
                                            ).value('.', 'NVARCHAR(MAX)'), 1, 1, '') AS '目前簽核人員'
                                    ,ISNULL((SELECT TOP 1 TG013 FROM [192.168.1.105].[TK].dbo.PURTG WHERE TG001=TG001_Value AND TG002=TG002_Value),'') AS 'TG013'
                                    ,ISNULL((SELECT TOP 1 TA006 FROM [192.168.1.105].[TK].dbo.INVTA WHERE TA001=TG001_Value AND TA002=TG002_Value),'')  AS 'TA006'
                                    ,TASK_ID
                                    FROM 
                                    (
                                    SELECT
                                    TASK_ID,
                                    DOC_NBR,
                                    TG001_Value,
                                    TG002_Value
                                    FROM 
                                    [UOF].[dbo].[TB_WKF_TASK]
                                    CROSS APPLY
                                    (SELECT
                                    CURRENT_DOC.value('(/Form/FormFieldValue/FieldItem[@fieldId=""TG001""]/@fieldValue)[1]', 'NVARCHAR(MAX)') AS TG001_Value,
                                    CURRENT_DOC.value('(/Form/FormFieldValue/FieldItem[@fieldId=""TG002""]/@fieldValue)[1]', 'NVARCHAR(MAX)') AS TG002_Value
                                    ) AS XMLData
                                    WHERE DOC_NBR LIKE 'PURTH%'
                                    AND XMLData.TG002_Value LIKE '{0}%'
                                    UNION ALL
                                    SELECT
                                    TASK_ID,
                                    DOC_NBR,
                                    TG001_Value,
                                    TG002_Value
                                    FROM
                                    [UOF].[dbo].[TB_WKF_TASK]
                                    CROSS APPLY
                                    (SELECT
                                    CURRENT_DOC.value('(/Form/FormFieldValue/FieldItem[@fieldId=""TA001""]/@fieldValue)[1]', 'NVARCHAR(MAX)') AS TG001_Value,
                                    CURRENT_DOC.value('(/Form/FormFieldValue/FieldItem[@fieldId=""TA002""]/@fieldValue)[1]', 'NVARCHAR(MAX)') AS TG002_Value
                                    ) AS XMLData
                                    WHERE DOC_NBR LIKE 'QCINVTATB%'
                                    AND XMLData.TG002_Value LIKE '{0}%'
                                    ) AS TEMP
                                    ) AS TEMP2
                                    WHERE 1 = 1
                                    AND ERP單號 LIKE '%{0}%'
                                    AND(TG013 + TA006) = 'N'
                                    ORDER BY ERP單別,ERP單號



                                    ", textBox1.Text.Trim());

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();

                if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    dataGridView1.DataSource = ds.Tables["TEMPds1"];

                    dataGridView1.Columns["表單"].Width = 200;
                    dataGridView1.Columns["ERP單別"].Width = 100;
                    dataGridView1.Columns["ERP單號"].Width = 100;
                    dataGridView1.Columns["目前簽核人員"].Width = 300;

                }
                else
                {
                    dataGridView1.DataSource = null;
                    MessageBox.Show("沒有待簽的表單");
                }


            }
            catch
            {

            }
            finally
            {

            }
        }
        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {           

            MESSAGESHOW MSGSHOW = new MESSAGESHOW();
            // 鎖定控制項
            this.Enabled = false;
            // 顯示跳出視窗
            MSGSHOW.Show();

            // 使用非同步操作執行長時間運行的操作
            Task.Run(() =>
            {
                // 使用非同步操作執行長時間運行的操作
                //Search();

                // 更新 UI，確保在主 UI 線程上執行
                Invoke(new Action(() =>
                {
                    // 更新 UI，確保在主 UI 線程上執行
                    Search();
                    // 關閉跳出視窗
                    MSGSHOW.Close();
                    // 解除鎖定
                    this.Enabled = true;
                }));
            });

        }
        #endregion
    }
}
