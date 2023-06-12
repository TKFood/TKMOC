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
    public partial class frnREPORTDEP : Form
    {
        public frnREPORTDEP()
        {
            InitializeComponent();

            comboBox1load();
        }


        #region FUNCTION

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

        public void comboBox1load()
        {
            LoadComboBoxData(comboBox1, "SELECT MD001,MD002 FROM [TK].dbo.CMSMD    WHERE MD002 IN ( '製一線','製二線')  ", "MD002", "MD002");
        }
        public void SETFASTREPORT(string TA021, string SDAY,string EDAY)
        {
            StringBuilder SQL1 = new StringBuilder();
            string CHECK_TA021 = "";           

            Report report1 = new Report();

            if (TA021.Equals("製一線"))
            {
                CHECK_TA021 = "03";
                report1.Load(@"REPORT\製造生產排程-小線.frx");
                
            }
            else  if (TA021.Equals("製二線"))
            {
                CHECK_TA021 = "02";
                report1.Load(@"REPORT\製造生產排程-大線.frx");
            }

            SQL1 = SETSQL1(CHECK_TA021, SDAY,  EDAY);
            

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            SqlConnection sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;



            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();


            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL1(string TA021, string SDAY, string EDAY)
        {
            StringBuilder SB = new StringBuilder();           
            SB.AppendFormat(@" 
                            SELECT SUBSTRING(TA003,5,2) AS '月' ,SUBSTRING(TA003,7,2) AS '日' ,TA003 AS '製令日期' ,TA001 AS '製令別',TA002 AS '製令編號',TA021 AS '生產線別',TA006 AS '品號',TA034 AS '品名',TA035 AS '規格',TA015 AS '預計產量',TA017 AS '實際產出',TA007 AS '單位',TA029 AS '備註',MB023,MB198
                            ,CASE WHEN MB198='2' THEN  CONVERT(NVARCHAR,DATEADD(DAY,-1,DATEADD(MONTH,MB023,TA003)),112) ELSE CONVERT(NVARCHAR,DATEADD(DAY,-1,DATEADD(DAY,MB023,TA003)),112) END AS '有效日期'
                            ,[ERPINVMB].[PCT] AS '比例'
                            ,[ERPINVMB].[ALLERGEN]  AS '過敏原'
                            ,[ERPINVMB].[SPEC] AS '餅體'
                            ,CONVERT(decimal(16,3),TA015/ISNULL(MC004,1)) AS '桶數'
                            ,CONVERT(decimal(16, 3), TA015 / ISNULL(MD007, 1)) AS '箱數'
                            ,MOCTA.UDF01 AS '順序'
                            ,ISNULL(MC004,1) MC004
                            ,ISNULL(MD007,1) AS MD007,ISNULL(MD010,1) AS MD010
                            ,(CASE WHEN TA021='02' THEN '大線' WHEN TA021='03' THEN '小線' END ) AS '線別'
                            ,TA021
                            FROM [TK].dbo.MOCTA
                            LEFT JOIN [TK].dbo.INVMB ON MB001=TA006
                            LEFT JOIN [TKMOC].[dbo].[ERPINVMB] ON [ERPINVMB].MB001=TA006
                            LEFT JOIN [TK].dbo.BOMMC ON MC001=TA006
                            LEFT JOIN [TK].dbo.BOMMD ON MD035 LIKE '%箱%' AND MD003 LIKE '2%' AND MD007>1 AND MD001=TA006
                            WHERE 1=1
                            AND TA034 NOT LIKE '%水麵%'
                            AND TA003>='{1}' AND TA003<='{2}' 
                            AND TA021='{0}'

                            ORDER BY TA003,REPLACE(MOCTA.UDF01,'△','')
                            ", TA021, SDAY, EDAY);

            return SB;

        }

       
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(comboBox1.Text.ToString(),dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
        }

        #endregion
    }
}
