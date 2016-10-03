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


namespace TKMOC
{
    public partial class frmEngineering : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();
        DataTable dt = new DataTable();
        DataTable dtTemp2 = new DataTable();
        DataTable dtTemp3 = new DataTable();
        DataTable dtMAINPARTS=new DataTable();
        DataTable dtMACHINEDAILYCHECK = new DataTable();
        DataTable dtMACHINEMAINWEEK = new DataTable();
        DataTable dtMACHINEMAINRECORD = new DataTable();
        DataTable dtMAINPARTSUSED = new DataTable();
        DataTable dtMACHINECARD = new DataTable();
        DataTable  dtMACHINEATTACH = new DataTable();
        DataGridViewRow drMAINAPPLY = new DataGridViewRow();
        DataGridViewRow drMAINAPPLYOUT = new DataGridViewRow();
        DataGridViewRow drMAINRECORD = new DataGridViewRow();

        string tablename = null;
        string EquipmentID;
        string MAINAPPLYID;
        string MAINAPPLYOUTID;
        string MAINRECORDID;
        string MACHINEID;
        string MACHINEDAILYCHECK;
        string MACHINEMAINRECORDID;
        string MACHINEMAINWEEKID;
        string MAINPARTSID;
        string MAINPARTSUSEDID;
        Thread TD;

        public frmEngineering()
        {
            InitializeComponent();
            combobox1load();
        }

        #region FUNCTION
        public void combobox1load()
        {

            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            String Sequel = "SELECT  [UNITID],[UNITNAME] FROM [TKMOC].[dbo].[ENDUNIT]";
            SqlDataAdapter da = new SqlDataAdapter(Sequel, sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("UNITID", typeof(string));
            dt.Columns.Add("UNITNAME", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "UNITID";
            comboBox1.DisplayMember = "UNITNAME";
            sqlConn.Close();           

        }
        public void Search()
        {
            StringBuilder Query = new StringBuilder();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                if(!string.IsNullOrEmpty(comboBox1.SelectedValue.ToString()))
                {
                    if(!(comboBox1.SelectedValue.ToString().Equals("0")))
                    {
                        Query.AppendFormat(" AND [ENDUNIT].UNITID='{0}' ", comboBox1.SelectedValue.ToString());
                    }
                }
                if(!string.IsNullOrEmpty(textBox1.Text.ToString()))
                {
                    Query.AppendFormat(" AND  [ENGEQUIPMENT].ID='{0}' ", textBox1.Text.ToString());
                }
         
                sbSql.Append(@" SELECT [ID] AS '設備編號',[NAME]  AS '設備名稱',[UNITNAME]  AS '單位',[FACTORY]  AS '廠牌',[TYPE]  AS '型別',[MAINTENANCE]  AS '保養',[CHEKCK]  AS '點檢',[STATUS]  AS '狀況說明'   ");
                sbSql.Append(@" FROM [TKMOC].[dbo].[ENGEQUIPMENT] WITH (NOLOCK),[TKMOC].[dbo].[ENDUNIT] WITH (NOLOCK)");
                sbSql.Append(@" WHERE [ENGEQUIPMENT].UNIT=[ENDUNIT].UNITID");
                sbSql.AppendFormat("  {0} ",Query.ToString());
                sbSql.Append(@" ORDER BY [ID] ");
                sbSql.Append(@" ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {
                    label1.Text = "找不到資料";
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        label1.Text = "有 " + ds.Tables["TEMPds1"].Rows.Count.ToString() + " 筆";
                        //dataGridView1.Rows.Clear();
                        dataGridView1.DataSource = ds.Tables["TEMPds1"];
                        dataGridView1.AutoResizeColumns();

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

        public void ExcelExport()
        {

            string NowDB = "TK";
            //建立Excel 2003檔案
            IWorkbook wb = new XSSFWorkbook();
            ISheet ws;

            XSSFCellStyle cs = (XSSFCellStyle)wb.CreateCellStyle();
            //框線樣式及顏色
            cs.BorderBottom = NPOI.SS.UserModel.BorderStyle.Double;
            cs.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            cs.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            cs.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            cs.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;
            cs.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;
            cs.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;
            cs.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;

            //Search();            
            dt = ds.Tables["TEMPds1"];

            if (dt.TableName != string.Empty)
            {
                ws = wb.CreateSheet(dt.TableName);
            }
            else
            {
                ws = wb.CreateSheet("Sheet1");
            }

            ws.CreateRow(0);//第一行為欄位名稱
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ws.GetRow(0).CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);
            }

            int j = 0;
            int k = dataGridView1.Rows.Count;
            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {
                ws.CreateRow(j + 1);
                ws.GetRow(j + 1).CreateCell(0).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString());
                ws.GetRow(j + 1).CreateCell(1).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());
                ws.GetRow(j + 1).CreateCell(2).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString());
                ws.GetRow(j + 1).CreateCell(3).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString());
                ws.GetRow(j + 1).CreateCell(4).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[4].ToString());
                ws.GetRow(j + 1).CreateCell(5).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[5].ToString());
                ws.GetRow(j + 1).CreateCell(6).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[6].ToString());
                ws.GetRow(j + 1).CreateCell(7).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[7].ToString());
                //ws.GetRow(j + 1).CreateCell(3).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString()));
                j++;
            }

            if (Directory.Exists(@"c:\temp\"))
            {
                //資料夾存在
            }
            else
            {
                //新增資料夾
                Directory.CreateDirectory(@"c:\temp\");
            }
            StringBuilder filename = new StringBuilder();
            filename.AppendFormat(@"c:\temp\機械堪用月記錄表{0}.xlsx", DateTime.Now.ToString("yyyyMMdd"));

            FileStream file = new FileStream(filename.ToString(), FileMode.Create);//產生檔案
            wb.Write(file);
            file.Close();

            MessageBox.Show("匯出完成-EXCEL放在-" + filename.ToString());
            FileInfo fi = new FileInfo(filename.ToString());
            if (fi.Exists)
            {
                System.Diagnostics.Process.Start(filename.ToString());
            }
            else
            {
                //file doesn't exist
            }


        }
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if(dataGridView1.Rows.Count>=1)
            {
                EquipmentID = dataGridView1.CurrentRow.Cells["設備編號"].Value.ToString();
            }
            
        }

        public void SearchMAINAPPLY()
        {
            DataSet ds = new DataSet();
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();               

                sbSql.Append(@" SELECT [APPLYUNIT] AS '申請單位',[APPDATE] AS '申請日期',[EQUIPMENTID] AS '機台編號' ,[EQUIPMENTNAME] AS '設備名稱',[FINDEMP] AS '發現者',[APPLYEMP] AS '申請人' ,[ERROR] AS '異常情形',[STATUS] AS '原因及處理方式',[REMARK] AS '備註',[MAINEMP] AS '維修者',[MAINDATE] AS '維修時間',[ID]  ");
                sbSql.Append(@" FROM [TKMOC].[dbo].[MAINAPPLY] ");
                sbSql.AppendFormat(@" WHERE [EQUIPMENTID] ='{0}'",EquipmentID.ToString());
                sbSql.Append(@" ORDER BY [APPDATE] DESC");
                sbSql.Append(@" ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds");
                sqlConn.Close();


                if (ds.Tables["TEMPds"].Rows.Count == 0)
                {
                    label1.Text = "找不到資料";
                    dataGridView2.DataSource = null;
                }
                else
                {
                    if (ds.Tables["TEMPds"].Rows.Count >= 1)
                    {                                              
                        dataGridView2.DataSource = ds.Tables["TEMPds"];
                        dataGridView2.AutoResizeColumns();

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

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView2.Rows.Count >= 1)
            {
                MAINAPPLYID = dataGridView2.CurrentRow.Cells["ID"].Value.ToString();
                drMAINAPPLY = dataGridView2.Rows[dataGridView2.SelectedCells[0].RowIndex];
            }
        }

        public void SearchMAINAPPLYOUT()
        {
            DataSet ds = new DataSet();
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.Append(@" SELECT  [APPLYUNIT] AS '申請單位',[EQUIPDATE] AS '出廠日期',[EQUIPMENTID] AS '設備編號',[EQUIPMENTNAME] AS '設備名稱',[APPLYEMP] AS '申請人',[ERROR] AS '異常情形',[STATUS] AS '原因及處理方式',[FACTROY] AS '維修廠商',[RETURNDATE] AS '預定回廠日',[RECEIVEDATE] AS '接收日' ,[ID]  ");
                sbSql.Append(@" FROM [TKMOC].[dbo].[MAINAPPLYOUT] ");
                sbSql.AppendFormat(@" WHERE [EQUIPMENTID] ='{0}'", EquipmentID.ToString());
                sbSql.Append(@" ORDER BY [RECEIVEDATE] DESC");
                sbSql.Append(@" ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds");
                sqlConn.Close();


                if (ds.Tables["TEMPds"].Rows.Count == 0)
                {
                    label1.Text = "找不到資料";
                    dataGridView3.DataSource = null; 
                }
                else
                {
                    if (ds.Tables["TEMPds"].Rows.Count >= 1)
                    {                        
                        dataGridView3.DataSource = ds.Tables["TEMPds"];
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

        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView3.Rows.Count >= 1)
            {
                MAINAPPLYOUTID = dataGridView3.CurrentRow.Cells["ID"].Value.ToString();
                drMAINAPPLYOUT= dataGridView3.Rows[dataGridView3.SelectedCells[0].RowIndex];
            }
        }

        public void SearchMAINRECORD()
        {
            DataSet ds = new DataSet();
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.Append(@" SELECT [EQUIPMENTID] AS '財產編號',[EQUIPMENTNAME] AS '設備名稱',[UNIT] AS '使用部門',[ERROR] AS '故障情形',[MAINDATEBEGIN] AS '維修時間起',[MAINDATEEND] AS '維修時間迄',[MAINDATHR] AS '維修時數',[MAINEMP] AS '維修人員',[MALFUNCIONID] AS '故障性質',[MAINSTATUS] AS '維修內容',[MAINUSED] AS '本次更換' ,[ID]  ");
                sbSql.Append(@" FROM [TKMOC].[dbo].[MAINRECORD] ");
                sbSql.AppendFormat(@" WHERE [EQUIPMENTID] ='{0}'", EquipmentID.ToString());
                sbSql.Append(@" ORDER BY [ID] DESC");
                sbSql.Append(@" ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds");
                sqlConn.Close();


                if (ds.Tables["TEMPds"].Rows.Count == 0)
                {
                    label1.Text = "找不到資料";
                    dataGridView4.DataSource = null;
                }
                else
                {
                    if (ds.Tables["TEMPds"].Rows.Count >= 1)
                    {
                        dataGridView4.DataSource = ds.Tables["TEMPds"];
                        dataGridView4.AutoResizeColumns();

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
        private void dataGridView4_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView4.Rows.Count >= 1)
            {
                MAINRECORDID = dataGridView4.CurrentRow.Cells["ID"].Value.ToString();
                drMAINRECORD = dataGridView4.Rows[dataGridView4.SelectedCells[0].RowIndex];
            }
        }

        public void SearchMCHINE()
        {
            DataSet ds = new DataSet();
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.Append(@" SELECT [EQUIPMENTID] AS '設備編號',[EQUIPMENTNAME] AS '設備名稱',[VALUE] AS '價值',[TYPE] AS '型號',[WEIGHT] AS '重量',[MACHINECODE] AS '機械製造號碼',[SIZE] AS '外形尺寸',[FACTORY] AS '製造廠商',[MACHINEID] AS '機器編號',[SELLFACTORY] AS '出售廠商',[MACHYEAR] AS '製造年份',[UNIT] AS '使用單位',[BUYDATE] AS '購入日期',[OWNER] AS '保管人',[STATUS] AS '重要規格'  ,USEWATER AS '用水tom/hr',USEPOWER AS '電力kW',USEAIR AS '空氣m3/min' ,MANAGER AS '主管' ,CREATOR  AS '建卡人',CREATEDATE  AS '建卡日期'  ");
                sbSql.Append(@" FROM [TKMOC].[dbo].[MACHINECARD] ");
                sbSql.AppendFormat(@" WHERE [EQUIPMENTID] ='{0}'", EquipmentID.ToString());
                sbSql.Append(@" ORDER BY [EQUIPMENTID] DESC ");
                sbSql.Append(@" ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds");
                sqlConn.Close();


                if (ds.Tables["TEMPds"].Rows.Count == 0)
                {
                    label1.Text = "找不到資料";
                    dataGridView5.DataSource = null;
                }
                else
                {
                    if (ds.Tables["TEMPds"].Rows.Count >= 1)
                    {
                        dataGridView5.DataSource = ds.Tables["TEMPds"];
                        dataGridView5.AutoResizeColumns();
                        dtMACHINECARD = ds.Tables["TEMPds"];
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
        private void dataGridView5_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView5.Rows.Count >= 1)
            {
                MACHINEID = dataGridView5.CurrentRow.Cells["設備編號"].Value.ToString();
            }
            
        }
        public void SearchMCHINEATTACH()
        {
            DataSet ds = new DataSet();
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.Append(@" SELECT [EQUIPMENTID] AS '設備編號',[EQUIPMENTNAME] AS '設備名稱',[ATTCHANAME] AS '附件名稱',[SPEC] AS '規格',[NUM] AS '數量',[ID] ");
                sbSql.Append(@" FROM [TKMOC].[dbo].[MACHINEATTACH] ");
                sbSql.AppendFormat(@" WHERE [EQUIPMENTID] ='{0}'", EquipmentID.ToString());
                sbSql.Append(@" ORDER BY [ID] DESC ");
                sbSql.Append(@" ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds");
                sqlConn.Close();


                if (ds.Tables["TEMPds"].Rows.Count == 0)
                {
                    label1.Text = "找不到資料";
                    dataGridView6.DataSource = null;
                }
                else
                {
                    if (ds.Tables["TEMPds"].Rows.Count >= 1)
                    {
                        dataGridView6.DataSource = ds.Tables["TEMPds"];
                        dataGridView6.AutoResizeColumns();
                        dtMACHINEATTACH = ds.Tables["TEMPds"];

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
        public void SearchMACHINEDAILYCHECK()
        {
            DataSet ds = new DataSet();
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.Append(@" SELECT [EQUIPMENTID] AS '設備編號',[EQUIPMENTNAME] AS '設備名稱',[UNIT] AS '使用單位',[MAINDATE] AS '保養日期',[CHECK1] AS '工作前機台內外部的清理消毒',[CHECK2] AS '各部螺絲確實鎖緊',[CHECK3] AS '各操作按鍵鈕正常無異',[CHECK4] AS '機台運行順暢無異常',[CHECK5] AS '各設定確實依作業標準書',[CHECK6] AS '機器運行正常無異聲',[CHECK7] AS '零件使用後確實清潔消毒',[CHECK8] AS '各指示燈確實亮起無異',[CHECK9] AS '各設定溫度時間確實達到',[CHECK10] AS '零件安裝固定完全',[CHECK11] AS '工作後機台內外部清潔消毒',[CHECKOR] AS ' 檢查者',[ID] ");
                sbSql.Append(@" FROM [TKMOC].[dbo].[MACHINEDAILYCHECK] ");
                sbSql.AppendFormat(@" WHERE [EQUIPMENTID] ='{0}'", EquipmentID.ToString());
                sbSql.Append(@" ORDER BY [ID] DESC ");
                sbSql.Append(@" ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds");
                sqlConn.Close();


                if (ds.Tables["TEMPds"].Rows.Count == 0)
                {
                    label1.Text = "找不到資料";
                    dataGridView7.DataSource = null;
                }
                else
                {
                    if (ds.Tables["TEMPds"].Rows.Count >= 1)
                    {
                        dataGridView7.DataSource = ds.Tables["TEMPds"];
                        dataGridView7.AutoResizeColumns();
                        dtMACHINEDAILYCHECK = ds.Tables["TEMPds"];

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
        private void dataGridView7_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView7.Rows.Count >= 1)
            {
                MACHINEDAILYCHECK = dataGridView7.CurrentRow.Cells["ID"].Value.ToString();
            }
        }

        public void SearchMACHINEMAINRECORD()
        {
            DataSet ds = new DataSet();
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.Append(@"  SELECT [EQUIPMENTID] AS '設備編號',[EQUIPMENTNAME] AS '設備名稱',[UNIT] AS '使用單位',[MAINDATE] AS '保養日期',[MAINEMP] AS '保養者',[STATUS] AS '原因及處理情形',[CHECKER] AS '審查者',[ID] ");
                sbSql.Append(@"  FROM [TKMOC].[dbo].[MACHINEMAINRECORD] ");
                sbSql.AppendFormat(@" WHERE [EQUIPMENTID] ='{0}'", EquipmentID.ToString());
                sbSql.Append(@" ORDER BY [ID] DESC ");
                sbSql.Append(@" ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds");
                sqlConn.Close();


                if (ds.Tables["TEMPds"].Rows.Count == 0)
                {
                    label1.Text = "找不到資料";
                    dataGridView8.DataSource = null;
                }
                else
                {
                    if (ds.Tables["TEMPds"].Rows.Count >= 1)
                    {
                        dataGridView8.DataSource = ds.Tables["TEMPds"];
                        dataGridView8.AutoResizeColumns();
                        dtMACHINEMAINRECORD = ds.Tables["TEMPds"];

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
        private void dataGridView8_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView8.Rows.Count >= 1)
            {
                MACHINEMAINRECORDID = dataGridView8.CurrentRow.Cells["ID"].Value.ToString();
            }
        }

        public void SearchMACHINEMAINWEEK()
        {
            DataSet ds = new DataSet();
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.Append(@"  SELECT [EQUIPMENTID] AS '設備編號',[EQUIPMENTNAME]  AS '設備名稱',[UNIT] AS '使用單位',[MAINDATE]  AS '保養日',[MAINYEAR] AS '保養年',[MAINMONTH] AS '保養月',[MAINWEEK] AS '保養週次',[ISMAIN] AS '是否保養' ,[ID]");
                sbSql.Append(@"  FROM [TKMOC].[dbo].[MACHINEMAINWEEK]  ");
                sbSql.AppendFormat(@" WHERE [EQUIPMENTID] ='{0}'", EquipmentID.ToString());
                sbSql.Append(@" ORDER BY [ID] DESC ");
                sbSql.Append(@" ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds");
                sqlConn.Close();


                if (ds.Tables["TEMPds"].Rows.Count == 0)
                {
                    label1.Text = "找不到資料";
                    dataGridView9.DataSource = null;
                }
                else
                {
                    if (ds.Tables["TEMPds"].Rows.Count >= 1)
                    {
                        dataGridView9.DataSource = ds.Tables["TEMPds"];
                        dataGridView9.AutoResizeColumns();
                        dtMACHINEMAINWEEK = ds.Tables["TEMPds"];

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
        private void dataGridView9_SelectionChanged(object sender, EventArgs e)
        {           
            if (dataGridView9.Rows.Count >= 1)
            {
                MACHINEMAINWEEKID = dataGridView9.CurrentRow.Cells["ID"].Value.ToString();
            }
        }
        public void SearchMAINPARTS()
        {
            DataSet ds = new DataSet();
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.Append(@"  SELECT [EQUIPMENTID] AS '設備編號',[EQUIPMENTNAME] AS '設備名稱',[PARTSNO] AS '備品編號',[PARTSNAME] AS '品名',[PARTSSPEC] AS '規格',CAST([PARTSPRICE] AS  DECIMAL(16,2) ) AS '單價',[PARTSFACTORY] AS '供應商',[TEL] AS '電話',[YEARS] AS '使用壽命',[STOCKNUM] AS '安全庫存',[PRETIME] AS '前置時間',[REMARK] AS '備註' ,[ID]");
                sbSql.Append(@"  FROM [TKMOC].[dbo].[MAINPARTS]  ");
                sbSql.AppendFormat(@" WHERE [EQUIPMENTID] ='{0}'", EquipmentID.ToString());
                sbSql.Append(@" ORDER BY [ID] DESC ");
                sbSql.Append(@" ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds");
                sqlConn.Close();


                if (ds.Tables["TEMPds"].Rows.Count == 0)
                {
                    label1.Text = "找不到資料";
                    dataGridView10.DataSource = null;
                }
                else
                {
                    if (ds.Tables["TEMPds"].Rows.Count >= 1)
                    {
                        dataGridView10.DataSource = ds.Tables["TEMPds"];
                        dataGridView10.AutoResizeColumns();
                        dtMAINPARTS = ds.Tables["TEMPds"];

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
        private void dataGridView10_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView10.Rows.Count >= 1)
            {
                MAINPARTSID = dataGridView10.CurrentRow.Cells["ID"].Value.ToString();
            }
        }

        public void SearchMAINPARTSUSED()
        {
            DataSet ds = new DataSet();
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.Append(@"  SELECT [EQUIPMENTID] AS '設備編號',[EQUIPMENTNAME]  AS '設備名稱',[PARTSNO]  AS '備品編號',[PARTSNAME] AS '品名',[PARTSSPEC]  AS '規格',CAST([PARTSPRICE] AS  DECIMAL(16,2) )AS '單價',[PARTSFACTORY]  AS '供應商',[TEL] AS '電話',[YEARS] AS '使用壽命',[STOCKNUM] AS '安全庫存',[NOWNUM] AS '現有庫存',[USEDDATE] AS '入/領用日',[INUM] AS '入庫數'  ,[USEDNUM] AS '領用數',[ID]");
                sbSql.Append(@"  FROM [TKMOC].[dbo].[MAINPARTSUSED]  ");
                sbSql.AppendFormat(@" WHERE [EQUIPMENTID] ='{0}'", EquipmentID.ToString());
                sbSql.Append(@" ORDER BY [ID] DESC ");
                sbSql.Append(@" ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds");
                sqlConn.Close();


                if (ds.Tables["TEMPds"].Rows.Count == 0)
                {
                    label1.Text = "找不到資料";
                    dataGridView11.DataSource = null;
                }
                else
                {
                    if (ds.Tables["TEMPds"].Rows.Count >= 1)
                    {
                        dataGridView11.DataSource = ds.Tables["TEMPds"];
                        dataGridView11.AutoResizeColumns();
                        dtMAINPARTSUSED = ds.Tables["TEMPds"];

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
        private void dataGridView11_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView10.Rows.Count >= 1)
            {
                MAINPARTSUSEDID = dataGridView11.CurrentRow.Cells["ID"].Value.ToString();
            }
        }

        public void ExcelExportMAINPARTS()
        {

            string NowDB = "TK";
            //建立Excel 2003檔案
            IWorkbook wb = new XSSFWorkbook();
            ISheet ws;

            XSSFCellStyle cs = (XSSFCellStyle)wb.CreateCellStyle();
            //框線樣式及顏色
            cs.BorderBottom = NPOI.SS.UserModel.BorderStyle.Double;
            cs.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            cs.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            cs.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            cs.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;
            cs.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;
            cs.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;
            cs.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;

            //Search();            
            dt = dtMAINPARTS;

            if (dt.TableName != string.Empty)
            {
                ws = wb.CreateSheet(dt.TableName);
            }
            else
            {
                ws = wb.CreateSheet("Sheet1");
            }

            ws.CreateRow(0);//第一行為欄位名稱
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ws.GetRow(0).CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);
            }

            int j = 0;
            int k = dataGridView10.Rows.Count;
            foreach (DataGridViewRow dr in this.dataGridView10.Rows)
            {
                ws.CreateRow(j + 1);
                ws.GetRow(j + 1).CreateCell(0).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString());
                ws.GetRow(j + 1).CreateCell(1).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());
                ws.GetRow(j + 1).CreateCell(2).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString());
                ws.GetRow(j + 1).CreateCell(3).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString());
                ws.GetRow(j + 1).CreateCell(4).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[4].ToString());
                ws.GetRow(j + 1).CreateCell(5).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[5].ToString());
                ws.GetRow(j + 1).CreateCell(6).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[6].ToString());
                ws.GetRow(j + 1).CreateCell(7).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[7].ToString());
                ws.GetRow(j + 1).CreateCell(8).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[8].ToString());
                ws.GetRow(j + 1).CreateCell(9).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[9].ToString());
                ws.GetRow(j + 1).CreateCell(10).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[10].ToString());
                ws.GetRow(j + 1).CreateCell(11).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[11].ToString());
                ws.GetRow(j + 1).CreateCell(12).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[12].ToString());
               
                //ws.GetRow(j + 1).CreateCell(3).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString()));
                j++;
            }

            if (Directory.Exists(@"c:\temp\"))
            {
                //資料夾存在
            }
            else
            {
                //新增資料夾
                Directory.CreateDirectory(@"c:\temp\");
            }
            StringBuilder filename = new StringBuilder();
            filename.AppendFormat(@"c:\temp\備品一覽表{0}.xlsx", DateTime.Now.ToString("yyyyMMdd"));

            FileStream file = new FileStream(filename.ToString(), FileMode.Create);//產生檔案
            wb.Write(file);
            file.Close();

            MessageBox.Show("匯出完成-EXCEL放在-" + filename.ToString());
            FileInfo fi = new FileInfo(filename.ToString());
            if (fi.Exists)
            {
                System.Diagnostics.Process.Start(filename.ToString());
            }
            else
            {
                //file doesn't exist
            }


        }
        public void PRINTMAINAPPLY()
        {
            // 首先把建立的範本檔案讀入MemoryStream
            //首先把建立的範本檔案讀入MemoryStream
            System.IO.MemoryStream _memoryStream = new System.IO.MemoryStream(Properties.Resources.維修申請單);

            //建立一個Document物件
            //並傳入MemoryStream
            Aspose.Words.Document doc = new Aspose.Words.Document(_memoryStream);

            //新增一個DataTable
            DataTable table = new DataTable();
            //建立Column
            table.Columns.Add("UNIT");
            table.Columns.Add("APPDATE");
            table.Columns.Add("TIMES");
            table.Columns.Add("EQUIPMENTID");
            table.Columns.Add("EQUIPMENTNAME");
            table.Columns.Add("FINDEMP");
            table.Columns.Add("APPLYEMP");
            table.Columns.Add("ERROR");
            table.Columns.Add("STATUS");
            table.Columns.Add("REMARK");
            table.Columns.Add("MAINEMP");
            table.Columns.Add("MAINDATE");

            //[APPLYUNIT] AS '申請單位',[APPDATE] AS '申請日期',[EQUIPMENTID] AS '機台編號' 
            //,[EQUIPMENTNAME] AS '設備名稱',[FINDEMP] AS '發現者',[APPLYEMP] AS '申請人' ,[ERROR] AS '異常情形'
            //,[STATUS] AS '原因及處理方式',[REMARK] AS '備註',[MAINEMP] AS '維修者',[MAINDATE] AS '維修時間'
            //透過建立的DataTable物件來New一個儲存資料的Row
            DataRow row = table.NewRow();
            //這些Row具有上面所建立相同的Column欄位
            //因此可以直接指定欄位名稱將資料填入裡面       
            DateTime dt = Convert.ToDateTime(drMAINAPPLY.Cells["申請日期"].Value.ToString());
            row["UNIT"] = FindUNIT(drMAINAPPLY.Cells["申請單位"].Value.ToString());
            //row["APPDATE"] = dt.Year.ToString() + "年" + dt.Month.ToString() + "月" + dt.Day.ToString() + "日";
            row["APPDATE"] = dt.ToString("yyyy/MM/dd");
            row["APPDATE"] = dt.ToString("yyyy/MM/dd");
            row["TIMES"] = dt.Hour.ToString() + ":" + dt.Minute.ToString();
            row["EQUIPMENTID"] = drMAINAPPLY.Cells["機台編號"].Value.ToString();
            row["EQUIPMENTNAME"] = drMAINAPPLY.Cells["設備名稱"].Value.ToString();
            row["FINDEMP"] = drMAINAPPLY.Cells["發現者"].Value.ToString();
            row["APPLYEMP"] = drMAINAPPLY.Cells["申請人"].Value.ToString();
            row["ERROR"] = drMAINAPPLY.Cells["異常情形"].Value.ToString();
            row["STATUS"] = drMAINAPPLY.Cells["原因及處理方式"].Value.ToString();
            row["REMARK"] = drMAINAPPLY.Cells["備註"].Value.ToString();
            row["MAINEMP"] = drMAINAPPLY.Cells["維修者"].Value.ToString();
            row["MAINDATE"] = drMAINAPPLY.Cells["維修時間"].Value.ToString();


            //把所建立的資料行加入Table的Row清單內
            table.Rows.Add(row);


            //將DataTable傳入Document的MailMerge.Execute()方法
            doc.MailMerge.Execute(table);
            //清空所有未被合併的功能變數
            doc.MailMerge.DeleteFields();

            if (Directory.Exists(@"c:\temp\"))
            {
                //資料夾存在
            }
            else
            {
                //新增資料夾
                Directory.CreateDirectory(@"c:\temp\");
            }
            //將檔案儲存至c:\
            StringBuilder filename = new StringBuilder();
            filename.AppendFormat(@"c:\temp\維修申請單{0}.doc", DateTime.Now.ToString("yyyyMMdd"));
            doc.Save(filename.ToString());

            MessageBox.Show("匯出完成-文件放在-" + filename.ToString());
            FileInfo fi = new FileInfo(filename.ToString());
            if (fi.Exists)
            {
                System.Diagnostics.Process.Start(filename.ToString());
            }
            else
            {
                //file doesn't exist
            }

        }

        public string FindUNIT(string UNITID)
        {

            DataSet ds = new DataSet();
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@" SELECT  [UNITID],[UNITNAME] FROM [TKMOC].[dbo].[ENDUNIT] WHERE   [UNITID] ='{0}'", UNITID.ToString());
                
                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds");
                sqlConn.Close();


                if (ds.Tables["TEMPds"].Rows.Count == 0)
                {
                    return "";  
                }
                else
                {
                    if (ds.Tables["TEMPds"].Rows.Count >= 1)
                    {
                       return  ds.Tables["TEMPds"].Rows[0]["UNITNAME"].ToString(); 
                    }
                    return "";
                }

            }
            catch
            {
                return "";
            }
            finally
            {
                
            }
            
        }

        public void PRINTMAINAPPLYOUT()
        {
            // 首先把建立的範本檔案讀入MemoryStream
            //首先把建立的範本檔案讀入MemoryStream
            System.IO.MemoryStream _memoryStream = new System.IO.MemoryStream(Properties.Resources.委外維修申請單);

            //建立一個Document物件
            //並傳入MemoryStream
            Aspose.Words.Document doc = new Aspose.Words.Document(_memoryStream);

            //新增一個DataTable
            DataTable table = new DataTable();
            //建立Column
            table.Columns.Add("APPLYUNIT");
            table.Columns.Add("EQUIPDATE");
            table.Columns.Add("APPLYEMP");
            table.Columns.Add("EQUIPMENTID");
            table.Columns.Add("EQUIPMENTNAME");
            table.Columns.Add("ERROR");
            table.Columns.Add("STATUS");
            table.Columns.Add("FACTROY");
            table.Columns.Add("RETURNDATE");
            table.Columns.Add("RECEIVEDATE");


            //[APPLYUNIT] AS '申請單位',[EQUIPDATE] AS '出廠日期',[EQUIPMENTID] AS '設備編號'
            //,[EQUIPMENTNAME] AS '設備名稱',[APPLYEMP] AS '申請人',[ERROR] AS '異常情形'
            //,[STATUS] AS '原因及處理方式',[FACTROY] AS '維修廠商',[RETURNDATE] AS '預定回廠日'
            //,[RECEIVEDATE] AS '接收日'
            //透過建立的DataTable物件來New一個儲存資料的Row
            DataRow row = table.NewRow();
            //這些Row具有上面所建立相同的Column欄位
            //因此可以直接指定欄位名稱將資料填入裡面       
            DateTime dt = Convert.ToDateTime(drMAINAPPLYOUT.Cells["出廠日期"].Value.ToString());
            DateTime dt2 = Convert.ToDateTime(drMAINAPPLYOUT.Cells["預定回廠日"].Value.ToString());
            DateTime dt3 = Convert.ToDateTime(drMAINAPPLYOUT.Cells["接收日"].Value.ToString());
            row["APPLYUNIT"] = FindUNIT(drMAINAPPLYOUT.Cells["申請單位"].Value.ToString());
            //row["APPDATE"] = dt.Year.ToString() + "年" + dt.Month.ToString() + "月" + dt.Day.ToString() + "日";
            row["EQUIPDATE"] = dt.ToString("yyyy/MM/dd");
            row["APPLYEMP"] = drMAINAPPLYOUT.Cells["申請人"].Value.ToString();
            row["EQUIPMENTID"] = drMAINAPPLYOUT.Cells["設備編號"].Value.ToString();
            row["EQUIPMENTNAME"] = drMAINAPPLYOUT.Cells["設備名稱"].Value.ToString();
            row["ERROR"] = drMAINAPPLYOUT.Cells["異常情形"].Value.ToString();
            row["STATUS"] = drMAINAPPLYOUT.Cells["原因及處理方式"].Value.ToString();
            row["FACTROY"] = drMAINAPPLYOUT.Cells["維修廠商"].Value.ToString();
            row["RETURNDATE"] = dt2.ToString("yyyy/MM/dd");
            row["RECEIVEDATE"] = dt3.ToString("yyyy/MM/dd");


            //把所建立的資料行加入Table的Row清單內
            table.Rows.Add(row);


            //將DataTable傳入Document的MailMerge.Execute()方法
            doc.MailMerge.Execute(table);
            //清空所有未被合併的功能變數
            doc.MailMerge.DeleteFields();

            if (Directory.Exists(@"c:\temp\"))
            {
                //資料夾存在
            }
            else
            {
                //新增資料夾
                Directory.CreateDirectory(@"c:\temp\");
            }
            //將檔案儲存至c:\
            StringBuilder filename = new StringBuilder();
            filename.AppendFormat(@"c:\temp\委外維修申請單{0}.doc", DateTime.Now.ToString("yyyyMMdd"));
            doc.Save(filename.ToString());

            MessageBox.Show("匯出完成-文件放在-" + filename.ToString());
            FileInfo fi = new FileInfo(filename.ToString());
            if (fi.Exists)
            {
                System.Diagnostics.Process.Start(filename.ToString());
            }
            else
            {
                //file doesn't exist
            }

        }

        public void PRINTMAINRECORD()
        {
            // 首先把建立的範本檔案讀入MemoryStream
            //首先把建立的範本檔案讀入MemoryStream
            System.IO.MemoryStream _memoryStream = new System.IO.MemoryStream(Properties.Resources.機械設備維修紀錄表);

            //建立一個Document物件
            //並傳入MemoryStream
            Aspose.Words.Document doc = new Aspose.Words.Document(_memoryStream);

            //新增一個DataTable
            DataTable table = new DataTable();
            //建立Column
            table.Columns.Add("EQUIPMENTNAME");
            table.Columns.Add("EQUIPMENTID");
            table.Columns.Add("UNIT");
            table.Columns.Add("ERROR");
            table.Columns.Add("MAINDATEBEGIN");
            table.Columns.Add("MAINDATEEND");
            table.Columns.Add("MAINDATHR");
            table.Columns.Add("MAINEMP");
            table.Columns.Add("MALFUNCIONID");
            table.Columns.Add("MAINSTATUS");
            table.Columns.Add("MAINUSED");


            //SELECT [EQUIPMENTID] AS '財產編號',[EQUIPMENTNAME] AS '設備名稱',[UNIT] AS '使用部門'
            //,[ERROR] AS '故障情形',[MAINDATEBEGIN] AS '維修時間起',[MAINDATEEND] AS '維修時間迄'
            //,[MAINDATHR] AS '維修時數',[MAINEMP] AS '維修人員',[MALFUNCIONID] AS '故障性質
            //',[MAINSTATUS] AS '維修內容',[MAINUSED] AS '本次更換'
            //透過建立的DataTable物件來New一個儲存資料的Row
            DataRow row = table.NewRow();
            //這些Row具有上面所建立相同的Column欄位
            //因此可以直接指定欄位名稱將資料填入裡面       
            DateTime dt = Convert.ToDateTime(drMAINRECORD.Cells["維修時間起"].Value.ToString());
            DateTime dt2 = Convert.ToDateTime(drMAINRECORD.Cells["維修時間迄"].Value.ToString());
            row["EQUIPMENTNAME"] = drMAINRECORD.Cells["設備名稱"].Value.ToString();
            //row["APPDATE"] = dt.Year.ToString() + "年" + dt.Month.ToString() + "月" + dt.Day.ToString() + "日";
            row["EQUIPMENTID"] = drMAINRECORD.Cells["財產編號"].Value.ToString(); 
            row["UNIT"] = FindUNIT(drMAINRECORD.Cells["使用部門"].Value.ToString());
            row["ERROR"] = drMAINRECORD.Cells["故障情形"].Value.ToString();
            row["MAINDATEBEGIN"] = dt.ToString("yyyy/MM/dd hh:mm");
            row["MAINDATEEND"] = dt2.ToString("yyyy/MM/dd hh:mm");
            row["MAINDATHR"] = drMAINRECORD.Cells["維修時數"].Value.ToString();
            row["MAINEMP"] = drMAINRECORD.Cells["維修人員"].Value.ToString();
            row["MALFUNCIONID"] = FindMALFUNCION(drMAINRECORD.Cells["故障性質"].Value.ToString());
            row["MAINSTATUS"] = drMAINRECORD.Cells["維修內容"].Value.ToString();
            row["MAINUSED"] = drMAINRECORD.Cells["本次更換"].Value.ToString();



            //把所建立的資料行加入Table的Row清單內
            table.Rows.Add(row);


            //將DataTable傳入Document的MailMerge.Execute()方法
            doc.MailMerge.Execute(table);
            //清空所有未被合併的功能變數
            doc.MailMerge.DeleteFields();

            if (Directory.Exists(@"c:\temp\"))
            {
                //資料夾存在
            }
            else
            {
                //新增資料夾
                Directory.CreateDirectory(@"c:\temp\");
            }
            //將檔案儲存至c:\
            StringBuilder filename = new StringBuilder();
            filename.AppendFormat(@"c:\temp\機械設備維修紀錄表{0}.doc", DateTime.Now.ToString("yyyyMMdd"));
            doc.Save(filename.ToString());

            MessageBox.Show("匯出完成-文件放在-" + filename.ToString());
            FileInfo fi = new FileInfo(filename.ToString());
            if (fi.Exists)
            {
                System.Diagnostics.Process.Start(filename.ToString());
            }
            else
            {
                //file doesn't exist
            }

        }
        public string FindMALFUNCION(string MALFUNCIONID)
        {

            DataSet ds = new DataSet();
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@" SELECT  [ID],[MALFUNCION]  FROM [TKMOC].[dbo].[MALFUNCION] WHERE   [ID] ='{0}'", MALFUNCIONID.ToString());

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds");
                sqlConn.Close();


                if (ds.Tables["TEMPds"].Rows.Count == 0)
                {
                    return "";
                }
                else
                {
                    if (ds.Tables["TEMPds"].Rows.Count >= 1)
                    {
                        return ds.Tables["TEMPds"].Rows[0]["MALFUNCION"].ToString();
                    }
                    return "";
                }

            }
            catch
            {
                return "";
            }
            finally
            {

            }

        }
        public void ExcelExportMACHINEDAILYCHECK()
        {

            string NowDB = "TK";
            //建立Excel 2003檔案
            IWorkbook wb = new XSSFWorkbook();
            ISheet ws;

            XSSFCellStyle cs = (XSSFCellStyle)wb.CreateCellStyle();
            //框線樣式及顏色
            cs.BorderBottom = NPOI.SS.UserModel.BorderStyle.Double;
            cs.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            cs.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            cs.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            cs.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;
            cs.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;
            cs.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;
            cs.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;

            //Search();            
            dt = dtMACHINEDAILYCHECK;

            if (dt.TableName != string.Empty)
            {
                ws = wb.CreateSheet(dt.TableName);
            }
            else
            {
                ws = wb.CreateSheet("Sheet1");
            }

            ws.CreateRow(0);//第一行為欄位名稱
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ws.GetRow(0).CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);
            }

            int j = 0;
            int k = dataGridView7.Rows.Count;
            foreach (DataGridViewRow dr in this.dataGridView7.Rows)
            {
                ws.CreateRow(j + 1);
                ws.GetRow(j + 1).CreateCell(0).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString());
                ws.GetRow(j + 1).CreateCell(1).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());
                ws.GetRow(j + 1).CreateCell(2).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString());
                ws.GetRow(j + 1).CreateCell(3).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString());
                ws.GetRow(j + 1).CreateCell(4).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[4].ToString());
                ws.GetRow(j + 1).CreateCell(5).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[5].ToString());
                ws.GetRow(j + 1).CreateCell(6).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[6].ToString());
                ws.GetRow(j + 1).CreateCell(7).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[7].ToString());
                ws.GetRow(j + 1).CreateCell(8).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[8].ToString());
                ws.GetRow(j + 1).CreateCell(9).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[9].ToString());
                ws.GetRow(j + 1).CreateCell(10).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[10].ToString());
                ws.GetRow(j + 1).CreateCell(11).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[11].ToString());
                ws.GetRow(j + 1).CreateCell(12).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[12].ToString());
                ws.GetRow(j + 1).CreateCell(13).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[13].ToString());
                ws.GetRow(j + 1).CreateCell(14).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[14].ToString());
                ws.GetRow(j + 1).CreateCell(15).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[15].ToString());

                //ws.GetRow(j + 1).CreateCell(3).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString()));
                j++;
            }

            if (Directory.Exists(@"c:\temp\"))
            {
                //資料夾存在
            }
            else
            {
                //新增資料夾
                Directory.CreateDirectory(@"c:\temp\");
            }
            StringBuilder filename = new StringBuilder();
            filename.AppendFormat(@"c:\temp\設備日常檢查表{0}.xlsx", DateTime.Now.ToString("yyyyMMdd"));

            FileStream file = new FileStream(filename.ToString(), FileMode.Create);//產生檔案
            wb.Write(file);
            file.Close();

            MessageBox.Show("匯出完成-EXCEL放在-" + filename.ToString());
            FileInfo fi = new FileInfo(filename.ToString());
            if (fi.Exists)
            {
                System.Diagnostics.Process.Start(filename.ToString());
            }
            else
            {
                //file doesn't exist
            }


        }
        public void ExcelExportMACHINEMAINWEEK()
        {

            string NowDB = "TK";
            //建立Excel 2003檔案
            IWorkbook wb = new XSSFWorkbook();
            ISheet ws;

            XSSFCellStyle cs = (XSSFCellStyle)wb.CreateCellStyle();
            //框線樣式及顏色
            cs.BorderBottom = NPOI.SS.UserModel.BorderStyle.Double;
            cs.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            cs.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            cs.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            cs.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;
            cs.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;
            cs.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;
            cs.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;

            //Search();            
            dt = dtMACHINEMAINWEEK;

            if (dt.TableName != string.Empty)
            {
                ws = wb.CreateSheet(dt.TableName);
            }
            else
            {
                ws = wb.CreateSheet("Sheet1");
            }

            ws.CreateRow(0);//第一行為欄位名稱
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ws.GetRow(0).CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);
            }

            int j = 0;
            int k = dataGridView9.Rows.Count;
            foreach (DataGridViewRow dr in this.dataGridView9.Rows)
            {
                ws.CreateRow(j + 1);
                ws.GetRow(j + 1).CreateCell(0).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString());
                ws.GetRow(j + 1).CreateCell(1).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());
                ws.GetRow(j + 1).CreateCell(2).SetCellValue(FindUNIT(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString()));
                ws.GetRow(j + 1).CreateCell(3).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString());
                ws.GetRow(j + 1).CreateCell(4).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[4].ToString());
                ws.GetRow(j + 1).CreateCell(5).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[5].ToString());
                ws.GetRow(j + 1).CreateCell(6).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[6].ToString());
                ws.GetRow(j + 1).CreateCell(7).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[7].ToString());
                ws.GetRow(j + 1).CreateCell(8).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[8].ToString());
            
                //ws.GetRow(j + 1).CreateCell(3).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString()));
                j++;
            }

            if (Directory.Exists(@"c:\temp\"))
            {
                //資料夾存在
            }
            else
            {
                //新增資料夾
                Directory.CreateDirectory(@"c:\temp\");
            }
            StringBuilder filename = new StringBuilder();
            filename.AppendFormat(@"c:\temp\定期維護保養計晝表{0}.xlsx", DateTime.Now.ToString("yyyyMMdd"));

            FileStream file = new FileStream(filename.ToString(), FileMode.Create);//產生檔案
            wb.Write(file);
            file.Close();

            MessageBox.Show("匯出完成-EXCEL放在-" + filename.ToString());
            FileInfo fi = new FileInfo(filename.ToString());
            if (fi.Exists)
            {
                System.Diagnostics.Process.Start(filename.ToString());
            }
            else
            {
                //file doesn't exist
            }


        }
        public void ExcelExportMACHINEMAINRECORD()
        {

            string NowDB = "TK";
            //建立Excel 2003檔案
            IWorkbook wb = new XSSFWorkbook();
            ISheet ws;

            XSSFCellStyle cs = (XSSFCellStyle)wb.CreateCellStyle();
            //框線樣式及顏色
            cs.BorderBottom = NPOI.SS.UserModel.BorderStyle.Double;
            cs.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            cs.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            cs.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            cs.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;
            cs.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;
            cs.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;
            cs.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;

            //Search();            
            dt = dtMACHINEMAINRECORD;

            if (dt.TableName != string.Empty)
            {
                ws = wb.CreateSheet(dt.TableName);
            }
            else
            {
                ws = wb.CreateSheet("Sheet1");
            }

            ws.CreateRow(0);//第一行為欄位名稱
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ws.GetRow(0).CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);
            }

            int j = 0;
            int k = dataGridView8.Rows.Count;
            foreach (DataGridViewRow dr in this.dataGridView8.Rows)
            {
                ws.CreateRow(j + 1);
                ws.GetRow(j + 1).CreateCell(0).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString());
                ws.GetRow(j + 1).CreateCell(1).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());
                ws.GetRow(j + 1).CreateCell(2).SetCellValue(FindUNIT(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString()));
                ws.GetRow(j + 1).CreateCell(3).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString());
                ws.GetRow(j + 1).CreateCell(4).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[4].ToString());
                ws.GetRow(j + 1).CreateCell(5).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[5].ToString());
                ws.GetRow(j + 1).CreateCell(6).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[6].ToString());
          
                //ws.GetRow(j + 1).CreateCell(3).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString()));
                j++;
            }

            if (Directory.Exists(@"c:\temp\"))
            {
                //資料夾存在
            }
            else
            {
                //新增資料夾
                Directory.CreateDirectory(@"c:\temp\");
            }
            StringBuilder filename = new StringBuilder();
            filename.AppendFormat(@"c:\temp\保養維護記錄卡{0}.xlsx", DateTime.Now.ToString("yyyyMMdd"));

            FileStream file = new FileStream(filename.ToString(), FileMode.Create);//產生檔案
            wb.Write(file);
            file.Close();

            MessageBox.Show("匯出完成-EXCEL放在-" + filename.ToString());
            FileInfo fi = new FileInfo(filename.ToString());
            if (fi.Exists)
            {
                System.Diagnostics.Process.Start(filename.ToString());
            }
            else
            {
                //file doesn't exist
            }


        }

        public void ExcelExportMAINPARTSUSED()
        {

            string NowDB = "TK";
            //建立Excel 2003檔案
            IWorkbook wb = new XSSFWorkbook();
            ISheet ws;

            XSSFCellStyle cs = (XSSFCellStyle)wb.CreateCellStyle();
            //框線樣式及顏色
            cs.BorderBottom = NPOI.SS.UserModel.BorderStyle.Double;
            cs.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            cs.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            cs.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            cs.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;
            cs.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;
            cs.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;
            cs.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;

            //Search();            
            dt = dtMAINPARTSUSED;

            if (dt.TableName != string.Empty)
            {
                ws = wb.CreateSheet(dt.TableName);
            }
            else
            {
                ws = wb.CreateSheet("Sheet1");
            }

            ws.CreateRow(0);//第一行為欄位名稱
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ws.GetRow(0).CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);
            }

            int j = 0;
            int k = dataGridView11.Rows.Count;
            foreach (DataGridViewRow dr in this.dataGridView11.Rows)
            {
                ws.CreateRow(j + 1);
                ws.GetRow(j + 1).CreateCell(0).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString());
                ws.GetRow(j + 1).CreateCell(1).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());
                ws.GetRow(j + 1).CreateCell(2).SetCellValue(FindUNIT(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString()));
                ws.GetRow(j + 1).CreateCell(3).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString());
                ws.GetRow(j + 1).CreateCell(4).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[4].ToString());
                ws.GetRow(j + 1).CreateCell(5).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[5].ToString());
                ws.GetRow(j + 1).CreateCell(6).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[6].ToString());
                ws.GetRow(j + 1).CreateCell(7).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[7].ToString());
                ws.GetRow(j + 1).CreateCell(8).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[8].ToString());
                ws.GetRow(j + 1).CreateCell(9).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[9].ToString());
                ws.GetRow(j + 1).CreateCell(10).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[10].ToString());
                ws.GetRow(j + 1).CreateCell(11).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[11].ToString());
                ws.GetRow(j + 1).CreateCell(12).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[12].ToString());
                ws.GetRow(j + 1).CreateCell(13).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[13].ToString());
                ws.GetRow(j + 1).CreateCell(14).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[14].ToString());
            
                //ws.GetRow(j + 1).CreateCell(3).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString()));
                j++;
            }

            if (Directory.Exists(@"c:\temp\"))
            {
                //資料夾存在
            }
            else
            {
                //新增資料夾
                Directory.CreateDirectory(@"c:\temp\");
            }
            StringBuilder filename = new StringBuilder();
            filename.AppendFormat(@"c:\temp\備品管制卡{0}.xlsx", DateTime.Now.ToString("yyyyMMdd"));

            FileStream file = new FileStream(filename.ToString(), FileMode.Create);//產生檔案
            wb.Write(file);
            file.Close();

            MessageBox.Show("匯出完成-EXCEL放在-" + filename.ToString());
            FileInfo fi = new FileInfo(filename.ToString());
            if (fi.Exists)
            {
                System.Diagnostics.Process.Start(filename.ToString());
            }
            else
            {
                //file doesn't exist
            }


        }
        public void ExcelExportMACHINECARD()
        {

            string NowDB = "TK";
            //建立Excel 2003檔案
            IWorkbook wb = new XSSFWorkbook();
            ISheet ws;

            XSSFCellStyle cs = (XSSFCellStyle)wb.CreateCellStyle();
            //框線樣式及顏色
            cs.BorderBottom = NPOI.SS.UserModel.BorderStyle.Double;
            cs.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            cs.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            cs.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            cs.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;
            cs.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;
            cs.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;
            cs.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;

            //Search();            
            dt = dtMACHINECARD;

            if (dt.TableName != string.Empty)
            {
                ws = wb.CreateSheet(dt.TableName);
            }
            else
            {
                ws = wb.CreateSheet("Sheet1");
            }

            ws.CreateRow(0);//第一行為欄位名稱
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ws.GetRow(0).CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);
            }

            int j = 0;
            int k = dataGridView5.Rows.Count;
            foreach (DataGridViewRow dr in this.dataGridView5.Rows)
            {
                ws.CreateRow(j + 1);
                ws.GetRow(j + 1).CreateCell(0).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString());
                ws.GetRow(j + 1).CreateCell(1).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());
                ws.GetRow(j + 1).CreateCell(2).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString());
                ws.GetRow(j + 1).CreateCell(3).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString());
                ws.GetRow(j + 1).CreateCell(4).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[4].ToString());
                ws.GetRow(j + 1).CreateCell(5).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[5].ToString());
                ws.GetRow(j + 1).CreateCell(6).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[6].ToString());
                ws.GetRow(j + 1).CreateCell(7).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[7].ToString());
                ws.GetRow(j + 1).CreateCell(8).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[8].ToString());
                ws.GetRow(j + 1).CreateCell(9).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[9].ToString());
                ws.GetRow(j + 1).CreateCell(10).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[10].ToString());
                ws.GetRow(j + 1).CreateCell(11).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[11].ToString());
                ws.GetRow(j + 1).CreateCell(12).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[12].ToString());
                ws.GetRow(j + 1).CreateCell(13).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[13].ToString());
                ws.GetRow(j + 1).CreateCell(14).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[14].ToString());
                ws.GetRow(j + 1).CreateCell(15).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[15].ToString());
                ws.GetRow(j + 1).CreateCell(16).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[16].ToString());
                ws.GetRow(j + 1).CreateCell(17).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[17].ToString());
                ws.GetRow(j + 1).CreateCell(18).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[18].ToString());
                ws.GetRow(j + 1).CreateCell(19).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[19].ToString());
                ws.GetRow(j + 1).CreateCell(20).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[20].ToString());
               

                //ws.GetRow(j + 1).CreateCell(3).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString()));
                j++;
            }

            if (Directory.Exists(@"c:\temp\"))
            {
                //資料夾存在
            }
            else
            {
                //新增資料夾
                Directory.CreateDirectory(@"c:\temp\");
            }
            StringBuilder filename = new StringBuilder();
            filename.AppendFormat(@"c:\temp\機器設備卡{0}.xlsx", DateTime.Now.ToString("yyyyMMdd"));

            FileStream file = new FileStream(filename.ToString(), FileMode.Create);//產生檔案
            wb.Write(file);
            file.Close();

            MessageBox.Show("匯出完成-EXCEL放在-" + filename.ToString());
            FileInfo fi = new FileInfo(filename.ToString());
            if (fi.Exists)
            {
                System.Diagnostics.Process.Start(filename.ToString());
            }
            else
            {
                //file doesn't exist
            }


        }
        public void ExcelExportMACHINEATTACH()
        {

            string NowDB = "TK";
            //建立Excel 2003檔案
            IWorkbook wb = new XSSFWorkbook();
            ISheet ws;

            XSSFCellStyle cs = (XSSFCellStyle)wb.CreateCellStyle();
            //框線樣式及顏色
            cs.BorderBottom = NPOI.SS.UserModel.BorderStyle.Double;
            cs.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            cs.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            cs.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            cs.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;
            cs.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;
            cs.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;
            cs.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;

            //Search();            
            dt = dtMACHINEATTACH;

            if (dt.TableName != string.Empty)
            {
                ws = wb.CreateSheet(dt.TableName);
            }
            else
            {
                ws = wb.CreateSheet("Sheet1");
            }

            ws.CreateRow(0);//第一行為欄位名稱
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ws.GetRow(0).CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);
            }

            int j = 0;
            int k = dataGridView6.Rows.Count;
            foreach (DataGridViewRow dr in this.dataGridView6.Rows)
            {
                ws.CreateRow(j + 1);
                ws.GetRow(j + 1).CreateCell(0).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString());
                ws.GetRow(j + 1).CreateCell(1).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());
                ws.GetRow(j + 1).CreateCell(2).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString());
                ws.GetRow(j + 1).CreateCell(3).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString());
                ws.GetRow(j + 1).CreateCell(4).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[4].ToString());
                ws.GetRow(j + 1).CreateCell(5).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[5].ToString());
               
                //ws.GetRow(j + 1).CreateCell(3).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString()));
                j++;
            }

            if (Directory.Exists(@"c:\temp\"))
            {
                //資料夾存在
            }
            else
            {
                //新增資料夾
                Directory.CreateDirectory(@"c:\temp\");
            }
            StringBuilder filename = new StringBuilder();
            filename.AppendFormat(@"c:\temp\機器設備卡附件{0}.xlsx", DateTime.Now.ToString("yyyyMMdd"));

            FileStream file = new FileStream(filename.ToString(), FileMode.Create);//產生檔案
            wb.Write(file);
            file.Close();

            MessageBox.Show("匯出完成-EXCEL放在-" + filename.ToString());
            FileInfo fi = new FileInfo(filename.ToString());
            if (fi.Exists)
            {
                System.Diagnostics.Process.Start(filename.ToString());
            }
            else
            {
                //file doesn't exist
            }


        }
        #endregion

        #region BUTTION
        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            frmEngineeringAddEditDel objfrmEngineeringAddEditDel = new frmEngineeringAddEditDel("");
            objfrmEngineeringAddEditDel.ShowDialog();
            Search();

        }
        private void button3_Click(object sender, EventArgs e)
        {
            frmEngineeringAddEditDel objfrmEngineeringAddEditDel = new frmEngineeringAddEditDel(EquipmentID);
            objfrmEngineeringAddEditDel.ShowDialog();
            Search();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            SearchMAINAPPLY();
            SearchMAINAPPLYOUT();
            SearchMAINRECORD();
        }
        private void button6_Click(object sender, EventArgs e)
        {
            frmMAINAPPLYAddEditDel objfrmMAINAPPLYAddEditDel = new frmMAINAPPLYAddEditDel(MAINAPPLYID);
            objfrmMAINAPPLYAddEditDel.ShowDialog();
            SearchMAINAPPLY();
            button4.PerformClick();
        }
        private void button5_Click(object sender, EventArgs e)
        {
            frmMAINAPPLYAddEditDel objfrmMAINAPPLYAddEditDel = new frmMAINAPPLYAddEditDel("");
            objfrmMAINAPPLYAddEditDel.ShowDialog();
            SearchMAINAPPLY();
            button4.PerformClick();
        }
        private void button9_Click(object sender, EventArgs e)
        {
            frmMAINAPPLYOUTAddEditDel objfrmMAINAPPLYOUTAddEditDel = new frmMAINAPPLYOUTAddEditDel(MAINAPPLYOUTID);
            objfrmMAINAPPLYOUTAddEditDel.ShowDialog();
            SearchMAINAPPLYOUT();
            button4.PerformClick();
        }
        private void button7_Click(object sender, EventArgs e)
        {
            frmMAINAPPLYOUTAddEditDel objfrmMAINAPPLYOUTAddEditDel = new frmMAINAPPLYOUTAddEditDel("");
            objfrmMAINAPPLYOUTAddEditDel.ShowDialog();
            SearchMAINAPPLYOUT();
            button4.PerformClick();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            frmMAINRECORDAddEditDel objfrmMAINRECORDAddEditDel = new frmMAINRECORDAddEditDel(MAINRECORDID);
            objfrmMAINRECORDAddEditDel.ShowDialog();
            SearchMAINRECORD();
            button4.PerformClick();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            frmMAINRECORDAddEditDel objfrmMAINRECORDAddEditDel = new frmMAINRECORDAddEditDel("");
            objfrmMAINRECORDAddEditDel.ShowDialog();
            SearchMAINRECORD();
            button4.PerformClick();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            SearchMCHINE();
            SearchMCHINEATTACH();
        }
        private void button13_Click(object sender, EventArgs e)
        {
            frmMACHINECARD objfrmMACHINECARD = new frmMACHINECARD(MACHINEID);
            objfrmMACHINECARD.ShowDialog();
            SearchMCHINE();
            button11.PerformClick();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            frmMACHINECARD objfrmMACHINECARD = new frmMACHINECARD("");
            objfrmMACHINECARD.ShowDialog();
            SearchMCHINE();
        }

        private void button15_Click(object sender, EventArgs e)
        {
            frmMACHINEATTACH objfrmMACHINEATTACH = new frmMACHINEATTACH(MACHINEID);
            objfrmMACHINEATTACH.ShowDialog();
            SearchMCHINEATTACH();
            button11.PerformClick();
        }

        private void button14_Click(object sender, EventArgs e)
        {
            frmMACHINEATTACH objfrmMACHINEATTACH = new frmMACHINEATTACH("");
            objfrmMACHINEATTACH.ShowDialog();
            SearchMCHINEATTACH();
            button11.PerformClick();
        }
        private void button16_Click(object sender, EventArgs e)
        {
            SearchMACHINEDAILYCHECK();
            SearchMACHINEMAINRECORD();
            SearchMACHINEMAINWEEK();
        }
        private void button18_Click(object sender, EventArgs e)
        {
            frmMACHINEDAILYCHECK objfrmMACHINEDAILYCHECK = new frmMACHINEDAILYCHECK(MACHINEDAILYCHECK);
            objfrmMACHINEDAILYCHECK.ShowDialog();
            SearchMACHINEDAILYCHECK();
            button16.PerformClick();
        }

        private void button17_Click(object sender, EventArgs e)
        {
            frmMACHINEDAILYCHECK objfrmMACHINEDAILYCHECK = new frmMACHINEDAILYCHECK("");
            objfrmMACHINEDAILYCHECK.ShowDialog();
            SearchMACHINEDAILYCHECK();
            button16.PerformClick();
        }

        private void button21_Click(object sender, EventArgs e)
        {
            frmMACHINEMAINRECORD objfrmMACHINEMAINRECORD = new frmMACHINEMAINRECORD(MACHINEMAINRECORDID);
            objfrmMACHINEMAINRECORD.ShowDialog();
            SearchMACHINEMAINRECORD();
            button16.PerformClick();
        }

        private void button19_Click(object sender, EventArgs e)
        {
            frmMACHINEMAINRECORD objfrmMACHINEMAINRECORD = new frmMACHINEMAINRECORD("");
            objfrmMACHINEMAINRECORD.ShowDialog();
            SearchMACHINEMAINRECORD();
            button16.PerformClick();
        }

        private void button22_Click(object sender, EventArgs e)
        {
            frmMACHINEMAINWEEK objfrmMACHINEMAINWEEK = new frmMACHINEMAINWEEK(MACHINEMAINWEEKID);
            objfrmMACHINEMAINWEEK.ShowDialog();
            SearchMACHINEMAINWEEK();
            button16.PerformClick();
        }

        private void button20_Click(object sender, EventArgs e)
        {
            frmMACHINEMAINWEEK objfrmMACHINEMAINWEEK = new frmMACHINEMAINWEEK("");
            objfrmMACHINEMAINWEEK.ShowDialog();
            SearchMACHINEMAINWEEK();
            button16.PerformClick();
        }
        private void button23_Click(object sender, EventArgs e)
        {
            SearchMAINPARTS();
            SearchMAINPARTSUSED();
        }
        private void button27_Click(object sender, EventArgs e)
        {
            frmMAINPARTS objfrmMAINPARTS = new frmMAINPARTS(MAINPARTSID);
            objfrmMAINPARTS.ShowDialog();
            SearchMAINPARTS();
            button23.PerformClick();
        }

        private void button25_Click(object sender, EventArgs e)
        {
            frmMAINPARTS objfrmMAINPARTS = new frmMAINPARTS("");
            objfrmMAINPARTS.ShowDialog();
            SearchMAINPARTS();
            button23.PerformClick();
        }

        private void button26_Click(object sender, EventArgs e)
        {
            frmMAINPARTSUSED objfrmMAINPARTSUSED = new frmMAINPARTSUSED(MAINPARTSUSEDID);
            objfrmMAINPARTSUSED.ShowDialog();
            SearchMAINPARTSUSED();
            button23.PerformClick();
        }

        private void button24_Click(object sender, EventArgs e)
        {
            frmMAINPARTSUSED objfrmMAINPARTSUSED = new frmMAINPARTSUSED("");
            objfrmMAINPARTSUSED.ShowDialog();
            SearchMAINPARTSUSED();
            button23.PerformClick();
        }



        private void button28_Click(object sender, EventArgs e)
        {
            ExcelExport();
        }
        private void button29_Click(object sender, EventArgs e)
        {
            ExcelExportMAINPARTS();
        }
        private void button30_Click(object sender, EventArgs e)
        {
            PRINTMAINAPPLY();
        }
        private void button31_Click(object sender, EventArgs e)
        {
            PRINTMAINAPPLYOUT();
        }
        private void button32_Click(object sender, EventArgs e)
        {
            PRINTMAINRECORD();
        }
        private void button33_Click(object sender, EventArgs e)
        {
            ExcelExportMACHINEDAILYCHECK();
        }
        private void button34_Click(object sender, EventArgs e)
        {
            ExcelExportMACHINEMAINWEEK();
        }
        private void button35_Click(object sender, EventArgs e)
        {
            ExcelExportMACHINEMAINRECORD();
        }
        private void button36_Click(object sender, EventArgs e)
        {
            ExcelExportMAINPARTSUSED();
        }
        private void button38_Click(object sender, EventArgs e)
        {
            ExcelExportMACHINECARD();
        }
        private void button37_Click(object sender, EventArgs e)
        {
            ExcelExportMACHINEATTACH();
        }

        #endregion


    }
}
