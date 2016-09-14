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
        string tablename = null;
        string EquipmentID;
        string MAINAPPLYID;
        string MAINAPPLYOUTID;
        string MAINRECORDID;
        string MACHINEID;
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
                ws.GetRow(j + 1).CreateCell(3).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString()));
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
            filename.AppendFormat(@"c:\temp\設備{0}.xlsx", DateTime.Now.ToString("yyyyMMdd"));

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
        }
        private void button5_Click(object sender, EventArgs e)
        {
            frmMAINAPPLYAddEditDel objfrmMAINAPPLYAddEditDel = new frmMAINAPPLYAddEditDel("");
            objfrmMAINAPPLYAddEditDel.ShowDialog();
            SearchMAINAPPLY();
        }
        private void button9_Click(object sender, EventArgs e)
        {
            frmMAINAPPLYOUTAddEditDel objfrmMAINAPPLYOUTAddEditDel = new frmMAINAPPLYOUTAddEditDel(MAINAPPLYOUTID);
            objfrmMAINAPPLYOUTAddEditDel.ShowDialog();
            SearchMAINAPPLYOUT();
        }
        private void button7_Click(object sender, EventArgs e)
        {
            frmMAINAPPLYOUTAddEditDel objfrmMAINAPPLYOUTAddEditDel = new frmMAINAPPLYOUTAddEditDel("");
            objfrmMAINAPPLYOUTAddEditDel.ShowDialog();
            SearchMAINAPPLYOUT();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            frmMAINRECORDAddEditDel objfrmMAINRECORDAddEditDel = new frmMAINRECORDAddEditDel(MAINRECORDID);
            objfrmMAINRECORDAddEditDel.ShowDialog();
            SearchMAINRECORD();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            frmMAINRECORDAddEditDel objfrmMAINRECORDAddEditDel = new frmMAINRECORDAddEditDel("");
            objfrmMAINRECORDAddEditDel.ShowDialog();
            SearchMAINRECORD();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            SearchMCHINE();
        }
        private void button13_Click(object sender, EventArgs e)
        {
            frmMACHINECARD objfrmMACHINECARD = new frmMACHINECARD(MACHINEID);
            objfrmMACHINECARD.ShowDialog();
            SearchMCHINE();

        }

        private void button12_Click(object sender, EventArgs e)
        {
            frmMACHINECARD objfrmMACHINECARD = new frmMACHINECARD("");
            objfrmMACHINECARD.ShowDialog();
            SearchMCHINE();
        }



        #endregion

        
    }
}
