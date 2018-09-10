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

namespace TKMOC
{
    public partial class FrmMOCWEEK : Form
    {
        private ComponentResourceManager _ResourceManager = new ComponentResourceManager();
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlDataAdapter adapter2 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        SqlDataAdapter adapterCALENDAR = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilderCALENDAR = new SqlCommandBuilder();
        SqlDataAdapter adapterCALENDAR2 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilderCALENDAR2 = new SqlCommandBuilder();

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();


        int result;

        public FrmMOCWEEK()
        {
            InitializeComponent();

            SETTODAY();
            SETFIRSTDAY();
        }

        #region FUNCTION
        public void SETTODAY()
        {
            dateTimePicker1.Value = DateTime.Now;
        }

        public void SETFIRSTDAY()
        {
            DateTime dt = dateTimePicker1.Value;

            dt.AddDays(-((int)dt.DayOfWeek));
            dateTimePicker2.Value = GetWeekFirstDayMon(dt); 
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            DateTime dt = dateTimePicker1.Value;

            dateTimePicker2.Value = GetWeekFirstDayMon(dt);
            dateTimePicker3.Value = dateTimePicker2.Value.AddDays(1);
            dateTimePicker4.Value = dateTimePicker2.Value.AddDays(2);
            dateTimePicker5.Value = dateTimePicker2.Value.AddDays(3);
            dateTimePicker6.Value = dateTimePicker2.Value.AddDays(4);
            dateTimePicker7.Value = dateTimePicker2.Value.AddDays(5);
            dateTimePicker8.Value = dateTimePicker2.Value.AddDays(6);
        }

        public DateTime GetWeekFirstDayMon(DateTime datetime)
        {
            //星期一为第一天
            int weeknow = Convert.ToInt32(datetime.DayOfWeek);

            //因为是以星期一为第一天，所以要判断weeknow等于0时，要向前推6天。
            weeknow = (weeknow == 0 ? (7 - 1) : (weeknow - 1));
            int daydiff = (-1) * weeknow;

            //本周第一天
            string FirstDay = datetime.AddDays(daydiff).ToString("yyyy-MM-dd");
            return Convert.ToDateTime(FirstDay);
        }
        #endregion




        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {

        }
        #endregion

        
    }
}
