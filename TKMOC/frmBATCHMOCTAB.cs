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
    public partial class frmBATCHMOCTAB : Form
    {
        private ComponentResourceManager _ResourceManager = new ComponentResourceManager();
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();

        public class ADDITEM
        {
            public string MB001;
            public double NUM;

        }

        public frmBATCHMOCTAB()
        {
            InitializeComponent();
        }

        public void TEST()
        {
            List<ADDITEM> ADDS = new List<ADDITEM>();

            ADDS.Add(new ADDITEM { MB001 = "Honda", NUM =1.23 });
            ADDS.Add(new ADDITEM { MB001 = "Vroom", NUM = 4.56 });

            foreach(var add in ADDS)
            {
                MessageBox.Show(add.MB001+" "+add.NUM);
            }
        }

        #region FUNCTION

        #endregion

        #region BUTTON
        private void button2_Click(object sender, EventArgs e)
        {
            TEST();
        }
        #endregion
    }
}
