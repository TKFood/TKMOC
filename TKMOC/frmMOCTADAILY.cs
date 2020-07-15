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
using FastReport;
using FastReport.Data;

namespace TKMOC
{
    public partial class frmMOCTADAILY : Form
    {
        private ComponentResourceManager _ResourceManager = new ComponentResourceManager();
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();

        int result;
        string STATUS = null;
        string ID = null;

        public frmMOCTADAILY()
        {
            InitializeComponent();
        }

        #region FUNCTION

        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {

        }

        #endregion
    }
}
