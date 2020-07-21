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

namespace TKMOC
{
    public partial class frmMOCMANULINESubTEMPADD : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();


        public frmMOCMANULINESubTEMPADD()
        {
            InitializeComponent();
        }

        public frmMOCMANULINESubTEMPADD(string MB001)
        {
            InitializeComponent();

            textBox1.Text = MB001;
            
        }

        #region FUNCTION
        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        #endregion

        #region BUTTON

        #endregion


    }
}
