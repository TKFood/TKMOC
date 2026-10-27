using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using TKITDLL;

namespace TKMOC
{
    /// <summary>
    /// 資料庫存取共用 helper。
    /// 取代各方法中重複的「讀取 dbconn → 解密帳號密碼 → 建立連線」樣板（20210902密）。
    /// </summary>
    internal static class Db
    {
        private static readonly object _lock = new object();
        private static string _connectionString;

        /// <summary>已解密帳號、密碼的 dbconn 連線字串（第一次使用時解密並快取）。</summary>
        public static string ConnectionString
        {
            get
            {
                if (_connectionString == null)
                {
                    lock (_lock)
                    {
                        if (_connectionString == null)
                        {
                            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(
                                ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                            //資料庫使用者密碼解密
                            Class1 TKID = new Class1();
                            sqlsb.Password = TKID.Decryption(sqlsb.Password);
                            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                            _connectionString = sqlsb.ConnectionString;
                        }
                    }
                }
                return _connectionString;
            }
        }

        /// <summary>建立（尚未開啟的）dbconn 連線。</summary>
        public static SqlConnection CreateConnection()
        {
            return new SqlConnection(ConnectionString);
        }

        /// <summary>執行查詢並回傳 DataTable。</summary>
        public static DataTable Query(string sql, params SqlParameter[] parameters)
        {
            using (SqlConnection conn = CreateConnection())
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
            {
                if (parameters != null && parameters.Length > 0)
                {
                    cmd.Parameters.AddRange(parameters);
                }

                DataTable table = new DataTable();
                adapter.Fill(table);
                return table;
            }
        }

        /// <summary>
        /// 在交易中執行異動 SQL：影響 0 筆時 Rollback，否則 Commit（與原本各方法的寫法相同）。
        /// 回傳影響筆數；發生例外時會 Rollback 並把例外往外丟。
        /// </summary>
        public static int ExecuteInTransaction(string sql, int commandTimeout = 60, params SqlParameter[] parameters)
        {
            using (SqlConnection conn = CreateConnection())
            {
                conn.Open();
                using (SqlTransaction tran = conn.BeginTransaction())
                using (SqlCommand cmd = new SqlCommand(sql, conn, tran))
                {
                    cmd.CommandTimeout = commandTimeout;
                    if (parameters != null && parameters.Length > 0)
                    {
                        cmd.Parameters.AddRange(parameters);
                    }

                    try
                    {
                        int rows = cmd.ExecuteNonQuery();
                        if (rows == 0)
                        {
                            tran.Rollback();    //交易取消
                        }
                        else
                        {
                            tran.Commit();      //執行交易
                        }
                        return rows;
                    }
                    catch
                    {
                        try { tran.Rollback(); } catch { /* 連線已中斷時 Rollback 可能失敗，忽略 */ }
                        throw;
                    }
                }
            }
        }
    }
}
