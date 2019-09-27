using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QLRUpdate
{
    class OleDbHelper
    {
        /// <summary>
        /// 生成table
        /// </summary>
        /// <param name="oleDb"></param>
        /// <returns></returns>
        public static DataTable QueryTable(string oleDb, string Path)
        {
            string OleDbConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Path;
            DataTable table = new DataTable();
            using (OleDbConnection conn = new OleDbConnection(OleDbConnectionString))
            {
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;
                cmd.CommandText = oleDb;
                try
                {
                    if (conn.State != ConnectionState.Open)
                    {
                        conn.Open();
                    }

                    OleDbDataAdapter adapter = new OleDbDataAdapter(cmd);
                    adapter.Fill(table);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                finally
                {
                    conn.Close();
                    cmd.Dispose();
                }
            }

            return table;
        }



        /// <summary>
        /// 执行sql语句
        /// </summary>
        /// <param name="oleDb"></param>
        /// <returns>受影响行数</returns>
        public static int RunCommand(string oleDb, string Path)
        {
            string OleDbConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Path;
            int result = 0;
            using (OleDbConnection connection = new OleDbConnection(OleDbConnectionString))
            {
                try
                {
                    if (connection.State != ConnectionState.Open)
                    {
                        connection.Open();
                    }

                    var command = connection.CreateCommand();//创建OleDbCommand对象
                    command.CommandText = oleDb;
                    result = command.ExecuteNonQuery();
                    connection.Close();
                }
                catch (Exception ex)
                {
                    if (connection.State != ConnectionState.Closed)
                    {
                        connection.Close();
                    }

                    throw ex;
                }


            }


            return result;

        }

        /// <summary>
        /// 执行事务
        /// </summary>
        /// <param name="oleDblist"></param>
        /// <returns></returns>
        public static int RunTransAction(List<string> oleDblist, string Path)
        {
            string OleDbConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Path;
            int result = 0;
            using (OleDbConnection connection = new OleDbConnection(OleDbConnectionString))
            {
                //创建连接对象
                connection.Open();
                OleDbTransaction oleDbTran = connection.BeginTransaction();//开始事务
                var command = connection.CreateCommand();//创建OleDbCommand对象
                command.Transaction = oleDbTran;//将OleDbCommand与OleDbTransaction关联起来
                try
                {
                    for (int i = 0; i < oleDblist.Count; i++)
                    {
                        command.CommandText = oleDblist[i];
                        int res = command.ExecuteNonQuery();
                        result = result + res;
                    }

                    oleDbTran.Commit();
                    connection.Close();

                }
                catch (Exception ex)
                {
                    oleDbTran.Rollback();
                    throw ex;
                }

            }

            return result;

        }
    }
}
