using System;
using UTDataValidator;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Linq;

namespace UnitTestProject1
{
    public class SampleEventValidator : IEventExcelValidator
    {
        private void Picking_ScanTrolley(Dictionary<string, string> parameters)
        {

        }

        private void Picking_ScanPart(Dictionary<string, string> parameters)
        {

        }

        public void ExecuteAction(TestAction action)
        {
            if (action.ActionName == "picking_scan_trolley")
            {
                Picking_ScanTrolley(action.Parameters);
            }
            else if (action.ActionName == "picking_scan_part")
            {
                Picking_ScanPart(action.Parameters);
            }
            else
            {
                throw new Exception("Action not found.");
            }
        }

        public void InitData(IEnumerable<ExcelDataDefinition> excelDataDefinition)
        {
            foreach (var item in excelDataDefinition)
            {
                using (SqlConnection connection = new SqlConnection("Data Source=.; Initial Catalog=ADM_SPAREPART_LENGKAP; Integrated Security=True; MultipleActiveResultSets=True; Connection Timeout=864000;"))
                {
                    connection.Open();
                    using (SqlCommand command = connection.CreateCommand())
                    {
                        string sqlCleanTable = $"DELETE FROM {item.Table};";
                        command.Connection = connection;
                        command.CommandType = CommandType.Text;
                        command.CommandText = sqlCleanTable;
                        command.ExecuteNonQuery();
                    }

                    foreach (DataRow row in item.Data.Rows)
                    {
                        using (SqlCommand command = connection.CreateCommand())
                        {
                            command.Connection = connection;
                            command.CommandType = CommandType.Text;

                            List<string> columns = new List<string>();
                            foreach (DataColumn column in item.Data.Columns)
                            {
                                if (row[column.ColumnName] == DBNull.Value) continue;
                                columns.Add(column.ColumnName);

                                command.Parameters.Add(new SqlParameter($"@{column.ColumnName}", row[column.ColumnName]));
                            }

                            string sql = $"INSERT INTO {item.Table} ({string.Join(", ", columns)}) VALUES ({string.Join(", ", columns.Select(f => "@" + f).ToList())});";
                            command.CommandText = sql;
                            command.ExecuteNonQuery();
                        }
                    }
                }
            }
        }

        public DataTable ReadTable(ExcelDataDefinition excelDataDefinition)
        {
            DataTable result = new DataTable();
            string sql = $"SELECT * FROM {excelDataDefinition.Table};";
            using (SqlConnection connection = new SqlConnection("Data Source=.; Initial Catalog=ADM_SPAREPART_LENGKAP; Integrated Security=True; MultipleActiveResultSets=True; Connection Timeout=864000;"))
            {
                connection.Open();
                using (SqlCommand command = connection.CreateCommand())
                {
                    command.Connection = connection;
                    command.CommandType = CommandType.Text;
                    command.CommandText = sql;

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        result.Load(reader);
                    }
                }
            }
            return result;
        }
    }
}
