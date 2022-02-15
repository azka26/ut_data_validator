using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using UTDataValidator;
using System.Data;
using System.Data.SqlClient;
using System.Data.Common;
using System.Collections.Generic;
using System.Linq;

namespace UnitTestProject1
{
    public class UnitTestBase
    {
        protected ExcelValidator GetExcelValidator(string excelPath, string worksheetInitData, string worksheetExpected)
        {
            ExcelValidator excelValidator = new ExcelValidator(
                excelPath: excelPath
                , worksheetInitData: "Sheet1"
                , worksheetExpectedData: "Sheet2",
                (ExcelDataDefinition definition) =>
                {
                    DataTable result = new DataTable();
                    string sql = $"SELECT * FROM {definition.Table};";
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
                },
                (IEnumerable<ExcelDataDefinition> data) =>
                {
                    foreach (var item in data)
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
            );

            excelValidator.ExecuteAction((TestAction action) =>
            {
                for (int i = 0; i < action.Loop; i++)
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
            });

            return excelValidator;
        }

        protected void Picking_ScanTrolley(Dictionary<string, string> parameters)
        {
            
        }

        protected void Picking_ScanPart(Dictionary<string, string> parameters)
        {

        }
    }

    [TestClass]
    public class OutboundUnitTest : UnitTestBase
    {
        [TestMethod]
        public void Picking_Test()
        {
            string excelPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "excel", "sample.xlsx");
            ExcelValidator excelValidator = GetExcelValidator(excelPath, "Sheet1", "Sheet2");
            excelValidator.Validate();
        }
    }
}
