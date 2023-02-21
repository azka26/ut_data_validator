using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using UTDataValidator;
using System.Data.Common;
using System.Data.SqlClient;
using Newtonsoft.Json;
using OfficeOpenXml;

namespace UnitTestProject1
{
    public class ExcelValidatorTest : IEventExcelValidator
    {
        private string ConnectionString = "Data Source=localhost,1433;Initial Catalog=UnitTest;User ID=sa;Password=Strong@Password;";
        [Test()]
        public void ExcelValidator_Test()
        {
            var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "excel", "sample-tipe-data.xlsx");
            ExcelValidator excelValidator = new ExcelValidator(new Assertion(), path, "init", "expected", this);

            ExecuteNonQuery($@"
                    INSERT INTO SampleTable (Name, CreatedDate, CreatedTime, CreatedDateTime, ValueInt, ValueDouble) 
                    VALUES ('Andika', '2022-01-03', '00:02:00', '2022-01-03 00:02:00.000','22','20.3');
                ");
            
            excelValidator.Validate();
        }

        [Test]
        public void JsonValidator_Test()
        {
            JsonDataValidator jsonDataValidator = new JsonDataValidator(new Assertion(),
                JsonConvert.SerializeObject(new { Id = "1", Name = "Rizal" }));
            jsonDataValidator.ValidateData(JsonConvert.SerializeObject(new { Id = "1", Name = "Rizal" }));
        }

        private int ExecuteNonQuery(string sql, List<SqlParameter> parameters = null)
        {
            using (var sqlConnection = new SqlConnection(ConnectionString))
            {
                sqlConnection.Open();
                using (var sqlCommand = sqlConnection.CreateCommand())
                {
                    sqlCommand.CommandType = CommandType.Text;
                    sqlCommand.CommandText = sql;

                    if (parameters != null)
                    {
                        sqlCommand.Parameters.AddRange(parameters.ToArray());
                    }
                    
                    return sqlCommand.ExecuteNonQuery();
                }
            }
        }
        
        private DataTable ExecuteQuery(string sql)
        {
            DataTable dataTable = new DataTable();
            using (var sqlConnection = new SqlConnection(ConnectionString))
            {
                sqlConnection.Open();
                using (var sqlCommand = sqlConnection.CreateCommand())
                {
                    sqlCommand.CommandType = CommandType.Text;
                    sqlCommand.CommandText = sql;
                    using (var reader = sqlCommand.ExecuteReader())
                    {
                        dataTable.Load(reader);
                    }
                }
            }

            return dataTable;
        }

        public DataTable ReadTable(ExcelDataDefinition excelDataDefinition)
        {
            return ExecuteQuery($"SELECT * FROM {excelDataDefinition.Table}");
        }

        public void InitData(IEnumerable<ExcelDataDefinition> excelDataDefinition)
        {
            foreach (var definition in excelDataDefinition)
            {
                ExecuteNonQuery($"DELETE FROM {definition.Table}");
                foreach (DataRow dataRow in definition.Data.Rows)
                {
                    var fields = new List<string>();
                    var fieldParameters = new List<string>();
                    var parameters = new List<SqlParameter>();
                    foreach (DataColumn column in definition.Data.Columns)
                    {
                        fields.Add(column.ColumnName);
                        fieldParameters.Add($"@{column.ColumnName}");
                        parameters.Add(new SqlParameter($"@{column.ColumnName}", dataRow[column.ColumnName]));
                    }

                    var sql =
                        $@"     SET IDENTITY_INSERT {definition.Table} ON;
                                INSERT INTO {definition.Table} ({string.Join(", ", fields)}) VALUES ({string.Join(", ", @fieldParameters)});
                                SET IDENTITY_INSERT {definition.Table} OFF;
                        ";
                    ExecuteNonQuery(sql, parameters);
                }

                try
                {
                    ExecuteNonQuery($@"
                    declare @lastId int
                    select @lastId = max(id) from {definition.Table} st
                    dbcc checkident('{definition.Table}', reseed, @lastId)
                ");
                } catch {}
            }
        }

        public void ExecuteAction(TestAction testAction)
        {
            
        }

        public bool ConvertType(Type type, ExcelRange excelRange, out object outputValue)
        {
            outputValue = null;
            return false;
        }
    }
    
    
}
