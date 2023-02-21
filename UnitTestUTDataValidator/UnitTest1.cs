using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using UTDataValidator;
using System.Data.Common;
using System.Data.SqlClient;
using OfficeOpenXml;

namespace UnitTestProject1
{
    public class ExcelValidatorTest : IEventExcelValidator
    {
        private string ConnectionString = "Data Source=localhost,1433;Initial Catalog=UnitTest;User ID=sa;Password=Strong@Password;";
        [Test()]
        public void ReadExcelData_Test()
        {
            var sqlDelete = "DELETE FROM SampleTable";
            var dataList = new List<string>()
            {
                $@" SET IDENTITY_INSERT dbo.SampleTable ON;  
                    INSERT INTO SampleTable (Id, Name, CreatedDate, CreatedTime, CreatedDateTime, ValueInt, ValueDouble) 
                    VALUES ('0', 'Rizal', '2022-01-01', '00:00:00', '2022-01-01 00:00:00.000','20','20.1');
                    SET IDENTITY_INSERT dbo.SampleTable OFF;
                ",
                $@" SET IDENTITY_INSERT dbo.SampleTable ON;
                    INSERT INTO SampleTable (Id, Name, CreatedDate, CreatedTime, CreatedDateTime, ValueInt, ValueDouble) 
                    VALUES ('1', 'Rosa', '2022-01-02', '00:01:00', '2022-01-02 00:01:00.000','21','20.2');
                    SET IDENTITY_INSERT dbo.SampleTable OFF;
                ",
                $@" SET IDENTITY_INSERT dbo.SampleTable ON;
                    INSERT INTO SampleTable (Id, Name, CreatedDate, CreatedTime, CreatedDateTime, ValueInt, ValueDouble) 
                    VALUES ('2', 'Andika', '2022-01-03', '00:02:00', '2022-01-03 00:02:00.000','22','20.3');
                    SET IDENTITY_INSERT dbo.SampleTable OFF;
                "
            };

            ExecuteNonQuery(sqlDelete);
            foreach (var sqlInsert in dataList)
            {
                ExecuteNonQuery(sqlInsert);
            }

            var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "excel", "sample-tipe-data.xlsx");
            ExcelValidator excelValidator = new ExcelValidator(new Assertion(), path, "init", "init", this);
            excelValidator.Validate();
        }

        private int ExecuteNonQuery(string sql)
        {
            using (var sqlConnection = new SqlConnection(ConnectionString))
            {
                sqlConnection.Open();
                using (var sqlCommand = sqlConnection.CreateCommand())
                {
                    sqlCommand.CommandType = CommandType.Text;
                    sqlCommand.CommandText = sql;
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
            DataTable dataTable = new DataTable();
            using (var sqlConnection = new SqlConnection(ConnectionString))
            {
                sqlConnection.Open();
                using (var sqlCommand = sqlConnection.CreateCommand())
                {
                    sqlCommand.CommandType = CommandType.Text;
                    sqlCommand.CommandText = "SELECT * FROM " + excelDataDefinition.Table;
                    using (var reader = sqlCommand.ExecuteReader())
                    {
                        dataTable.Load(reader);
                    }
                }
            }

            return dataTable;
        }

        public void InitData(IEnumerable<ExcelDataDefinition> excelDataDefinition)
        {
            
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
