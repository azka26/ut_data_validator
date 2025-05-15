using System.Data;
using System.Data.Common;
using Microsoft.Data.SqlClient;
using OfficeOpenXml;
using UTDataValidator;

namespace UnitTest;

public abstract class SqlServerUnitTestBase : ExcelValidationUTBase, IDisposable
{
    protected abstract string GetSqlServerConnectionString();
    private DbConnection? _dbConnection;
    private DbConnection GetConnection()
    {
        if (_dbConnection == null)
        {
            _dbConnection = new SqlConnection(GetSqlServerConnectionString());
        }

        if (_dbConnection.State != ConnectionState.Open)
        {
            _dbConnection.Open();
        }

        return _dbConnection;
    }

    protected override void ResetConnection()
    {
        if (_dbConnection != null)
        {
            _dbConnection.Close();
            _dbConnection.Dispose();
            _dbConnection = null;
        }
    }

    #region Assertions
    public override void AreEqual<T>(T expected, T actual, string message)
    {
        var dateTimeFormat = "{0:yyyy-MM-dd HH:mm:ss}";
        if (expected != null && actual != null)
        {
            if (expected is DateTime expectedDate && actual is DateTime actualDate)
            {
                var expectedString = string.Format(dateTimeFormat, expectedDate);
                var actualString = string.Format(dateTimeFormat, actualDate);
                Assert.True(expectedString == actualString, message);
                return;
            }
        }

        Assert.True(expected.Equals(actual), message);
    }

    public override void IsTrue(bool condition, string message)
    {
        Assert.True(condition, message);
    }
    #endregion

    public override bool ConvertType(Type type, ExcelRange excelRange, out object outputValue)
    {
        if (type == typeof(string))
        {
            outputValue = excelRange.GetValue<string>();
            return true;
        }

        if (type == typeof(DateTime))
        {
            outputValue = excelRange.GetValue<DateTime>();
            return true;
        }

        if (type == typeof(int) || type == typeof(Int32))
        {
            outputValue = excelRange.GetValue<Int32>();
            return true;
        }

        if (type == typeof(Int64))
        {
            outputValue = excelRange.GetValue<Int64>();
            return true;
        }

        if (type == typeof(double))
        {
            outputValue = excelRange.GetValue<double>();
            return true;
        }

        if (type == typeof(float))
        {
            outputValue = excelRange.GetValue<float>();
            return true;
        }

        if (type == typeof(bool))
        {
            outputValue = excelRange.GetValue<bool>();
            return true;
        }

        outputValue = null;
        return false;
    }

    protected virtual int ExecuteNonQuery(string sql)
    {
        using var command = GetConnection().CreateCommand();
        command.CommandText = sql;
        return command.ExecuteNonQuery();
    }

    protected virtual DataTable ExecuteQuery(string sql)
    {
        using var command = GetConnection().CreateCommand();
        command.CommandText = sql;
        var dataTable = new DataTable();
        using var dbReader = command.ExecuteReader();
        dataTable.Load(dbReader);
        return dataTable;
    }

    protected virtual bool IsAutoIncrement(DataTable dataTable)
    {
        foreach (DataColumn column in dataTable.Columns)
        {
            if (column.AutoIncrement)
            {
                return true;
            }
        }

        return false;
    }

    public override void InitData(IEnumerable<ExcelDataDefinition> excelDataDefinition)
    {
        var listDefinitions = excelDataDefinition.ToList();

        var deleteSequences = listDefinitions.ToList();
        deleteSequences.Reverse();
        foreach (var item in deleteSequences)
        {
            ExecuteNonQuery($"DELETE FROM {item.Table}");
        }

        foreach (var itemDefinition in listDefinitions)
        {
            var isAutoIncrement = IsAutoIncrement(itemDefinition.Data);
            var keyValueMap = new Dictionary<string, string>();
            foreach (DataColumn column in itemDefinition.Data.Columns)
            {
                keyValueMap.Add(column.ColumnName, $"@{column.ColumnName}");
            }

            var sqlInsert = $@"
                INSERT INTO {itemDefinition.Table} ({string.Join(", ", keyValueMap.Keys.ToList())}) 
                    VALUES ({string.Join(", ", keyValueMap.Values.ToList())});
            ";

            var lastId = 0;
            foreach (DataRow dataRow in itemDefinition.Data.Rows)
            {
                using var command = GetConnection().CreateCommand();
                command.CommandType = CommandType.Text;

                var sqlList = new List<string>();
                if (isAutoIncrement)
                {
                    sqlList.Add($"SET IDENTITY_INSERT {itemDefinition.Table} ON");
                }

                sqlList.Add(sqlInsert);

                if (isAutoIncrement)
                {
                    sqlList.Add($"SET IDENTITY_INSERT {itemDefinition.Table} OFF");
                }

                foreach (DataColumn dataColumn in itemDefinition.Data.Columns)
                {
                    var sqlParameter =
                        new Microsoft.Data.SqlClient.SqlParameter($"@{dataColumn.ColumnName}", dataRow[dataColumn.ColumnName]);
                    command.Parameters.Add(sqlParameter);
                    if (dataColumn.AutoIncrement)
                    {
                        var id = Convert.ToInt32(dataRow[dataColumn.ColumnName]);
                        if (lastId < id)
                        {
                            lastId = id;
                        }
                    }
                }

                command.CommandText = string.Join(Environment.NewLine, sqlList);
                command.ExecuteNonQuery();
            }

            if (isAutoIncrement)
            {
                using var command = GetConnection().CreateCommand();
                command.CommandType = CommandType.Text;
                // kasih command untuk update seed
                command.CommandText = $"DBCC CHECKIDENT ('{itemDefinition.Table}', RESEED, {lastId});";
                command.ExecuteNonQuery();
            }
        }
    }

    public override DataTable ReadTable(ExcelDataDefinition excelDataDefinition)
    {
        var fields = excelDataDefinition
                .ColumnMapping
                .Select(f => f.Value.ColumnName)
                .ToList();

        var sql = $"SELECT {string.Join(", ", fields)} FROM {excelDataDefinition.Table}";

        var dataTable = new DataTable();
        using var command = GetConnection().CreateCommand();
        command.CommandType = CommandType.Text;
        command.CommandText = sql;
        using var reader = command.ExecuteReader();
        dataTable.Load(reader);
        reader.Close();

        return dataTable;
    }

    public void Dispose()
    {
        if (_dbConnection != null)
        {
            if (_dbConnection.State != ConnectionState.Closed)
            {
                _dbConnection.Close();
            }

            _dbConnection.Dispose();
            _dbConnection = null; 
        }
    }
}
