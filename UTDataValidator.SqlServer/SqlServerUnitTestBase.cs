using System.Data;
using System.Data.Common;
using Microsoft.Data.SqlClient;
using OfficeOpenXml;

namespace UTDataValidator.SqlServer;

public abstract class SqlServerUnitTestBase : ExcelValidationUTBase, IDisposable
{
    private DbConnection? _dbConnection;
    private DbConnection GetConnection()
    {
        if (_dbConnection == null)
        {
            _dbConnection = new SqlConnection(GetConnectionString());
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
        var connection = GetConnection();
        
        // Start transaction for better performance
        using var transaction = connection.BeginTransaction();
        
        try
        {
            // Delete data in reverse order
            var deleteSequences = listDefinitions.ToList();
            deleteSequences.Reverse();
            foreach (var item in deleteSequences)
            {
                using var command = connection.CreateCommand();
                command.Transaction = transaction;
                
                // Use TRUNCATE if possible (faster), fallback to DELETE
                try
                {
                    command.CommandText = $"TRUNCATE TABLE {item.Table}";
                    command.ExecuteNonQuery();
                }
                catch
                {
                    // Fallback to DELETE if TRUNCATE fails (e.g., foreign keys)
                    command.CommandText = $"DELETE FROM {item.Table}";
                    command.ExecuteNonQuery();
                }
            }

            // Insert data
            foreach (var itemDefinition in listDefinitions)
            {
                if (itemDefinition.Data.Rows.Count == 0)
                    continue;

                var isAutoIncrement = IsAutoIncrement(itemDefinition.Data);
                var columns = new List<string>();
                foreach (DataColumn column in itemDefinition.Data.Columns)
                {
                    columns.Add(column.ColumnName);
                }

                var parameterNames = columns.Select(c => $"@{c}").ToList();
                var sqlInsert = $@"INSERT INTO {itemDefinition.Table} ({string.Join(", ", columns)}) 
                                   VALUES ({string.Join(", ", parameterNames)});";

                // Toggle IDENTITY_INSERT once per table (not per row)
                if (isAutoIncrement)
                {
                    using var identityCommand = connection.CreateCommand();
                    identityCommand.Transaction = transaction;
                    identityCommand.CommandText = $"SET IDENTITY_INSERT {itemDefinition.Table} ON";
                    identityCommand.ExecuteNonQuery();
                }

                var lastId = 0;
                
                // Prepare command once, reuse for all rows
                using var command = connection.CreateCommand();
                command.Transaction = transaction;
                command.CommandText = sqlInsert;
                command.CommandType = CommandType.Text;

                // Add parameters (will be updated with values in loop)
                foreach (var columnName in columns)
                {
                    command.Parameters.Add(new SqlParameter($"@{columnName}", SqlDbType.Variant));
                }

                // Batch insert rows
                foreach (DataRow dataRow in itemDefinition.Data.Rows)
                {
                    // Update parameter values
                    for (int i = 0; i < columns.Count; i++)
                    {
                        var columnName = columns[i];
                        var value = dataRow[columnName];
                        command.Parameters[$"@{columnName}"].Value = value ?? DBNull.Value;
                        
                        // Track last ID for auto-increment
                        if (itemDefinition.Data.Columns[columnName].AutoIncrement)
                        {
                            var id = Convert.ToInt32(value);
                            if (lastId < id)
                            {
                                lastId = id;
                            }
                        }
                    }

                    command.ExecuteNonQuery();
                }

                // Turn off IDENTITY_INSERT and reseed
                if (isAutoIncrement)
                {
                    using var identityOffCommand = connection.CreateCommand();
                    identityOffCommand.Transaction = transaction;
                    identityOffCommand.CommandText = $"SET IDENTITY_INSERT {itemDefinition.Table} OFF";
                    identityOffCommand.ExecuteNonQuery();

                    if (lastId > 0)
                    {
                        using var reseedCommand = connection.CreateCommand();
                        reseedCommand.Transaction = transaction;
                        reseedCommand.CommandText = $"DBCC CHECKIDENT ('{itemDefinition.Table}', RESEED, {lastId});";
                        reseedCommand.ExecuteNonQuery();
                    }
                }
            }

            transaction.Commit();
        }
        catch
        {
            transaction.Rollback();
            throw;
        }
        finally
        {
            ResetConnection();
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
