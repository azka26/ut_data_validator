using System.Data;
using System.Data.Common;
using Npgsql;
using OfficeOpenXml;

namespace UTDataValidator.PostgreSql;

public abstract class PostgreSqlUnitTestBase : ExcelValidationUTBase, IDisposable
{
    private DbConnection? _dbConnection;
    private DbConnection GetConnection()
    {
        if (_dbConnection == null)
        {
            _dbConnection = new NpgsqlConnection(GetConnectionString());
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

    protected virtual bool IsAutoIncrement(DataTable dataTable, string tableName)
    {
        // First, check the standard AutoIncrement property
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
                    command.CommandText = $"TRUNCATE TABLE {item.Table} CASCADE";
                    command.ExecuteNonQuery();
                }
                catch
                {
                    // Fallback to DELETE if TRUNCATE fails
                    command.CommandText = $"DELETE FROM {item.Table}";
                    command.ExecuteNonQuery();
                }
            }

            // Insert data
            foreach (var itemDefinition in listDefinitions)
            {
                if (itemDefinition.Data.Rows.Count == 0)
                    continue;

                var isAutoIncrement = IsAutoIncrement(itemDefinition.Data, itemDefinition.Table);
                var columns = new List<string>();
                foreach (DataColumn column in itemDefinition.Data.Columns)
                {
                    columns.Add(column.ColumnName);
                }

                var parameterNames = columns.Select(c => $"@{c}").ToList();
                var sqlInsert = $@"INSERT INTO {itemDefinition.Table} ({string.Join(", ", columns)}) 
                                   VALUES ({string.Join(", ", parameterNames)})";

                var lastId = 0;
                
                // Prepare command once, reuse for all rows
                using var command = connection.CreateCommand();
                command.Transaction = transaction;
                command.CommandText = sqlInsert;
                command.CommandType = CommandType.Text;

                // Add parameters (will be updated with values in loop)
                foreach (var columnName in columns)
                {
                    command.Parameters.Add(new NpgsqlParameter($"@{columnName}", DBNull.Value));
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
                        if (itemDefinition.Data.Columns[columnName].AutoIncrement || 
                            (isAutoIncrement && columnName.ToLower() == "id"))
                        {
                            if (value != null && value != DBNull.Value)
                            {
                                var id = Convert.ToInt32(value);
                                if (lastId < id)
                                {
                                    lastId = id;
                                }
                            }
                        }
                    }

                    command.ExecuteNonQuery();
                }

                // Reset sequence if needed
                if (isAutoIncrement && lastId > 0)
                {
                    // Get actual sequence name from database
                    var sequenceQuery = $@"
                        SELECT pg_get_serial_sequence('{itemDefinition.Table}', 'id') as seq_name";
                    
                    using var seqCommand = connection.CreateCommand();
                    seqCommand.Transaction = transaction;
                    seqCommand.CommandText = sequenceQuery;
                    
                    var seqName = seqCommand.ExecuteScalar()?.ToString();
                    if (!string.IsNullOrEmpty(seqName))
                    {
                        using var reseedCommand = connection.CreateCommand();
                        reseedCommand.Transaction = transaction;
                        reseedCommand.CommandText = $"SELECT setval('{seqName}', {lastId}, true)";
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
        finally {
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

        sql = $@"
                SELECT column_name
                FROM information_schema.columns
                WHERE table_name = '{excelDataDefinition.Table}' AND table_schema = 'public' AND is_identity = 'YES'
                LIMIT 1;
        ";
        var resultTablePk = ExecuteQuery(sql);
        if (resultTablePk.Rows.Count > 0 && resultTablePk.Rows[0]["column_name"] != null)
        {
            var pkColumn = resultTablePk.Rows[0]["column_name"].ToString();
            if (!string.IsNullOrEmpty(pkColumn) && dataTable.Columns.Contains(pkColumn))
            {
                var column = dataTable.Columns[pkColumn];
                if (column != null)
                {
                    column.AutoIncrement = true;
                    column.AutoIncrementSeed = 1;
                    column.AutoIncrementStep = 1;
                }
            }
        }

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
