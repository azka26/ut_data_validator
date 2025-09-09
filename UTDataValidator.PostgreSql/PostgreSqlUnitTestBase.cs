using System.Data;
using System.Data.Common;
using Npgsql;
using OfficeOpenXml;

namespace UTDataValidator.PostgreSql;

public abstract class PostgreSqlUnitTestBase : ExcelValidationUTBase, IDisposable
{
    protected abstract string GetPostgreSqlConnectionString();
    private DbConnection? _dbConnection;
    private DbConnection GetConnection()
    {
        if (_dbConnection == null)
        {
            _dbConnection = new NpgsqlConnection(GetPostgreSqlConnectionString());
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

        var deleteSequences = listDefinitions.ToList();
        deleteSequences.Reverse();
        foreach (var item in deleteSequences)
        {
            ExecuteNonQuery($"DELETE FROM {item.Table}");
        }

        foreach (var itemDefinition in listDefinitions)
        {
            var isAutoIncrement = IsAutoIncrement(itemDefinition.Data, itemDefinition.Table);
            var keyValueMap = new Dictionary<string, string>();
            var allColumns = new List<string>();
            foreach (DataColumn column in itemDefinition.Data.Columns)
            {
                keyValueMap.Add(column.ColumnName, $"@{column.ColumnName}");
                allColumns.Add(column.ColumnName);
            }

            var overridingClause = ""; // Removed OVERRIDING for compatibility
            var sqlInsert = $@"
                INSERT INTO {itemDefinition.Table} ({string.Join(", ", allColumns)}) 
                    VALUES ({string.Join(", ", keyValueMap.Values.ToList())}){overridingClause};
            ";

            var lastId = 0;
            foreach (DataRow dataRow in itemDefinition.Data.Rows)
            {
                using var command = GetConnection().CreateCommand();
                command.CommandType = CommandType.Text;

                command.CommandText = sqlInsert;

                foreach (DataColumn dataColumn in itemDefinition.Data.Columns)
                {
                    var sqlParameter = new NpgsqlParameter($"@{dataColumn.ColumnName}", dataRow[dataColumn.ColumnName]);
                    command.Parameters.Add(sqlParameter);
                    if (dataColumn.AutoIncrement || (isAutoIncrement && dataColumn.ColumnName.ToLower() == "id"))
                    {
                        var id = Convert.ToInt32(dataRow[dataColumn.ColumnName]);
                        if (lastId < id)
                        {
                            lastId = id;
                        }
                    }
                }

                command.ExecuteNonQuery();
            }

            if (isAutoIncrement && lastId > 0)
            {
                // Assuming sequence name is table_column_seq, common in PostgreSQL
                var sequenceName = $"{itemDefinition.Table}_id_seq"; // Adjust if different
                using var command = GetConnection().CreateCommand();
                command.CommandType = CommandType.Text;
                command.CommandText = $"ALTER SEQUENCE {sequenceName} RESTART WITH {lastId + 1};";
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
