using System;
using System.Data;
using Microsoft.Data.SqlClient;

namespace UnitTest.SampleService;

public class SampleService : ISampleService
{
    private readonly SqlConnection _sqlConnection;
    public SampleService(SqlConnection sqlConnection)
    {
        _sqlConnection = sqlConnection;
        if (_sqlConnection.State != ConnectionState.Open)
        {
            _sqlConnection.Open();
        }
    }

    private readonly List<string> _errors = new List<string>();
    public async Task DoSomethingAsync(string input)
    {
        if (input == "error_1")
        {
            _errors.Add("Error Message 1");
            return;
        }

        if (input == "error_2")
        {
            _errors.Add("Error Message 1");
            _errors.Add("Error Message 2");
            return;
        }

        if (input == "error_3")
        {
            _errors.Add("Error Message 1");
            _errors.Add("Error Message 2");
            _errors.Add("Error Message 3");
            return;
        }

        using (var command = _sqlConnection.CreateCommand())
        {
            command.CommandText = "UPDATE item set processed = 1 where 1=1;";
            command.CommandType = CommandType.Text;
            command.ExecuteNonQuery();
        }

        return;
    }

    public List<string> GetErrors()
    {
        return _errors;
    }
}
