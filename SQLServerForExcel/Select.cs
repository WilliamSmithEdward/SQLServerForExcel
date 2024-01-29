using ExcelDna.Integration;
using System.Data.SqlClient;

namespace SQLServerForExcel
{
    public class Select
    {
        [ExcelFunction(IsThreadSafe = true, Description = "Select Scalar")]
        public static object SelectScalar(
            [ExcelArgument(Name = "connectionString", Description = "SQL Server connection string")]
            string connectionString,
            [ExcelArgument(Name = "sql", Description = "SQL string to be executed against the server")]
            string sql
        )
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                // Create SqlCommand
                using (SqlCommand command = new SqlCommand(sql, connection))
                {
                    connection.Open();

                    var result = command.ExecuteScalar();

                    if (result is null) return ExcelError.ExcelErrorValue;
                    else return result.ToString() ?? string.Empty;
                }
            }
        }
    }
}
