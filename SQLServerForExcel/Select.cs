using ExcelDna.Integration;
using System.Data.SqlClient;

namespace SQLServerForExcel
{
    public class Select
    {
        [ExcelFunction(IsThreadSafe = true, Description = "Select Scalar")]
        public static object SelectScalar(string sql)
        {
            string connectionString = "Data Source=DESKTOP-8M9VJFP\\SQLEXPRESS;Initial Catalog=PowerBI;Trusted_Connection=True;";

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
