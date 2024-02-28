using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using MySql.Data.MySqlClient;

namespace ExcelPractice
{
    public class DbConnectionFactory
    {
        private readonly string connectionString;

        public DbConnectionFactory(string connectionString)
        {
            this.connectionString = connectionString;
        }

        public IDbConnection CreateConnection()
        {
            // Use MySqlConnection for MySQL
            return new MySqlConnection(connectionString);
        }
    }
}
