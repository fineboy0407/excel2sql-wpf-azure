using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace ExcelToDatabase
{
    /// <summary>
    /// Interaction logic for SourceCodeWindow.xaml
    /// </summary>
    public partial class SourceCodeWindow : Window
    {
        public SourceCodeWindow()
        {
            InitializeComponent();
            codeBox.Text = @"
    // CODE BY (ROMAN MATVIIENKO  matviienkoroman366@gmail.com)

    internal class DataManipulator
    {
        private async Task ReadAllExcelFilesInDirectory(string path, List<User> usersList)
        {
            var dirInfo = new DirectoryInfo(path);
            var files = dirInfo.GetFiles(""*.xlsx"");
            foreach (var file in files)
                await ReadExcelFile(file, usersList);
        }

        private async Task ReadExcelFile(FileInfo fileName, List<User> usersList)
        {
            if (fileName == null) throw new ArgumentNullException(nameof(fileName));
            if (usersList == null) throw new ArgumentNullException(nameof(usersList));

            var package = new ExcelPackage(fileName);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var currentSheet = package.Workbook.Worksheets;
            var workSheet = currentSheet.First();
            var noOfCol = workSheet.Dimension.End.Column;
            var noOfRow = workSheet.Dimension.End.Row;
            for (var rowIterator = 2; rowIterator <= noOfRow; rowIterator++)
            {
                var user = new User
                {
                    Name = workSheet.Cells[rowIterator, 1].Value?.ToString(),
                    Email = workSheet.Cells[rowIterator, 3].Value?.ToString(),
                };


                usersList.Add(user);
            }
        }

        private async Task InsertUsersIntoDatabase(List<User> usersList)
        {
            var connString = ConfigurationManager.ConnectionStrings[""Development""].ConnectionString;
            var conn = new SqlConnection(connString);
            await conn.OpenAsync();
            await InsertAsync(usersList, ""[GameAPI].[dbo].[A101Orders]"", conn, CancellationToken.None);
        }


        private async Task InsertAsync<T>(IEnumerable<T> items, string tableName, SqlConnection connection,
                CancellationToken cancellationToken)
        {
            var bulk = new SqlBulkCopy(connection);
            var reader = ObjectReader.Create(items);
            bulk.DestinationTableName = tableName;
            var properties = typeof(T).GetProperties();
            foreach (var prop in properties)
                bulk.ColumnMappings.Add(new SqlBulkCopyColumnMapping(prop.Name, prop.Name));
            await bulk.WriteToServerAsync(reader, cancellationToken);
        }

        private async Task InsertAsync(List<User> users)
        {
            using (var conn = new MySqlConnection(builder.ConnectionString))
            {
                // Opening connection
                await conn.OpenAsync();

                using (var command = conn.CreateCommand())
                {
                    // Drop Table (if exists)
                    command.CommandText = ""DROP TABLE IF EXISTS inventory;"";
                    await command.ExecuteNonQueryAsync();

                    // Create Table
                    command.CommandText = ""CREATE TABLE inventory (id serial PRIMARY KEY, email VARCHAR(50), name VARCHAR(50));"";
                    await command.ExecuteNonQueryAsync();

                    string baseQueryString = ""INSERT INTO inventory (name, email) VALUES"";

                    for(var i=1; i<=users.Count; i++)
                    {
                        if(i!=users.Count)
                            baseQueryString += string.Format("" (@name{0}, @email{0}),"", i);
                        else
                            baseQueryString += string.Format("" (@name{0}, @email{0});"", i);

                    }

                    command.CommandText = baseQueryString;

                    for (var i = 1; i <= users.Count; i++)
                    {
                        command.Parameters.AddWithValue(string.Format(""name{0}"", i), users[i].Name);
                        command.Parameters.AddWithValue(string.Format(""email{0}"", i), users[i].Email);
                    }

                }

                // Close connection
            }
        }

        // MySQL connection in Azure
        private MySqlConnectionStringBuilder builder = new MySqlConnectionStringBuilder
        {
            Server = ""YOUR-SERVER.mysql.database.azure.com"",
            Database = ""YOUR-DATABASE"",
            UserID = ""USER@YOUR-SERVER"",
            Password = ""PASSWORD"",
            SslMode = MySqlSslMode.Required,
        };

        // Main Entry Method
        public async void ReadAndInsert(string path)
        {
            try
            {
                var usersList = new List<User>();
                await ReadAllExcelFilesInDirectory(path, usersList);
                await InsertUsersIntoDatabase(usersList);
                MessageBox.Show(""Done"", ""Success"");
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, ""Error"");
                throw;
            }
        }

    }";
        }
    }
}
