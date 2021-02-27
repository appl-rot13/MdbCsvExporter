
namespace MdbCsvExporter
{
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.Data.OleDb;
    using System.IO;
    using System.Linq;
    using System.Windows;

    public partial class MainWindow
    {
        public MainWindow()
        {
            this.InitializeComponent();
        }

        protected override void OnPreviewDragOver(DragEventArgs e)
        {
            base.OnPreviewDragOver(e);

            e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
            e.Handled = true;
        }

        protected override void OnPreviewDrop(DragEventArgs e)
        {
            base.OnPreviewDrop(e);

            if (!e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                return;
            }

            var filePaths = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (filePaths == null)
            {
                return;
            }

            foreach (var filePath in filePaths)
            {
                try
                {
                    ExportCsv(filePath);
                }
                catch (Exception exception)
                {
                    MessageBox.Show($"{exception}");
                }
            }
        }

        private static void ExportCsv(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath) || Path.GetExtension(filePath) != ".mdb")
            {
                return;
            }

            var outputDirectoryPath = Path.GetDirectoryName(filePath);
            if (string.IsNullOrWhiteSpace(outputDirectoryPath))
            {
                return;
            }

            var connectionStringBuilder = new OleDbConnectionStringBuilder
            {
                Provider = "Microsoft.Jet.OLEDB.4.0",
                DataSource = filePath
            };

            using (var connection = new OleDbConnection($"{connectionStringBuilder}"))
            {
                connection.Open();

                var schema = connection.GetSchema("Tables", new[] { null, null, null, "TABLE" });
                var tableNames = schema.Rows.OfType<DataRow>().Select(row => row["TABLE_NAME"]);
                foreach (var tableName in tableNames)
                {
                    var table = new DataTable();
                    using (var command = new OleDbCommand($"SELECT * FROM {tableName}", connection))
                    using (var reader = command.ExecuteReader())
                    {
                        table.Load(reader);
                    }

                    var outputTableData = new List<IEnumerable<string>>();
                    var columnNames = table.Columns.OfType<DataColumn>().Select(column => column.ColumnName);
                    outputTableData.Add(columnNames);

                    foreach (var row in table.Rows.OfType<DataRow>())
                    {
                        outputTableData.Add(row.ItemArray.Select(value => $"{value}"));
                    }

                    var outputFilePath = Path.Combine(outputDirectoryPath, $"{tableName}.csv");
                    using (var writer = new StreamWriter(outputFilePath))
                    {
                        foreach (var row in outputTableData)
                        {
                            writer.WriteLine(string.Join(",", row));
                        }
                    }
                }
            }
        }
    }
}
