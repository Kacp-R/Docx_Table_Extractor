using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Win32;

namespace TableExtractor
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        private void SelectFileButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Word Documents|*.docx";
            openFileDialog.Title = "Select a DOCX File";

            if (openFileDialog.ShowDialog() == true)
            {
                FilePathTextBox.Text = openFileDialog.FileName;
            }
        }
        private void ExtractTablesButton_Click(object sender, RoutedEventArgs e)
        {
            string filePath = FilePathTextBox.Text;
            if (File.Exists(filePath))
            {
                if (Path.GetExtension(filePath) == ".docx")
                {
                    ClearAll();
                    using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, false))
                    {
                        var tables = doc.MainDocumentPart.Document.Body.Elements<Table>();

                        foreach (var table in tables)
                        {
                            var tableData = ExtractTableData(table);
                            AddArray(tableData);
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Złe rozszerzenie!");
                }
            }
            else
            {
                MessageBox.Show("Plik nie istnieje!");
            }
        }

        private string[,] ExtractTableData(Table table)
        {
            var rows = table.Elements<TableRow>().ToList();
            int rowCount = rows.Count;
            int colCount = rows.Max(r => r.Elements<TableCell>().Count());
            string[,] tableData = new string[rowCount, colCount];

            for (int i = 0; i < rowCount; i++)
            {
                var cells = rows[i].Elements<TableCell>().ToList();
                for (int j = 0; j < colCount; j++)
                {
                    if (j < cells.Count)
                    {
                        tableData[i, j] = GetTableCellText(cells[j]);
                    }
                    else
                    {
                        tableData[i, j] = "";
                    }
                }
            }

            return tableData;
        }
        private string GetTableCellText(TableCell cell)
        {
            var texts = cell.Descendants<Text>().Select(t => t.Text);
            return string.Join(Environment.NewLine, texts);
        }

        private void AddArray(string[,] data)
        {
            Grid grid = new Grid();
            grid.Margin = new Thickness(1);
            for (int i = 0; i < data.GetLength(0); i++)
            {
                grid.RowDefinitions.Add(new RowDefinition());
            }

            for (int j = 0; j < data.GetLength(1); j++)
            {
                grid.ColumnDefinitions.Add(new ColumnDefinition());
            }
            for (int i = 0; i < data.GetLength(0); i++)
            {
                for (int j = 0; j < data.GetLength(1); j++)
                {
                    System.Windows.Controls.Border cellBorder = new System.Windows.Controls.Border();
                    cellBorder.BorderThickness = new Thickness(1);
                    cellBorder.BorderBrush = Brushes.Black;
                    cellBorder.Margin = new Thickness(1);

                    TextBlock textBlock = new TextBlock
                    {
                        Text = data[i, j],
                        VerticalAlignment = VerticalAlignment.Center,
                        HorizontalAlignment = HorizontalAlignment.Center
                    };
                    cellBorder.Child = textBlock;
                    grid.Children.Add(cellBorder);
                    Grid.SetRow(cellBorder, i);
                    Grid.SetColumn(cellBorder, j);
                }
            }
            System.Windows.Controls.Border gridBorder = new System.Windows.Controls.Border();
            gridBorder.BorderThickness = new Thickness(2);
            gridBorder.BorderBrush = Brushes.Black;
            gridBorder.Margin = new Thickness(5);
            gridBorder.Child = grid;

            ArrayMain.Children.Add(gridBorder);
        }

        private void ClearAll()
        {
            ArrayMain.Children.Clear();
        }

    }
}
