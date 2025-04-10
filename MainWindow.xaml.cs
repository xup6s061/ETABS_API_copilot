using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Metadata.W3cXsd2001;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ETABS_API_copilot.Models;
using ETABS_API_copilot.ViewModels; // 引用包含 MainViewModel 的命名空間

namespace ETABS_API_copilot
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            DataContext = new MainViewModel(); // 設置 DataContext
        }

        private void DataGrid_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {

        }
        private void DataGrid_KeyDown(object sender, KeyEventArgs e)
        {
            var dataGrid = sender as DataGrid;

            if (dataGrid == null) return;

            // 複製操作 (Ctrl+C)
            if (e.Key == Key.C && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                CopySelectedCells(dataGrid);
                e.Handled = true;
            }

            // 貼上操作 (Ctrl+V)
            if (e.Key == Key.V && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                PasteData(dataGrid);
                e.Handled = true;
            }
        }

        private void CopySelectedCells(DataGrid dataGrid)
        {
            var selectedCells = dataGrid.SelectedCells;
            if (selectedCells.Count == 0) return;

            var clipboardText = new System.Text.StringBuilder();
            var rowGroups = selectedCells.GroupBy(cell => cell.Item);

            foreach (var rowGroup in rowGroups)
            {
                var rowText = string.Join("\t", rowGroup.Select(cell =>
                {
                    var binding = (cell.Column as DataGridBoundColumn)?.Binding as System.Windows.Data.Binding;
                    var propertyName = binding?.Path.Path;
                    var property = cell.Item.GetType().GetProperty(propertyName);
                    return property?.GetValue(cell.Item)?.ToString() ?? string.Empty;
                }));
                clipboardText.AppendLine(rowText);
            }

            Clipboard.SetText(clipboardText.ToString());
        }

        private void PasteData(DataGrid dataGrid)
        {
            if (dataGrid.SelectedCells.Count == 0 && !dataGrid.CanUserAddRows) return;

            // 從剪貼簿獲取資料
            var clipboardText = Clipboard.GetText();
            if (string.IsNullOrEmpty(clipboardText)) return;

            var rows = clipboardText.Split(new[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);

            // 根據 DataGrid 的 ItemsSource 類型動態處理
            if (dataGrid.ItemsSource is IList<Material> materialCollection)
            {
                PasteRowsToCollection(rows, materialCollection);
            }
            else if (dataGrid.ItemsSource is IList<Building> buildingCollection)
            {
                PasteRowsToCollection(rows, buildingCollection);
            }
            else if (dataGrid.ItemsSource is IList<SectionProperty> sectionCollection)
            {
                PasteRowsToCollection(rows, sectionCollection);
            }

            // 通知 UI 更新
            dataGrid.Items.Refresh();
        }

        private void PasteRowsToCollection<T>(string[] rows, IList<T> collection) where T : new()
        {
            foreach (var row in rows)
            {
                var values = row.Split('\t');
                if (values.Length == 0) continue;

                var newItem = new T();
                var properties = newItem.GetType().GetProperties();

                for (int i = 0; i < values.Length && i < properties.Length; i++)
                {
                    var property = properties[i];
                    if (property.CanWrite)
                    {
                        var convertedValue = Convert.ChangeType(values[i], property.PropertyType);
                        property.SetValue(newItem, convertedValue);
                    }
                }

                collection.Add(newItem);
            }
        }


    }
}
