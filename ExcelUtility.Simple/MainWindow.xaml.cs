using ExcelUtility.Simple.Tables;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Windows;

namespace ExcelUtility.Simple
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private UserTable UserTable { get; } = new UserTable();
        public DataTable Table { get; set; }

        public MainWindow()
        {
            InitializeComponent();
            Table = StrategyHelper.FieldsToTable(UserTable.Datas);
            DataContext = this;
            ExcelHelper.WriteToExcelAsync(Path.Combine(Environment.CurrentDirectory, "1.xlsx"), new List<ExcelUtility.Model.Base.IExcelSheet>() { UserTable });
        }
    }
}