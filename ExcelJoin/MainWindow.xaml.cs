using ExcelJoin.Actions;
using ExcelJoin.Models;
using ExcelJoin.Providers.Epplus;
using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace ExcelJoin
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        private Workbook book1, book2;
        private int sheet1Pos, sheet2Pos;
        private Book bookItem1, bookItem2;
        private OpenFileDialog openFileDialog;

        public delegate void WindowLoaded();

        public MainWindow()
        {
            openFileDialog = new OpenFileDialog();
            openFileDialog.DefaultExt = "Excel文件(*.xlsx)";

            InitializeComponent();

            //后台线程
            var _thread1 = new Thread(MainWindow1_Loaded);
            _thread1.IsBackground = true;
            _thread1.Start();
            //var load1 = new WindowLoaded(MainWindow1_Loaded);
            //Dispatcher.BeginInvoke(load1);
        }

        private void SelectSheet_Selected(object sender, RoutedEventArgs e)
        {
            var select = sender as ComboBox;
            var selectIndex = select.SelectedIndex;
            if (selectIndex == -1) { Debug.WriteLine("选择为null"); return; }
            if (select == this.SelectSheet1)
            {
                sheet1Pos = selectIndex+1;
            }
            else if (select == this.SelectSheet2)
            {
                sheet2Pos = selectIndex+1;
            }
        }

        private void btnJoin_Click(object sender, RoutedEventArgs e)
        {
            var sp = new SheetProvider(book1.Book.Worksheets[sheet1Pos], true);
            var sp2 = new SheetProvider(book1.Book.Worksheets[sheet2Pos], true);
            var outPath = InputPath3.Text;
            var sheetName = inputSheetName.Text;
            int col1, col2;
            if (!int.TryParse(InputCol1.Text, out col1) || !int.TryParse(InputCol1.Text, out col2))
            {
                return;
            }
            var sheet1 = sp.Get(col1);
            var sheet2 = sp2.Get(col2);
            JoinAction action = new JoinAction();
            action.Export(sheet1, sheet2, outPath, sheetName,true);
        }

        /// <summary>
        /// 路径 TextBox 变化，更新选择列表
        /// </summary>
        /// <param name="sender">路径 TextBox</param>
        /// <param name="e"></param>
        private void InputPath_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (sender == InputPath1)
            {
                UpdateSelect(ComboBoxIns.ComboBox1, InputPath1.Text);
            }
            else if (sender == InputPath2)
            {
                UpdateSelect(ComboBoxIns.ComboBox2, InputPath2.Text);
            }
        }

        /// <summary>
        /// 启动后 两秒 开始执行
        /// </summary>
        private void MainWindow1_Loaded()
        {
            Thread.Sleep(1000);
            Dispatcher.Invoke(() =>
            {
                this.InputPath1.Text = "./files/xlsx/class1.xlsx";
                this.InputPath2.Text = "./files/xlsx/class1.xlsx";
                this.InputPath3.Text = "./files/xlsx/test.xlsx";
                this.InputCol1.Text = "1";
                this.InputCol2.Text = "1";
                this.inputSheetName.Text = "result";
            });
        }

        private void btnChoose1_Click(object sender, RoutedEventArgs e)
        {
            if (openFileDialog.ShowDialog(this) != true) { return; }
            UpdateSelect(ComboBoxIns.ComboBox1, openFileDialog.FileName);
            this.InputPath1.Text = openFileDialog.FileName;
        }

        /// <summary>
        /// 读取excel，读出Sheet集合，设置workbook,bookItem
        /// 设置下拉列表选项
        /// </summary>
        /// <param name="ins">标识</param>
        /// <param name="fileName">要读取的excel文件</param>
        private void UpdateSelect(ComboBoxIns ins, string fileName)
        {
            var workbook = new Workbook(new FileInfo(fileName));
            var bp = new BookProvider(workbook.Book, true);
            var bookItem = bp.GetSimple();
            ComboBox select=null;
            switch (ins)
            {
                case ComboBoxIns.ComboBox1:
                    book1 = workbook;
                    bookItem1 = bookItem;
                    select = SelectSheet1;
                    break;
                case ComboBoxIns.ComboBox2:
                    book2 = workbook;
                    bookItem2 = bookItem;
                    select = SelectSheet2;
                    break;
                default:
                    break;
            }
            select.ItemsSource = bookItem.Sheets;
        }

        private void btnChoose2_Click(object sender, RoutedEventArgs e)
        {
            if (openFileDialog.ShowDialog(this) != true) { return; }
            UpdateSelect(ComboBoxIns.ComboBox2, openFileDialog.FileName);
            this.InputPath1.Text = openFileDialog.FileName;
        }
    }

    enum ComboBoxIns
    {
        ComboBox1,ComboBox2
    }
}
