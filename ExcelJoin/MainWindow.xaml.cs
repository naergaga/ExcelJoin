using ExcelJoin.Actions;
using ExcelJoin.Models;
using ExcelJoin.Providers;
using ExcelJoin.Providers.EDR;
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
        private int sheet1Pos, sheet2Pos;
        private Book bookItem1, bookItem2;
        private OpenFileDialog openFileDialog;
        private ExportConfig config = new ExportConfig { DateTimeIsHourMinute = true };
        private MainWindowModel DC => (MainWindowModel)DataContext;

        public delegate void WindowLoaded();

        public MainWindow()
        {
            openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel文件 (*.xlsx,*.xls)|*.xlsx;*.xls";

            InitializeComponent();
        }

        private void SelectSheet_Selected(object sender, RoutedEventArgs e)
        {
            var select = sender as ComboBox;
            var selectIndex = select.SelectedIndex;
            if (selectIndex == -1) { Debug.WriteLine("选择为null"); return; }
            if (select == this.SelectSheet1)
            {
                sheet1Pos = selectIndex + 1;
            }
            else if (select == this.SelectSheet2)
            {
                sheet2Pos = selectIndex + 1;
            }
        }

        private void btnJoin_Click(object sender, RoutedEventArgs e)
        {
            var sp = new EDRSheetProvider();
            var sheet1 = sp.Get(this.InputPath1.Text, sheet1Pos,DC.ColumnIndex1);
            var sheet2 = sp.Get(this.InputPath2.Text, sheet2Pos, DC.ColumnIndex2);
            JoinAction action = new JoinAction(config);
            action.Export(sheet1, sheet2, InputPath3.Text, inputSheetName.Text, DC.HeadTitle1, DC.HeadTitle2);
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

        private void CheckBox_Click(object sender, RoutedEventArgs e)
        {
            if (sender == this.CbTimeFormat1)
            {
                this.CbTimeFormat2.IsChecked = false;
            }
            else if (sender == this.CbTimeFormat2)
            {
                this.CbTimeFormat1.IsChecked = false;
            }
            this.config.DateTimeIsHourMinute = this.CbTimeFormat1.IsChecked == true;
        }

        private void CheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            if (sender == this.CbTimeFormat1)
            {
                this.CbTimeFormat2.IsChecked = true;
            }
            else if (sender == this.CbTimeFormat2)
            {
                this.CbTimeFormat1.IsChecked = true;
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel文件 (*.xlsx)|*.xlsx";
            if (saveFileDialog.ShowDialog(this) != true) { return; }
            this.InputPath3.Text = saveFileDialog.FileName;
        }

        private void btnChoose1_Click(object sender, RoutedEventArgs e)
        {
            if (openFileDialog.ShowDialog(this) != true) { return; }
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
            var bp = new EDRBookProvider();
            if (!File.Exists(fileName)) { return; }
            var bookItem = bp.GetSimple(fileName);
            ComboBox select = null;
            switch (ins)
            {
                case ComboBoxIns.ComboBox1:
                    //book1 = workbook;
                    bookItem1 = bookItem;
                    select = SelectSheet1;
                    break;
                case ComboBoxIns.ComboBox2:
                    //book2 = workbook;
                    bookItem2 = bookItem;
                    select = SelectSheet2;
                    break;
                default:
                    break;
            }
            this.tbBookInfo.Text += InfoProvider.GetBook(bookItem)+'\n';
            select.ItemsSource = bookItem.Sheets;
        }

        private void btnChoose2_Click(object sender, RoutedEventArgs e)
        {
            if (openFileDialog.ShowDialog(this) != true) { return; }
            this.InputPath2.Text = openFileDialog.FileName;
        }
    }

    enum ComboBoxIns
    {
        ComboBox1, ComboBox2
    }
}
