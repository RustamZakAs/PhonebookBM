using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Runtime;
using Microsoft.Win32;
using System.Reflection;
using System.Windows.Data;
using System.Globalization;
using System.Windows.Input;
using System.Windows.Media;
using System.ComponentModel;
using System.Windows.Shapes;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Xml.Serialization;
using System.Windows.Navigation;
using System.Collections.Generic;
using System.Windows.Media.Imaging;
using System.Runtime.Serialization;
using System.Collections.ObjectModel;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.Serialization.Formatters.Binary;

namespace PhonebookBM
{
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        private ObservableCollection<MyContact> ocMyContacts;
        public ObservableCollection<MyContact> OCMyContacts { get => ocMyContacts; set => Set(ref ocMyContacts, value); }

        private ObservableCollection<MyContact> ocCMyContactsNotСonfirmed;
        public ObservableCollection<MyContact> OCMyContactsNotСonfirmed { get => ocCMyContactsNotСonfirmed; set => Set(ref ocCMyContactsNotСonfirmed, value); }

        private string searchText;
        public string SearchText
        {
            get => searchText;
            set
            {
                Set(ref searchText, value);
                MySearch(value);
            }
        }

        private int userStatus; //User //Administrator
        public int UserStatus
        {
            get => userStatus;
            set
            {
                Set(ref userStatus, value);
                if (value == 0) lbluser.Content = "İstifadəçi";
                else lbluser.Content = "Administrator";
            }
        }

        private bool myIsEnabled = false;
        public bool MyIsEnabled { get => myIsEnabled; set => Set(ref myIsEnabled, value); }

        private bool adminKeyPress;
        public bool AdminKeyPress { get => adminKeyPress; set => Set(ref adminKeyPress, value); }

        private string excelFilePath;
        public string ExcelFilePath { get => excelFilePath; set => Set(ref excelFilePath, value); }

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
            
            OCMyContacts = OCMyContactsNotСonfirmed;
        }

        public event PropertyChangedEventHandler PropertyChanged;
        public void Set<T>(ref T field, T value, [System.Runtime.CompilerServices.CallerMemberName]string prop = "")
        {
            field = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        }

        private void MySearch(string value)
        {
            var linqResults1 = from user in OCMyContactsNotСonfirmed
                                   //from lang in user.Languages
                               where user.ContactName.Contains(value) ||
                               user.ContactSurname.Contains(value) ||
                               user.Department.Contains(value) ||
                               user.UnderDepartment.Contains(value)
                               //where user.ContactSurname == "%" + value + "%"
                               select user;

            OCMyContacts = new ObservableCollection<MyContact>(linqResults1);
        }

        private void NumberInsert(object sender, TextCompositionEventArgs e)
        {
            if (e.Text != "." && e.Text != "+" && IsNumber(e.Text) == false)
                e.Handled = true;
            else if (e.Text == ".")
            {
                if (((TextBox)sender).Text.IndexOf(e.Text) > -1)
                    e.Handled = true;
            }
            else if (e.Text == "+")
            {
                if (((TextBox)sender).Text.IndexOf(e.Text) > -1)
                    e.Handled = true;
                if (((TextBox)sender).Text.Length > 1) e.Handled = true;
                if (((TextBox)sender).Text.StartsWith(e.Text) == true)
                    e.Handled = true;
            }
        }

        private bool IsNumber(string text)
        {
            int output;
            return int.TryParse(text, out output);
        }

        private void AdminKey(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.R & Keyboard.Modifiers == ModifierKeys.Control)
            { 
                AdminKeyPress = true;
            }
            else if ((e.Key == Key.Z & Keyboard.Modifiers == ModifierKeys.Control) && AdminKeyPress)
            {
                if (this.UserStatus == 0)
                {
                    this.UserStatus = 1;
                    MyIsEnabled = true;
                }
                else
                {
                    this.UserStatus = 0;
                    MyIsEnabled = false;
                }
                AdminKeyPress = false;
            }
            else AdminKeyPress = false;
        }

        private string OpenFileDialogAndReturnExcelFilePath()
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd = new OpenFileDialog();
            ofd.InitialDirectory = @"C:\";
            ofd.RestoreDirectory = true;
            ofd.Title = "Browse Excel Files";
            ofd.DefaultExt = "xls";
            ofd.Filter = "Excel files (*.xls)|*.xls|Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            ofd.FilterIndex = 2;
            ofd.CheckFileExists = true;
            ofd.CheckPathExists = true;
            ofd.ShowDialog();
            return ofd.FileName;
        }

        private void LoadFromExcel(object sender, RoutedEventArgs e)
        {
            ExcelFilePath = OpenFileDialogAndReturnExcelFilePath();
            MyExcel myExcel = new MyExcel(ExcelFilePath);
            OCMyContactsNotСonfirmed = myExcel.ReadExcel();
        }
    }
}
