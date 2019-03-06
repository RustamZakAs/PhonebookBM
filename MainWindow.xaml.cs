using System.Runtime.Serialization.Formatters.Binary;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Runtime.Serialization.Json;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Collections.ObjectModel;
using GalaSoft.MvvmLight.CommandWpf;
using System.Runtime.Serialization;
using System.Windows.Media.Imaging;
using System.Collections.Generic;
using System.Windows.Navigation;
using System.Xml.Serialization;
using System.Windows.Documents;
using System.Windows.Controls;
using System.Windows.Shapes;
using System.ComponentModel;
using System.Windows.Media;
using System.Windows.Input;
using System.Globalization;
using System.Windows.Data;
using System.Reflection;
using Microsoft.Win32;
using System.Runtime;
using System.Windows;
using System.Text;
using System.Linq;
using System.IO;
using System;

namespace PhonebookBM
{
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        private ObservableCollection<MyContact> ocMyContactsAll;
        public ObservableCollection<MyContact> OCMyContactsAll { get => ocMyContactsAll; set => Set(ref ocMyContactsAll, value); }

        private ObservableCollection<MyContact> ocMyContactsFiltered;
        public ObservableCollection<MyContact> OCMyContactsFiltered { get => ocMyContactsFiltered; set => Set(ref ocMyContactsFiltered, value); }

        private MyContact selectedContact;
        public MyContact SelectedContact { get => selectedContact; set => Set(ref selectedContact, value); }

        private string searchText;
        public string SearchText
        {
            get => searchText;
            set
            {
                Set(ref searchText, value);
                OCMyContactsFiltered = MySearch(OCMyContactsAll, value);
            }
        }

        private bool isChange;
        public bool IsChange { get => isChange; set => Set(ref isChange, value); }

        private int userStatus = 1; //User = 1 //Administrator = 0
        public int UserStatus
        {
            get => userStatus;
            set
            {
                Set(ref userStatus, value);
                if (value == 0) lbluser.Text = "Administrator";
                else lbluser.Text = "İstifadəçi";
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
        }

        public event PropertyChangedEventHandler PropertyChanged;
        public void Set<T>(ref T field, T value, [System.Runtime.CompilerServices.CallerMemberName]string prop = "")
        {
            field = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        }

        private ObservableCollection<MyContact> MySearch(ObservableCollection<MyContact> contacts, string value)
        {
            if (contacts != null && contacts.Count > 0)
            { 
                var linqResults1 = from user in contacts
                                   where user.ContactName.Contains(value) ||
                                   user.ContactSurname.Contains(value) ||
                                   user.Department.Contains(value) ||
                                   user.UnderDepartment.Contains(value)
                                   select user;
                return new ObservableCollection<MyContact>(linqResults1);
            }
            return new ObservableCollection<MyContact>();
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
                if (this.UserStatus == 1)
                {
                    this.UserStatus = 0;
                    MyIsEnabled = true;
                }
                else
                {
                    this.UserStatus = 1;
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
            if (ExcelFilePath != null && ExcelFilePath.Length > 0)
                OCMyContactsFiltered = OCMyContactsAll = myExcel.ReadExcel();

            MyJSON.Save(OCMyContactsAll);
        }

        private void MainWindow_Closing(object sender, CancelEventArgs e)
        {
            MyJSON.Save(OCMyContactsAll);
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            OCMyContactsAll = OCMyContactsFiltered = MyJSON.Load();

            if (OCMyContactsAll.Count == 0)
            {
                MyExcel excel = new MyExcel();
                OCMyContactsAll = OCMyContactsFiltered = excel.TestValue();
            }
        }
    }
}
