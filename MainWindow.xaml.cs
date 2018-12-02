using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace PhonebookBM
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    /// 


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

        private bool isEnabled;
        public bool MyIsEnabled { get => isEnabled; set => Set(ref isEnabled, value); }


        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
            OCMyContactsNotСonfirmed = MyExcel.ReadExcel();
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
                               user.ContactSurname.Contains(value)
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
            if (e.Key == Key.F5)
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
            }

        }
    }
}
