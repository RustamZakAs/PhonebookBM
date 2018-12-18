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
    public class MyContact : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        public void Set<T>(ref T field, T value, [System.Runtime.CompilerServices.CallerMemberName]string prop = "")
        {
            field = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        }

        private int id = 0;
        public int Id { get => id; set => Set(ref id, value); }

        private string departmentIcon = "";
        public string DepartmentIcon { get => departmentIcon; set => Set(ref departmentIcon, value); }

        private string department = "";
        public string Department { get => department; set => Set(ref department, value); }

        private string underDepartment = "";
        public string UnderDepartment { get => underDepartment; set => Set(ref underDepartment, value); }

        private string contactName = "";
        public string ContactName { get => contactName; set => Set(ref contactName, value); }

        private string contactSurname = "";
        public string ContactSurname { get => contactSurname; set => Set(ref contactSurname, value); }

        private string profession = "";
        public string Profession { get => profession; set => Set(ref profession, value); }

        private string telNumber = "";
        public string TelNumber { get => telNumber; set => Set(ref telNumber, value); }

        private int contactState = 0; //0-User add row //1-Admin add row
        public int ContactState { get => contactState; set => Set(ref contactState, value); }
    }
}
