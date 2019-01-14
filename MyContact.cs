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
    [DataContract]
    public class MyContact : INotifyPropertyChanged, ICloneable
    {
        public event PropertyChangedEventHandler PropertyChanged;
        public void Set<T>(ref T field, T value, [System.Runtime.CompilerServices.CallerMemberName]string prop = "")
        {
            field = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        }

        public object Clone()
        {
            return new MyContact {
                Id = this.Id,
                DepartmentIcon = this.DepartmentIcon,
                Department = this.Department,
                UnderDepartment = this.UnderDepartment,
                ContactName = this.ContactName,
                ContactSurname = this.ContactSurname,
                Profession = this.Profession,
                TelNumber = this.TelNumber,
                ContactState = this.ContactState,
                Confirmed = this.Confirmed,
                Deleted = this.Deleted
            };
            //return this.MemberwiseClone();
            //throw new NotImplementedException();
        }

        [DataMember] //1
        private int id = 0; 
        public int Id { get => id; set => Set(ref id, value); }

        [DataMember] //2
        private string departmentIcon = "";
        public string DepartmentIcon { get => departmentIcon; set => Set(ref departmentIcon, value); }

        [DataMember] //3
        private string department = "";
        public string Department { get => department; set => Set(ref department, value); }

        [DataMember] //4
        private string underDepartment = "";
        public string UnderDepartment { get => underDepartment; set => Set(ref underDepartment, value); }

        [DataMember] //5
        private string contactName = "";
        public string ContactName { get => contactName; set => Set(ref contactName, value); }

        [DataMember] //6
        private string contactSurname = "";
        public string ContactSurname { get => contactSurname; set => Set(ref contactSurname, value); }

        [DataMember] //7
        private string profession = "";
        public string Profession { get => profession; set => Set(ref profession, value); }

        [DataMember] //8
        private string telNumber = "";
        public string TelNumber { get => telNumber; set => Set(ref telNumber, value); }

        [DataMember] //9
        private int contactState = 0; //0-User add row //1-Admin add row
        public int ContactState { get => contactState; set => Set(ref contactState, value); }

        [DataMember] //10
        private bool confirmed = false; //подтверждено
        public bool Confirmed { get => confirmed; set => Set(ref confirmed, value); }

        [DataMember] //11
        private bool deleted = false; //удвлено пользователем
        public bool Deleted { get => deleted; set => Set(ref deleted, value); }
    }
}
