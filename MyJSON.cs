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
    public class MyJSON
    {
        public MyJSON()
        {

        }

        public static void Save(ObservableCollection<MyContact> collection)
        {
            DataContractJsonSerializer jsonFormatter = new DataContractJsonSerializer(typeof(ObservableCollection<MyContact>));
            using (FileStream fs = new FileStream("Contacts.json", FileMode.Create))
            {
                jsonFormatter.WriteObject(fs, collection);
            }
        }

        public static ObservableCollection<MyContact> Load()
        {
            if (File.Exists("Contacts.json"))
            {
                DataContractJsonSerializer jsonFormatter = new DataContractJsonSerializer(typeof(ObservableCollection<MyContact>));
                using (FileStream fs = new FileStream("Contacts.json", FileMode.Open))
                {
                    if (fs.Length > 3)
                        return (ObservableCollection<MyContact>)jsonFormatter.ReadObject(fs);
                    else
                        MessageBox.Show("file is empty");
                }
            }
            return new ObservableCollection<MyContact>();
        }
    }
}
