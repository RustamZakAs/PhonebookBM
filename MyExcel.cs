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
    public class MyExcel
    {
        string filePath = "";
        public List<MyContact> contacs;
        ObservableCollection<MyContact> OCMyContacts = new ObservableCollection<MyContact>();

        public MyExcel()
        {

        }

        public MyExcel(string filePath)
        {
            this.filePath = filePath;
        }

        public ObservableCollection<MyContact> ReadExcel()
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(this.filePath);
            Excel.Worksheet xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            int col = 0;
            int goodRow = 0;
            for (int i = 2; i <= rowCount; i++)
            {
                MyContact row = new MyContact();
                //************************************************************************
                col = 2;
                if (xlRange.Cells[i, col] != null && xlRange.Cells[i, col].Value2 != null)
                    row.DepartmentIcon = xlRange.Cells[i, col].Value2.ToString();
                //************************************************************************
                col = 3;
                if (xlRange.Cells[i, col] != null && xlRange.Cells[i, col].Value2 != null)
                    row.Department = xlRange.Cells[i, col].Value2.ToString();
                //************************************************************************
                col = 4;
                if (xlRange.Cells[i, col] != null && xlRange.Cells[i, col].Value2 != null)
                    row.UnderDepartment = xlRange.Cells[i, col].Value2.ToString();
                //************************************************************************
                col = 5;
                if (xlRange.Cells[i, col] != null && xlRange.Cells[i, col].Value2 != null)
                    row.ContactName = xlRange.Cells[i, col].Value2.ToString();
                //************************************************************************
                col = 6;
                if (xlRange.Cells[i, col] != null && xlRange.Cells[i, col].Value2 != null)
                    row.ContactSurname = xlRange.Cells[i, col].Value2.ToString();
                //************************************************************************
                col = 7;
                if (xlRange.Cells[i, col] != null && xlRange.Cells[i, col].Value2 != null)
                    row.Profession = xlRange.Cells[i, col].Value2.ToString();
                //************************************************************************
                col = 8;
                if (xlRange.Cells[i, col] != null && xlRange.Cells[i, col].Value2 != null)
                    row.TelNumber = xlRange.Cells[i, col].Value2.ToString();
                //************************************************************************
                col = 9;
                if (xlRange.Cells[i, col] != null && xlRange.Cells[i, col].Value2 != null)
                    row.ContactState = xlRange.Cells[i, col].Value2.ToString();
                //************************************************************************
                col = 10;
                if (xlRange.Cells[i, col] != null && xlRange.Cells[i, col].Value2 != null)
                    row.Сonfirmed = xlRange.Cells[i, col].Value2.ToString();
                //************************************************************************
                if (row.Department.Length > 0 && row.UnderDepartment.Length > 0 && row.ContactName.Length > 0)
                {
                    row.Id = goodRow++;
                    this.OCMyContacts.Add(row);
                }
            }

            #region CloseClearExcel
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            #endregion

            int k = 1;
            {
                MyContact row = new MyContact();
                row.Id = k++;
                row.Department = "İnformasiya Texnologiyaları";
                row.ContactName = "Rüstam";
                row.ContactSurname = "Zəkəryəyev";
                row.Profession = "Proqram təmitatı üzrə mütəxəssis";
                row.UnderDepartment = "Proqramlaşdırma";
                row.TelNumber = "+994772709923";
                row.DepartmentIcon = @".\logo.png";
                row.ContactState = 0;
                this.OCMyContacts.Add(row);

                row = new MyContact();
                row.Id = k++;
                row.Department = "İnformasiya Texnologiyaları";
                row.ContactName = "Zakir";
                row.ContactSurname = "Vəliyev";
                row.Profession = "İnformasiya Texnologiyaları üzrə Departament müdiri";
                row.UnderDepartment = "";
                row.TelNumber = "+994772709940";
                row.DepartmentIcon = @".\logo.png";
                row.ContactState = 1;
                this.OCMyContacts.Add(row);
            }
            return OCMyContacts;
        }
    }
}
