using System;
using System.IO;
using System.Text;
using System.Linq;
using System.Runtime;
using System.Windows;
using System.Reflection;
using System.Windows.Data;
using System.Globalization;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;
using System.ComponentModel;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Xml.Serialization;
using System.Windows.Navigation;
using System.Collections.Generic;
using System.Windows.Media.Imaging;
using System.Runtime.Serialization;
using System.Collections.ObjectModel;
using System.Runtime.CompilerServices;
using System.Runtime.Serialization.Formatters.Binary;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Win32;

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

        private int id;
        public int Id { get => id; set => Set(ref id, value); }

        private string departmentIcon;
        public string DepartmentIcon { get => departmentIcon; set => Set(ref departmentIcon, value); }

        private string department;
        public string Department { get => department; set => Set(ref department, value); }

        private string underDepartment;
        public string UnderDepartment { get => underDepartment; set => Set(ref underDepartment, value); }

        private string contactName;
        public string ContactName { get => contactName; set => Set(ref contactName, value); }

        private string contactSurname;
        public string ContactSurname { get => contactSurname; set => Set(ref contactSurname, value); }

        private string profession;
        public string Profession { get => profession; set => Set(ref profession, value); }

        private string telNumber;
        public string TelNumber { get => telNumber; set => Set(ref telNumber, value); }

        private int contactState; //0-User add row //1-Admin add row
        public int ContactState { get => contactState; set => Set(ref contactState, value); }
    }

    public class MyExcel
    {
        public static List<MyContact> bankAccounts;
        MyExcel()
        {
            bankAccounts = new List<MyContact> {
            new MyContact {
                          Id = 345678,
                          Department = "541.27"
                        },
            new MyContact {
                  Id = 1230221,
                  Department = "-127.44"
                }
            };

            MyExcel.DisplayInExcel(bankAccounts);
        }

        public static ObservableCollection<MyContact> ReadExcel()
        {
#if Debug
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = @"C:\";
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.Title = "Browse Excel Files";
            openFileDialog1.DefaultExt = "xls";
            openFileDialog1.Filter = "Excel files (*.xls)|*.xls|Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.CheckFileExists = true;
            openFileDialog1.CheckPathExists = true;
            openFileDialog1.ShowDialog();
            string filename = openFileDialog1.FileName;

            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filename);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!
            //for (int k = 1; k <= 3; k++)
            //{
            //    for (int m = 1; m <= 3; m++)
            //    {
            //        //new line
            //        if (m == 1)
            //            Console.Write("\r\n");
            //        //write the value to the console
            //        if (xlRange.Cells[k, m] != null && xlRange.Cells[k, m].Value2 != null)
            //            Console.Write(xlRange.Cells[k, m].Value2.ToString() + "\t");
            //        //add useful things here!   
            //    }
            //}
#endif
            ObservableCollection<MyContact> OCMyContacts = new ObservableCollection<MyContact>();
            

            int i = 1;
            //int j = 1;
            //while (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
            {
                MyContact myContact = new MyContact();

                myContact.Id = i++;
                myContact.Department = "İnformasiya Texnologiyaları";
                myContact.ContactName = "Rüstam";
                myContact.ContactSurname = "Zəkəryəyev";
                myContact.Profession = "Proqram təmitatı üzrə mütəxəssis";
                myContact.UnderDepartment = "Proqramlaşdırma";
                myContact.TelNumber = "+994772709923";
                myContact.DepartmentIcon = @"C:\Users\User\source\repos\PhonebookBM\logo.png";
                myContact.ContactState = 0;

                OCMyContacts.Add(myContact);

                myContact = new MyContact();

                myContact.Id = i++;
                myContact.Department = "İnformasiya Texnologiyaları";
                myContact.ContactName = "Zakir";
                myContact.ContactSurname = "Vəliyev";
                myContact.Profession = "İnformasiya Texnologiyaları üzrə Departament müdiri";
                myContact.UnderDepartment = "";
                myContact.TelNumber = "+994772709940";
                myContact.DepartmentIcon = @"C:\Users\User\source\repos\PhonebookBM\logo.png";
                myContact.ContactState = 1;

                OCMyContacts.Add(myContact);
            }

            return OCMyContacts;
        }
        

        public static void DisplayInExcel(IEnumerable<MyContact> myContact)
        {
            var excelApp = new Excel.Application();
            // Make the object visible.
            excelApp.Visible = true;

            // Create a new, empty workbook and add it to the collection returned 
            // by property Workbooks. The new workbook becomes the active workbook.
            // Add has an optional parameter for specifying a praticular template. 
            // Because no argument is sent in this example, Add creates a new workbook. 
            excelApp.Workbooks.Add();

            // This example uses a single workSheet. The explicit type casting is
            // removed in a later procedure.
            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;

            workSheet.Cells[1, "A"] = "ID Number";
            workSheet.Cells[1, "B"] = "Current Balance";

            var row = 1;
            foreach (var acct in bankAccounts)
            {
                row++;
                workSheet.Cells[row, "A"] = acct.Id;
                workSheet.Cells[row, "B"] = acct.Department;
            }

            workSheet.Columns[1].AutoFit();
            workSheet.Columns[2].AutoFit();

            ((Excel.Range)workSheet.Columns[1]).AutoFit();
            ((Excel.Range)workSheet.Columns[2]).AutoFit();

            // Put the spreadsheet contents on the clipboard. The Copy method has one
            // optional parameter for specifying a destination. Because no argument  
            // is sent, the destination is the Clipboard.
            workSheet.Range["A1:B3"].Copy();

            // Call to AutoFormat in Visual C# 2010.
            workSheet.Range["A1", "B3"].AutoFormat(Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic2);

            // Call to AutoFormat in Visual C# 2010.
            workSheet.Range["A4", "B6"].AutoFormat(Format:Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic2);

            // The AutoFormat method has seven optional value parameters. The
            // following call specifies a value for the first parameter, and uses 
            // the default values for the other six. 

            // Call to AutoFormat in Visual C# 2008. This code is not part of the
            // current solution.
            excelApp.get_Range("A7", "B9").AutoFormat(Excel.XlRangeAutoFormat.xlRangeAutoFormatTable3,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        }
    }

    public static class MyWord
    {
        public static void CreateIconInWordDoc()
        {
            var wordApp = new Word.Application();
            wordApp.Visible = true;

            // The Add method has four reference parameters, all of which are 
            // optional. Visual C# allows you to omit arguments for them if
            // the default values are what you want.
            wordApp.Documents.Add();

            // PasteSpecial has seven reference parameters, all of which are 
            // optional. This example uses named arguments to specify values 
            // for two of the parameters. Although these are reference 
            // parameters, you do not need to use the ref keyword, or to create 
            // variables to send in as arguments. You can send the values directly.
            wordApp.Selection.PasteSpecial(Link: true, DisplayAsIcon: true);
        }

        public static void CreateIconInWordDoc2008()
        {
            var wordApp = new Word.Application();
            wordApp.Visible = true;

            // The Add method has four parameters, all of which are optional. 
            // In Visual C# 2008 and earlier versions, an argument has to be sent 
            // for every parameter. Because the parameters are reference  
            // parameters of type object, you have to create an object variable
            // for the arguments that represents 'no value'. 

            object useDefaultValue = Type.Missing;

            wordApp.Documents.Add(ref useDefaultValue, ref useDefaultValue,
                ref useDefaultValue, ref useDefaultValue);

            // PasteSpecial has seven reference parameters, all of which are
            // optional. In this example, only two of the parameters require
            // specified values, but in Visual C# 2008 an argument must be sent
            // for each parameter. Because the parameters are reference parameters,
            // you have to contruct variables for the arguments.
            object link = true;
            object displayAsIcon = true;

            wordApp.Selection.PasteSpecial(ref useDefaultValue,
                                            ref link,
                                            ref useDefaultValue,
                                            ref displayAsIcon,
                                            ref useDefaultValue,
                                            ref useDefaultValue,
                                            ref useDefaultValue);
        }
    }

    public class MyJSON
    {

    }

}
