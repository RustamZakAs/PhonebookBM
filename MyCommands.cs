﻿using GalaSoft.MvvmLight.CommandWpf;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace PhonebookBM
{
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        private RelayCommand itemDeleteCommand;
        public RelayCommand ItemDeleteCommand
        {
            get => itemDeleteCommand ?? (itemDeleteCommand = new RelayCommand(
                 () =>
                 {
                     //MessageBox.Show(SelectedContact.ContactName);
                     MessageBoxResult flag = MessageBox.Show("Silinməsinə əminsinizmi?",
                                                             MyApp.Name,
                                                             MessageBoxButton.YesNo,
                                                             MessageBoxImage.Question);
                     if (flag == MessageBoxResult.Yes)
                     {
                         if (UserStatus == 0)
                         {
                             int x = lbItems.SelectedIndex;
                             if (lbItems.SelectedItem != null && x > -1)
                             {
                                 OCMyContactsFiltered.Remove(SelectedContact);
                                 OCMyContactsAll.Remove(SelectedContact);
                             }
                         }
                         else
                         {
                             int x = lbItems.SelectedIndex;
                             if (lbItems.SelectedItem != null && x > -1)
                             {
                                 OCMyContactsFiltered[x].Confirmed = true;
                                 OCMyContactsFiltered[x].ContactState = UserStatus;

                                 OCMyContactsAll[x].Confirmed = true;
                                 OCMyContactsAll[x].ContactState = UserStatus;
                             }
                         }
                     }
                 }
                 ));
        }

        private RelayCommand itemChangeCommand;
        public RelayCommand ItemChangeCommand
        {
            get => itemChangeCommand ?? (itemChangeCommand = new RelayCommand(
                 () =>
                 {
                     if (!IsChange)
                     {
                         if (SelectedContact == null)
                         {
                             MessageBox.Show("Sətir seçilməyib", MyApp.Name, MessageBoxButton.OK, MessageBoxImage.Information);
                             return;
                         }

                         ObservableCollection<MyContact> temp = new ObservableCollection<MyContact>();
                         for (int i = 0; i < OCMyContactsAll.Count; i++)
                         {
                             temp.Add(new MyContact(OCMyContactsAll[i]));
                         }
                         OCMyContactsAll = temp;

                         MyContact mc = SelectedContact;
                         mc.Confirmed = true;
                         mc.Deleted = false;
                         OCMyContactsFiltered.Clear();
                         OCMyContactsFiltered.Add(mc);
                         IsChange = true;
                     }
                     else
                     {
                         for (int i = 0; i < OCMyContactsFiltered.Count; i++)
                         {
                             var tempItem = OCMyContactsAll.Where(x => x.Id == OCMyContactsFiltered[i].Id).FirstOrDefault();
                             tempItem.Confirmed = true;
                             tempItem.Deleted = false;

                             OCMyContactsAll.Add(new MyContact(OCMyContactsFiltered[i]));
                             OCMyContactsAll.Remove(tempItem);
                         }

                         ObservableCollection<MyContact> temp = new ObservableCollection<MyContact>();
                         for (int i = 0; i < OCMyContactsAll.Count; i++)
                         {
                             temp.Add(new MyContact(OCMyContactsAll[i]));
                         }
                         OCMyContactsFiltered = temp;

                         IsChange = false;
                     }
                 }
                 ));
        }

        private RelayCommand itemAddCommand;
        public RelayCommand ItemAddCommand
        {
            get => itemAddCommand ?? (itemAddCommand = new RelayCommand(
                () =>
                {
                    if(!IsChange)
                    {
                        ObservableCollection<MyContact> temp = new ObservableCollection<MyContact>();
                        for (int i = 0; i < OCMyContactsAll.Count; i++)
                        {
                            temp.Add(new MyContact(OCMyContactsAll[i]));
                        }
                        //OCMyContactsFiltered = temp;

                        OCMyContactsFiltered.Clear();
                        int maxId = temp.Max(x => x.Id) + 1;
                        OCMyContactsFiltered.Add(new MyContact(maxId));

                        temp.Add(OCMyContactsFiltered[0]);

                        OCMyContactsAll = temp;

                        IsChange = true;
                    }
                    else
                    {
                        OCMyContactsFiltered = OCMyContactsAll;

                        IsChange = false;
                    }
                }
                ));
        }
    }

    static class Extensions
    {
        public static IList<T> Clone<T>(this IList<T> listToClone) where T : ICloneable
        {
            return listToClone.Select(item => (T)item.Clone()).ToList();
        }
    }

    class MyCommands
    {
    }
}
