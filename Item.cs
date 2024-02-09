using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace WpfApp1
{
    public class Item : INotifyPropertyChanged
    {
        string name;
        double price;
        int amount;
        public string Name
        {
            get
            {
                return name;
            }
            set
            {

                name = value;
                OnPropertyChanged("Name");
            }
        }
        public int Amount { 
            get
            {
                return amount;
            }
            set
            {
                
                amount = value;
                OnPropertyChanged("Amount");
                OnPropertyChanged("Sum");
            }
        }
        public double Price
        {
            get
            {
                return price;
            }
            set
            {
                price = value;
                OnPropertyChanged("Price");
                OnPropertyChanged("Sum");
            }
        }
        public double Sum
        {
            get
            {
                return amount * price;
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        void OnPropertyChanged(string prop)
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(prop));
        }
    }
}
