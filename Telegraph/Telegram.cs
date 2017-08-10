using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace Telegraph
{

    public class Telegram : INotifyPropertyChanged
    {
        private int number;
        private string to;
        private string text;
        private string subnum;
        private string date;
        private int urgency;
        private string dispatcher;

        public int id { get; set; }

        public int Number
        {
            get { return number; }
            set
            {
                number = value;
                OnPropertyChanged("Number");
            }
        }

        public string To
        {
            get { return to; }
            set
            {
                to = value;
                OnPropertyChanged("To");
            }
        }

        public string Text
        {
            get { return text; }
            set
            {
                text = value;
                OnPropertyChanged("Text");
            }
        }

        public string Subnum
        {
            get { return subnum; }
            set
            {
                subnum = value;
                OnPropertyChanged("Subnum");
            }
        }

        public string Date
        {
            get { return date; }
            set
            {
                date = value;
                OnPropertyChanged("Date");
            }
        }

        public int Urgency
        {
            get { return urgency; }
            set
            {
                urgency = value;
                OnPropertyChanged("Urgency");
            }
        }

        public string Dispatcher
        {
            get { return dispatcher; }
            set
            {
                dispatcher = value;
                OnPropertyChanged("Dispatcher");
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged([CallerMemberName]string prop = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        }
    }
}