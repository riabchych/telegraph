using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace Telegraph
{
    public class Telegram : INotifyPropertyChanged
    {
        private int number;
        private string from;
        private string to;
        private string text;
        private string subnum;
        private string date;
        private int executor;
        private int urgency;

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
        public string From
        {
            get { return from; }
            set
            {
                from = value;
                OnPropertyChanged("From");
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

        public int Executor
        {
            get { return executor; }
            set
            {
                executor = value;
                OnPropertyChanged("Executor");
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

        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged([CallerMemberName]string prop = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        }
    }
}