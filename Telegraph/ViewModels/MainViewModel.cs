using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Data;

namespace Telegraph.ViewModels
{
    public abstract class MainViewModel : PropertyChangedNotification
    {
        protected struct TlgRegex
        {
            private const string footerRegex = @"^([ÂÈÊ\s]+[\.])*[\s]?([À-ß¨¯²ª¥\1]+){1}[\s]*([\s\d\.\-\/\:]+)?[\s]*$[\s\-]+^([ÏÐÄ\s]+[\.])*[\s]?([À-ß¨¯²ª¥\1]+){1}[\s]*([\s\d\.\-\/\:]+)?[\s]*$[\s\-]+^([ÏÐÍ\s]+[\.])*[\s]?([À-ß¨¯²ª¥\1]+)*[\s]*([\s\d\.\/\:]+)*$";
            private const string mainRegex = @".ÍÐ[\s\.]+([0-9]*)?[À-ß¨¯²ª¥\s\d\W]+^([À-ß¨¯²ª¥\s\d]*).ÍÐ[\s\.]+([0-9]*)?[\s]?$[\s]+^([À-ß¨¯²ª¥\d\/\s\(\)]+)[\s\n\r]+^([À-ß¨¯²ª¥\w\W\s]+)^[\s]*ÍÐ[\.]?[\s]*([\d\/\s\-]+[\s]*[\d]+)[\s]+([À-ß¨¯²ª¥\s\.\-0-9]+)[\r\n]^([\d]+[\s\.\\\/]+[\d]+[\s\.\\\/]+[\d]+)[\s]+([À-ß¨¯²ª¥\s\-\1]+)$[\s]+([À-ß¨¯²ª¥\-\1]+)?[\s]+([À-ß¨¯²ª¥\s\1\.]+)$[\r\n]?[\s\-]+";
            private const string urgencyRegex = @"[\s\d]{10,}....(ÒÅÐÌ²ÍÎÂ.)";

            public static string FooterRegex => footerRegex;
            public static string MainRegex => mainRegex;
            public static string UrgencyRegex => urgencyRegex;
        }

        private ApplicationContext db;
        private Task dbTask;

        protected string GetTempFile(string fileExtension)
        {
            string res = string.Empty;

            while (true)
            {
                res = System.IO.Path.ChangeExtension(System.IO.Path.GetTempFileName(), fileExtension);
                if (!System.IO.File.Exists(res))
                {
                    try
                    {
                        System.IO.File.Create(res).Close();
                        break;
                    }
                    catch (Exception)
                    {

                    }
                }
            }
            return res;
        }

        protected bool isMSWordInstalled()
        {
            using (var regWord = Registry.ClassesRoot.OpenSubKey("Word.Application"))
            {
                return (regWord == null) ? false : true;
            }
        }

        protected bool CheckData(Telegram data) => ((data.IncNum < 1) ||
                string.IsNullOrWhiteSpace(data.From) ||
                string.IsNullOrWhiteSpace(data.To) ||
                string.IsNullOrWhiteSpace(data.Text) ||
                string.IsNullOrWhiteSpace(data.SubNum) ||
                string.IsNullOrWhiteSpace(data.Date) ||
                string.IsNullOrWhiteSpace(data.SenderPos) ||
                string.IsNullOrWhiteSpace(data.SenderRank) ||
                string.IsNullOrWhiteSpace(data.SenderName) ||
                string.IsNullOrWhiteSpace(data.Executor) ||
                string.IsNullOrWhiteSpace(data.HandedBy))
            ? false : true;

        public IEnumerable<Telegram> Telegrams
        {
            get { return GetValue(() => Telegrams); }
            set { SetValue(() => Telegrams, value); }
        }

        public CollectionViewSource TelegramsViewSource
        {
            get { return GetValue(() => TelegramsViewSource); }
            set { SetValue(() => TelegramsViewSource, value); }
        }

        public ICollectionView SourceCollection
        {
            get { return TelegramsViewSource.View; }
        }

        public bool IsBusy
        {
            get { return GetValue(() => IsBusy); }
            set { SetValue(() => IsBusy, value); }
        }

        protected Task DbTask { get => dbTask; set => dbTask = value; }
        protected ApplicationContext Db { get => db; set => db = value; }
    }
}