using Microsoft.Win32;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Data;

namespace Telegraph.ViewModels
{
    public struct TlgRegex
    {
        private const string footerRegex = @"^([���\s]+[\.])*[\s]?([�-ߨ����\1]+){1}[\s]*([\s\d\.\-\/\:]+)?[\s]*$[\s\-]+^([���\s]+[\.])*[\s]?([�-ߨ����\1]+){1}[\s]*([\s\d\.\-\/\:]+)?[\s]*$[\s\-]+^([���\s]+[\.])*[\s]?([�-ߨ����\1]+)*[\s]*([\s\d\.\/\:]+)*$";
        private const string mainRegex = @".��[\s\.]+([0-9]*)?[�-ߨ����\s\d\W]+^([�-ߨ����\s\d]*).��[\s\.]+([0-9]*)?[\s]?$[\s]+^([�-ߨ����\d\/\s\(\)]+)[\s\n\r]+^([�-ߨ����\w\W\s]+)^[\s]*��[\.]?[\s]*([\d\/\s\-]+[\s]*[\d]+)[\s]+([�-ߨ����\s\.\-0-9]+)[\r\n]^([\d]+[\s\.\\\/]+[\d]+[\s\.\\\/]+[\d]+)[\s]+([�-ߨ����\s\-\1]+)$[\s]+([�-ߨ����\-\1]+)?[\s]+([�-ߨ����\s\1\.]+)$[\r\n]?[\s\-]+";
        private const string urgencyRegex = @"[\s\d]{10,}....(���̲���.)";

        public static string FooterRegex => footerRegex;
        public static string MainRegex => mainRegex;
        public static string UrgencyRegex => urgencyRegex;
    }

    public abstract class MainViewModel : PropertyChangedNotification
    {
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
                return regWord != null;
            }
        }

        protected bool CheckData(Telegram data) => !((data.IncNum < 1) ||
                string.IsNullOrWhiteSpace(data.From) ||
                string.IsNullOrWhiteSpace(data.To) ||
                string.IsNullOrWhiteSpace(data.Text) ||
                string.IsNullOrWhiteSpace(data.SubNum) ||
                string.IsNullOrWhiteSpace(data.Date) ||
                string.IsNullOrWhiteSpace(data.SenderPos) ||
                string.IsNullOrWhiteSpace(data.SenderRank) ||
                string.IsNullOrWhiteSpace(data.SenderName) ||
                string.IsNullOrWhiteSpace(data.Executor) ||
                string.IsNullOrWhiteSpace(data.HandedBy));

        public ICollection<Telegram> Telegrams
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

        protected IList<t> ConvertListOfSelectedItem<t>(object param)
        {
            return (param as IList).Cast<t>().ToList();
        }

        protected Task DbTask { get => dbTask; set => dbTask = value; }
        protected ApplicationContext Db { get => db; set => db = value; }
    }
}