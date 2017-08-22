using System;
using System.ComponentModel.DataAnnotations;
using Telegraph.CustomValidationAttributes;

namespace Telegraph
{

    public class Telegram : PropertyChangedNotification, ICloneable
    {

        public int Id { get; set; }

        [Unique(ErrorMessage = "Телеграма з даним ID вже існує")]
        [Required(AllowEmptyStrings = false, ErrorMessage = "Поле не повинно бути пустим")]
        [Range(1, 1000000, ErrorMessage = "Значення повинно бути від 1 до 1000000")]
        public int SelfNum
        {
            get { return GetValue(() => SelfNum); }
            set { SetValue(() => SelfNum, value); }
        }

        [Unique(ErrorMessage = "Телеграма з даним вхідним номером вже існує")]
        [Required(AllowEmptyStrings = false, ErrorMessage = "Поле не повинно бути пустим")]
        [Range(1, 1000000, ErrorMessage = "Значення повинно бути від 1 до 1000000")]
        public int IncNum
        {
            get { return GetValue(() => IncNum); }
            set { SetValue(() => IncNum, value); }
        }

        [Required(AllowEmptyStrings = false, ErrorMessage = "Поле не повинно бути пустим")]
        public string From
        {
            get { return GetValue(() => From); }
            set { SetValue(() => From, value); }
        }

        [Required(AllowEmptyStrings = false, ErrorMessage = "Поле не повинно бути пустим")]
        public string To
        {
            get { return GetValue(() => To); }
            set { SetValue(() => To, value); }
        }

        [Required(AllowEmptyStrings = false, ErrorMessage = "Поле не повинно бути пустим")]
        public string Text
        {
            get { return GetValue(() => Text); }
            set { SetValue(() => Text, value); }
        }

        [Required(AllowEmptyStrings = false, ErrorMessage = "Поле не повинно бути пустим")]
        public string SubNum
        {
            get { return GetValue(() => SubNum); }
            set { SetValue(() => SubNum, value); }
        }

        [Required(AllowEmptyStrings = false, ErrorMessage = "Поле не повинно бути пустим")]
        public string Date
        {
            get { return GetValue(() => Date); }
            set { SetValue(() => Date, value); }
        }

        [Range(0, 1, ErrorMessage = "Значення повинно бути 0 або 1")]
        public int Urgency
        {
            get { return GetValue(() => Urgency); }
            set { SetValue(() => Urgency, value); }
        }

        [Required(AllowEmptyStrings = false, ErrorMessage = "Поле не повинно бути пустим")]
        public string Dispatcher
        {
            get { return GetValue(() => Dispatcher); }
            set { SetValue(() => Dispatcher, value); }
        }

        [Required(AllowEmptyStrings = false, ErrorMessage = "Поле не повинно бути пустим")]
        public string SenderPos
        {
            get { return GetValue(() => SenderPos); }
            set { SetValue(() => SenderPos, value); }
        }

        [Required(AllowEmptyStrings = false, ErrorMessage = "Поле не повинно бути пустим")]
        public string SenderName
        {
            get { return GetValue(() => SenderName); }
            set { SetValue(() => SenderName, value); }
        }

        [Required(AllowEmptyStrings = false, ErrorMessage = "Поле не повинно бути пустим")]
        public string SenderRank
        {
            get { return GetValue(() => SenderRank); }
            set { SetValue(() => SenderRank, value); }
        }

        [Required(AllowEmptyStrings = false, ErrorMessage = "Поле не повинно бути пустим")]
        public string Executor
        {
            get { return GetValue(() => Executor); }
            set { SetValue(() => Executor, value); }
        }

        [MaxLength(13, ErrorMessage = "Значення поля перевищує 13 символів")]
        public string Phone
        {
            get { return GetValue(() => Phone); }
            set { SetValue(() => Phone, value); }
        }


        [Required(AllowEmptyStrings = false, ErrorMessage = "Поле не повинно бути пустим")]
        [MaxLength(13, ErrorMessage = "Значення поля перевищує 20 символів")]
        public string HandedBy
        {
            get { return GetValue(() => HandedBy); }
            set { SetValue(() => HandedBy, value); }
        }

        public string Time
        {
            get { return GetValue(() => Time); }
            set { SetValue(() => Time, value); }
        }

        public object Clone()
        {
            return new Telegram
            {
                Id = this.Id,
                SelfNum = this.SelfNum,
                IncNum = this.IncNum,
                From = this.From,
                To = this.To,
                Text = this.Text,
                SubNum = this.SubNum,
                Date = this.Date,
                SenderPos = this.SenderPos,
                SenderRank = this.SenderRank,
                SenderName = this.SenderName,
                Executor = this.Executor,
                Phone = this.Phone,
                HandedBy = this.HandedBy,
                Urgency = this.Urgency,
                Dispatcher = this.Dispatcher,
                Time = this.Time
            };
        }
    }
}