using System.ComponentModel.DataAnnotations;
using System.Linq;
using Telegraph.ViewModels;

namespace Telegraph.CustomValidationAttributes
{
    public class Unique : ValidationAttribute
    {
        protected override ValidationResult IsValid(object value, ValidationContext validationContext)
        {

            int count = 0, val = int.Parse(value.ToString());
            ApplicationViewModel vm = ApplicationViewModel.SharedViewModel();
            Telegram instance = (Telegram)validationContext.ObjectInstance;

            if (vm.Telegrams == null)
            {
                return ValidationResult.Success;
            }

            switch (validationContext.MemberName)
            {
                case "SelfNum":
                    count = vm.Telegrams.Count() > 0 ? vm.Telegrams.Where(x => x.SelfNum.Equals(val) && !x.Id.Equals(instance.Id)).Count() : 0;
                    break;
                case "IncNum":
                    count = vm.Telegrams.Count() > 0 ? vm.Telegrams.Where(x => x.IncNum.Equals(val) && !x.Id.Equals(instance.Id)).Count() : 0;
                    break;
                default:
                    count = 0;
                    break;
            }

            return (count > 0) ? new ValidationResult(FormatErrorMessage(validationContext.DisplayName)) : ValidationResult.Success;
        }
    }
}
