using System.ComponentModel.DataAnnotations;
using System.Linq;

namespace Telegraph.CustomValidationAttributes
{
    public class Unique : ValidationAttribute
    {
        private readonly string field;

        public Unique(string val)
        {
            field = val;
        }

        protected override ValidationResult IsValid(object value, ValidationContext validationContext)
        {
            int val = int.Parse(value.ToString());
            bool result = false;

            switch (field)
            {
                case "SelfNum":
                    if (val != ApplicationViewModel.SharedViewModel().EditSelfNum)
                    {
                        if (ApplicationViewModel.SharedViewModel().Telegrams.Select(x => x.SelfNum).Contains(val))
                        {
                            result = false;
                        }
                        else
                        {
                            result = true;
                        }
                    }
                    else
                    {
                        result = true;
                    }
                    break;
                case "IncNum":
                    if (val != ApplicationViewModel.SharedViewModel().EditIncNum)
                    {
                        if (ApplicationViewModel.SharedViewModel().Telegrams.Select(x => x.IncNum).Contains(val))
                        {
                            result = false;
                        }
                        else
                        {
                            result = true;
                        }
                    }
                    else
                    {
                        result = true;
                    }
                    break;
            }

            if (!result)
                return new ValidationResult(FormatErrorMessage(validationContext.DisplayName));
            else
                return ValidationResult.Success;
        }
    }
}
