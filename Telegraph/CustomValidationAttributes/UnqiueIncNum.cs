using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;
using System.Windows;

namespace Telegraph.CustomValidationAttributes
{
    public class UnqiueIncNum : ValidationAttribute
    {
        protected override ValidationResult IsValid(object value, ValidationContext validationContext)
        {
            var contains = ApplicationViewModel.SharedViewModel().Telegrams.Select(x => x.IncNum).Contains(int.Parse(value.ToString()));

            if (contains)
                return new ValidationResult(FormatErrorMessage(validationContext.DisplayName));
            else
                return ValidationResult.Success;
        }
    }
}
