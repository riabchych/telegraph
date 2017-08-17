using GalaSoft.MvvmLight.Command;
using System.Windows.Controls;

namespace Telegraph.Converters
{
    public class TextChangedEventArgsConverter: IEventArgsConverter
    {
        public object Convert(object value, object parameter)
        {
            var textChangedEventArgs = (TextChangedEventArgs)value;
            return ((TextBox)textChangedEventArgs.Source).Text;
        }
    }
}
