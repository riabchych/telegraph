/*
  In App.xaml:
  <Application.Resources>
      <vm:ViewModelLocator xmlns:vm="clr-namespace:Telegraph"
                           x:Key="Locator" />
  </Application.Resources>
  
  In the View:
  DataContext="{Binding Source={StaticResource Locator}, Path=ViewModelName}"

  You can also use Blend to do all this with the tool's support.
  See http://www.galasoft.ch/mvvm
*/

using GalaSoft.MvvmLight;

namespace Telegraph.ViewModels
{
    /// <summary>
    /// This class contains static references to all the view models in the
    /// application and provides an entry point for the bindings.
    /// </summary>
    public class ViewModelLocator
    {
        private ITelegramDataService _dataService;
        /// <summary>
        /// Initializes a new instance of the ViewModelLocator class.
        /// </summary>
        public ViewModelLocator()
        {            

            if (ViewModelBase.IsInDesignModeStatic)
            {
                _dataService = new TelegramDisignDataService();
            }
            else
            {
                _dataService = new TelegramDataService();
            }
            App = new ApplicationViewModel(_dataService);
        }

        public ApplicationViewModel App
        {
            get; private set;
        }
    }
}