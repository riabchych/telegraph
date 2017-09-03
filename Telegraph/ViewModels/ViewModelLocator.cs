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
using GalaSoft.MvvmLight.Ioc;
using Microsoft.Practices.ServiceLocation;

namespace Telegraph.ViewModel
{
    /// <summary>
    /// This class contains static references to all the view models in the
    /// application and provides an entry point for the bindings.
    /// </summary>
    public class ViewModelLocator
    {
        private ITelegramDataService _dataService;
        private ApplicationViewModel _app;
        /// <summary>
        /// Initializes a new instance of the ViewModelLocator class.
        /// </summary>
        public ViewModelLocator()
        {
            ServiceLocator.SetLocatorProvider(() => SimpleIoc.Default);
            

            ServiceLocator.SetLocatorProvider(() => SimpleIoc.Default);
            if (ViewModelBase.IsInDesignModeStatic)
            {
                _dataService = new TelegramDataService();
            }
            else
            {
                _dataService = new TelegramDataService();
            }
            _app = new ApplicationViewModel(_dataService);

            SimpleIoc.Default.Register<MainViewModel>();
            SimpleIoc.Default.Register<ApplicationViewModel>();
            SimpleIoc.Default.Register<ImportViewModel>();
        }

        public ApplicationViewModel App
        {
            get; private set;
        }

        public static void Cleanup()
        {
            // TODO Clear the ViewModels
        }
    }
}