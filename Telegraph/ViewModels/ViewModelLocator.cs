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
using MvvmDialogs;
using Telegraph.LogModule;
using Telegraph.LogModule.Loggers;
using Telegraph.LogModule.Loggers.TraceLogger;

namespace Telegraph.ViewModels
{
    /// <summary>
    /// This class contains static references to all the view models in the
    /// application and provides an entry point for the bindings.
    /// </summary>
    public class ViewModelLocator
    {
        /// <summary>
        /// Initializes a new instance of the ViewModelLocator class.
        /// </summary>
        public ViewModelLocator()
        {
            SimpleIoc.Default.Reset();
            ServiceLocator.SetLocatorProvider(() => SimpleIoc.Default);

            if (ViewModelBase.IsInDesignModeStatic)
            {
                SimpleIoc.Default.Register<ITelegramDataService>(() => new TelegramDisignDataService());
            }
            else
            {
                SimpleIoc.Default.Register<ITelegramDataService>(() => new TelegramDataService());
            }

            SimpleIoc.Default.Register<ILogger>(() => LoggerFactory.Create<TraceLogger>());
            SimpleIoc.Default.Register<IDialogService>(() => new DialogService());
            SimpleIoc.Default.Register(() => new ApplicationViewModel());
            SimpleIoc.Default.Register(() => new ImportViewModel());
        }

        public ApplicationViewModel App
        {
            get
            {
                return ServiceLocator.Current.GetInstance<ApplicationViewModel>();
            }
        }

        public ImportViewModel Import
        {
            get
            {
                return ServiceLocator.Current.GetInstance<ImportViewModel>();
            }
        }
    }
}