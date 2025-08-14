using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Codumentor.ViewModels
{
    public class TitleBarViewModel : BindableBase
    {
        public DelegateCommand MinimizeCommand { get; }
        public DelegateCommand CloseCommand { get; }

        public TitleBarViewModel()
        {
            MinimizeCommand = new DelegateCommand(OnMinimize);
            CloseCommand = new DelegateCommand(OnClose);
        }

        private void OnMinimize()
        {
            Application.Current.MainWindow.WindowState = WindowState.Minimized;
        }

        private void OnClose()
        {
            Application.Current.MainWindow.Close();
        }
    }
}
