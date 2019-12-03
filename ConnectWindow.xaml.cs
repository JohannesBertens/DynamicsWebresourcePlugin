using System;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using Microsoft.VisualStudio.PlatformUI;

namespace WebResourcePlugin
{
    /// <summary>
    /// Interaction logic for ConnectWindow.xaml
    /// </summary>
    public partial class ConnectWindow : DialogWindow
    {
        public ConnectWindow()
        {
            InitializeComponent();
            FillSettings();
        }

        private void FillSettings()
        {
            var settings = ConnectToDynamics.Instance.ConnectionSettings;
            DynamicsUrl.Text = settings.Url;
            DynamicsUser.Text = settings.User;
            DynamicsDomain.Text = settings.Domain;

            if (ConnectToDynamics.Instance.IsConnected)
            {
                ConnectedToDynamics();
            }
        }
        
        public void OnConnect(object sender, RoutedEventArgs e)
        {
            var url = DynamicsUrl.Text;
            var domain = DynamicsDomain.Text;
            var user = DynamicsUser.Text;
            var password = DynamicsPassword.Password;
            var saveSettings = SaveOnConnect.IsChecked ?? false;

            if (ConnectToDynamics.Instance.MakeConnection(url, domain, user, password, saveSettings))
            {
                this.StatusLabel.Content = $"CONNECTED! to {this.DynamicsUrl.Text}";
                ConnectedToDynamics();
            }
            else
            {
                this.StatusLabel.Content = $"Not connected to {this.DynamicsUrl.Text}";
            }
        }

        private void ConnectedToDynamics()
        {
            FillCombobox(ConnectToDynamics.Instance.GetSolutions());
            DynamicsSolutions.IsEnabled = true;
            ButtonOk.IsEnabled = true;
        }

        private void OnConnectKey(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Return || e.Key == Key.Enter)
            {
                OnConnect(sender, e);
                e.Handled = true;
            }
        }

        public void OnCancel(object sender, RoutedEventArgs e)
        {
            this.Hide();
        }

        public void OnOk(object sender, RoutedEventArgs e)
        {
            var selectedSolution = (string) DynamicsSolutions.SelectedValue;

            var settings = ConnectToDynamics.Instance.ConnectionSettings;
            if (!string.IsNullOrEmpty((string)DynamicsSolutions.SelectedValue) 
                && !DynamicsSolutions.SelectedValue.Equals(settings.Solution) 
                && SaveOnConnect.IsChecked.HasValue && SaveOnConnect.IsChecked.Value)
            {
                settings.Solution = selectedSolution;
                ConnectToDynamics.Instance.ConnectionSettings = settings;
            }

            ConnectToDynamics.Instance.SelectedSolution = selectedSolution;

            this.Hide();
        }

        public void FillCombobox(string[] solutions)
        {
            DynamicsSolutions.ItemsSource = solutions;

            var settings = ConnectToDynamics.Instance.ConnectionSettings;
            for (var i = 0;i < solutions.Length; i++)
            {
                if (solutions[i].Equals(settings.Solution))
                {
                    DynamicsSolutions.SelectedIndex = i;
                }
            }
        }
    }
}
