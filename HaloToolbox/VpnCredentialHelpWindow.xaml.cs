using System.Diagnostics;
using System.Windows;

namespace HaloToolbox;

public partial class VpnCredentialHelpWindow : Window
{
    private const string NordCredentialsUrl =
        "https://my.nordaccount.com/dashboard/nordvpn/manual-configuration/service-credentials/";

    public VpnCredentialHelpWindow()
    {
        InitializeComponent();
    }

    public static void OpenNordCredentialsPage()
    {
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = NordCredentialsUrl,
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Could not open the Nord Account page.\n\n{ex.Message}\n\n{NordCredentialsUrl}",
                "Halo Toolbox",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
    }

    private void OpenCredentials_Click(object sender, RoutedEventArgs e) =>
        OpenNordCredentialsPage();

    private void Close_Click(object sender, RoutedEventArgs e) => Close();
}
