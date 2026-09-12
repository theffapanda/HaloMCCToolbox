using System.Windows;
using System.Windows.Input;
using System.Windows.Media;

namespace HaloToolbox;

public partial class OverlayRepositionHelpWindow : Window
{
    public OverlayRepositionHelpWindow()
    {
        InitializeComponent();
    }

    private void DismissButton_Click(object sender, RoutedEventArgs e) => Close();

    private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ChangedButton != MouseButton.Left)
            return;

        for (DependencyObject? element = e.OriginalSource as DependencyObject;
             element is not null && !ReferenceEquals(element, sender);
             element = VisualTreeHelper.GetParent(element))
        {
            if (element is System.Windows.Controls.Button)
                return;
        }

        DragMove();
    }

    private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.Escape)
            return;

        Close();
        e.Handled = true;
    }
}
