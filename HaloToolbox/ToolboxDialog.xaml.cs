using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace HaloToolbox;

public partial class ToolboxDialog : Window
{
    private readonly MessageBoxButton _buttons;
    private readonly MessageBoxResult _closeResult;
    private Button? _defaultButton;
    private MessageBoxResult _result;

    private ToolboxDialog(
        string message,
        string caption,
        MessageBoxButton buttons,
        MessageBoxImage image,
        MessageBoxResult defaultResult)
    {
        InitializeComponent();
        _buttons = buttons;
        _closeResult = SafeCloseResult(buttons, defaultResult);
        _result = _closeResult;

        Title = string.IsNullOrWhiteSpace(caption) ? "Halo MCC Toolbox" : caption;
        DialogTitleText.Text = NormalizeCaption(caption).ToUpperInvariant();
        DialogMessageText.Text = message;
        ConfigureIcon(image);
        BuildButtons(defaultResult);
    }

    public static MessageBoxResult Show(string messageBoxText) =>
        Show(messageBoxText, "Halo MCC Toolbox", MessageBoxButton.OK, MessageBoxImage.None);

    public static MessageBoxResult Show(string messageBoxText, string caption) =>
        Show(messageBoxText, caption, MessageBoxButton.OK, MessageBoxImage.None);

    public static MessageBoxResult Show(
        string messageBoxText,
        string caption,
        MessageBoxButton button) =>
        Show(messageBoxText, caption, button, MessageBoxImage.None);

    public static MessageBoxResult Show(
        string messageBoxText,
        string caption,
        MessageBoxButton button,
        MessageBoxImage icon) =>
        Show(messageBoxText, caption, button, icon, MessageBoxResult.None);

    public static MessageBoxResult Show(
        string messageBoxText,
        string caption,
        MessageBoxButton button,
        MessageBoxImage icon,
        MessageBoxResult defaultResult) =>
        ShowCore(null, messageBoxText, caption, button, icon, defaultResult);

    public static MessageBoxResult Show(
        Window owner,
        string messageBoxText,
        string caption,
        MessageBoxButton button,
        MessageBoxImage icon) =>
        ShowCore(owner, messageBoxText, caption, button, icon, MessageBoxResult.None);

    public static MessageBoxResult Show(
        Window owner,
        string messageBoxText,
        string caption,
        MessageBoxButton button,
        MessageBoxImage icon,
        MessageBoxResult defaultResult) =>
        ShowCore(owner, messageBoxText, caption, button, icon, defaultResult);

    private static MessageBoxResult ShowCore(
        Window? owner,
        string message,
        string caption,
        MessageBoxButton buttons,
        MessageBoxImage image,
        MessageBoxResult defaultResult)
    {
        var dispatcher = Application.Current?.Dispatcher;
        if (dispatcher is not null && !dispatcher.CheckAccess())
        {
            return dispatcher.Invoke(() =>
                ShowCore(owner, message, caption, buttons, image, defaultResult));
        }

        owner ??= Application.Current?.Windows
            .OfType<Window>()
            .FirstOrDefault(window => window.IsActive && window.IsVisible);
        owner ??= Application.Current?.MainWindow is { IsVisible: true } mainWindow
            ? mainWindow
            : null;

        var dialog = new ToolboxDialog(
            message ?? "",
            caption ?? "Halo MCC Toolbox",
            buttons,
            image,
            defaultResult);
        if (owner is not null && !ReferenceEquals(owner, dialog))
            dialog.Owner = owner;

        dialog.ShowDialog();
        return dialog._result;
    }

    private void ConfigureIcon(MessageBoxImage image)
    {
        string glyph;
        string brushKey;
        switch (image)
        {
            case MessageBoxImage.Error:
                glyph = "×";
                brushKey = "RedBrush";
                break;
            case MessageBoxImage.Warning:
                glyph = "!";
                brushKey = "OrangeBrush";
                break;
            case MessageBoxImage.Question:
                glyph = "?";
                brushKey = "AccentBrush";
                break;
            case MessageBoxImage.None:
                glyph = "•";
                brushKey = "MutedBrush";
                break;
            default:
                glyph = "i";
                brushKey = "AccentBrush";
                break;
        }

        DialogIconText.Text = glyph;
        if (TryFindResource(brushKey) is Brush brush)
        {
            DialogIconText.Foreground = brush;
            DialogIconBorder.BorderBrush = brush;
        }
    }

    private void BuildButtons(MessageBoxResult requestedDefault)
    {
        var definitions = _buttons switch
        {
            MessageBoxButton.OK => new[] { (MessageBoxResult.OK, "OK") },
            MessageBoxButton.OKCancel => new[]
            {
                (MessageBoxResult.Cancel, "CANCEL"),
                (MessageBoxResult.OK, "OK")
            },
            MessageBoxButton.YesNoCancel => new[]
            {
                (MessageBoxResult.Cancel, "CANCEL"),
                (MessageBoxResult.No, "NO"),
                (MessageBoxResult.Yes, "YES")
            },
            _ => new[]
            {
                (MessageBoxResult.No, "NO"),
                (MessageBoxResult.Yes, "YES")
            }
        };

        MessageBoxResult effectiveDefault = IsResultAvailable(requestedDefault)
            ? requestedDefault
            : definitions[^1].Item1;

        foreach (var (result, label) in definitions)
        {
            bool isPrimary = result == effectiveDefault;
            var button = new Button
            {
                Content = label,
                Margin = new Thickness(10, 0, 0, 0),
                Style = (Style)FindResource(
                    isPrimary ? "DialogPrimaryButton" : "DialogSecondaryButton"),
                IsDefault = isPrimary,
                IsCancel = result == _closeResult
            };
            button.Click += (_, _) => Complete(result);
            DialogButtons.Children.Add(button);
            if (isPrimary)
                _defaultButton = button;
        }
    }

    private bool IsResultAvailable(MessageBoxResult result) => _buttons switch
    {
        MessageBoxButton.OK => result == MessageBoxResult.OK,
        MessageBoxButton.OKCancel =>
            result is MessageBoxResult.OK or MessageBoxResult.Cancel,
        MessageBoxButton.YesNo =>
            result is MessageBoxResult.Yes or MessageBoxResult.No,
        MessageBoxButton.YesNoCancel =>
            result is MessageBoxResult.Yes or MessageBoxResult.No or MessageBoxResult.Cancel,
        _ => false
    };

    private static MessageBoxResult SafeCloseResult(
        MessageBoxButton buttons,
        MessageBoxResult requestedDefault) => buttons switch
    {
        MessageBoxButton.OK => MessageBoxResult.OK,
        MessageBoxButton.OKCancel => MessageBoxResult.Cancel,
        MessageBoxButton.YesNoCancel => MessageBoxResult.Cancel,
        MessageBoxButton.YesNo when requestedDefault == MessageBoxResult.Yes =>
            MessageBoxResult.Yes,
        _ => MessageBoxResult.No
    };

    private static string NormalizeCaption(string caption)
    {
        string normalized = Regex.Replace(
            caption.Trim(),
            @"\s*(?:--|—|-)\s*Halo\s*MCC\s*Toolbox\s*$",
            "",
            RegexOptions.IgnoreCase);
        return string.IsNullOrWhiteSpace(normalized)
            ? "Halo MCC Toolbox"
            : normalized;
    }

    private void Complete(MessageBoxResult result)
    {
        _result = result;
        Close();
    }

    private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ChangedButton != MouseButton.Left)
            return;

        for (DependencyObject? element = e.OriginalSource as DependencyObject;
             element is not null && !ReferenceEquals(element, sender);
             element = VisualTreeHelper.GetParent(element))
        {
            if (element is Button)
                return;
        }

        DragMove();
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e) =>
        Complete(_closeResult);

    private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Escape)
        {
            Complete(_closeResult);
            e.Handled = true;
        }
        else if (e.Key == Key.Enter && _defaultButton is not null)
        {
            _defaultButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            e.Handled = true;
        }
    }
}
