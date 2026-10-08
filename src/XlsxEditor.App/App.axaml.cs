using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Markup.Xaml;

namespace XlsxEditor.App;

public partial class App : Application
{
    public override void Initialize() => AvaloniaXamlLoader.Load(this);

    public override void OnFrameworkInitializationCompleted()
    {
        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
        {
            var window = new MainWindow();
            desktop.MainWindow = window;
            // "Open With" on Windows passes the file as an argument.
            if (desktop.Args is [var path, ..]) _ = window.OpenAsync(path);
        }
        base.OnFrameworkInitializationCompleted();
    }
}
