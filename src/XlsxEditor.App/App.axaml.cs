using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Markup.Xaml;
using Avalonia.Platform.Storage;

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
            // On macOS, Finder (double-click / Open With) sends it as an activation event instead.
            if (this.TryGetFeature<IActivatableLifetime>() is { } activatable)
                activatable.Activated += (_, e) =>
                {
                    if (e is FileActivatedEventArgs { Files: [var file, ..] } && file.TryGetLocalPath() is { } p) _ = window.OpenAsync(p);
                };
            // Cmd+Q / logout: let the window ask about unsaved changes first.
            desktop.ShutdownRequested += (_, e) =>
            {
                if (!window.HasUnsavedChanges) return;
                e.Cancel = true;
                window.Close(); // prompts; quits if the user saves or discards
            };
        }
        base.OnFrameworkInitializationCompleted();
    }
}
