using System.Windows;

namespace ScanBridgeTest;

public sealed class AppMessage
{
    public string Title { get; set; } = "";
    public string Body { get; set; } = "";
    public string? Link { get; set; }
    public Visibility HasLinkVisibility => string.IsNullOrWhiteSpace(Link) ? Visibility.Collapsed : Visibility.Visible;
}