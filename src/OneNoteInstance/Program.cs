using System.Diagnostics;
using Microsoft.Extensions.Configuration;

if (args.Contains("--help") || args.Contains("-h") || args.Contains("-?"))
{
    WriteHelpMessage();
    return 0;
}

string? oneNoteContentUrl = ValidateArgsAndGetOneNoteContentUrl(args);
if (oneNoteContentUrl is null)
{
    WriteHelpMessage();
    return 1;
}

string? oneNoteExePath = GetOneNoteExePathFromAppSettings();
if (string.IsNullOrEmpty(oneNoteExePath))
{
    return 1;
}

// Open a new instance of the OneNote application.

Process? oneNoteProcess = Process.Start(new ProcessStartInfo
    {
        FileName = oneNoteExePath,
        UseShellExecute = true
    });

if (oneNoteProcess is null)
{
    Console.Error.WriteLine("Failed to start new OneNote application instance.");
    return 1;
}

// Wait for the process to be ready for input with a timeout of 15 seconds.
if (!oneNoteProcess.WaitForInputIdle(15000))
{
    Console.Error.WriteLine("Timed-out waiting for new OneNote application instance.");
    return 1;
}

// Open the specified OneNote content.
// We let the Windows Shell open the URL with the default
// app registered for the `onenote:` protocol.
// Assuming that app is OneNote.
// It will open the content in the most recent instance
// of the OneNote application. The one we just started.

Process.Start(new ProcessStartInfo
    {
        FileName = oneNoteContentUrl,
        UseShellExecute = true
    });

return 0;

#region Helper Methods

static string? GetOneNoteExePathFromAppSettings()
{
    IConfiguration config = new ConfigurationBuilder()
        .SetBasePath(AppContext.BaseDirectory)
        .AddJsonFile("appSettings.json", optional: false, reloadOnChange: true)
        .Build();

    string? oneNoteExePath = config["OneNoteExePath"];

    if (string.IsNullOrEmpty(oneNoteExePath))
    {
        Console.Error.WriteLine("The 'OneNoteExePath' setting is not set in the appSettings.json file.");
    }

    return oneNoteExePath;
}

static string? ValidateArgsAndGetOneNoteContentUrl(string[] args)
{
    if (args.Length != 1)
    {
        Console.Error.WriteLine("<OneNoteContentURL> argument is required.");
        return null;
    }

    string oneNoteContentUrl = args[0];

    // Ensure the URL specifies the "onenote:" protocol.
    if (!oneNoteContentUrl.StartsWith("onenote:", StringComparison.OrdinalIgnoreCase))
    {
        Console.Error.WriteLine("The <OneNoteContentURL> must start with the 'onenote:' protocol.");
        Console.Error.WriteLine($"  Given URL: {oneNoteContentUrl}");
        return null;
    }

    return oneNoteContentUrl;
}

static void WriteHelpMessage()
{
    Console.WriteLine();
    Console.WriteLine("Usage: OneNoteInstance <OneNoteContentURL> [-?|-h|--help]");
    Console.WriteLine("Arguments:");
    Console.WriteLine("  <OneNoteContentURL>  The URL of the OneNote content to open.");
    Console.WriteLine("                       Must use the 'onenote:' protocol.");
    Console.WriteLine("                       Wrap the URL in quotation marks.");
    Console.WriteLine("Options:");
    Console.WriteLine("  --?|-h|--help        Show this help message and exit.");
}

#endregion Helper Methods
