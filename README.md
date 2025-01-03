# OneNoteInstance

A Windows console app for opening content at a specified OneNote URL
in a new OneNote application instance.

## Why?

I often work with multiple OneNote notebooks, sections, or pages open, and I needed
a way to open additional OneNote content in a new OneNote instances.
I couldn't find a way to do that from OneNote's command line, so I created `OneNoteInstance`.
And it allows me to create scripts or Windows shorcuts so I can open specified
content in a new instance of OneNote.

## Usage

```text
OneNoteInstance <OneNoteContentURL> [-?|-h|--help]

Arguments:
  <OneNoteContentURL>  The URL of the OneNote content to open.
                       Must use the 'onenote:' protocol.
                       Wrap the URL in quotation marks.
Options:
  --?|-h|--help        Show this help message and exit.
```

### Example

```powershell
OneNoteInstance.exe "onenote:https://some-url-to-onenote-content"
```
- The URL must use the `onenote:` protocol used for opening content in
the OneNote desktop app.
- Wrap the URL in quotation marks because they often contain characters
that would otherwise be interpreted by the shell you're invoking the
command in.

## Configuration

The path to the OneNote Windows executable (`.exe`) is specified in
the `appsettings.json` file's `OneNoteExePath` setting.

The default path is:
```
C:/Program Files/Microsoft Office/root/Office16/ONENOTE.EXE
```
Edit that setting if you're using a different version,
or have it installed at a different location.

## The Code

### Build
```powershell
cd src/OneNoteInstance

dotnet build
```
