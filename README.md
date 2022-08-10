# Convert Eml to Pst Outlook - Aspose Email

This is easy code for convert .eml to .pst, I share this here because I don't found one free application for convert, very thanks the Aspose for help us, below has the code you need change only variable "dirPath" is should be root path (where has subpaths) or path only should be have the .eml files.

```
using Aspose.Email;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;

var dirPath = "{PATH}";
var outFileName = $"{dirPath}\\outputFile.pst";

if (File.Exists(outFileName))
    File.Delete(outFileName);

using (var personalStorage = PersonalStorage.Create(outFileName, FileFormatVersion.Unicode))
{
    var directories = Directory.GetDirectories(dirPath);
    if (directories.Any())
    {
        foreach (var d in directories)
            SetFilesIntoBox(personalStorage, d);
    }
    else
        SetFilesIntoBox(personalStorage, dirPath);
}

void SetFilesIntoBox(PersonalStorage personalStorage, string directoryPath)
{
    var pathBox = personalStorage.RootFolder.AddSubFolder(Path.GetFileName(directoryPath));
    Console.WriteLine($"Create box: {pathBox.DisplayName}");
    foreach (var f in Directory.GetFiles(directoryPath, "*.eml"))
    {
        using (var message = MailMessage.Load(f))
        {
            pathBox.AddMessage(MapiMessage.FromMailMessage(message, MapiConversionOptions.UnicodeFormat));
            Console.WriteLine($"Add message: {Path.GetFileName(f)}");
        }
    }
}
```

Nuget Package Reference
```
Install-Package Aspose.Email
````

After executed this code, one file the "outputFile.pst" are will create in the root path, this file is easy to import on Outlook by "Export/Import", select the option import by pst.

Note: If you has sub-paths in the root, in the pst will create path too.

