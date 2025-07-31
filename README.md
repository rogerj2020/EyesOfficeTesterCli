# EyesOfficeTesterCli

A command-line utility for executing visual testing of Microsoft Word documents and Microsoft Excel workbooks using [Applitools Eyes](https://applitools.com/platform/eyes/).

## Description

EyesOfficeTesterCli is a command-line utility that enables use of [Applitools Eyes](https://applitools.com/platform/eyes/) to visually test Microsoft Word documents and Microsoft Excel workbooks stored in a specified directory, when running tests under the Windows operating system (Windows 10 or newer). EyesOfficeTesterCli can be used to integrate visual testing of Microsoft Word documents and Microsoft Excel workbooks into any test automation pipeline where .NET 8.0 and Microsoft Office 2016 (or newer) are installed, within a Windows-based execution environment. Please see below for more details regarding available command line arguments and options, when using EyesOfficeTesterCli.

## Getting Started

### Dependencies

* An [Applitools Eyes](https://applitools.com/platform/eyes/) API Key
* Windows 10 or newer
* Microsoft .NET 8.0 or newer
* Microsoft Office 2016 or newer
* Microsoft Office Interop Libraries (integrated into this project via Nuget)

### Installing

* Option 1: Download or clone this repo, and build the solution. (And optionally copy all files from the build output to a permanent location for future test execution.)
* Option 2: Download a pre-built zip package from the Releases page, and extract it to a permanent location for future test execution.

### Executing EyesOfficeTesterCli

* Run the `EyesOfficeTesterCli.exe` binary with any of the command line arguments provided below.<br><br>
    
    ```
    EyesOfficeTesterCli.exe --directory C:\OfficeFiles -a <API_KEY> -u https://eyesapi.applitools.com --saveImages --notify --failOnDiff
    ```

* Command Line Options / Arguments
    | Argument | Type | Description | Default |
    |---|---|---|---|
    | `-d`, `--directory` | String | Specify the directory to scan for Microsoft Word documents and Microsoft Excel workbooks.<br>(Optional. EyesOfficeTesterCli.exe will scan the directory where `EyesOfficeTesterCli.exe` is located if not specified.) | `None` |
    | `-a`, `--apiKey` | String | Specify the Applitools Eyes API Key.<br>(Optional. EyesOfficeTesterCli.exe will also check for the Applitools Eyes API key in the `APPLITOOLS_API_KEY` environment variable if this option is not specified.) | `None` |
    | `-u`, `--serverUrl` | String | Specify the Applitools Eyes Server URL.<br>(Optional. EyesOfficeTesterCli.exe will also check for the Applitools Eyes Server URL in the `APPLITOOLS_SERVER_URL` environment variable if this option is not specified.) | `https://eyesapi.applitools.com` |
    | `-s`, `--saveImages` | Boolean (Switch) | Save image files for captured screenshots of pages and sheets, in the directory where `EyesOfficeTesterCli.exe` is located. (Optional)| `false` |
    | `-p`, `--progressBar` | Boolean (Switch) | Display a progress bar for command output. (Optional. Not Recommended in pipelines.) | `false` |
    | `-n`, `--notify` | Boolean (Switch) | Send Applitools Eyes Batch Notification when all visual tests have completed. (Optional) | `false` |
    | `-f`, `--failOnDiff` | Boolean (Switch) | Throw an exception to exit with error status if visual diffs are detected when all tests have completed.<br>(Optional. Helpful for indicating failure status when visual diffs are detected during pipeline executions.) | `false` |




## Help

For help, you can contact the Applitools Support Team at: support@applitools.com, reference the code in this repo, and mention that this utility uses the .NET Eyes.Images SDK.


## Authors

Roger Jefferies (germdman@gmail.com)  



## Acknowledgments

Inspiration, code snippets, etc.
* [Applitools](https://applitools.com/)
* [Applitools Eyes](https://applitools.com/platform/eyes/)
* [awesome-readme](https://github.com/matiassingers/awesome-readme)