# SpreadsheetLight Cross Platform

[![MIT](https://img.shields.io/github/license/NF-Software-Inc/spreadsheetlight-cross-platform)](https://github.com/NF-Software-Inc/spreadsheetlight-cross-platform/blob/master/license.txt)
[![NuGet](https://img.shields.io/nuget/v/SpreadsheetLight.Cross.Platform.svg)](https://www.nuget.org/packages/SpreadsheetLight.Cross.Platform/)
[![Build](https://img.shields.io/github/actions/workflow/status/NF-Software-Inc/spreadsheetlight-cross-platform/build.yml)](https://github.com/NF-Software-Inc/spreadsheetlight-cross-platform/actions/workflows/build.yml)
[![Publish](https://img.shields.io/github/actions/workflow/status/NF-Software-Inc/spreadsheetlight-cross-platform/publish.yml?label=publish)](https://github.com/NF-Software-Inc/spreadsheetlight-cross-platform/actions/workflows/publish.yml)

This project is a fork of the [SpreadsheetLight library](https://spreadsheetlight.com/).

The purpose of this fork is to create a version of the library that runs on .NET Core and is also capable of running in cross-platform environments.

The library uses System.Drawing.Common, which required some platform checks to ensure those code sections only run on Windows machines. This does mean that certain features will not work on other platforms, these features will typically silently fail. Identified so far:

* `.InsertPicture()` method [Issue Reference](https://github.com/NF-Software-Inc/spreadsheetlight-cross-platform/issues/10)

## Getting Started

These instuctions can be used to acquire and implement the library.

### Installation

To use this library either clone a copy of the repository or check out the [NuGet package](https://www.nuget.org/packages/SpreadsheetLight.Cross.Platform/)

### Usage

Here are some examples, there are more in the examples directory. You can run all of them by executing the **Test-Examples.ps1** PowerShell script found in that directory.

* [The Hello World](https://github.com/NF-Software-Inc/spreadsheetlight-cross-platform/HelloWorld/HelloWorld.cs)
* [How to modify an existing spreadsheet](https://github.com/NF-Software-Inc/spreadsheetlight-cross-platform/ModifyExistingSpreadsheet/ModifyExistingSpreadsheet.cs)
* [How to format numbers and dates](https://github.com/NF-Software-Inc/spreadsheetlight-cross-platform/NumberFormat/NumberFormat.cs)
* [How to set font settings](https://github.com/NF-Software-Inc/spreadsheetlight-cross-platform/Font/Font.cs)
* [How to copy cells](https://github.com/NF-Software-Inc/spreadsheetlight-cross-platform/CopyCell/CopyCell.cs)
* [How to merge and unmerge cells](https://github.com/NF-Software-Inc/spreadsheetlight-cross-platform/MergeCells/MergeCells.cs)
* [How to autofit row heights and column widths](https://github.com/NF-Software-Inc/spreadsheetlight-cross-platform/AutoFitRowColumn/AutoFitRowColumn.cs)
* [How to insert hyperlinks](https://github.com/NF-Software-Inc/spreadsheetlight-cross-platform/Hyperlinks/Hyperlinks.cs)
* [How to insert tables](https://github.com/NF-Software-Inc/spreadsheetlight-cross-platform/Tables/Tables.cs)
* [How to import data from a DataTable](https://github.com/NF-Software-Inc/spreadsheetlight-cross-platform/ImportDataTable/ImportDataTable.cs)

## Authors

* **Vincent Tan**
* **NF Software Inc.**

## License

This project is licensed under the MIT License - see the [LICENSE](license.txt) file for details

## Acknowledgments

Thank you to:

* [Freepik](https://www.flaticon.com/authors/freepik) for the project icon
