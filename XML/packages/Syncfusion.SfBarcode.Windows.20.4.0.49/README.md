### Syncfusion WinForms Barcode Component
The Syncfusion [WinForms Barcode](https://www.syncfusion.com/winforms-ui-controls/barcode?utm_source=nuget&utm_medium=listing&utm_campaign=winforms-barcode-nuget) for Windows Forms is a .NET control that renders barcode in any .NET application without requiring fonts. It supports major 1D and 2D barcodes including Code 128 and QR code.

![Windows Barcode](https://cdn.syncfusion.com/nuget-readme/winforms/winforms-barcode-generator.png)

[Features overview](https://www.syncfusion.com/winforms-ui-controls/barcode?utm_source=nuget&utm_medium=listing&utm_campaign=winforms-barcode-nuget) | [Docs](https://help.syncfusion.com/file-formats/pdf/working-with-barcode) | [API Reference](https://help.syncfusion.com/cr/file-formats/Syncfusion.Pdf.Barcode.html?utm_source=nuget&utm_medium=listing&utm_campaign=winforms-barcode-nuget) | [Online Demo](https://ej2.syncfusion.com/demos/#/bootstrap5/barcode/ean8.html?utm_source=nuget&utm_medium=listing&utm_campaign=winforms-barcode-nuget) | [Blogs](https://www.syncfusion.com/blogs/?s=barcode?utm_source=nuget&utm_medium=listing&utm_campaign=winforms-barcode-nuget) | [Support](https://support.syncfusion.com/create?utm_source=nuget&utm_medium=listing&utm_campaign=winforms-barcode-nuget) | [Forums](https://www.syncfusion.com/forums?utm_source=nuget&utm_medium=listing&utm_campaign=winforms-barcode-nuget) | [Feedback](https://www.syncfusion.com/feedback/pdf?utm_source=nuget&utm_medium=listing&utm_campaign=winforms-barcode-nuget).

### Key Features

* Complete customization of barcode to augment its appearance.
* Support most common 1D and 2D barcodes.
* Easy to use and integrate.
* QR code
* Data Matrix
* Code 39
* Code 39 Extended
* Code 128
* Code 11
* Codabar
* Code 93
* Code 93 Extended
* Code 32
* UPC

### System Requirements

* [System Requirements](https://help.syncfusion.com/file-formats/installation-and-upgrade/system-requirements?utm_source=nuget&utm_medium=listing&utm_campaign=winforms-barcode-nuget).

### Getting started

You can fetch the Syncfusion .NET PDF library NuGet by simply running the command Install-Package [Syncfusion.SfBarcode.Windows](https://www.nuget.org/packages/Syncfusion.SfBarcode.Windows/?utm_source=nuget&utm_medium=listing&utm_campaign=winforms-barcode-nuget) from the Package Manager Console in Visual Studio.

Try the following code example to create a hello world PDF document.

```csharp
//Creating new PDF Document
PdfDocument doc = new PdfDocument();
//Adding new page to PDF document
PdfPage page = doc.Pages.Add();
//Drawing Code39 barcode 
PdfCode39Barcode barcode = new PdfCode39Barcode(); 
//Setting height of the barcode 
barcode.BarHeight = 45; 
barcode.Text = "CODE39$"; 
//Printing barcode on to the Pdf. 
barcode.Draw(page, new PointF(25, 70 ));
//Saving the Documents
doc.Save("CODE39.pdf");
```

### License

This is a commercial product and requires a paid license for possession or use. Syncfusion’s licensed software, including this component, is subject to the terms and conditions of [Syncfusion's EULA](https://www.syncfusion.com/eula/es/?utm_source=nuget&utm_medium=listing&utm_campaign=winforms-barcode-nuget). You can purchase a license [here](https://www.syncfusion.com/sales/products?utm_source=nuget&utm_medium=listing&utm_campaign=winforms-barcode-nuget) or start a free 30-day trial [here](https://www.syncfusion.com/account/manage-trials/start-trials?utm_source=nuget&utm_medium=listing&utm_campaign=winforms-barcode-nuget).

### About Syncfusion

Founded in 2001 and headquartered in Research Triangle Park, N.C., Syncfusion has more than 27,000+ customers and more than 1 million users, including large financial institutions, Fortune 500 companies, and global IT consultancies.
 
Today, we provide 1700+ components and frameworks for web ([Blazor](https://www.syncfusion.com/blazor-components?utm_source=nuget&utm_medium=listing&utm_campaign=winforms-barcode-nuget), [Flutter](https://www.syncfusion.com/flutter-widgets?utm_source=nuget&utm_medium=listing&utm_campaign=winforms-barcode-nuget), [ASP.NET Core](https://www.syncfusion.com/aspnet-core-ui-controls?utm_source=nuget&utm_medium=listing&utm_campaign=winforms-barcode-nuget), [ASP.NET MVC](https://www.syncfusion.com/aspnet-mvc-ui-controls?utm_source=nuget&utm_medium=listing&utm_campaign=winforms-barcode-nuget), [ASP.NET Web Forms](https://www.syncfusion.com/jquery/aspnet-webforms-ui-controls?utm_source=nuget&utm_medium=listing&utm_campaign=winforms-barcode-nuget), [JavaScript](https://www.syncfusion.com/javascript-ui-controls?utm_source=nuget&utm_medium=listing&utm_campaign=winforms-barcode-nuget), [Angular](https://www.syncfusion.com/angular-ui-components?utm_source=nuget&utm_medium=listing&utm_campaign=winforms-barcode-nuget), [React](https://www.syncfusion.com/react-ui-components?utm_source=nuget&utm_medium=listing&utm_campaign=winforms-barcode-nuget), [Vue](https://www.syncfusion.com/vue-ui-components?utm_source=nuget&utm_medium=listing&utm_campaign=winforms-barcode-nuget), and [jQuery](https://www.syncfusion.com/jquery-ui-widgets?utm_source=nuget&utm_medium=listing&utm_campaign=winforms-barcode-nuget)), mobile ([.NET MAUI (Preview)](https://www.syncfusion.com/maui-controls?utm_source=nuget&utm_medium=listing&utm_campaign=winforms-barcode-nuget), [Flutter](https://www.syncfusion.com/flutter-widgets?utm_source=nuget&utm_medium=listing&utm_campaign=winforms-barcode-nuget), [Xamarin](https://www.syncfusion.com/xamarin-ui-controls?utm_source=nuget&utm_medium=listing&utm_campaign=winforms-barcode-nuget), [UWP](https://www.syncfusion.com/uwp-ui-controls?utm_source=nuget&utm_medium=listing&utm_campaign=winforms-barcode-nuget), and [JavaScript](https://www.syncfusion.com/javascript-ui-controls?utm_source=nuget&utm_medium=listing&utm_campaign=winforms-barcode-nuget)), and desktop development ([WinForms](https://www.syncfusion.com/winforms-ui-controls?utm_source=nuget&utm_medium=listing&utm_campaign=winforms-barcode-nuget), [WPF](https://www.syncfusion.com/wpf-controls?utm_source=nuget&utm_medium=listing&utm_campaign=winforms-barcode-nuget), [WinUI](https://www.syncfusion.com/winui-controls?utm_source=nuget&utm_medium=listing&utm_campaign=winforms-barcode-nuget), [.NET MAUI (Preview)](https://www.syncfusion.com/maui-controls?utm_source=nuget&utm_medium=listing&utm_campaign=winforms-barcode-nuget), [Flutter](https://www.syncfusion.com/flutter-widgets?utm_source=nuget&utm_medium=listing&utm_campaign=winforms-barcode-nuget), [Xamarin](https://www.syncfusion.com/xamarin-ui-controls?utm_source=nuget&utm_medium=listing&utm_campaign=winforms-barcode-nuget), and [UWP](https://www.syncfusion.com/uwp-ui-controls?utm_source=nuget&utm_medium=listing&utm_campaign=winforms-barcode-nuget)). We provide ready-to-deploy enterprise software for dashboards, reports, data integration, and big data processing. Many customers have saved millions in licensing fees by deploying our software.

[sales@syncfusion.com](mailto:sales@syncfusion.com?Subject=Syncfusion%20WinForms%20Barcode-%20NuGet) | [www.syncfusion.com](https://www.syncfusion.com?utm_source=nuget&utm_medium=listing&utm_campaign=winforms-barcode-nuget) | Toll Free: 1-888-9 DOTNET


