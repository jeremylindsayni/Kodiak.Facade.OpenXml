# Kodiak.Facade.OpenXml

Opinionated wrapper for the OpenXml library.  

The example below shows how to use with the .NET MVC framework. This example depends on [Kodiak.Mvc.Actions](https://github.com/jeremylindsayni/Kodiak.Mvc.Actions), and obviously the [OpenXml](https://www.nuget.org/packages/DocumentFormat.OpenXml) library, and a dotx/docx file with merge fields named "Name" and "addr1".

```C#
public IMsWordService MsWordService { get; set; }

public HomeController()
{
    this.MsWordService = new MsWordService();
}

public ActionResult MsWordDocumentStream()
{
    MsWordService.SetPageTemplate(@"C:\Users\WebApplication\AddressTemplate.dotx");
    MsWordService.MergeTextToFieldInMainDocumentPart("Name", "Jeremy");
    MsWordService.MergeTextToFieldInMainDocumentPart("addr1", "1 Main St");

    return new WordStreamResult(MsWordService.GetMemoryStream(), "doc.docx");
}
```
