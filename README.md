# 選股志願抽號器

[Live Demo](https://vote.stillu.cc)

## Why

Because sorting preferences by hand is too much of a hassle. That and I just wanted to write a Blazor app for fun 😛.

## Dependencies

- ZXing.Net for QR code generation
- Tewr.Blazor.FileReader for CSV uploading
- CsvHelper for CSV parsing
- Blazorise for frontend component

As of 2020/04/13:
```
Project 'RaffleBlazor' has the following package references
   [netstandard2.1]:
   Top-level Package                                            Requested                Resolved
   > Blazorise.Bootstrap                                        0.9.0-preview3           0.9.0-preview3
   > Blazorise.Icons.FontAwesome                                0.9.0-preview3           0.9.0-preview3
   > Blazorise.Sidebar                                          0.9.0-preview3           0.9.0-preview3
   > Blazorise.Snackbar                                         0.9.0-preview3           0.9.0-preview3
   > CsvHelper                                                  15.0.4                   15.0.4
   > DocumentFormat.OpenXml                                     2.10.1                   2.10.1
   > Microsoft.AspNetCore.Blazor.HttpClient                     3.2.0-preview3.20168.3   3.2.0-preview3.20168.3
   > Microsoft.AspNetCore.Components.WebAssembly                3.2.0-preview3.20168.3   3.2.0-preview3.20168.3
   > Microsoft.AspNetCore.Components.WebAssembly.Build          3.2.0-preview3.20168.3   3.2.0-preview3.20168.3
   > Microsoft.AspNetCore.Components.WebAssembly.DevServer      3.2.0-preview3.20168.3   3.2.0-preview3.20168.3
   > Tewr.Blazor.FileReader                                     1.5.0.20093              1.5.0.20093
   > ZXing.Net                                                  0.16.5                   0.16.5
```