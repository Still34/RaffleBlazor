using System;
using System.Threading.Tasks;
using Blazor.FileReader;
using RaffleBlazor.Utility;
using Blazorise;
using Blazorise.Bootstrap;
using Blazorise.Icons.FontAwesome;
using Microsoft.AspNetCore.Components.WebAssembly.Hosting;
using Microsoft.Extensions.DependencyInjection;
using ZXing;
using ZXing.QrCode;
using ZXing.QrCode.Internal;
using ZXing.Rendering;

namespace RaffleBlazor
{
    public class Program
    {
        public static async Task Main(string[] args)
        {
            var builder = WebAssemblyHostBuilder.CreateDefault(args);

            builder.RootComponents.Add<App>("app");

            var qrWriter = QrWriterFactory();

            builder.Services
                .AddBlazorise(options =>
                {
                    options.ChangeTextOnKeyPress = true;
                })
                .AddBootstrapProviders()
                .AddFontAwesomeIcons()
                .AddBaseAddressHttpClient()
                .AddSingleton(qrWriter)
                .AddTransient<ExportProvider>()
                .AddTransient<EnumerableUtility>()
                .AddSingleton<Random>()
                .AddFileReaderService(options => options.UseWasmSharedBuffer = true);

            var host = builder.Build();
            
            host.Services
                .UseBootstrapProviders()
                .UseFontAwesomeIcons();

            await host.RunAsync();
        }
        public static BarcodeWriterSvg QrWriterFactory()
        {
            return new BarcodeWriterSvg
            {
                Format = BarcodeFormat.QR_CODE,
                Options = new QrCodeEncodingOptions() {ErrorCorrection = ErrorCorrectionLevel.M},
                Renderer = new SvgRenderer()
            };
        }
    }
}