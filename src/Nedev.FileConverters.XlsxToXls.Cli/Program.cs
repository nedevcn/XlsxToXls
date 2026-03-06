using System;
using Nedev.FileConverters.XlsxToXls;
using Nedev.FileConverters.Core;
using Microsoft.Extensions.DependencyInjection;

static void ShowUsage()
{
    Console.WriteLine("Usage: XlsxToXls.Cli <input.xlsx> <output.xls>");
    Console.WriteLine("Converts an XLSX file to XLS format using the Nedev.FileConverters.XlsxToXls library (via core abstraction).");
}

if (args.Length != 2)
{
    ShowUsage();
    return;
}

var input = args[0];
var output = args[1];

// set up dependency injection and register available converters
IServiceCollection services = new ServiceCollection();
services.AddFileConverter("xlsx", "xls", new XlsxToXlsConverter.FileConverterAdapter());
var provider = services.BuildServiceProvider();
var converter = provider.GetRequiredService<IFileConverter>();

try
{
    Console.WriteLine($"Converting '{input}' to '{output}' using core abstraction...");
    using var inStream = File.OpenRead(input);
    using var result = converter.Convert(inStream);
    using var outStream = File.Create(output);
    result.CopyTo(outStream);
    Console.WriteLine("Conversion completed.");
}
catch (Exception ex)
{
    Console.Error.WriteLine("Error: " + ex.Message);
    Environment.Exit(1);
}
