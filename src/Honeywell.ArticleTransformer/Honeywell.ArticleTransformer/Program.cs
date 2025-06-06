using System;
using System.IO;
using Honeywell.ArticleTransformer.Core;

namespace Honeywell.ArticleTransformer
{
    public static class Program
    {
        public static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: Honeywell.ArticleTransformer <input.csv|input.xlsx>");
                return;
            }

            var inputPath = args[0];
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            var outputExtension = Path.GetExtension(inputPath).Equals(".xlsx", StringComparison.OrdinalIgnoreCase)
                ? ".xlsx" : ".csv";
            var outputPath = Path.Combine(Path.GetDirectoryName(inputPath) ?? string.Empty, $"output{outputExtension}");

            Transformer.ProcessFile(inputPath, outputPath);
            Console.WriteLine($"Processed file saved to {Path.GetFileName(outputPath)}");
        }
    }
}
