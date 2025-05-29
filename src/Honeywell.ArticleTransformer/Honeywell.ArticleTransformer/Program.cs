using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Honeywell.ArticleTransformer
{
    public static class Program
    {
        public static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: Honeywell.ArticleTransformer <input.csv>");
                return;
            }

            var inputPath = args[0];
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            var lines = File.ReadAllLines(inputPath);
            var outputLines = lines.Select(ProcessLine).ToArray();
            File.WriteAllText("output.csv", string.Join("\r\n", outputLines), Encoding.UTF8);
            Console.WriteLine("Processed file saved to output.csv");
        }

        private static string ProcessLine(string line)
        {
            var cells = line.Split(';');
            for (var i = 0; i < cells.Length; i++)
            {
                cells[i] = TransformCell(cells[i]);
            }

            return string.Join(';', cells);
        }

        private static string TransformCell(string cell)
        {
            if (string.IsNullOrEmpty(cell))
                return cell;

            var text = cell.Replace("_x000D_", string.Empty)
                            .Replace("\r\n", " ")
                            .Replace("\n", " ")
                            .Replace("\r", " ");

            text = Regex.Replace(text, " {2,}", " ");

            text = Regex.Replace(text, "\\s+(Anschl\\u00fcsse:|Technische Daten:|Betriebsspannung:)", "\r\n$1");

            text = Regex.Replace(text, "(?<!Technische\\s)Daten:", m => "\r\n" + m.Value);
            text = Regex.Replace(text, "Daten:\\s*$", "Daten:\r\n");

            text = Regex.Replace(text, ":\\s*", ":\r\n");
            text = Regex.Replace(text, "\\s*-", "\r\n-" );

            return text.Trim();
        }
    }
}
