using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace Honeywell.ArticleTransformer.Core
{
    public static class Transformer
    {
        public static void ProcessFile(string inputPath, string outputPath)
        {
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file '{inputPath}' not found.");

            var lines = Path.GetExtension(inputPath).Equals(".xlsx", StringComparison.OrdinalIgnoreCase)
                ? ReadXlsx(inputPath)
                : File.ReadAllLines(inputPath);

            var outputLines = lines.Select(ProcessLine).ToArray();

            if (Path.GetExtension(outputPath).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                WriteXlsx(outputPath, outputLines);
            }
            else
            {
                File.WriteAllText(outputPath, string.Join("\r\n", outputLines), Encoding.UTF8);
            }
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

        private static string[] ReadXlsx(string path)
        {
            using var archive = ZipFile.OpenRead(path);
            var sheetEntry = archive.GetEntry("xl/worksheets/sheet1.xml");
            if (sheetEntry == null)
                return Array.Empty<string>();

            var strings = Array.Empty<string>();
            var stringsEntry = archive.GetEntry("xl/sharedStrings.xml");
            if (stringsEntry != null)
            {
                using var reader = new StreamReader(stringsEntry.Open());
                var doc = XDocument.Load(reader);
                var ns = doc.Root!.Name.Namespace;
                strings = doc.Descendants(ns + "t").Select(t => t.Value).ToArray();
            }

            using var sheetReader = new StreamReader(sheetEntry.Open());
            var sheetDoc = XDocument.Load(sheetReader);
            var n = sheetDoc.Root!.Name.Namespace;
            var rows = sheetDoc.Descendants(n + "row");
            return rows.Select(r => string.Join(';', r.Elements(n + "c").Select(c =>
            {
                var t = c.Attribute("t")?.Value;
                if (t == "s")
                {
                    var v = c.Element(n + "v")?.Value;
                    if (int.TryParse(v, out var idx) && idx >= 0 && idx < strings.Length)
                        return strings[idx];
                    return string.Empty;
                }
                if (t == "inlineStr")
                {
                    return c.Element(n + "is")?.Element(n + "t")?.Value ?? string.Empty;
                }
                return c.Element(n + "v")?.Value ?? string.Empty;
            }))).ToArray();
        }

        private static void WriteXlsx(string path, string[] lines)
        {
            if (File.Exists(path))
                File.Delete(path);

            using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
            AddEntry(archive, "[Content_Types].xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
                "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
                "</Types>");

            AddEntry(archive, "_rels/.rels",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
                "</Relationships>");

            AddEntry(archive, "xl/_rels/workbook.xml.rels",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>" +
                "</Relationships>");

            AddEntry(archive, "xl/workbook.xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                "<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets>" +
                "</workbook>");

            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>");
            for (var r = 0; r < lines.Length; r++)
            {
                var cells = lines[r].Split(';');
                sb.Append($"<row r=\"{r + 1}\">");
                for (var c = 0; c < cells.Length; c++)
                {
                    var cellRef = ColumnName(c) + (r + 1);
                    sb.Append($"<c r=\"{cellRef}\" t=\"inlineStr\"><is><t>{EscapeXml(cells[c])}</t></is></c>");
                }
                sb.Append("</row>");
            }
            sb.Append("</sheetData></worksheet>");
            AddEntry(archive, "xl/worksheets/sheet1.xml", sb.ToString());
        }

        private static void AddEntry(ZipArchive archive, string path, string content)
        {
            var entry = archive.CreateEntry(path);
            using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
            writer.Write(content);
        }

        private static string ColumnName(int index)
        {
            var dividend = index + 1;
            var columnName = string.Empty;
            while (dividend > 0)
            {
                var modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }
            return columnName;
        }

        private static string EscapeXml(string value)
        {
            return SecurityElement.Escape(value) ?? string.Empty;
        }
    }
}
