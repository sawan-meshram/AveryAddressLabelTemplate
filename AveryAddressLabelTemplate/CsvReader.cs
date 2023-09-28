using System;
using System.Collections.Generic;
using System.IO;

namespace AveryAddressLabelTemplate
{
    public class CsvReader : IDisposable
    {
        private readonly char delimiter;
        private readonly StreamReader reader;
        private readonly bool ignoreHeader;

        public CsvReader(string filePath, char delimiter = ',', bool ignoreHeader = false)
        {
            this.delimiter = delimiter;
            this.ignoreHeader = ignoreHeader;
            reader = new StreamReader(filePath);
            if (ignoreHeader && !reader.EndOfStream)
            {
                // Skip the header row
                reader.ReadLine();
            }
        }

        public IEnumerable<string[]> ReadAll()
        {
            while (!reader.EndOfStream)
            {
                var line = reader.ReadLine();
                if (line != null)
                {
                    var values = ParseCsvLine(line);
                    yield return values;
                }
            }
        }

        private string[] ParseCsvLine(string line)
        {
            var values = new List<string>();
            bool inQuotes = false;
            int startIndex = 0;

            for (int i = 0; i < line.Length; i++)
            {
                if (line[i] == '"')
                {
                    inQuotes = !inQuotes;
                }

                if (line[i] == delimiter && !inQuotes)
                {
                    values.Add(line.Substring(startIndex, i - startIndex).Trim('"'));
                    startIndex = i + 1;
                }
            }

            values.Add(line.Substring(startIndex).Trim('"'));

            return values.ToArray();
        }

        public void Close()
        {
            reader.Close();
        }

        public void Dispose()
        {
            Close();
        }
    }
}
