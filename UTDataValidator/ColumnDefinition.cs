using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace UTDataValidator
{
    public class ColumnDefinition
    {
        public ColumnDefinition(string columnValue, int columnIndex)
        {
            NeedValidation = true;
            ColumnName = columnValue;
            ColumnIndex = columnIndex;
            if (columnValue.Contains(":"))
            {
                string[] split = columnValue.Split(':');
                ColumnName = split[0].Trim();

                var extensions = split[1]
                    .Split('|')
                    .Where(f => !string.IsNullOrWhiteSpace(f))
                    .Select(f => f.Trim())
                    .ToList();

                var regex = new Regex(@"(?:validation)\s*\(([\S\s]+)\)", RegexOptions.IgnoreCase | RegexOptions.Multiline);
                foreach (var extension in extensions)
                {   
                    if (extension == "0")
                    {
                        NeedValidation = false;
                        continue;
                    }

                    var match = regex.Match(extension);
                    if (match.Success)
                    {
                        var group1 = match.Groups[1].Value;
                        var validations = group1.Split(',')
                            .Where(f => !string.IsNullOrWhiteSpace(f))
                            .Select(f => f.Trim().ToUpper())
                            .ToList();

                        foreach (var value in validations)
                        {
                            CustomValidations.Add(value);
                        }
                    }
                }
            }
        }

        public string ColumnName { get; private set; }
        public bool NeedValidation { get; private set; }
        public List<string> CustomValidations { get; set; } = new List<string>();
        public int ColumnIndex { get; set; }
    }
}
