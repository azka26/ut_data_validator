using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace UTDataValidator
{
    public static class KeyWordValidator
    {
        private static Regex TablePattern => new Regex(@"(table|tabel)(\s*)([:])(\s*)(\w+)", RegexOptions.IgnoreCase);
        
        public static bool IsTableInfo(this string value)
        {
            var regex = TablePattern;
            if (!regex.IsMatch(value))
            {
                return false;
            }
            
            var match = regex.Match(value);
            return match.Groups[0].Value.Trim() == value.Trim();
        }

        public static string GetTableName(this string value)
        {
            var regex = TablePattern;
            if (!regex.IsMatch(value))
            {
                throw new Exception($"Invalid Format Table Name = \"{value}\".");
            }
            
            var match = regex.Match(value);
            if (match.Groups[0].Value.Trim() != value.Trim())
            {
                throw new Exception($"Invalid Format Table Name = \"{value}\".");
            }
            
            return match.Groups[5].Value;
        }
    }
}