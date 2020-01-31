using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using System.Text.RegularExpressions;


namespace ExcelMazRegex
{
    public static class ExcelMazRegex
    {
        [ExcelFunction(Description = "Find the pattern in the input, return the first matching string.")]

        public static object RegexMatch(
            [ExcelArgument(Name = "input", Description = "input text")]
            String input,
            [ExcelArgument(Name = "pattern", Description = "regular expression pattern to search for")]
            String pattern,
            [ExcelArgument
                (
                    Name = "options",
                    Description = "Regular expression option flags, add value for each requested option: \n" +
                    "IgnoreCase = 1, " +
                    "Multiline = 2, " +
                    "ExplicitCapture = 4, " +
                    "Compiled = 8, " +
                    "Singleline = 16, " +
                    "IgnorePatternWhitespace = 32, " +
                    "RightToLeft = 64, " +
                    "ECMAScript = 256, " +
                    "CultureInvariant = 512"
                )
            ]
            int options,
            [ExcelArgument(Name = "replacement", Description = "optional replacement pattern for found text")]
            String replacement
            )
        {
            // the replacement string is not checked for an empty string because that is a valid replacement pattern
            if (String.IsNullOrEmpty(input) || String.IsNullOrEmpty(pattern) )
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorValue;
            }
            else
            {
                RegexOptions ro = (RegexOptions)options;
                Regex rx = new Regex(pattern, ro);
                Match rm = rx.Match(input);
                if (rm.Success)
                {
                    if ( string.IsNullOrEmpty( replacement ) )
                    {
                        return rm.Value;
                    } else
                    {
                        return rm.Result(replacement);
                    }
                } 
                {
                    return ExcelDna.Integration.ExcelError.ExcelErrorNA;
                }
            }
        }


        [ExcelFunction(Description = "Find the pattern in the input, substitute all matches with the replacement pattern")]

        public static object RegexReplace(
            [ExcelArgument(Name = "input", Description = "input text")]
            String input,
            [ExcelArgument(Name = "pattern", Description = "regular expression pattern to search for")]
            String pattern,
            [ExcelArgument
                (
                    Name = "options",
                    Description = "Regular expression option flags, add value for each requested option: \n" +
                    "IgnoreCase = 1, " +
                    "Multiline = 2, " +
                    "ExplicitCapture = 4, " +
                    "Compiled = 8, " +
                    "Singleline = 16, " +
                    "IgnorePatternWhitespace = 32, " +
                    "RightToLeft = 64, " +
                    "ECMAScript = 256, " +
                    "CultureInvariant = 512"
                )
            ]
            int options,
            [ExcelArgument(Name = "replacement", Description = "replacement for each occurrence found")]
            String replacement
            )
        {
            // the replacement string is not checked for an empty string because that is a valid replacement pattern
            if (String.IsNullOrEmpty(input) || String.IsNullOrEmpty(pattern) || replacement == null)
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorValue;
            }
            else
            {
                RegexOptions ro = (RegexOptions)options;
                Regex rx = new Regex(pattern, ro);
                return rx.Replace(input, replacement);
            }
        }

        [ExcelFunction(Description = "My first .NET function", IsThreadSafe=true)]
        public static string HelloWorld(
            [ExcelArgument(Name = "name", Description = "foo!")]
            string name)
        {
            return "Hello " + name;
        }
    }
}
