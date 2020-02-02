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
            [ExcelArgument( Name = "input",     Description = "Input text"  )]
            String input,
            [ExcelArgument( Name = "pattern",   Description = "Regular expression pattern to search for"    )]
            String pattern,
            [ExcelArgument( Name = "options",   Description = "Regex option flags, sum for each active option: IgnoreCase = 1, Multiline = 2, ExplicitCapture = 4, Compiled = 8, Singleline = 16, IgnorePatternWhitespace = 32, RightToLeft = 64, ECMAScript = 256, CultureInvariant = 512" ) ]
            int options,
            [ExcelArgument( Name = "replacement", Description = "Optional replacement pattern for result")]
            String replacement
        )
        {
            // the replacement string is not checked for an empty string or null because empty is a valid replacement pattern and null when ommited (it's optional)
            if (String.IsNullOrEmpty(input) || String.IsNullOrEmpty(pattern) )
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorValue;
            }
            else
            {
                RegexOptions ro = (RegexOptions)options;
                if (string.IsNullOrEmpty(replacement))
                {
                    return Regex.Match(input, pattern, ro).Value;
                }
                else
                {
                    return Regex.Match(input, pattern, ro).Result(replacement);
                    // if not found, how do we return the NA? Test first with plain Result, see how it's returned
                    // return ExcelDna.Integration.ExcelError.ExcelErrorNA;
                }
            }
        }



        [ExcelFunction(Description = "Find the pattern in the input, substitute all matches with the replacement pattern")]

        public static object RegexReplace(
            [ExcelArgument( Name = "input",     Description = "Input text"  )]
            String input,
            [ExcelArgument( Name = "pattern",   Description = "Regular expression pattern to search for"    )]
            String pattern,
            [ExcelArgument( Name = "options",   Description = "Regex option flags, sum for each active option: IgnoreCase = 1, Multiline = 2, ExplicitCapture = 4, Compiled = 8, Singleline = 16, IgnorePatternWhitespace = 32, RightToLeft = 64, ECMAScript = 256, CultureInvariant = 512" ) ]
            int options,
            [ExcelArgument( Name = "replacement", Description = "Replacement pattern for result")]
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
                // Changed to static function call because caching
                return Regex.Replace(input, pattern, replacement, ro);
            }
        }


        [ExcelFunction(Description = "Return string with special characters escaped to remove special meaning", IsThreadSafe = true)]

        public static string RegexEscape(
            [ExcelArgument(Name = "input", Description = "String to return with special characters escaped")]
            string input)
        {
            return Regex.Escape(input);
        }


        [ExcelFunction(Description = "Find the pattern in the input, return TRUE if matched, FALSE otherwise")]
        
        public static object IsRegexMatch(
            [ExcelArgument( Name = "input",     Description = "Input text"  )]
            String input,
            [ExcelArgument( Name = "pattern",   Description = "Regular expression pattern to search for"    )]
            String pattern,
            [ExcelArgument( Name = "options",   Description = "Regex option flags, sum for each active option: IgnoreCase = 1, Multiline = 2, ExplicitCapture = 4, Compiled = 8, Singleline = 16, IgnorePatternWhitespace = 32, RightToLeft = 64, ECMAScript = 256, CultureInvariant = 512" ) ]
            int options
         )
        {
            if (String.IsNullOrEmpty(input) || String.IsNullOrEmpty(pattern))
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorValue;
            }
            else
            {
                RegexOptions ro = (RegexOptions)options;
                // Changed from instantiating a Regex and Match objects to calling the static function, because caching
                // Patterns are cached this way; better performance and less chance of memory leaks
                // Regex rx = new Regex(pattern, ro);
                // return rx.IsMatch(input);
                return Regex.IsMatch(input, pattern, ro);
            }
        }


        [ExcelFunction(Description = "Search for input for the pattern, return a comma delimited list of names or numbers of matching capture groups")]

        public static object RegexMatchGroups(
            [ExcelArgument( Name = "input",     Description = "Input text"  )]
            String input,
            [ExcelArgument( Name = "pattern",   Description = "Regular expression pattern to search for"    )]
            String pattern,
            [ExcelArgument( Name = "options",   Description = "Regex option flags, sum for each active option: IgnoreCase = 1, Multiline = 2, ExplicitCapture = 4, Compiled = 8, Singleline = 16, IgnorePatternWhitespace = 32, RightToLeft = 64, ECMAScript = 256, CultureInvariant = 512" ) ]
            int options,
            [ExcelArgument( Name = "MaxMatches", Description = "Maximum number of matches to execute on the input (omit or 0 to return all matches)")]
            int MaxMatches,
            [ExcelArgument( Name = "MaxGroups", Description = "Maximum number of group names or numbers to return for each match (omit or 0 for all groups)")]
            int MaxGroups
        )
        {
            if (String.IsNullOrEmpty(input) || String.IsNullOrEmpty(pattern))
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorValue;
            }
            else
            {
                RegexOptions ro = (RegexOptions)options;
                // For each match requested, walk the list of groups, concatenating the names/numbers of successful captures
                // If the pattern has no capture groups, we return a "0" group number if successful.
                // If the pattern does have capture groups we skip group 0 in each collection, as it represents the whole regex
                // The groups are not in strict order of appearance in the pattern, first come all the numbered (unnamed) groups, then all the named ones,
                // see https://docs.microsoft.com/en-us/dotnet/standard/base-types/grouping-constructs-in-regular-expressions?view=netframework-4.8#grouping-constructs-and-regular-expression-objects
                // so caveat emptor: don't assume the order of the groups indicates the order of appearance in the input string.
                string rs = ""; int i = 0; int j = 0; int m = MaxMatches;
                // Performance note: regex matches are executed lazily, so iterating on the matches collection 
                // has no penalty when requesting less matches than the total potential matches.
                foreach (Match rm in Regex.Matches(input, pattern, ro))
                {
                    int g = MaxGroups;
                    foreach (Group rg in rm.Groups)
                    {
                        if (i++ == 0)
                        {   // Group 0's succcess is the success of the whole match, return #NA if not
                            if (!g.Success) { return ExcelDna.Integration.ExcelError.ExcelErrorNA; }
                        }
                        else
                        {
                            if (g.Success)
                            {
                                rs += ',' + g.Name;
                                if (--MaxGroups == 0) { break; }
                            }
                        }
                    }
                }
                return rs.Substring(1);
            }
        }


    }

}
