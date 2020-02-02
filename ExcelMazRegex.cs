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


        [ExcelFunction(Description = "Search the input for matches of the pattern, return a comma delimited list of matching capture group names/numbers")]

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
            int MaxGroups,
            [ExcelArgument( Name = "IncludeDuplicates", Description = "Default TRUE: print group names every time they're found in a match. FALSE: Only return the first instance of each capture group")]
            bool IncludeDuplicates = true
        )
        {
            if (String.IsNullOrEmpty(input) || String.IsNullOrEmpty(pattern))
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorValue;
            }
            else
            {
                RegexOptions ro = (RegexOptions)options;
                HashSet<String> seengroup = new HashSet<String>();
                // Walk the matches, for each match, walk the capture groups, concatenating the names/numbers of successful captures, checking for max count or duplication limits
                // If the pattern has no capture groups, we return group number "0" for each qualifying match.
                // The groups within each match don't show in the same order as in the pattern, first come all the numbered (unnamed) groups, then all the named ones,
                // see https://docs.microsoft.com/en-us/dotnet/standard/base-types/grouping-constructs-in-regular-expressions?view=netframework-4.8#grouping-constructs-and-regular-expression-objects
                // so caveat emptor: don't assume the order of the groups inside each match is in pattern appearance order, nor input match index order.

                string matchlist = ""; bool isfirst = true; int gnum; int gfound;
                // Performance note: regex matches are executed lazily (it's an iterator, not prepopulated), 
                // so walking the matches collection incurrs no penalty when stopping before the end
                foreach (Match rm in Regex.Matches(input, pattern, ro))
                {
                    gnum = 0; gfound = MaxGroups;
                    foreach (Group rg in rm.Groups)
                    {
                        if (gnum++ == 0)     // First group of every match, check the first one, skip the rest
                        {
                            if (isfirst)      // First group of all matches: if no match return #NA
                            {
                                if (!rg.Success) return ExcelDna.Integration.ExcelError.ExcelErrorNA;
                                isfirst = false;
                            }
                        }
                        else
                            if (rg.Success && ( IncludeDuplicates || ! seengroup.Contains(rg.Name) ) )
                            {
                                matchlist += "," + rg.Name;
                                if (--gfound == 0) break;
                                if( !IncludeDuplicates ) seengroup.Add(rg.Name);   // avoid lookup maintenance overhead if not required
                            }
                    }
                    // If i==1 here then the match was successful, but no capturing groups above zero exist, 
                    // so we add group "0" (which we skipped, see comments above)
                    if (gnum == 1) matchlist += ",0";         
                    if (--MaxMatches == 0) break;
                }
                return matchlist.Substring(1);
            }
        }


    }

}
