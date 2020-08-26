using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using System.Text.RegularExpressions;
using System.Reflection;
using System.Runtime.Remoting.Messaging;

// ToDo: implement versioning from code instead of project properties
// [assembly: AssemblyVersion("1.2.0.1")]

namespace ExcelMazRegex
{
    public static class ExcelMazRegex
    {

        // For RegexGroupMatches we need to remember the order in which we added items to the set
        public class OrderedHashSet<T> : KeyedCollection<T, T>
        {
            protected override T GetKeyForItem(T item)
            {
                return item;
            }
        }

        [ExcelFunction(Description = "Returns the version number.", IsThreadSafe = true)]

        public static string RegexVersionNumber()
        {
            return "1.3.1";
        }


        [ExcelFunction(Description = "Find the pattern in the input, return the first matching string." , IsThreadSafe = true ) ]

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
            if ( input == null || String.IsNullOrEmpty(pattern) )
                return ExcelDna.Integration.ExcelError.ExcelErrorValue;
            else
            {
                //if (String.IsNullOrEmpty(input)) input = "";
                RegexOptions ro = (RegexOptions)options;
                Match rm = Regex.Match(input, pattern, ro);
                if (rm.Success)
                    if (string.IsNullOrEmpty(replacement))
                        return rm.Value;
                    else
                        return Regex.Match(input, pattern, ro).Result(replacement);
                else 
                    return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }



        [ExcelFunction(Description = "Find the pattern in the input, substitute all matches with the replacement pattern" , IsThreadSafe = true ) ]

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
            if ( input==null || String.IsNullOrEmpty(pattern) || replacement == null )
                return ExcelDna.Integration.ExcelError.ExcelErrorValue;
            else
            {
                RegexOptions ro = (RegexOptions)options;
                // Changed to static function call because caching
                return Regex.Replace(input, pattern, replacement, ro);
            }
        }


        [ExcelFunction(Description = "Return string with special characters escaped to remove special meaning", IsThreadSafe = true ) ]

        public static string RegexEscape(
            [ExcelArgument(Name = "input", Description = "String to return with special characters escaped")]
            string input)
        {
            return Regex.Escape(input);
        }


        [ExcelFunction(Description = "Find the pattern in the input, return TRUE if matched, FALSE otherwise" , IsThreadSafe = true ) ]

        public static object IsRegexMatch(
            [ExcelArgument( Name = "input",     Description = "Input text"  )]
            String input,
            [ExcelArgument( Name = "pattern",   Description = "Regular expression pattern to search for"    )]
            String pattern,
            [ExcelArgument( Name = "options",   Description = "Regex option flags, sum for each active option: IgnoreCase = 1, Multiline = 2, ExplicitCapture = 4, Compiled = 8, Singleline = 16, IgnorePatternWhitespace = 32, RightToLeft = 64, ECMAScript = 256, CultureInvariant = 512" ) ]
            int options
         )
        {
            if ( input==null || String.IsNullOrEmpty(pattern) )
                return ExcelDna.Integration.ExcelError.ExcelErrorValue;
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


        [ExcelFunction(Description = "Search the input for matches of the pattern, return a comma delimited list of matching capture group names/numbers in match order" , IsThreadSafe = true ) ]

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
            object MaxGroups,
            [ExcelArgument( Name = "IncludeDuplicates", Description = "Default TRUE: Print group names every time they're found in a match. FALSE: Only return the first instance of each capture group")]
            object IncludeDuplicates
        )
        {
            if ( input == null || String.IsNullOrEmpty(pattern) )
                return ExcelDna.Integration.ExcelError.ExcelErrorValue;
            else
            {
                RegexOptions ro = (RegexOptions)options;
                bool incdups = (
                    IncludeDuplicates is ExcelMissing
                        ? true
                        : ( IncludeDuplicates is bool
                            ? (bool) IncludeDuplicates 
                            : ( IncludeDuplicates is string 
                                ? String.IsNullOrEmpty((string)IncludeDuplicates) 
                                : ( (IncludeDuplicates is int ) || (IncludeDuplicates is double ) 
                                    ? ( (double) IncludeDuplicates ) > 0
                                    : false
                ) ) ) );
                int maxgrp = 0;
                if ((MaxGroups is int) || (MaxGroups is double))
                {
                    maxgrp = (int)(double)MaxGroups;
                }
                else
                {
                    if (MaxGroups is string)
                    {

                        if (String.IsNullOrEmpty((string)MaxGroups))
                        {
                            maxgrp = 0;
                        }
                        else
                        {
                            if (!int.TryParse((string)MaxGroups, out maxgrp)) return ExcelDna.Integration.ExcelError.ExcelErrorValue;
                        }
                    }
                    else
                    {
                        if (MaxGroups is bool & (bool)MaxGroups)
                        {
                            maxgrp = 1;
                        }
                    }
                }

                HashSet<String> seengroup = new HashSet<String>();
                // Walk the matches, for each match, walk the capture groups, concatenating the names/numbers of successful captures, checking for max count or duplication limits
                // If the pattern has no capture groups, we return group number "0" for each qualifying match.
                // The groups within each match don't show in the same order as in the pattern, first come all the numbered (unnamed) groups, then all the named ones,
                // see https://docs.microsoft.com/en-us/dotnet/standard/base-types/grouping-constructs-in-regular-expressions?view=netframework-4.8#grouping-constructs-and-regular-expression-objects
                // so caveat emptor: don't assume the order of the groups inside each match is in pattern appearance order, nor input match index order.
                // this behavior can be controlled if using regex option flag ExplicitCapture (4), which ignores non-named capture groups.

                string matchlist = ""; bool isfirst = true; int gnum; int gfound;
                // Performance note: regex matches are executed lazily (it's an iterator, not prepopulated), 
                // so walking the matches collection incurrs no penalty when stopping before the end
                foreach (Match rm in Regex.Matches(input, pattern, ro))
                {
                    gnum = 0; gfound = maxgrp;
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
                            if (rg.Success && (incdups || ! seengroup.Contains(rg.Name) ) )
                            {
                                matchlist += "," + rg.Name;
                                if (--gfound == 0) break;
                                if( ! incdups) seengroup.Add(rg.Name);   // avoid lookup maintenance overhead if not required
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


        [ExcelFunction(Description = "Finds all the occurrences of the pattern in the input. Returns delimiter-separated list of matches with optional replacement pattern." , IsThreadSafe = true ) ]

        public static object RegexMatches(
            [ExcelArgument( Name = "input",     Description = "Input text"  )]
            String input,
            [ExcelArgument( Name = "pattern",   Description = "Regular expression pattern to search for"    )]
            String pattern,
            [ExcelArgument( Name = "options",   Description = "Regex option flags, sum for each active option: IgnoreCase = 1, Multiline = 2, ExplicitCapture = 4, Compiled = 8, Singleline = 16, IgnorePatternWhitespace = 32, RightToLeft = 64, ECMAScript = 256, CultureInvariant = 512" ) ]
            int options,
            [ExcelArgument( Name = "replacement", Description = "Optional replacement pattern for result")]
            String replacement,
            [ExcelArgument( Name = "delimiter", Description = "Delimiter for the list of results, default ','")]
            object delimiter,
            [ExcelArgument( Name = "MaxMatches", Description = "Maximum number of matches to execute on the input (omit or 0 to return all matches)")]
            int MaxMatches,
            [ExcelArgument( Name = "IncludeDuplicates", Description = "Default TRUE: return every match. FALSE: Only return the first instance of each match")]
            object IncludeDuplicates

        )
        {
            // the replacement string is not checked for an empty string or null because empty is a valid replacement pattern and null when ommited (it's optional)
            if ( input == null || String.IsNullOrEmpty(pattern) )
                return ExcelDna.Integration.ExcelError.ExcelErrorValue;
            else
            {
                RegexOptions ro = (RegexOptions)options;
                bool incdups = (
                    IncludeDuplicates is ExcelMissing
                        ? true
                        : (IncludeDuplicates is bool
                            ? (bool)IncludeDuplicates
                            : (IncludeDuplicates is string
                                ? String.IsNullOrEmpty((string)IncludeDuplicates)
                                : ((IncludeDuplicates is int) || (IncludeDuplicates is double)
                                    ? ((double)IncludeDuplicates) > 0
                                    : false
                ))));
                HashSet<String> seengroup = new HashSet<String>();
                string delim = (delimiter is ExcelMissing ? "," : (string)delimiter);
                string rs = "";
                bool hasReplacement = ! string.IsNullOrEmpty(replacement);
                foreach (Match rm in Regex.Matches(input, pattern, ro))
                {
                    if (rm.Success && (incdups || !seengroup.Contains(rm.Value)))
                    {
                        rs += delim + ( hasReplacement ? rm.Result(replacement) : rm.Value );
                        if (!incdups) seengroup.Add(rm.Value);
                        if (--MaxMatches == 0) break;
                    }
                }
                
                if (rs == "")
                    return ExcelDna.Integration.ExcelError.ExcelErrorNA;
                else
                    return rs.Substring( delim.Length );
            }
        }

        [ExcelFunction(Description = "Search the input for matches of the pattern, return a comma delimited list of matching capture group names/numbers in capture group order within the pattern" , IsThreadSafe = true ) ]

        public static object RegexGroupMatches(
            [ExcelArgument( Name = "input",     Description = "Input text"  )]
            String input,
            [ExcelArgument( Name = "pattern",   Description = "Regular expression pattern to search for"    )]
            String pattern,
            [ExcelArgument( Name = "options",   Description = "Regex option flags, sum for each active option: IgnoreCase = 1, Multiline = 2, ExplicitCapture = 4, Compiled = 8, Singleline = 16, IgnorePatternWhitespace = 32, RightToLeft = 64, ECMAScript = 256, CultureInvariant = 512" ) ]
            int options,
            [ExcelArgument( Name = "MaxMatches", Description = "Maximum number of matches to execute on the input (omit or 0 to return all matches)")]
            int MaxMatches,
            [ExcelArgument( Name = "MaxGroups", Description = "Maximum total number of group names or numbers to return (omit or 0 for all groups)")]
            int MaxGroups,
            [ExcelArgument( Name = "IncludeDuplicates", Description = "Default TRUE: Print group names every time they're found in a match. FALSE: Only return the first instance of each capture group")]
            object IncludeDuplicates,
            [ExcelArgument( Name = "GroupNamesTransformPattern",  Description = "Transform group names on output list, regex pattern for search"    )]
            String GroupNamesTransformPattern,
            [ExcelArgument( Name = "GroupNamesTransformReplacement",  Description = "Transform group names on output list, regex replacement pattern"    )]
            String GroupNamesTransformReplacement
        )
        {
            if ( input == null || String.IsNullOrEmpty(pattern) )
                return ExcelDna.Integration.ExcelError.ExcelErrorValue;
            else
            {
                // Walk the matches, for each match, walk the capture groups, concatenating the names/numbers of successful captures, checking for max count or duplication limits
                // If the pattern has no capture groups, we return group number "0" for each qualifying match.
                // The groups within each match don't show in the same order as in the pattern, first come all the numbered (unnamed) groups, then all the named ones,
                // see https://docs.microsoft.com/en-us/dotnet/standard/base-types/grouping-constructs-in-regular-expressions?view=netframework-4.8#grouping-constructs-and-regular-expression-objects
                // so caveat emptor: don't assume the order of the groups inside each match is in pattern appearance order, nor input match index order.
                // this behavior can be controlled if using regex option flag ExplicitCapture (4), which ignores non-named capture groups.

                RegexOptions ro = (RegexOptions)options;
                bool incdups = (
                    IncludeDuplicates is ExcelMissing
                        ? true
                        : (IncludeDuplicates is bool
                            ? (bool)IncludeDuplicates
                            : (IncludeDuplicates is string
                                ? String.IsNullOrEmpty((string)IncludeDuplicates)
                                : ((IncludeDuplicates is int) || (IncludeDuplicates is double)
                                    ? ((double)IncludeDuplicates) > 0
                                    : false
                ))));
                string matchlist = ""; string[] gnames; int[] gmatched;
                int gcount = 0, gnum = 0;
                
                
                // If we get a match, retrieve the capture group names and initialize data structures, else return #NA
                MatchCollection rmc = Regex.Matches(input, pattern, ro);
                Match firstmatch = rmc[0];
                if (! firstmatch.Success) return ExcelDna.Integration.ExcelError.ExcelErrorNA;
                
                GroupCollection rgc = firstmatch.Groups;
                gcount = rgc.Count;
                gmatched = new    int[gcount];  // For capture group match counts
                gnames   = new string[gcount];  // For capture group names
                foreach (Group rg in rgc)
                {
                    gnames[gnum++] = rg.Name;
                }
                
                // Now walk a maximum of of MaxMatches, recording successfully matched capture groups
                foreach (Match rm in rmc)
                {
                    gnum = 0; 
                    foreach (Group rg in rm.Groups)
                    {
                        if (rg.Success) ++gmatched[gnum];
                        ++gnum;
                    }
                    if (--MaxMatches==0) break;
                }

                // Now construct output list, with MaxGroups names, honoring IncludeDuplicates and GroupNamesTransformPattern
                HashSet<String> seengroup = new HashSet<String>(); string gname;
                bool stripnames; string[] gstripped = new string[gcount];

                if (stripnames = ! String.IsNullOrEmpty(GroupNamesTransformPattern))
                {
                    if (String.IsNullOrEmpty(GroupNamesTransformReplacement)) GroupNamesTransformReplacement = "";
                    for (int i = 0; i < gcount; i++)
                        gstripped[i] = Regex.Replace(gnames[i], GroupNamesTransformPattern, GroupNamesTransformReplacement);
                }


                for (int i = ( gcount == 1 ? 0 : 1); i < gcount; i++)           // If there are defined capture groups, skip group 0 (corresponding to the full pattern)
                {
                    if (gmatched[i] > 0)
                    {
                        gname = ( stripnames ? gstripped[i] : gnames[i] );
                        if (incdups || ! seengroup.Contains(gname))
                        {
                            matchlist += "," + gname;
                            if (--MaxGroups == 0) break;
                            if ( ! incdups ) seengroup.Add(gname);
                        }
                    }
                } 

                // debugging output. ToDo: Remove later
                // return $"{gcount}|" + String.Join(",", (stripnames ? gstripped : gnames)) + "|" + matchlist;
                return matchlist.Substring(1);
            }
        }

    }

}
