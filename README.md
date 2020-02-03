# ExcelMazRegex
Excel regular expression add-in using .NET regex engine and ExcelDNA integration

 ## Version History:
- 1.2 MAZ 2020-02-02
  - Added `RegexMatches()`
  - Fixed `RegexMatch()` return value when not matched (did not return #NA correctly)
- 1.1 MAZ 2020-02-02
  - Added `IsRegexMatch()`, `RegexEscape()`, `RegexMatchGroups()`
  - Few minor fixes 
- 1.0 MAZ 2020-01-29
   - Released with `RegexMatch()` and `RegexReplace()`
   - Implemented Excel Intellisense functionality using decorations
   - ToDo: 
     - Add thread-safe ExcelDNA registration. This assumes we *never* have a Regex.Match object jump threads.
   - Consider for later: 
     - V1 function calls from inside Excel are now pretty fast, but consider memoization for the future

## Overview
ExcelMAZRegez is a simple, fast .NET regular expression library for Excel. 
As of v1 it's only for Excel formulas inside a worksheet; a later version might implement it for use inside VBA.

It's several orders of magnitude faster than using the VBA scripting library, and it's less convoluted than the few other Excel .NET regex libraries I found out there.

See the [.Net regular expression documentation](https://docs.microsoft.com/en-us/dotnet/standard/base-types/regular-expressions) for a full description of regexes, 
including search and replacement patterns, and how options work.


The formula use and syntax show up in the Excel Intellisense UI, and below for reference:

## Installation:
- Download the `\*-packed.xll` and `ExcelDna.IntelliSense.dll` files from the project repository [Releases page](https://github.com/mazamorac/ExcelMazRegex/releases) 
- Copy the appropriate `\*.xll` add-in file (32 or 64 bit) to your Excel add-ins folder (`%APPDATA%\Microsoft\AddIns` on Windows).
- Copy the `ExcelDna.IntelliSense.dll` file into the add-ins folder.
- Turn on the add-in in the Excel manage addins dialog (Alt+t,i).

## Documentation:
### Function RegexEscape()
`RegexEscape( text )`

Return the input text with special characters escaped. Useful to construct regex patterns with arbitrary text that will not be interpreted for special pattern interpretations, such as when a string includes  "\[", "$", etc.

### Function RegexMatch()
`RegexMatch( input, pattern [, options [, replacement ] ] )`

Finds and returns the text of the first instance of the regular expression pattern inside the input string, optionally modified with the option flags, and optionally with a replacement pattern.

#### Parameters
The `options` are bit flags (see below), and sould be added up to specify more than one optoin. E.g.: ignore case plus multilines is 3 (1 + 2).
* 1 = IgnoreCase
* 2 = Multiline
* 4 = ExplicitCapture
* 8 = Compiled
* 16 = Singleline
* 32 = IgnorePatternWhitespace
* 64 = RightToLeft
* 256 = ECMAScript
* 512 = CultureInvariant

if not specified, the replacement patterns defaults to "$0".

#### Returns:
* First instance of text in input that conforms to the input pattern and options, optionally modified by a replacement pattern
* #VALUE error if the input or pattern are empty strings
* #NA error if the pattern is not found

### Function RegexMatches()
`RegexMatches( input, pattern [, options [, replacement ] ] )`

Finds all the occurrences of the pattern in the input. Returns delimiter-separated list of matches with optional replacement pattern.

#### Parameters
- Same as `RegexMatch()`, plus...
- `delimiter`: Delimiter for the list of results, default ','

#### Returns:
- String with delimiter-separated list of matches found, optionally modified by replacement pattern.

### Function IsRegexMatch()
`IsRegexMatch( input, pattern [, options ] )`

#### Parameters
- Same as `RegexMatch()`, except for `replacement`, which is not used.

#### Returns:
TRUE if the pattern is found in the input, FALSE otherwise. 

### Function RegexMatchGroups()
`RegexMatchGroups( input, pattern [, options [, MaxMatches [, MaxGroups [, IncludeDuplicates ] ] ] ] )`

Search the input for matches of the pattern, return a comma delimited list of matching capture group names/numbers.

Useful to find out what chunks of a regular expression were matched against, without actually caring what the text that matched was. I personally use it a lot to label data, see the examples below.

#### Parameters:
- The `input`, `pattern`, and `options` parameters are the same as for `RegexMatch()`.
- `MaxMatches`: Maximum number of matches to execute on the input (omit or 0 to return all matches)
- `MaxGroups`: Maximum number of group names or numbers to return for each match (omit or 0 for all groups)
- `IncludeDuplicates`: Default TRUE: Print group names every time they're found in a match. FALSE: Only return the first instance of each capture group.

#### Notes:
- If the pattern has no capture groups, we return group number "0" for each qualifying match. 
- The groups within each match don't show in the same order as in the pattern, first come all the numbered (unnamed) groups, then all the named ones, so caveat emptor: don't assume the order of the groups inside each match is in pattern appearance order, nor input match index order. See https://docs.microsoft.com/en-us/dotnet/standard/base-types/grouping-constructs-in-regular-expressions?view=netframework-4.8#grouping-constructs-and-regular-expression-objects


#### Examples:
The formula:
- `=RegexMatchGroups('liliac,red,mauve,green','<?<primary>red|green|blue)(?<artsy>mauve|lilac|haze)'`)
- returns: `artsy,primary,artsy,primary`

To only return the first match, you'd use:
- `=RegexMatchGroups('mauve,red,green','<?<primary>red|green|blue)(?<artsy>mauve|lilac|haze)'`,,1)
- returns: `artsy`

To return all matches but not repeat group names, you set `IncludeDuplicates=FALSE`:
- `=RegexMatchGroups('mauve,red,green','<?<primary>red|green|blue)(?<artsy>mauve|lilac|haze)'`,,,,FALSE)
- returns: `artsy,primary`

The `MaxGroups` sets the max number of groups _per match_. Handy, for example, when different subexpressions may match the same text, and you only care for the first group that does. For example, see the difference between not using `MaxGroups`:
- `=RegexMatchGroups('mauve,red,green','(?<funky>green|lilac)<?<primary>red|green|blue)(?<artsy>mauve|lilac|haze)'`)
- returns: `artsy,primary,funky,primary`

... and setting `MaxGroups=1`: 
- `=RegexMatchGroups('mauve,red,green','(?<funky>green|lilac)<?<primary>red|green|blue)(?<artsy>mauve|lilac|haze)'`,,,1)
- returns: `artsy,primary,funky`

### Function RegexReplace()
`RegexReplace( input, pattern [, options [, replacement ] ] )`

Finds all the instances of the search pattern in the input text, optionally modified with the option flags, replaces them with the replacement pattern, and returns the modifed input. Similar to `RegexMatch()` but searching and replacing the entire input, and the `replacement` pattern default is an empty string ("").

The parameters, and options for RegexReplace() are the same as for RegexMatch(), except for the `replacement` default. 

#### Returns:
* The input text with every instance of the search pattern + options replaced by the replacement pattern
* #VALUE error if the input or pattern are empty strings

## General notes:
* If you ever get a #NUM error, it's an internal error and shouldn't happen. Please raise an issue if it does, with replicable example(s).
