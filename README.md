# ExcelMazRegex
 Excel regular expressions add-in using ExcelDNA 1.0

 ## Version History:
 - 1.0 MAZ 2020-01-29
   - Released with RegexMatch and RegexReplace
   - Implemented ExcelIntellisense functionality with decorations
   - ToDo: 
     - Add thread-safe ExcelDNA registration. This assumes we *never* have a Regex.Match object jump threads.
   - Consider for later: 
     - V1 function calls from inside Excel are pretty fast, but consider memoization for the future

## Documentation:
### Summary
ExcelMAZRegez is a simple, fast .NET regular expression library for Excel. 
As of v1 it's only for Excel formulas inside a worksheet; a later version might implement it for use inside VBA.

It's several orders of magnitude faster than using the VBA scripting library, and it's less convoluted than the few other Excel .NET regex libraries I found out there.

See the [.Net regular expression documentation](https://docs.microsoft.com/en-us/dotnet/standard/base-types/regular-expressions) for a full description of regexes, 
including search and replacement patterns, and how options work.


The formula use and syntax show up in the Excel Intellisense UI, and below for reference:

#### Function RegexMatch()
`RegexMatch( input, pattern [, options [, replacement ] ] )`

Finds and returns the text of the first instance of the regular expression pattern inside the input string, optionally modified with the option flags, and optionally with a replacement pattern.

The options are bit flags (see below), and sould be added up to specify more than one optoin. E.g.: ignore case plus multilines is 4 (1 + 3).
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

Returns:
* first instance of text in input that conforms to the input pattern and options, optionally modified by a replacement pattern
* #VALUE error if the input or pattern are empty strings
* #NA error if the pattern is not found
* #NUM error for internal errors (please raise an issue if this happens (it shouldn't))


#### Function RegexReplace()
`RegexReplace( input, pattern [, options [, replacement ] ] )`

Finds all the instances of the search pattern in the input text, optionally modified with the option flags, replaces them with the replacement pattern, and returns the modifed input.

The parameters, and options for RegexReplace() are the same as for RegexReplace().

Returns:
* all the input text, with every instance of the search pattern + options replaced by the replacement pattern
* #VALUE error if the input or pattern are empty strings
* #NUM error for internal errors (please raise an issue if this happens (it shouldn't))
