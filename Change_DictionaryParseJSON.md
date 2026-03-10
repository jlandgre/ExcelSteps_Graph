## Purpose
Add `Dictionary.ParseStringToDictProcedure` (`Dictionary.cls`) as a new public class method that parses a JSON-like string into an already-instanced `Dictionary` object (empty or pre-populated).

## Background
The `Dictionary` class in ExcelSteps provides a cross-platform dictionary. We want to add a procedure to populate a new Dictionary or add key/value pairs to an existing one by parsing a JSON-like string like these examples:
```
{"key1":"value1", "key2":"value2"}   value1 and value2 are Strings
{"key1":True, "key2":False}          values are Boolean
{"key1":23.1, "key2":100}           values are numeric
{"name":"Alice", "age":25, "active":True}  mixed types
```

Reference documents:
- [[VBA Project Architecture]] - Standard VBA project structure
- [[VBA Project Changes]] - Planning mode process
- Dictionary.cls source code: `Excel_Steps/src/Dictionary.cls`
- Existing tests: `Excel_Steps/tests/tests_Dictionary.bas`

## Data I/O Descriptions

**Input:**
- `dict` (Object) - Dictionary instance (can be empty or pre-populated)
- `jsonStr` (String) - JSON-like string with format: `{"key":"value", "key2":value2}`

**String Format Requirements:**
- Outer braces: `{` and `}`
- Key-value pairs: `"key":value` or `key:value`
- Separator: comma between pairs
- String values: quoted `"value"`
- Numeric values: unquoted numbers (Integer/Long/Double)
- Boolean values: unquoted `True` or `False` (case-insensitive)
- Whitespace: tolerant of spaces around delimiters

**Output:**
- Dictionary populated with parsed key-value pairs
- Returns Boolean (True if successful, False on error)
- Existing dictionary entries preserved unless keys conflict (overwrite behavior)

**Data Type Mapping:**
- Quoted strings → String
- Numeric literals → Double
- `True`/`False` → Boolean
- Empty value → Empty

## Project Architecture

**Module Structure:**
- **Dictionary.cls** - Add one new public entry-point function plus private helper methods in the same class
- No new classes required
- No changes to Validation.bas (New_Dictionary already exists)

**Procedure Pattern:**
`Dictionary.cls` exposes one public entry-point procedure function: `dict.ParseStringToDictProcedure(jsonStr)`.
The procedure function calls methods in sequence (validate, split, parse, type-detect, add), and requires `dict` to be initialized before the call (empty or pre-populated).

## Test Architecture

**Test Module:** `tests_Dictionary.bas` (existing module)
**Procedure Group:** `procs.dictionary` (existing Procedure attribute in Procedures.cls)

**New Tests to Add:**
- `test_ParseStringToDict` - Main functionality test
- Additional test functions as needed for edge cases

## Discussion: JSON Syntax Flexibility

**Design Decision: Quote Requirements**
- **Keys**: Support both quoted (`"key"`) and unquoted (`key`) for flexibility
- **String values**: Require quotes to distinguish from Boolean/Numeric
- **Boolean/Numeric**: Unquoted (standard JSON behavior)

**Rationale:** Simplifies user experience while maintaining type safety. Unquoted keys are easier to type; quoted string values prevent ambiguity.

**String Parsing Strategy:**
1. Validate outer braces
2. Strip braces and split by comma (handling quoted strings)
3. For each pair: split by colon
4. Trim whitespace from keys/values
5. Type detection: detect quotes (String), True/False (Boolean), numeric pattern (Double), else error

## Testing Considerations

**Test Module Structure:** Add tests to existing `tests_Dictionary.bas` module under `procs.dictionary` procedure group

**Test Coverage Scope:** Unit tests of individual methods and integration test of complete method:
- Parse simple string pairs (single key-value)
- Parse multiple pairs with mixed types
- Parse with various whitespace patterns
- Parse with quoted and unquoted keys
- Parse into initialized empty vs initialized pre-existing dictionary
- Error cases: malformed JSON, unmatched braces, invalid syntax

**Impact on Existing Tests:** None - new tests added only

**Test Data Requirements:** No external test files needed; all test strings defined within test code

**Cross-workbook Instantiation:** ExcelSteps.New_Dictionary already exists in Validation.bas

**Edge Cases and Validation:**
- Empty string input
- Malformed JSON (missing braces, unmatched quotes)
- Invalid types (unquoted strings that aren't True/False)
- Comma within quoted string values
- Escaped characters in strings (if supported)
- Duplicate keys (test overwrite behavior)
- Very long strings (performance)

## Procedure Outline

**ParseStringToDictProcedure** - Public Dictionary class method to parse JSON-like string and populate Dictionary
* **`ParseStringToDictProcedure`**: [[procPlan_ParseStringToDictProcedure]]
* `ValidateAndStripBraces` - Validate outer braces and remove them; return inner string
* `SplitIntoPairs` - Split string into array of key:value pair strings (handle commas in quoted strings)
* `ParsePair` - Parse single "key:value" string into key and value; detect value type
* `DetectValueType` - Determine if value is String, Boolean, Double or invalid
* `AddParsedValue` - Add parsed key-value to dictionary using dict.Add with IsObject check
