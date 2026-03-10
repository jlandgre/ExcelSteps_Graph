## Procedure Purpose
Parse a JSON-like string and populate a Dictionary with key-value pairs. Supports String, Boolean, and Double values with flexible whitespace.

All methods in this procedure are Public. `ParseStringToDictProcedure` is the procedure method, and the remaining methods are called in sequence by it.

## Procedure Detailed Requirements

**Overall Action:**
Parse `{"key1":"value1", "key2":value2, "key3":True}`-style input and add each pair via `.Add`. The procedure tolerates whitespace and supports quoted or unquoted keys.

**Input Requirements:**
- `dict` (Dictionary object): Pre-initialized via `Set dict = ExcelSteps.New_Dictionary`; may be empty or pre-populated
- `jsonStr` (String): JSON-like string with outer braces `{}`, comma-separated `key:value` pairs

**String Format Requirements:**
- Outer braces: `{` and `}`
- Key-value pairs: `"key":value` or `key:value`
- Separator: comma between pairs
- String values: quoted `"value"`
- Numeric values: unquoted numbers (Integer/Long/Double)
- Boolean values: unquoted `True` or `False` (case-insensitive)
- Whitespace: tolerant of spaces around delimiters

**Output:**
- Returns `Boolean`: `True` on success, `False` on validation/parsing errors
- Modifies `dict` by adding parsed key-value pairs via `.Add`
- Preserves existing entries; duplicate keys are updated per `.Add` behavior

**Type Handling:**
- String values: Enclosed in double quotes `"value"`
- Boolean values: Unquoted `True` or `False` (case-insensitive)
- Numeric values: Unquoted numeric literals parsed as Double
- Keys: Support both quoted `"key"` and unquoted `key` formats

**Data Type Mapping:**
- Quoted strings -> String
- Numeric literals -> Double
- `True`/`False` -> Boolean
- Empty value -> Empty

## Procedure Method/Sub-Procedure Descriptions

### ValidateAndStripBraces
- **Action**: Validate that jsonStr starts with `{` and ends with `}` after trimming; remove braces and return inner content
- **Inputs**: `jsonStr` (String) - raw input string with expected braces
- **Outputs**: Returns String (inner content without braces); sets function False and returns empty string if validation fails
- **Logic Steps**:
  1. Trim whitespace from jsonStr using `Trim()`
  2. Validate length > 2 via `errs.IsFail(Len(trimmedStr) < 3, 1)`
  3. Validate first char = `{` via `errs.IsFail(Left(trimmedStr, 1) <> "{", 2)`
  4. Validate last char = `}` via `errs.IsFail(Right(trimmedStr, 1) <> "}", 3)`
  5. Return `Mid(trimmedStr, 2, Len(trimmedStr) - 2)` (substring without first and last chars)
- **Validation/Error Conditions**: Check length >= 3; check first/last characters are braces

### SplitIntoPairs
- **Action**: Split inner string into array of "key:value" pair strings, handling commas within quoted values
- **Inputs**: `innerStr` (String) - content between outer braces; `pairsArray()` (String array) - ByRef output array
- **Outputs**: Returns Boolean; Sets pairsArray() to array of pair strings; returns count as Long ByRef argument
- **Logic Steps**:
  1. Initialize tracking variables: `inQuotes As Boolean`, `currentPair As String`, collection for pairs
  2. Loop through each character in innerStr using `Mid(innerStr, i, 1)`
  3. If char is `"` (double quote): toggle `inQuotes` flag
  4. If char is `,` and `Not inQuotes`: append currentPair to collection; reset currentPair
  5. Else: append char to currentPair
  6. After loop: append final currentPair to collection
  7. Size pairsArray and populate from collection
- **Validation/Error Conditions**: Check innerStr length > 0; return False if empty

### ParsePair
- **Action**: Parse single "key:value" pair string into separate key and typed value
- **Inputs**: `pairStr` (String) - single pair like "name":"Alice" or age:25; `key` (String) - ByRef output; `value` (Variant) - ByRef output
- **Outputs**: Returns Boolean; Sets key (String) and value (Variant with detected type)
- **Logic Steps**:
  1. Find colon position using `InStr(pairStr, ":")`
  2. Validate colon exists via `errs.IsFail(colonPos = 0, 1)`
  3. Extract key: `Left(pairStr, colonPos - 1)`; extract valueStr: `Mid(pairStr, colonPos + 1)`
  4. Trim and remove quotes from key if present: Check `Left(keyTrimmed, 1) = """` and `Right(keyTrimmed, 1) = """`
  5. Call DetectValueType to parse valueStr into typed value
  6. Return key and value via ByRef arguments
- **Validation/Error Conditions**: Validate colon exists in pair string; validate detectValueType succeeds

### DetectValueType
- **Action**: Analyze value string and return appropriately typed Variant (String, Boolean, or Double)
- **Inputs**: `valueStr` (String) - trimmed value portion after colon
- **Outputs**: Returns Boolean; `typedValue` (Variant) - ByRef output with detected type
- **Logic Steps**:
  1. Trim valueStr
  2. Check if quoted string: `Left(trimmed, 1) = """` and `Right(trimmed, 1) = """`
     - If yes: remove quotes and return as String
  3. Check if Boolean: `UCase(trimmed) = "TRUE"` or `UCase(trimmed) = "FALSE"`
     - If yes: return Boolean via `CBool(UCase(trimmed) = "TRUE")`
  4. Check if numeric: Use `IsNumeric(trimmed)`
     - If yes: return `CDbl(trimmed)`
  5. Else: invalid type, set error via `errs.IsFail(True, 1)` and return False
- **Validation/Error Conditions**: Return error if value doesn't match any recognized type

### AddParsedValue
- **Action**: Add parsed key-value pair to Dictionary using existing .Add method
- **Inputs**: `dict` (Dictionary object); `key` (String); `value` (Variant)
- **Outputs**: Returns Boolean (True if successful)
- **Logic Steps**:
  1. Call `dict.Add key, value`
  2. .Add method handles both new keys and updates to existing keys automatically
  3. Return True if no errors
- **Validation/Error Conditions**: Rely on .Add method's existing validation

## Testing Requirements

**Test Module Location:** tests_Dictionary.bas; Procedure group: procs.dictionary

**Test Coverage Scope:** Unit tests of individual methods and integration test of the full procedure method:
- Parse simple string pairs (single key-value)
- Parse multiple pairs with mixed types
- Parse with various whitespace patterns
- Parse with quoted and unquoted keys
- Parse into initialized empty vs initialized pre-existing dictionary
- Error cases: malformed JSON, unmatched braces, invalid syntax

**Impact on Existing Tests:** None - new tests added only

**Test Data Requirements:** No external test files needed; all test strings are defined in test code

**Cross-workbook Instantiation:** `ExcelSteps.New_Dictionary` already exists in `Validation.bas`

**Test Setup Pattern:**
```vb
Dim dict As Object
Set dict = ExcelSteps.New_Dictionary
```

**Success Criteria:**
- Parse simple single pair: `dict.ParseStringToDictProcedure("{""name"":""Alice""}")` -> `dict.Item("name") = "Alice"`
- Parse multiple pairs: `{"name":"Alice", "age":25}` → two items with correct values
- Parse mixed types: String, Boolean, Double all correctly typed
- Parse with whitespace: `{ "key1" : "val1" , "key2" : 123 }` succeeds
- Parse into existing dict: existing keys preserved, new keys added
- Parse with unquoted keys: `{name:"Alice"}` succeeds

**Edge Cases to Test:**
- Empty string: `""` returns False
- Empty braces: `"{}"` succeeds with no items added
- Missing braces: `"key:value"` returns False
- Unmatched braces: `"{key:value"` returns False
- Invalid value type: `{key:invalid}` returns False (not quoted, not Boolean, not numeric)
- Comma in quoted string: `{"name":"Last, First"}` parses correctly
- Duplicate keys: Second occurrence updates value
- Boolean values: `{active:True}`, `{active:true}`, `{active:FALSE}` all work
- Numeric values: Integer, decimal, negative numbers
- Special characters in string values
- Very long strings (performance check)

**Shared Test Setup Subroutines:**
If multiple tests share setup pattern, create a shared setup subroutine:
```vb
Sub SetupParseAndValidate(tst As Test, dict As Object, jsonStr As String, expectedSize As Long)
    tst.Assert tst, dict.ParseStringToDictProcedure(jsonStr)
    tst.Assert tst, dict.Size = expectedSize
End Sub
```
