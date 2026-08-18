# Pure VB6 Crypto

## General guidelines

- All VB6 source files are ASCII (cp1251) with CRLF line endings
- Use only batch or PowerShell scripts and remember to `cd FullPath` before executing any of these

## Coding Style for this project

- Declare all helper procedures `Private`
- Never pass `Source` to `Err.Raise` i.e. `Err.Raise ERR_NUMBER, , ERR_DESCRIPTION`
- Put `On Error GoTo 0` before a manual `Err.Raise` so the local handler does not trap/log it
- One blank line between procedures
- One blank line between the `Dim` block and the first executable line inside a procedure; no other blank lines inside a procedure body
- Use hungarian notation: `s` - String, `l` - Long/LongPtr, `n` - Integer, `b` - Boolean, `o` - Object, `c` - Collection, `d` - Date, `dbl` - Double, `sng` - Single, `byt` - Byte, `u` - UDTs, `h` - Handles (incl. hResult), `cy` - Currency, `e` - Enums, `clr` - OLE_COLOR
- Use `p` for interface "pointers" i.e. anything declared `As IVBXxxx` incl. weak references
- Use `ba` prefix for byte arrays
- Use `a` for arrays without element type prefix i.e. `m_aRows` instead of `m_auRows`
- Use `m_` prefix for member variables and `g_` for global ones
- Use `md` prefix for standard modules, `c` for classes, `frm` for forms, `ctx` for user-controls
- Use `lab` prefix for Labels, not `lbl`
- Use `Ucs` for enums/UDTs names, use `ucsXxx` for enum entries where Xxx is a unique prefix
- For UDT members stop using hungarian notation i.e. `uRow.Values(1)` not `aValues`
- Declare all variables at the beginning of the procedure, separated from the code by a blank line
- One local variable declaration per line, data-type aligned at column 25
- Order local variable declarations by how early each is used in the procedure
- Align API consts data-types at column 45
- Align module variables data-types at column 37, module consts at 41 (without STR_MODULE_NAME)
- Declare API consts local to a routine if not used in any other routine
- Group API consts by what they are under a `'---` divider naming the group, sorted by value, no blank lines between groups, group names on a single line; keep it that way whenever a const is added
- API declares use "dllname" without the .dll suffix, always Unicode versions (aliased to names without the W suffix)
- Use `LongPtr` with API declares i.e. for VB6 use enum hack
- Separate logical sections inside a procedure with `'---` comments, not blank lines
- Start comments with `'---` instead of a single `'`, except for a comment banner at the start of a procedure/module
- Put only one statement per logical line i.e. don't use : to separate multiple statements
- Put `If` statements on separate lines i.e. don't merge `If Cond Then Stmt` on a single line
- Use `Call` only for API functions whose result is discarded (not used in an `If`); never otherwise
- Use `QH` label for cleanup before `EH` label
- Align `Case`es after `Select Case` at the same column i.e. don't indent
- All `Long` const hex literals between &H8000 and &HFFFF must use the & type character (e.g. &H8000&) or risk being sign-extended
- Never use `Next Var` i.e. just `Next`
- Always omit `ByRef`; specify `ByVal` only when needed
- Order of procedures in module: public events/enums/types, API declares, member variables and private enums/types, properties, methods, event handlers, base class events (e.g. Class_Terminate)
- Within properties/methods order by visibility: public, friend, private
- Use `pv` prefix for private procedures and `fr` for friend ones
- Name procedures Verb+[Adj+]Noun: `Get` for cheap values, `Count`/`Calc`/`Build`/`Measure` where work happens, `Hit` for point lookups, `Handle` for message reactions, `Track`/`End` for drag pairs, `Is`/`Has`/`Needs`/`Uses` for predicates
- Use `lIdx`, `lJdx`, etc. instead of single-letter index variable `i`, `j`, etc.
- Prefer `Long` to `Integer` i.e. don't use `Integer` indexers
- ReDim uses explicit data-type i.e. `ReDim aName(0 To 100) As String`
- Write code which works with `Break on All Errors` setting in the IDE i.e. don't depend on `On Error Resume Next` for normal workflow
- Put `DefObj A-Z` on all modules so untyped vars/params do not remain of Variant type
- Put `Public enums`, `Public events`, `API`, `Constants and member variables`, `Error management`, `Properties`, `Methods`, `Functions`, `Control events`, `Base class events`, `Interface Xxx` separators in all modules in this order
- Classes have `Methods` section, standard modules have `Functions` section
- Keep `API` section ordered by consts first, then declares, types last
- Always use named params for optional parameters and boolean flags
- Drop hungarian notation on optional parameters i.e. `Optional Max As Long` not `lMax`
- Instead of `For lIdx = 1 To m_cItems.Count` prefer `For Each` loops as these are significantly faster with `VBA.Collections`
- Don't use any error handler in `Class_Terminate` (or methods called by it) as this will clear current values in `Err` object
- Inside procedure bodies don't write comments more than one line long. Procedure banners can be two lines long. Don't put comments on consts, member variables, UDTs and enums
