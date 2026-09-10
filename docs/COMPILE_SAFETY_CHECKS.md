# Compile-safety checks

Issue #45 adds two independent results to `tools/vba_lint.py`:

1. `error-handler labels resolve within their procedure`
2. `assigned identifiers are declared`

The active v1.4.1 gate has **15 checks**: the previous 13, including #51's
provenance fixtures, plus these two. Published v1.4.0 evidence remains 12/12.
Issue #64 hardens the assignment-declaration parser without adding a new gate
result, so the active count remains 15.

## Supported source grammar

The scanner recursively reads `.bas`, `.cls` and `.frm` exports in `src/`,
`test/` and `demo/`. It normalizes line endings in memory without changing files.
Object-module scope is determined from the export path: `.cls` and `.frm` are
object modules even when the exported header does not contain `VERSION ... CLASS`.

- Public, Private, Friend and implicit-public Sub/Function/Property procedures;
  Static procedures; multiline declarations/signatures; case-insensitive names.
- Escaped double-quoted strings, apostrophe/Rem comments, colon-separated
  statements, named arguments and single-line If/Then/Else arms.
- Procedure-local symbolic and numeric handler labels, including numbered lines.
  `On Error GoTo 0` and `On Error GoTo -1` are excluded by this rule; exclusion
  is not a claim about their runtime semantics or validity in every VBA host.
- Dim, Static, Const, parameters, module declarations and public standard-module
  declarations in other project files. Type fields are not standalone variables;
  class and UserForm instance fields remain object-module scoped and do not
  become project-global variables.
- `Property Let` and `Property Set` names are assignable within their own module.
  They are not exported into the project-wide assignment symbol set merely
  because the property is Public. A `Property Get` name is valid as a return
  assignment only inside that getter unless a Let/Set of the same name declares
  an assignable property in the module.
- Ordinary, Let and Set assignments; array-element assignments; For/For Each
  targets; ReDim/Preserve targets. Function and Property Get result assignments
  are permitted only within their own procedure.
- Qualified member assignments, including With-block members, are excluded:
  this rule does not resolve object types, members or default properties.

For this project's explicit-declaration discipline, `ReDim` targets must already
be declared. VBA's implicit declaration through ReDim is intentionally not an
escape hatch. A declaration match does not prove that an identifier is writable:
constant/enum mutability, type compatibility, bounds and RHS references remain
the compiler's responsibility.

## Conditional compilation

Each file is evaluated separately for pre-VBA7 Windows, VBA7 Win32 and VBA7
Win64. The scanner understands `#If`, `#ElseIf`, `#Else`, `#End If`, VBA7,
Win64, Win32, Mac, True/False, parentheses, Not, And, Or and Xor. Windows profiles
set Win32 true (including Win64) and Mac false. Each profile has separate symbol
and label tables: inactive branches cannot satisfy a reference.

Unknown conditional expressions/directives, including unimplemented `#Const`
syntax, fail both checks with an unsupported-source diagnostic. This is bounded
support for the repository's grammar, not a complete conditional-expression
interpreter. Extend the parser and fixtures before introducing other forms.

Bracketed identifiers, date-literal parsing, arbitrary VBA statement grammar,
external references, object-member resolution and full compiler/type semantics
are outside these checks. A pass does not establish VBA compilation. Diagnostics
identify the file, procedure, logical statement's first physical line, missing
identifier/label and conditional profile.

## Verification

```bash
python tools/test_vba_compile_safety.py
python tools/test_vba_compile_safety.py --integration
python tools/vba_lint.py --json vba-lint-results.json
```

The ordinary linter runs the local fixture matrix as part of the respective
two results, without adding a third check. Each single-defect fixture requires
the other check to remain green. #64 expands the matrix from 41 to **47 fixtures**
with Property Let/Set, Get-only, undeclared object-member and UserForm-scope cases.
The optional integration test copies the source tree to a temporary directory and
drives the real CLI in three cases: valid input, missing handler and undeclared
assignment. Each negative case must return 1 and fail only its intended JSON
result. The copy has no Git identity; this test proves CLI wiring, not
release-history verification. Exact-SHA gate evidence must come from a real clean
checkout separately.

No workflow or VBA-source change is required. Real Excel compilation and
execution remain separate certification gates.
