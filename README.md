# Unexpected IsEmpty Behavior with Numeric Empty Values in VBScript

This repository demonstrates a subtle but potentially problematic quirk in VBScript's `IsEmpty` function.  The `IsEmpty` function doesn't reliably detect if a numeric variable has been initialized to `Empty`. This can lead to errors in situations where you intend to check if a numeric variable has a value assigned or not.

## The Bug

The provided `bug.vbs` script shows how `IsEmpty` fails to correctly identify an `Empty` numeric value.

## The Solution

The `bugSolution.vbs` script provides a workaround using the `VarType` function to more reliably check if a variable is uninitialized or assigned the value Empty for numeric values. Note that an alternative method is to leverage the fact that Empty is equal to 0 when involved in arithmetic operations.