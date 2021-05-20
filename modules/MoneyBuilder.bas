Attribute VB_Name = "MoneyBuilder"
Option Explicit

' Input Number    UP  DOWN    CEILING FLOOR   HALF_UP HALF_DOWN   HALF_EVEN   UNNECESSARY
' 5.5             6   5         6      5       6       5           6           throw ArithmeticException
' 2.5             3   2         3   2       3   2   2   throw ArithmeticException
' 1.6             2   1         2   1       2   2   2   throw ArithmeticException
' 1.1             2   1         2   1       1   1   1   throw ArithmeticException
' 1.0             1   1         1   1       1   1   1   1
' -1.0           -1  -1         -1      -1  -1  -1  -1  -1
' -1.1           -2  -1         -1      -2  -1  -1  -1  throw ArithmeticException
' -1.6           -2  -1         -1      -2  -2  -2  -2  throw ArithmeticException
' -2.5           -3  -2         -2      -3  -3  -2  -2  throw ArithmeticException
' -5.5           -6   -5        -5      -6  -6  -5  -6  throw ArithmeticException

'Public Type RoundingMode
    'Rounding mode to round towards positive infinity.
    'Ceiling
    'Rounding mode to round towards zero.
    'Down
    'Rounding mode to round towards negative infinity.
    'Floor
    'Rounding mode to round towards "nearest neighbor" unless both neighbors are equidistant, in which case round down.
    'Half_Down
    'Rounding mode to round towards the "nearest neighbor" unless both neighbors are equidistant, in which case, round towards the even neighbor.
    'Half_Even
    'Rounding mode to round towards "nearest neighbor" unless both neighbors are equidistant, in which case round up.
    'Half_Up
    'Rounding mode to assert that the requested operation has an exact result, hence no rounding is necessary.
    'UNNECESSARY
    'Rounding mode to round away from zero.
    'Up
'End Type
