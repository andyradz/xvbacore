VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Money"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

  '
  ' The money amount.
  ' Never null.
  ' @serial
  '
  Private amount As Variant
  
  '
  ' The currency of the money, such as US Dollars or Euros.
  ' Never null.
  ' @serial
  '
  Private curry As Currency
  
  '
  ' The rounding style to be used.
  ' See {@link BigDecimal}.
  ' @serial
  '
  'private final RoundingMode rounding;

