# Calendar Schedule Builder Formula

## Purpose

These two formulas build calendar-like schedule tables in Google Sheets based on a selected month.  
The generated tables are later used by the `monthly-schedule-email.js` script to create and send a formatted HTML schedule email to contractors.

## What the first formula does

- Selects the correct month data from a calendar source sheet
- Adds leading empty cells depending on the weekday the month starts on
- Adds trailing empty cells depending on the weekday the month ends on
- Combines day numbers with weekday labels
- Splits the generated sequence into a calendar grid
- Prepares the output for further formatting and email automation

## What the second formula does

- Selects the correct month data from a calendar source sheet
- Detects which weekday the selected month starts on
- Adds leading placeholder cells so the first day appears under the correct weekday
- Combines calendar data with weekday labels
- Converts a linear sequence into a table-like grid
- Uses temporary text markers ('x') to shape the layout
- Removes helper marker characters from the final output
- Inserts line breaks for Saturday/week boundary handling

## Technologies / techniques used

- `ARRAYFORMULA`
- `SPLIT`
- `TRANSPOSE`
- `REGEXREPLACE`
- `FILTER`
- `IFS`
- `JOIN`
- `RIGHT`
- `CHAR(10)`
- dynamic range selection
- spreadsheet-based data transformation

## First sanitized Google Sheets formula

```text
=ArrayFormula(SPLIT(TRANSPOSE(SPLIT(REGEXREPLACE(REGEXREPLACE(IFS(
FILTER('Calendar'!$AH$4:$OH$4, 'Calendar'!$AH$2:$OH$2='Schedule Email'!A1)="Mo", "3 ", 
FILTER('Calendar'!$AH$4:$OH$4, 'Calendar'!$AH$2:$OH$2='Schedule Email'!A1)="Tu", "3 3 ", 
FILTER('Calendar'!$AH$4:$OH$4, 'Calendar'!$AH$2:$OH$2='Schedule Email'!A1)="We", "3 3 3 ", 
FILTER('Calendar'!$AH$4:$OH$4, 'Calendar'!$AH$2:$OH$2='Schedule Email'!A1)="Th", "3 3 3 3 ", 
FILTER('Calendar'!$AH$4:$OH$4, 'Calendar'!$AH$2:$OH$2='Schedule Email'!A1)="Fr", "3 3 3 3 3 ", 
FILTER('Calendar'!$AH$4:$OH$4, 'Calendar'!$AH$2:$OH$2='Schedule Email'!A1)="Sa", "3 3 3 3 3 3 ", 
FILTER('Calendar'!$AH$4:$OH$4, 'Calendar'!$AH$2:$OH$2='Schedule Email'!A1)="Su", "")
&
JOIN(" ", ArrayFormula(SPLIT(IFS(
'Schedule Email'!A1="January", JOIN(" ",ARRAYFORMULA('Calendar'!$AH$5:$BL$5)),
'Schedule Email'!A1="February", JOIN(" ",ARRAYFORMULA('Calendar'!$BM$5:$CN$5)),
'Schedule Email'!A1="March", JOIN(" ",ARRAYFORMULA('Calendar'!$CO$5:$DS$5)),
'Schedule Email'!A1="April", JOIN(" ",ARRAYFORMULA('Calendar'!$DT$5:$EW$5)), 
'Schedule Email'!A1="May", JOIN(" ",ARRAYFORMULA('Calendar'!$EX$5:$GB$5)),
'Schedule Email'!A1="June", JOIN(" ",ARRAYFORMULA('Calendar'!$GC$5:$HF$5)),
'Schedule Email'!A1="July", JOIN(" ",ARRAYFORMULA('Calendar'!$HG$5:$IK$5)),
'Schedule Email'!A1="August", JOIN(" ",ARRAYFORMULA('Calendar'!$IL$5:$JP$5)),
'Schedule Email'!A1="September", JOIN(" ",ARRAYFORMULA('Calendar'!$JQ$5:$KT$5)),
'Schedule Email'!A1="October", JOIN(" ",ARRAYFORMULA('Calendar'!$KU$5:$LY$5)),
'Schedule Email'!A1="November", JOIN(" ",ARRAYFORMULA('Calendar'!$LZ$5:$NC$5)),
'Schedule Email'!A1="December", JOIN(" ",ARRAYFORMULA('Calendar'!$ND$5:$OH$5))), 
" ")
&
SPLIT(IFS(
'Schedule Email'!A1="January", JOIN(" ",ARRAYFORMULA('Calendar'!$AH$4:$BL$4)),
'Schedule Email'!A1="February", JOIN(" ",ARRAYFORMULA('Calendar'!$BM$4:$CN$4)),
'Schedule Email'!A1="March", JOIN(" ",ARRAYFORMULA('Calendar'!$CO$4:$DS$4)),
'Schedule Email'!A1="April", JOIN(" ",ARRAYFORMULA('Calendar'!$DT$4:$EW$4)), 
'Schedule Email'!A1="May", JOIN(" ",ARRAYFORMULA('Calendar'!$EX$4:$GB$4)),
'Schedule Email'!A1="June", JOIN(" ",ARRAYFORMULA('Calendar'!$GC$4:$HF$4)),
'Schedule Email'!A1="July", JOIN(" ",ARRAYFORMULA('Calendar'!$HG$4:$IK$4)),
'Schedule Email'!A1="August", JOIN(" ",ARRAYFORMULA('Calendar'!$IL$4:$JP$4)),
'Schedule Email'!A1="September", JOIN(" ",ARRAYFORMULA('Calendar'!$JQ$4:$KT$4)),
'Schedule Email'!A1="October", JOIN(" ",ARRAYFORMULA('Calendar'!$KU$4:$LY$4)),
'Schedule Email'!A1="November", JOIN(" ",ARRAYFORMULA('Calendar'!$LZ$4:$NC$4)),
'Schedule Email'!A1="December", JOIN(" ",ARRAYFORMULA('Calendar'!$ND$4:$OH$4))), 
" ")))
&
IFS(
RIGHT(IFS(
'Schedule Email'!A1="January", JOIN(" ",ARRAYFORMULA('Calendar'!$AH$4:$BL$4)),
'Schedule Email'!A1="February", JOIN(" ",ARRAYFORMULA('Calendar'!$BM$4:$CN$4)),
'Schedule Email'!A1="March", JOIN(" ",ARRAYFORMULA('Calendar'!$CO$4:$DS$4)),
'Schedule Email'!A1="April", JOIN(" ",ARRAYFORMULA('Calendar'!$DT$4:$EW$4)), 
'Schedule Email'!A1="May", JOIN(" ",ARRAYFORMULA('Calendar'!$EX$4:$GB$4)),
'Schedule Email'!A1="June", JOIN(" ",ARRAYFORMULA('Calendar'!$GC$4:$HF$4)),
'Schedule Email'!A1="July", JOIN(" ",ARRAYFORMULA('Calendar'!$HG$4:$IK$4)),
'Schedule Email'!A1="August", JOIN(" ",ARRAYFORMULA('Calendar'!$IL$4:$JP$4)),
'Schedule Email'!A1="September", JOIN(" ",ARRAYFORMULA('Calendar'!$JQ$4:$KT$4)),
'Schedule Email'!A1="October", JOIN(" ",ARRAYFORMULA('Calendar'!$KU$4:$LY$4)),
'Schedule Email'!A1="November", JOIN(" ",ARRAYFORMULA('Calendar'!$LZ$4:$NC$4)),
'Schedule Email'!A1="December", JOIN(" ",ARRAYFORMULA('Calendar'!$ND$4:$OH$4))), 2)="Su", " 3 3 3 3 3 3 ",
RIGHT(IFS(
'Schedule Email'!A1="January", JOIN(" ",ARRAYFORMULA('Calendar'!$AH$4:$BL$4)),
'Schedule Email'!A1="February", JOIN(" ",ARRAYFORMULA('Calendar'!$BM$4:$CN$4)),
'Schedule Email'!A1="March", JOIN(" ",ARRAYFORMULA('Calendar'!$CO$4:$DS$4)),
'Schedule Email'!A1="April", JOIN(" ",ARRAYFORMULA('Calendar'!$DT$4:$EW$4)), 
'Schedule Email'!A1="May", JOIN(" ",ARRAYFORMULA('Calendar'!$EX$4:$GB$4)),
'Schedule Email'!A1="June", JOIN(" ",ARRAYFORMULA('Calendar'!$GC$4:$HF$4)),
'Schedule Email'!A1="July", JOIN(" ",ARRAYFORMULA('Calendar'!$HG$4:$IK$4)),
'Schedule Email'!A1="August", JOIN(" ",ARRAYFORMULA('Calendar'!$IL$4:$JP$4)),
'Schedule Email'!A1="September", JOIN(" ",ARRAYFORMULA('Calendar'!$JQ$4:$KT$4)),
'Schedule Email'!A1="October", JOIN(" ",ARRAYFORMULA('Calendar'!$KU$4:$LY$4)),
'Schedule Email'!A1="November", JOIN(" ",ARRAYFORMULA('Calendar'!$LZ$4:$NC$4)),
'Schedule Email'!A1="December", JOIN(" ",ARRAYFORMULA('Calendar'!$ND$4:$OH$4))), 2)="Mo", " 3 3 3 3 3 ", 
RIGHT(IFS(
'Schedule Email'!A1="January", JOIN(" ",ARRAYFORMULA('Calendar'!$AH$4:$BL$4)),
'Schedule Email'!A1="February", JOIN(" ",ARRAYFORMULA('Calendar'!$BM$4:$CN$4)),
'Schedule Email'!A1="March", JOIN(" ",ARRAYFORMULA('Calendar'!$CO$4:$DS$4)),
'Schedule Email'!A1="April", JOIN(" ",ARRAYFORMULA('Calendar'!$DT$4:$EW$4)), 
'Schedule Email'!A1="May", JOIN(" ",ARRAYFORMULA('Calendar'!$EX$4:$GB$4)),
'Schedule Email'!A1="June", JOIN(" ",ARRAYFORMULA('Calendar'!$GC$4:$HF$4)),
'Schedule Email'!A1="July", JOIN(" ",ARRAYFORMULA('Calendar'!$HG$4:$IK$4)),
'Schedule Email'!A1="August", JOIN(" ",ARRAYFORMULA('Calendar'!$IL$4:$JP$4)),
'Schedule Email'!A1="September", JOIN(" ",ARRAYFORMULA('Calendar'!$JQ$4:$KT$4)),
'Schedule Email'!A1="October", JOIN(" ",ARRAYFORMULA('Calendar'!$KU$4:$LY$4)),
'Schedule Email'!A1="November", JOIN(" ",ARRAYFORMULA('Calendar'!$LZ$4:$NC$4)),
'Schedule Email'!A1="December", JOIN(" ",ARRAYFORMULA('Calendar'!$ND$4:$OH$4))), 2)="Tu", " 3 3 3 3 ", 
RIGHT(IFS(
'Schedule Email'!A1="January", JOIN(" ",ARRAYFORMULA('Calendar'!$AH$4:$BL$4)),
'Schedule Email'!A1="February", JOIN(" ",ARRAYFORMULA('Calendar'!$BM$4:$CN$4)),
'Schedule Email'!A1="March", JOIN(" ",ARRAYFORMULA('Calendar'!$CO$4:$DS$4)),
'Schedule Email'!A1="April", JOIN(" ",ARRAYFORMULA('Calendar'!$DT$4:$EW$4)), 
'Schedule Email'!A1="May", JOIN(" ",ARRAYFORMULA('Calendar'!$EX$4:$GB$4)),
'Schedule Email'!A1="June", JOIN(" ",ARRAYFORMULA('Calendar'!$GC$4:$HF$4)),
'Schedule Email'!A1="July", JOIN(" ",ARRAYFORMULA('Calendar'!$HG$4:$IK$4)),
'Schedule Email'!A1="August", JOIN(" ",ARRAYFORMULA('Calendar'!$IL$4:$JP$4)),
'Schedule Email'!A1="September", JOIN(" ",ARRAYFORMULA('Calendar'!$JQ$4:$KT$4)),
'Schedule Email'!A1="October", JOIN(" ",ARRAYFORMULA('Calendar'!$KU$4:$LY$4)),
'Schedule Email'!A1="November", JOIN(" ",ARRAYFORMULA('Calendar'!$LZ$4:$NC$4)),
'Schedule Email'!A1="December", JOIN(" ",ARRAYFORMULA('Calendar'!$ND$4:$OH$4))), 2)="We", " 3 3 3 ", 
RIGHT(IFS(
'Schedule Email'!A1="January", JOIN(" ",ARRAYFORMULA('Calendar'!$AH$4:$BL$4)),
'Schedule Email'!A1="February", JOIN(" ",ARRAYFORMULA('Calendar'!$BM$4:$CN$4)),
'Schedule Email'!A1="March", JOIN(" ",ARRAYFORMULA('Calendar'!$CO$4:$DS$4)),
'Schedule Email'!A1="April", JOIN(" ",ARRAYFORMULA('Calendar'!$DT$4:$EW$4)), 
'Schedule Email'!A1="May", JOIN(" ",ARRAYFORMULA('Calendar'!$EX$4:$GB$4)),
'Schedule Email'!A1="June", JOIN(" ",ARRAYFORMULA('Calendar'!$GC$4:$HF$4)),
'Schedule Email'!A1="July", JOIN(" ",ARRAYFORMULA('Calendar'!$HG$4:$IK$4)),
'Schedule Email'!A1="August", JOIN(" ",ARRAYFORMULA('Calendar'!$IL$4:$JP$4)),
'Schedule Email'!A1="September", JOIN(" ",ARRAYFORMULA('Calendar'!$JQ$4:$KT$4)),
'Schedule Email'!A1="October", JOIN(" ",ARRAYFORMULA('Calendar'!$KU$4:$LY$4)),
'Schedule Email'!A1="November", JOIN(" ",ARRAYFORMULA('Calendar'!$LZ$4:$NC$4)),
'Schedule Email'!A1="December", JOIN(" ",ARRAYFORMULA('Calendar'!$ND$4:$OH$4))), 2)="Th", " 3 3 ", 
RIGHT(IFS(
'Schedule Email'!A1="January", JOIN(" ",ARRAYFORMULA('Calendar'!$AH$4:$BL$4)),
'Schedule Email'!A1="February", JOIN(" ",ARRAYFORMULA('Calendar'!$BM$4:$CN$4)),
'Schedule Email'!A1="March", JOIN(" ",ARRAYFORMULA('Calendar'!$CO$4:$DS$4)),
'Schedule Email'!A1="April", JOIN(" ",ARRAYFORMULA('Calendar'!$DT$4:$EW$4)), 
'Schedule Email'!A1="May", JOIN(" ",ARRAYFORMULA('Calendar'!$EX$4:$GB$4)),
'Schedule Email'!A1="June", JOIN(" ",ARRAYFORMULA('Calendar'!$GC$4:$HF$4)),
'Schedule Email'!A1="July", JOIN(" ",ARRAYFORMULA('Calendar'!$HG$4:$IK$4)),
'Schedule Email'!A1="August", JOIN(" ",ARRAYFORMULA('Calendar'!$IL$4:$JP$4)),
'Schedule Email'!A1="September", JOIN(" ",ARRAYFORMULA('Calendar'!$JQ$4:$KT$4)),
'Schedule Email'!A1="October", JOIN(" ",ARRAYFORMULA('Calendar'!$KU$4:$LY$4)),
'Schedule Email'!A1="November", JOIN(" ",ARRAYFORMULA('Calendar'!$LZ$4:$NC$4)),
'Schedule Email'!A1="December", JOIN(" ",ARRAYFORMULA('Calendar'!$ND$4:$OH$4))), 2)="Fr", " 3 ", 
RIGHT(IFS(
'Schedule Email'!A1="January", JOIN(" ",ARRAYFORMULA('Calendar'!$AH$4:$BL$4)),
'Schedule Email'!A1="February", JOIN(" ",ARRAYFORMULA('Calendar'!$BM$4:$CN$4)),
'Schedule Email'!A1="March", JOIN(" ",ARRAYFORMULA('Calendar'!$CO$4:$DS$4)),
'Schedule Email'!A1="April", JOIN(" ",ARRAYFORMULA('Calendar'!$DT$4:$EW$4)), 
'Schedule Email'!A1="May", JOIN(" ",ARRAYFORMULA('Calendar'!$EX$4:$GB$4)),
'Schedule Email'!A1="June", JOIN(" ",ARRAYFORMULA('Calendar'!$GC$4:$HF$4)),
'Schedule Email'!A1="July", JOIN(" ",ARRAYFORMULA('Calendar'!$HG$4:$IK$4)),
'Schedule Email'!A1="August", JOIN(" ",ARRAYFORMULA('Calendar'!$IL$4:$JP$4)),
'Schedule Email'!A1="September", JOIN(" ",ARRAYFORMULA('Calendar'!$JQ$4:$KT$4)),
'Schedule Email'!A1="October", JOIN(" ",ARRAYFORMULA('Calendar'!$KU$4:$LY$4)),
'Schedule Email'!A1="November", JOIN(" ",ARRAYFORMULA('Calendar'!$LZ$4:$NC$4)),
'Schedule Email'!A1="December", JOIN(" ",ARRAYFORMULA('Calendar'!$ND$4:$OH$4))), 2)="Sa", ""),
"Sa","Sa "&CHAR(10)), "[a-zA-Z]", ""),CHAR(10)))," "))
```

## Second sanitized Google Sheets formula

```text
=ArrayFormula(REGEXREPLACE(SPLIT(TRANSPOSE(SPLIT(REGEXREPLACE(IFS(
FILTER('Calendar'!$AH$4:$OH$4, 'Calendar'!$AH$2:$OH$2='Schedule Email'!A1)="Mo", "x ", 
FILTER('Calendar'!$AH$4:$OH$4, 'Calendar'!$AH$2:$OH$2='Schedule Email'!A1)="Tu", "x x ", 
FILTER('Calendar'!$AH$4:$OH$4, 'Calendar'!$AH$2:$OH$2='Schedule Email'!A1)="We", "x x x ", 
FILTER('Calendar'!$AH$4:$OH$4, 'Calendar'!$AH$2:$OH$2='Schedule Email'!A1)="Th", "x x x x ", 
FILTER('Calendar'!$AH$4:$OH$4, 'Calendar'!$AH$2:$OH$2='Schedule Email'!A1)="Fr", "x x x x x ", 
FILTER('Calendar'!$AH$4:$OH$4, 'Calendar'!$AH$2:$OH$2='Schedule Email'!A1)="Sa", "x x x x x x ", 
FILTER('Calendar'!$AH$4:$OH$4, 'Calendar'!$AH$2:$OH$2='Schedule Email'!A1)="Su", "")
&
JOIN(" ", ArrayFormula(SPLIT(IFS(
'Schedule Email'!A1="January", JOIN(" ",ARRAYFORMULA('Calendar'!$AH$3:$BL$3)),
'Schedule Email'!A1="February", JOIN(" ",ARRAYFORMULA('Calendar'!$BM$3:$CN$3)),
'Schedule Email'!A1="March", JOIN(" ",ARRAYFORMULA('Calendar'!$CO$3:$DS$3)),
'Schedule Email'!A1="April", JOIN(" ",ARRAYFORMULA('Calendar'!$DT$3:$EW$3)), 
'Schedule Email'!A1="May", JOIN(" ",ARRAYFORMULA('Calendar'!$EX$3:$GB$3)),
'Schedule Email'!A1="June", JOIN(" ",ARRAYFORMULA('Calendar'!$GC$3:$HF$3)),
'Schedule Email'!A1="July", JOIN(" ",ARRAYFORMULA('Calendar'!$HG$3:$IK$3)),
'Schedule Email'!A1="August", JOIN(" ",ARRAYFORMULA('Calendar'!$IL$3:$JP$3)),
'Schedule Email'!A1="September", JOIN(" ",ARRAYFORMULA('Calendar'!$JQ$3:$KT$3)),
'Schedule Email'!A1="October", JOIN(" ",ARRAYFORMULA('Calendar'!$KU$3:$LY$3)),
'Schedule Email'!A1="November", JOIN(" ",ARRAYFORMULA('Calendar'!$LZ$3:$NC$3)),
'Schedule Email'!A1="December", JOIN(" ",ARRAYFORMULA('Calendar'!$ND$3:$OH$3)))," ")
&
SPLIT(IFS(
'Schedule Email'!A1="January", JOIN(" ",ARRAYFORMULA('Calendar'!$AH$4:$BL$4)),
'Schedule Email'!A1="February", JOIN(" ",ARRAYFORMULA('Calendar'!$BM$4:$CN$4)),
'Schedule Email'!A1="March", JOIN(" ",ARRAYFORMULA('Calendar'!$CO$4:$DS$4)),
'Schedule Email'!A1="April", JOIN(" ",ARRAYFORMULA('Calendar'!$DT$4:$EW$4)), 
'Schedule Email'!A1="May", JOIN(" ",ARRAYFORMULA('Calendar'!$EX$4:$GB$4)),
'Schedule Email'!A1="June", JOIN(" ",ARRAYFORMULA('Calendar'!$GC$4:$HF$4)),
'Schedule Email'!A1="July", JOIN(" ",ARRAYFORMULA('Calendar'!$HG$4:$IK$4)),
'Schedule Email'!A1="August", JOIN(" ",ARRAYFORMULA('Calendar'!$IL$4:$JP$4)),
'Schedule Email'!A1="September", JOIN(" ",ARRAYFORMULA('Calendar'!$JQ$4:$KT$4)),
'Schedule Email'!A1="October", JOIN(" ",ARRAYFORMULA('Calendar'!$KU$4:$LY$4)),
'Schedule Email'!A1="November", JOIN(" ",ARRAYFORMULA('Calendar'!$LZ$4:$NC$4)),
'Schedule Email'!A1="December", JOIN(" ",ARRAYFORMULA('Calendar'!$ND$4:$OH$4))), " "))), "Sa","Sa"&CHAR(10)), CHAR(10))), " "), "[a-zA-Z]", ""))