# Excel Lambda Utils — Documentation

## Types

```Types
  Type Date    : Number (y/M/d h:m)
  Type Time    : Number (h:m)
  Type Duration: Number ([h]:m)
  Type Point   : HSTACK (Number, Number)
```
## Time Calculations

OverlapTime: Calculates overlapping duration between date/time ranges.
```OverlapTime
Function OverlapTime (
  StateDate, EndDate: Date
  MinTime, MaxTime  : Time
): Duration
```

TimeToDecimal: Converts a time (TimeDate) to a decimal number representing hours.
```TimeToDecimal
Function TimeToDecimal (
  T: Time
): Number (Decimal)
```

DecimalToTime: Converts a decimal hour value back to a time (TimeDate).
```DecimalToTime
Function DecimalToTime (
  F: Number (Decimal)
): Time
```

MonthsBetween: Counts the number of months between two dates (supports end-of-month or day-before-next-month conventions).
```MonthsBetween
Function MonthsBetween (
  StartDate, EndDate: Date
  EndOfMonth        : Boolean (optional)
  PivotDay          : Number (optional)
): Number
```

LastDay: Returns the last day of a specified date.
```LastDay
Function LastDay (
  TargetDate: Date
): Number
```

DiffDaysTime: Subtracts days and times from a given date/time, with an optional daily time span.
```DiffDaysTime
Function DiffDaysTime (
  StartDays : Number
  StartTime : Time
  SubDays   : Number
  SubTime   : Time
  TimePerDay: Time (optional)
): HSTACK (Number, Time)
```

## Coordinate Calculations

DistancePoint: Calculates the distance between two 2D points.
```DistancePoint
Function DistancePoint (
  PointA, PointB: Point
): Point
```

PolygonArea: Computes the area of a polygon defined by coordinates.
```PolygonArea
Function PolygonArea (
  Points: VSTACK (Point)
): Number
```

FootPoint: Returns the perpendicular intersection of a point onto a line and the distance from the point.
```FootPoint
Function FootPoint(
  Line : HSTACK(Point, Point)
  Point: Point
): HSTACK (Point, Number)
```

CrossPoint: Returns the intersection point of two lines and a boolean indicating whether the intersection lies within the specified line segments.
```CrossPoint
Function CrossPoint(
  LineA: HSTACK(Point, Point)
  LineB: HSTACK(Point, Point)
): HSTACK (Point, Boolean)
```

## Area / Volume Calculations

PolygonArea handles arbitrary polygons, while TriangleArea functions handle triangles from sides and/or angles.

TriangleAreaSSS: Calculates the area of a triangle from three side lengths.
```TriangleAreaSSS
Function TriangleAreaSSS(
  A, B, C: Number
): Number
```

TriangleAreaSAS: Calculates the area of a triangle from two sides and the angle between them.
```TriangleAreaSAS
Function TriangleAreaSAS(
  A: Number
  R: Number (degrees)
  B: Number
): Number
```

TriangleAreaASA: Calculates the area of a triangle from one side and its two end angles.
```TriangleAreaASA
Function TriangleAreaASA(
  A     : Number
  LR, RR: Number (degrees)
): Number
```

SectionVolume: Calculate volume from SP cross section (average section method, prismoidal method, Toda correction)
```SectionVolume
Function SectionVolume(
  SPRange, ARange: VSTACK (Number)
  UniformSpan    : Boolean (optional)
  Alpha          : Number (optional)
): VSTACK ( HSTACK (Number, Number) )
```

## Curve Calculations

PowerCurve: Generates a power curve and computes its value.
```PowerCurve
Function PowerCurve (
  XRange: VSTACK (Number)
  YRange: VSTACK (Number)
  x     : Number
): Number
```

ExpCurveSimple: Generates a simple exponential curve and computes its value.
```ExpCurveSimple
Function ExpCurveSimple (
  XRange: VSTACK (Number)
  YRange: VSTACK (Number)
  x     : Number
): Number
```

ExpCurveModified: Generates a modified exponential curve and computes its value.
```ExpCurveModified
Function ExpCurveModified (
  XRange: VSTACK (Number)
  YRange: VSTACK (Number)
  L     : Number (optional)
  x     : Number
): Number
```

LogisticCurve: Generates a logistic curve and computes its value.
```LogisticCurve
Function LogisticCurve (
  XRange: VSTACK (Number)
  YRange: VSTACK (Number)
  L     : Number (optional)
  Xo    : Number (optional)
  x     : Number
): Number
```

## String Operations

SmartSplit: Safely splits CSV/JSON-like strings, handling quotes and escape characters.
```SmartSplit
Function SmartSplit (
  S: string
): VSTACK (string)
```

SmartJoin: Joins a range into a CSV-like string, quoting values and escaping quotes as [""] .
```SmartJoin
Function SmartJoin (
  Rng: VSTACK (string)
): string
```

Converts a number into the corresponding Excel column label.
```NumberToColumn
Function NumberToColumn (
  Num: Number
): string
```

Converts an Excel column label into its corresponding number.
```ColumnToNumber
Function ColumnToNumber (
  Str: string
): Number
```
