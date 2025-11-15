# 🕒 OverlapTime — Excel LAMBDA関数で時刻の重複を計算

時刻の重複を計算します  
時間帯によって単価が異なる契約の実績時間など、時間帯別集計の計算に使用できます

**引数**

| 引数      | 型                      | 説明                 |
| --------- | ----------------------- | -------------------- |
| StartDate | Number (Excel Datetime) | 開始日時             |
| EndDate   | Number (Excel Datetime) | 終了日時             |
| MinTime   | Number (Excel Datetime) | 重複開始時刻         |
| MaxTime   | Number (Excel Datetime) | 重複終了時刻         |
| Result    | Number (Excel Datetime) | 重複時間の合計を返す |

**備考**

- Resultは戻り値です。引数としては不要です。
- 返り値は「日単位の数値」です。
- 時間表示にする場合は、セルの表示形式を`[h]:mm`等に設定してください。
- MinTime > MaxTime (22:00-5:00のような日またぎ範囲) にも対応しています。

**コード**

```excel
= LAMBDA(StartDate,EndDate,MinTime,MaxTime, LET(
  StartTime, MOD(StartDate, 1),
  EndTime  , MOD(EndDate, 1),
  MinT     , MOD(MinTime, 1),
  MaxT     , MOD(MaxTime, 1),
  Buf      , INT(EndDate) - INT(StartDate),
  Days     , IF(Buf > 0, Buf, IF(StartTime >= EndTime, 1, 0)),

  TimeOfOneDay, MaxT - MinT + IF(MinT < MaxT, 0, 1),
  TimeOfDays  , (Days - 1) * TimeOfOneDay,

  FirstDayUpperLimit, MaxT,
  FirstDayUpperTime , IF(Days > 0, 1, EndTime),
  FirstDayUpper     , IF(FirstDayUpperLimit < FirstDayUpperTime,
    FirstDayUpperLimit, FirstDayUpperTime
  ),

  FirstDayLowerLimit, IF(MinT >= MaxT, 0, MinT),
  FirstDayLowerTime , StartTime,
  FirstDayLower     , IF(FirstDayLowerLimit < FirstDayLowerTime,
    FirstDayLowerTime, FirstDayLowerLimit
  ),
  TimeOfFirstDay, FirstDayUpper - FirstDayLower,

  TimeOfFirstDayBefore, IF((MinT >= MaxT) * (MaxT < StartTime) = 0, 0,
    IF(Days > 0, 1, EndTime) - IF(StartTime < MinT, MinT, StartTime)
  ),

  LastDayUpper, IF(MaxT < EndTime, MaxT, EndTime),

  LastDayLowerLimit, IF(MinT >= MaxT, 0, MinT),
  LastDayLowerTime , IF(StartTime < EndTime, StartTime, 0),
  LastDayLower     , IF(LastDayLowerLimit < LastDayLowerTime,
    LastDayLowerTime, LastDayLowerLimit
  ),

  TimeOfLastDay, IF(Days > 0, LastDayUpper - LastDayLower, 0),

  IF(TimeOfFirstDay < 0, 0, TimeOfFirstDay) +
  IF(TimeOfFirstDayBefore < 0, 0, TimeOfFirstDayBefore) +
  IF(TimeOfDays < 0, 0, TimeOfDays) +
  IF(TimeOfLastDay < 0, 0, TimeOfLastDay)
))
```

**変数の詳細**

- StartTime: Number (Excel Datetime), StartDate の時刻部分
- EndTime  : Number (Excel Datetime), EndDate の時刻部分
- MinT     : Number (Excel Datetime), MinTime の時刻部分
- MaxT     : Number (Excel Datetime), MaxTime の時刻部分
- Buf      : 開始日と終了日の差
- Days     : 経過日数。StartTime >= EndTime のときは1日補正

- TimeOfOneDay: 1日あたりの重複時間。MinT > MaxT のときは1日補正
- TimeOfDays  : 中間日の合計重複時間

- FirstDayUpperLimit: 初日の上限基準
- FirstDayUpperTime : 初日の終了時刻
- FirstDayUpper     : 重複する上限時間 (MIN(FirstDayUpperLimit, FirstDayUpperTime))
- FirstDayLower     : 初日の下限時間
- TimeOfFirstDay    : 初日の重複時間

- TimeOfFirstDayBefore: 日またぎ時の時間調整

- LastDayLowerLimit: 最終日の下限基準
- LastDayLowerTime : 最終日の開始時刻
- LastDayLower     : 重複する下限時間 (MAX(LastDayLowerLimit, LastDayLowerTime))
- LastDayUpper     : 最終日の上限時間
- TimeOfLastDay    : 最終日の重複時間

**使用例**

OverlapTime という名前で、ブックに登録しているものとします
> スピルにも対応しています

```excel
= OverlapTime(A1:A10, B1:B10, TIMEVALUE("08:30"), TIMEVALUE("17:15"))
```
