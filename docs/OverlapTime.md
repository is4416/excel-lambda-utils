# 🕒 OverlapTime — Excel LAMBDA関数で時刻の重複を計算

時刻の重複を計算します  
時間帯によって単価が異なる契約の実績時間など、時間帯別集計の計算に使用できます

**引数**

| 引数      | 型                      | 説明                                             |
| --------- | ----------------------- | ------------------------------------------------ |
| StartDate | Number (Excel Datetime) | 開始日時 - EndDateと逆転しないよう入力           |
| EndDate   | Number (Excel Datetime) | 終了日時 - 日をまたぐ場合は [25:00] なども可     |
| MinDate   | Number (Excel Datetime) | 重複開始日時 - MaxDateと逆転しないよう入力       |
| MaxDate   | Number (Excel Datetime) | 重複終了日時 - 日をまたぐ場合は [25:00] なども可 |
| Result    | Number (Excel Datetime) | 重複時間の合計を返す                             |

**備考**

- Resultは戻り値です。引数としては不要です。
- 返り値は「日単位の数値」です。
- 時間表示にする場合は、セルの表示形式を`[h]:mm`等に設定してください。

**コード**

```excel
= LAMBDA(StartDate, EndDate, MinDate, MaxDate, LET(
  Days          , INT(EndDate) - INT(StartDate),
  StartTime     , MOD(StartDate, 1),
  EndTime       , MOD(EndDate, 1),
  MinTime       , MOD(MinDate, 1),
  MaxTime       , MOD(MaxDate, 1),
  UpperTime     , IF(EndDate > 1, 1, EndDate),
  UpperLimit    , IF(MaxDate > 1, 1, MaxDate),
  UpperFirst    , IF(UpperTime < UpperLimit, UpperTime, UpperLimit),
  LowerLimit    , IF((INT(MaxDate) > 0) * (StartTime < MaxTime) = 1, 0, MinTime),
  LowerFirst    , IF(StartTime < LowerLimit, LowerLimit, StartTime),
  TimeOfFirstDay, IF(UpperFirst - LowerFirst > 0, UpperFirst - LowerFirst, 0),
  TimeOfDays    , IF(Days > 1, (Days - 1) * (MaxDate - MinDate), 0),
  UpperLast     , IF(MinTime > MaxTime, 0, MinTime),
  LowerLast     , IF(EndTime < MaxTime, EndTime, MaxTime),
  TimeOfLastDay , IF(Days > 0, IF(LowerLast - UpperLast > 0, LowerLast - UpperLast, 0), 0),
  TimeOfFirstDay + TimeOfDays + TimeOfLastDay
))
```

**変数の詳細**

- Days          : Integer, 開始時刻-終了時刻の経過日数
- StartTime     : Number (Excel Datetime), 開始時刻の時間部分
- EndTime       : Number (Excel Datetime), 終了時刻の時間部分
- MinTime       : Number (Excel Datetime), 重複開始時刻の時間部分
- MaxTime       : Number (Excel Datetime), 重複終了時刻の時間部分
- UpperTime     : Number (Excel Datetime), 初日の上限時間 (24:00もしくはEndDate)
- UpperLimit    : Number (Excel Datetime), 初日の重複上限時間 (24:00もしくはMaxDate)
- UpperFirst    : Number (Excel Datetime), 初日の上限時間 (UpperTimeかUpperLimitの大きい方)
- LowerLimit    : Number (Excel Datetime), 重複が日またぎのとき、重複開始時刻の起算点を調整
- LowerFirst    : Number (Excel Datetime), 初日の下限時間 (StartTimeかLowerLimitの大きい方)
- TimeOfFirstDay: Number (Excel Datetime), 初日の重複時間
- TimeOfDays    : Number (Excel Datetime), 経過日の重複時間
- UpperLast     : Number (Excel Datetime), 終了日の上限時間 (0:00もしくはMinTime)
- LowerLast     : Number (Excel Datetime), 終了日の下限時間
- TimeOfLastDay : Number (Excel Datetime), 終了日の重複時間

**使用例**

OverlapTime という名前で、ブックに登録しているものとします
> スピルにも対応しています

```excel
= OverlapTime(A1:A10, B1:B10, TIMEVALUE("08:30"), TIMEVALUE("17:15"))
```
