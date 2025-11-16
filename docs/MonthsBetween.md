# 🕒 MonthsBetween — Excel LAMBDA関数で月数をカウント

月数をカウントします  
月末締め若しくは、翌月の前日締めでカウントします

**引数**

| 引数       | 型                      | 説明                |
| ---------- | ----------------------- | ------------------- |
| StartDate  | Number (Excel Datetime) | 開始日              |
| EndDate    | Number (Excel Datetime) | 終了日              |
| EndOfMonth | Boolean                 | 集計方法 (省略可能) |

**備考**

- 戻り値は経過月数
- EndOfMonthを省略した場合、TRUE が設定されます  
  TRUE = 月末締め、FALSE = 翌月の前日締め

**コード**

```excel
= LAMBDA(StartDate, EndDate, EndOfMonth, LET(
  EOM, IF(ISOMITTED(EndOfMonth), TRUE, EndOfMonth),
  Y_1, YEAR(StartDate),
  M_1, MONTH(StartDate),
  D_1, DAY(StartDate),
  Y_2, YEAR(EndDate),
  M_2, MONTH(EndDate),
  D_2, DAY(EndDate),

  Correction, IF(EOM + (D_1 = 1) + (D_1 <= D_2), 1, 0),

  (Y_2 - Y_1) * 12 + M_2 - M_1 + Correction
))
```

**変数の詳細**

- EOM: Boolean, EndOfMonth を省略した場合は、TRUE となる
- Y_1: Number, YEAR(StartDate)
- M_1: Number, MONTH(StartDate)
- D_1: Number, DAY(StartDate)
- Y_2: Number, YEAR(EndDate)
- M_2: Number, MONTH(EndDate)
- D_2: Number, DAY(EndDate)
- Correction: Number, 日数補正

**使用例**

MonthsBetWeen という名前で、ブックに登録しているものとします
> スピルにも対応しています

```excel
= MonthsBetWeen(A1:A3, B1:B3, FALSE)
```
