# 🕒 DiffDaysTime — Excel LAMBDA関数で勤務時間計算を行う

日 + 時間 から、日 + 時間を控除します  
勤務時間管理を目的としているため、1日当たりの時間を指定できます  
1日当たりの時間を省略した場合、7時間45分となります

**引数**

| 引数       | 型                       | 説明                       |
| ---------- | ------------------------ | -------------------------- |
| StartDays  | Number                   | 当初の日数                 |
| StartTime  | Number (Excel Datetime)  | 当初の時間                 |
| SubDays    | Number                   | 控除する日数               |
| SubTime    | Number (Excel Datetime)  | 控除する時間               |
| TimePerDay | Number (Excel Datetime)  | 1日当たりの時間 (省略可能) |
| Result     | Range (Number, Datetime) | 控除後の日数, 時間         |

**備考**

- Result は戻り値です。引数としては不要です
- Result は HSTACK で返ります
- TimePerDay が省略された場合、1日当たり 7時間45分となります

**コード**

```excel
= LAMBDA(StartDays, StartTime, SubDays, SubTime, TimePerDay, LET(
  TPD, IF(ISOMITTED(TimePerDay), TIME(7, 45, 0), TimePerDay),
  T  , IF(StartTime < SubTime, TPD, 0) + StartTime - SubTime,
  D  , StartDays - SubDays - IF(StartTime < SubTime, 1, 0),
  HSTACK(D, T)
))
```

**使用例**

DiffDaysTime という名前で、ブックに登録しているものとします
> スピルにも対応しています

```excel
= DiffDaysTime(30, TIME(0, 30, 0), 0, TIME(3, 0, 0),)
```

結果:
30日 0時間 30分 - 0日 3時間 0分 = 29日 5時間 15分
