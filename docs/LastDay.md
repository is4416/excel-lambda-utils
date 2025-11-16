# 🕒 LastDay — Excel LAMBDA関数で最終日を返す

指定された日を含む月の最終日を返します  
`EOMONTH(Target, 0)` と同じ値を返しますが、スピル対応しています

**引数**

| 引数       | 型                      | 説明   |
| ---------- | ----------------------- | ------ |
| TargetDate | Number (Excel Datetime) | 指定日 |

**備考**

- 戻り値は最終日 (Excel Datetime)

**コード**

```excel
= LAMBDA(TargetDate, DATE(YEAR(TargetDate), MONTH(TargetDate) + 1, 0))
```

**使用例**

LastDay という名前で、ブックに登録しているものとします
> スピルにも対応しています

```excel
= LastDay(A1:A3)
```
