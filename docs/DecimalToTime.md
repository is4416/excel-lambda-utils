# 🕒 DecimalToTime — Excel LAMBDA関数で時間 (Float) を時刻 (TimeDate) に変換

時間 (Float) を、時刻 (TimeDate) に変換します

**引数**

| 引数      | 型                      | 説明            |
| --------- | ----------------------- | --------------- |
| F         | Number (Float)          | 時間 (Float)    |
| Result    | Number (Excel Datetime) | 時刻 (TimeDate) |

**備考**

- Resultは戻り値です。引数としては不要です。

**コード**

```excel
= LAMBDA(F, LET(
  TotalSec, ROUND(F * 3600, 0),
  H       , INT(TotalSec / 3600),
  M       , INT(MOD(TotalSec, 3600) / 60),
  S       , MOD(TotalSec, 60),
  TIME(H, M, S)
))
```

**引数の詳細**

- TotalSec: Number, 全体を秒に変換
- H       : Number, 時間
- M       : Number, 分
- S       : Number, 秒

**使用例**

DecimalToTime という名前で、ブックに登録しているものとします
> スピルにも対応しています

```excel
= TimeToDecimal(TIMEVALUE("1:45"))
= DecimalToTime(1.75)
```

結果: 1:45
