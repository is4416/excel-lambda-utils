# 🕒 TimeToDecimal — Excel LAMBDA関数で時刻 (TimeDate) を時間 (Float) に変換

時刻 (TimeDate) を、時間 (Float) に変換します

**引数**

| 引数      | 型                      | 説明            |
| --------- | ----------------------- | --------------- |
| T         | Number (Excel Datetime) | 時刻 (TimeDate) |
| Result    | Number (Float)          | 時間 (Float)    |

**備考**

- Resultは戻り値です。引数としては不要です。

**コード**

```vb
= LAMBDA(T,
  HOUR(T) + MINUTE(T) / 60 + SECOND(T) / 3600
)
```

**使用例**

TimeToDecimal という名前で、ブックに登録しているものとします
> スピルにも対応しています

```vb
= TimeToDecimal(TIMEVALUE("1:45"))
```

結果: 1.75
