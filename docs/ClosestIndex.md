# 🔍 ClosestIndex ー Excel LAMBDA関数でしきい値に一番近い値の、最初のインデックス番号を返す

指定された範囲から、しきい値に最も近い最初のインデックス番号を返します。

**引数**

= LAMBDA(Values, Threshold, LET(

| 引数      | 型     | 説明                                           |
| --------- | ------ | ---------------------------------------------- |
| Values    | Range  | 値の範囲                                       |
| Threshold | Number | しきい値                                       |
| Result    | Number | しきい値に一番近い値の、最初のインデックス番号 |

**備考**

- Result は戻り値です。引数としては不要です。

**コード**

```vb
= LAMBDA(Values, Threshold, LET(
  Diff, ABS(Threshold - Values),
  MATCH(MIN(Diff), Diff, FALSE)
)```

**変数の詳細**

- Diff : 値としきい値の差の絶対値の配列

**使用例**

ClosestIndex という名前で、ブックに登録しているものとします
> スピルには、対応していません

```vb
= ClosestIndex(A1:A5, 10)
```
