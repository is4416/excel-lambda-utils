# 📐 PolygonArea — Excel LAMBDA関数で多角形の面積を計算

2列（X, Y）の範囲から、多角形の面積を計算します。
範囲は少なくとも3行以上必要です。閉じた多角形の場合、最終行は最初の行と同じ座標にする必要はありません（関数内で自動で閉じます）。

**引数**

| 引数   | 型     | 説明                             |
| ------ | ------ | -------------------------------- |
| Points | Range  | 2列（X, Y）の座標範囲（最低3行） |
| Result | Number | 多角形の面積を返す               |

**備考**

- Result は戻り値です。引数としては不要です。
- Points は必ず2列の範囲で、1行目:X, 2列目:Y を示す座標である必要があります。
- 座標は半時計回りで配置する必要があります。
- 時計回りで配置した場合、面積は負（減算扱い）となります。

**コード**

```excel
= LAMBDA(Points, LET(
  RowCount  , ROWS(Points),
  DoubleArea, MAP(SEQUENCE(RowCount),
    LAMBDA(i, LET(
        i_before, IF(i = 1, RowCount, i - 1),
        i_after , IF(i = RowCount, 1, i + 1),
        (INDEX(Points, i_after, 2) - INDEX(Points, i_before, 2)) * INDEX(Points, i, 1)
      )
    )
  ),
  SUM(DoubleArea) / 2
))
```

**変数の詳細**

- Points    : Range, 2列 (X, Y) の座標範囲 (最低3行必要)
- RowCount  : Number, 座標の行数
- DoubleArea: Array, 各点に対する倍面積
- i         : カウンター
- i_before  : カウンターより1つ前の行番号か、RowCount
- i_after   : カウンターより1つ後の行番号か、1

**使用例**

PolygonArea という名前で、ブックに登録しているものとします。

```excel
= PolygonArea(A1:B10)
```

- A列：X座標
- B列：Y座標
- 範囲は3行以上
- 閉じた多角形の場合、最終行は自動で閉じられます
- 半時計回りの座標配置で正の面積が計算されます
