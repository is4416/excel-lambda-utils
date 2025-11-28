# 📏 DistancePoint — Excel LAMBDA関数で2点間の距離を計算

2次元平面上の2点間の距離（ユークリッド距離）を計算します。  
座標データを渡すだけで、隣接する点や任意の2点間の距離を求めることができます。

**引数**

| 引数   | 型     | 説明                           |
| ------ | ------ | ------------------------------ |
| PointA | Range  | 2列 (X, Y) の座標範囲          |
| PointB | Range  | 2列 (X, Y) の座標範囲          |
| Result | Number | 2点間の距離を返す              |

**備考**

- Result は戻り値です。引数としては不要です。

**コード**

```excel
= LAMBDA(PointA, PointB, LET(
  AX, TAKE(PointA,,1),
  AY, TAKE(DROP(PointA,,1),,1),
  BX, TAKE(PointB,,1),
  BY, TAKE(DROP(PointB,,1),,1),
  SQRT(
    (BX - AX)^2 + (BY - AY)^2
  )
))
```

**変数の詳細**

- AX: Number, PointAのX座標
- AY: Number, PointAのY座標
- BX: Number, PointBのX座標
- BY: Number, PointBのY座標

**使用例**

DistancePoint という名前で、ブックに登録しているものとします。
> スピルにも対応しています。

```excel
= DistancePoint(
  A2:B10,
  VSTACK(A3:B10, A2:B2)
)
```
