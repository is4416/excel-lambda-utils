# 📏 DistancePoint — Excel LAMBDA関数で2点間の距離を計算

2次元平面上の2点間の距離（ユークリッド距離）を計算します。  
座標データを渡すだけで、隣接する点や任意の2点間の距離を求めることができます。

**引数**

| 引数   | 型     | 説明                           |
| ------ | ------ | ------------------------------ |
| P1X    | Number | 1点目のX座標（数値または範囲） |
| P1Y    | Number | 1点目のY座標（数値または範囲） |
| P2X    | Number | 2点目のX座標（数値または範囲） |
| P2Y    | Number | 2点目のY座標（数値または範囲） |
| Result | Number | 2点間の距離を返す              |

**備考**

- Result は戻り値です。引数としては不要です。

**コード**

```excel
= LAMBDA(P1X, P1Y, P2X, P2Y,
  SQRT(
    (P1X - P2X)^2 + (P1Y - P2Y)^2
  )
)
```

**変数の詳細**

- P1X: Number, 1点目のX座標
- P1Y: Number, 1点目のY座標
- P2X: Number, 2点目のX座標
- P2Y: Number, 2点目のY座標

**使用例**

DistancePoint という名前で、ブックに登録しているものとします。
> スピルにも対応しています。

```excel
= DistancePoint(A2:A10, B2:B10, VSTACK(A3:A10, A2), VSTACK(B3:B10, B2))
```
