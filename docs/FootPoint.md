# 📐 FootPoint — Excel LAMBDA関数で線と点が垂直に交差する点と、距離を返す

線と点が垂直に交差する点と、距離を返します。  
スピルにも対応しています

**引数**

| 引数   | 型                     | 説明                                           |
| ------ | ---------------------- | ---------------------------------------------- |
| Line   | Range                  | 4列 (線の基点・終点のX, Y) の座標範囲          |
| Point  | Range                  | 2列 (X, Y) の座標範囲                          |
| Result | Number, Number, Number | 垂直に交差する点の座標 (X, Y) と、点からの距離 |

**備考**

- Result は戻り値です。引数としては不要です。
- Result は、X座標, Y座標, 点からの距離 を HSTACKで返します
- Line   は必ず4列の範囲で、1列目:線の基点X, 2列目:線の基点Y, 3列目:線の終点X, 4列目:線の終点Y
- Point  は必ず2列の範囲で、1列目:X, 2列目:Y

**コード**

```excel
= LAMBDA(Line, Point, LET(
  X  , TAKE(Line,,1),
  XX , TAKE(DROP(Line,,2),,1),
  XXX, TAKE(Point,,1),

  Y  , TAKE(DROP(Line,,1),,1),
  YY , TAKE(DROP(Line,,3),,1),
  YYY, TAKE(DROP(Point,,1),,1),

  DX , XX - X,
  DY , YY - Y,
  L  , DX^2 + DY^2,
  T  , IF(L = 0,
    NA(),
    ((XXX - X) * DX + (YYY - Y) * DY) / L
  ),

  FX, X + T * DX,
  FY, Y + T * DY,

  HSTACK(FX, FY, SQRT((FX - XXX)^2 + (FY - YYY)^2))
))
```

**変数の詳細**

- X  : Number, 線の基点X座標
- XX : Number, 線の終点X座標
- XXX: Number, 点のX座標
- Y  : Number, 線の基点Y座標
- YY : Number, 線の終点Y座標
- YYY: Number, 点のY座標
- DX : Number, 線のX幅
- DY : Number, 線のY幅
- L  : Number, 線の長さの2乗
- T  : Number, 割合計算 (L = 0 の時は、エラー値を返す)
- FX : Number, 垂直に交わる点のX座標
- FY : Number, 垂直に交わる点のY座標

**使用例**

FootPoint という名前で。ブックに登録しているものとします  
以下、3つの計算をスピルを用いて出力している例です

```excel
= FootPoint(A1:D3, E1:F3)
```

A列: 線の基点 X座標
B列: 線の基点 Y座標
C列: 線の終点 X座標
D列: 線の終点 Y座標
E列: 任意の点 X座標
F列: 任意の点 Y座標
