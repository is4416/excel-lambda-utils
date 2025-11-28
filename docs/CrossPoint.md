# 📐 CrossPoint — Excel LAMBDA関数で線の交差する座標と、交差が指定した線の内部か判定します

線と線が交差する点と、交差が指定した線の内部か判定します。  
スピルにも対応しています

**引数**

| 引数   | 型                      | 説明                                           |
| ------ | ----------------------- | ---------------------------------------------- |
| LineA  | Range                   | 4列 (線の起点・終点のX, Y) の座標範囲          |
| LineB  | Range                   | 4列 (線の起点・終点のX, Y) の座標範囲          |
| Result | Number, Number, Boolean | 線の交点座標 (X, Y) と、線の内部か判定します   |

**備考**

- Result は戻り値です。引数としては不要です。
- Result は、X座標、Y座標、線の内部か (Boolean) を HSTACKで返します
- LineA, LineB は、必ず4列の範囲で、1列目:線の起点X, 2列目:線の起点Y, 3列目:線の終点X, 4列目:線の終点Y

**コード**

```excel
= LAMBDA(LineA, LineB, LET(
  AX , TAKE(LineA,,1),
  AY , TAKE(DROP(LineA,,1),,1),
  AXX, TAKE(DROP(LineA,,2),,1),
  AYY, TAKE(DROP(LineA,,3),,1),
  VA , AX = AXX,
  AA , IF(VA, 0, (AYY - AY) / (AXX - AX)),
  AB , IF(VA, 0, AY - AA * AX),

  BX , TAKE(LineB,,1),
  BY , TAKE(DROP(LineB,,1),,1),
  BXX, TAKE(DROP(LineB,,2),,1),
  BYY, TAKE(DROP(LineB,,3),,1),
  VB , BX = BXX,
  BA , IF(VB, 0, (BYY - BY) / (BXX - BX)),
  BB , IF(VB, 0, BY - BA * BX),

  CX, IF(VA = VB,
    IF(AA = BA,
      NA(),
      (BB - AB) / (AA - BA)
    ),
    IF(VA, AX, BX)
  ),

  CY, IF(ISNA(CX),
    NA(),
    IF(VA,
      BA * CX + BB,
      AA * CX + AB
    )
  ),

  LineAMinX, IF(AX < AXX, AX, AXX),
  LineAMaxX, IF(AX < AXX, AXX, AX),

  LineBMinX, IF(BX < BXX, BX, BXX),
  LineBMaxX, IF(BX < BXX, BXX, BX),

  MinX, IF(LineAMinX < LineBMinX, LineBMinX, LineAMinX),
  MaxX, IF(LineAMaxX < LineBMaxX, LineAMaxX, LineBMaxX),

  InLine, IF(ISNA(CX),
    NA(),
    (MinX <= CX) * (CX <= MaxX) <> 0
  ),

  HSTACK(CX, CY, InLine)
))
```

**変数の詳細**

- AX : Number, LineAの起点X座標
- AY : Number, LineAの起点Y座標
- AXX: Number, LineAの終点X座標
- AYY: Number, LineAの終点Y座標
- VA : Boolean, LineAが垂直か
- AA : Number, LineAの傾き
- AB : Number, LineAの切片
- BX, BY, BXX, BYY, VB, BA, BB: は、上記の LineBに対する変数
- CX, CY: LineAとLineBの交点。平行または同一線上にある場合、NA
- LineAMinX: Number, LineAの、X座標の小さい方
- LineAMaxX: Number, LineAの、X座標の大きい方
- LineBMinX, LineBMaxX: Number, 上記のLineBに対する変数
- MinX, MaxX: Number, 線の上限、下限のX値
- InLine: Boolean, 線の内部で重複しているか。平行または同一線上にある場合、NA

**使用例**

CrossPoint という名前で。ブックに登録しているものとします  
以下、3つの計算をスピルを用いて出力している例です

```excel
= CrossPoint(A1:D3, E1:H3)
```

A列: LineAの起点 X座標  
B列: LineAの起点 Y座標  
C列: LineAの終点 X座標  
D列: LineAの終点 Y座標  

E列: LineBの起点 X座標
F列: LineBの起点 Y座標
G列: LineBの終点 X座標
H列: LineBの終点 Y座標