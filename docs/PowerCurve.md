# 📈 PowerCurve ー Excel LAMBDA関数でべき乗曲線の結果を計算

X列, Y列からべき乗曲線を作成し、結果を計算します。  
X列, Y列は少なくとも2行以上必要です。

**引数**

| 引数   | 型     | 説明                       |
| ------ | ------ | -------------------------- |
| XRange | Range  | X列の範囲（最低2行）       |
| YRange | Range  | Y列の範囲（XRangeと同数）  |
| x      | Number | 計算するX値                |
| Result | Number | べき乗曲線の計算結果を返す |

**備考**

- Result は戻り値です。引数としては不要です。

**コード**

```excel
= LAMBDA(XRange, YRange, x, LET(
  LnX, MAP(XRange, LAMBDA(val, LN(val))),
  LnY, MAP(YRange, LAMBDA(val, LN(val))),
  Res, LINEST(LnY, LnX),
  a  , EXP(INDEX(Res, 2)),
  b  , INDEX(Res, 1),
  a * x^b
))
```

**変数の詳細**

- LnX: Range, XRangeの対数配列
- LnY: Range, YRangeの対数配列
- Res: 最小二乗法の結果
- a  : 切片
- b  : 傾き

**式**

$y = ax^{b}$

$ log (y) = log(ax^{b}) $  
$ log (y) = log(a) + log(x^{b}) $  
$ log (y) = log(a) + b * log(x) $ ※ 線形に変換

**使用例**

PowerCurve という名前で、ブックに登録しているものとします
> スピルにも対応しています

```excel
= PowerCurve(A1:A10, B1:B10, A1:A20)
```
