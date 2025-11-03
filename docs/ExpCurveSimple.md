# 📈 ExpCurveSimple ー Excel LAMBDA関数で単純指数曲線の結果を計算

X列, Y列から単純指数曲線を作成し、結果を計算します。  
X列, Y列は少なくとも2行以上必要です。

**引数**

| 引数   | 型     | 説明                         |
| ------ | ------ | ---------------------------- |
| XRange | Range  | X列の範囲（最低2行）         |
| YRange | Range  | Y列の範囲（XRangeと同数）    |
| x      | Number | 計算するX値                  |
| Result | Number | 単純指数曲線の計算結果を返す |

**備考**

- Result は戻り値です。引数としては不要です。

**コード**

```excel
= LAMBDA(XRange, YRange, x, LET(
  y0 , INDEX(YRange, 1),
  LnY, MAP(YRange, LAMBDA(val, LN(val))),
  Res, LINEST(LnY, XRange),
  k  , INDEX(Res, 1),
  y0 * EXP(k * x)
))
```

**変数の詳細**

- y0 : Number, Y列の最初の値
- LnY: Range, YRangeの対数配列
- Res: 最小二乗法の結果
- k  : Number, 傾き

**式**

$y = y_0 * e^{kx}$

$ log (y) = log (y_0 * e^{kx}) $  
$ log (y) = log(y_0) + log(e^{kx}) $  
$ log (y) = log(y_0) + kx * log(e) $  
$ log (y) = log(y_0) + kx $ 線形に変換

**使用例**

ExpCurveSimple という名前で、ブックに登録しているものとします
> スピルにも対応しています

```excel
= ExpCurveSimple(A1:A10, B1:B10, A1:A20)
```
