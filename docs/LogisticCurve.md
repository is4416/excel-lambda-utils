# 📈 LogisticCurve ー Excel LAMBDA関数でロジスティック曲線の結果を計算

X列, Y列からロジスティック曲線を作成し、結果を計算します。  
X列, Y列は少なくとも2行以上必要です。  
上限値Lを省略する場合、""を指定してください。  
中心点Xoを省略する場合、""を指定してください。  
※ 最小二乗法による計算であり、精度に限界があります

**引数**

| 引数   | 型     | 説明                               |
| ------ | ------ | ---------------------------------- |
| XRange | Range  | X列の範囲（最低2行）               |
| YRange | Range  | Y列の範囲（XRangeと同数）          |
| L      | Number | 上限値（省略可能）                 |
| Xo     | Number | 中心点（省略可能）                 |
| x      | Number | 計算するX値                        |
| Result | Number | ロジスティック曲線の計算結果を返す |

**備考**

- Result は戻り値です。引数としては不要です。
- L に "" を指定した場合、`MAX(YRange)` が自動設定されます。  
- Xo に "" を指定した場合、最小二乗法の結果から自動推定されます。  
- 本関数は、相対誤差(%)を最小化します。  
- Y値が上限値 L に近い場合やノイズが大きい場合には誤差が増大します。  
- 高精度が必要な場合、Solver または非線形最小二乗法を使用してください。 

**コード**

```excel
= LAMBDA(XRange, YRange, L, Xo, x, LET(
  Limit, IF(L = "", MAX(YRange), L),
  Y    , MAP(YRange, LAMBDA(val, LN((Limit - val) / val))),
  Res  , LINEST(Y, XRange),
  k    , - INDEX(Res, 1, 1),
  Mid  , IF(Xo = "", INDEX(Res, 1, 2) / k, Xo),
  Limit / (1 + EXP(- k * (x - Mid)))
))
```

**変数の詳細**

- Limit: 上限値
- Y    : Range, Y列の対数配列
- Res  : 最小二乗法の結果
- k    : Number, 傾き
- Mid  : Number, 中心値

**式**

$ y = \frac{L}{1 + e^{-k(x-x_0)}} $

$ \frac{1}{y} = \frac{(1 + e^{-k(x-x_0)})}{L} $  
$ L = y * (1 + e^{-k(x-x_0)}) $  
$ L = y + y * e^{-k(x-x_0)} $  
$ L - y = y * e^{-k(x-x_0)} $  
$ \frac{L - y}{y} = e^{-k(x-x_0)} $  
$ log (\frac{L - y}{y}) = log(e^{-k(x-x_0)}) $  
$ log (\frac{L - y}{y}) = -k(x-x_0) * log(e) $  
$ log (\frac{L - y}{y}) = -k(x-x_0) $

**使用例**

LogisticCurve という名前で、ブックに登録しているものとします
> スピルにも対応しています

```excel
= LogisticCurve(A1:A10, B1:B10, "", "", A1:A20)
```
