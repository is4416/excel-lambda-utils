# 📈 ExpCurveModified ー Excel LAMBDA関数で修正指数曲線の結果を計算

X列, Y列から修正指数曲線を作成し、結果を計算します。  
X列, Y列は少なくとも2行以上必要です。  
上限値Lを省略する場合、""を指定してください。

**引数**

| 引数   | 型     | 説明                         |
| ------ | ------ | ---------------------------- |
| XRange | Range  | X列の範囲（最低2行）         |
| YRange | Range  | Y列の範囲（XRangeと同数）    |
| L      | Number | 上限値（省略可能）           |
| x      | Number | 計算するX値                  |
| Result | Number | 修正指数曲線の計算結果を返す |

**備考**

- Result は戻り値です。引数としては不要です。
- L に""を指定した場合、ROUNDUP(MAX(YRange), -1)が設定されます。

**コード**

```excel
= LAMBDA(XRange, YRange, L, x, LET(
  Limit, IF(L = "", MAX(YRange), L),
  Y    , MAP(YRange, LAMBDA(val, LN(Limit - val))),
  Res  , LINEST(Y, XRange),
  k    , INDEX(Res, 1),
  Limit - (Limit - INDEX(YRange, 1)) * EXP(- k * x)
))
```

**変数の詳細**

- Limit: 上限値
- Y    : Range, Y列の対数配列
- Res  : 最小二乗法の結果
- k    : Number, 傾き

**式**

$y = L - (L - y_0) * e^{-kx}$

$ (L - y_0) * e^{-kx} = L - y $  
$ log ((L - y_0) * e^{-k)x} = log (L - y) $  
$ log (L - y_0) + log (e^{-kx}) = log (L - y) $  
$ log (L - y_0) - kx * log(e) = log (L - y) $  
$ log (L - y) = log (L - y_0) - kx $ ※ 線形に変換

**使用例**

ExpCurveModified という名前で、ブックに登録しているものとします
> スピルにも対応しています

```excel
= ExpCurveModified(A1:A10, B1:B10, "", A1:A20)
```
