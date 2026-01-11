# ExpCurveModified 使用例

## 説明
2つの列から、修正指数曲線を作成し、値を計算します。  
上限値を指定しない場合、Y値の最大値付近 `ROUNDUP(MAX(YRange), -1)` を設定します。

## X値がA1:A10、Y値がB1:B10から単純指数曲線を作成し、A1:A20の値を計算する例
= ExpCurveModified(A1:A10, B1:B10, , A1:A20)
