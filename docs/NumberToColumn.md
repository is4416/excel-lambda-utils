# ✂️ NumberToColumn ー Excel LAMBDA関数で、数字 <==> 文字を変換

数字をエクセルに対応したCOLUMNの文字に変換します  
逆変換関数も用意しています

**引数**

| 引数   | 型     | 説明             |
| ------ | ------ | ---------------- |
| Num    | Number | 数字             |
| Result | String | COLUMNを表す文字 |

**備考**

- Result は戻り値です。引数としては不要です。

**コード**

```vb
NumberToColumn = LAMBDA(Num, LET(
  L, INT(LOG(Num * 25 + 1, 26)),
  Keys, MAP(SEQUENCE(L), LAMBDA(i, LET(
    T, REDUCE(Num, SEQUENCE(i), LAMBDA(R,j, IF(j = 1, R, INT((R - 1) / 26)))),
    CHAR(65 + MOD(T - 1, 26))
  ))),
  REDUCE("", Keys, LAMBDA(R,Key, Key & R))
))
```

**変数の詳細**

- L: Number, 変換後の文字の長さ
- Keys: String, 結合前の文字

**使用例**

NumberToColumn という名前でブックに登録しているものとします。

```vb
= NumberToColumn(4)
```

出力:
D

スピル対応ができなかったため、連続する値を処理する場合には、MAPでラップしてください。

```vb
= MAP(A1:A10, NumberToColumn)
```

**逆関数**

逆関数として、ColumnToNumber も用意しました  
こちらもスカラー関数であるため、連続する値を処理する場合には、MAPでラップしてください。

**コード**

```vb
ColumnToNumber = LAMBDA(Str, LET(
  L, LEN(Str),
  Codes, MAP(SEQUENCE(L), LAMBDA(i, CODE(MID(Str, i, 1)) - 64)),
  REDUCE(0, SEQUENCE(L), LAMBDA(R, i, R * 26 + INDEX(Codes, i)))
))
```
