# ✂️ SmartSplit ー Excel LAMBDA関数で安全に文字列を分割

カンマ区切りの文字列を、ダブルクォート（`"`）やエスケープ（`\`）を考慮して分割します。  
CSV風・JSON風の文字列を扱う際に、単純な `TEXTSPLIT` では対応できないケースに使用できます。

**引数**

| 引数   | 型     | 説明                                      |
| ------ | ------ | ----------------------------------------- |
| Text   | String | 分割対象の文字列                          |
| Result | String | 配列（スピル）として、各要素を1行ずつ返す |


**備考**

- Result は戻り値です。引数としては不要です。
- `"\""` によるエスケープ `"\""` は `"\""` 1つ分として扱います。
- ダブルクォートで囲まれていないカンマ `,` のみを区切り文字として認識します。
- JSON配列（`["A","B","C"]`）や、CSV風（`"A,B","C"`）文字列の分割が可能です。
- 要素の前後に空白がある場合、自動で `TRIM` されます。
- ダブルクォートのない要素も受け付けますが、結果はすべて「文字列」として扱われます。

**コード**

```excel
= LAMBDA(Text, LET(
  Buf       , TRIM(Text),
  OuterLeft , LEFT(Buf, 1),
  OuterRight, RIGHT(Buf, 1),
  S, IF(OR(
    AND(OuterLeft = "{", OuterRight = "}"),
    AND(OuterLeft = "[", OuterRight = "]")
  ), MID(Buf, 2, LEN(Buf) - 2), Buf),
  NumberList, SEQUENCE(LEN(S)),
  Result    , REDUCE(
    { {}, "", False },
    NumberList,
    LAMBDA(acc, i, LET(
      List   , INDEX(acc, 1),
      Buf    , INDEX(acc, 2),
      Flag   , INDEX(acc, 3),
      C      , MID(S, i, 1),
      Esc    , IF(i > 1, MID(S, i - 1, 1) = "\", False),
      NewFlag, IF(AND(C = """", NOT(Esc)), NOT(Flag), Flag),
      NewList, IF(OR(AND(C = ",", NewFlag = False), i = LEN(S)),
        VSTACK(List, Buf),
        List
      ),
      NewBuf , IF(OR(C = ",", AND(C = """", NOT(NewFlag))),
        "",
        IF(Esc,
          LEFT(Buf, MAX(0, LEN(Buf) - 1)) & C,
          Buf & C
        )
      ),
      { NewList, NewBuf, NewFlag }
    ))
  ),
  INDEX(Result, 1)
))
```

**変数の詳細**

- Buf       : String, Textを 'TRIM' した文字
- OuterLeft : String, 左端の文字
- OuterRight: String, 右端の文字
- S         : String, Bufの外側に `{}[]` があれば削除した文字
- NumberList: Range, S の各文字インデックス（1〜LEN(S)）
- Result    : Array, `REDUCE` による最終結果 `リスト, バッファ, フラグ`


**使用例**

SmartSplit という名前でブックに登録しているものとします。

```excel
= SmartSplit(A1)
```

| A列の内容         | 結果（スピル出力） |
| ----------------- | ------------------ |
| `["A","B","C"]`   | A<br>B<br>C        |
| `["A,B","C"]`     | A,B<br>C           |
| `{"X", "Y", "Z"}` | X<br>Y<br>Z        |

**応用例**

CSVセルの中に `"A,B,C"` のような複雑な構造がある場合に、`SmartSplit` で安全に分解して再利用できます。  
JSON構文のような括弧・クォートを含む場合にも対応可能です。
