# ✂️ SmartSplit ー Excel LAMBDA関数で安全に文字列を分割

カンマ区切りの文字列を、ダブルクォート（`"`）やエスケープ（`""`）を考慮して分割します。  
CSV風の文字列を扱う際に、単純な `TEXTSPLIT` では対応できないケースに使用できます。

**引数**

| 引数   | 型     | 説明                                      |
| ------ | ------ | ----------------------------------------- |
| S      | String | 分割対象の文字列                          |
| Result | String | 配列（スピル）として、各要素を1行ずつ返す |

**備考**

- Result は戻り値です。引数としては不要です。
- `""` によるエスケープ `""` は `"` 1つ分として扱います。
- ダブルクォートで囲まれていないカンマ `,` のみを区切り文字として認識します。
- CSV風（`"A,B","C"`）文字列の分割が可能です。
- 要素の前後に空白がある場合、自動で `TRIM` されます。
- ダブルクォートのない要素も受け付けますが、結果はすべて「文字列」として扱われます。

**コード**

```excel
= LAMBDA(S, LET(
  S_Count, LEN(S),
  S_List, SEQUENCE(S_Count),

  QuotList, MAP(
    S_List,
    LAMBDA(i, LET(
      C, MID(S, i, 1),
      PrevC, IF(i = 1, "", MID(S, i - 1, 1)),
      NextC, IF(i = LEN(S), "", MID(S, i + 1, 1)),
      IF(AND(C = """", PrevC <> """", NextC <> """"), i, 0)
    ))
  ),
  CleanQuotList, FILTER(QuotList, QuotList > 0),

  Q_Count, ROWS(CleanQuotList),
  Q_List, SEQUENCE(CEILING(Q_Count / 2, 1)),

  EscapeAreas_Start, IF(Q_Count = 0, Q_List, MAP(
    Q_List,
    LAMBDA(i, INDEX(CleanQuotList, (i - 1) * 2 + 1))
  )),

  EscapeAreas_End, IF(Q_Count = 0, Q_List, MAP(
    Q_List,
    LAMBDA(i, IF(i * 2 > Q_Count, 0, INDEX(CleanQuotList, i * 2)))
  )),

  DelimiterList, MAP(
    S_List,
    LAMBDA(i, IF(i = S_Count, i, LET(
      c, MID(S, i, 1),
      IF(c = ",", i, 0)
    )))
  ),

  DelimiterFlags, MAP(
    DelimiterList,
    LAMBDA(i, IF(i = 0,
      FALSE,
      SUM((EscapeAreas_End > 0) * (EscapeAreas_Start < i) * (EscapeAreas_End > i)) = 0
    ))
  ),
  CleanDelimiterList, FILTER(DelimiterList, DelimiterFlags),

  TextList, TRIM(MAP(
    SEQUENCE(ROWS(CleanDelimiterList)),
    LAMBDA(i, LET(
      Start, IF(i = 1, 1, INDEX(CleanDelimiterList, i - 1) + 1),
      Size, INDEX(CleanDelimiterList, i) - Start + 1,
      Buf, TRIM(MID(S, Start, Size)),
      Item, IF(i = ROWS(CleanDelimiterList), Buf, MID(Buf, 1, LEN(Buf) - 1)),
      L, IF(LEN(Item) > 0, LEFT(Item, 1), ""),
      R, IF(LEN(Item) > 1, RIGHT(Item, 1), ""),
      SUBSTITUTE(
        IF(AND(L = """", R = """"), MID(Item, 2, LEN(Item) - 2), Item),
        """""",
        """"
      )
    ))
  )),

  TextList
))
```

**変数の詳細**

- S_Count           : Number, S の文字数（LEN(S)）
- S_List            : Range, 1〜S_Count の連番（SEQUENCE(S_Count)）で各文字位置を走査
- QuotList          : Range, 各文字が単独の "（直前も直後も " でない）ならその位置、そうでなければ 0
- CleanQuotList     : Range, QuotList から 0 を除いた実際のクォート位置の一覧
- Q_Count           : Number, クォートの出現数（ROWS(CleanQuotList)）
- Q_List            : Range, 開始・終了ペア処理用のインデックス配列（SEQUENCE(CEILING(Q_Count / 2, 1))）
- EscapeAreas_Start : Range, クォートで囲まれた範囲（エスケープ領域）の開始位置の一覧
- EscapeAreas_End   : Range, クォートで囲まれた範囲（エスケープ領域）の終了位置の一覧
- DelimiterList     : Range, 各文字がカンマならその位置、そうでなければ 0。最終文字位置も含む
- DelimiterFlags    : Range, 各カンマがクォート外にある場合のみ TRUE（内側は FALSE）
- CleanDelimiterList: Range, DelimiterList から FALSE の位置を除いた実際の区切り位置の一覧
- TextList          : Range, 各区切りで抽出した最終的な文字列配列（スピル出力）

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
