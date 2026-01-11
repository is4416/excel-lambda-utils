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

```vb
= LAMBDA(S, LET(
  S_Count, LEN(S),
  S_List , SEQUENCE(S_Count),

  QuotList, FILTER(S_List, MID(S, S_List, 1) = """"),

  Q_Count, ROWS(QuotList),
  Q_List , SEQUENCE(CEILING(Q_Count / 2, 1)),

  EscapeAreas_Start, IF(Q_Count = 0, Q_List, MAP(
    Q_List,
    LAMBDA(i, INDEX(QuotList, (i - 1) * 2 + 1))
  )),
  EscapeAreas_End, IF(Q_Count = 0, Q_List, MAP(
    Q_List,
    LAMBDA(i, IF(i * 2 > Q_Count, 0, INDEX(QuotList, i * 2)))
  )),

  DelimiterList, MAP(
    S_List,
    LAMBDA(i, IF(i = S_Count,
      i,
      IF(MID(S, i, 1) = ",", i, 0)
    ))
  ),

  DelimiterFlags, MAP(
    DelimiterList,
    LAMBDA(i, IF(i = 0,
      FALSE,
      SUM((EscapeAreas_End > 0) * (EscapeAreas_Start < i) * (EscapeAreas_End > i)) = 0
    ))
  ),
  CleanDelimiterList, FILTER(DelimiterList, DelimiterFlags),

  TRIM(MAP(
    SEQUENCE(ROWS(CleanDelimiterList)),
    LAMBDA(i, LET(
      Start, IF(i = 1, 1, INDEX(CleanDelimiterList, i - 1) + 1),
      Size , INDEX(CleanDelimiterList, i) - Start + 1,
      Buf  , TRIM(MID(S, Start, Size)),
      Item , IF(i = ROWS(CleanDelimiterList), Buf, MID(Buf, 1, LEN(Buf) - 1)),
      L    , IF(LEN(Item) > 0, LEFT(Item, 1), ""),
      R    , IF(LEN(Item) > 1, RIGHT(Item, 1), ""),
      SUBSTITUTE(
        IF(AND(L = """", R = """"), MID(Item, 2, LEN(Item) - 2), Item),
        """""",
        """"
      )
    ))
  ))
))
```

**変数の詳細**

- S_Count           : Number, S の文字数（LEN(S)）
- S_List            : Range, 1〜S_Count の連番（SEQUENCE(S_Count)）で各文字位置を走査
- QuotList          : Range, 各文字が " ならその位置、そうでなければ 0（すべての " を拾う）
- Q_Count           : Number, QuotList に含まれる " の数（ROWS(QuotList)）
- Q_List            : Range, 開始・終了ペア処理用のインデックス配列（SEQUENCE(CEILING(Q_Count / 2, 1))）
- EscapeAreas_Start : Range, 偶数・奇数ペアで決まる、引用（エスケープ）領域の開始位置一覧
- EscapeAreas_End   : Range, 偶数・奇数ペアで決まる、引用（エスケープ）領域の終了位置一覧
- DelimiterList     : Range, 各文字がカンマならその位置、そうでなければ 0。最終文字位置も含む
- DelimiterFlags    : Range, 各カンマが 引用外にある場合のみ TRUE、引用内は FALSE
- CleanDelimiterList: Range, DelimiterList から FALSE の位置を除いた、実際の区切り（カンマ）位置の一覧

**使用例**

SmartSplit という名前でブックに登録しているものとします。

```vb
= SmartSplit(A1)
```

| A列の内容         | 結果（スピル出力） |
| ----------------- | ------------------ |
| `["A","B","C"]`   | A<br>B<br>C        |
| `["A,B","C"]`     | A,B<br>C           |

**応用例**

CSVセルの中に `"A,B,C"` のような複雑な構造がある場合に、`SmartSplit` で安全に分解して再利用できます。  
JSON構文のような括弧・クォートを含む場合にも対応可能です。
