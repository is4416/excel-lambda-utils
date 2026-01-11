# 📐 SectionVolume ー Excel LAMBDA関数でSP断面から体積を計算

SP断面から、体積を計算します  
平均断面法またはプリスモイダル法による体積計算及び戸田式補正を行います

> **戸田式補正**は、現地実態にあわせて任意の補正値 (0.1 - 0.2程度) を乗じるものなので、計算で得られるものではありません

**引数**

| 引数        | 型      | 説明                             |
| ----------- | ------- | -------------------------------- |
| SPRange     | Range   | SP配列 (最低2行)                 |
| ARange      | Range   | 断面積配列 (最低2行)             |
| UniformSpan | Boolean | Prismoidal法で計算 (省略可能)    |
| Alpha       | Number  | 戸田式補正 (省略可能)            |
| Result      | VSTACK(Number, Number)| SPと区間体積の配列 |

**備考**

- Result は戻り値です。引数としては不要です。
- SPRange と ARange の行数は一致させる必要があります
- プリスモイダル法による計算を行う場合、偶数行を中間点として計算します
- 総行数が偶数の場合、中間点が取得できなかった最後の区間は、平均断面法により計算されます
- UniformSpan を省略した場合、SPが等間隔に並んでいればプリスモイダル法、そうでなければ平均断面法となります
- 測量誤差等によりSP距離に誤差があった場合でも、UniformSpan を TRUE に設定するとプリスモイダル法により計算します
- Alpha を0以外の数値に設定した場合、計算方法の如何に関わらず、結果に補正を行います

**コード**

```vb
= LAMBDA(SPRange, ARange, UniformSpan, Alpha, LET(
  US, IF(ISOMITTED(UniformSpan),
    LET(
      Span, INDEX(SPRange, 2) - INDEX(SPRange, 1),
      AND(MAP(
        SEQUENCE(ROWS(SPRange) - 2),
        LAMBDA(i, Span = INDEX(SPRange, i + 2) - INDEX(SPRange, i + 1))
      ))
    ),
    UniformSpan
  ),

  Alp, IF(ISOMITTED(Alpha), 0, Alpha),

  N, ROWS(SPRange),
  IndexList, SEQUENCE(IF(US,
    INT((N - 1) / 2),
    N - 1
  )),

  L, IF(US,
    MAP(
      IndexList,
      LAMBDA(i, INDEX(SPRange, (i - 1) * 2 + 3) - INDEX(SPRange, (i - 1) * 2 + 1))
    ),
    MAP(
      IndexList,
      LAMBDA(i, INDEX(SPRange, i + 1) - INDEX(SPRange, i))
    )
  ),

  SPList, IF(US,
    MAP(
      IndexList,
      LAMBDA(i, INDEX(SPRange, (i - 1) * 2 + 1))
    ),
    DROP(SPRange, - 1)
  ),

  A, IF(US,
    MAP(
      IndexList,
      LAMBDA(i, INDEX(ARange, (i - 1) * 2 + 1))
    ),
    DROP(ARange, - 1)
  ),

  AA, IF(US,
    MAP(
      IndexList,
      LAMBDA(i, INDEX(ARange, (i - 1) * 2 + 3))
    ),
    DROP(ARange, 1)
  ),

  M, IF(US,
    MAP(
      IndexList,
      LAMBDA(i, INDEX(ARange, (i - 1) * 2 + 2))
    ),
    (A + AA) / 2
  ),

  Volumes, (A + 4 * M + AA) / 6 * L + Alp * (AA - A) * L,

  ExFlag  , US * (MOD(N - 1, 2) = 1),
  ExSPList, INDEX(SPRange, N - 1),
  ExA     , INDEX(ARange, N - 1),
  ExAA    , INDEX(ARange, N),
  ExL     , INDEX(SPRange, N) - INDEX(SPRange, N - 1),
  ExVolume, (ExA + ExAA) / 2 * ExL,

  HSTACK(
    IF(ExFlag,
      VSTACK(SPList, ExSPList),
      SPList
    ),
    IF(ExFlag,
      VSTACK(Volumes, ExVolume),
      Volumes
    )
  )
))
```

**変数の詳細**

- US       : Boolean, Prismoidal法を使用するか判定
- Alp      : Number , 補正値 (0なら補正なし)
- N        : Number , SPRangeの総行数
- IndexList: Range  , 計算に使用する区間のインデックス
- L        : Range  , 区間長
- SPList   : Range  , 区間開始のSP値
- A        : Range  , 区間開始の断面積
- AA       : Range  , 区間終了の断面積
- M        : Range  , 中間点断面積 (Prismoidal用)
- Volumes  : Range  , 各区間の体積 (補正含む)
- ExFlag   : Boolean, 区間の最後で、平均断面法の補完が必要か判定
- ExSPList : Number , 最後の区間開始SP値
- ExA      : Number , 最後の区間開始断面積
- ExAA     : Number , 最後の区間終了断面積
- ExL      : Number , 最後の区間長
- ExVolume : Number , 最後の区間体積 (平均断面法)

**使用例**

SectionVolume という名前で、ブックに登録しているものとします  

| 列  | 内容   | 備考               |
| --- | ------ | ------------------ |
| A   | SP     | 最低2行もしくは3行 |
| B   | 断面積 | 最低2行もしくは3行 |

---

*平均断面法、補正なし*

`= SectionVolume(A, B, FALSE,)`

*平均断面法、補正あり (0.2)*

`= SectionVolume(A, B, FALSE, 0.2)`

*プリスモイダル法、補正なし*

偶数行を中間点として扱うため、最低3行以上が必要です  
総行数が偶数の場合、中間点が取得できなかった最後の区間は、平均断面法により計算されます

`= SectionVolume(A, B,,)`

プリスモイダル法を使用するためには、SPが等間隔で並んでいる必要があります  
測量誤差等を許容してプリスモイダル法により計算するためには、フラグを TRUE に設定してください

`= SectionVolume(A, B, TRUE,)`

補正を行う場合には、オプションに0以外を指定します

`= SectionVolume(A, B, TRUE, 0.2)`

**備考**

SectionVolume の返値は `VSTACK(HSTACK(SP, 体積))` となります  
体積の合計だけ取得したい場合は `= SUM( DROP(SectionVolume(A,B,,),,1) )` などとしてください  
※ DROP で体積列のみを取り出しています
