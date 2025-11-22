# ⚙️ Excel Lambda Utils

このリポジトリは、Excel LAMBDA関数を使ったユーティリティ集です。  
現在、以下の関数を公開しています。

[時間の計算]
- **OverlapTime**     : 時間の重複計算
- **TimeToDecimal**   : 時刻 (TimeDate) を時間 (Float) に変換
- **DecimalToTime**   : 時間 (Float) を時刻 (TimeToDate) に変換
- **MonthsBetween**   : 月数をカウント (月末締め or 翌月の前日締め)
- **LastDay**         : 指定日の最終日を返す
- **DiffDaysTime**    : 日+時刻から、日+時刻を控除 (1日当たりの時間を指定)

[座標計算]
- **DistancePoint**   : 2次元座標間の距離
- **PolygonArea**     : 座標による面積計算
- **FootPoint**       : 線と点の垂直に交わる交点の座標と、点との距離を返す

[曲線計算]
- **PowerCurve**      : べき乗曲線を作成し、結果を計算
- **ExpCurveSimple**  : 単純指数曲線を作成し、結果を計算
- **ExpCurveModified**: 修正指数曲線を作成し、結果を計算
- **LogisticCurve**   : ロジスティック曲線を作成し、結果を計算

[文字列操作]
- **SmartSplit**      : CSV/JSON風文字列を、ダブルクォートとエスケープを考慮して安全に分割
- **SmartJoin**       : 範囲から、CSV風文字列を作成。ダブルクォートで囲い、["] は [""] に置き換える

## 📂 構成

```
excel-lambda-utils/
├── index.md
├── LICENSE
├── README.md
├── TODO.md
├── docs/
│   ├── index.md
│   │
│   ├── OverlapTime.md
│   ├── TimeToDecimal.md
│   ├── DecimalToTime.md
│   ├── MonthsBetween.md
│   ├── LastDay.md
│   ├── DiffDaysTime.md
│   │
│   ├── DistancePoint.md
│   ├── PolygonArea.md
│   ├── FootPoint.md
│   │
│   ├── PowerCurve.md
│   ├── ExpCurveSimple.md
│   ├── ExpCurveModified.md
│   ├── LogisticCurve.md
│   │
│   ├── SmartSplit.md
│   └── SmartJoin.md
│
├── examples/
│   ├── excel-lambda-utils.ods
│   │
│   ├── OverlapTime-example.txt
│   ├── TimeToDecimal-example.txt
│   ├── DecimalToTIme-example.txt
│   ├── MonthsBetween-example.txt
│   ├── LastDay-example.txt
│   ├── DiffDaysTime-example.txt
│   │
│   ├── DistancePoint-example.txt
│   ├── PolygonArea-example.txt
│   ├── FootPoint-example.txt
│   │
│   ├── PowerCurve-example.txt
│   ├── ExpCurveSimple-example.txt
│   ├── ExpCurveModified-example.txt
│   ├── LogisticCurve-example.txt
│   │
│   ├── SmartSplit-example.txt
│   └── SmartJoin-example.txt
│
└── src/
     ├── OverlapTime.txt
     ├── TimeToDecimal.txt
     ├── DecimalToTime.txt
     ├── MonthsBetween.txt
     ├── LastDay.txt
     ├── DiffDaysTime.txt
     │
     ├── DistancePoint.txt
     ├── PolygonArea.txt
     ├── FootPoint.txt
     │
     ├── PowerCurve.txt
     ├── ExpCurveSimple.txt
     ├── ExpCurveModified.txt
     ├── LogisticCurve.txt
     │
     ├── SmartSplit.txt
     └── SmartJoin.txt
```

## 📖 ドキュメント

それぞれの関数の詳しい解説は docs 内のファイルを参照してください。

## 📝 使用例

それぞれの関数の使用例は examples 内のファイルを参照してください。

## ⚖️ ライセンス

このリポジトリは [MIT](LICENSE) ライセンスの下で公開されています。
