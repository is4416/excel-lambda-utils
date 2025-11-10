# 🕒 Excel Lambda Utils

このリポジトリは、Excel LAMBDA関数を使ったユーティリティ集です。  
現在、以下の関数を公開しています。

- **OverlapTime**     : 時間の重複計算
- **DistancePoint**   : 2次元座標間の距離
- **PolygonArea**     : 座標による面積計算
- **PowerCurve**      : べき乗曲線を作成し、結果を計算
- **ExpCurveSimple**  : 単純指数曲線を作成し、結果を計算
- **ExpCurveModified**: 修正指数曲線を作成し、結果を計算
- **LogisticCurve**   : ロジスティック曲線を作成し、結果を計算
- **SmartSplit**      : CSV/JSON風文字列を、ダブルクォートとエスケープを考慮して安全に分割
- **SmartJoin**       : 範囲から、CSV風文字列を作成。ダブルクォートで囲い、["] は [""] に置き換える

## 📂 構成

```
excel-lambda-utils/
├── LICENSE
├── README.md
├── TODO.md
├── docs/
│   ├── OverlapTime.md
│   ├── DistancePoint.md
│   ├── PolygonArea.md
│   ├── PowerCurve.md
│   ├── ExpCurveSimple.md
│   ├── ExpCurveModified.md
│   ├── LogisticCurve.md
│   ├── SmartSplit.md
│   └── SmartJoin.md
├── examples/
│   ├── OverlapTime-example.txt
│   ├── DistancePoint-example.txt
│   ├── PolygonArea-example.txt
│   ├── PowerCurve-example.txt
│   ├── ExpCurveSimple-example.txt
│   ├── ExpCurveModified-example.txt
│   ├── LogisticCurve-example.txt
│   ├── SmartSplit-example.txt
│   └── SmartJoin-example.txt
└── src/
     ├── OverlapTime.txt
     ├── DistancePoint.txt
     ├── PolygonArea.txt
     ├── PowerCurve.txt
     ├── ExpCurveSimple.txt
     ├── ExpCurveModified.txt
     ├── LogisticCurve.txt
     ├── SmartSplit.txt
     └── SmartJoin.txt
```

## 📖 ドキュメント

それぞれの関数の詳しい解説は docs 内のファイルを参照してください。

## 📝 使用例

それぞれの関数の使用例は examples 内のファイルを参照してください。

## ⚖️ ライセンス

このリポジトリは [MIT](LICENSE) ライセンスの下で公開されています。

