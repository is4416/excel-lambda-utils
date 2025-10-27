# 🕒 Excel Lambda Utils

このリポジトリは、Excel LAMBDA関数を使ったユーティリティ集です。  
現在、時刻の重複計算用に `OverlapTime` 関数を公開しています。

## 📂 構成

```
excel-lambda-utils/
├── LICENSE                  ← MITライセンス
├── docs/
│   └── OverlapTime.md       ← 関数の詳細解説
├── examples/
│   └── OverlapTime-example.txt ← 使用例
└── src/
     └── OverlapTime.txt      ← LAMBDA関数本体
```

## 📖 ドキュメント

`OverlapTime` 関数の詳しい解説は [docs/OverlapTime.md](docs/OverlapTime.md) を参照してください。

## 📝 使用例

OverlapTime という名前でブックに登録しているものとします。

```excel
= OverlapTime(A1:A10, B1:B10, TIMEVALUE("08:30"), TIMEVALUE("17:15"))
```

※ スピルにも対応しています。

## ⚖️ ライセンス

このリポジトリは [MIT](LICENSE) ライセンスの下で公開されています。

