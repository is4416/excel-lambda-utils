# ⚙️ Excel LAMBDA Utils

Excel LAMBDA関数を使った汎用ユーティリティ集です。  
勤務時間計算、座標計算、曲線作成、文字列操作など、様々な関数を提供します。

- ライセンス: [MIT](LICENSE)

## ⏱️ 時間の計算
- **[OverlapTime](docs/OverlapTime.md)**     : 時間の重複計算
- **[TimeToDecimal](docs/TimeToDecimal.md)** : 時刻 (TimeDate) を時間 (Float) に変換
- **[DecimalToTime](docs/DecimalToTime.md)** : 時間 (Float) を時刻 (TimeToDate) に変換
- **[MonthsBetween](docs/MonthsBetween.md)** : 月数をカウント (月末締め or 翌月の前日締め)
- **[LastDay](docs/LastDay.md)**             : 指定日の最終日を返す
- **[DiffDaysTime](docs/DiffDaysTime.md)**   : 日+時刻から、日+時刻を控除 (1日当たりの時間を指定)

## 📍 座標計算
- **[DistancePoint](docs/DistancePoint.md)** : 2次元座標間の距離
- **[PolygonArea](docs/PolygonArea.md)**     : 座標による面積計算
- **[FootPoint](docs/FootPoint.md)**         : 線と点の垂直に交わる交点の座標と、点との距離を返す

## 📈 曲線計算
- **[PowerCurve](docs/PowerCurve.md)**             : べき乗曲線を作成し、結果を計算
- **[ExpCurveSimple](docs/ExpCurveSimple.md)**     : 単純指数曲線を作成し、結果を計算
- **[ExpCurveModified](docs/ExpCurveModified.md)** : 修正指数曲線を作成し、結果を計算
- **[LogisticCurve](docs/LogisticCurve.md)**       : ロジスティック曲線を作成し、結果を計算

## ✂️ 文字列操作
- **[SmartSplit](docs/SmartSplit.md)** : CSV/JSON風文字列を、ダブルクォートとエスケープを考慮して安全に分割
- **[SmartJoin](docs/SmartJoin.md)**   : 範囲から、CSV風文字列を作成。ダブルクォートで囲い、["] は [""] に置き換える
