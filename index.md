# ⚙️ Excel LAMBDA Utils

Excel LAMBDA関数を使った汎用ユーティリティ集です。  
勤務時間計算、座標計算、曲線作成、文字列操作など、様々な関数を提供します。

- [A simple list in English](docs/index.md)
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
- **[CrossPoint](docs/CrossPoint.md)**       : 線と線が交差する点と、交差が指定した線の内部か判定する

## 🔺 面積/体積計算
- **[TriangleAreaSSS](docs/TriangleArea.md)** : 3辺から三角形の面積を計算
- **[TriangleAreaSAS](docs/TriangleArea.md)** : 2辺とその間の角度から三角形の面積を計算
- **[TriangleAreaASA](docs/TriangleArea.md)** : 1辺とその両端の角度から三角形の面積を計算
- **[SectionVolume](docs/SectionVolume.md)**  : SP断面から体積を計算 (平均断面法、プリスモイダル法、戸田式補正)

## 📈 曲線計算
- **[PowerCurve](docs/PowerCurve.md)**             : べき乗曲線を作成し、結果を計算
- **[ExpCurveSimple](docs/ExpCurveSimple.md)**     : 単純指数曲線を作成し、結果を計算
- **[ExpCurveModified](docs/ExpCurveModified.md)** : 修正指数曲線を作成し、結果を計算
- **[LogisticCurve](docs/LogisticCurve.md)**       : ロジスティック曲線を作成し、結果を計算

## ✂️ 文字列操作
- **[SmartSplit](docs/SmartSplit.md)**: CSV/JSON風文字列を、ダブルクォートとエスケープを考慮して安全に分割
- **[SmartJoin](docs/SmartJoin.md)**  : 範囲から、CSV風文字列を作成。ダブルクォートで囲い、["] は [""] に置き換える
- **[Words](docs/Words.md)**          : スペースで区切られた文字を分割して取得
- **[NumberToColumn](docs/NumberToColumn.md)** : 数字を、エクセルのCOLUMNに対応する文字列に変換する
- **[ColumnToNumber](docs/NumberToColumn.md)** : エクセルのCOLUMNに対応する文字列を、数字に変換する

## 🔍 検索
- **[ClosestIndex](docs/ClosestIndex.md)**: 指定された範囲から、しきい値に一番近い値の、最初のインデックス番号を返す
