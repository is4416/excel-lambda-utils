# OverlapTime 使用例

# 単一の時間帯での使用例
= OverlapTime(A1, B1, TIMEVALUE("08:30"), TIMEVALUE("17:15"))

# 複数行にスピルしてまとめて計算する例
= OverlapTime(A1:A10, B1:B10, TIMEVALUE("08:30"), TIMEVALUE("17:15"))

# 別の時間帯での計算例
= OverlapTime(C1:C10, D1:D10, TIMEVALUE("07:00"), TIMEVALUE("19:00"))

# セル表示形式を [h]:mm にすると時間表示できる
# 例: 2:45 → 2時間45分

