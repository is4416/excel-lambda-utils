# MonthsBetween 使用例

# 月末締めで、月数をカウントする例
= MonthsBetween(A1, B1,,)

# 契約日締め 'DAY(StartDate)' で、スピルでまとめて計算する例
= MonthsBetween(A1:A10, A1:A10, FALSE,)

# 15日締めで計算する例
= MonthsBetween(A1:A10, B1:B10, FALSE, 15)
