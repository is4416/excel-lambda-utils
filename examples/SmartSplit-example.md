# SmartSplit 使用例

## 説明
クォートで囲まれた要素を考慮して、文字列を分割します。  
カンマ区切りのCSV風やJSON風の配列文字列を安全に分割できます。

## 例1
["A","B","C"]  
= SmartSplit(A1)

出力:
A  
B  
C

## 例2
["A,B","C"]  
= SmartSplit(A2)

出力:
A,B  
C

## 例3
{"X", "Y", "Z"}  
= SmartSplit(A3)

出力:
X  
Y  
Z
