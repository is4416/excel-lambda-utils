# 🔺 TriangleArea — Excel LAMBDA関数で三角形の面積を計算

辺の長さや角度から、三角形の面積を計算します。

- 3辺がわかっている場合、TriangleAreaSSS
- 2辺とその間の角度がわかっている場合、TriangleAreaSAS
- 1辺とその両端の角度がわかっている場合、TriangleAreaASA

を使用してください

---

## TriangleAreaSSS

3辺がわかっている場合

**引数**

| 引数    | 型     | 説明               |
| ------- | ------ | ------------------ |
| A, B, C | Number | 辺の長さ           |
| Result  | Number | 三角形の面積を返す |

**備考**

- Result は戻り値です。引数としては不要です。
- 三角形が成り立たない場合、NAを返します

**コード**

```vb
= LAMBDA(A, B, C, LET(
  S, (A + B + C) / 2,
  IF((A < B + C) * (B < A + C) * (C < A + B),
    SQRT(S * (S - A) * (S - B) * (S - C)),
    NA()
  )
))
```

**変数の詳細**

- S: Number, (A + B + C) / 2

**使用例**

TriangleAreaSSS という名前で、ブックに登録しているものとします
> スピルにも対応しています

3辺 (A, B, C) から三角形の面積を求める

```vb
= TriangleAreaSSS(A, B, C)
```

---

## TriangleAreaSAS

2辺とその間の角度 (degrees) がわかっている場合

**引数**

| 引数   | 型               | 説明               |
| ------ | ---------------- | ------------------ |
| A      | Number           | 辺の長さ           |
| R      | Number (degrees) | 間の角度 (degrees) |
| B      | Number           | 辺の長さ           |
| Result | Number           | 三角形の面積を返す |

**備考**

- Result は戻り値です。引数としては不要です。
- 三角形が成り立たない場合、NAを返します
- 角度は、度数 (degrees) で入力します

**コード**

```vb
= LAMBDA(A, R, B, IF(
  (A <= 0) + (B <= 0) + (R >= 180),
  NA(),
  A * B * SIN(RADIANS(R)) / 2
))
```
**使用例**

TriangleAreaSAS という名前で、ブックに登録しているものとします
> スピルにも対応しています

2辺 (A, B) とその間の角度 (R) から三角形の面積を求める

```vb
= TriangleAreaSAS(A, R, B)
```

---

## TriangleAreaASA

1辺と両端の角度 (degrees) がわかっている場合

**引数**

| 引数   | 型               | 説明                 |
| ------ | ---------------- | -------------------- |
| A      | Number           | 1辺の長さ            |
| LR, RR | Number (degrees) | 両端の角度 (degrees) |
| Result | Number           | 三角形の面積を返す   |

**備考**

- Result は戻り値です。引数としては不要です。
- 三角形が成り立たない場合、NAを返します
- 角度は、度数 (degrees) で入力します

**コード**
```vb
= LAMBDA(A, LR, RR, IF(
  (A <= 0) + (LR <= 0) + (RR <= 0) + (LR + RR >= 180),
  NA(),
  A^2 * SIN(RADIANS(LR)) * SIN(RADIANS(RR)) / (2 * SIN(RADIANS(LR + RR)))
))
```
**使用例**

TriangleAreaASA という名前で、ブックに登録しているものとします
> スピルにも対応しています

1辺 (A) と両端の角度 (LR, RR) から三角形の面積を求める

```vb
= TriangleAreaASA(A, LR, RR)
```
