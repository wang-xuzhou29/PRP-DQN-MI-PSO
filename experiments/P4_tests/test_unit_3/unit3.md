​    

```python
# === 验证规则执行函数 ===
def execute_Tr(dx, dy, dz):
    """
    执行验证规则，返回触发的规则集合
    """
    b = set()

    # 使用全局的质量参数
    global quality_x, quality_y, quality_z
    quality_x = dx
    quality_y = dy
    quality_z = dz

    # 异常类型1：质量参数乘积异常
    if ((quality_y * quality_z) / (quality_x + 1) > 80) != ((quality_y * quality_y * quality_z) / (quality_x + 1) > 80):
        b.add(1)
    if ((quality_y * quality_z) / (quality_x + 1) > 80) != ((quality_y * quality_z * quality_z) / (quality_x + 1) > 80):
        b.add(2)
    if ((quality_y * quality_z) / (quality_x + 1) > 80) != ((quality_y * quality_x * quality_z) / (quality_x + 1) > 80):
        b.add(3)
    if ((quality_y * quality_z) / (quality_x + 1) > 80) != ((quality_y * quality_z) / (quality_x + 1) > 60):
        b.add(4)
    if ((quality_y * quality_z) / (quality_x + 1) > 80) != ((quality_y * quality_z) / (quality_x + 10) > 80):
        b.add(5)
    if ((quality_y * quality_z) / (quality_x + 1) > 80) != ((quality_y * quality_z) / (quality_x + 13) > 80):
        b.add(6)
    if ((quality_y * quality_z) / (quality_x + 1) > 80) != ((quality_y * quality_z * 5) / (quality_x + 1) > 80):
        b.add(7)
    if ((quality_y * quality_z) / (quality_x + 1) > 80) != ((quality_y * quality_z * 2) / (quality_x + 1) > 80):
        b.add(8)
    if ((quality_y * quality_z) / (quality_x + 1) > 80) != ((quality_y * quality_z) / (quality_x + 1) > 40):
        b.add(9)
    if ((quality_y * quality_z) / (quality_x + 1) > 80) != ((quality_y * quality_x) / (quality_x + 1) > 80):
        b.add(10)
    if ((quality_y * quality_z) / (quality_x + 1) > 80) != ((quality_y * quality_y) / (quality_x + 1) > 80):
        b.add(11)
    if ((quality_y * quality_z) / (quality_x + 1) > 80) != ((quality_z * quality_z) / (quality_x + 1) > 80):
        b.add(12)

    # 异常类型2：质量差值异常
    if ((quality_z - quality_x) < 0.4 * quality_y) != ((quality_z - quality_x) < 0.3 * quality_y):
        b.add(13)
    if ((quality_z - quality_x) < 0.4 * quality_y) != ((quality_z - quality_x) < 0.5 * quality_y):
        b.add(14)
    if ((quality_z - quality_x) < 0.4 * quality_y) != ((quality_z - quality_x) < 0.4 * quality_z):
        b.add(15)
    if ((quality_z - quality_x) < 0.4 * quality_y) != ((quality_z - quality_x) < 0.4 * quality_x):
        b.add(16)
    if ((quality_z - quality_x) < 0.4 * quality_y) != ((quality_z * 1.1 - quality_x) < 0.4 * quality_y):
        b.add(17)
    if ((quality_z - quality_x) < 0.4 * quality_y) != ((quality_z * 2 - quality_x) < 0.4 * quality_y):
        b.add(18)
    if ((quality_z - quality_x) < 0.4 * quality_y) != ((quality_z * quality_z - quality_x) < 0.4 * quality_y):
        b.add(19)
    if ((quality_z - quality_x) < 0.4 * quality_y) != ((quality_z * quality_x - quality_x) < 0.4 * quality_y):
        b.add(20)
    if ((quality_z - quality_x) < 0.4 * quality_y) != ((quality_z * quality_y - quality_x) < 0.4 * quality_y):
        b.add(21)
    if ((quality_z - quality_x) < 0.4 * quality_y) != ((quality_z * 1.5 - quality_x) < 0.4 * quality_y):
        b.add(22)

    # 异常类型3：质量立方关系
    if ((quality_x ** 3 + quality_y ** 3) < quality_z ** 2) != ((quality_x ** 2 + quality_y ** 3) < quality_z ** 2):
        b.add(23)
    if ((quality_x ** 3 + quality_y ** 3) < quality_z ** 2) != ((quality_x ** 3 + quality_y ** 2) < quality_z ** 2):
        b.add(24)
    if ((quality_x ** 3 + quality_y ** 3) < quality_z ** 2) != ((quality_x ** 3 + quality_y ** 3) < quality_z ** 1):
        b.add(25)
    if ((quality_x ** 3 + quality_y ** 3) < quality_z ** 2) != ((quality_x ** 3 + quality_y ** 3) < quality_z ** 3):
        b.add(26)
    if ((quality_x ** 3 + quality_y ** 3) < quality_z ** 2) != ((quality_x ** 3 + quality_y ** 3) < quality_z ** 4):
        b.add(27)
    if ((quality_x ** 3 + quality_y ** 3) < quality_z ** 2) != ((quality_x ** 3 + quality_y ** 4) < quality_z ** 2):
        b.add(28)
    if ((quality_x ** 3 + quality_y ** 3) < quality_z ** 2) != ((quality_x ** 3 + quality_x ** 3) < quality_z ** 2):
        b.add(29)
    if ((quality_x ** 3 + quality_y ** 3) < quality_z ** 2) != ((quality_x ** 1 + quality_y ** 3) < quality_z ** 2):
        b.add(30)
    if ((quality_x ** 3 + quality_y ** 3) < quality_z ** 2) != ((quality_x ** 3 + quality_y ** 1) < quality_z ** 2):
        b.add(31)
    if ((quality_x ** 3 + quality_y ** 3) < quality_z ** 2) != (
            ((quality_x ** 3) * 2 + quality_y ** 3) < quality_z ** 2):
        b.add(32)
    if ((quality_x ** 3 + quality_y ** 3) < quality_z ** 2) != (
            (quality_x ** 3 + (quality_y ** 3) * 2) < quality_z ** 2):
        b.add(33)
    if ((quality_x ** 3 + quality_y ** 3) < quality_z ** 2) != ((quality_y ** 3 + quality_y ** 3) < quality_z ** 2):
        b.add(34)
    if ((quality_x ** 3 + quality_y ** 3) < quality_z ** 2) != (
            (quality_x ** 3 + quality_y ** 3) < (quality_z ** 2) * 2):
        b.add(35)
    if ((quality_x ** 3 + quality_y ** 3) < quality_z ** 2) != (
            (quality_x ** 3 + quality_y ** 3) < (quality_x ** 2) * 2):
        b.add(36)
    if ((quality_x ** 3 + quality_y ** 3) < quality_z ** 2) != (
            (quality_x ** 3 + quality_y ** 3) < (quality_y ** 2) * 2):
        b.add(37)

    # 异常类型6：质量同步性检查
    if (abs(quality_x % 1 - quality_y % 1) < 0.1 and abs(quality_y % 1 - quality_z % 1) < 0.1) != (
            abs(quality_x % 2 - quality_y % 1) < 0.1 and abs(quality_y % 1 - quality_z % 1) < 0.1):
        b.add(38)
    if (abs(quality_x % 1 - quality_y % 1) < 0.1 and abs(quality_y % 1 - quality_z % 1) < 0.1) != (
            abs(quality_x % 3 - quality_y % 1) < 0.1 and abs(quality_y % 1 - quality_z % 1) < 0.1):
        b.add(39)
    if (abs(quality_x % 1 - quality_y % 1) < 0.1 and abs(quality_y % 1 - quality_z % 1) < 0.1) != (
            abs(quality_x % 1 - quality_y % 2) < 0.1 and abs(quality_y % 1 - quality_z % 1) < 0.1):
        b.add(40)
    if (abs(quality_x % 1 - quality_y % 1) < 0.1 and abs(quality_y % 1 - quality_z % 1) < 0.1) != (
            abs(quality_x % 1 - quality_y % 3) < 0.1 and abs(quality_y % 1 - quality_z % 1) < 0.1):
        b.add(41)
    if (abs(quality_x % 1 - quality_y % 1) < 0.1 and abs(quality_y % 1 - quality_z % 1) < 0.1) != (
            abs(quality_x % 1 - quality_y % 5) < 0.1 and abs(quality_y % 1 - quality_z % 1) < 0.1):
        b.add(42)
    if (abs(quality_x % 1 - quality_y % 1) < 0.1 and abs(quality_y % 1 - quality_z % 1) < 0.1) != (
            abs(quality_x % 5 - quality_y % 1) < 0.1 and abs(quality_y % 1 - quality_z % 1) < 0.1):
        b.add(43)
    if (abs(quality_x % 1 - quality_y % 1) < 0.1 and abs(quality_y % 1 - quality_z % 1) < 0.1) != (
            abs(quality_x % 1 - quality_y % 1) < 0.1 and abs(quality_y % 2 - quality_z % 1) < 0.1):
        b.add(44)
    if (abs(quality_x % 1 - quality_y % 1) < 0.1 and abs(quality_y % 1 - quality_z % 1) < 0.1) != (
            abs(quality_x % 1 - quality_y % 1) < 0.1 and abs(quality_y % 3 - quality_z % 1) < 0.1):
        b.add(45)
    if (abs(quality_x % 1 - quality_y % 1) < 0.1 and abs(quality_y % 1 - quality_z % 1) < 0.1) != (
            abs(quality_x % 1 - quality_y % 1) < 0.1 and abs(quality_y % 5 - quality_z % 1) < 0.1):
        b.add(46)
    if (abs(quality_x % 1 - quality_y % 1) < 0.1 and abs(quality_y % 1 - quality_z % 1) < 0.1) != (
            abs(quality_x % 1 - quality_y % 1) < 0.1 and abs(quality_y % 1 - quality_z % 2) < 0.1):
        b.add(47)
    if (abs(quality_x % 1 - quality_y % 1) < 0.1 and abs(quality_y % 1 - quality_z % 1) < 0.1) != (
            abs(quality_x % 1 - quality_y % 1) < 0.1 and abs(quality_y % 1 - quality_z % 3) < 0.1):
        b.add(48)
    if (abs(quality_x % 1 - quality_y % 1) < 0.1 and abs(quality_y % 1 - quality_z % 1) < 0.1) != (
            abs(quality_x % 1 - quality_y % 1) < 0.1 and abs(quality_y % 1 - quality_z % 5) < 0.1):
        b.add(49)
    if (abs(quality_x % 1 - quality_y % 1) < 0.1 and abs(quality_y % 1 - quality_z % 1) < 0.1) != (
            abs(quality_x % 1 - quality_y % 1) < 0.1 and abs(quality_y % 4 - quality_z % 1) < 0.1):
        b.add(50)
    if (abs(quality_x % 1 - quality_y % 1) < 0.1 and abs(quality_y % 1 - quality_z % 1) < 0.1) != (
            abs(quality_x % 1 - quality_y % 1) < 0.1 and abs(quality_y % 1 - quality_z % 4) < 0.1):
        b.add(51)
    if (abs(quality_x % 1 - quality_y % 1) < 0.1 and abs(quality_y % 1 - quality_z % 1) < 0.1) != (
            abs(quality_x % 4 - quality_y % 1) < 0.1 and abs(quality_y % 1 - quality_z % 1) < 0.1):
        b.add(52)
    if (abs(quality_x % 1 - quality_y % 1) < 0.1 and abs(quality_y % 1 - quality_z % 1) < 0.1) != (
            abs(quality_x % 1 - quality_y % 6) < 0.1 and abs(quality_y % 1 - quality_z % 1) < 0.1):
        b.add(53)
    if (abs(quality_x % 1 - quality_y % 1) < 0.1 and abs(quality_y % 1 - quality_z % 1) < 0.1) != (
            abs(quality_x % 6 - quality_y % 1) < 0.1 and abs((quality_y * 2) % 1 - quality_z % 1) < 0.1):
        b.add(54)
    if (abs(quality_x % 1 - quality_y % 1) < 0.1 and abs(quality_y % 1 - quality_z % 1) < 0.1) != (
            abs(quality_x % 1 - quality_y % 1) < 0.1 and abs(quality_y % 6 - (quality_z * 2) % 1) < 0.1):
        b.add(55)

    # 其他复杂检查逻辑
    if (quality_x * quality_y * quality_z > 500000 and (quality_x + quality_y + quality_z) / 3 < 85) != (
            quality_x * quality_y * quality_z > 500000 and (quality_x + quality_y + quality_z) / 2 < 85):
        b.add(56)
    if (quality_x * quality_y * quality_z > 500000 and (quality_x + quality_y + quality_z) / 3 < 85) != (
            quality_x * quality_y * quality_z > 500000 and (quality_x + quality_y + quality_z) / 4 < 85):
        b.add(57)
    if (quality_x * quality_y * quality_z > 500000 and (quality_x + quality_y + quality_z) / 3 < 85) != (
            quality_x * quality_y * quality_z > 500000 and (quality_x * 2 + quality_y + quality_z) / 3 < 85):
        b.add(58)
    if (quality_x * quality_y * quality_z > 500000 and (quality_x + quality_y + quality_z) / 3 < 85) != (
            quality_x * quality_y * quality_z > 500000 and (quality_x + quality_y * 2 + quality_z) / 3 < 85):
        b.add(59)
    if (quality_x * quality_y * quality_z > 500000 and (quality_x + quality_y + quality_z) / 3 < 85) != (
            quality_x * quality_y * quality_z > 500000 and (quality_x + quality_y - quality_z) / 3 < 85):
        b.add(60)
    if (quality_x * quality_y * quality_z > 500000 and (quality_x + quality_y + quality_z) / 3 < 85) != (
            quality_x * quality_y * quality_z > 500000 and (quality_x + quality_y * quality_z * 2) / 3 < 85):
        b.add(61)
    if (quality_x * quality_y * quality_z > 500000 and (quality_x + quality_y + quality_z) / 3 < 85) != (
            quality_x * quality_y * quality_z > 500000 and (quality_x + quality_y + quality_x) / 3 < 85):
        b.add(62)
    if (quality_x * quality_y * quality_z > 500000 and (quality_x + quality_y + quality_z) / 3 < 85) != (
            quality_x * quality_y * quality_z > 500000 and (quality_x + quality_y + quality_y) / 3 < 85):
        b.add(63)
    if (quality_x * quality_y * quality_z > 500000 and (quality_x + quality_y + quality_z) / 3 < 85) != (
            quality_x * quality_y * quality_z > 600000 and (quality_x + quality_y + quality_z) / 3 < 85):
        b.add(64)
    if (quality_x * quality_y * quality_z > 500000 and (quality_x + quality_y + quality_z) / 3 < 85) != (
            quality_x * quality_y * quality_z * 2 > 500000 and (quality_x + quality_y + quality_z) / 3 < 85):
        b.add(65)
    if (quality_x * quality_y * quality_z > 500000 and (quality_x + quality_y + quality_z) / 3 < 85) != (
            quality_y * quality_y * quality_z > 500000 and (quality_x + quality_y + quality_z) / 3 < 85):
        b.add(66)
    if (quality_x * quality_y * quality_z > 500000 and (quality_x + quality_y + quality_z) / 3 < 85) != (
            quality_z * quality_y * quality_z > 500000 and (quality_x + quality_y + quality_z) / 3 < 85):
        b.add(67)
    if (quality_x * quality_y * quality_z > 500000 and (quality_x + quality_y + quality_z) / 3 < 85) != (
            quality_x * quality_x * quality_z > 500000 and (quality_x + quality_y + quality_z) / 3 < 85):
        b.add(68)
    if (quality_x * quality_y * quality_z > 500000 and (quality_x + quality_y + quality_z) / 3 < 85) != (
            quality_x * quality_y * quality_x > 500000 and (quality_x + quality_y + quality_z) / 3 < 85):
        b.add(69)
    if (quality_x * quality_y * quality_z > 500000 and (quality_x + quality_y + quality_z) / 3 < 85) != (
            quality_x * quality_z * quality_z > 500000 and (quality_x + quality_y + quality_z) / 3 < 85):
        b.add(70)
    if (quality_x * quality_y * quality_z > 500000 and (quality_x + quality_y + quality_z) / 3 < 85) != (
            quality_x * quality_y * quality_y > 500000 and (quality_x + quality_y + quality_z) / 3 < 85):
        b.add(71)

    return b
# === 目标路径定义 ===
target_paths = [
    [1, 2, 3, 4, 7, 8, 9, 10, 11, 19, 20, 21, 27, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 49, 50, 51, 52, 53, 54, 55,
     56, 58, 59, 61, 62, 63, 64, 67, 68, 70],
    [1, 2, 3, 4, 7, 8, 9, 12, 18, 19, 20, 21, 22, 27, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53,
     54, 55, 56, 58, 59, 61, 64, 66, 71],
    [1, 2, 3, 4, 7, 8, 9, 12, 13, 17, 18, 19, 20, 21, 22, 26, 27, 40, 41, 42, 44, 45, 46, 48, 49, 50, 51, 52, 53, 55,
     56, 58, 59, 61, 64, 66, 69, 71],
    [5, 6, 10, 11, 13, 17, 18, 19, 20, 21, 22, 26, 27, 38, 39, 40, 42, 43, 44, 46, 49, 50, 52, 53, 54, 55, 56, 58, 59,
     61, 64, 66, 69, 71],
    [1, 2, 3, 7, 12, 17, 18, 19, 20, 21, 22, 26, 27, 30, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 49, 50, 51, 52, 53, 54,
     55, 65, 68, 70],
    [1, 2, 3, 7, 8, 9, 12, 17, 18, 19, 20, 21, 22, 26, 27, 38, 39, 40, 41, 42, 44, 45, 46, 47, 48, 50, 51, 52, 53, 54,
     55, 57, 60, 63],
    [1, 2, 3, 4, 7, 8, 9, 12, 23, 24, 26, 27, 30, 31, 34, 35, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51,
     52, 53, 54, 55],
    [16, 18, 19, 20, 21, 22, 27, 38, 40, 41, 42, 43, 44, 45, 46, 48, 49, 50, 51, 52, 53, 54, 55, 56, 58, 59, 61, 63, 64,
     68, 69, 70],
    [1, 2, 3, 4, 7, 8, 9, 12, 14, 15, 16, 26, 27, 38, 40, 41, 42, 43, 44, 45, 46, 48, 49, 50, 51, 52, 53, 54, 55, 57,
     60, 62, 63],
    [5, 6, 10, 13, 15, 16, 18, 19, 20, 21, 22, 27, 29, 31, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53,
     54, 55],
    [1, 2, 3, 7, 8, 9, 12, 25, 28, 29, 32, 33, 36, 37, 38, 40, 41, 42, 43, 44, 45, 46, 47, 48, 50, 51, 52, 53, 54, 55,
     70]
]


```

