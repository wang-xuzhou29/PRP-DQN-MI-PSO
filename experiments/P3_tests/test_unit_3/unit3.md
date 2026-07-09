test_data = [path_depth(1, 15), file_count(1, 110000), access_level(1, 4)]

```python
# === 执行编排规则函数 (更新后) ===

def execute_orchestration_rules(a):
    """
    替换原有的编排规则函数
    参数a: 包含3个元素的元组或数组，分别对应path_depth, file_count, access_level
    返回: 触发的规则编号集合
    """
    path_depth, file_count, access_level = float(a[0]), int(a[1]), float(a[2])
    triggered = set()
# 创建一个字典来存储b数组的值，用于跟踪哪些规则被触发
b = {}

# 按照原始代码的逻辑进行条件判断
if (path_depth >= 10) != (path_depth >= 12):
    b[0] = 1
    triggered.add(1)
if (path_depth >= 10) != (path_depth == 10):
    b[1] = 2
    triggered.add(2)
if (path_depth >= 10) != (path_depth >= 8):
    b[2] = 3
    triggered.add(3)

if path_depth >= 10:
    if (file_count >= 10000 and access_level >= 3) != (file_count >= 10000 and access_level >= 4.3):
        b[3] = 4
        triggered.add(4)
    if (file_count >= 10000 and access_level >= 3) != (file_count >= 10000 and access_level == 3):
        b[4] = 5
        triggered.add(5)
    if (file_count >= 10000 and access_level >= 3) != (file_count >= 10000 and access_level >= 3.2):
        b[5] = 6
        triggered.add(6)
    if (file_count >= 10000 and access_level >= 3) != (file_count == 10000 and access_level >= 3):
        b[6] = 7
        triggered.add(7)
    if (file_count >= 10000 and access_level >= 3) != (file_count >= 10000 and access_level >= 4):
        b[7] = 8
        triggered.add(8)
    if (file_count >= 10000 and access_level >= 3) != (file_count >= 13400 and access_level >= 3.9):
        b[8] = 9
        triggered.add(9)
    if (file_count >= 10000 and access_level >= 3) != (file_count != 10000 and access_level >= 3):
        b[9] = 10
        triggered.add(10)
    if (file_count >= 10000 and access_level >= 3) != (file_count >= 10000 and access_level >= 5):
        b[10] = 11
        triggered.add(11)

    if file_count >= 10000 and access_level >= 3:
        if (access_level == 3 and file_count >= 50000) != (access_level != 3 and file_count >= 50000):
            b[11] = 12
            triggered.add(12)
        if (access_level == 3 and file_count >= 50000) != (access_level >= 3 and file_count >= 50000):
            b[12] = 13
            triggered.add(13)
        if (access_level == 3 and file_count >= 50000) != (access_level <= 3 and file_count >= 50000):
            b[13] = 14
            triggered.add(14)
        if (access_level == 3 and file_count >= 50000) != (access_level == 3 or file_count >= 50000):
            b[14] = 15
            triggered.add(15)

        if (file_count >= 25000) != (file_count != 25000):
            b[15] = 16
            triggered.add(16)
        if (file_count >= 25000) != (file_count >= 30000):
            b[16] = 17
            triggered.add(17)
        if (file_count >= 25000) != (file_count >= 29000):
            b[17] = 18
            triggered.add(18)

    if (file_count >= 5000 and access_level >= 2) != (file_count >= 5000 and access_level >= 2.5):
        b[18] = 19
        triggered.add(19)
    if (file_count >= 5000 and access_level >= 2) != (file_count >= 5000 or access_level >= 3.5):
        b[19] = 20
        triggered.add(20)
    if (file_count >= 5000 and access_level >= 2) != (file_count >= 5000 and access_level != 2):
        b[20] = 21
        triggered.add(21)
    if (file_count >= 5000 and access_level >= 2) != (file_count != 5000 and access_level >= 2):
        b[21] = 22
        triggered.add(22)
    if (file_count >= 5000 and access_level >= 2) != (file_count >= 5000 and access_level >= 3):
        b[22] = 23
        triggered.add(23)
    if (file_count >= 5000 and access_level >= 2) != (file_count >= 5000 and access_level >= 3.3):
        b[23] = 24
        triggered.add(24)

    if (file_count >= 1000) != (file_count >= 43000):
        b[24] = 25
        triggered.add(25)
    if (file_count >= 1000) != (file_count >= 25000):
        b[25] = 26
        triggered.add(26)
    if (file_count >= 1000) != (file_count >= 50000):
        b[26] = 27
        triggered.add(27)

if (path_depth >= 6) != (path_depth >= 7):
    b[27] = 28
    triggered.add(28)
if (path_depth >= 6) != (path_depth != 6):
    b[28] = 29
    triggered.add(29)
if (path_depth >= 6) != (path_depth >= 7):
    b[29] = 30
    triggered.add(30)

elif path_depth >= 6:
    if (file_count >= 20000 and access_level >= 3) != (file_count >= 20000 and access_level >= 4):
        b[30] = 31
        triggered.add(31)
    if (file_count >= 20000 and access_level >= 3) != (file_count >= 20000 and access_level != 3):
        b[31] = 32
        triggered.add(32)
    if (file_count >= 20000 and access_level >= 3) != (file_count >= 20000 and access_level == 3):
        b[32] = 33
        triggered.add(33)
    if (file_count >= 20000 and access_level >= 3) != (file_count == 20000 and access_level >= 3):
        b[33] = 34
        triggered.add(34)
    if (file_count >= 20000 and access_level >= 3) != (file_count != 20000 and access_level >= 3):
        b[34] = 35
        triggered.add(35)

    if (file_count >= 8000 and access_level >= 2) != (file_count >= 54000 and access_level >= 2):
        b[35] = 36
        triggered.add(36)
    if (file_count >= 8000 and access_level >= 2) != (file_count != 8000 and access_level >= 2):
        b[36] = 37
        triggered.add(37)
    if (file_count >= 8000 and access_level >= 2) != (file_count >= 45000 and access_level >= 2):
        b[37] = 38
        triggered.add(38)
    if (file_count >= 8000 and access_level >= 2) != (file_count >= 8000 and access_level != 2):
        b[38] = 39
        triggered.add(39)
    if (file_count >= 8000 and access_level >= 2) != (file_count >= 8000 and access_level >= 3.3):
        b[39] = 40
        triggered.add(40)
    if (file_count >= 8000 and access_level >= 2) != (file_count >= 8000 and access_level >= 2.5):
        b[40] = 41
        triggered.add(41)
    if (file_count >= 8000 and access_level >= 2) != (file_count >= 80500 and access_level >= 2):
        b[41] = 42
        triggered.add(42)

    elif file_count >= 8000 and access_level >= 2:
        if (access_level == 3) != (access_level <= 3):
            b[42] = 43
            triggered.add(43)
        if (access_level == 3) != (access_level >= 3):
            b[43] = 44
            triggered.add(44)

    if (file_count >= 2000) != (file_count != 2000):
        b[44] = 45
        triggered.add(45)
    if (file_count >= 2000) != (file_count >= 62000):
        b[45] = 46
        triggered.add(46)

if (path_depth >= 3) != (path_depth != 3):
    b[46] = 47
    triggered.add(47)
if (path_depth >= 3) != (path_depth >= 3.5):
    b[47] = 48
    triggered.add(48)

elif path_depth >= 3:
    if (file_count >= 15000 and access_level >= 2) != (file_count != 15000 and access_level >= 2):
        b[48] = 49
        triggered.add(49)
    if (file_count >= 15000 and access_level >= 2) != (file_count == 15000 and access_level >= 2):
        b[49] = 50
        triggered.add(50)
    if (file_count >= 15000 and access_level >= 2) != (file_count >= 15000 and access_level != 2):
        b[50] = 51
        triggered.add(51)
    if (file_count >= 15000 and access_level >= 2) != (file_count >= 15000 and access_level == 2):
        b[51] = 52
        triggered.add(52)
    if (file_count >= 15000 and access_level >= 2) != (file_count >= 15000 and access_level >= 2.5):
        b[52] = 53
        triggered.add(53)
    if (file_count >= 15000 and access_level >= 2) != (file_count >= 17500 and access_level >= 2):
        b[53] = 54
        triggered.add(54)
    if (file_count >= 15000 and access_level >= 2) != (file_count >= 33330 and access_level >= 2):
        b[54] = 55
        triggered.add(55)

    if file_count >= 15000 and access_level >= 2:
        if (access_level == 3) != (access_level >= 3):
            b[55] = 56
            triggered.add(56)
        if (access_level == 3) != (access_level <= 3):
            b[56] = 57
            triggered.add(57)

    if (file_count >= 5000) != (file_count != 5000):
        b[57] = 58
        triggered.add(58)
    if (file_count >= 5000) != (file_count >= 55000):
        b[58] = 59
        triggered.add(59)

# 文件数量维度处理
if (file_count >= 100000) != (file_count >= 90000):
    b[59] = 60
    triggered.add(60)
if (file_count >= 100000) != (file_count == 100000):
    b[60] = 61
    triggered.add(61)

if file_count >= 100000:
    if (path_depth >= 8 and access_level >= 3) != (path_depth >= 8 and access_level != 3):
        b[61] = 62
        triggered.add(62)
    if (path_depth >= 8 and access_level >= 3) != (path_depth >= 8 and access_level == 3):
        b[62] = 63
        triggered.add(63)
    if (path_depth >= 8 and access_level >= 3) != (path_depth == 8 and access_level >= 3):
        b[63] = 64
        triggered.add(64)
    if (path_depth >= 8 and access_level >= 3) != (path_depth != 8 and access_level >= 3):
        b[64] = 65
        triggered.add(65)
    if (path_depth >= 8 and access_level >= 3) != (path_depth >= 9 and access_level >= 3):
        b[65] = 66
        triggered.add(66)
    if (path_depth >= 8 and access_level >= 3) != (path_depth >= 11 and access_level >= 3):
        b[66] = 67
        triggered.add(67)
    if (path_depth >= 8 and access_level >= 3) != (path_depth >= 15 and access_level >= 3):
        b[67] = 68
        triggered.add(68)

    if (path_depth >= 5 and access_level >= 2) != (path_depth != 5 and access_level >= 2):
        b[68] = 69
        triggered.add(69)
    if (path_depth >= 5 and access_level >= 2) != (path_depth == 5 and access_level >= 2):
        b[69] = 70
        triggered.add(70)
    if (path_depth >= 5 and access_level >= 2) != (path_depth >= 5 and access_level >= 2.9):
        b[70] = 71
        triggered.add(71)
    if (path_depth >= 5 and access_level >= 2) != (path_depth >= 5 and access_level != 2):
        b[71] = 72
        triggered.add(72)
    if (path_depth >= 5 and access_level >= 2) != (path_depth >= 5 and access_level == 2):
        b[72] = 73
        triggered.add(73)
    if (path_depth >= 5 and access_level >= 2) != (path_depth >= 5 and access_level >= 2.5):
        b[73] = 74
        triggered.add(74)
    if (path_depth >= 5 and access_level >= 2) != (path_depth >= 7 and access_level >= 2):
        b[74] = 75
        triggered.add(75)

if (file_count >= 50000) != (file_count != 50000):
    b[75] = 76
    triggered.add(76)
if (file_count >= 50000) != (file_count == 50000):
    b[76] = 77
    triggered.add(77)

elif file_count >= 50000:
    if (access_level >= 3) != (access_level != 3):
        b[77] = 78
        triggered.add(78)
    if (access_level >= 3) != (access_level == 3):
        b[78] = 79
        triggered.add(79)

if (file_count >= 10000) != (file_count != 10000):
    b[79] = 80
    triggered.add(80)
if (file_count >= 10000) != (file_count == 10000):
    b[80] = 81
    triggered.add(81)

# 访问级别维度的扫描序列规划
if (access_level == 3) != (access_level <= 3):
    b[81] = 82
    triggered.add(82)
if (access_level == 3) != (access_level >= 3):
    b[82] = 83
    triggered.add(83)

if access_level == 3:
    if (path_depth >= 8 and file_count >= 20000) != (path_depth != 8 and file_count >= 20000):
        b[83] = 84
        triggered.add(84)
    if (path_depth >= 8 and file_count >= 20000) != (path_depth == 8 and file_count >= 20000):
        b[84] = 85
        triggered.add(85)
    if (path_depth >= 8 and file_count >= 20000) != (path_depth >= 7 and file_count >= 20000):
        b[85] = 86
        triggered.add(86)
    if (path_depth >= 8 and file_count >= 20000) != (path_depth >= 8 and file_count != 20000):
        b[86] = 87
        triggered.add(87)
    if (path_depth >= 8 and file_count >= 20000) != (path_depth >= 8 and file_count == 20000):
        b[87] = 88
        triggered.add(88)
    if (path_depth >= 8 and file_count >= 20000) != (path_depth >= 11 and file_count == 20000):
        b[88] = 89
        triggered.add(89)
    if (path_depth >= 8 and file_count >= 20000) != (path_depth >= 8 and file_count == 25500):
        b[89] = 90
        triggered.add(90)

    if (path_depth >= 5) != (path_depth != 5):
        b[90] = 91
        triggered.add(91)
    if (path_depth >= 5) != (path_depth >= 3.5):
        b[91] = 92
        triggered.add(92)
    if (path_depth >= 5) != (path_depth >= 6):
        b[92] = 93
        triggered.add(93)

if (access_level == 2) != (access_level >= 2):
    b[93] = 94
    triggered.add(94)
if (access_level == 2) != (access_level <= 2):
    b[94] = 95
    triggered.add(95)

return triggered
# === 目标路径组 ===

targetPaths = [
    # A1
    {1, 4, 6, 7, 8, 9, 11, 15, 17, 24, 25, 27, 31, 32, 34, 36, 38, 40, 42, 46, 50, 52, 55, 59, 62, 64, 67, 68, 70, 73, 76, 78, 81, 85, 88, 89, 90, 94},
    # A2
    {1, 2, 4, 6, 7, 8, 9, 11, 15, 17, 24, 25, 27, 31, 32, 34, 36, 38, 40, 42, 46, 50, 52, 55, 59, 62, 64, 68, 70, 73, 76, 78, 81, 85, 88, 89, 90, 94},
    # A3
    {3, 4, 5, 7, 11, 17, 25, 27, 33, 34, 36, 38, 42, 44, 46, 50, 52, 55, 56, 59, 63, 65, 66, 67, 68, 70, 73, 76, 79, 81, 83, 84, 88, 89, 90, 94},
    # A4
    {17, 19, 21, 23, 24, 25, 27, 28, 29, 30, 32, 36, 38, 39, 40, 41, 42, 43, 46, 50, 51, 53, 55, 57, 59, 70, 71, 72, 74, 75, 76, 78, 81, 82, 84},
    # A5
    {19, 21, 23, 24, 25, 27, 29, 32, 36, 38, 39, 40, 41, 42, 43, 46, 50, 51, 53, 55, 57, 59, 69, 71, 72, 74, 75, 76, 78, 81, 82, 84, 91, 93},
    # A6
    {16, 19, 21, 23, 24, 25, 26, 27, 29, 36, 38, 39, 40, 41, 42, 43, 46, 50, 51, 53, 55, 57, 59, 69, 71, 72, 74, 75, 76, 78, 81, 82, 91, 93},
    # A7
    {3, 4, 6, 7, 8, 9, 11, 15, 16, 24, 25, 26, 27, 35, 36, 38, 40, 42, 46, 50, 52, 54, 55, 59, 62, 64, 67, 68, 70, 73, 76, 78, 81, 87, 94},
    # A8
    {1, 12, 14, 15, 19, 21, 23, 24, 32, 39, 40, 41, 42, 43, 46, 50, 51, 53, 57, 59, 62, 70, 71, 72, 74, 77, 78, 81, 82, 85, 88, 89, 90},
    # A9
    {3, 4, 5, 7, 11, 12, 13, 15, 33, 34, 42, 44, 46, 50, 52, 56, 59, 63, 65, 66, 67, 68, 70, 73, 77, 79, 81, 83, 84, 88, 89, 90, 94},
    # A10
    {17, 19, 21, 23, 24, 25, 27, 32, 36, 38, 39, 40, 41, 42, 43, 46, 50, 51, 53, 55, 57, 59, 70, 71, 72, 74, 76, 78, 81, 82, 84, 86},
    # A11
    {17, 19, 21, 23, 24, 25, 27, 29, 32, 36, 38, 39, 40, 41, 42, 43, 46, 47, 48, 50, 51, 53, 55, 57, 59, 69, 76, 78, 81, 82, 84, 91},
    # A12
    {16, 19, 21, 23, 24, 25, 26, 27, 28, 29, 30, 36, 38, 39, 40, 41, 42, 43, 46, 49, 57, 59, 70, 71, 72, 74, 75, 76, 78, 80, 82},
    # A13
    {16, 19, 21, 23, 24, 25, 26, 27, 29, 36, 38, 39, 40, 41, 42, 43, 46, 50, 51, 53, 55, 57, 59, 69, 76, 78, 81, 82, 91, 92},
    # A14
    {1, 10, 16, 20, 22, 25, 26, 27, 35, 37, 44, 45, 49, 56, 58, 63, 64, 67, 68, 70, 73, 76, 79, 80, 83, 87, 94},
    # A15
    {1, 2, 17, 20, 21, 25, 27, 32, 39, 43, 46, 51, 57, 59, 62, 72, 76, 78, 81, 82, 85, 88, 89, 90, 95},
    # A16
    {4, 5, 7, 11, 17, 18, 25, 27, 29, 33, 34, 36, 38, 42, 44, 46, 50, 52, 55, 56, 59, 65, 69, 73, 75, 76, 79, 81, 83, 84, 91, 93, 94},
    # A17
    {4, 6, 7, 8, 9, 11, 12, 24, 31, 32, 34, 40, 50, 52, 61, 65, 70, 73, 77, 78, 81, 84, 86, 94},
    # A18
    {1, 2, 4, 6, 7, 8, 9, 11, 12, 24, 31, 32, 34, 40, 50, 52, 60, 62, 64, 68, 70, 73, 77, 78, 81, 85, 88, 89, 90, 94}
    ]
```

