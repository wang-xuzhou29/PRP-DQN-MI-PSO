test_data = [(config_depth(1, 7), param_count(1, 180), rsecurity_weight(0, 1)]


```python
=== 执行验证规则函数 (更新后) ===
def execute_validation_rules(a):
    """
    替换原有的配置验证函数
    参数a: 包含3个元素的元组或数组，分别对应config_depth, param_count, security_weight
    返回: 触发的规则编号集合
    """
    config_depth, param_count, security_weight = float(a[0]), int(a[1]), float(a[2])
    triggered = set()
# 创建一个字典来存储b数组的值，用于跟踪哪些规则被触发
b = {}

# 按照原始代码的逻辑进行条件判断
if (config_depth >= 5) != (config_depth >= 5.8):
    b[0] = 1
    triggered.add(1)
if (config_depth >= 5) != (config_depth == 5):
    b[1] = 2
    triggered.add(2)

if config_depth >= 5:
    if (param_count >= 100 and security_weight >= 0.8) != (param_count != 100 and security_weight >= 0.8):
        b[2] = 3
        triggered.add(3)
    if (param_count >= 100 and security_weight >= 0.8) != (param_count == 100 and security_weight >= 0.8):
        b[3] = 4
        triggered.add(4)
    if (param_count >= 100 and security_weight >= 0.8) != (param_count >= 100 and security_weight != 0.8):
        b[4] = 5
        triggered.add(5)
    if (param_count >= 100 and security_weight >= 0.8) != (param_count >= 100 and security_weight == 0.8):
        b[5] = 6
        triggered.add(6)
    if (param_count >= 100 and security_weight >= 0.8) != (param_count >= 100 or security_weight >= 0.8):
        b[6] = 7
        triggered.add(7)

    if param_count >= 100 and security_weight >= 0.8:
        if (security_weight >= 0.95) != (security_weight == 0.95):
            b[7] = 8
            triggered.add(8)
        if (security_weight >= 0.95) != (security_weight != 0.95):
            b[8] = 9
            triggered.add(9)
        if (security_weight >= 0.9) != (security_weight != 0.9):
            b[9] = 10
            triggered.add(10)
        if (security_weight >= 0.9) != (security_weight == 0.9):
            b[10] = 11
            triggered.add(11)

    if (param_count >= 50 and security_weight >= 0.6) != (param_count != 50 and security_weight >= 0.6):
        b[11] = 12
        triggered.add(12)
    if (param_count >= 50 and security_weight >= 0.6) != (param_count == 50 and security_weight >= 0.6):
        b[12] = 13
        triggered.add(13)
    if (param_count >= 50 and security_weight >= 0.6) != (param_count >= 60 and security_weight >= 0.6):
        b[13] = 14
        triggered.add(14)
    if (param_count >= 50 and security_weight >= 0.6) != (param_count >= 50 and security_weight != 0.6):
        b[14] = 15
        triggered.add(15)
    if (param_count >= 50 and security_weight >= 0.6) != (param_count >= 50 and security_weight == 0.6):
        b[15] = 16
        triggered.add(16)
    if (param_count >= 50 and security_weight >= 0.6) != (param_count >= 53 and security_weight >= 0.6):
        b[16] = 17
        triggered.add(17)

    elif param_count >= 50 and security_weight >= 0.6:
        if (security_weight >= 0.8) != (security_weight != 0.8):
            b[17] = 18
            triggered.add(18)
        if (security_weight >= 0.8) != (security_weight == 0.8):
            b[18] = 19
            triggered.add(19)

    if (param_count >= 25) != (param_count != 25):
        b[19] = 20
        triggered.add(20)
    if (param_count >= 25) != (param_count >= 35):
        b[20] = 21
        triggered.add(21)

if (config_depth >= 3) != (config_depth != 3):
    b[21] = 22
    triggered.add(22)
if (config_depth >= 3) != (config_depth >= 3.9):
    b[22] = 23
    triggered.add(23)

elif config_depth >= 3:
    if (param_count >= 75 and security_weight >= 0.7) != (param_count >= 75 or security_weight >= 0.7):
        b[23] = 24
        triggered.add(24)
    if (param_count >= 75 and security_weight >= 0.7) != (param_count == 75 and security_weight >= 0.7):
        b[24] = 25
        triggered.add(25)
    if (param_count >= 75 and security_weight >= 0.7) != (param_count >= 75 and security_weight != 0.7):
        b[25] = 26
        triggered.add(26)
    if (param_count >= 75 and security_weight >= 0.7) != (param_count >= 75 and security_weight == 0.7):
        b[26] = 27
        triggered.add(27)
    if (param_count >= 75 and security_weight >= 0.7) != (param_count != 75 and security_weight >= 0.7):
        b[27] = 28
        triggered.add(28)

    if param_count >= 75 and security_weight >= 0.7:
        if (security_weight >= 0.9) != (security_weight != 0.9):
            b[28] = 29
            triggered.add(29)
        if (security_weight >= 0.9) != (security_weight == 0.9):
            b[29] = 30
            triggered.add(30)
        if (security_weight >= 0.8) != (security_weight != 0.8):
            b[30] = 31
            triggered.add(31)
        if (security_weight >= 0.8) != (security_weight == 0.8):
            b[31] = 32
            triggered.add(32)

    if (param_count >= 40 and security_weight >= 0.5) != (param_count >= 40 or security_weight >= 0.5):
        b[32] = 33
        triggered.add(33)
    if (param_count >= 40 and security_weight >= 0.5) != (param_count >= 40 and security_weight != 0.5):
        b[33] = 34
        triggered.add(34)
    if (param_count >= 40 and security_weight >= 0.5) != (param_count >= 40 and security_weight == 0.5):
        b[34] = 35
        triggered.add(35)
    if (param_count >= 40 and security_weight >= 0.5) != (param_count != 40 and security_weight >= 0.5):
        b[35] = 36
        triggered.add(36)
    if (param_count >= 40 and security_weight >= 0.5) != (param_count == 40 and security_weight >= 0.5):
        b[36] = 37
        triggered.add(37)

if (config_depth >= 2) != (config_depth != 2):
    b[37] = 38
    triggered.add(38)
if (config_depth >= 2) != (config_depth >= 3):
    b[38] = 39
    triggered.add(39)

elif config_depth >= 2:
    if (param_count >= 50 and security_weight >= 0.6) != (param_count != 50 and security_weight >= 0.6):
        b[39] = 40
        triggered.add(40)
    if (param_count >= 50 and security_weight >= 0.6) != (param_count == 50 and security_weight >= 0.6):
        b[40] = 41
        triggered.add(41)
    if (param_count >= 50 and security_weight >= 0.6) != (param_count >= 50 or security_weight >= 0.6):
        b[41] = 42
        triggered.add(42)
    if (param_count >= 50 and security_weight >= 0.6) != (param_count >= 50 and security_weight == 0.6):
        b[42] = 43
        triggered.add(43)
    if (param_count >= 50 and security_weight >= 0.6) != (param_count >= 50 and security_weight != 0.6):
        b[43] = 44
        triggered.add(44)

    if (param_count >= 30) != (param_count != 30):
        b[44] = 45
        triggered.add(45)
    if (param_count >= 30) != (param_count >= 45):
        b[45] = 46
        triggered.add(46)

# 参数数量维度分析
if (param_count >= 150) != (param_count != 150):
    b[46] = 47
    triggered.add(47)
if (param_count >= 150) != (param_count >= 100):
    b[47] = 48
    triggered.add(48)

if param_count >= 150:
    if (config_depth >= 4 and security_weight >= 0.8) != (config_depth >= 4 or security_weight >= 0.8):
        b[48] = 49
        triggered.add(49)
    if (config_depth >= 4 and security_weight >= 0.8) != (config_depth == 4 and security_weight >= 0.8):
        b[49] = 50
        triggered.add(50)
    if (config_depth >= 4 and security_weight >= 0.8) != (config_depth != 4 and security_weight >= 0.8):
        b[50] = 51
        triggered.add(51)
    if (config_depth >= 4 and security_weight >= 0.8) != (config_depth >= 4 and security_weight == 0.8):
        b[51] = 52
        triggered.add(52)
    if (config_depth >= 4 and security_weight >= 0.8) != (config_depth >= 4 and security_weight != 0.8):
        b[52] = 53
        triggered.add(53)

    if (config_depth >= 3 and security_weight >= 0.7) != (config_depth >= 3 and security_weight != 0.7):
        b[53] = 54
        triggered.add(54)
    if (config_depth >= 3 and security_weight >= 0.7) != (config_depth >= 3 and security_weight == 0.7):
        b[54] = 55
        triggered.add(55)
    if (config_depth >= 3 and security_weight >= 0.7) != (config_depth >= 3 or security_weight >= 0.7):
        b[55] = 56
        triggered.add(56)
    if (config_depth >= 3 and security_weight >= 0.7) != (config_depth != 3 and security_weight >= 0.7):
        b[56] = 57
        triggered.add(57)
    if (config_depth >= 3 and security_weight >= 0.7) != (config_depth == 3 and security_weight >= 0.7):
        b[57] = 58
        triggered.add(58)

if (param_count >= 100) != (param_count >= 125):
    b[58] = 59
    triggered.add(59)
if (param_count >= 100) != (param_count >= 200):
    b[59] = 60
    triggered.add(60)

elif param_count >= 100:
    if (config_depth >= 3 and security_weight >= 0.7) != (config_depth >= 5 and security_weight >= 0.81):
        b[60] = 61
        triggered.add(61)
    if (config_depth >= 3 and security_weight >= 0.7) != (config_depth >= 5 and security_weight >= 0.7):
        b[61] = 62
        triggered.add(62)
    if (security_weight >= 0.6) != (security_weight == 0.6):
        b[62] = 63
        triggered.add(63)
    if (security_weight >= 0.6) != (security_weight != 0.6):
        b[63] = 64
        triggered.add(64)

if (param_count >= 50) != (param_count != 50):
    b[64] = 65
    triggered.add(65)
if (param_count >= 50) != (param_count >= 74):
    b[65] = 66
    triggered.add(66)

elif param_count >= 50:
    if (security_weight >= 0.8) != (security_weight != 0.8):
        b[66] = 67
        triggered.add(67)
    if (security_weight >= 0.8) != (security_weight == 0.8):
        b[67] = 68
        triggered.add(68)
    if (security_weight >= 0.6) != (security_weight != 0.6):
        b[68] = 69
        triggered.add(69)
    if (security_weight >= 0.6) != (security_weight == 0.6):
        b[69] = 70
        triggered.add(70)

# 安全权重维度检查
if (security_weight >= 0.9) != (security_weight != 0.9):
    b[70] = 71
    triggered.add(71)
if (security_weight >= 0.9) != (security_weight == 0.9):
    b[71] = 72
    triggered.add(72)

if security_weight >= 0.9:
    if (config_depth >= 4 and param_count >= 80) != (config_depth != 4 and param_count >= 80):
        b[72] = 73
        triggered.add(73)
    if (config_depth >= 4 and param_count >= 80) != (config_depth == 4 and param_count >= 80):
        b[73] = 74
        triggered.add(74)
    if (config_depth >= 4 and param_count >= 80) != (config_depth >= 4 and param_count != 80):
        b[74] = 75
        triggered.add(75)
    if (config_depth >= 4 and param_count >= 80) != (config_depth >= 4 and param_count == 80):
        b[75] = 76
        triggered.add(76)
    if (config_depth >= 4 and param_count >= 80) != (config_depth >= 4 and param_count >= 88):
        b[76] = 77
        triggered.add(77)
    if (config_depth >= 4 and param_count >= 80) != (config_depth >= 4.4 and param_count >= 80):
        b[77] = 78
        triggered.add(78)

    if (config_depth >= 3 and param_count >= 50) != (config_depth >= 3 and param_count != 50):
        b[78] = 79
        triggered.add(79)
    if (config_depth >= 3 and param_count >= 50) != (config_depth >= 3 and param_count == 50):
        b[79] = 80
        triggered.add(80)
    if (config_depth >= 3 and param_count >= 50) != (config_depth != 3 and param_count >= 50):
        b[80] = 81
        triggered.add(81)
    if (config_depth >= 3 and param_count >= 50) != (config_depth == 3 and param_count >= 50):
        b[81] = 82
        triggered.add(82)
    if (config_depth >= 3 and param_count >= 50) != (config_depth >= 3.6 and param_count >= 50):
        b[82] = 83
        triggered.add(83)
    if (config_depth >= 3 and param_count >= 50) != (config_depth >= 3 and param_count >= 66):
        b[83] = 84
        triggered.add(84)

if (security_weight >= 0.7) != (security_weight != 0.7):
    b[84] = 85
    triggered.add(85)
if (security_weight >= 0.7) != (security_weight == 0.7):
    b[85] = 86
    triggered.add(86)

elif security_weight >= 0.7:
    if (config_depth >= 3 and param_count >= 60) != (config_depth >= 3 and param_count >= 66):
        b[86] = 87
        triggered.add(87)
    if (config_depth >= 3 and param_count >= 60) != (config_depth >= 3 and param_count == 60):
        b[87] = 88
        triggered.add(88)
    if (config_depth >= 3 and param_count >= 60) != (config_depth != 3 and param_count >= 60):
        b[88] = 89
        triggered.add(89)
    if (config_depth >= 3 and param_count >= 60) != (config_depth == 3 and param_count >= 60):
        b[89] = 90
        triggered.add(90)
    if (config_depth >= 3 and param_count >= 60) != (config_depth >= 3 and param_count >= 72):
        b[90] = 91
        triggered.add(91)
    if (config_depth >= 3 and param_count >= 60) != (config_depth >= 3.3 and param_count >= 60):
        b[91] = 92
        triggered.add(92)

    if (param_count >= 40) != (param_count != 40):
        b[92] = 93
        triggered.add(93)
    if (param_count >= 40) != (param_count >= 25):
        b[93] = 94
        triggered.add(94)

if (security_weight >= 0.5) != (security_weight != 0.5):
    b[94] = 95
    triggered.add(95)
if (security_weight >= 0.5) != (security_weight == 0.5):
    b[95] = 96
    triggered.add(96)

elif security_weight >= 0.5:
    if (param_count >= 30) != (param_count >= 80):
        b[96] = 97
        triggered.add(97)
    if (param_count >= 30) != (param_count >= 100):
        b[97] = 98
        triggered.add(98)

# Missing parameters detection based on dimensions
if (config_depth < 3) != (config_depth < 4.7):
    b[98] = 99
    triggered.add(99)
if (config_depth < 3) != (config_depth < 5):
    b[99] = 100
    triggered.add(100)
if (param_count < 50) != (param_count < 76):
    b[100] = 101
    triggered.add(101)
if (param_count < 50) != (param_count < 67):
    b[101] = 102
    triggered.add(102)
if (security_weight < 0.6) != (security_weight == 0.63):
    b[102] = 103
    triggered.add(103)
if (security_weight < 0.6) != (security_weight != 0.6):
    b[103] = 104
    triggered.add(104)

return triggered
# === 目标路径组 ===
targetPaths = [
    # A1
    {3, 7, 8, 11, 12, 14, 16, 17, 19, 22, 23, 24, 28, 30, 32, 35, 37, 40, 43, 47, 49, 51, 55, 57, 61, 62, 63, 65, 66, 68, 70, 72, 79, 81, 83, 84, 86, 96, 97, 98, 99, 100, 101, 102, 104},
    # A2
    {3, 7, 8, 11, 12, 14, 16, 17, 19, 24, 28, 30, 32, 35, 37, 40, 43, 47, 51, 52, 55, 58, 61, 62, 63, 65, 66, 68, 70, 72, 75, 79, 82, 84, 86, 96, 97, 98, 99, 100, 101, 102, 104},
    # A3
    {3, 6, 8, 11, 13, 16, 19, 22, 23, 25, 27, 30, 32, 35, 37, 41, 43, 47, 48, 49, 51, 55, 57, 59, 60, 61, 62, 63, 68, 70, 72, 73, 80, 81, 83, 86, 88, 89, 92, 96, 99, 100, 104},
    # A4
    {3, 7, 8, 11, 12, 19, 21, 24, 28, 30, 32, 33, 36, 40, 42, 45, 46, 47, 51, 52, 55, 58, 61, 62, 63, 65, 68, 70, 72, 75, 79, 86, 93, 94, 96, 97, 98, 99, 100, 104},
    # A5
    {1, 3, 7, 8, 11, 13, 16, 19, 24, 28, 30, 32, 35, 37, 41, 43, 47, 50, 52, 55, 58, 63, 66, 68, 70, 72, 75, 80, 82, 84, 86, 87, 90, 91, 96, 97, 98, 101, 102, 104},
    # A6
    {2, 3, 7, 8, 11, 12, 19, 20, 21, 24, 28, 30, 32, 33, 36, 40, 42, 45, 47, 50, 52, 55, 58, 63, 65, 68, 70, 72, 75, 79, 86, 93, 94, 96, 104},
    # A7
    {1, 9, 10, 15, 18, 24, 26, 29, 31, 33, 34, 42, 44, 47, 49, 53, 54, 56, 64, 67, 69, 71, 74, 75, 77, 80, 82, 85, 88, 90, 95, 98, 103},
    # A8
    {5, 7, 9, 10, 15, 18, 22, 24, 26, 29, 31, 33, 34, 38, 42, 44, 47, 60, 64, 67, 69, 71, 73, 81, 85, 89, 95, 103},
    # A9
    {4, 6, 8, 11, 13, 16, 19, 25, 27, 30, 32, 35, 37, 41, 43, 47, 51, 52, 55, 58, 60, 61, 62, 63, 68, 70, 72, 73, 76, 78, 80, 82, 86, 88, 90, 96, 99, 100, 104},
    # A10
    {3, 7, 8, 11, 13, 16, 19, 22, 27, 28, 30, 32, 35, 37, 38, 39, 41, 43, 47, 49, 51, 56, 57, 63, 68, 70, 72, 81, 86, 89, 96, 97, 98, 101, 104}
]
```