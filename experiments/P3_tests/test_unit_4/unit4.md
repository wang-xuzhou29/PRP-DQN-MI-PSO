

```python
# === 执行威胁分析规则函数 (更新后) ===

def execute_threat_analysis_rules(a):
    """
    替换原有的威胁分析规则函数
    参数a: 包含3个元素的元组或数组，分别对应threat_count, severity_level, confidence_score
    返回: 触发的规则编号集合
    """
    threat_count, severity_level, confidence_score = int(a[0]), float(a[1]), float(a[2])
    triggered = set()
# 创建一个字典来存储b数组的值，用于跟踪哪些规则被触发
b = {}

# 按照原始代码的逻辑进行条件判断
if (threat_count >= 100) != (threat_count >= 110):
    b[0] = 1
    triggered.add(1)
if (threat_count >= 100) != (threat_count == 100):
    b[1] = 2
    triggered.add(2)
if (threat_count >= 100) != (threat_count >= 75):
    b[2] = 3
    triggered.add(3)

if threat_count >= 100:
    if (severity_level >= 8.0 and confidence_score >= 90) != (severity_level >= 8.0 and confidence_score != 90):
        b[3] = 4
        triggered.add(4)
    if (severity_level >= 8.0 and confidence_score >= 90) != (severity_level >= 8.0 and confidence_score >= 95):
        b[4] = 5
        triggered.add(5)
    if (severity_level >= 8.0 and confidence_score >= 90) != (severity_level != 8.0 and confidence_score >= 90):
        b[5] = 6
        triggered.add(6)
    if (severity_level >= 8.0 and confidence_score >= 90) != (severity_level >= 9.5 and confidence_score >= 90):
        b[6] = 7
        triggered.add(7)
    if (severity_level >= 8.0 and confidence_score >= 90) != (severity_level >= 9.3 and confidence_score >= 90):
        b[7] = 8
        triggered.add(8)
    if (severity_level >= 8.0 and confidence_score >= 90) != (severity_level >= 8.0 and confidence_score >= 88):
        b[8] = 9
        triggered.add(9)
    if (severity_level >= 8.0 and confidence_score >= 90) != (severity_level >= 8.0 and confidence_score >= 95):
        b[9] = 10
        triggered.add(10)
    if (severity_level >= 8.0 and confidence_score >= 90) != (severity_level >= 8.3 and confidence_score >= 90):
        b[10] = 11
        triggered.add(11)

    if severity_level >= 8.0 and confidence_score >= 90:
        if (confidence_score >= 95) != (confidence_score >= 93):
            b[11] = 12
            triggered.add(12)
        if (confidence_score >= 95) != (confidence_score == 95):
            b[12] = 13
            triggered.add(13)
        if (confidence_score >= 95) != (confidence_score >= 97):
            b[13] = 14
            triggered.add(14)

        if (severity_level >= 9.0) != (severity_level >= 9.1):
            b[14] = 15
            triggered.add(15)
        if (severity_level >= 9.0) != (severity_level == 9.0):
            b[15] = 16
            triggered.add(16)
        if (severity_level >= 9.0) != (severity_level >= 8.0):
            b[16] = 17
            triggered.add(17)

    if (severity_level >= 6.0 and confidence_score >= 80) != (severity_level != 6.0 and confidence_score >= 80):
        b[17] = 18
        triggered.add(18)
    if (severity_level >= 6.0 and confidence_score >= 80) != (severity_level == 6.0 and confidence_score >= 80):
        b[18] = 19
        triggered.add(19)
    if (severity_level >= 6.0 and confidence_score >= 80) != (severity_level >= 6.0 and confidence_score != 80):
        b[19] = 20
        triggered.add(20)
    if (severity_level >= 6.0 and confidence_score >= 80) != (severity_level >= 6.0 and confidence_score == 80):
        b[20] = 21
        triggered.add(21)
    if (severity_level >= 6.0 and confidence_score >= 80) != (severity_level >= 6.0 and confidence_score >= 82):
        b[21] = 22
        triggered.add(22)
    if (severity_level >= 6.0 and confidence_score >= 80) != (severity_level >= 6.0 and confidence_score >= 88):
        b[22] = 23
        triggered.add(23)
    if (severity_level >= 6.0 and confidence_score >= 80) != (severity_level >= 7.7 and confidence_score >= 82):
        b[23] = 24
        triggered.add(24)

    if (severity_level >= 4.0) != (severity_level != 4.0):
        b[24] = 25
        triggered.add(25)
    if (severity_level >= 4.0) != (severity_level >= 7.0):
        b[25] = 26
        triggered.add(26)
    if (severity_level >= 4.0) != (severity_level >= 5.0):
        b[26] = 27
        triggered.add(27)

if (threat_count >= 50) != (threat_count >= 75):
    b[27] = 28
    triggered.add(28)
if (threat_count >= 50) != (threat_count >= 59):
    b[28] = 29
    triggered.add(29)
if (threat_count >= 50) != (threat_count >= 55):
    b[29] = 30
    triggered.add(30)

elif threat_count >= 50:
    if (severity_level >= 7.0 and confidence_score >= 85) != (severity_level != 7.0 and confidence_score >= 85):
        b[30] = 31
        triggered.add(31)
    if (severity_level >= 7.0 and confidence_score >= 85) != (severity_level == 7.0 and confidence_score >= 85):
        b[31] = 32
        triggered.add(32)
    if (severity_level >= 7.0 and confidence_score >= 85) != (severity_level >= 7.0 and confidence_score != 85):
        b[32] = 33
        triggered.add(33)
    if (severity_level >= 7.0 and confidence_score >= 85) != (severity_level >= 7.0 and confidence_score == 85):
        b[33] = 34
        triggered.add(34)
    if (severity_level >= 7.0 and confidence_score >= 85) != (severity_level >= 7.0 and confidence_score >= 88):
        b[34] = 35
        triggered.add(35)

    if severity_level >= 7.0 and confidence_score >= 85:
        if (confidence_score >= 95) != (confidence_score >= 85):
            b[35] = 36
            triggered.add(36)
        if (confidence_score >= 95) != (confidence_score == 95):
            b[36] = 37
            triggered.add(37)

    if (severity_level >= 5.0 and confidence_score >= 70) != (severity_level >= 5.0 and confidence_score != 70):
        b[37] = 38
        triggered.add(38)
    if (severity_level >= 5.0 and confidence_score >= 70) != (severity_level >= 5.0 and confidence_score == 70):
        b[38] = 39
        triggered.add(39)
    if (severity_level >= 5.0 and confidence_score >= 70) != (severity_level >= 5.0 and confidence_score >= 75):
        b[39] = 40
        triggered.add(40)
    if (severity_level >= 5.0 and confidence_score >= 70) != (severity_level != 5.0 and confidence_score >= 70):
        b[40] = 41
        triggered.add(41)
    if (severity_level >= 5.0 and confidence_score >= 70) != (severity_level == 5.0 and confidence_score >= 70):
        b[41] = 42
        triggered.add(42)
    if (severity_level >= 5.0 and confidence_score >= 70) != (severity_level >= 5.0 and confidence_score >= 74):
        b[42] = 43
        triggered.add(43)
    if (severity_level >= 5.0 and confidence_score >= 70) != (severity_level >= 6.2 and confidence_score >= 70):
        b[43] = 44
        triggered.add(44)

if (threat_count >= 20) != (threat_count != 20):
    b[44] = 45
    triggered.add(45)
if (threat_count >= 20) != (threat_count >= 30):
    b[45] = 46
    triggered.add(46)
if (threat_count >= 20) != (threat_count >= 25):
    b[46] = 47
    triggered.add(47)

elif threat_count >= 20:
    if (severity_level >= 6.0 and confidence_score >= 80) != (severity_level >= 6.0 and confidence_score != 80):
        b[47] = 48
        triggered.add(48)
    if (severity_level >= 6.0 and confidence_score >= 80) != (severity_level >= 6.0 and confidence_score == 80):
        b[48] = 49
        triggered.add(49)
    if (severity_level >= 6.0 and confidence_score >= 80) != (severity_level >= 6.0 and confidence_score >= 87):
        b[49] = 50
        triggered.add(50)
    if (severity_level >= 6.0 and confidence_score >= 80) != (severity_level != 6.0 and confidence_score >= 80):
        b[50] = 51
        triggered.add(51)
    if (severity_level >= 6.0 and confidence_score >= 80) != (severity_level == 6.0 and confidence_score >= 80):
        b[51] = 52
        triggered.add(52)

    if (severity_level >= 4.0) != (severity_level != 4.0):
        b[52] = 53
        triggered.add(53)
    if (severity_level >= 4.0) != (severity_level >= 2.5):
        b[53] = 54
        triggered.add(54)

if (threat_count >= 5) != (threat_count != 5):
    b[54] = 55
    triggered.add(55)
if (threat_count >= 5) != (threat_count >= 7.2):
    b[55] = 56
    triggered.add(56)

elif threat_count >= 5:
    if (severity_level >= 7.0) != (severity_level >= 8.2):
        b[56] = 57
        triggered.add(57)
    if (severity_level >= 7.0) != (severity_level == 7.0):
        b[57] = 58
        triggered.add(58)
    if (severity_level >= 7.0) != (severity_level >= 7.5):
        b[58] = 59
        triggered.add(59)

# 严重程度维度的利益相关者通知
if (severity_level >= 9.0) != (severity_level >= 9.4):
    b[59] = 60
    triggered.add(60)
if (severity_level >= 9.0) != (severity_level == 9.0):
    b[60] = 61
    triggered.add(61)

if severity_level >= 9.0:
    if (threat_count >= 50 and confidence_score >= 90) != (threat_count >= 50 and confidence_score >= 93):
        b[61] = 62
        triggered.add(62)
    if (threat_count >= 50 and confidence_score >= 90) != (threat_count >= 50 and confidence_score == 90):
        b[62] = 63
        triggered.add(63)
    if (threat_count >= 50 and confidence_score >= 90) != (threat_count >= 55 and confidence_score >= 93):
        b[63] = 64
        triggered.add(64)
    if (threat_count >= 50 and confidence_score >= 90) != (threat_count >= 66 and confidence_score >= 90):
        b[64] = 65
        triggered.add(65)
    if (threat_count >= 50 and confidence_score >= 90) != (threat_count >= 50 and confidence_score == 90):
        b[65] = 66
        triggered.add(66)

    if (threat_count >= 20 or confidence_score >= 85) != (threat_count != 20 or confidence_score >= 85):
        b[66] = 67
        triggered.add(67)
    if (threat_count >= 20 or confidence_score >= 85) != (threat_count >= 50 or confidence_score >= 85):
        b[67] = 68
        triggered.add(68)
    if (threat_count >= 20 or confidence_score >= 85) != (threat_count >= 35 or confidence_score >= 85):
        b[68] = 69
        triggered.add(69)
    if (threat_count >= 20 or confidence_score >= 85) != (threat_count >= 20 or confidence_score != 85):
        b[69] = 70
        triggered.add(70)
    if (threat_count >= 20 or confidence_score >= 85) != (threat_count >= 20 or confidence_score == 85):
        b[70] = 71
        triggered.add(71)

if (severity_level >= 7.0) != (severity_level >= 5.5):
    b[71] = 72
    triggered.add(72)
if (severity_level >= 7.0) != (severity_level == 7.0):
    b[72] = 73
    triggered.add(73)

elif severity_level >= 7.0:
    if (confidence_score >= 85) != (confidence_score >= 87):
        b[73] = 74
        triggered.add(74)
    if (confidence_score >= 85) != (confidence_score == 85):
        b[74] = 75
        triggered.add(75)

if (severity_level >= 5.0) != (severity_level != 5.0):
    b[75] = 76
    triggered.add(76)
if (severity_level >= 5.0) != (severity_level >= 6.0):
    b[76] = 77
    triggered.add(77)

# 置信度分数维度的行动建议
if (confidence_score >= 95) != (confidence_score >= 97):
    b[77] = 78
    triggered.add(78)
if (confidence_score >= 95) != (confidence_score >= 89):
    b[78] = 79
    triggered.add(79)
if (confidence_score >= 95) != (confidence_score == 95):
    b[79] = 80
    triggered.add(80)

if confidence_score >= 95:
    if (severity_level >= 8.0 and threat_count >= 30) != (severity_level >= 8.0 and threat_count != 30):
        b[80] = 81
        triggered.add(81)
    if (severity_level >= 8.0 and threat_count >= 30) != (severity_level >= 8.0 and threat_count == 30):
        b[81] = 82
        triggered.add(82)
    if (severity_level >= 8.0 and threat_count >= 30) != (severity_level >= 8.0 and threat_count >= 31):
        b[82] = 83
        triggered.add(83)
    if (severity_level >= 8.0 and threat_count >= 30) != (severity_level >= 8.0 and threat_count >= 45):
        b[83] = 84
        triggered.add(84)
    if (severity_level >= 8.0 and threat_count >= 30) != (severity_level == 8.0 and threat_count >= 30):
        b[84] = 85
        triggered.add(85)

    if (severity_level >= 6.0) != (severity_level >= 6.6):
        b[85] = 86
        triggered.add(86)
    if (severity_level >= 6.0) != (severity_level >= 7.3):
        b[86] = 87
        triggered.add(87)

if (confidence_score >= 85) != (confidence_score >= 78):
    b[87] = 88
    triggered.add(88)
if (confidence_score >= 85) != (confidence_score >= 87):
    b[88] = 89
    triggered.add(89)

elif confidence_score >= 85:
    if (severity_level >= 7.0) != (severity_level >= 8.7):
        b[89] = 90
        triggered.add(90)
    if (severity_level >= 7.0) != (severity_level == 7.0):
        b[90] = 91
        triggered.add(91)
    if (severity_level >= 7.0) != (severity_level >= 7.2):
        b[91] = 92
        triggered.add(92)

    if (severity_level >= 5.0) != (severity_level >= 6.3):
        b[92] = 93
        triggered.add(93)
    if (severity_level >= 5.0) != (severity_level >= 3.7):
        b[93] = 94
        triggered.add(94)

if (confidence_score >= 70) != (confidence_score >= 77):
    b[94] = 95
    triggered.add(95)
if (confidence_score >= 70) != (confidence_score >= 86):
    b[95] = 96
    triggered.add(96)

return triggered
# === 目标路径组 ===

targetPaths = [
    # A1
    {6, 7, 8, 11, 13, 17, 19, 21, 28, 29, 30, 32, 34, 37, 39, 42, 49, 52, 57, 58, 63, 64, 65, 66, 73, 75, 80, 82, 90, 91},
    # A2
    {5, 7, 8, 10, 15, 19, 21, 28, 29, 32, 34, 36, 39, 42, 49, 52, 58, 60, 62, 63, 64, 65, 66, 73, 75, 79, 82, 85, 91},
    # A3
    {6, 7, 8, 11, 14, 17, 19, 21, 28, 29, 30, 32, 34, 39, 42, 49, 52, 57, 58, 63, 64, 65, 66, 73, 75, 78, 82, 90, 91},
    # A4
    {4, 5, 7, 8, 10, 15, 19, 21, 28, 29, 30, 32, 34, 36, 39, 42, 49, 52, 58, 60, 62, 64, 65, 73, 75, 79, 82, 85, 91},
    # A5
    {5, 6, 7, 8, 10, 11, 12, 17, 19, 21, 28, 32, 34, 36, 39, 42, 49, 52, 57, 58, 63, 65, 66, 73, 75, 79, 82, 90, 91},
    # A6
    {5, 10, 16, 19, 21, 28, 29, 30, 32, 34, 36, 39, 42, 49, 52, 58, 61, 62, 63, 64, 65, 66, 73, 75, 79, 82, 85, 91},
    # A7
    {1, 2, 6, 7, 8, 11, 13, 14, 17, 19, 21, 32, 34, 37, 39, 42, 49, 52, 57, 58, 63, 66, 73, 75, 78, 80, 82, 90, 91},
    # A8
    {5, 6, 7, 8, 10, 11, 12, 17, 19, 21, 32, 34, 36, 39, 42, 45, 49, 52, 55, 57, 58, 71, 73, 75, 79, 81, 90, 91},
    # A9
    {3, 6, 7, 8, 11, 13, 14, 17, 19, 21, 32, 34, 37, 39, 42, 49, 52, 57, 58, 63, 66, 73, 75, 78, 80, 82, 90, 91},
    # A10
    {4, 5, 6, 7, 8, 10, 11, 17, 19, 21, 32, 34, 36, 39, 42, 45, 46, 47, 49, 52, 57, 58, 73, 75, 79, 81, 90, 91},
    # A11
    {6, 7, 8, 11, 13, 17, 19, 21, 32, 34, 37, 39, 42, 45, 49, 52, 55, 56, 57, 58, 71, 73, 75, 80, 81, 90, 91},
    # A12
    {4, 15, 19, 21, 23, 28, 29, 30, 32, 34, 35, 36, 39, 42, 49, 50, 52, 58, 60, 73, 74, 75, 82, 85, 89, 91},
    # A13
    {4, 17, 19, 21, 23, 32, 33, 35, 36, 39, 42, 45, 46, 47, 49, 50, 52, 57, 58, 73, 74, 81, 89, 90, 91, 96},
    # A14
    {4, 17, 19, 21, 23, 32, 33, 35, 36, 39, 42, 45, 49, 50, 52, 56, 57, 58, 70, 73, 74, 81, 89, 90, 91, 96},
    # A15
    {4, 17, 19, 21, 23, 33, 39, 42, 45, 46, 47, 49, 50, 52, 57, 58, 67, 68, 69, 73, 81, 88, 90, 91, 96},
    # A16
    {4, 17, 19, 21, 22, 23, 24, 33, 39, 42, 45, 49, 50, 52, 55, 57, 58, 67, 70, 73, 81, 88, 90, 91, 96},
    # A17
    {6, 13, 14, 18, 21, 24, 26, 28, 31, 37, 39, 42, 44, 49, 51, 63, 65, 66, 72, 75, 78, 80, 86, 87, 93},
    # A18
    {6, 13, 14, 19, 21, 24, 28, 31, 34, 37, 39, 42, 49, 52, 57, 59, 63, 65, 66, 75, 78, 80, 87, 90, 92},
    # A19
    {5, 6, 7, 8, 10, 11, 17, 19, 21, 32, 34, 36, 39, 42, 49, 52, 57, 58, 73, 75, 79, 82, 84, 90, 91},
    # A20
    {4, 16, 19, 20, 22, 23, 24, 28, 29, 30, 33, 39, 42, 48, 50, 52, 58, 61, 73, 82, 85, 88, 91, 96},
    # A21
    {6, 13, 14, 18, 26, 28, 29, 30, 31, 37, 39, 41, 44, 51, 63, 64, 65, 66, 75, 76, 77, 78, 80, 93},
    # A22
    {6, 7, 8, 11, 14, 17, 19, 21, 32, 34, 39, 42, 49, 52, 57, 58, 73, 75, 78, 81, 83, 84, 90, 91},
    # A23
    {4, 9, 17, 19, 21, 32, 34, 36, 39, 42, 45, 46, 47, 49, 52, 57, 58, 73, 75, 79, 81, 90, 91},
    # A24
    {4, 17, 20, 33, 38, 40, 42, 43, 45, 48, 55, 56, 57, 58, 67, 70, 73, 81, 90, 91, 95, 96},
    # A25
    {6, 18, 25, 26, 27, 28, 29, 30, 31, 36, 41, 51, 53, 62, 63, 64, 65, 66, 75, 76, 79, 94},
    # A26
    {6, 12, 18, 25, 28, 29, 30, 31, 36, 41, 51, 53, 54, 63, 64, 65, 66, 75, 76, 79}
]
```

