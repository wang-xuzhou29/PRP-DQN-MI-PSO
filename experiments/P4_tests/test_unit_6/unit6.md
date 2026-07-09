

```python
def execute_Tr(x, y, z):
    # 初始化分支覆盖数组
    b = [0] * 107  # 根据分支数量调整大小

    # 定义质量参数
    quality_x = x
    quality_y = y
    quality_z = z

    # 定义故障参数
    fault_x = x
    fault_y = y
    fault_z = z

    # 定义过程参数
    process_a = random.randint(2, 100)
    process_b = random.randint(20, 150)
    process_c = random.randint(30, 200)

    if ((process_b * process_c) / (process_a + 1) > 110) != ((process_a * process_c) / (process_a + 1) > 110): b[0] = 1
    if ((process_b * process_c) / (process_a + 1) > 110) != ((process_c * process_c) / (process_a + 1) > 110): b[1] = 2
    if ((process_b * process_c) / (process_a + 1) > 110) != ((process_b * process_b) / (process_a + 1) > 110): b[2] = 3
    if ((process_b * process_c) / (process_a + 1) > 110) != ((process_b * process_a) / (process_a + 1) > 110): b[3] = 4
    if ((process_b * process_c) / (process_a + 1) > 110) != ((process_b * 12) / (process_a + 1) > 110): b[4] = 5
    if ((process_b * process_c) / (process_a + 1) > 110) != ((process_b * process_c) / (process_b + 1) > 110): b[5] = 6
    if ((process_b * process_c) / (process_a + 1) > 110) != ((process_b * process_c) / (process_c + 1) > 110): b[6] = 7
    if ((process_b * process_c) / (process_a + 1) > 110) != ((process_b * process_c) / (process_a + 5) > 110): b[7] = 8
    if ((process_b * process_c) / (process_a + 1) > 110) != ((process_b * process_c) / (process_a + 1) > 130): b[8] = 9
    if ((process_b * process_c) / (process_a + 1) > 110) != ((50 * process_c) / (process_a + 1) > 110): b[9] = 10
    if ((process_b * process_c) / (process_a + 1) > 110) != ((process_b * process_c) / (process_a - 1) > 110): b[
        10] = 11
    if ((process_b * process_c) / (process_a + 1) > 110) != ((process_b * process_c * 3) / (process_a + 1) > 110): b[
        11] = 12
    if ((process_b * process_c) / (process_a + 1) > 110) != ((process_b * process_c) / (process_a * 0.5 + 1) > 110): b[
        12] = 13
    if ((process_b * process_c) / (process_a + 1) > 110) != ((process_b * process_c) / (process_a * 2 + 1) > 110): b[
        13] = 14
    if ((process_b * process_c) / (process_a + 1) > 110) != ((process_b * process_c) / (60 + 1) > 110): b[14] = 15

    # 优化策略2：参数关系调整
    if ((process_c - process_a) < 0.28 * process_b) != ((process_b - process_a) < 0.28 * process_b): b[15] = 16
    if ((process_c - process_a) < 0.28 * process_b) != ((process_c * 2 - process_a) < 0.28 * process_b): b[16] = 17
    if ((process_c - process_a) < 0.28 * process_b) != ((process_c * 1.5 - process_a) < 0.28 * process_b): b[17] = 18
    if ((process_c - process_a) < 0.28 * process_b) != ((process_c - process_b) < 0.28 * process_b): b[18] = 19
    if ((process_c - process_a) < 0.28 * process_b) != ((process_c - process_a) < 0.28 * process_a): b[19] = 20
    if ((process_c - process_a) < 0.28 * process_b) != ((process_c - process_a) < 0.28 * process_c): b[20] = 21
    if ((process_c - process_a) < 0.28 * process_b) != ((process_c - process_a) < 0.48 * process_b): b[21] = 22
    if ((process_c - process_a) < 0.28 * process_b) != ((process_c - process_a) < 0.18 * process_b): b[22] = 23
    if ((process_c - process_a) < 0.28 * process_b) != ((process_c + process_a) < 0.28 * process_b): b[23] = 24
    if ((process_c - process_a) < 0.28 * process_b) != ((process_c - process_a * 1.2) < 0.28 * process_b): b[24] = 25
    if ((process_c - process_a) < 0.28 * process_b) != ((process_c - process_a * 2) < 0.28 * process_b): b[25] = 26

    # 优化策略3：非线性效率优化35
    if ((process_a ** 3 + process_b ** 3) < process_c ** 2) != ((process_a ** 2.5 + process_b ** 3) < process_c ** 2): b[
        26] = 27
    if ((process_a ** 3 + process_b ** 3) < process_c ** 2) != ((process_a ** 3 + process_b ** 2.5) < process_c ** 2): b[
        27] = 28
    if ((process_a ** 3 + process_b ** 3) < process_c ** 2) != ((process_a ** 3 - process_b ** 3) < process_c ** 2): b[
        28] = 29
    if ((process_a ** 3 + process_b ** 3) < process_c ** 2) != ((process_b ** 3 + process_b ** 3) < process_c ** 2): b[
        29] = 30
    if ((process_a ** 3 + process_b ** 3) < process_c ** 2) != ((process_c ** 3 + process_b ** 3) < process_c ** 2): b[
        30] = 31
    if ((process_a ** 3 + process_b ** 3) < process_c ** 2) != ((process_a ** 3 + process_a ** 3) < process_c ** 2): b[
        31] = 32
    if ((process_a ** 3 + process_b ** 3) < process_c ** 2) != ((process_a ** 3 + process_c ** 3) < process_c ** 2): b[
        32] = 33
    if ((process_a ** 3 + process_b ** 3) < process_c ** 2) != ((process_a ** 3 + process_b ** 2.5) < process_c ** 2): b[
        33] = 34
    if ((process_a ** 3 + process_b ** 3) < process_c ** 2) != ((process_a ** 3 + process_b ** 3) < process_a ** 2): b[
        34] = 35
    if ((process_a ** 3 + process_b ** 3) < process_c ** 2) != ((process_a ** 3 + process_b ** 3) < process_b ** 2): b[
        35] = 36
    if ((process_a ** 3 + process_b ** 3) < process_c ** 2) != ((process_a ** 3 + process_b ** 3) < process_c ** 2.3): b[
        36] = 37
    if ((process_a ** 3 + process_b ** 3) < process_c ** 2) != ((process_a ** 3 + process_b ** 3) < process_c ** 1.5): b[
        37] = 38
    if ((process_a ** 3 + process_b ** 3) < process_c ** 2) != ((process_a ** 3 + process_b * 3) < process_c ** 2): b[
        38] = 39
    if ((process_a ** 3 + process_b ** 3) < process_c ** 2) != ((process_a * 3 + process_b ** 3) < process_c ** 2): b[
        39] = 40
    if ((process_a ** 3 + process_b ** 3) < process_c ** 2) != ((process_a ** 3 + process_b ** 3) < process_c * 2): b[
        40] = 41

    # 优化策略5：理想线性工艺
    if (abs(process_c - (process_a + process_b)) < 2.5 and abs(process_b - process_a * 1.25) < 1.5) != (
            abs(process_c * 2 - (process_a + process_b)) < 2.5 and abs(process_b - process_a * 1.25) < 1.5): b[41] = 42
    if (abs(process_c - (process_a + process_b)) < 2.5 and abs(process_b - process_a * 1.25) < 1.5) != (
            abs(process_c - (process_a + process_a)) < 2.5 and abs(process_b - process_a * 1.25) < 1.5): b[42] = 43
    if (abs(process_c - (process_a + process_b)) < 2.5 and abs(process_b - process_a * 1.25) < 1.5) != (
            abs(process_c - (process_a - process_b)) < 2.5 and abs(process_b - process_a * 1.25) < 1.5): b[43] = 44
    if (abs(process_c - (process_a + process_b)) < 2.5 and abs(process_b - process_a * 1.25) < 1.5) != (
            abs(process_c - (process_b + process_b)) < 2.5 and abs(process_b - process_a * 1.25) < 1.5): b[44] = 45
    if (abs(process_c - (process_a + process_b)) < 2.5 and abs(process_b - process_a * 1.25) < 1.5) != (
            abs(process_c - (process_a * 0.9 + process_b)) < 2.5 and abs(process_b - process_a * 1.25) < 1.5): b[
        45] = 46
    if (abs(process_c - (process_a + process_b)) < 2.5 and abs(process_b - process_a * 1.25) < 1.5) != (
            abs(process_c - (process_a + process_b * 0.8)) < 2.5 and abs(process_b - process_a * 1.25) < 1.5): b[
        46] = 47
    if (abs(process_c - (process_a + process_b)) < 2.5 and abs(process_b - process_a * 1.25) < 1.5) != (
            abs(process_c - (process_a + process_b)) < 7 and abs(process_b - process_a * 1.25) < 1.5): b[47] = 48
    if (abs(process_c - (process_a + process_b)) < 2.5 and abs(process_b - process_a * 1.25) < 1.5) != (
            abs(process_c - (process_a + process_b)) < 2.5 and abs(process_c - process_a * 1.25) < 1.5): b[48] = 49
    if (abs(process_c - (process_a + process_b)) < 2.5 and abs(process_b - process_a * 1.25) < 1.5) != (
            abs(process_c - (process_a + process_b)) < 2.5 and abs(process_b * 1.5 - process_a * 1.25) < 1.5): b[
        49] = 50
    if (abs(process_c - (process_a + process_b)) < 2.5 and abs(process_b - process_a * 1.25) < 1.5) != (
            abs(process_c - (process_a + process_b)) < 2.5 and abs(process_b - process_a * 1.75) < 1.5): b[50] = 51
    if (abs(process_c - (process_a + process_b)) < 2.5 and abs(process_b - process_a * 1.25) < 1.5) != (
            abs(process_c - (process_a + process_b)) < 2.5 and abs(process_b - process_c * 1.25) < 1.5): b[51] = 52
    if (abs(process_c - (process_a + process_b)) < 2.5 and abs(process_b - process_a * 1.25) < 1.5) != (
            abs(process_c - (process_a + process_b)) < 2.5 and abs(process_b - process_a * 1.25) < 8): b[52] = 53
    if (abs(process_c - (process_a + process_b)) < 2.5 and abs(process_b - process_a * 1.25) < 1.5) != (
            abs(process_c - (process_a + process_b)) < 2.5 and abs(process_b - process_a ** 1.25) < 1.5): b[53] = 54

    # 优化策略7：产能平衡调整
    if ((process_a / (process_b + 0.01)) > 3.5 and (process_b / (process_c + 0.01)) < 0.3) != (
            (process_a * 1.3 / (process_b + 0.01)) > 3.5 and (process_b / (process_c + 0.01)) < 0.3): b[54] = 55
    if ((process_a / (process_b + 0.01)) > 3.5 and (process_b / (process_c + 0.01)) < 0.3) != (
            (process_a / (process_b * 0.6 + 0.01)) > 3.5 and (process_b / (process_c + 0.01)) < 0.3): b[55] = 56
    if ((process_a / (process_b + 0.01)) > 3.5 and (process_b / (process_c + 0.01)) < 0.3) != (
            (process_a / (process_b + 0.01)) > 8 and (process_b / (process_c + 0.01)) < 0.3): b[56] = 67
    if ((process_a / (process_b + 0.01)) > 3.5 and (process_b / (process_c + 0.01)) < 0.3) != (
            (process_a / (process_b + 0.01)) > 3.5 and (process_c / (process_c + 0.01)) < 0.3): b[57] = 58
    if ((process_a / (process_b + 0.01)) > 3.5 and (process_b / (process_c + 0.01)) < 0.3) != (
            (process_a / (process_b + 0.01)) > 3.5 and (process_a / (process_c + 0.01)) < 0.3): b[58] = 59
    if ((process_a / (process_b + 0.01)) > 3.5 and (process_b / (process_c + 0.01)) < 0.3) != (
            (process_b / (process_b + 0.01)) > 3.5 and (process_b / (process_c + 0.01)) < 0.3): b[59] = 60
    if ((process_a / (process_b + 0.01)) > 3.5 and (process_b / (process_c + 0.01)) < 0.3) != (
            (process_c / (process_b + 0.01)) > 3.5 and (process_b / (process_c + 0.01)) < 0.3): b[60] = 61
    if ((process_a / (process_b + 0.01)) > 3.5 and (process_b / (process_c + 0.01)) < 0.3) != (
            (process_a % (process_b + 0.01)) > 3.5 and (process_b / (process_c + 0.01)) < 0.3): b[61] = 62
    if ((process_a / (process_b + 0.01)) > 3.5 and (process_b / (process_c + 0.01)) < 0.3) != (
            (process_a / (process_b + 0.01)) > 3.5 and (process_b % (process_c + 0.01)) < 0.3): b[62] = 63
    if ((process_a / (process_b + 0.01)) > 3.5 and (process_b / (process_c + 0.01)) < 0.3) != (
            (process_a / (process_b + 0.01)) > 3.5 or (process_b / (process_c + 0.01)) < 0.3): b[63] = 64
    if ((process_a / (process_b + 0.01)) > 3.5 and (process_b / (process_c + 0.01)) < 0.3) != (
            (process_a / (process_b + 0.01)) > 3.5 and (process_b / (process_c + 0.01)) < 0.2): b[64] = 65
    if ((process_a / (process_b + 0.01)) > 3.5 and (process_b / (process_c + 0.01)) < 0.3) != (
            (process_a / (process_b + 0.01)) > 3.5 and (process_b / (process_c + 0.01)) < 1.3): b[65] = 66
    if ((process_a / (process_b + 0.01)) > 3.5 and (process_b / (process_c + 0.01)) < 0.3) != (
            (process_a / (process_b + 0.01)) > 3.5 and (process_b / (process_c * 0.7 + 0.01)) < 0.3): b[66] = 67

    # 优化策略8：工艺波动控制
    if (abs(process_a - process_b) > 14 and abs(process_b - process_c) > 16 and abs(process_a - process_c) < 7) != (
            abs(process_a * 1.2 - process_b) > 14 and abs(process_b - process_c) > 16 and abs(
        process_a - process_c) < 7): b[67] = 68
    if (abs(process_a - process_b) > 14 and abs(process_b - process_c) > 16 and abs(process_a - process_c) < 7) != (
            abs(process_a * 1.6 - process_b) > 14 and abs(process_b - process_c) > 16 and abs(
        process_a - process_c) < 7): b[68] = 69
    if (abs(process_a - process_b) > 14 and abs(process_b - process_c) > 16 and abs(process_a - process_c) < 7) != (
            abs(process_a - process_b * 1.2) > 14 and abs(process_b - process_c) > 16 and abs(process_a - process_c) < 7): b[69] = 70
    if (abs(process_a - process_b) > 14 and abs(process_b - process_c) > 16 and abs(process_a - process_c) < 7) != (
            abs(process_a - process_c) > 14 and abs(process_b - process_c) > 16 and abs(process_a - process_c) < 7): b[
        70] = 71
    if (abs(process_a - process_b) > 14 and abs(process_b - process_c) > 16 and abs(process_a - process_c) < 7) != (
            abs(process_a - process_b) > 14 and abs(process_a - process_c) > 16 and abs(process_a - process_c) < 7): b[
        71] = 72
    if (abs(process_a - process_b) > 14 and abs(process_b - process_c) > 16 and abs(process_a - process_c) < 7) != (
            abs(process_a - process_b) > 14 and abs(process_b - process_a) > 16 and abs(process_a - process_c) < 7): b[
        72] = 73
    if (abs(process_a - process_b) > 14 and abs(process_b - process_c) > 16 and abs(process_a - process_c) < 7) != (
            abs(process_a - process_b) > 14 and abs(process_b - process_c) > 16 and abs(process_b - process_c) < 7): b[
        73] = 74
    if (abs(process_a - process_b) > 14 and abs(process_b - process_c) > 16 and abs(process_a - process_c) < 7) != (
            abs(process_a - process_b) > 14 and abs(process_b - process_c) > 16 and abs(process_a - process_b) < 7): b[
        74] = 75
    if (abs(process_a - process_b) > 14 and abs(process_b - process_c) > 16 and abs(process_a - process_c) < 7) != (
            abs(process_a - process_b) > 14 and abs(process_b - process_c) > 16 and abs(process_a - process_c) < 12): b[
        75] = 76
    if (abs(process_a - process_b) > 14 and abs(process_b - process_c) > 16 and abs(process_a - process_c) < 7) != (
            abs(process_a - process_b) > 14 and abs(process_b - process_c) > 16 and abs(process_a * 2 - process_c) < 7): b[
        76] = 77
    if (abs(process_a - process_b) > 14 and abs(process_b - process_c) > 16 and abs(process_a - process_c) < 7) != (
            abs(process_a - process_b) > 14 and abs(process_b * 3 - process_c) > 16 and abs(process_a - process_c) < 7): b[
        77] = 78
    if (abs(process_a - process_b) > 14 and abs(process_b - process_c) > 16 and abs(process_a - process_c) < 7) != (
            abs(process_a - process_b) > 14 and abs(process_b - process_c * 1.6) > 16 and abs(process_a - process_c) < 7): b[78] = 79
    if (abs(process_a - process_b) > 14 and abs(process_b - process_c) > 16 and abs(process_a - process_c) < 7) != (
            abs(process_a - process_b) > 14 and abs(process_b - process_c * 0.7) > 16 and abs(process_a - process_c) < 7): b[79] = 80
    if (abs(process_a - process_b) > 14 and abs(process_b - process_c) > 16 and abs(process_a - process_c) < 7) != (
            abs(process_a - process_b) > 20 and abs(process_b - process_c) > 16 and abs(process_a - process_c) < 7): b[
        80] = 81

    # 优化策略9：极值工艺处理
    if ((process_a > 85 or process_a < 8) and (process_b > 80 or process_b < 5) and (
            process_c > 75 or process_c < 4)) != (
            (process_a * process_a > 85 or process_a < 8) and (process_b > 80 or process_b < 5) and (
            process_c > 75 or process_c < 4)): b[81] = 82
    if ((process_a > 85 or process_a < 8) and (process_b > 80 or process_b < 5) and (
            process_c > 75 or process_c < 4)) != (
            (process_a * process_b > 85 or process_a < 8) and (process_b > 80 or process_b < 5) and (
            process_c > 75 or process_c < 4)): b[82] = 83
    if ((process_a > 85 or process_a < 8) and (process_b > 80 or process_b < 5) and (
            process_c > 75 or process_c < 4)) != (
            (process_a * process_c > 85 or process_a < 8) and (process_b > 80 or process_b < 5) and (
            process_c > 75 or process_c < 4)): b[83] = 84
    if ((process_a > 85 or process_a < 8) and (process_b > 80 or process_b < 5) and (
            process_c > 75 or process_c < 4)) != (
            (process_a > 15 or process_a < 8) and (process_b > 80 or process_b < 5) and (
            process_c > 75 or process_c < 4)): b[84] = 85
    if ((process_a > 85 or process_a < 8) and (process_b > 80 or process_b < 5) and (
            process_c > 75 or process_c < 4)) != (
            (process_a > 85 or process_a < 8) and (process_b * process_a > 80 or process_b < 5) and (
            process_c > 75 or process_c < 4)): b[85] = 86
    if ((process_a > 85 or process_a < 8) and (process_b > 80 or process_b < 5) and (
            process_c > 75 or process_c < 4)) != (
            (process_a > 85 or process_a < 8) and (process_b * process_b > 80 or process_b < 5) and (
            process_c > 75 or process_c < 4)): b[86] = 87
    if ((process_a > 85 or process_a < 8) and (process_b > 80 or process_b < 5) and (
            process_c > 75 or process_c < 4)) != (
            (process_a > 85 or process_a < 8) and (process_b * process_c > 80 or process_b < 5) and (
            process_c > 75 or process_c < 4)): b[87] = 88
    if ((process_a > 85 or process_a < 8) and (process_b > 80 or process_b < 5) and (
            process_c > 75 or process_c < 4)) != (
            (process_a > 85 or process_a < 8) and (process_b > 10 or process_b < 5) and (
            process_c > 75 or process_c < 4)): b[88] = 89
    if ((process_a > 85 or process_a < 8) and (process_b > 80 or process_b < 5) and (
            process_c > 75 or process_c < 4)) != (
            (process_a * 50 > 85 or process_a < 8) and (process_b > 80 or process_b < 5) and (
            process_c > 75 or process_c < 4)): b[89] = 90
    if ((process_a > 85 or process_a < 8) and (process_b > 80 or process_b < 5) and (
            process_c > 75 or process_c < 4)) != (
            (process_a * 80 > 85 or process_a < 8) and (process_b > 80 or process_b < 5) and (
            process_c > 75 or process_c < 4)): b[90] = 91
    if ((process_a > 85 or process_a < 8) and (process_b > 80 or process_b < 5) and (
            process_c > 75 or process_c < 4)) != (
            (process_a > 15 or process_a < 8) and (process_b > 80 or process_b < 5) and (
            process_c > 75 or process_c < 4)): b[91] = 92
    if ((process_a > 85 or process_a < 8) and (process_b > 80 or process_b < 5) and (
            process_c > 75 or process_c < 4)) != (
            (process_a > 85 or process_a < 8) and (process_b > 80 or process_b < 5) and (
            process_c * process_c > 75 or process_c < 4)): b[92] = 93

    if ((process_a + process_b) ** 1.2 < process_c ** 1.8 and (process_a + process_b + process_c) / 3 > 45) != (
            (process_b + process_b) ** 1.2 < process_c ** 1.8 and (process_a + process_b + process_c) / 3 > 45): b[93] = 94
    if ((process_a + process_b) ** 1.2 < process_c ** 1.8 and (process_a + process_b + process_c) / 3 > 45) != (
            (process_c + process_b) ** 1.2 < process_c ** 1.8 and (process_a + process_b + process_c) / 3 > 45): b[94] = 95
    if ((process_a + process_b) ** 1.2 < process_c ** 1.8 and (process_a + process_b + process_c) / 3 > 45) != (
            (process_a + process_a) ** 1.2 < process_c ** 1.8 and (process_a + process_b + process_c) / 3 > 45): b[95] = 96
    if ((process_a + process_b) ** 1.2 < process_c ** 1.8 and (process_a + process_b + process_c) / 3 > 45) != (
            (process_a + process_c) ** 1.2 < process_c ** 1.8 and (process_a + process_b + process_c) / 3 > 45): b[96] = 97
    if ((process_a + process_b) ** 1.2 < process_c ** 1.8 and (process_a + process_b + process_c) / 3 > 45) != (
            (process_a + process_b) ** 1.8 < process_c ** 1.8 and (process_a + process_b + process_c) / 3 > 45): b[97] = 98
    if ((process_a + process_b) ** 1.2 < process_c ** 1.8 and (process_a + process_b + process_c) / 3 > 45) != (
            (process_a + process_b) ** 1.2 < process_a ** 1.8 and (process_a + process_b + process_c) / 3 > 45): b[98] = 99
    if ((process_a + process_b) ** 1.2 < process_c ** 1.8 and (process_a + process_b + process_c) / 3 > 45) != (
            (process_a + process_b) ** 1.2 < process_b ** 1.8 and (process_a + process_b + process_c) / 3 > 45): b[99] = 100
    if ((process_a + process_b) ** 1.2 < process_c ** 1.8 and (process_a + process_b + process_c) / 3 > 45) != (
            (process_a + process_b) ** 0.8 < process_c ** 1.8 and (process_a + process_b + process_c) / 3 > 45): b[
        100] = 101
    if ((process_a + process_b) ** 1.2 < process_c ** 1.8 and (process_a + process_b + process_c) / 3 > 45) != (
            (process_a + process_b) ** 1.2 < process_c ** 1.8 and (process_a + process_a + process_c) / 3 > 45): b[
        101] = 102
    if ((process_a + process_b) ** 1.2 < process_c ** 1.8 and (process_a + process_b + process_c) / 3 > 45) != (
            (process_a + process_b) ** 1.2 < process_c ** 1.8 and (process_a + process_c + process_c) / 3 > 45): b[
        102] = 103
    if ((process_a + process_b) ** 1.2 < process_c ** 1.8 and (process_a + process_b + process_c) / 3 > 45) != (
            (process_a + process_b) ** 1.2 < process_c ** 1.8 and (process_a + process_b + process_a) / 3 > 45): b[
        103] = 104
    if ((process_a + process_b) ** 1.2 < process_c ** 1.8 and (process_a + process_b + process_c) / 3 > 45) != (
            (process_a + process_b) ** 1.2 < process_c ** 1.8 and (process_a + process_b + process_b) / 3 > 45): b[
        104] = 105
    if ((process_a + process_b) ** 1.2 < process_c ** 1.8 and (process_a + process_b + process_c) / 3 > 45) != (
            (process_a + process_b) ** 1.2 < process_c ** 1.8 and (process_a + 50 + process_c) / 3 > 45): b[105] = 106
    if ((process_a + process_b) ** 1.2 < process_c ** 1.8 and (process_a + process_b + process_c) / 3 > 45) != (
            (process_a + process_b) ** 1.2 < process_c ** 1.8 and (60 + process_b + process_c) / 3 > 45): b[106] = 107

# 目标路径定义
targetPaths = [
    {1, 2, 5, 6, 8, 9, 10, 14, 16, 17, 18, 24, 29, 69, 71, 72, 74, 75, 77, 79, 82, 83, 84, 85, 90, 91, 92, 98},
    {1, 2, 5, 6, 8, 9, 10, 14, 16, 17, 18, 24, 29, 71, 72, 74, 75, 77, 79, 82, 83, 84, 85, 90, 91, 92, 98},
    {1, 2, 4, 5, 6, 7, 8, 9, 10, 14, 17, 18, 24, 29, 71, 72, 74, 75, 77, 82, 83, 84, 85, 90, 91, 92, 98},
    {3, 4, 7, 12, 13, 15, 17, 18, 24, 29, 68, 71, 72, 74, 75, 77, 79, 82, 83, 84, 85, 90, 91, 92, 98},
    {1, 2, 4, 5, 6, 7, 10, 14, 15, 16, 17, 18, 20, 21, 24, 29, 77, 94, 98, 99, 102, 103, 104, 106},
    {1, 3, 4, 5, 7, 15, 30, 31, 33, 35, 36, 38, 41, 61, 62, 64, 86, 87, 88, 89, 99, 102, 104, 105},
    {4, 5, 7, 16, 26, 29, 42, 43, 44, 45, 46, 47, 49, 50, 51, 52, 54, 82, 83, 84, 85, 90, 91, 92},
    {1, 3, 4, 5, 6, 7, 8, 9, 10, 14, 16, 19, 21, 22, 25, 26, 29, 82, 83, 84, 85, 90, 91, 92, 98},
    {1, 4, 6, 7, 15, 16, 17, 18, 20, 21, 23, 24, 29, 32, 39, 93, 94, 98, 99, 102, 103, 104, 106},
    {1, 4, 5, 6, 7, 15, 28, 29, 32, 34, 37, 39, 61, 62, 64, 86, 87, 88, 89, 99, 102, 104, 105},
    {16, 20, 21, 22, 25, 26, 57, 58, 59, 60, 62, 63, 65, 67, 76, 86, 87, 88, 89, 98, 100, 105},
    {2, 16, 20, 21, 25, 26, 40, 49, 57, 58, 59, 60, 63, 65, 86, 87, 88, 89, 98, 100, 105},
    {3, 4, 5, 7, 8, 9, 14, 15, 16, 27, 28, 29, 30, 34, 37, 39, 40, 61, 62, 64, 104, 105},
    {1, 3, 4, 5, 6, 7, 8, 9, 14, 15, 28, 29, 32, 34, 37, 39, 61, 62, 64, 103, 106, 107},
    {1, 2, 5, 6, 9, 10, 14, 15, 16, 24, 29, 71, 72, 74, 75, 77, 96, 97, 99, 100, 101},
    {12, 13, 15, 17, 18, 19, 24, 70, 71, 72, 74, 75, 77, 80, 81, 86, 87, 88, 89, 98},
    {3, 12, 13, 15, 17, 18, 24, 29, 73, 78, 79, 80, 82, 83, 84, 85, 90, 91, 92, 98},
    {1, 6, 15, 16, 17, 18, 20, 21, 29, 32, 39, 93, 94, 95, 98, 99, 102, 103, 106},
    {3, 4, 7, 11, 12, 13, 15, 17, 18, 24, 29, 68, 71, 72, 73, 74, 75, 77, 81, 98},
    {2, 16, 20, 21, 22, 25, 26, 55, 56, 61, 62, 64, 76, 86, 87, 88, 89, 98},
    {4, 5, 7, 10, 14, 16, 26, 29, 46, 48, 82, 83, 84, 85, 90, 91, 92, 98},
    {3, 4, 5, 7, 10, 14, 16, 26, 29, 53, 82, 83, 84, 85, 90, 91, 92},
    {17, 18, 19, 24, 64, 66, 76, 86, 87, 88, 89, 98}
]

```



