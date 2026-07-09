

```python

def execute_Tr(x, y, z):
    temp = x
    pressure = y
    speed = z
    # 初始化分支覆盖数组
    b = [0] * 99  # 根据分支数量调整大小

    if ((temp * pressure) / (speed + 1) > 150) != ((temp * pressure) / (speed + 1) > 200): b[0] = 1
    if ((temp * pressure) / (speed + 1) > 150) != ((temp * pressure) / (speed * 2 + 1) > 150): b[1] = 2
    if ((temp * pressure) / (speed + 1) > 150) != ((temp * temp) / (speed + 1) > 150): b[2] = 3
    if ((temp * pressure) / (speed + 1) > 150) != ((temp * 2 * pressure) / (speed + 1) > 150): b[3] = 4
    if ((temp * pressure) / (speed + 1) > 150) != ((pressure * pressure) / (speed + 1) > 150): b[4] = 5
    if ((temp * pressure) / (speed + 1) > 150) != ((temp * pressure) / (speed + 1) > 500): b[5] = 6
    if ((temp * pressure) / (speed + 1) > 150) != ((temp * 0.5 * pressure) / (speed + 1) > 150): b[6] = 7
    if ((temp * pressure) / (speed + 1) > 150) != ((temp * pressure) / (speed + 10) > 150): b[7] = 8
    if ((temp * pressure) / (speed + 1) > 150) != ((temp * pressure) / (speed * speed + 1) > 150): b[8] = 9
    if ((temp * pressure) / (speed + 1) > 150) != ((temp / pressure) / (speed + 1) > 150): b[9] = 10

    # 验证规则2：相对偏差检测
    if ((pressure - temp) < 0.2 * speed) != ((pressure - temp * 2) < 0.2 * speed): b[10] = 11
    if ((pressure - temp) < 0.2 * speed) != ((pressure - temp) < 0.1 * speed): b[11] = 12
    if ((pressure - temp) < 0.2 * speed) != ((pressure - temp) < 0.3 * speed): b[12] = 13
    if ((pressure - temp) < 0.2 * speed) != ((pressure - temp) < 0.5 * speed): b[13] = 14
    if ((pressure - temp) < 0.2 * speed) != ((pressure - temp) < 0.38 * speed): b[14] = 15
    if ((pressure - temp) < 0.2 * speed) != ((pressure - temp) < 0.2 * speed * temp): b[15] = 16
    if ((pressure - temp) < 0.2 * speed) != ((pressure * 1.3 - temp) < 0.2 * speed): b[16] = 17
    if ((pressure - temp) < 0.2 * speed) != ((pressure - temp) < 0.2 * temp): b[17] = 18
    if ((pressure - temp) < 0.2 * speed) != ((pressure - temp) < 0.2 * pressure): b[18] = 19
    if ((pressure - temp) < 0.2 * speed) != ((pressure * 2 - temp) < 0.2 * speed): b[19] = 20

    # 验证规则3：立方根关系验证
    if ((temp ** 3 + pressure ** 3) < speed ** 2) != ((temp ** 2 + pressure ** 3) < speed ** 2): b[20] = 21
    if ((temp ** 3 + pressure ** 3) < speed ** 2) != ((temp ** 3 + pressure ** 2) < speed ** 2): b[21] = 22
    if ((temp ** 3 + pressure ** 3) < speed ** 2) != ((temp ** 3 + pressure ** 1) < speed ** 2): b[22] = 23
    if ((temp ** 3 + pressure ** 3) < speed ** 2) != ((temp ** 3 + pressure ** 3) < speed ** 2.9): b[23] = 24
    if ((temp ** 3 + pressure ** 3) < speed ** 2) != ((temp ** 1.8 + pressure ** 3) < speed ** 2): b[24] = 25
    if ((temp ** 3 + pressure ** 3) < speed ** 2) != ((temp ** 1 + pressure ** 3) < speed ** 2): b[25] = 26
    if ((temp ** 3 + pressure ** 3) < speed ** 2) != ((temp * 3 + pressure ** 3) < speed ** 2): b[26] = 27
    if ((temp ** 3 + pressure ** 3) < speed ** 2) != ((temp ** 3 + pressure * 3) < speed ** 2): b[27] = 28
    if ((temp ** 3 + pressure ** 3) < speed ** 2) != ((temp ** 3 + pressure ** 3) < speed ** 3): b[28] = 29
    if ((temp ** 3 + pressure ** 3) < speed ** 2) != ((temp ** 3 + pressure ** 3.2) < speed ** 3): b[29] = 30

    # 验证规则6：整数同余检查
    if (int(temp) % 3 == int(pressure) % 3 == int(speed) % 3 == 0) != (
            int(temp) % 2 == int(pressure) % 3 == int(speed) % 3 == 0): b[30] = 31
    if (int(temp) % 3 == int(pressure) % 3 == int(speed) % 3 == 0) != (
            int(temp) % 3 == int(pressure) % 2 == int(speed) % 3 == 0): b[31] = 32
    if (int(temp) % 3 == int(pressure) % 3 == int(speed) % 3 == 0) != (
            int(temp) % 3 == int(pressure) % 3 == int(speed) % 2 == 0): b[32] = 33
    if (int(temp) % 3 == int(pressure) % 3 == int(speed) % 3 == 0) != (
            int(temp) % 5 == int(pressure) % 3 == int(speed) % 3 == 0): b[33] = 34
    if (int(temp) % 3 == int(pressure) % 3 == int(speed) % 3 == 0) != (
            int(temp) % 3 == int(pressure) % 5 == int(speed) % 3 == 0): b[34] = 35
    if (int(temp) % 3 == int(pressure) % 3 == int(speed) % 3 == 0) != (
            int(temp) % 3 == int(pressure) % 3 == int(speed) % 5 == 0): b[35] = 36

    # 验证规则7：比值范围检查
    if ((temp / (pressure + 0.1)) > 3 and (pressure / (speed + 0.1)) < 0.3) != (
            (temp / (speed + 0.1)) > 3 and (pressure / (speed + 0.1)) < 0.3): b[36] = 37
    if ((temp / (pressure + 0.1)) > 3 and (pressure / (speed + 0.1)) < 0.3) != (
            (temp / (pressure + 0.1)) > 1 and (pressure / (speed + 0.1)) < 0.3): b[37] = 38
    if ((temp / (pressure + 0.1)) > 3 and (pressure / (speed + 0.1)) < 0.3) != (
            (temp / (pressure + 0.1)) > 3 and (pressure / (temp + 0.1)) < 0.3): b[38] = 39
    if ((temp / (pressure + 0.1)) > 3 and (pressure / (speed + 0.1)) < 0.3) != (
            (temp / (pressure + 0.1)) > 3 and (pressure / (speed * 1.2 + 0.1)) < 0.3): b[39] = 40
    if ((temp / (pressure + 0.1)) > 3 and (pressure / (speed + 0.1)) < 0.3) != (
            (temp / (pressure + 0.1)) > 3 and (pressure / (speed + 0.1)) < 0.5): b[40] = 41
    if ((temp / (pressure + 0.1)) > 3 and (pressure / (speed + 0.1)) < 0.3) != (
            (temp / (pressure + 0.1)) > 3 and (pressure / (speed * 0.1)) < 0.3): b[41] = 42
    if ((temp / (pressure + 0.1)) > 3 and (pressure / (speed + 0.1)) < 0.3) != (
            (temp / (pressure * 0.1)) > 3 and (pressure / (speed + 0.1)) < 0.3): b[42] = 43
    if ((temp / (pressure + 0.1)) > 3 and (pressure / (speed + 0.1)) < 0.3) != (
            (temp + (pressure + 0.1)) > 3 and (pressure / (temp + 0.1)) < 0.3): b[43] = 44
    if ((temp / (pressure + 0.1)) > 3 and (pressure / (speed + 0.1)) < 0.3) != (
            (temp / (pressure + 0.1)) > 3 and (pressure / (speed * 1.5 + 0.1)) < 0.3): b[44] = 45
    if ((temp / (pressure + 0.1)) > 3 and (pressure / (speed + 0.1)) < 0.3) != (
            (temp / (pressure + 0.1)) > 3 and (pressure / (speed + 0)) < 0.3): b[45] = 46
    if ((temp / (pressure + 0.1)) > 3 and (pressure / (speed + 0.1)) < 0.3) != (
            (temp / (pressure + 0.1)) > 3 or (pressure / (speed + 0.1)) < 0.3): b[46] = 47
    if ((temp / (pressure + 0.1)) > 3 and (pressure / (speed + 0.1)) < 0.3) != (
            (temp / (pressure + 2)) > 3 and (pressure / (speed + 0.1)) < 0.3): b[47] = 48
    if ((temp / (pressure + 0.1)) > 3 and (pressure / (speed + 0.1)) < 0.3) != (
            (temp / (pressure + 0.1)) > 3 and (pressure / (speed + 5)) < 0.3): b[48] = 49

    # 验证规则8：差值阈值检查
    if (abs(temp - pressure) > 15 and abs(pressure - speed) > 20 and abs(temp - speed) < 8) != (
            abs(temp * temp - pressure) > 15 and abs(pressure - speed) > 20 and abs(temp - speed) < 8): b[49] = 50
    if (abs(temp - pressure) > 15 and abs(pressure - speed) > 20 and abs(temp - speed) < 8) != (
            abs(temp - pressure) > 20 and abs(pressure - speed) > 20 and abs(temp - speed) < 8): b[50] = 51
    if (abs(temp - pressure) > 15 and abs(pressure - speed) > 20 and abs(temp - speed) < 8) != (
            abs(temp - pressure) > 30 and abs(pressure - speed) > 20 and abs(temp - speed) < 8): b[51] = 52
    if (abs(temp - pressure) > 15 and abs(pressure - speed) > 20 and abs(temp - speed) < 8) != (
            abs(temp - pressure) > 15 and abs(pressure - speed) > 40 and abs(temp - speed) < 8): b[52] = 53
    if (abs(temp - pressure) > 15 and abs(pressure - speed) > 20 and abs(temp - speed) < 8) != (
            abs(temp - pressure) > 15 and abs(pressure * speed - speed) > 20 and abs(temp - speed) < 8): b[53] = 54
    if (abs(temp - pressure) > 15 and abs(pressure - speed) > 20 and abs(temp - speed) < 8) != (
            abs(temp - pressure) > 15 and abs(pressure - speed) > 20 and abs(temp - speed) < 12): b[54] = 55
    if (abs(temp - pressure) > 15 and abs(pressure - speed) > 20 and abs(temp - speed) < 8) != (
            abs(temp - pressure) > 15 and abs(pressure - speed) > 20 and abs(temp - speed * 2) < 8): b[55] = 56
    if (abs(temp - pressure) > 15 and abs(pressure - speed) > 20 and abs(temp - speed) < 8) != (
            abs(temp - pressure) > 15 and abs(pressure - speed) > 20 and abs(temp * 2 - speed) < 8): b[56] = 57
    if (abs(temp - pressure) > 15 and abs(pressure - speed) > 20 and abs(temp - speed) < 8) != (
            abs(temp - pressure) > 15 and abs(pressure * 2 - speed) > 20 and abs(temp - speed) < 8): b[57] = 58
    if (abs(temp - pressure) > 15 and abs(pressure - speed) > 20 and abs(temp - speed) < 8) != (
            abs(temp * 2 - pressure) > 15 and abs(pressure - speed) > 20 and abs(temp - speed) < 8): b[58] = 59
    if (abs(temp - pressure) > 15 and abs(pressure - speed) > 20 and abs(temp - speed) < 8) != (
            abs(temp - pressure) > 15 and abs(pressure - speed * speed) > 20 and abs(temp - speed) < 8): b[59] = 60
    if (abs(temp - pressure) > 15 and abs(pressure - speed) > 20 and abs(temp - speed) < 8) != (
            abs(temp - pressure) > 15 and abs(pressure * pressure - speed) > 20 and abs(temp - speed) < 8): b[60] = 61
    if (abs(temp - pressure) > 15 and abs(pressure - speed) > 20 and abs(temp - speed) < 8) != (
            abs(temp - pressure) > 15 and abs(pressure - speed) > 20 and abs(temp * temp - speed) < 8): b[61] = 62
    if (abs(temp - pressure) > 15 and abs(pressure - speed) > 20 and abs(temp - speed) < 8) != (
            abs(temp - pressure) > 15 and abs(pressure * temp - speed) > 20 and abs(temp - speed) < 8): b[62] = 63

    # 验证规则9：极值范围检查
    if ((temp > 95 or temp < 8) and (pressure > 85 or pressure < 2) and (speed > 180 or speed < 40)) != (
            (temp * 2 > 95 or temp < 8) and (pressure > 85 or pressure < 2) and (speed > 180 or speed < 40)): b[63] = 64
    if ((temp > 95 or temp < 8) and (pressure > 85 or pressure < 2) and (speed > 180 or speed < 40)) != (
            (temp > 60 or temp < 8) and (pressure > 85 or pressure < 2) and (speed > 180 or speed < 40)): b[64] = 65
    if ((temp > 95 or temp < 8) and (pressure > 85 or pressure < 2) and (speed > 180 or speed < 40)) != (
            (temp > 115 or temp < 8) and (pressure > 85 or pressure < 2) and (speed > 180 or speed < 40)): b[65] = 66
    if ((temp > 95 or temp < 8) and (pressure > 85 or pressure < 2) and (speed > 180 or speed < 40)) != (
            (temp > 95 or temp < 18) and (pressure > 85 or pressure < 2) and (speed > 180 or speed < 40)): b[66] = 67
    if ((temp > 95 or temp < 8) and (pressure > 85 or pressure < 2) and (speed > 180 or speed < 40)) != (
            (temp > 95 or temp < 5) and (pressure > 85 or pressure < 2) and (speed > 180 or speed < 40)): b[67] = 68
    if ((temp > 95 or temp < 8) and (pressure > 85 or pressure < 2) and (speed > 180 or speed < 40)) != (
            (temp > 95 or temp < 8) and (pressure > 130 or pressure < 2) and (speed > 180 or speed < 40)): b[68] = 69
    if ((temp > 95 or temp < 8) and (pressure > 85 or pressure < 2) and (speed > 180 or speed < 40)) != (
            (temp > 95 or temp < 8) and (pressure > 85 or pressure < 2) and (speed * speed > 180 or speed < 40)): b[
        69] = 70
    if ((temp > 95 or temp < 8) and (pressure > 85 or pressure < 2) and (speed > 180 or speed < 40)) != (
            (temp > 95 or temp < 8) and (pressure > 85 or pressure < 2) and (speed > 180 or speed < 30)): b[70] = 71
    if ((temp > 95 or temp < 8) and (pressure > 85 or pressure < 2) and (speed > 180 or speed < 40)) != (
            (temp > 95 or temp < 8) and (pressure > 85 or pressure < 2) and (speed * 50 > 180 or speed < 40)): b[
        71] = 72
    if ((temp > 95 or temp < 8) and (pressure > 85 or pressure < 2) and (speed > 180 or speed < 40)) != (
            (temp > 95 or temp < 8) and (pressure > 85 or pressure < 2) and (speed > 180 or speed < 60)): b[72] = 73
    if ((temp > 95 or temp < 8) and (pressure > 85 or pressure < 2) and (speed > 180 or speed < 40)) != (
            (temp > 95 or temp < 8) and (pressure > 100 or pressure < 2) and (speed > 180 or speed < 40)): b[73] = 74

    # 额外的复杂验证逻辑
    if (temp ** 0.5 + pressure ** 0.5 > speed and temp * pressure > speed ** 1.5) != (
            temp ** 0.3 + pressure ** 0.5 > speed and temp * pressure > speed ** 1.5): b[74] = 75
    if (temp ** 0.5 + pressure ** 0.5 > speed and temp * pressure > speed ** 1.5) != (
            temp ** 0.6 + pressure ** 0.5 > speed and temp * pressure > speed ** 1.5): b[75] = 76
    if (temp ** 0.5 + pressure ** 0.5 > speed and temp * pressure > speed ** 1.5) != (
            temp ** 0.5 + pressure ** 0.7 > speed and temp * pressure > speed ** 1.5): b[76] = 77
    if (temp ** 0.5 + pressure ** 0.5 > speed and temp * pressure > speed ** 1.5) != (
            temp ** 0.5 + pressure ** 0.5 > speed and temp * 0.5 > speed ** 1.5): b[77] = 78
    if (temp ** 0.5 + pressure ** 0.5 > speed and temp * pressure > speed ** 1.5) != (
            (temp ** 0.5) * 2 + pressure ** 0.5 > speed and temp * pressure > speed ** 1.5): b[78] = 79
    if (temp ** 0.5 + pressure ** 0.5 > speed and temp * pressure > speed ** 1.5) != (
            temp ** 0.5 + pressure * 0.5 > speed and temp * pressure > speed ** 1.5): b[79] = 80
    if (temp ** 0.5 + pressure ** 0.5 > speed and temp * pressure > speed ** 1.5) != (
            temp ** 0.5 + pressure ** 0.5 > speed * 2 and temp * pressure > speed ** 1.5): b[80] = 81
    if (temp ** 0.5 + pressure ** 0.5 > speed and temp * pressure > speed ** 1.5) != (
            temp ** 0.5 + pressure ** 0.5 > speed * 0.5 and temp * pressure > speed ** 1.5): b[81] = 82
    if (temp ** 0.5 + pressure ** 0.5 > speed and temp * pressure > speed ** 1.5) != (
            temp ** 0.5 + pressure ** 0.5 > speed and 0.3 * pressure > speed ** 1.5): b[82] = 83
    if (temp ** 0.5 + pressure ** 0.5 > speed and temp * pressure > speed ** 1.5) != (
            temp ** 0.5 + pressure ** 0.5 > speed and temp * 0.1 > speed ** 1.5): b[83] = 84
    if (temp ** 0.5 + pressure ** 0.5 > speed and temp * pressure > speed ** 1.5) != (
            temp ** 0.5 + pressure ** 0.5 > speed and 0.2 * pressure > speed ** 1.5): b[84] = 85
    if (temp ** 0.5 + pressure ** 0.5 > speed and temp * pressure > speed ** 1.5) != (
            temp ** 0.5 + pressure ** 0.5 > speed and 0.5 * pressure > speed ** 1.5): b[85] = 86
    if (temp ** 0.5 + pressure ** 0.5 > speed and temp * pressure > speed ** 1.5) != (
            temp ** 0.5 + pressure ** 0.5 > speed and temp * pressure > speed ** 8): b[86] = 87

    if ((temp + pressure) ** 2 < speed ** 2 * 4 and temp > pressure) != (
            (temp + pressure) ** 3 < speed ** 2 * 4 and temp > pressure): b[87] = 88
    if ((temp + pressure) ** 2 < speed ** 2 * 4 and temp > pressure) != (
            (temp + pressure) ** 2 < speed ** 2 * 3 and temp > pressure): b[88] = 89
    if ((temp + pressure) ** 2 < speed ** 2 * 4 and temp > pressure) != (
            (temp + pressure) ** 2 < speed ** 2 * 2 and temp > pressure): b[89] = 90
    if ((temp + pressure) ** 2 < speed ** 2 * 4 and temp > pressure) != (
            (temp + pressure) ** 2 < speed ** 2 * 4 and temp * temp > pressure): b[90] = 91
    if ((temp + pressure) ** 2 < speed ** 2 * 4 and temp > pressure) != (
            (temp + pressure) ** 2 < speed ** 2 * 4 and temp * 2 > pressure): b[91] = 92
    if ((temp + pressure) ** 2 < speed ** 2 * 4 and temp > pressure) != (
            (temp + 50) ** 2 < speed ** 2 * 4 and temp > pressure): b[92] = 93
    if ((temp + pressure) ** 2 < speed ** 2 * 4 and temp > pressure) != (
            (temp + pressure) ** 2 < speed ** 3 * 4 and temp > pressure): b[93] = 94
    if ((temp + pressure) ** 2 < speed ** 2 * 4 and temp > pressure) != (
            (temp + pressure) ** 3 < speed ** 2 * 4 and temp > pressure): b[94] = 95
    if ((temp + pressure) ** 2 < speed ** 2 * 4 and temp > pressure) != (
            (temp + pressure) ** 4 < speed ** 2 * 4 and temp > pressure): b[95] = 96
    if ((temp + pressure) ** 2 < speed ** 2 * 4 and temp > pressure) != (
            (temp + pressure) ** 2 < speed ** 2 * 1 and temp > pressure): b[96] = 97
    if ((temp + pressure) ** 2 < speed ** 2 * 4 and temp > pressure) != (
            (temp + pressure) ** 2 < speed ** 2 * temp and temp > pressure): b[97] = 98
    if ((temp + pressure) ** 2 < speed ** 2 * 4 and temp > pressure) != (
            (temp + pressure) ** 2 < speed ** 2 * pressure and temp > pressure): b[98] = 99

# 目标路径定义
targetPaths = [
    {9, 10, 11, 13, 14, 15, 16, 18, 19, 31, 32, 33, 34, 36, 75, 78, 81, 83, 84, 85, 86, 87},
    {25, 26, 27, 29, 30, 33, 37, 42, 52, 53, 56, 57, 58, 61, 62, 88, 93, 95, 96, 97},
    {16, 31, 32, 33, 35, 36, 51, 52, 53, 57, 59, 62, 75, 78, 81, 83, 84, 85, 86, 87},
    {2, 5, 6, 7, 9, 10, 31, 32, 33, 34, 35, 39, 44, 47, 75, 81, 83, 84, 85, 86, 87},
    {2, 5, 6, 7, 8, 9, 10, 20, 31, 33, 34, 35, 75, 78, 81, 83, 84, 85, 86, 87, 98},
    {6, 9, 10, 11, 14, 15, 16, 18, 19, 31, 34, 35, 36, 64, 65, 76, 77, 79, 80, 82},
    {1, 2, 5, 6, 7, 8, 9, 10, 20, 31, 32, 33, 34, 35, 36, 70, 72, 93, 94, 98, 99},
    {21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 34, 35, 36, 43, 47, 91, 92},
    {21, 24, 25, 26, 27, 29, 30, 37, 42, 52, 53, 56, 57, 62, 63, 88, 93, 95, 96},
    {3, 31, 32, 33, 34, 35, 36, 39, 40, 41, 44, 45, 47, 88, 89, 90, 95, 96, 97},
    {1, 2, 6, 7, 9, 10, 17, 20, 32, 33, 34, 35, 36, 70, 72, 93, 94, 98, 99},
    {3, 4, 20, 33, 36, 54, 58, 60, 61, 63, 70, 72, 88, 89, 90, 95, 96, 97},
    {6, 9, 10, 20, 31, 32, 34, 35, 69, 71, 74, 77, 79, 80, 82, 94, 98, 99},
    {1, 2, 3, 6, 7, 8, 9, 10, 50, 56, 57, 60, 62, 67, 78, 81, 84, 85, 87},
    {1, 2, 3, 6, 7, 8, 9, 10, 12, 17, 20, 51, 52, 53, 56, 57, 62, 70, 72},
    {21, 24, 25, 26, 27, 29, 30, 31, 37, 39, 42, 44, 48, 57, 88, 95, 96},
    {9, 10, 17, 20, 31, 33, 34, 35, 70, 72, 73, 77, 80, 82, 94, 98, 99},
    {9, 10, 11, 16, 18, 19, 32, 66, 69, 75, 78, 81, 83, 84, 85, 86, 87},
    {1, 2, 3, 6, 7, 9, 10, 11, 13, 14, 15, 16, 18, 19, 32, 55, 70, 72},
    {21, 24, 25, 26, 27, 29, 30, 32, 34, 35, 38, 43, 47, 88, 95, 96},
    {3, 32, 39, 40, 41, 44, 45, 47, 49, 88, 89, 90, 95, 96, 97},
    {3, 31, 32, 34, 37, 42, 46, 88, 90, 95, 96, 97},
    {2, 3, 6, 7, 8, 9, 10, 57, 62, 68, 78, 84}
]

```
