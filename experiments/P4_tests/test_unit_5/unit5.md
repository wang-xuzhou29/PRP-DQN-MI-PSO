

```python
变异分支：
    if ((fault_y * fault_z) / (fault_x + 1) > 85) != ((fault_y * fault_z) / (fault_x + 2) > 85): b[0] = 1
    if ((fault_y * fault_z) / (fault_x + 1) > 85) != ((fault_y * fault_z) / (fault_y + 1) > 85): b[1] = 2
    if ((fault_y * fault_z) / (fault_x + 1) > 85) != ((fault_y * fault_z) / (fault_z + 1) > 85): b[2] = 3
    if ((fault_y * fault_z) / (fault_x + 1) > 85) != ((fault_y * fault_z) / (fault_x * 1) > 85): b[3] = 4
    if ((fault_y * fault_z) / (fault_x + 1) > 85) != ((fault_y * fault_y) / (fault_x + 1) > 85): b[4] = 5
    if ((fault_y * fault_z) / (fault_x + 1) > 85) != ((fault_y * fault_x) / (fault_x + 1) > 85): b[5] = 6
    if ((fault_y * fault_z) / (fault_x + 1) > 85) != ((fault_x * fault_z) / (fault_x + 1) > 85): b[6] = 7
    if ((fault_y * fault_z) / (fault_x + 1) > 85) != ((fault_z * fault_z) / (fault_x + 1) > 85): b[7] = 8
    if ((fault_y * fault_z) / (fault_x + 1) > 85) != ((10 * fault_z) / (fault_x + 1) > 85): b[8] = 9
    if ((fault_y * fault_z) / (fault_x + 1) > 85) != ((fault_y * fault_z) - (fault_x + 1) > 85): b[9] = 10
    if ((fault_y * fault_z) / (fault_x + 1) > 85) != ((fault_y * fault_z) / (fault_x + 1) > 105): b[10] = 11
    if ((fault_y * fault_z) / (fault_x + 1) > 85) != ((fault_y * fault_z) / (fault_x - 1) > 85): b[11] = 12
    if ((fault_y * fault_z) / (fault_x + 1) > 85) != ((fault_y * 2 * fault_z) / (fault_x + 1) > 85): b[12] = 13
    if ((fault_y * fault_z) / (fault_x + 1) > 85) != ((fault_y / 2 * fault_z) / (fault_x + 1) > 85): b[13] = 14
    if ((fault_y * fault_z) / (fault_x + 1) > 85) != ((fault_y * 15) / (fault_x + 1) > 85): b[14] = 15
if ((fault_z - fault_x) < 0.25 * fault_y) != ((fault_y - fault_x) < 0.25 * fault_y): b[15] = 16
if ((fault_z - fault_x) < 0.25 * fault_y) != ((fault_z * 1.2 - fault_x) < 0.25 * fault_y): b[16] = 17
if ((fault_z - fault_x) < 0.25 * fault_y) != ((fault_z - fault_x) < 0.3 * fault_y): b[17] = 18
if ((fault_z - fault_x) < 0.25 * fault_y) != ((fault_z - fault_x) < 0.4 * fault_y): b[18] = 19
if ((fault_z - fault_x) < 0.25 * fault_y) != ((fault_z - fault_x) < 0.25 * fault_x): b[19] = 20
if ((fault_z - fault_x) < 0.25 * fault_y) != ((fault_z - fault_x) < 0.25 * fault_z): b[20] = 21
if ((fault_z - fault_x) < 0.25 * fault_y) != ((fault_y - fault_x) < 0.25 * fault_y): b[21] = 22
if ((fault_z - fault_x) < 0.25 * fault_y) != ((fault_z - fault_x * 0.7) < 0.25 * fault_y): b[22] = 23
if ((fault_z - fault_x) < 0.25 * fault_y) != ((fault_z - fault_x) < 0.25 ** fault_y): b[23] = 24
if ((fault_z - fault_x) < 0.25 * fault_y) != ((fault_z / fault_x) < 0.25 * fault_y): b[24] = 25
if ((fault_z - fault_x) < 0.25 * fault_y) != ((fault_z + fault_x) < 0.25 * fault_y): b[25] = 26
if ((fault_z - fault_x) < 0.25 * fault_y) != ((80 - fault_x) < 0.25 * fault_y): b[26] = 27
if ((fault_z - fault_x) < 0.25 * fault_y) != ((fault_z - 60) < 0.25 * fault_y): b[27] = 28
if ((fault_x ** 3 + fault_y ** 3) < fault_z ** 2) != ((fault_y ** 3 + fault_y ** 3) < fault_z ** 2): b[28] = 29
if ((fault_x ** 3 + fault_y ** 3) < fault_z ** 2) != ((fault_x * 3 + fault_y ** 3) < fault_z ** 2): b[29] = 30
if ((fault_x ** 3 + fault_y ** 3) < fault_z ** 2) != ((fault_x ** 3 - fault_y ** 3) < fault_z ** 2): b[30] = 31
if ((fault_x ** 3 + fault_y ** 3) < fault_z ** 2) != ((fault_x ** 3 + fault_x ** 3) < fault_z ** 2): b[31] = 32
if ((fault_x ** 3 + fault_y ** 3) < fault_z ** 2) != ((fault_x ** 3 + fault_y * 3) < fault_z ** 2): b[32] = 33
if ((fault_x ** 3 + fault_y ** 3) < fault_z ** 2) != ((fault_x ** 3 + fault_z ** 3) < fault_z ** 2): b[33] = 34
if ((fault_x ** 3 + fault_y ** 3) < fault_z ** 2) != ((fault_z ** 3 + fault_y ** 3) < fault_z ** 2): b[34] = 35
if ((fault_x ** 3 + fault_y ** 3) < fault_z ** 2) != ((fault_x ** 3 + fault_y ** 3) < fault_x ** 2): b[35] = 36
if ((fault_x ** 3 + fault_y ** 3) < fault_z ** 2) != ((fault_x ** 3 + fault_y ** 3) < fault_y ** 2): b[36] = 37
if ((fault_x ** 3 + fault_y ** 3) < fault_z ** 2) != ((fault_x ** 3 + fault_y ** 3) < fault_z * 2): b[37] = 38
if ((fault_x ** 3 + fault_y ** 3) < fault_z ** 2) != ((fault_x ** 3 + fault_y ** 2.8) < fault_z ** 2): b[38] = 39
if ((fault_x ** 3 + fault_y ** 3) < fault_z ** 2) != ((fault_x ** 2.4 + fault_y ** 3) < fault_z ** 2): b[39] = 40
if ((fault_x ** 3 + fault_y ** 3) < fault_z ** 2) != ((fault_x ** 3 + fault_y ** 3) < fault_z ** 2.3): b[40] = 41
if ((fault_x ** 3 + fault_y ** 3) < fault_z ** 2) != ((20 ** 3 + fault_y ** 3) < fault_z ** 2): b[41] = 42
if ((fault_x ** 3 + fault_y ** 3) < fault_z ** 2) != ((fault_x ** 3 + 15 ** 3) < fault_z ** 2): b[42] = 43
if ((fault_x ** 3 + fault_y ** 3) < fault_z ** 2) != ((fault_x ** 3 + fault_y ** 3) < 50 ** 2): b[43] = 44
if ((fault_x / (fault_y + 0.01)) > 4.5 and (fault_y / (fault_z + 0.01)) < 0.22) != (       (fault_z / (fault_y + 0.01)) > 4.5 and (fault_y / (fault_z + 0.01)) < 0.22): b[44] = 45
if ((fault_x / (fault_y + 0.01)) > 4.5 and (fault_y / (fault_z + 0.01)) < 0.22) != (       (fault_x / (fault_y + 0.01)) > 1 and (fault_y / (fault_z + 0.01)) < 0.22): b[45] = 46
if ((fault_x / (fault_y + 0.01)) > 4.5 and (fault_y / (fault_z + 0.01)) < 0.22) != (      (fault_x / (fault_y + 0.01)) > 0.5 and (fault_y / (fault_z + 0.01)) < 0.22): b[46] = 47
if ((fault_x / (fault_y + 0.01)) > 4.5 and (fault_y / (fault_z + 0.01)) < 0.22) != (    (fault_x / (fault_x + 0.01)) > 4.5 and (fault_y / (fault_z + 0.01)) < 0.22): b[47] = 48
if ((fault_x / (fault_y + 0.01)) > 4.5 and (fault_y / (fault_z + 0.01)) < 0.22) != (      (fault_x*0.6 / (fault_y + 0.01)) > 4.5 and (fault_y / (fault_z + 0.01)) < 0): b[48] = 49
if ((fault_x / (fault_y + 0.01)) > 4.5 and (fault_y / (fault_z + 0.01)) < 0.22) != (     (fault_x / (fault_y + 0.01)) > 7.5 and (fault_y / (fault_z + 0.01)) < 0.22): b[49] = 50
if ((fault_x / (fault_y + 0.01)) > 4.5 and (fault_y / (fault_z + 0.01)) < 0.22) != (     (fault_x / (fault_y + 0.01)) > 4.5 and (fault_z / (fault_z + 0.01)) < 0.22): b[50] = 51
if ((fault_x / (fault_y + 0.01)) > 4.5 and (fault_y / (fault_z + 0.01)) < 0.22) != (    (fault_x / (fault_y + 0.01)) > 4.5 and (fault_x / (fault_z + 0.01)) < 0.22): b[51] = 51
if ((fault_x / (fault_y + 0.01)) > 4.5 and (fault_y / (fault_z + 0.01)) < 0.22) != (  (fault_x / (fault_y + 0.01)) > 4.5 and (fault_y / (fault_x + 0.01)) < 0.22): b[52] = 53
if ((fault_x / (fault_y + 0.01)) > 4.5 and (fault_y / (fault_z + 0.01)) < 0.22) != (   (fault_x / (fault_y + 0.01)) > 4.5 and (fault_y / (fault_y + 0.01)) < 0.22): b[53] = 54
if ((fault_x / (fault_y + 0.01)) > 4.5 and (fault_y / (fault_z + 0.01)) < 0.22) != ((fault_x / (fault_y + 0.01)) > 2.5 and (fault_y / (fault_z + 0.01)) < 0.22): b[54] = 55
if ((fault_x / (fault_y + 0.01)) > 4.5 and (fault_y / (fault_z + 0.01)) < 0.22) != ((fault_x / (fault_y + 0.01)) > 3.5 and (fault_y / (fault_z + 0.01)) < 0.22): b[55] = 56
if (abs(fault_x - fault_y) > 13 and abs(fault_y - fault_z) > 17 and abs(fault_x - fault_z) < 8) != ( abs(fault_x - fault_y) > 13 and abs(fault_y - fault_z) > 17 and abs(fault_x*2 - fault_z) < 8): b[56] = 57
if (abs(fault_x - fault_y) > 13 and abs(fault_y - fault_z) > 17 and abs(fault_x - fault_z) < 8) != (abs(fault_x  + fault_y) > 13 and abs(fault_y - fault_z) > 17 and abs(fault_x - fault_z) < 8): b[ 57] = 58
if (abs(fault_x - fault_y) > 13 and abs(fault_y - fault_z) > 17 and abs(fault_x - fault_z) < 8) != (abs(fault_x - fault_z) > 13 and abs(fault_y - fault_z) > 17 and abs(fault_x - fault_z) < 8): b[58] = 59
if (abs(fault_x - fault_y) > 13 and abs(fault_y - fault_z) > 17 and abs(fault_x - fault_z) < 8) != (  abs(fault_x - fault_y * 1.1) > 13 and abs(fault_y - fault_z) > 17 and abs(fault_x - fault_z) < 8): b[59] = 60
if (abs(fault_x - fault_y) > 13 and abs(fault_y - fault_z) > 17 and abs(fault_x - fault_z) < 8) != ( abs(fault_x - fault_y) > 18 and abs(fault_y - fault_z) > 17 and abs(fault_x - fault_z) < 8): b[60] = 61
if (abs(fault_x - fault_y) > 13 and abs(fault_y - fault_z) > 17 and abs(fault_x - fault_z) < 8) != (abs(fault_x - fault_y) > 13 and abs(fault_x - fault_z) > 17 and abs(fault_x - fault_z) < 8): b[61] = 62
if (abs(fault_x - fault_y) > 13 and abs(fault_y - fault_z) > 17 and abs(fault_x - fault_z) < 8) != (abs(fault_x - fault_y) > 13 and abs(fault_y - fault_z * 0.9) > 17 and abs(fault_x - fault_z) < 8): b[62] = 63
if (abs(fault_x - fault_y) > 13 and abs(fault_y - fault_z) > 17 and abs(fault_x - fault_z) < 8) != (abs(fault_x - fault_y) > 13 and abs(fault_y * 1.4 - fault_z) > 17 and abs(fault_x - fault_z) < 8): b[ 63] = 64
if (abs(fault_x - fault_y) > 13 and abs(fault_y - fault_z) > 17 and abs(fault_x - fault_z) < 8) != (  abs(fault_x - fault_y) > 13 and abs(fault_y - fault_z) > 20 and abs(fault_x - fault_z) < 8): b[64] = 65
if (abs(fault_x - fault_y) > 13 and abs(fault_y - fault_z) > 17 and abs(fault_x - fault_z) < 8) != (abs(fault_x - fault_y) > 13 and abs(fault_y - fault_z) > 17 and abs(fault_y - fault_z) < 8): b[65] = 66
if (abs(fault_x - fault_y) > 13 and abs(fault_y - fault_z) > 17 and abs(fault_x - fault_z) < 8) != (abs(fault_x - fault_y) > 13 and abs(fault_y - fault_z) > 17 and abs(fault_x * 1.5 - fault_z) < 8): b[66] = 67
if (abs(fault_x - fault_y) > 13 and abs(fault_y - fault_z) > 17 and abs(fault_x - fault_z) < 8) != (abs(fault_x - fault_y) > 13 and abs(fault_y - fault_z) > 17 and abs(fault_x - fault_z) < 4): b[67] = 68
if ((fault_x > 92 or fault_x < 6) and (fault_y > 87 or fault_y < 3) and (fault_z > 83 or fault_z < 2)) != (
        (fault_x * 3 > 92 or fault_x < 6) and (fault_y > 87 or fault_y < 3) and (fault_z > 83 or fault_z < 2)): b[
    68] = 69
if ((fault_x > 92 or fault_x < 6) and (fault_y > 87 or fault_y < 3) and (fault_z > 83 or fault_z < 2)) != (
        (fault_x * fault_x > 92 or fault_x < 6) and (fault_y > 87 or fault_y < 3) and (
        fault_z > 83 or fault_z < 2)): b[69] = 70
if ((fault_x > 92 or fault_x < 6) and (fault_y > 87 or fault_y < 3) and (fault_z > 83 or fault_z < 2)) != (
        (fault_x * fault_y > 92 or fault_x < 6) and (fault_y > 87 or fault_y < 3) and (
        fault_z > 83 or fault_z < 2)): b[70] = 71
if ((fault_x > 92 or fault_x < 6) and (fault_y > 87 or fault_y < 3) and (fault_z > 83 or fault_z < 2)) != (
        (fault_x * fault_z > 92 or fault_x < 6) and (fault_y > 87 or fault_y < 3) and (
        fault_z > 83 or fault_z < 2)): b[71] = 72
if ((fault_x > 92 or fault_x < 6) and (fault_y > 87 or fault_y < 3) and (fault_z > 83 or fault_z < 2)) != (
        (fault_x > 92 or fault_x < 4) and (fault_y > 87 or fault_y < 3) and (fault_z > 83 or fault_z < 2)): b[
    72] = 73
if ((fault_x > 92 or fault_x < 6) and (fault_y > 87 or fault_y < 3) and (fault_z > 83 or fault_z < 2)) != (
        (fault_x > 92 or fault_x < 6) and (fault_y * fault_x > 87 or fault_y < 3) and (
        fault_z > 83 or fault_z < 2)): b[73] = 74
if ((fault_x > 92 or fault_x < 6) and (fault_y > 87 or fault_y < 3) and (fault_z > 83 or fault_z < 2)) != (
        (fault_x > 92 or fault_x < 6) and (fault_y * fault_y > 87 or fault_y < 3) and (
        fault_z > 83 or fault_z < 2)): b[74] = 75
if ((fault_x > 92 or fault_x < 6) and (fault_y > 87 or fault_y < 3) and (fault_z > 83 or fault_z < 2)) != (
        (fault_x > 92 or fault_x < 6) and (fault_y * fault_z > 87 or fault_y < 3) and (
        fault_z > 83 or fault_z < 2)): b[75] = 76
if ((fault_x > 92 or fault_x < 6) and (fault_y > 87 or fault_y < 3) and (fault_z > 83 or fault_z < 2)) != (
        (fault_x > 92 or fault_x < 6) and (fault_y > 87 or fault_y < 3) and (
        fault_z * fault_x > 83 or fault_z < 2)): b[76] = 77
if ((fault_x > 92 or fault_x < 6) and (fault_y > 87 or fault_y < 3) and (fault_z > 83 or fault_z < 2)) != (
        (fault_x > 92 or fault_x < 6) and (fault_y > 87 or fault_y < 3) and (
        fault_z * fault_y > 83 or fault_z < 2)): b[77] = 78
if ((fault_x > 92 or fault_x < 6) and (fault_y > 87 or fault_y < 3) and (fault_z > 83 or fault_z < 2)) != (
        (fault_x > 92 or fault_x < 6) and (fault_y > 87 or fault_y < 3) and (
        fault_z * fault_z > 83 or fault_z < 2)): b[78] = 79
if ((fault_x > 92 or fault_x < 6) and (fault_y > 87 or fault_y < 3) and (fault_z > 83 or fault_z < 2)) != (
        (fault_x > 92 or fault_x < 6) and (fault_y > 87 or fault_y < 3) and (fault_z * 50 > 83 or fault_z < 2)): b[
    79] = 80
if ((fault_x > 92 or fault_x < 6) and (fault_y > 87 or fault_y < 3) and (fault_z > 83 or fault_z < 2)) != (
        (fault_x * 60 > 92 or fault_x < 6) and (fault_y > 87 or fault_y < 3) and (fault_z > 83 or fault_z < 2)): b[
    80] = 81
if ((fault_x > 92 or fault_x < 6) and (fault_y > 87 or fault_y < 3) and (fault_z > 83 or fault_z < 2)) != (
        (fault_x > 92 or fault_x < 6) and (fault_y * 75 > 87 or fault_y < 3) and (fault_z > 83 or fault_z < 2)): b[
    81] = 82
   路径：
A1=[2,7,8,10,13,17,23,24,26,28,57,59,60,61,62,63,64,66,67,68,74,75,76,82]
A2=[2,7,9,14,15,16,22,23,24,26,31,57,59,62,66,67,68,69,70,71,72,81]
A3=[2,7,8,10,13,16,22,25,27,29,30,42,45,46,47,55,56,57,74,75,76,82]
A4=[2,7,8,10,16,22,25,27,29,30,42,48,49,50,51,52,54,57,74,75,76,82]
A5=[2,7,9,11,14,15,23,24,26,31,57,59,62,65,66,67,69,70,71,72,81]
A6=[2,7,8,10,16,20,21,22,25,27,30,48,49,50,51,52,54,74,75,76,82]
A7=[1,2,3,6,7,8,9,11,14,15,16,22,24,26,27,31,57,59,62,66,67,68]
A8=[3,5,6,9,14,15,16,22,25,29,30,31,33,40,41,42,43,45,46,47]
A9=[3,5,6,9,14,15,16,22,29,30,31,33,39,40,41,42,43,45,46,47]
A10=[3,5,6,9,14,15,16,22,32,34,35,36,37,38,44,45,46,47]
A11=[3,4,5,6,10,12,13,16,22,26,27,31,57,59,62,66,67,68]
A12=[7,10,17,23,26,28,53,57,59,62,66,67,68,74,75,76,82]
A13=[7,8,10,16,18,19,20,21,22,25,27,57,59,62,66,67,68]
A14=[2,7,16,17,20,21,22,24,26,27,31,32,33,78,79,80]
A15=[2,6,7,9,18,19,25,28,31,32,33,77,78,79,80]
A16=[2,6,7,25,31,32,33,43,73]
A17=[2,7,8,9,11,14,15,26,31,58,60]

```
