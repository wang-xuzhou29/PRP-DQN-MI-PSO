```python
变异分支：
if ((energy_y * energy_z) / (energy_x + 1) > 140) != ((energy_y * energy_y) / (energy_x + 1) > 140): b[0] = 1
    if ((energy_y * energy_z) / (energy_x + 1) > 140) != ((energy_z * energy_z) / (energy_x + 1) > 140): b[1] = 2
    if ((energy_y * energy_z) / (energy_x + 1) > 140) != ((energy_y * energy_x) / (energy_x + 1) > 140): b[2] = 3
    if ((energy_y * energy_z) / (energy_x + 1) > 140) != ((energy_y * energy_z) / (energy_x + 3) > 140): b[3] = 4
    if ((energy_y * energy_z) / (energy_x + 1) > 140) != ((energy_y * energy_z) / (energy_x - 1) > 140): b[4] = 5
    if ((energy_y * energy_z) / (energy_x + 1) > 140) != ((energy_y * energy_z * 2) / (energy_y + 1) > 140): b[5] = 6
    if ((energy_y * energy_z) / (energy_x + 1) > 140) != ((energy_y * energy_z) / (energy_x + 1) > 100): b[6] = 7
    if ((energy_y * energy_z) / (energy_x + 1) > 140) != ((energy_y * energy_z) / (energy_x + 1) > 180): b[7] = 8
    if ((energy_y * energy_z) / (energy_x + 1) > 140) != ((energy_y * energy_z) / (energy_x + 10) > 140): b[8] = 9
    if ((energy_y * energy_z) / (energy_x + 1) > 140) != ((energy_y * energy_z) / (energy_x * 1) > 140): b[9] = 10
    if ((energy_y * energy_z) / (energy_x + 1) > 140) != ((energy_y * 30) / (energy_x + 1) > 140): b[10] = 11


if ((energy_z - energy_x) < 0.22 * energy_y) != ((energy_z - energy_x) < 0.22 * energy_x): b[11] = 12
if ((energy_z - energy_x) < 0.22 * energy_y) != ((energy_z - energy_x) < 0.22 * energy_z): b[12] = 13
if ((energy_z - energy_x) < 0.22 * energy_y) != ((energy_z - energy_x) < 0.32 * energy_y): b[13] = 14
if ((energy_z - energy_x) < 0.22 * energy_y) != ((energy_z - energy_x) < 0.12 * energy_y): b[14] = 15
if ((energy_z - energy_x) < 0.22 * energy_y) != ((energy_z * 2 - energy_x) < 0.22 * energy_y): b[15] = 16
if ((energy_z - energy_x) < 0.22 * energy_y) != ((energy_z - energy_x * 1.2) < 0.22 * energy_y): b[16] = 17
if ((energy_z - energy_x) < 0.22 * energy_y) != ((energy_z + energy_x) < 0.22 * energy_y): b[17] = 18
if ((energy_z - energy_x) < 0.22 * energy_y) != ((energy_z - 20) < 0.22 * energy_y): b[18] = 19
if ((energy_z - energy_x) < 0.22 * energy_y) != ((90 - energy_x) < 0.22 * energy_y): b[19] = 20
if ((energy_z - energy_x) < 0.22 * energy_y) != ((energy_z - energy_x) < 0.4 * energy_y): b[20] = 21

if ((energy_x ** 3 + energy_y ** 3) < energy_z ** 2) != ((energy_x ** 2.7 + energy_y ** 3) < energy_z ** 2): b[   21] = 22
if ((energy_x ** 3 + energy_y ** 3) < energy_z ** 2) != ((energy_x ** 3 + energy_y ** 2.6) < energy_z ** 2): b[  22] = 23
if ((energy_x ** 3 + energy_y ** 3) < energy_z ** 2) != ((energy_x ** 3 + energy_y ** 3) < energy_z ** 1.8): b[  23] = 24
if ((energy_x ** 3 + energy_y ** 3) < energy_z ** 2) != ((energy_x ** 3 - energy_y ** 3) < energy_z ** 2): b[  24] = 25
if ((energy_x ** 3 + energy_y ** 3) < energy_z ** 2) != ((energy_y ** 3 + energy_y ** 3) < energy_z ** 2): b[   25] = 26
if ((energy_x ** 3 + energy_y ** 3) < energy_z ** 2) != ((energy_z ** 3 + energy_y ** 3) < energy_z ** 2): b[  26] = 27
if ((energy_x ** 3 + energy_y ** 3) < energy_z ** 2) != ((energy_x ** 3 + energy_x ** 3) < energy_z ** 2): b[27] = 28
if ((energy_x ** 3 + energy_y ** 3) < energy_z ** 2) != ((energy_x ** 3 + energy_z ** 3) < energy_z ** 2): b[  28] = 29
if ((energy_x ** 3 + energy_y ** 3) < energy_z ** 2) != ((energy_x ** 3 + energy_y ** 3) < energy_x ** 2): b[  29] = 30
if ((energy_x ** 3 + energy_y ** 3) < energy_z ** 2) != ((energy_x ** 3 + energy_y ** 3) < energy_y ** 2): b[30] = 31
if ((energy_x ** 3 + energy_y ** 3) < energy_z ** 2) != ((energy_x ** 3 + energy_y ** 3) < energy_z ** 2.5): b[31] = 32

if ((energy_x / (energy_y + 0.01)) > 5 and (energy_y / (energy_z + 0.01)) < 0.2) != ((energy_x / (energy_z + 0.01)) > 5 and (energy_y / (energy_z + 0.01)) < 0.2): b[32] = 33
if ((energy_x / (energy_y + 0.01)) > 5 and (energy_y / (energy_z + 0.01)) < 0.2) != ( (energy_x / (energy_x + 0.01)) > 5 and (energy_y / (energy_z + 0.01)) < 0.2): b[33] = 34
if ((energy_x / (energy_y + 0.01)) > 5 and (energy_y / (energy_z + 0.01)) < 0.2) != ( (energy_z / (energy_y + 0.01)) > 5 and (energy_y / (energy_z + 0.01)) < 0.2): b[34] = 35
if ((energy_x / (energy_y + 0.01)) > 5 and (energy_y / (energy_z + 0.01)) < 0.2) != (     (energy_y / (energy_y + 0.01)) > 5 and (energy_y / (energy_z + 0.01)) < 0.2): b[35] = 36
if ((energy_x / (energy_y + 0.01)) > 5 and (energy_y / (energy_z + 0.01)) < 0.2) != (  (energy_x / (energy_y + 0.01)) > 5 and (energy_z / (energy_z + 0.01)) < 0.2): b[36] = 37
if ((energy_x / (energy_y + 0.01)) > 5 and (energy_y / (energy_z + 0.01)) < 0.2) != ( (energy_x / (energy_y + 0.01)) > 5 and (energy_x / (energy_z + 0.01)) < 0.2): b[37] = 38
if ((energy_x / (energy_y + 0.01)) > 5 and (energy_y / (energy_z + 0.01)) < 0.2) != (   (energy_x / (energy_y + 0.01)) > 5 and (energy_y / (energy_y + 0.01)) < 0.2): b[38] = 39
if ((energy_x / (energy_y + 0.01)) > 5 and (energy_y / (energy_z + 0.01)) < 0.2) != (     (energy_x / (energy_y + 0.01)) > 5 and (energy_y / (energy_x + 0.01)) < 0.2): b[39] = 40
if ((energy_x / (energy_y + 0.01)) > 5 and (energy_y / (energy_z + 0.01)) < 0.2) != (  (energy_x / (energy_y + 0.01)) > 5 and (energy_y / (energy_z + 0.01)) < 0.15): b[40] = 41
if ((energy_x / (energy_y + 0.01)) > 5 and (energy_y / (energy_z + 0.01)) < 0.2) != ( (energy_x / (energy_y + 0.01)) > 7 and (energy_y / (energy_z + 0.01)) < 0.2): b[41] = 42

if (abs(energy_x - energy_y) > 16 and abs(energy_y - energy_z) > 18 and abs(energy_x - energy_z) < 9) != (abs(energy_x * 1.2 - energy_y) > 16 and abs(energy_y - energy_z) > 18 and abs(energy_x - energy_z) < 9): b[42] = 43
if (abs(energy_x - energy_y) > 16 and abs(energy_y - energy_z) > 18 and abs(energy_x - energy_z) < 9) != (abs(energy_x - energy_y ) > 16 and abs(energy_y*2 - energy_z) > 18 and abs(energy_x - energy_z) < 9): b[43] = 44
if (abs(energy_x - energy_y) > 16 and abs(energy_y - energy_z) > 18 and abs(energy_x - energy_z) < 9) != ( abs(energy_x - energy_y) > 19 and abs(energy_y - energy_z) > 18 and abs(energy_y - energy_z) < 9): b[44] = 45
if (abs(energy_x - energy_y) > 16 and abs(energy_y - energy_z) > 18 and abs(energy_x - energy_z) < 9) != (  abs(energy_x - energy_y) > 16 and abs(energy_x - energy_z) > 18 and abs(energy_x - energy_z) < 9): b[45] = 46
if (abs(energy_x - energy_y) > 16 and abs(energy_y - energy_z) > 18 and abs(energy_x - energy_z) < 9) != ( abs(energy_x - energy_y) > 16 and abs(energy_y - energy_z) > 40 and abs(energy_x - energy_z) < 9): b[46] = 47
if (abs(energy_x - energy_y) > 16 and abs(energy_y - energy_z) > 18 and abs(energy_x - energy_z) < 9) != ( abs(energy_x - energy_y) > 16 and abs(energy_y  - energy_z) > 18 and abs(energy_x*2 - energy_z) < 9): b[47] = 48
if (abs(energy_x - energy_y) > 16 and abs(energy_y - energy_z) > 18 and abs(energy_x - energy_z) < 9) != (   abs(energy_x - energy_y) > 16 and abs(energy_y - energy_z *0.2) > 18 and abs(energy_x - energy_z) < 9): b[48] = 49
if (abs(energy_x - energy_y) > 16 and abs(energy_y - energy_z) > 18 and abs(energy_x - energy_z) < 9) != (  abs(energy_x - energy_y) > 16 and abs(energy_y - energy_z) > 18 and abs(energy_x * 1.5 - energy_z) < 9): b[49] = 50
if (abs(energy_x - energy_y) > 16 and abs(energy_y - energy_z) > 18 and abs(energy_x - energy_z) < 9) != ( abs(energy_x - energy_y) > 16 and abs(energy_y - energy_z) > 18 and abs(energy_x - energy_z * 0.87) < 9): b[50] = 51
if (abs(energy_x - energy_y) > 16 and abs(energy_y - energy_z) > 18 and abs(energy_x - energy_z) < 9) != ( abs(energy_x - energy_y) > 16 and abs(energy_y - energy_z) > 18 and abs(energy_x - energy_z) < 7.8): b[51] = 52

if ((energy_x > 95 or energy_x < 5) and (energy_y > 90 or energy_y < 3) and (energy_z > 85 or energy_z < 2)) != ( (energy_x > 95 or energy_x < 5) and (energy_y * energy_y > 90 or energy_y < 3) and (  energy_z > 85 or energy_z < 2)): b[52] = 53
if ((energy_x > 95 or energy_x < 5) and (energy_y > 90 or energy_y < 3) and (energy_z > 85 or energy_z < 2)) != (    (energy_x > 95 or energy_x < 5) and (energy_y * energy_x > 90 or energy_y < 3) and ( energy_z > 85 or energy_z < 2)): b[53] = 54
if ((energy_x > 95 or energy_x < 5) and (energy_y > 90 or energy_y < 3) and (energy_z > 85 or energy_z < 2)) != ( (energy_x > 95 or energy_x < 5) and (energy_y * energy_z > 90 or energy_y < 3) and (  energy_z > 85 or energy_z < 2)): b[54] = 55
if ((energy_x > 95 or energy_x < 5) and (energy_y > 90 or energy_y < 3) and (energy_z > 85 or energy_z < 2)) != ((energy_x > 95 or energy_x < 5) and (energy_y*80 > 90 or energy_y < 3) and (energy_z > 85 or energy_z < 2)): b[55] = 56
if ((energy_x > 95 or energy_x < 5) and (energy_y > 90 or energy_y < 3) and (energy_z > 85 or energy_z < 2)) != ((energy_x * energy_y > 95 or energy_x < 5) and (energy_y > 90 or energy_y < 3) and ( energy_z > 85 or energy_z < 2)): b[56] = 57
if ((energy_x > 95 or energy_x < 5) and (energy_y > 90 or energy_y < 3) and (energy_z > 85 or energy_z < 2)) != ( (energy_x * energy_x > 95 or energy_x < 5) and (energy_y > 90 or energy_y < 3) and ( energy_z > 85 or energy_z < 2)): b[57] = 58
if ((energy_x > 95 or energy_x < 5) and (energy_y > 90 or energy_y < 3) and (energy_z > 85 or energy_z < 2)) != (  (energy_x * energy_z > 95 or energy_x < 5) and (energy_y > 90 or energy_y < 3) and (energy_z > 85 or energy_z < 2)): b[58] = 59
if ((energy_x > 95 or energy_x < 5) and (energy_y > 90 or energy_y < 3) and (energy_z > 85 or energy_z < 2)) != (  (energy_x*50 > 95 or energy_x < 5) and (energy_y > 90 or energy_y < 3) and (  energy_z  > 85 or energy_z < 2)): b[59] = 60
if ((energy_x > 95 or energy_x < 5) and (energy_y > 90 or energy_y < 3) and (energy_z > 85 or energy_z < 2)) != ( (energy_x > 95 or energy_x < 5) and (energy_y*40 > 90 or energy_y < 3) and (energy_z  > 85 or energy_z < 2)): b[60] = 61
if ((energy_x > 95 or energy_x < 5) and (energy_y > 90 or energy_y < 3) and (energy_z > 85 or energy_z < 2)) != (  (energy_x > 95 or energy_x < 5) and (energy_y > 90 or energy_y < 3) and (energy_z * energy_z > 85 or energy_z < 2)): b[61] = 62
if ((energy_x > 95 or energy_x < 5) and (energy_y > 90 or energy_y < 3) and (energy_z > 85 or energy_z < 2)) != ((energy_x*40 > 95 or energy_x < 5) and (energy_y > 90 or energy_y < 3) and (energy_z > 85 or energy_z < 2)): b[ 62] = 63

if (energy_x ** 0.7 + energy_y ** 0.7 > energy_z ** 0.9 and energy_x + energy_y + energy_z < 180) != (energy_x ** 0.6 + energy_y ** 0.7 > energy_z ** 0.9 and energy_x + energy_y + energy_z < 180): b[63] = 64
if (energy_x ** 0.7 + energy_y ** 0.7 > energy_z ** 0.9 and energy_x + energy_y + energy_z < 180) != (energy_x ** 0.7 + energy_y ** 0.7 > energy_z ** 0.9 and energy_z + energy_y + energy_z < 180): b[64] = 65
if (energy_x ** 0.7 + energy_y ** 0.7 > energy_z ** 0.9 and energy_x + energy_y + energy_z < 180) != ( energy_x ** 0.7 + energy_y ** 0.8 > energy_z ** 0.9 and energy_x + energy_y + energy_z < 180): b[65] = 66
if (energy_x ** 0.7 + energy_y ** 0.7 > energy_z ** 0.9 and energy_x + energy_y + energy_z < 180) != (energy_x ** 0.7 + energy_y ** 0.7 > energy_z ** 0.8 and energy_x + energy_y + energy_z < 180): b[66] = 67
if (energy_x ** 0.7 + energy_y ** 0.7 > energy_z ** 0.9 and energy_x + energy_y + energy_z < 180) != (energy_x ** 0.7 + energy_z ** 0.7 > energy_z ** 0.9 and energy_x + energy_y + energy_z < 180): b[67] = 68
if (energy_x ** 0.7 + energy_y ** 0.7 > energy_z ** 0.9 and energy_x + energy_y + energy_z < 180) != (energy_y ** 0.7 + energy_y ** 0.7 > energy_z ** 0.9 and energy_x + energy_y + energy_z < 180): b[68] = 69
if (energy_x ** 0.7 + energy_y ** 0.7 > energy_z ** 0.9 and energy_x + energy_y + energy_z < 180) != ( energy_z ** 0.7 + energy_y ** 0.7 > energy_z ** 0.9 and energy_x + energy_y + energy_z < 180): b[69] = 70
if (energy_x ** 0.7 + energy_y ** 0.7 > energy_z ** 0.9 and energy_x + energy_y + energy_z < 180) != (energy_x ** 0.7 + energy_x ** 0.7 > energy_z ** 0.9 and energy_x + energy_y + energy_z < 180): b[70] = 71
if (energy_x ** 0.7 + energy_y ** 0.7 > energy_z ** 0.9 and energy_x + energy_y + energy_z < 180) != ( energy_x ** 0.7 + energy_y ** 0.7 > energy_x ** 0.9 and energy_x + energy_y + energy_z < 180): b[71] = 72
if (energy_x ** 0.7 + energy_y ** 0.7 > energy_z ** 0.9 and energy_x + energy_y + energy_z < 180) != (energy_x ** 0.7 + energy_y ** 0.7 > energy_z ** 0.9 and energy_y + energy_y + energy_z < 180): b[72] = 73
if (energy_x ** 0.7 + energy_y ** 0.7 > energy_z ** 0.9 and energy_x + energy_y + energy_z < 180) != (energy_x ** 0.7 + energy_y ** 0.7 > energy_z ** 0.9 and energy_z + energy_y + energy_z < 180): b[73] = 74
if (energy_x ** 0.7 + energy_y ** 0.7 > energy_z ** 0.9 and energy_x + energy_y + energy_z < 180) != (energy_x ** 0.7 + energy_y ** 0.7 > energy_z ** 0.9 and energy_x + energy_x + energy_z < 180): b[74] = 75

if ((energy_x + energy_y) ** 1.3 < energy_z ** 1.6 and energy_x + energy_y + energy_z / 3 > 35) != ((energy_y + energy_y) ** 1.3 < energy_z ** 1.6 and energy_x + energy_y + energy_z / 3 > 35): b[75] = 76
if ((energy_x + energy_y) ** 1.3 < energy_z ** 1.6 and energy_x + energy_y + energy_z / 3 > 35) != ((energy_z + energy_y) ** 1.3 < energy_z ** 1.6 and energy_x + energy_y + energy_z / 3 > 35): b[76] = 77
if ((energy_x + energy_y) ** 1.3 < energy_z ** 1.6 and energy_x + energy_y + energy_z / 3 > 35) != ((energy_x + energy_x) ** 1.3 < energy_z ** 1.6 and energy_x + energy_y + energy_z / 3 > 35): b[77] = 78
if ((energy_x + energy_y) ** 1.3 < energy_z ** 1.6 and energy_x + energy_y + energy_z / 3 > 35) != ((energy_x + energy_z) ** 1.3 < energy_z ** 1.6 and energy_x + energy_y + energy_z / 3 > 35): b[78] = 79
if ((energy_x + energy_y) ** 1.3 < energy_z ** 1.6 and energy_x + energy_y + energy_z / 3 > 35) != ((energy_x + 20) ** 1.3 < energy_z ** 1.6 and energy_x + energy_y + energy_z / 3 > 35): b[79] = 80
if ((energy_x + energy_y) ** 1.3 < energy_z ** 1.6 and energy_x + energy_y + energy_z / 3 > 35) != ((energy_x + energy_y) ** 1 < energy_z ** 1.6 and energy_x + energy_y + energy_z / 3 > 35): b[80] = 81
if ((energy_x + energy_y) ** 1.3 < energy_z ** 1.6 and energy_x + energy_y + energy_z / 3 > 35) != ((energy_x + energy_y) ** 1.3 < energy_z ** 1.7 and energy_x + energy_y + energy_z / 3 > 35): b[81] = 82
if ((energy_x + energy_y) ** 1.3 < energy_z ** 1.6 and energy_x + energy_y + energy_z / 3 > 35) != ((energy_x + energy_y) ** 1.2 < energy_z ** 1.6 and energy_x + energy_y + energy_z / 3 > 35): b[82] = 83
if ((energy_x + energy_y) ** 1.3 < energy_z ** 1.6 and energy_x + energy_y + energy_z / 3 > 35) != ((energy_x + energy_y) ** 1.3 < energy_z ** 1.6 and energy_y + energy_y + energy_z / 3 > 35): b[83] = 84
if ((energy_x + energy_y) ** 1.3 < energy_z ** 1.6 and energy_x + energy_y + energy_z / 3 > 35) != ((energy_x + energy_y) ** 1.3 < energy_z ** 1.6 and energy_x + energy_y - energy_z / 3 > 35): b[84] = 85

路径:
A1=[2,3,6,8,9,11,12,13,15,16,18,19,20,25,48,50,65,68,71,73,74,78,79,80,81,82,83]
A2=[6,16,18,19,26,33,34,36,37,38,39,41,42,45,46,48,49,50,51,52,53,54,55,56,61]
A3=[2,3,4,6,8,9,11,12,13,15,16,18,20,25,48,50,68,71,73,78,79,80,81,82,83]
A4=[6,12,13,17,20,21,26,33,34,36,37,38,39,45,46,48,49,50,53,54,55,56,61]
A5=[1,16,18,19,20,25,45,46,47,48,50,64,65,68,71,73,74,78,79,80,81,82,83]
A6=[2,3,4,6,8,9,11,12,13,15,16,18,19,20,25,50,51,75,78,79,80,81,82,83]
A7=[12,13,17,20,26,33,34,36,37,38,39,45,46,48,49,50,51,57,58,59,60,63]
A8=[16,18,19,20,25,43,45,46,47,48,50,51,52,65,74,75,78,79,80,81,82,83]
A9=[16,18,19,20,25,43,45,46,47,48,50,64,68,71,73,77,78,79,80,81,82,83]
A10=[2,3,6,8,9,11,14,17,21,25,48,64,65,68,71,73,74,78,79,80,81,82,83]
A11=[1,7,12,13,16,18,19,20,25,50,51,65,68,71,73,74,78,79,80,81,82,83]
A12=[1,5,7,10,12,13,15,16,18,20,25,48,50,51,68,71,73,78,79,80,81]
A13=[12,13,17,20,26,33,34,36,37,38,39,45,46,48,50,58,59,60,63,84]
A14=[18,19,20,26,33,34,36,37,38,39,41,64,69,70,72,75,76,77,81,83]
A15=[16,18,19,20,26,33,34,35,36,37,38,39,41,67,68,71,78,79,80,84]
A16=[3,6,12,13,15,16,18,20,25,28,62,65,68,71,73,74,78,79,80,81]
A17=[16,18,19,20,26,33,34,36,37,38,39,67,68,71,78,79,80,84,85]
A18=[18,19,20,26,33,34,36,37,38,39,66,67,68,71,76,77,81,82,83]
A19=[2,6,24,27,28,29,30,31,33,34,36,37,39,58,59,60,63,85]
A20=[2,6,22,26,32,33,34,36,37,38,39,58,59,60,63,85]
A21=[12,13,14,17,21,26,44,45,46,47,48,49,50,84,85]
A22=[18,19,20,26,40,66,67,68,71,76,77,81,82,83]
A23=[3,23,25,28,32,35,53,54,55,56,61,85]

```
