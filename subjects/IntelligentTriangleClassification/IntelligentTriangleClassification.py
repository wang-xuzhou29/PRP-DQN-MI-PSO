from typing import List, Set, Dict, Tuple
class TriangleRuleAnalyzer:
    def __init__(self):
        self.rule_descriptions = {
        }

    def execute_rules(self, sides: List[int]) -> Set[int]:

        x, y, z = sides[0], sides[1], sides[2]
        triggered = set()

        if (x > y):triggered.add(1)
        if (x > 5):triggered.add(2)
        if (x > z):triggered.add(3)
        if (x * x > y):triggered.add(4)
        if (x > y):triggered.add(5)
        if (x > y * y):triggered.add(6)

        if (y > z):triggered.add(7)
        if (x > 10):triggered.add(8)
        if (x > z):triggered.add(9)
        if (x * x > z):triggered.add(10)
        if (x > z):triggered.add(11)
        if (x > z * z):triggered.add(12)

        if (y > z):triggered.add(13)
        if (y > 8):triggered.add(14)
        if (y > z):triggered.add(15)
        if (y * y > z):triggered.add(16)
        if (y > z):triggered.add(17)
        if (y > z * z):triggered.add(18)
        if (y > z):triggered.add(19)
        if (10 > z):triggered.add(20)

        if (x + y <= z):triggered.add(21)
        if (x + y <= z * x):triggered.add(22)
        if (x + y <= z):triggered.add(23)
        if (x + y <= z * y):triggered.add(24)
        if (x + y <= z):triggered.add(25)
        if (x == y == z):triggered.add(26)
        return triggered

    def analyze_triangle(self, sides: List[int]) -> Dict:

        triggered_rules = self.execute_rules(sides)

        a, b, c = sorted(sides)
        triangle_type = ""

        if a + b <= c:
            triangle_type = "Non-triangular (not satisfying the triangle inequality)"
        elif a == b == c:
            triangle_type = "Equilateral triangle"
        elif a == b or b == c or a == c:
            triangle_type = "Isosceles triangle"
        else:
            triangle_type = "Scalene triangle"


        angle_type = ""
        if a * a + b * b == c * c:
            angle_type = "Right triangle"
        elif a * a + b * b > c * c:
            angle_type = "Acute triangle"
        else:
            angle_type = "Obtuse triangle"

        return {
            "sides": sides,
            "triangle_type": triangle_type,
            "angle_type": angle_type,
            "triggered_rules": triggered_rules,
            "rule_details": {rule: self.rule_descriptions[rule] for rule in triggered_rules}
        }



def main():
    analyzer = TriangleRuleAnalyzer()

    test_cases = [
        [3, 4, 5],
        [5, 5, 5],
        [5, 5, 8],
        [2, 3, 4],
        [2, 2, 3],
        [1, 1, 2]
    ]

    for sides in test_cases:
        result = analyzer.analyze_triangle(sides)
        print(f"边长: {sides}")
        print(f"Triangle type: {result['triangle_type']}")
        print(f"Angle type: {result['angle_type']}")
        print(f"The number of triggered rules: {len(result['triggered_rules'])}")
        print("Details of the triggered rules:")
        for rule_id, description in result['rule_details'].items():
            print(f"   {rule_id}: {description}")
        print("-" * 50)
if __name__ == "__main__":
    main()