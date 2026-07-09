import random
import datetime


class AdvancedTrafficControllerMassiveExtended:
    def __init__(self):
        # 初始化信号灯状态
        self.main_road_light = "red"  # 主路信号灯
        self.secondary_road_light = "green"  # 辅路信号灯
        self.pedestrian_light = "red"  # 行人信号灯
        self.turning_light = "red"  # 转向信号灯
        self.bicycle_light = "red"  # 自行车信号灯
        self.bus_light = "red"  # 公交车专用道信号灯

        # 初始化其他属性
        self.warning_light = "off"  # 警示灯
        self.traffic_police_alert = "off"  # 交警支援请求
        self.emergency_light = "off"  # 应急灯
        self.fog_light = "off"  # 雾灯
        self.speed_limit_display = "normal"  # 限速显示
        self.noise_alert = "off"  # 噪音警示
        self.air_quality_alert = "off"  # 空气质量警示

        # 信号灯时长参数(秒)
        self.max_main = 70
        self.min_main = 15
        self.max_secondary = 50
        self.min_secondary = 10
        self.base_pedestrian = 15
        self.base_turning = 10
        self.base_bicycle = 12
        self.base_bus = 20

        # 天气编码：1=sunny, 2=rainy, 3=foggy, 4=snowy, 5=windy, 6=stormy
        self.weather_codes = [1, 2, 3, 4, 5, 6]
        self.weather_names = ["sunny", "rainy", "foggy", "snowy", "windy", "stormy"]

        # 时间段编码：1=rush_morning, 2=rush_evening, 3=lunch, 4=night, 5=weekend, 6=holiday
        self.time_codes = [1, 2, 3, 4, 5, 6]
        self.time_names = ["rush_morning", "rush_evening", "lunch", "night", "weekend", "holiday"]

        self.special_events = ["none", "festival", "sports_game", "concert", "emergency", "construction"]

    def get_sensor_data(self):
        """获取传感器数据（模拟）"""
        # 生成模拟数据 (x:主路, y:辅路, z:行人)
        x = random.uniform(5, 95)  # 主路车流量(0-100)
        y = random.uniform(5, 95)  # 辅路车流量(0-100)
        z = random.uniform(0, 60)  # 行人数量(0-60)

        # 新增传感器数据 - 使用数字编码
        weather = random.choice(self.weather_codes)  # 1-6的数字
        time_period = random.choice(self.time_codes)  # 1-6的数字
        special_event = random.choice(self.special_events)

        # 车辆类型分布
        vehicle_types = {
            "cars": random.uniform(60, 85),
            "trucks": random.uniform(5, 20),
            "buses": random.uniform(3, 15),
            "motorcycles": random.uniform(2, 12),
            "bicycles": random.uniform(5, 25)
        }

        # 其他环境因素
        emergency_vehicles = random.randint(0, 3)
        air_quality = random.uniform(30, 150)  # AQI指数
        noise_level = random.uniform(40, 85)  # 分贝
        visibility = random.uniform(50, 1000)  # 米
        temperature = random.uniform(-10, 40)  # 摄氏度
        wind_speed = random.uniform(0, 25)  # 米/秒

        # 区域特征
        school_zones_active = random.choice([True, False])
        business_district_active = random.choice([True, False])
        hospital_nearby = random.choice([True, False])
        construction_zone = random.choice([True, False])

        # 交通事件
        accident_probability = random.uniform(0, 100)
        road_condition = random.choice(["excellent", "good", "fair", "poor", "very_poor"])

        return (x, y, z, weather, time_period, special_event, vehicle_types,
                emergency_vehicles, air_quality, noise_level, visibility,
                temperature, wind_speed, school_zones_active, business_district_active,
                hospital_nearby, construction_zone, accident_probability, road_condition)

    def block1_x_y_z_combinations(self, x, y, z, weather, time_period, special_event,
                                  vehicle_types, emergency_vehicles, air_quality,
                                  noise_level, visibility, temperature, wind_speed,
                                  school_zones_active, business_district_active,
                                  hospital_nearby, construction_zone, accident_probability, road_condition):
        """大块1: X, Y, Z流量组合处理（只包含这三个变量）- 大规模扩展版"""
        actions = []
        adjustments = {}

        # === 只包含X, Y, Z的组合 - 45个IF语句 ===

        # 1. 主路高流量、辅路低流量、行人少
        if x > 85 and y < 40 and z < 25:
            actions.append("主路严重拥堵，辅路畅通，行人少，最大化主路通行")
            adjustments["main"] = self.max_main + 15
            adjustments["secondary"] = self.min_secondary - 3

        # 2. 主路高流量、辅路低流量、行人多
        if x > 80 and y < 45 and z > 40:
            actions.append("主路拥堵，辅路畅通，行人多，平衡车辆与行人")
            adjustments["main"] = self.max_main - 5
            adjustments["secondary"] = self.min_secondary + 5
            adjustments["pedestrian"] = self.base_pedestrian + 15

        # 3. 主路极高流量、辅路极低流量、行人极少
        if x > 92 and y < 30 and z < 15:
            actions.append("主路极度拥堵，辅路极畅通，行人极少，极大化主路")
            adjustments["main"] = self.max_main + 25
            adjustments["secondary"] = self.min_secondary - 5
            self.traffic_police_alert = "on"

        # 4. 主路高流量、辅路中等、行人密集
        if x > 75 and 45 < y < 65 and z > 50:
            actions.append("主路拥堵，辅路中等，行人密集，三方平衡")
            adjustments["main"] = 45
            adjustments["secondary"] = 35
            adjustments["pedestrian"] = self.base_pedestrian + 20

        # 5. 主路低流量、辅路高流量、行人少
        if x < 50 and y > 80 and z < 25:
            actions.append("主路畅通，辅路拥堵，行人少，最大化辅路通行")
            adjustments["main"] = self.min_main - 2
            adjustments["secondary"] = self.max_secondary + 12

        # 6. 主路低流量、辅路高流量、行人多
        if x < 45 and y > 75 and z > 40:
            actions.append("主路畅通，辅路拥堵，行人多，平衡辅路与行人")
            adjustments["main"] = self.min_main + 5
            adjustments["secondary"] = self.max_secondary - 5
            adjustments["pedestrian"] = self.base_pedestrian + 12

        # 7. 主路极低流量、辅路极高流量、行人极少
        if x < 30 and y > 92 and z < 15:
            actions.append("主路极畅通，辅路极度拥堵，行人极少，极大化辅路")
            adjustments["main"] = self.min_main - 5
            adjustments["secondary"] = self.max_secondary + 20
            self.traffic_police_alert = "on"

        # 8. 主路中等、辅路高流量、行人密集
        if 45 < x < 65 and y > 75 and z > 50:
            actions.append("主路中等，辅路拥堵，行人密集，三方平衡")
            adjustments["main"] = 35
            adjustments["secondary"] = 45
            adjustments["pedestrian"] = self.base_pedestrian + 18

        # 9. 主辅路均高流量、行人少
        if x > 75 and y > 75 and z < 30:
            actions.append("主辅路均拥堵，行人少，均衡分配车辆时间")
            adjustments["main"] = 48
            adjustments["secondary"] = 42

        # 10. 主辅路均高流量、行人多
        if x > 70 and y > 70 and z > 45:
            actions.append("主辅路均拥堵，行人多，综合平衡分配")
            adjustments["main"] = 40
            adjustments["secondary"] = 35
            adjustments["pedestrian"] = self.base_pedestrian + 18

        # 11. 主辅路均极高流量、行人适中
        if x > 88 and y > 88 and 25 < z < 45:
            actions.append("主辅路均极度拥堵，行人适中，启动管制")
            adjustments["main"] = 35
            adjustments["secondary"] = 30
            adjustments["pedestrian"] = self.base_pedestrian + 10
            self.traffic_police_alert = "on"

        # 12. 主辅路均高流量、行人极多
        if x > 75 and y > 75 and z > 55:
            actions.append("主辅路均拥堵，行人极多，行人优先保护")
            adjustments["main"] = 35
            adjustments["secondary"] = 30
            adjustments["pedestrian"] = self.base_pedestrian + 25
            self.emergency_light = "on"

        # 13. 主辅路均低流量、行人少
        if x < 35 and y < 35 and z < 25:
            actions.append("主辅路均低流量，行人少，最小周期")
            adjustments["main"] = 25
            adjustments["secondary"] = 20
            adjustments["pedestrian"] = self.base_pedestrian - 3

        # 14. 主辅路均低流量、行人多
        if x < 40 and y < 40 and z > 40:
            actions.append("主辅路均低流量，行人多，优先行人通行")
            adjustments["main"] = 28
            adjustments["secondary"] = 22
            adjustments["pedestrian"] = self.base_pedestrian + 15

        # 15. 主辅路均极低流量、行人适中
        if x < 25 and y < 25 and 20 < z < 40:
            actions.append("主辅路均极低流量，行人适中，行人为主")
            adjustments["main"] = 20
            adjustments["secondary"] = 18
            adjustments["pedestrian"] = self.base_pedestrian + 8

        # 16. 主辅路均极低流量、行人极少
        if x < 20 and y < 20 and z < 15:
            actions.append("主辅路均极低流量，行人极少，极小周期")
            adjustments["main"] = 18
            adjustments["secondary"] = 15
            adjustments["pedestrian"] = self.base_pedestrian - 5

        # 17. 主路极高、辅路极高、行人极多
        if x > 90 and y > 90 and z > 55:
            actions.append("主辅路极度拥堵，行人极多，启动最高级别管制")
            adjustments["main"] = 30
            adjustments["secondary"] = 25
            adjustments["pedestrian"] = self.base_pedestrian + 25
            self.traffic_police_alert = "on"
            self.emergency_light = "on"

        # 18. 主路适中、辅路适中、行人适中
        if 45 < x < 70 and 45 < y < 70 and 25 < z < 45:
            actions.append("主辅路适中流量，行人适中，标准均衡配时")
            adjustments["main"] = 42
            adjustments["secondary"] = 36
            adjustments["pedestrian"] = self.base_pedestrian

        # 19. 主路高、辅路适中、行人多
        if x > 75 and 40 < y < 65 and z > 45:
            actions.append("主路拥堵，辅路适中，行人多，优先主路和行人")
            adjustments["main"] = 50
            adjustments["secondary"] = 30
            adjustments["pedestrian"] = self.base_pedestrian + 16

        # 20. 主路适中、辅路高、行人多
        if 40 < x < 65 and y > 75 and z > 45:
            actions.append("主路适中，辅路拥堵，行人多，优先辅路和行人")
            adjustments["main"] = 30
            adjustments["secondary"] = 48
            adjustments["pedestrian"] = self.base_pedestrian + 16

        # 21. 主路低、辅路适中、行人密集
        if x < 45 and 40 < y < 70 and z > 50:
            actions.append("主路低流量，辅路适中，行人密集，行人优先")
            adjustments["main"] = 25
            adjustments["secondary"] = 35
            adjustments["pedestrian"] = self.base_pedestrian + 20

        # 22. 主路适中、辅路低、行人密集
        if 40 < x < 70 and y < 45 and z > 50:
            actions.append("主路适中，辅路低流量，行人密集，行人优先")
            adjustments["main"] = 35
            adjustments["secondary"] = 25
            adjustments["pedestrian"] = self.base_pedestrian + 20

        # 23. 主路高、辅路低、行人适中
        if x > 78 and y < 42 and 20 < z < 40:
            actions.append("主路拥堵，辅路畅通，行人适中，主路为主")
            adjustments["main"] = self.max_main + 8
            adjustments["secondary"] = self.min_secondary + 3
            adjustments["pedestrian"] = self.base_pedestrian + 5

        # 24. 主路低、辅路高、行人适中
        if x < 42 and y > 78 and 20 < z < 40:
            actions.append("主路畅通，辅路拥堵，行人适中，辅路为主")
            adjustments["main"] = self.min_main + 3
            adjustments["secondary"] = self.max_secondary + 8
            adjustments["pedestrian"] = self.base_pedestrian + 5

        # 25. 主路中等偏高、辅路中等偏低、行人少
        if 60 < x < 80 and 30 < y < 50 and z < 30:
            actions.append("主路中等偏高，辅路中等偏低，行人少，轻微优先主路")
            adjustments["main"] = 48
            adjustments["secondary"] = 32

        # 26. 主路中等偏低、辅路中等偏高、行人少
        if 30 < x < 50 and 60 < y < 80 and z < 30:
            actions.append("主路中等偏低，辅路中等偏高，行人少，轻微优先辅路")
            adjustments["main"] = 32
            adjustments["secondary"] = 48

        # 27. 主路极高、辅路中等、行人中等
        if x > 90 and 40 < y < 65 and 25 < z < 45:
            actions.append("主路极高，辅路中等，行人中等，主路绝对优先")
            adjustments["main"] = self.max_main + 20
            adjustments["secondary"] = self.min_secondary
            adjustments["pedestrian"] = self.base_pedestrian + 8

        # 28. 主路中等、辅路极高、行人中等
        if 40 < x < 65 and y > 90 and 25 < z < 45:
            actions.append("主路中等，辅路极高，行人中等，辅路绝对优先")
            adjustments["main"] = self.min_main
            adjustments["secondary"] = self.max_secondary + 20
            adjustments["pedestrian"] = self.base_pedestrian + 8

        # 29. 主路高、辅路高、行人极少
        if x > 78 and y > 78 and z < 18:
            actions.append("主路高，辅路高，行人极少，车辆极大化通行")
            adjustments["main"] = 52
            adjustments["secondary"] = 45
            adjustments["pedestrian"] = self.base_pedestrian - 5

        # 30. 主路低、辅路低、行人极多
        if x < 35 and y < 35 and z > 55:
            actions.append("主路低，辅路低，行人极多，行人极大化通行")
            adjustments["main"] = 22
            adjustments["secondary"] = 18
            adjustments["pedestrian"] = self.base_pedestrian + 30

        # 31. 主路中高、辅路中低、行人中高
        if 65 < x < 85 and 25 < y < 45 and 35 < z < 50:
            actions.append("主路中高，辅路中低，行人中高，主路行人并重")
            adjustments["main"] = 46
            adjustments["secondary"] = 28
            adjustments["pedestrian"] = self.base_pedestrian + 12

        # 32. 主路中低、辅路中高、行人中高
        if 25 < x < 45 and 65 < y < 85 and 35 < z < 50:
            actions.append("主路中低，辅路中高，行人中高，辅路行人并重")
            adjustments["main"] = 28
            adjustments["secondary"] = 46
            adjustments["pedestrian"] = self.base_pedestrian + 12

        # 33. 主路极低、辅路中等、行人少
        if x < 25 and 45 < y < 70 and z < 25:
            actions.append("主路极低，辅路中等，行人少，辅路优化")
            adjustments["main"] = 18
            adjustments["secondary"] = 40
            adjustments["pedestrian"] = self.base_pedestrian - 2

        # 34. 主路中等、辅路极低、行人少
        if 45 < x < 70 and y < 25 and z < 25:
            actions.append("主路中等，辅路极低，行人少，主路优化")
            adjustments["main"] = 40
            adjustments["secondary"] = 18
            adjustments["pedestrian"] = self.base_pedestrian - 2

        # 35. 主路极高、辅路低、行人少
        if x > 90 and y < 40 and z < 20:
            actions.append("主路极高，辅路低，行人少，主路最大化")
            adjustments["main"] = self.max_main + 25
            adjustments["secondary"] = self.min_secondary - 5
            adjustments["pedestrian"] = self.base_pedestrian - 3

        # 36. 主路低、辅路极高、行人少
        if x < 40 and y > 90 and z < 20:
            actions.append("主路低，辅路极高，行人少，辅路最大化")
            adjustments["main"] = self.min_main - 5
            adjustments["secondary"] = self.max_secondary + 25
            adjustments["pedestrian"] = self.base_pedestrian - 3

        # 37. 主路中等偏高、辅路中等偏高、行人少
        if 60 < x < 78 and 60 < y < 78 and z < 25:
            actions.append("主路中等偏高，辅路中等偏高，行人少，均衡高效")
            adjustments["main"] = 44
            adjustments["secondary"] = 40

        # 38. 主路中等偏低、辅路中等偏低、行人多
        if 30 < x < 48 and 30 < y < 48 and z > 45:
            actions.append("主路中等偏低，辅路中等偏低，行人多，行人优先")
            adjustments["main"] = 30
            adjustments["secondary"] = 26
            adjustments["pedestrian"] = self.base_pedestrian + 18

        # 39. 主路高、辅路中低、行人极少
        if x > 75 and 25 < y < 45 and z < 15:
            actions.append("主路高，辅路中低，行人极少，主路强化")
            adjustments["main"] = self.max_main + 12
            adjustments["secondary"] = self.min_secondary + 5
            adjustments["pedestrian"] = self.base_pedestrian - 5

        # 40. 主路中低、辅路高、行人极少
        if 25 < x < 45 and y > 75 and z < 15:
            actions.append("主路中低，辅路高，行人极少，辅路强化")
            adjustments["main"] = self.min_main + 5
            adjustments["secondary"] = self.max_secondary + 12
            adjustments["pedestrian"] = self.base_pedestrian - 5

        # 41. 主路极低、辅路极低、行人中等
        if x < 22 and y < 22 and 25 < z < 45:
            actions.append("主路极低，辅路极低，行人中等，行人主导")
            adjustments["main"] = 20
            adjustments["secondary"] = 16
            adjustments["pedestrian"] = self.base_pedestrian + 10

        # 42. 主路极高、辅路极低、行人中等
        if x > 92 and y < 22 and 25 < z < 45:
            actions.append("主路极高，辅路极低，行人中等，主路行人平衡")
            adjustments["main"] = self.max_main + 15
            adjustments["secondary"] = self.min_secondary - 5
            adjustments["pedestrian"] = self.base_pedestrian + 10

        # 43. 主路极低、辅路极高、行人中等
        if x < 22 and y > 92 and 25 < z < 45:
            actions.append("主路极低，辅路极高，行人中等，辅路行人平衡")
            adjustments["main"] = self.min_main - 5
            adjustments["secondary"] = self.max_secondary + 15
            adjustments["pedestrian"] = self.base_pedestrian + 10

        # 44. 主路中等、辅路中等、行人极多
        if 45 < x < 70 and 45 < y < 70 and z > 55:
            actions.append("主路中等，辅路中等，行人极多，行人绝对优先")
            adjustments["main"] = 30
            adjustments["secondary"] = 25
            adjustments["pedestrian"] = self.base_pedestrian + 28

        # 45. 主路中等、辅路中等、行人极少
        if 45 < x < 70 and 45 < y < 70 and z < 12:
            actions.append("主路中等，辅路中等，行人极少，车辆优化")
            adjustments["main"] = 45
            adjustments["secondary"] = 40
            adjustments["pedestrian"] = self.base_pedestrian - 8

            # === 两变量比较 - X和Y对比 ===
            # 46. 主路显著高于辅路
            if x > y + 30:
                actions.append("主路流量显著高于辅路，主路绝对优先")
                adjustments["main"] = self.max_main + 8
                adjustments["secondary"] = self.min_secondary + 2

            # 47. 辅路显著高于主路
            if y > x + 30:
                actions.append("辅路流量显著高于主路，辅路绝对优先")
                adjustments["main"] = self.min_main + 2
                adjustments["secondary"] = self.max_secondary + 8

            # 48. 主辅路流量相近
            if abs(x - y) < 10:
                actions.append("主辅路流量相近，均衡标准配时")
                adjustments["main"] = 40
                adjustments["secondary"] = 36

            # 49. 主路极度超越辅路
            if x > y + 50:
                actions.append("主路极度超越辅路，主路最大化")
                adjustments["main"] = self.max_main + 20
                adjustments["secondary"] = self.min_secondary - 3
                self.traffic_police_alert = "on"

            # 50. 辅路极度超越主路
            if y > x + 50:
                actions.append("辅路极度超越主路，辅路最大化")
                adjustments["main"] = self.min_main - 3
                adjustments["secondary"] = self.max_secondary + 20
                self.traffic_police_alert = "on"

            # === 两变量比较 - X和Z对比 ===
            # 51. 主路流量远超行人密度
            if x > z + 40:
                actions.append("主路流量远超行人密度，车辆优先")
                adjustments["main"] = self.max_main + 5
                adjustments["pedestrian"] = self.base_pedestrian - 2

            # 52. 行人密度超过主路流量
            if z > x - 20:
                actions.append("行人密度超过主路流量，行人优先")
                adjustments["main"] = self.max_main - 8
                adjustments["pedestrian"] = self.base_pedestrian + 12

            # 53. 主路与行人需求平衡
            if abs(x - z) < 15:
                actions.append("主路与行人需求平衡，协调配时")
                adjustments["main"] = 42
                adjustments["pedestrian"] = self.base_pedestrian + 3

            # 54. 主路流量极度超过行人
            if x > z + 60:
                actions.append("主路流量极度超过行人，车流绝对优先")
                adjustments["main"] = self.max_main + 15
                adjustments["pedestrian"] = self.base_pedestrian - 5

            # 55. 行人密度极度超过主路
            if z > x + 10:
                actions.append("行人密度极度超过主路，行人绝对优先")
                adjustments["main"] = self.max_main - 15
                adjustments["pedestrian"] = self.base_pedestrian + 20

            # === 两变量比较 - Y和Z对比 ===
            # 56. 辅路流量远超行人密度
            if y > z + 40:
                actions.append("辅路流量远超行人密度，辅路车辆优先")
                adjustments["secondary"] = self.max_secondary + 5
                adjustments["pedestrian"] = self.base_pedestrian - 2

            # 57. 行人密度超过辅路流量
            if z > y - 20:
                actions.append("行人密度超过辅路流量，行人优先于辅路")
                adjustments["secondary"] = self.max_secondary - 8
                adjustments["pedestrian"] = self.base_pedestrian + 12

            # 58. 辅路与行人需求平衡
            if abs(y - z) < 15:
                actions.append("辅路与行人需求平衡，协调配时")
                adjustments["secondary"] = 38
                adjustments["pedestrian"] = self.base_pedestrian + 3

            # 59. 辅路流量极度超过行人
            if y > z + 60:
                actions.append("辅路流量极度超过行人，辅路绝对优先")
                adjustments["secondary"] = self.max_secondary + 15
                adjustments["pedestrian"] = self.base_pedestrian - 5

            # 60. 行人密度极度超过辅路
            if z > y + 10:
                actions.append("行人密度极度超过辅路，行人绝对优先于辅路")
                adjustments["secondary"] = self.max_secondary - 15
                adjustments["pedestrian"] = self.base_pedestrian + 20

            return actions, adjustments

    def block2_weather_x_y_combinations(self, x, y, z, weather, time_period, special_event,
                                        vehicle_types, emergency_vehicles, air_quality,
                                        noise_level, visibility, temperature, wind_speed,
                                        school_zones_active, business_district_active,
                                        hospital_nearby, construction_zone, accident_probability, road_condition):
        """大块2: weather(数字), X, Y组合处理（只包含这三个变量）- 大规模扩展版"""
        actions = []
        adjustments = {}

        # === 只包含weather(数字), X, Y的组合 - 40个IF语句 ===
        # weather: 1=sunny, 2=rainy, 3=foggy, 4=snowy, 5=windy, 6=stormy

        # 1-10: 晴天组合
        # 1. 晴天主路拥堵辅路畅通
        if weather == 1 and x > 75 and y < 50:
            actions.append("晴天主路拥堵辅路畅通，标准优化配时")
            adjustments["main"] = self.max_main + 8
            adjustments["secondary"] = self.min_secondary

        # 2. 晴天主路畅通辅路拥堵
        if weather == 1 and x < 50 and y > 75:
            actions.append("晴天主路畅通辅路拥堵，优先辅路通行")
            adjustments["main"] = self.min_main
            adjustments["secondary"] = self.max_secondary + 8

        # 3. 晴天主辅路均拥堵
        if weather == 1 and x > 70 and y > 70:
            actions.append("晴天主辅路均拥堵，高效均衡分配")
            adjustments["main"] = self.max_main + 5
            adjustments["secondary"] = self.max_secondary + 5

        # 4. 晴天主辅路均低流量
        if weather == 1 and x < 40 and y < 40:
            actions.append("晴天主辅路均低流量，鼓励绿色出行")
            adjustments["main"] = 30
            adjustments["secondary"] = 25
            adjustments["bicycle"] = self.base_bicycle + 8

        # 5. 晴天主路极高辅路极低
        if weather == 1 and x > 90 and y < 30:
            actions.append("晴天主路极高辅路极低，最大化主路")
            adjustments["main"] = self.max_main + 20
            adjustments["secondary"] = self.min_secondary - 5

        # 6. 晴天主路极低辅路极高
        if weather == 1 and x < 30 and y > 90:
            actions.append("晴天主路极低辅路极高，最大化辅路")
            adjustments["main"] = self.min_main - 5
            adjustments["secondary"] = self.max_secondary + 20

        # 7. 晴天主路中等辅路高
        if weather == 1 and 45 < x < 70 and y > 78:
            actions.append("晴天主路中等辅路高，优先辅路")
            adjustments["main"] = 32
            adjustments["secondary"] = self.max_secondary + 6

        # 8. 晴天主路高辅路中等
        if weather == 1 and x > 78 and 45 < y < 70:
            actions.append("晴天主路高辅路中等，优先主路")
            adjustments["main"] = self.max_main + 6
            adjustments["secondary"] = 32

        # 9. 晴天主辅路均极低
        if weather == 1 and x < 25 and y < 25:
            actions.append("晴天主辅路均极低，极小优化配时")
            adjustments["main"] = 22
            adjustments["secondary"] = 18
            adjustments["bicycle"] = self.base_bicycle + 10

        # 10. 晴天主辅路均极高
        if weather == 1 and x > 88 and y > 88:
            actions.append("晴天主辅路均极高，高效管制")
            adjustments["main"] = self.max_main + 10
            adjustments["secondary"] = self.max_secondary + 8

        # 11-20: 雨天组合
        # 11. 雨天主路拥堵辅路畅通
        if weather == 2 and x > 75 and y < 50:
            actions.append("雨天主路拥堵辅路畅通，降低主路速度延长时间")
            adjustments["main"] = self.max_main - 8
            adjustments["secondary"] = self.min_secondary + 5
            self.speed_limit_display = "reduced"

        # 12. 雨天主路畅通辅路拥堵
        if weather == 2 and x < 50 and y > 75:
            actions.append("雨天主路畅通辅路拥堵，优先辅路谨慎通行")
            adjustments["main"] = self.min_main + 5
            adjustments["secondary"] = self.max_secondary - 8
            self.speed_limit_display = "reduced"

        # 13. 雨天主辅路均拥堵
        if weather == 2 and x > 70 and y > 70:
            actions.append("雨天主辅路均拥堵，降低通行强度")
            adjustments["main"] = self.max_main - 15
            adjustments["secondary"] = self.max_secondary - 12
            self.warning_light = "on"

        # 14. 雨天主辅路均低流量
        if weather == 2 and x < 35 and y < 35:
            actions.append("雨天主辅路均低流量，谨慎最小配时")
            adjustments["main"] = 28
            adjustments["secondary"] = 22
            self.speed_limit_display = "slightly_reduced"

        # 15. 雨天主路极高辅路中等
        if weather == 2 and x > 85 and 45 < y < 70:
            actions.append("雨天主路极高辅路中等，雨天管制模式")
            adjustments["main"] = self.max_main - 12
            adjustments["secondary"] = self.min_secondary + 8
            self.warning_light = "on"

        # 16. 雨天主路中等辅路极高
        if weather == 2 and 45 < x < 70 and y > 85:
            actions.append("雨天主路中等辅路极高，雨天辅路管制")
            adjustments["main"] = self.min_main + 8
            adjustments["secondary"] = self.max_secondary - 12
            self.warning_light = "on"

        # 17. 雨天主路低辅路中等
        if weather == 2 and x < 45 and 50 < y < 75:
            actions.append("雨天主路低辅路中等，轻微调整辅路")
            adjustments["main"] = self.min_main + 8
            adjustments["secondary"] = self.max_secondary - 5
            self.speed_limit_display = "slightly_reduced"

        # 18. 雨天主路中等辅路低
        if weather == 2 and 50 < x < 75 and y < 45:
            actions.append("雨天主路中等辅路低，轻微调整主路")
            adjustments["main"] = self.max_main - 5
            adjustments["secondary"] = self.min_secondary + 8
            self.speed_limit_display = "slightly_reduced"

        # 19. 雨天主辅路均极低
        if weather == 2 and x < 25 and y < 25:
            actions.append("雨天主辅路均极低，雨天极小配时")
            adjustments["main"] = 25
            adjustments["secondary"] = 20
            self.speed_limit_display = "reduced"

        # 20. 雨天主辅路均极高
        if weather == 2 and x > 88 and y > 88:
            actions.append("雨天主辅路均极高，雨天紧急管制")
            adjustments["main"] = self.max_main - 20
            adjustments["secondary"] = self.max_secondary - 18
            self.emergency_light = "on"

        # 21-28: 雾天组合
        # 21. 雾天主路拥堵辅路畅通
        if weather == 3 and x > 65 and y < 50:
            actions.append("雾天主路拥堵辅路畅通，启动雾灯降低主路速度")
            adjustments["main"] = min(self.max_main - 25, 35)
            adjustments["secondary"] = self.min_secondary + 10
            self.fog_light = "on"
            self.speed_limit_display = "very_reduced"

        # 22. 雾天主路畅通辅路拥堵
        if weather == 3 and x < 50 and y > 65:
            actions.append("雾天主路畅通辅路拥堵，启动雾灯降低辅路速度")
            adjustments["main"] = self.min_main + 10
            adjustments["secondary"] = min(self.max_secondary - 25, 30)
            self.fog_light = "on"
            self.speed_limit_display = "very_reduced"

        # 23. 雾天主辅路均拥堵
        if weather == 3 and x > 60 and y > 60:
            actions.append("雾天主辅路均拥堵，雾天紧急模式")
            adjustments["main"] = 30
            adjustments["secondary"] = 25
            self.fog_light = "on"
            self.emergency_light = "on"

        # 24. 雾天主辅路均低流量
        if weather == 3 and x < 40 and y < 40:
            actions.append("雾天主辅路均低流量，雾天谨慎模式")
            adjustments["main"] = 25
            adjustments["secondary"] = 20
            self.fog_light = "on"
            self.speed_limit_display = "reduced"

        # 25. 雾天主路适中辅路高
        if weather == 3 and 40 < x < 65 and y > 75:
            actions.append("雾天主路适中辅路高，雾天辅路优先")
            adjustments["main"] = self.min_main + 8
            adjustments["secondary"] = min(self.max_secondary - 15, 35)
            self.fog_light = "on"

        # 26. 雾天主路高辅路适中
        if weather == 3 and x > 75 and 40 < y < 65:
            actions.append("雾天主路高辅路适中，雾天主路优先")
            adjustments["main"] = min(self.max_main - 15, 35)
            adjustments["secondary"] = self.min_secondary + 8
            self.fog_light = "on"

        # 27. 雾天主辅路均极低
        if weather == 3 and x < 25 and y < 25:
            actions.append("雾天主辅路均极低，雾天极谨慎")
            adjustments["main"] = 20
            adjustments["secondary"] = 18
            self.fog_light = "on"
            self.speed_limit_display = "very_reduced"

        # 28. 雾天主辅路均极高
        if weather == 3 and x > 85 and y > 85:
            actions.append("雾天主辅路均极高，雾天最高警戒")
            adjustments["main"] = 25
            adjustments["secondary"] = 20
            self.fog_light = "on"
            self.emergency_light = "on"

        # 29-34: 雪天组合
        # 29. 雪天主路拥堵辅路畅通
        if weather == 4 and x > 70 and y < 50:
            actions.append("雪天主路拥堵辅路畅通，启动雪天主路模式")
            adjustments["main"] = max(self.max_main - 20, 30)
            adjustments["secondary"] = self.min_secondary + 8
            self.speed_limit_display = "very_reduced"

        # 30. 雪天主路畅通辅路拥堵
        if weather == 4 and x < 50 and y > 70:
            actions.append("雪天主路畅通辅路拥堵，启动雪天辅路模式")
            adjustments["main"] = self.min_main + 8
            adjustments["secondary"] = max(self.max_secondary - 20, 25)
            self.speed_limit_display = "very_reduced"

        # 31. 雪天主辅路均拥堵
        if weather == 4 and x > 65 and y > 65:
            actions.append("雪天主辅路均拥堵，雪天紧急模式")
            adjustments["main"] = 25
            adjustments["secondary"] = 20
            self.emergency_light = "on"
            self.speed_limit_display = "very_reduced"

        # 32. 雪天主辅路均低流量
        if weather == 4 and x < 35 and y < 35:
            actions.append("雪天主辅路均低流量，雪天最小谨慎配时")
            adjustments["main"] = 25
            adjustments["secondary"] = 20
            self.speed_limit_display = "reduced"

        # 33. 雪天主路高辅路适中
        if weather == 4 and x > 75 and 35 < y < 65:
            actions.append("雪天主路高辅路适中，雪天主路优先")
            adjustments["main"] = max(self.max_main - 25, 35)
            adjustments["secondary"] = self.min_secondary + 10
            self.warning_light = "on"

        # 34. 雪天主路适中辅路高
        if weather == 4 and 35 < x < 65 and y > 75:
            actions.append("雪天主路适中辅路高，雪天辅路优先")
            adjustments["main"] = self.min_main + 10
            adjustments["secondary"] = max(self.max_secondary - 25, 35)
            self.warning_light = "on"

        # 35-38: 强风组合
        # 35. 强风主路拥堵辅路畅通
        if weather == 5 and x > 75 and y < 45:
            actions.append("强风主路拥堵辅路畅通，注意大型车辆主路通行")
            adjustments["main"] = self.max_main - 10
            adjustments["secondary"] = self.min_secondary + 3
            self.warning_light = "on"

        # 36. 强风主路畅通辅路拥堵
        if weather == 5 and x < 45 and y > 75:
            actions.append("强风主路畅通辅路拥堵，注意大型车辆辅路通行")
            adjustments["main"] = self.min_main + 3
            adjustments["secondary"] = self.max_secondary - 10
            self.warning_light = "on"

        # 37. 强风主辅路均拥堵
        if weather == 5 and x > 70 and y > 70:
            actions.append("强风主辅路均拥堵，强风管制模式")
            adjustments["main"] = self.max_main - 15
            adjustments["secondary"] = self.max_secondary - 12
            self.warning_light = "on"

        # 38. 强风主辅路均低流量
        if weather == 5 and x < 40 and y < 40:
            actions.append("强风主辅路均低流量，强风谨慎模式")
            adjustments["main"] = 32
            adjustments["secondary"] = 28
            self.speed_limit_display = "slightly_reduced"

        # 39-40: 暴风雨组合
        # 39. 暴风雨主辅路拥堵
        if weather == 6 and x > 60 and y > 60:
            actions.append("暴风雨主辅路拥堵，启动最高级紧急模式")
            adjustments["main"] = 25
            adjustments["secondary"] = 20
            self.emergency_light = "on"
            self.speed_limit_display = "very_reduced"

        # 40. 暴风雨主辅路低流量
        if weather == 6 and x < 40 and y < 40:
            actions.append("暴风雨主辅路低流量，暴雨最小配时")
            adjustments["main"] = 20
            adjustments["secondary"] = 18
            self.emergency_light = "on"
            self.speed_limit_display = "very_reduced"

        # 41-50: weather vs x 天气与主路流量比较
        # 41. 晴天主路极高流量
        if weather == 1 and x > 90:
            actions.append("晴天主路极高流量，最大化通行效率")
            adjustments["main"] = adjustments.get("main", self.max_main) + 12
            adjustments["secondary"] = max(adjustments.get("secondary", self.max_secondary) - 5, self.min_secondary)

        # 42. 晴天主路极低流量
        if weather == 1 and x < 20:
            actions.append("晴天主路极低流量，节能环保配时")
            adjustments["main"] = 18
            adjustments["bicycle"] = self.base_bicycle + 12

        # 43. 恶劣天气主路高流量
        if weather in [2, 3, 4, 6] and x > 80:
            actions.append("恶劣天气主路高流量，强制减速延时")
            adjustments["main"] = max(adjustments.get("main", self.max_main) - 18, 25)
            self.speed_limit_display = "very_reduced"
            self.warning_light = "on"

        # 44. 恶劣天气主路低流量
        if weather in [2, 3, 4, 6] and x < 30:
            actions.append("恶劣天气主路低流量，安全最小配时")
            adjustments["main"] = 20
            self.speed_limit_display = "reduced"

        # 45. 雾雪天主路中高流量
        if weather in [3, 4] and 60 < x < 85:
            actions.append("雾雪天主路中高流量，能见度优先控制")
            adjustments["main"] = max(adjustments.get("main", self.max_main) - 15, 30)
            self.fog_light = "on" if weather == 3 else None
            self.speed_limit_display = "very_reduced"

        # 46. 强风天主路高流量
        if weather == 5 and x > 75:
            actions.append("强风天主路高流量，大型车辆管制")
            adjustments["main"] = max(adjustments.get("main", self.max_main) - 12, 30)
            self.warning_light = "on"

        # 47. 雨天主路中等流量
        if weather == 2 and 45 < x < 70:
            actions.append("雨天主路中等流量，防滑谨慎配时")
            adjustments["main"] = max(adjustments.get("main", self.max_main) - 8, 30)
            self.speed_limit_display = "slightly_reduced"

        # 48. 暴风雨主路任意流量
        if weather == 6 and x > 30:
            actions.append("暴风雨主路通行，极端天气应急")
            adjustments["main"] = 22
            self.emergency_light = "on"
            self.speed_limit_display = "very_reduced"

        # 49. 晴天主路中高流量
        if weather == 1 and 65 < x < 90:
            actions.append("晴天主路中高流量，标准高效配时")
            adjustments["main"] = adjustments.get("main", self.max_main) + 5

        # 50. 雪天主路极低流量
        if weather == 4 and x < 25:
            actions.append("雪天主路极低流量，防冰最小配时")
            adjustments["main"] = 18
            self.speed_limit_display = "reduced"

        # 51-60: weather vs y 天气与辅路流量比较
        # 51. 晴天辅路极高流量
        if weather == 1 and y > 90:
            actions.append("晴天辅路极高流量，辅路最大化通行")
            adjustments["secondary"] = adjustments.get("secondary", self.max_secondary) + 12
            adjustments["main"] = max(adjustments.get("main", self.max_main) - 5, self.min_main)

        # 52. 晴天辅路极低流量
        if weather == 1 and y < 20:
            actions.append("晴天辅路极低流量，辅路最小配时")
            adjustments["secondary"] = 15
            adjustments["main"] = adjustments.get("main", self.max_main) + 5

        # 53. 恶劣天气辅路高流量
        if weather in [2, 3, 4, 6] and y > 80:
            actions.append("恶劣天气辅路高流量，辅路谨慎延时")
            adjustments["secondary"] = max(adjustments.get("secondary", self.max_secondary) - 18, 22)
            self.speed_limit_display = "very_reduced"
            self.warning_light = "on"

        # 54. 恶劣天气辅路低流量
        if weather in [2, 3, 4, 6] and y < 30:
            actions.append("恶劣天气辅路低流量，辅路安全最小")
            adjustments["secondary"] = 18
            self.speed_limit_display = "reduced"

        # 55. 雾雪天辅路中高流量
        if weather in [3, 4] and 60 < y < 85:
            actions.append("雾雪天辅路中高流量，辅路能见度控制")
            adjustments["secondary"] = max(adjustments.get("secondary", self.max_secondary) - 15, 28)
            self.fog_light = "on" if weather == 3 else None
            self.speed_limit_display = "very_reduced"

        # 56. 强风天辅路高流量
        if weather == 5 and y > 75:
            actions.append("强风天辅路高流量，辅路风险管制")
            adjustments["secondary"] = max(adjustments.get("secondary", self.max_secondary) - 12, 28)
            self.warning_light = "on"

        # 57. 雨天辅路中等流量
        if weather == 2 and 45 < y < 70:
            actions.append("雨天辅路中等流量，辅路防滑配时")
            adjustments["secondary"] = max(adjustments.get("secondary", self.max_secondary) - 8, 28)
            self.speed_limit_display = "slightly_reduced"

        # 58. 暴风雨辅路任意流量
        if weather == 6 and y > 30:
            actions.append("暴风雨辅路通行，辅路极端应急")
            adjustments["secondary"] = 20
            self.emergency_light = "on"
            self.speed_limit_display = "very_reduced"

        # 59. 晴天辅路中高流量
        if weather == 1 and 65 < y < 90:
            actions.append("晴天辅路中高流量，辅路标准高效")
            adjustments["secondary"] = adjustments.get("secondary", self.max_secondary) + 5

        # 60. 雪天辅路极低流量
        if weather == 4 and y < 25:
            actions.append("雪天辅路极低流量，辅路防冰最小")
            adjustments["secondary"] = 16
            self.speed_limit_display = "reduced"

        # 61-70: x vs y 主路流量与辅路流量比较
        # 61. 主路极高辅路极低
        if x > 92 and y < 15:
            actions.append("主路极高辅路极低，主路绝对优先")
            adjustments["main"] = adjustments.get("main", self.max_main) + 15
            adjustments["secondary"] = max(adjustments.get("secondary", self.min_secondary) - 8, 12)

        # 62. 主路极低辅路极高
        if x < 15 and y > 92:
            actions.append("主路极低辅路极高，辅路绝对优先")
            adjustments["secondary"] = adjustments.get("secondary", self.max_secondary) + 15
            adjustments["main"] = max(adjustments.get("main", self.min_main) - 8, 12)

        # 63. 主辅路流量差距巨大(主路优势)
        if x - y > 60 and x > 70:
            actions.append("主辅路流量差距巨大主路优势，强化主路")
            adjustments["main"] = adjustments.get("main", self.max_main) + 10
            adjustments["secondary"] = max(adjustments.get("secondary", self.max_secondary) - 10, self.min_secondary)

        # 64. 主辅路流量差距巨大(辅路优势)
        if y - x > 60 and y > 70:
            actions.append("主辅路流量差距巨大辅路优势，强化辅路")
            adjustments["secondary"] = adjustments.get("secondary", self.max_secondary) + 10
            adjustments["main"] = max(adjustments.get("main", self.max_main) - 10, self.min_main)

        # 65. 主辅路均极高且接近
        if x > 90 and y > 90 and abs(x - y) < 10:
            actions.append("主辅路均极高且接近，超高强度均衡")
            adjustments["main"] = adjustments.get("main", self.max_main) + 8
            adjustments["secondary"] = adjustments.get("secondary", self.max_secondary) + 8

        # 66. 主辅路均极低且接近
        if x < 15 and y < 15 and abs(x - y) < 8:
            actions.append("主辅路均极低且接近，超低流量节能")
            adjustments["main"] = 15
            adjustments["secondary"] = 12
            adjustments["bicycle"] = self.base_bicycle + 15

        # 67. 主路中等辅路极高
        if 40 < x < 70 and y > 88:
            actions.append("主路中等辅路极高，辅路强化优先")
            adjustments["secondary"] = adjustments.get("secondary", self.max_secondary) + 8
            adjustments["main"] = max(adjustments.get("main", self.max_main) - 8, 25)

        # 68. 主路极高辅路中等
        if x > 88 and 40 < y < 70:
            actions.append("主路极高辅路中等，主路强化优先")
            adjustments["main"] = adjustments.get("main", self.max_main) + 8
            adjustments["secondary"] = max(adjustments.get("secondary", self.max_secondary) - 8, 25)

        # 69. 主辅路流量均衡中高
        if 70 < x < 85 and 70 < y < 85 and abs(x - y) < 10:
            actions.append("主辅路流量均衡中高，精细均衡配时")
            adjustments["main"] = adjustments.get("main", self.max_main) + 3
            adjustments["secondary"] = adjustments.get("secondary", self.max_secondary) + 3

        # 70. 主辅路流量均衡中低
        if 25 < x < 45 and 25 < y < 45 and abs(x - y) < 12:
            actions.append("主辅路流量均衡中低，标准均衡配时")
            adjustments["main"] = 30
            adjustments["secondary"] = 28

        return actions, adjustments

    def block3_weather_time_z_combinations(self, x, y, z, weather, time_period, special_event,
                                           vehicle_types, emergency_vehicles, air_quality,
                                           noise_level, visibility, temperature, wind_speed,
                                           school_zones_active, business_district_active,
                                           hospital_nearby, construction_zone, accident_probability, road_condition):
        """大块3: weather(数字), time_period(数字), Z组合处理（只包含这三个变量）- 大规模扩展版 + 补充缺失条件"""
        actions = []
        adjustments = {}

        # === 原有的71个IF语句 ===
        # weather: 1=sunny, 2=rainy, 3=foggy, 4=snowy, 5=windy, 6=stormy
        # time_period: 1=rush_morning, 2=rush_evening, 3=lunch, 4=night, 5=weekend, 6=holiday

        # 1-6: 晴天各时段组合
        # 1. 晴天早高峰行人多
        if weather == 1 and time_period == 1 and z > 40:
            actions.append("晴天早高峰行人多，舒适通行配时")
            adjustments["pedestrian"] = self.base_pedestrian + 12

        # 2. 晴天晚高峰行人多
        if weather == 1 and time_period == 2 and z > 40:
            actions.append("晴天晚高峰行人多，标准行人配时")
            adjustments["pedestrian"] = self.base_pedestrian + 10

        # 3. 晴天午餐时间行人密集
        if weather == 1 and time_period == 3 and z > 45:
            actions.append("晴天午餐时间行人密集，舒适通行")
            adjustments["pedestrian"] = self.base_pedestrian + 12

        # 4. 晴天夜间行人少
        if weather == 1 and time_period == 4 and z < 25:
            actions.append("晴天夜间行人少，标准夜间配时")
            adjustments["pedestrian"] = self.base_pedestrian - 2

        # 5. 晴天周末行人多
        if weather == 1 and time_period == 5 and z > 35:
            actions.append("晴天周末行人多，休闲模式")
            adjustments["pedestrian"] = self.base_pedestrian + 10

        # 6. 晴天假日行人密集
        if weather == 1 and time_period == 6 and z > 45:
            actions.append("晴天假日行人密集，节日舒适模式")
            adjustments["pedestrian"] = self.base_pedestrian + 15

        # 7-12: 雨天各时段组合
        # 7. 雨天早高峰行人多
        if weather == 2 and time_period == 1 and z > 40:
            actions.append("雨天早高峰行人多，增强行人保护")
            adjustments["pedestrian"] = self.base_pedestrian + 20
            self.speed_limit_display = "reduced"
            self.warning_light = "on"

        # 8. 雨天晚高峰行人多
        if weather == 2 and time_period == 2 and z > 40:
            actions.append("雨天晚高峰行人多，延长行人通行时间")
            adjustments["pedestrian"] = self.base_pedestrian + 18
            self.speed_limit_display = "reduced"

        # 9. 雨天午餐时间行人多
        if weather == 2 and time_period == 3 and z > 35:
            actions.append("雨天午餐时间行人多，雨天行人保护")
            adjustments["pedestrian"] = self.base_pedestrian + 16
            self.speed_limit_display = "slightly_reduced"

        # 10. 雨天夜间行人适中
        if weather == 2 and time_period == 4 and 20 < z < 40:
            actions.append("雨天夜间行人适中，雨夜安全模式")
            adjustments["pedestrian"] = self.base_pedestrian + 8
            self.warning_light = "on"

        # 11. 雨天周末行人适中
        if weather == 2 and time_period == 5 and 25 < z < 45:
            actions.append("雨天周末行人适中，休闲雨天模式")
            adjustments["pedestrian"] = self.base_pedestrian + 10
            self.speed_limit_display = "slightly_reduced"

        # 12. 雨天假日行人少
        if weather == 2 and time_period == 6 and z < 30:
            actions.append("雨天假日行人少，节日雨天模式")
            adjustments["pedestrian"] = self.base_pedestrian + 8
            self.speed_limit_display = "reduced"

        # 13-18: 雾天各时段组合
        # 13. 雾天早高峰行人多
        if weather == 3 and time_period == 1 and z > 35:
            actions.append("雾天早高峰行人多，极度警戒模式")
            adjustments["pedestrian"] = self.base_pedestrian + 30
            self.fog_light = "on"
            self.emergency_light = "on"

        # 14. 雾天晚高峰行人多
        if weather == 3 and time_period == 2 and z > 35:
            actions.append("雾天晚高峰行人多，雾天高度警戒")
            adjustments["pedestrian"] = self.base_pedestrian + 25
            self.fog_light = "on"
            self.warning_light = "on"

        # 15. 雾天午餐时间行人多
        if weather == 3 and time_period == 3 and z > 30:
            actions.append("雾天午餐时间行人多，雾天行人保护")
            adjustments["pedestrian"] = self.base_pedestrian + 22
            self.fog_light = "on"
            self.warning_light = "on"

        # 16. 雾天夜间行人适中
        if weather == 3 and time_period == 4 and 20 < z < 40:
            actions.append("雾天夜间行人适中，雾夜最高警戒")
            adjustments["pedestrian"] = self.base_pedestrian + 18
            self.fog_light = "on"
            self.emergency_light = "on"

        # 17. 雾天周末行人少
        if weather == 3 and time_period == 5 and z < 30:
            actions.append("雾天周末行人少，雾天谨慎模式")
            adjustments["pedestrian"] = self.base_pedestrian + 12
            self.fog_light = "on"

        # 18. 雾天假日行人适中
        if weather == 3 and time_period == 6 and 20 < z < 40:
            actions.append("雾天假日行人适中，节日雾天保护")
            adjustments["pedestrian"] = self.base_pedestrian + 15
            self.fog_light = "on"
            self.warning_light = "on"

        # 19-24: 雪天各时段组合
        # 19. 雪天早高峰行人多
        if weather == 4 and time_period == 1 and z > 35:
            actions.append("雪天早高峰行人多，最大保护措施")
            adjustments["pedestrian"] = self.base_pedestrian + 25
            self.speed_limit_display = "very_reduced"
            self.warning_light = "on"

        # 20. 雪天晚高峰行人多
        if weather == 4 and time_period == 2 and z > 35:
            actions.append("雪天晚高峰行人多，雪天强化保护")
            adjustments["pedestrian"] = self.base_pedestrian + 22
            self.speed_limit_display = "very_reduced"
            self.warning_light = "on"

        # 21. 雪天午餐时间行人适中
        if weather == 4 and time_period == 3 and 25 < z < 45:
            actions.append("雪天午餐时间行人适中，雪天行人保护")
            adjustments["pedestrian"] = self.base_pedestrian + 15
            self.speed_limit_display = "reduced"

        # 22. 雪天夜间行人少
        if weather == 4 and time_period == 4 and z < 20:
            actions.append("雪天夜间行人少，雪夜谨慎配时")
            adjustments["pedestrian"] = self.base_pedestrian + 8
            self.speed_limit_display = "reduced"

        # 23. 雪天周末行人适中
        if weather == 4 and time_period == 5 and 20 < z < 40:
            actions.append("雪天周末行人适中，雪天休闲保护")
            adjustments["pedestrian"] = self.base_pedestrian + 12
            self.speed_limit_display = "reduced"

        # 24. 雪天假日行人少
        if weather == 4 and time_period == 6 and z < 30:
            actions.append("雪天假日行人少，节日雪天模式")
            adjustments["pedestrian"] = self.base_pedestrian + 8
            self.speed_limit_display = "reduced"

        # 25-29: 强风各时段组合
        # 25. 强风早高峰行人多
        if weather == 5 and time_period == 1 and z > 40:
            actions.append("强风早高峰行人多，强风保护措施")
            adjustments["pedestrian"] = self.base_pedestrian + 18
            self.warning_light = "on"

        # 26. 强风晚高峰行人密集
        if weather == 5 and time_period == 2 and z > 45:
            actions.append("强风晚高峰行人密集，强风行人保护")
            adjustments["pedestrian"] = self.base_pedestrian + 20
            self.warning_light = "on"

        # 27. 强风午餐时间行人适中
        if weather == 5 and time_period == 3 and 25 < z < 45:
            actions.append("强风午餐时间行人适中，强风轻度保护")
            adjustments["pedestrian"] = self.base_pedestrian + 10
            self.speed_limit_display = "slightly_reduced"

        # 28. 强风周末行人多
        if weather == 5 and time_period == 5 and z > 40:
            actions.append("强风周末行人多，注意行人安全")
            adjustments["pedestrian"] = self.base_pedestrian + 15
            self.warning_light = "on"

        # 29. 强风夜间行人少
        if weather == 5 and time_period == 4 and z < 25:
            actions.append("强风夜间行人少，强风夜间模式")
            adjustments["pedestrian"] = self.base_pedestrian + 5
            self.speed_limit_display = "slightly_reduced"

        # 30-35: 暴风雨各时段组合
        # 30. 暴风雨早高峰行人多
        if weather == 6 and time_period == 1 and z > 30:
            actions.append("暴风雨早高峰行人多，最高级紧急保护")
            adjustments["pedestrian"] = self.base_pedestrian + 30
            self.emergency_light = "on"

        # 31. 暴风雨晚高峰行人多
        if weather == 6 and time_period == 2 and z > 30:
            actions.append("暴风雨晚高峰行人多，紧急保护模式")
            adjustments["pedestrian"] = self.base_pedestrian + 28
            self.emergency_light = "on"

        # 32. 暴风雨午餐时间行人适中
        if weather == 6 and time_period == 3 and 20 < z < 40:
            actions.append("暴风雨午餐时间行人适中，暴雨紧急保护")
            adjustments["pedestrian"] = self.base_pedestrian + 22
            self.emergency_light = "on"

        # 33. 暴风雨夜间行人极少
        if weather == 6 and time_period == 4 and z < 15:
            actions.append("暴风雨夜间行人极少，暴雨紧急最小配时")
            adjustments["pedestrian"] = self.base_pedestrian + 10
            self.emergency_light = "on"

        # 34. 暴风雨周末行人少
        if weather == 6 and time_period == 5 and z < 25:
            actions.append("暴风雨周末行人少，暴雨周末保护")
            adjustments["pedestrian"] = self.base_pedestrian + 15
            self.emergency_light = "on"

        # 35. 暴风雨假日行人适中
        if weather == 6 and time_period == 6 and 15 < z < 35:
            actions.append("暴风雨假日行人适中，暴雨节日保护")
            adjustments["pedestrian"] = self.base_pedestrian + 20
            self.emergency_light = "on"

        # 36. 晴天早高峰基础配时
        if weather == 1 and time_period == 1:
            actions.append("晴天早高峰基础配时优化")
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian + 8

        # 37. 晴天晚高峰基础配时
        if weather == 1 and time_period == 2:
            actions.append("晴天晚高峰基础配时优化")
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian + 6

        # 38. 雨天早高峰基础安全
        if weather == 2 and time_period == 1:
            actions.append("雨天早高峰基础安全措施")
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian + 12
            self.speed_limit_display = "slightly_reduced"

        # 39. 雨天夜间基础保护
        if weather == 2 and time_period == 4:
            actions.append("雨天夜间基础保护模式")
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian + 5
            self.warning_light = "on"

        # 40. 雾天早高峰基础警戒
        if weather == 3 and time_period == 1:
            actions.append("雾天早高峰基础警戒模式")
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian + 20
            self.fog_light = "on"

        # 41. 雾天夜间基础警戒
        if weather == 3 and time_period == 4:
            actions.append("雾天夜间基础警戒模式")
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian + 15
            self.fog_light = "on"
            self.warning_light = "on"

        # 42. 雪天早高峰基础保护
        if weather == 4 and time_period == 1:
            actions.append("雪天早高峰基础保护")
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian + 18
            self.speed_limit_display = "reduced"

        # 43. 雪天周末基础保护
        if weather == 4 and time_period == 5:
            actions.append("雪天周末基础保护")
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian + 10
            self.speed_limit_display = "slightly_reduced"

        # 44. 强风晚高峰基础措施
        if weather == 5 and time_period == 2:
            actions.append("强风晚高峰基础措施")
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian + 12
            self.warning_light = "on"

        # 45. 强风假日基础保护
        if weather == 5 and time_period == 6:
            actions.append("强风假日基础保护")
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian + 8

        # 46. 暴风雨午餐基础紧急
        if weather == 6 and time_period == 3:
            actions.append("暴风雨午餐基础紧急模式")
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian + 25
            self.emergency_light = "on"

        # 47. 暴风雨周末基础紧急
        if weather == 6 and time_period == 5:
            actions.append("暴风雨周末基础紧急模式")
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian + 20
            self.emergency_light = "on"

        # 48-59: weather + z 组合（不考虑time_period）
        # 48. 晴天行人多
        if weather == 1 and z > 50:
            actions.append("晴天行人密集，优化通行效率")
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian + 15

        # 49. 晴天行人少
        if weather == 1 and z < 20:
            actions.append("晴天行人稀少，标准配时")
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian

        # 50. 雨天行人多需保护
        if weather == 2 and z > 40:
            actions.append("雨天行人多需额外保护")
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian + 20
            self.speed_limit_display = "reduced"

        # 51. 雨天行人少仍需保护
        if weather == 2 and z < 25:
            actions.append("雨天行人少仍需基础保护")
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian + 8
            self.warning_light = "on"

        # 52. 雾天行人多高警戒
        if weather == 3 and z > 35:
            actions.append("雾天行人多高度警戒")
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian + 25
            self.fog_light = "on"
            self.emergency_light = "on"

        # 53. 雾天行人少也需警戒
        if weather == 3 and z < 20:
            actions.append("雾天行人少也需警戒")
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian + 12
            self.fog_light = "on"

        # 54. 雪天行人多强化保护
        if weather == 4 and z > 35:
            actions.append("雪天行人多强化保护措施")
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian + 22
            self.speed_limit_display = "very_reduced"

        # 55. 雪天行人少基础保护
        if weather == 4 and z < 20:
            actions.append("雪天行人少基础保护")
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian + 10
            self.speed_limit_display = "reduced"

        # 56. 强风行人多警告措施
        if weather == 5 and z > 40:
            actions.append("强风行人多警告措施")
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian + 18
            self.warning_light = "on"

        # 57. 强风行人少轻度措施
        if weather == 5 and z < 25:
            actions.append("强风行人少轻度保护措施")
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian + 6

        # 58. 暴风雨行人多最高警戒
        if weather == 6 and z > 30:
            actions.append("暴风雨行人多最高警戒")
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian + 35
            self.emergency_light = "on"
            self.speed_limit_display = "very_reduced"

        # 59. 暴风雨行人少紧急保护
        if weather == 6 and z < 20:
            actions.append("暴风雨行人少紧急保护")
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian + 15
            self.emergency_light = "on"

        # 60-71: time_period + z 组合（不考虑weather）
        # 60. 早高峰行人超多
        if time_period == 1 and z > 50:
            actions.append("早高峰行人超多，最大通行时间")
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian + 20

        # 61. 早高峰行人少
        if time_period == 1 and z < 20:
            actions.append("早高峰行人少，车辆优先")
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian + 5

        # 62. 晚高峰行人密集
        if time_period == 2 and z > 45:
            actions.append("晚高峰行人密集，延长行人时间")
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian + 18

        # 63. 晚高峰行人稀少
        if time_period == 2 and z < 25:
            actions.append("晚高峰行人稀少，车辆优先配时")
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian + 3

        # 64. 午餐时间行人爆满
        if time_period == 3 and z > 55:
            actions.append("午餐时间行人爆满，最大行人配时")
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian + 25

        # 65. 午餐时间行人少
        if time_period == 3 and z < 20:
            actions.append("午餐时间行人少，标准配时")
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian + 2

        # 66. 夜间行人多
        if time_period == 4 and z > 35:
            actions.append("夜间行人多，夜间行人保护")
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian + 12
            self.warning_light = "on"

        # 67. 夜间行人极少
        if time_period == 4 and z < 10:
            actions.append("夜间行人极少，最小配时")
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian - 5

        # 68. 周末行人密集
        if time_period == 5 and z > 50:
            actions.append("周末行人密集，休闲行人优先")
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian + 20

        # 69. 周末行人少
        if time_period == 5 and z < 25:
            actions.append("周末行人少，标准休闲配时")
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian + 5

        # 70. 假日行人爆满
        if time_period == 6 and z > 60:
            actions.append("假日行人爆满，节日行人优先")
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian + 30

        # 71. 假日行人少
        if time_period == 6 and z < 20:
            actions.append("假日行人少，节日标准配时")
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian + 8

        # ===== 新增补充条件：对应第二个代码第三层的缺失条件类型 =====

        # 72. 单变量z条件 - z < 40（中等偏少行人）适配天气时段
        if z < 40:
            actions.append("行人数量中等偏少，根据天气时段调整行人配时")
            base_adjustment = self.base_pedestrian - 2

            # 根据天气调整
            if weather in [2, 3, 4, 6]:  # 恶劣天气仍需保护
                base_adjustment = self.base_pedestrian + 3
            elif weather in [1, 5]:  # 晴天和强风适中调整
                base_adjustment = self.base_pedestrian - 1

            # 根据时段进一步调整
            if time_period in [1, 2]:  # 高峰期稍微增加
                base_adjustment += 5
            elif time_period == 4:  # 夜间减少
                base_adjustment -= 3

            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = max(base_adjustment, self.base_pedestrian - 5)

        # 73. 单变量z条件 - z > 20（不算很少行人）适配天气时段
        if z > 20:
            actions.append("行人数量不算很少，保障基础行人权益并考虑天气时段")
            base_adjustment = self.base_pedestrian + 3

            # 根据天气调整
            if weather in [2, 3, 4, 6]:  # 恶劣天气加强保护
                base_adjustment = self.base_pedestrian + 8
            elif weather == 1:  # 晴天标准保护
                base_adjustment = self.base_pedestrian + 3
            elif weather == 5:  # 强风适中保护
                base_adjustment = self.base_pedestrian + 5

            # 根据时段调整
            if time_period in [1, 2]:  # 高峰期
                base_adjustment += 5
            elif time_period in [5, 6]:  # 休息时段
                base_adjustment += 3
            elif time_period == 4:  # 夜间
                base_adjustment += 2

            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = base_adjustment

        # 74. 范围条件 - 10 < z < 50（中等行人范围）适配天气时段
        if 10 < z < 50:
            actions.append("行人数量在中等范围，根据天气时段平衡配时")
            base_adjustment = self.base_pedestrian + 6

            # 根据天气精细化调整
            if weather == 1:  # 晴天
                base_adjustment = self.base_pedestrian + 5
            elif weather == 2:  # 雨天
                base_adjustment = self.base_pedestrian + 10
            elif weather == 3:  # 雾天
                base_adjustment = self.base_pedestrian + 15
            elif weather == 4:  # 雪天
                base_adjustment = self.base_pedestrian + 12
            elif weather == 5:  # 强风
                base_adjustment = self.base_pedestrian + 8
            elif weather == 6:  # 暴风雨
                base_adjustment = self.base_pedestrian + 18

            # 根据时段调整
            if time_period in [1, 2]:  # 高峰期延长
                base_adjustment += 8
            elif time_period == 3:  # 午餐时间
                base_adjustment += 6
            elif time_period in [5, 6]:  # 休息时段
                base_adjustment += 4
            elif time_period == 4:  # 夜间
                base_adjustment += 2

            # 在中等范围内根据具体值微调
            if z < 25:
                base_adjustment -= 3  # 偏少的中等范围
            elif z > 35:
                base_adjustment += 3  # 偏多的中等范围

            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = base_adjustment

        # 75. 范围条件 - 15 < z < 45（常规行人范围）适配天气时段
        if 15 < z < 45:
            actions.append("行人数量在常规范围，标准天气时段行人保护")
            base_adjustment = self.base_pedestrian + 7

            # 根据天气调整
            if weather == 1:  # 晴天标准
                base_adjustment = self.base_pedestrian + 6
            elif weather == 2:  # 雨天增强
                base_adjustment = self.base_pedestrian + 12
            elif weather == 3:  # 雾天强化
                base_adjustment = self.base_pedestrian + 18
            elif weather == 4:  # 雪天保护
                base_adjustment = self.base_pedestrian + 15
            elif weather == 5:  # 强风警示
                base_adjustment = self.base_pedestrian + 10
            elif weather == 6:  # 暴雨紧急
                base_adjustment = self.base_pedestrian + 22

            # 根据时段精细调整
            if time_period == 1:  # 早高峰
                base_adjustment += 10
            elif time_period == 2:  # 晚高峰
                base_adjustment += 8
            elif time_period == 3:  # 午餐时间
                base_adjustment += 12
            elif time_period == 4:  # 夜间
                base_adjustment += 5
                if weather in [2, 3, 4, 6]:  # 恶劣天气夜间额外保护
                    base_adjustment += 5
            elif time_period == 5:  # 周末
                base_adjustment += 7
            elif time_period == 6:  # 假日
                base_adjustment += 9

            # 常规范围内的精细调整
            if z < 25:
                base_adjustment -= 2  # 常规偏少
            elif z > 35:
                base_adjustment += 4  # 常规偏多

            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = base_adjustment

        # ===== 特殊天气时段组合补充 =====

        # 76. 恶劣天气 + z < 40（中等偏少行人）
        if weather in [2, 3, 4, 6] and z < 40:
            actions.append("恶劣天气下行人中等偏少，仍需基础保护")
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian + 8
            # 启动相应警示
            if weather == 2:
                self.speed_limit_display = "slightly_reduced"
            elif weather == 3:
                self.fog_light = "on"
            elif weather == 4:
                self.speed_limit_display = "reduced"
            elif weather == 6:
                self.emergency_light = "on"

        # 77. 高峰时段 + z > 20（不算很少行人）
        if time_period in [1, 2] and z > 20:
            actions.append("高峰时段行人不算很少，加强通行保障")
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian + 12

        # 78. 夜间时段 + 10 < z < 50（中等行人范围）
        if time_period == 4 and 10 < z < 50:
            actions.append("夜间中等行人范围，夜间行人保护模式")
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian + 8
            self.warning_light = "on"

        # 79. 休息时段 + 15 < z < 45（常规行人范围）
        if time_period in [5, 6] and 15 < z < 45:
            actions.append("休息时段常规行人范围，舒适通行模式")
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian + 10

        # 80. 多重恶劣条件组合
        if weather in [3, 4, 6] and time_period in [1, 2, 4] and z > 20:
            actions.append("恶劣天气关键时段有行人，最高保护级别")
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian + 25
            self.emergency_light = "on"
            if weather == 3:
                self.fog_light = "on"

        return actions, adjustments

    def block4_time_x_y_combinations(self, x, y, z, weather, time_period, special_event,
                                     vehicle_types, emergency_vehicles, air_quality,
                                     noise_level, visibility, temperature, wind_speed,
                                     school_zones_active, business_district_active,
                                     hospital_nearby, construction_zone, accident_probability, road_condition):
        """大块4: time_period(数字), X, Y组合处理（只包含这三个变量）- 大规模扩展版 + 完整覆盖"""
        actions = []
        adjustments = {}

        # === 原有的35个IF语句保持不变 ===
        # time_period: 1=rush_morning, 2=rush_evening, 3=lunch, 4=night, 5=weekend, 6=holiday

        # 1-8: 早高峰组合
        # 1. 早高峰主路拥堵辅路畅通
        if time_period == 1 and x > 75 and y < 50:
            actions.append("早高峰主路拥堵辅路畅通，最大化主路通行")
            adjustments["main"] = self.max_main + 20
            adjustments["secondary"] = self.min_secondary - 3

        # 2. 早高峰主路畅通辅路拥堵
        if time_period == 1 and x < 50 and y > 75:
            actions.append("早高峰主路畅通辅路拥堵，优先辅路通行")
            adjustments["main"] = self.min_main - 2
            adjustments["secondary"] = self.max_secondary + 15

        # 3. 早高峰主辅路均拥堵
        if time_period == 1 and x > 70 and y > 70:
            actions.append("早高峰主辅路均拥堵，均衡高效分配")
            adjustments["main"] = 55
            adjustments["secondary"] = 48

        # 4. 早高峰主辅路均低流量（异常）
        if time_period == 1 and x < 40 and y < 40:
            actions.append("早高峰主辅路异常低流量，检查系统")
            adjustments["main"] = 30
            adjustments["secondary"] = 25
            self.warning_light = "on"

        # 5. 早高峰主路极高辅路中等
        if time_period == 1 and x > 88 and 45 < y < 70:
            actions.append("早高峰主路极高辅路中等，主路绝对优先")
            adjustments["main"] = self.max_main + 25
            adjustments["secondary"] = self.min_secondary - 5
            self.traffic_police_alert = "on"

        # 6. 早高峰主路中等辅路极高
        if time_period == 1 and 45 < x < 70 and y > 88:
            actions.append("早高峰主路中等辅路极高，辅路绝对优先")
            adjustments["main"] = self.min_main - 5
            adjustments["secondary"] = self.max_secondary + 25
            self.traffic_police_alert = "on"

        # 7. 早高峰主路低辅路中等
        if time_period == 1 and x < 45 and 50 < y < 75:
            actions.append("早高峰主路低辅路中等，辅路优先")
            adjustments["main"] = self.min_main + 3
            adjustments["secondary"] = self.max_secondary - 5

        # 8. 早高峰主路中等辅路低
        if time_period == 1 and 50 < x < 75 and y < 45:
            actions.append("早高峰主路中等辅路低，主路优先")
            adjustments["main"] = self.max_main - 5
            adjustments["secondary"] = self.min_secondary + 3

        # 9-16: 晚高峰组合
        # 9. 晚高峰主路拥堵辅路畅通
        if time_period == 2 and x > 75 and y < 50:
            actions.append("晚高峰主路拥堵辅路畅通，延长主路绿灯")
            adjustments["main"] = self.max_main + 15
            adjustments["secondary"] = self.min_secondary

        # 10. 晚高峰主路畅通辅路拥堵
        if time_period == 2 and x < 50 and y > 75:
            actions.append("晚高峰主路畅通辅路拥堵，延长辅路绿灯")
            adjustments["main"] = self.min_main + 3
            adjustments["secondary"] = self.max_secondary + 12

        # 11. 晚高峰主辅路均拥堵
        if time_period == 2 and x > 70 and y > 70:
            actions.append("晚高峰主辅路均拥堵，平衡分配")
            adjustments["main"] = 50
            adjustments["secondary"] = 45

        # 12. 晚高峰主辅路均低流量
        if time_period == 2 and x < 45 and y < 45:
            actions.append("晚高峰主辅路低流量，轻度优化配时")
            adjustments["main"] = 35
            adjustments["secondary"] = 30

        # 13. 晚高峰主路中等辅路极高
        if time_period == 2 and 45 < x < 70 and y > 88:
            actions.append("晚高峰主路中等辅路极高，辅路绝对优先")
            adjustments["main"] = self.min_secondary - 5
            adjustments["secondary"] = self.max_secondary + 20
            self.traffic_police_alert = "on"

        # 14. 晚高峰主路极高辅路中等
        if time_period == 2 and x > 88 and 45 < y < 70:
            actions.append("晚高峰主路极高辅路中等，主路绝对优先")
            adjustments["main"] = self.max_main + 20
            adjustments["secondary"] = self.min_secondary - 5
            self.traffic_police_alert = "on"

        # 15. 晚高峰主路适中辅路低
        if time_period == 2 and 55 < x < 75 and y < 40:
            actions.append("晚高峰主路适中辅路低，主路轻度优先")
            adjustments["main"] = self.max_main - 8
            adjustments["secondary"] = self.min_secondary + 5

        # 16. 晚高峰主路低辅路适中
        if time_period == 2 and x < 40 and 55 < y < 75:
            actions.append("晚高峰主路低辅路适中，辅路轻度优先")
            adjustments["main"] = self.min_main + 5
            adjustments["secondary"] = self.max_secondary - 8

        # 17-21: 午餐时间组合
        # 17. 午餐时间主路拥堵辅路适中
        if time_period == 3 and x > 70 and 40 < y < 65:
            actions.append("午餐时间主路拥堵辅路适中，优先主路")
            adjustments["main"] = self.max_main - 5
            adjustments["secondary"] = 32

        # 18. 午餐时间主路适中辅路拥堵
        if time_period == 3 and 40 < x < 65 and y > 70:
            actions.append("午餐时间主路适中辅路拥堵，优先辅路")
            adjustments["main"] = 32
            adjustments["secondary"] = self.max_secondary - 5

        # 19. 午餐时间主辅路均适中
        if time_period == 3 and 45 < x < 70 and 45 < y < 70:
            actions.append("午餐时间主辅路均适中，均衡配时")
            adjustments["main"] = 40
            adjustments["secondary"] = 35

        # 20. 午餐时间主辅路均低
        if time_period == 3 and x < 40 and y < 40:
            actions.append("午餐时间主辅路均低，舒适配时")
            adjustments["main"] = 32
            adjustments["secondary"] = 28

        # 21. 午餐时间主路高辅路低
        if time_period == 3 and x > 75 and y < 45:
            actions.append("午餐时间主路高辅路低，主路优先")
            adjustments["main"] = self.max_main - 8
            adjustments["secondary"] = self.min_secondary + 5

        # 22-27: 夜间组合
        # 22. 夜间主辅路均低流量
        if time_period == 4 and x < 35 and y < 35:
            actions.append("夜间主辅路均低流量，最小夜间周期")
            adjustments["main"] = 22
            adjustments["secondary"] = 18

        # 23. 夜间主路拥堵辅路畅通（异常）
        if time_period == 4 and x > 70 and y < 40:
            actions.append("夜间主路异常拥堵辅路畅通，特殊处理")
            adjustments["main"] = self.max_main - 15
            adjustments["secondary"] = self.min_secondary + 5
            self.warning_light = "on"

        # 24. 夜间主路畅通辅路拥堵（异常）
        if time_period == 4 and x < 40 and y > 70:
            actions.append("夜间主路畅通辅路异常拥堵，特殊处理")
            adjustments["main"] = self.min_main + 5
            adjustments["secondary"] = self.max_secondary - 15
            self.warning_light = "on"

        # 25. 夜间主辅路均中等（异常）
        if time_period == 4 and 45 < x < 70 and 45 < y < 70:
            actions.append("夜间主辅路异常中等流量，检查情况")
            adjustments["main"] = 35
            adjustments["secondary"] = 30
            self.warning_light = "on"

        # 26. 夜间主路极低辅路极低
        if time_period == 4 and x < 25 and y < 25:
            actions.append("夜间主辅路极低流量，极小夜间配时")
            adjustments["main"] = 18
            adjustments["secondary"] = 15

        # 27. 夜间主路中等辅路低
        if time_period == 4 and 40 < x < 65 and y < 35:
            actions.append("夜间主路中等辅路低，夜间主路优先")
            adjustments["main"] = 35
            adjustments["secondary"] = 20

        # 28-30: 周末组合
        # 28. 周末主辅路均低流量
        if time_period == 5 and x < 50 and y < 50:
            actions.append("周末主辅路均低流量，休闲模式")
            adjustments["main"] = 35
            adjustments["secondary"] = 30
            adjustments["bicycle"] = self.base_bicycle + 10

        # 29. 周末主辅路均高流量
        if time_period == 5 and x > 75 and y > 75:
            actions.append("周末主辅路均高流量，周末拥堵模式")
            adjustments["main"] = 45
            adjustments["secondary"] = 40

        # 30. 周末主路高辅路低
        if time_period == 5 and x > 70 and y < 45:
            actions.append("周末主路高辅路低，周末主路优先")
            adjustments["main"] = self.max_main - 10
            adjustments["secondary"] = self.min_secondary + 8

        # 31-35: 假日组合
        # 31. 假日主辅路适中流量
        if time_period == 6 and 40 < x < 70 and 40 < y < 70:
            actions.append("假日主辅路适中流量，节日标准配时")
            adjustments["main"] = 40
            adjustments["secondary"] = 35

        # 32. 假日主辅路均低流量
        if time_period == 6 and x < 45 and y < 45:
            actions.append("假日主辅路均低流量，节日休闲配时")
            adjustments["main"] = 32
            adjustments["secondary"] = 28

        # 33. 假日主路高辅路低
        if time_period == 6 and x > 70 and y < 50:
            actions.append("假日主路高辅路低，节日主路优先")
            adjustments["main"] = self.max_main - 12
            adjustments["secondary"] = self.min_secondary + 8

        # 34. 假日主路低辅路高
        if time_period == 6 and x < 50 and y > 70:
            actions.append("假日主路低辅路高，节日辅路优先")
            adjustments["main"] = self.min_main + 8
            adjustments["secondary"] = self.max_secondary - 12

        # 35. 假日主辅路均高流量
        if time_period == 6 and x > 75 and y > 75:
            actions.append("假日主辅路均高流量，节日拥堵管理")
            adjustments["main"] = 42
            adjustments["secondary"] = 38

        # ===== 新增补充条件：覆盖缺失的14种条件类型 =====

        # 36-41: 天气条件补充 (weather == 1, 2, 3, 4, 5, 6)
        # 36. 晴天时间段优化
        if weather == 1:
            actions.append("晴天条件，优化标准配时")
            if "main" not in adjustments:
                adjustments["main"] = self.max_main - 5
            if "secondary" not in adjustments:
                adjustments["secondary"] = self.max_secondary - 5

        # 37. 雨天时间段保护
        if weather == 2:
            actions.append("雨天条件，启动安全保护配时")
            if "main" not in adjustments:
                adjustments["main"] = max(self.max_main - 10, 30)
            if "secondary" not in adjustments:
                adjustments["secondary"] = max(self.max_secondary - 8, 25)
            self.speed_limit_display = "reduced"

        # 38. 雾天时间段警戒
        if weather == 3:
            actions.append("雾天条件，启动高度警戒配时")
            if "main" not in adjustments:
                adjustments["main"] = max(self.max_main - 15, 25)
            if "secondary" not in adjustments:
                adjustments["secondary"] = max(self.max_secondary - 12, 22)
            self.fog_light = "on"
            self.emergency_light = "on"

        # 39. 雪天时间段保护
        if weather == 4:
            actions.append("雪天条件，启动最高保护配时")
            if "main" not in adjustments:
                adjustments["main"] = max(self.max_main - 20, 25)
            if "secondary" not in adjustments:
                adjustments["secondary"] = max(self.max_secondary - 18, 20)
            self.speed_limit_display = "very_reduced"
            self.warning_light = "on"

        # 40. 强风时间段警示
        if weather == 5:
            actions.append("强风条件，启动风险警示配时")
            if "main" not in adjustments:
                adjustments["main"] = max(self.max_main - 8, 30)
            if "secondary" not in adjustments:
                adjustments["secondary"] = max(self.max_secondary - 6, 28)
            self.warning_light = "on"

        # 41. 暴风雨时间段紧急
        if weather == 6:
            actions.append("暴风雨条件，启动紧急配时")
            if "main" not in adjustments:
                adjustments["main"] = max(self.max_main - 25, 20)
            if "secondary" not in adjustments:
                adjustments["secondary"] = max(self.max_secondary - 22, 18)
            self.emergency_light = "on"
            self.speed_limit_display = "very_reduced"

        # 42-44: 天气集合条件 (weather in [1, 3, 5], weather in [2, 4, 6])
        # 42. 奇数天气条件
        if weather in [1, 3, 5]:
            actions.append("奇数天气条件，启动奇数天气策略")
            if time_period in [1, 2] and x > 70:
                if "main" not in adjustments:
                    adjustments["main"] = self.max_main + 3

        # 43. 偶数天气条件
        if weather in [2, 4, 6]:
            actions.append("偶数天气条件，启动偶数天气策略")
            if time_period in [1, 2] and y > 70:
                if "secondary" not in adjustments:
                    adjustments["secondary"] = self.max_secondary + 3

        # 44. 极端天气组合
        if weather in [3, 4, 6]:
            actions.append("极端天气条件，启动特殊保护")
            if "main" not in adjustments:
                adjustments["main"] = max(self.max_main - 12, 25)
            if "secondary" not in adjustments:
                adjustments["secondary"] = max(self.max_secondary - 10, 22)

        # 45-46: 复合条件 (weather + time_period > 6, <= 6)
        # 45. 天气时间段值较大
        if weather + time_period > 6:
            actions.append("天气时间段复合值较大，高级别管制")
            if x > 50 and y > 50:
                if "main" not in adjustments:
                    adjustments["main"] = max(self.max_main - 8, 30)
                if "secondary" not in adjustments:
                    adjustments["secondary"] = max(self.max_secondary - 6, 28)

        # 46. 天气时间段值较小
        if weather + time_period <= 6:
            actions.append("天气时间段复合值较小，标准管制")
            if x < 50 and y < 50:
                if "main" not in adjustments:
                    adjustments["main"] = 35
                if "secondary" not in adjustments:
                    adjustments["secondary"] = 30

        # 47-48: 数值关系条件 (weather % 2 == time_period % 2)
        # 47. 天气时间段同奇偶性
        if weather % 2 == time_period % 2:
            actions.append("天气时间段同奇偶性，协调配时")
            if x % 10 < 5 and y % 10 >= 5:
                if "main" not in adjustments:
                    adjustments["main"] = 40
                if "secondary" not in adjustments:
                    adjustments["secondary"] = 38

        # 48. 天气时间段不同奇偶性
        if weather % 2 != time_period % 2:
            actions.append("天气时间段不同奇偶性，对比配时")
            if x > 75 or y > 75:
                if "main" not in adjustments:
                    adjustments["main"] = self.max_main + 5
                if "secondary" not in adjustments:
                    adjustments["secondary"] = self.max_secondary + 5

        # 49-50: x和y的比较条件 (x > y, x < y)
        # 49. 主路流量大于辅路
        if x > y:
            actions.append("主路流量大于辅路，主路优先策略")
            flow_diff = x - y
            if flow_diff > 30:
                if "main" not in adjustments:
                    adjustments["main"] = self.max_main + int(flow_diff / 10)
                if "secondary" not in adjustments:
                    adjustments["secondary"] = max(self.max_secondary - int(flow_diff / 15), self.min_secondary)

        # 50. 辅路流量大于主路
        if x < y:
            actions.append("辅路流量大于主路，辅路优先策略")
            flow_diff = y - x
            if flow_diff > 30:
                if "secondary" not in adjustments:
                    adjustments["secondary"] = self.max_secondary + int(flow_diff / 10)
                if "main" not in adjustments:
                    adjustments["main"] = max(self.max_main - int(flow_diff / 15), self.min_main)

        # 51-52: 距离条件 (abs(x - y) < 20, abs(x - y) > 30)
        # 51. 主辅路流量接近
        if abs(x - y) < 20:
            actions.append("主辅路流量接近，均衡配时策略")
            if "main" not in adjustments:
                adjustments["main"] = 42
            if "secondary" not in adjustments:
                adjustments["secondary"] = 38

        # 52. 主辅路流量差距大
        if abs(x - y) > 30:
            actions.append("主辅路流量差距大，差异化配时")
            if x > y:
                if "main" not in adjustments:
                    adjustments["main"] = self.max_main + 8
                if "secondary" not in adjustments:
                    adjustments["secondary"] = max(self.max_secondary - 8, self.min_secondary)
            else:
                if "secondary" not in adjustments:
                    adjustments["secondary"] = self.max_secondary + 8
                if "main" not in adjustments:
                    adjustments["main"] = max(self.max_main - 8, self.min_main)

        # 53-55: 求和条件 (x + y > 100, x + y < 80, (x + y) // 2 > 50)
        # 53. 总流量高
        if x + y > 100:
            actions.append("总流量高，高强度管制模式")
            if "main" not in adjustments:
                adjustments["main"] = self.max_main + 5
            if "secondary" not in adjustments:
                adjustments["secondary"] = self.max_secondary + 3
            self.traffic_police_alert = "on"

        # 54. 总流量低
        if x + y < 80:
            actions.append("总流量低，节能优化模式")
            if "main" not in adjustments:
                adjustments["main"] = max(self.max_main - 10, 30)
            if "secondary" not in adjustments:
                adjustments["secondary"] = max(self.max_secondary - 8, 25)

        # 55. 平均流量高
        if (x + y) // 2 > 50:
            actions.append("平均流量高，平衡高效模式")
            if "main" not in adjustments:
                adjustments["main"] = 48
            if "secondary" not in adjustments:
                adjustments["secondary"] = 42

        # 56-58: 模运算条件 (x % 10 < 5, (x + y) % 3 == 0, (x * y) % 7 == 0)
        # 56. x模运算条件
        if x % 10 < 5:
            actions.append("主路流量个位数小于5，精细调节")
            if "main" not in adjustments:
                adjustments["main"] = self.max_main - 3

        # 57. 总和模运算条件
        if (x + y) % 3 == 0:
            actions.append("总流量能被3整除，标准化配时")
            if "main" not in adjustments:
                adjustments["main"] = 45
            if "secondary" not in adjustments:
                adjustments["secondary"] = 39

        # 58. 乘积模运算条件
        if (x * y) % 7 == 0:
            actions.append("流量乘积能被7整除，特殊同步配时")
            if "main" not in adjustments:
                adjustments["main"] = 42
            if "secondary" not in adjustments:
                adjustments["secondary"] = 42

        # 59-61: 复杂数学条件
        # 59. 三元复合模运算
        if (weather * time_period + z) % 7 == 0:
            actions.append("天气时段行人复合模运算匹配，高级同步")
            if "main" not in adjustments:
                adjustments["main"] = 40
            if "secondary" not in adjustments:
                adjustments["secondary"] = 40
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian + 10

        # 60. 参数相等条件
        if weather + time_period == z // 10:
            actions.append("天气时段和等于行人十位数，精准匹配")
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian + 15

        # 61. 绝对值距离条件
        if abs(weather - time_period) * 10 <= z:
            actions.append("天气时段差值与行人数匹配，协调配时")
            if "pedestrian" not in adjustments:
                adjustments["pedestrian"] = self.base_pedestrian + 8

        # 62-64: 复合逻辑条件 (x > 50 and y > 50, x < 50 and y < 50)
        # 62. 主辅路均中高流量
        if x > 50 and y > 50:
            actions.append("主辅路均中高流量，双重强化")
            if "main" not in adjustments:
                adjustments["main"] = self.max_main + 3
            if "secondary" not in adjustments:
                adjustments["secondary"] = self.max_secondary + 3

        # 63. 主辅路均低流量
        if x < 50 and y < 50:
            actions.append("主辅路均低流量，双重优化")
            if "main" not in adjustments:
                adjustments["main"] = max(self.max_main - 8, 30)
            if "secondary" not in adjustments:
                adjustments["secondary"] = max(self.max_secondary - 6, 25)

        # 64. 任一路高流量
        if x > 60 or y > 60:
            actions.append("任一路高流量，灵活响应")
            if x > 60:
                if "main" not in adjustments:
                    adjustments["main"] = self.max_main + 5
            if y > 60:
                if "secondary" not in adjustments:
                    adjustments["secondary"] = self.max_secondary + 5

        # 65-67: 或逻辑条件扩展
        # 65. 任一路低流量
        if x < 40 or y < 40:
            actions.append("任一路低流量，节能配时")
            if x < 40 and y >= 40:
                if "main" not in adjustments:
                    adjustments["main"] = max(self.max_main - 10, 25)
            elif y < 40 and x >= 40:
                if "secondary" not in adjustments:
                    adjustments["secondary"] = max(self.max_secondary - 8, 22)

        # 66. 任一路极高流量
        if x > 75 or y > 75:
            actions.append("任一路极高流量，紧急响应")
            if x > 75:
                if "main" not in adjustments:
                    adjustments["main"] = self.max_main + 12
                self.traffic_police_alert = "on"
            if y > 75:
                if "secondary" not in adjustments:
                    adjustments["secondary"] = self.max_secondary + 10
                self.traffic_police_alert = "on"

        # 67. 任一路极低流量
        if x < 25 or y < 25:
            actions.append("任一路极低流量，最小配时")
            if x < 25:
                if "main" not in adjustments:
                    adjustments["main"] = 20
            if y < 25:
                if "secondary" not in adjustments:
                    adjustments["secondary"] = 18

        # 68: 最大最小值条件 (max(x, y) - min(x, y) > 40)
        # 68. 流量差距极大
        if max(x, y) - min(x, y) > 40:
            actions.append("流量差距极大，极端优先配时")
            max_flow = max(x, y)
            min_flow = min(x, y)
            if x == max_flow:
                if "main" not in adjustments:
                    adjustments["main"] = self.max_main + 15
                if "secondary" not in adjustments:
                    adjustments["secondary"] = max(self.min_secondary - 5, 10)
            else:
                if "secondary" not in adjustments:
                    adjustments["secondary"] = self.max_secondary + 15
                if "main" not in adjustments:
                    adjustments["main"] = max(self.min_main - 5, 10)
            self.traffic_police_alert = "on"

        # 69: 复合乘积条件
        if weather * time_period > 15:
            actions.append("天气时段乘积值大，高级管制")
            if "main" not in adjustments:
                adjustments["main"] = max(self.max_main - 5, 30)
            if "secondary" not in adjustments:
                adjustments["secondary"] = max(self.max_secondary - 3, 28)

        # 70: 复合最值条件
        if max(weather, time_period) * min(x, y) > 150:
            actions.append("复合最值条件满足，智能优化配时")
            min_flow = min(x, y)
            if min_flow == x:
                if "main" not in adjustments:
                    adjustments["main"] = self.max_main + 8
            else:
                if "secondary" not in adjustments:
                    adjustments["secondary"] = self.max_secondary + 8

        return actions, adjustments

    def block5_special_emergency_visibility_combinations(self, x, y, z, weather, time_period, special_event,
                                                         vehicle_types, emergency_vehicles, air_quality,
                                                         noise_level, visibility, temperature, wind_speed,
                                                         school_zones_active, business_district_active,
                                                         hospital_nearby, construction_zone, accident_probability,
                                                         road_condition):
        """大块5: special_event_code(1-100), emergency_vehicles(1-400), visibility(1-1000)组合处理 - 大规模扩展版"""
        actions = []
        adjustments = {}

        # === 注意：需要修改get_sensor_data()函数 ===
        # special_event_code应该是1-100的整数
        # emergency_vehicles应该是1-400的整数
        # visibility应该是1-1000的整数

        # 假设special_event现在是数字编码1-100
        # 1-20: 常规事件, 21-40: 小型活动, 41-60: 中型活动, 61-80: 大型活动, 81-100: 紧急/危机事件

        # === Original 50 IF statements from second code ===

        # 1-10: 低特殊事件值组合 (1-30)
        # 1. 低事件+低紧急车辆+高能见度
        if special_event < 30 and emergency_vehicles < 100 and visibility > 700:
            actions.append("低级别事件，少量紧急车辆，能见度优秀，标准优化配时")
            adjustments["main"] = self.max_main - 5
            adjustments["secondary"] = self.max_secondary - 3

        # 2. 低事件+低紧急车辆+低能见度
        if special_event < 25 and emergency_vehicles < 80 and visibility < 300:
            actions.append("低级别事件，少量紧急车辆，能见度差，启动安全保护")
            adjustments["main"] = max(self.max_main - 20, 25)
            adjustments["secondary"] = max(self.max_secondary - 18, 22)
            self.fog_light = "on"
            self.speed_limit_display = "very_reduced"

        # 3. 低事件+中等紧急车辆+高能见度
        if special_event < 30 and 100 <= emergency_vehicles < 250 and visibility > 650:
            actions.append("低级别事件，中等紧急车辆，能见度良好，优先紧急通道")
            adjustments["main"] = self.max_main + 10
            adjustments["secondary"] = self.max_secondary - 8
            self.emergency_light = "on"

        # 4. 低事件+高紧急车辆+中等能见度
        if special_event < 30 and emergency_vehicles >= 250 and 400 < visibility < 700:
            actions.append("低级别事件，大量紧急车辆，中等能见度，强化应急通道")
            adjustments["main"] = self.max_main + 20
            adjustments["secondary"] = self.min_secondary - 5
            self.emergency_light = "on"
            self.traffic_police_alert = "on"

        # 5. 低事件+极高紧急车辆+低能见度
        if special_event < 25 and emergency_vehicles > 300 and visibility < 350:
            actions.append("低级别事件，极多紧急车辆，低能见度，紧急谨慎模式")
            adjustments["main"] = 30
            adjustments["secondary"] = 25
            self.emergency_light = "on"
            self.fog_light = "on"
            self.traffic_police_alert = "on"

        # 6. 低事件+低紧急车辆+极低能见度
        if special_event < 30 and emergency_vehicles < 100 and visibility < 200:
            actions.append("低级别事件，少量紧急车辆，极低能见度，最高警戒")
            adjustments["main"] = 25
            adjustments["secondary"] = 20
            self.fog_light = "on"
            self.warning_light = "on"
            self.speed_limit_display = "very_reduced"

        # 7. 低事件+中等紧急车辆+中等能见度
        if special_event < 30 and 100 <= emergency_vehicles < 250 and 400 < visibility < 700:
            actions.append("低级别事件，中等紧急车辆，中等能见度，标准应急配时")
            adjustments["main"] = self.max_main + 5
            adjustments["secondary"] = self.max_secondary - 5
            self.emergency_light = "on"

        # 8. 低事件+高紧急车辆+高能见度
        if special_event < 30 and emergency_vehicles >= 250 and visibility > 750:
            actions.append("低级别事件，大量紧急车辆，能见度优秀，最优应急通道")
            adjustments["main"] = self.max_main + 25
            adjustments["secondary"] = self.min_secondary - 8
            self.emergency_light = "on"
            self.traffic_police_alert = "on"

        # 9. 低事件+极低紧急车辆+极高能见度
        if special_event < 20 and emergency_vehicles < 50 and visibility > 850:
            actions.append("极低级别事件，极少紧急车辆，能见度极佳，最优标准配时")
            adjustments["main"] = self.max_main
            adjustments["secondary"] = self.max_secondary

        # 10. 低事件+紧急车辆临界+能见度临界
        if special_event < 30 and 90 < emergency_vehicles < 110 and 380 < visibility < 420:
            actions.append("低级别事件，紧急车辆临界值，能见度临界值，精细平衡")
            adjustments["main"] = 42
            adjustments["secondary"] = 38
            self.warning_light = "on"

        # 11-20: 中低特殊事件值组合 (31-50)
        # 11. 中低事件+低紧急车辆+高能见度
        if 31 <= special_event <= 50 and emergency_vehicles < 100 and visibility > 700:
            actions.append("中低级别事件，少量紧急车辆，能见度良好，轻度疏导")
            adjustments["main"] = self.max_main - 8
            adjustments["secondary"] = self.max_secondary - 6
            adjustments["pedestrian"] = self.base_pedestrian + 10

        # 12. 中低事件+低紧急车辆+低能见度
        if 31 <= special_event <= 50 and emergency_vehicles < 100 and visibility < 300:
            actions.append("中低级别事件，少量紧急车辆，低能见度，事件安全保护")
            adjustments["main"] = max(self.max_main - 22, 25)
            adjustments["secondary"] = max(self.max_secondary - 20, 22)
            adjustments["pedestrian"] = self.base_pedestrian + 15
            self.fog_light = "on"
            self.warning_light = "on"

        # 13. 中低事件+中等紧急车辆+高能见度
        if 31 <= special_event <= 50 and 100 <= emergency_vehicles < 250 and visibility > 650:
            actions.append("中低级别事件，中等紧急车辆，能见度良好，事件应急优先")
            adjustments["main"] = self.max_main + 8
            adjustments["secondary"] = self.max_secondary - 10
            adjustments["pedestrian"] = self.base_pedestrian + 12
            self.emergency_light = "on"

        # 14. 中低事件+高紧急车辆+低能见度
        if 31 <= special_event <= 50 and emergency_vehicles >= 250 and visibility < 350:
            actions.append("中低级别事件，大量紧急车辆，低能见度，事件紧急警戒")
            adjustments["main"] = 32
            adjustments["secondary"] = 27
            adjustments["pedestrian"] = self.base_pedestrian + 15
            self.emergency_light = "on"
            self.fog_light = "on"
            self.traffic_police_alert = "on"

        # 15. 中低事件+极高紧急车辆+中等能见度
        if 31 <= special_event <= 50 and emergency_vehicles > 300 and 400 < visibility < 700:
            actions.append("中低级别事件，极多紧急车辆，中等能见度，强化应急管制")
            adjustments["main"] = self.max_main + 18
            adjustments["secondary"] = self.min_secondary
            adjustments["pedestrian"] = self.base_pedestrian + 12
            self.emergency_light = "on"
            self.traffic_police_alert = "on"

        # 16. 中低事件+中等紧急车辆+中等能见度
        if 31 <= special_event <= 50 and 100 <= emergency_vehicles < 250 and 400 < visibility < 700:
            actions.append("中低级别事件，中等紧急车辆，中等能见度，标准事件配时")
            adjustments["main"] = 40
            adjustments["secondary"] = 35
            adjustments["pedestrian"] = self.base_pedestrian + 10

        # 17. 中低事件+极高紧急车辆+极低能见度
        if 31 <= special_event <= 50 and emergency_vehicles > 320 and visibility < 200:
            actions.append("中低级别事件，极多紧急车辆，极低能见度，极度警戒")
            adjustments["main"] = 28
            adjustments["secondary"] = 23
            adjustments["pedestrian"] = self.base_pedestrian + 18
            self.emergency_light = "on"
            self.fog_light = "on"
            self.warning_light = "on"
            self.traffic_police_alert = "on"

        # 18. 中低事件+低紧急车辆+极高能见度
        if 31 <= special_event <= 50 and emergency_vehicles < 80 and visibility > 850:
            actions.append("中低级别事件，少量紧急车辆，能见度极佳，优化事件配时")
            adjustments["main"] = self.max_main - 5
            adjustments["secondary"] = self.max_secondary - 3
            adjustments["pedestrian"] = self.base_pedestrian + 12

        # 19. 中低事件+高紧急车辆+高能见度
        if 31 <= special_event <= 50 and 250 <= emergency_vehicles < 320 and visibility > 700:
            actions.append("中低级别事件，大量紧急车辆，能见度良好，高效应急")
            adjustments["main"] = self.max_main + 15
            adjustments["secondary"] = self.min_secondary + 3
            self.emergency_light = "on"
            self.traffic_police_alert = "on"

        # 20. 中低事件+中高紧急车辆+低能见度
        if 35 <= special_event <= 48 and 200 <= emergency_vehicles < 280 and visibility < 320:
            actions.append("中低级别事件，中高紧急车辆，低能见度，综合管制")
            adjustments["main"] = 35
            adjustments["secondary"] = 30
            adjustments["pedestrian"] = self.base_pedestrian + 14
            self.emergency_light = "on"
            self.fog_light = "on"

        # 21-30: 中高特殊事件值组合 (51-70)
        # 21. 中高事件+低紧急车辆+高能见度
        if 51 <= special_event <= 70 and emergency_vehicles < 100 and visibility > 700:
            actions.append("中高级别事件，少量紧急车辆，能见度良好，中型活动疏导")
            adjustments["main"] = self.max_main - 12
            adjustments["secondary"] = self.max_secondary - 10
            adjustments["pedestrian"] = self.base_pedestrian + 18

        # 22. 中高事件+低紧急车辆+低能见度
        if 51 <= special_event <= 70 and emergency_vehicles < 100 and visibility < 300:
            actions.append("中高级别事件，少量紧急车辆，低能见度，活动安全优先")
            adjustments["main"] = max(self.max_main - 25, 25)
            adjustments["secondary"] = max(self.max_secondary - 22, 22)
            adjustments["pedestrian"] = self.base_pedestrian + 22
            self.fog_light = "on"
            self.warning_light = "on"

        # 23. 中高事件+中等紧急车辆+高能见度
        if 51 <= special_event <= 70 and 100 <= emergency_vehicles < 250 and visibility > 650:
            actions.append("中高级别事件，中等紧急车辆，能见度良好，活动应急模式")
            adjustments["main"] = self.max_main + 5
            adjustments["secondary"] = self.max_secondary - 12
            adjustments["pedestrian"] = self.base_pedestrian + 18
            self.emergency_light = "on"

        # 24. 中高事件+高紧急车辆+低能见度
        if 51 <= special_event <= 70 and emergency_vehicles >= 250 and visibility < 350:
            actions.append("中高级别事件，大量紧急车辆，低能见度，活动紧急管制")
            adjustments["main"] = 30
            adjustments["secondary"] = 25
            adjustments["pedestrian"] = self.base_pedestrian + 20
            self.emergency_light = "on"
            self.fog_light = "on"
            self.traffic_police_alert = "on"

        # 25. 中高事件+极高紧急车辆+中等能见度
        if 51 <= special_event <= 70 and emergency_vehicles > 300 and 400 < visibility < 700:
            actions.append("中高级别事件，极多紧急车辆，中等能见度，活动强化应急")
            adjustments["main"] = self.max_main + 15
            adjustments["secondary"] = self.min_secondary + 3
            adjustments["pedestrian"] = self.base_pedestrian + 18
            self.emergency_light = "on"
            self.traffic_police_alert = "on"

        # 26. 中高事件+中等紧急车辆+中等能见度
        if 51 <= special_event <= 70 and 100 <= emergency_vehicles < 250 and 400 < visibility < 700:
            actions.append("中高级别事件，中等紧急车辆，中等能见度，活动标准配时")
            adjustments["main"] = 38
            adjustments["secondary"] = 33
            adjustments["pedestrian"] = self.base_pedestrian + 16

        # 27. 中高事件+极高紧急车辆+极低能见度
        if 51 <= special_event <= 70 and emergency_vehicles > 320 and visibility < 200:
            actions.append("中高级别事件，极多紧急车辆，极低能见度，活动最高警戒")
            adjustments["main"] = 26
            adjustments["secondary"] = 21
            adjustments["pedestrian"] = self.base_pedestrian + 25
            self.emergency_light = "on"
            self.fog_light = "on"
            self.warning_light = "on"
            self.traffic_police_alert = "on"

        # 28. 中高事件+低紧急车辆+极高能见度
        if 51 <= special_event <= 70 and emergency_vehicles < 80 and visibility > 850:
            actions.append("中高级别事件，少量紧急车辆，能见度极佳，活动最优配时")
            adjustments["main"] = self.max_main - 8
            adjustments["secondary"] = self.max_secondary - 6
            adjustments["pedestrian"] = self.base_pedestrian + 20

        # 29. 中高事件+高紧急车辆+高能见度
        if 51 <= special_event <= 70 and 250 <= emergency_vehicles < 320 and visibility > 700:
            actions.append("中高级别事件，大量紧急车辆，能见度良好，活动高效应急")
            adjustments["main"] = self.max_main + 12
            adjustments["secondary"] = self.min_secondary + 5
            adjustments["pedestrian"] = self.base_pedestrian + 16
            self.emergency_light = "on"
            self.traffic_police_alert = "on"

        # 30. 中高事件+中高紧急车辆+低中能见度
        if 55 <= special_event <= 68 and 180 <= emergency_vehicles < 280 and 280 < visibility < 450:
            actions.append("中高级别事件，中高紧急车辆，低中能见度，活动平衡管制")
            adjustments["main"] = 36
            adjustments["secondary"] = 31
            adjustments["pedestrian"] = self.base_pedestrian + 18
            self.emergency_light = "on"
            self.warning_light = "on"

        # 31-40: 高特殊事件值组合 (71-100)
        # 31. 高事件+低紧急车辆+高能见度
        if special_event > 70 and emergency_vehicles < 100 and visibility > 700:
            actions.append("高级别事件，少量紧急车辆，能见度良好，大型活动待命")
            adjustments["main"] = self.max_main
            adjustments["secondary"] = self.max_secondary - 5
            adjustments["pedestrian"] = self.base_pedestrian + 15
            self.warning_light = "on"

        # 32. 高事件+低紧急车辆+低能见度
        if special_event > 70 and emergency_vehicles < 100 and visibility < 300:
            actions.append("高级别事件，少量紧急车辆，低能见度，双重危机模式")
            adjustments["main"] = 25
            adjustments["secondary"] = 20
            adjustments["pedestrian"] = self.base_pedestrian + 20
            self.emergency_light = "on"
            self.fog_light = "on"
            self.traffic_police_alert = "on"

        # 33. 高事件+中等紧急车辆+高能见度
        if special_event > 70 and 100 <= emergency_vehicles < 250 and visibility > 650:
            actions.append("高级别事件，中等紧急车辆，能见度良好，高优先级通道")
            adjustments["main"] = self.max_main + 15
            adjustments["secondary"] = self.min_secondary
            adjustments["pedestrian"] = self.base_pedestrian + 12
            self.emergency_light = "on"
            self.traffic_police_alert = "on"

        # 34. 高事件+高紧急车辆+低能见度
        if special_event > 70 and emergency_vehicles >= 250 and visibility < 350:
            actions.append("高级别事件，大量紧急车辆，低能见度，极限危机管理")
            adjustments["main"] = 28
            adjustments["secondary"] = 22
            adjustments["pedestrian"] = self.base_pedestrian + 18
            self.emergency_light = "on"
            self.fog_light = "on"
            self.warning_light = "on"
            self.traffic_police_alert = "on"
            self.speed_limit_display = "very_reduced"

        # 35. 高事件+极高紧急车辆+中等能见度
        if special_event > 70 and emergency_vehicles > 300 and 400 < visibility < 700:
            actions.append("高级别事件，极多紧急车辆，中等能见度，紧急强化管制")
            adjustments["main"] = self.max_main + 22
            adjustments["secondary"] = self.min_secondary - 5
            adjustments["pedestrian"] = self.base_pedestrian + 15
            self.emergency_light = "on"
            self.traffic_police_alert = "on"

        # 36. 高事件+超高紧急车辆+任意能见度
        if special_event > 70 and emergency_vehicles > 350:
            actions.append("高级别事件，超多紧急车辆，最高级别应急响应")
            adjustments["main"] = self.max_main + 30
            adjustments["secondary"] = self.min_secondary - 10
            adjustments["pedestrian"] = self.base_pedestrian + 10
            self.emergency_light = "on"
            self.traffic_police_alert = "on"
            self.warning_light = "on"

        # 37. 高事件+中等紧急车辆+中等能见度
        if special_event > 70 and 100 <= emergency_vehicles < 250 and 400 < visibility < 700:
            actions.append("高级别事件，中等紧急车辆，中等能见度，紧急准备模式")
            adjustments["main"] = self.max_main + 8
            adjustments["secondary"] = self.max_secondary - 8
            adjustments["pedestrian"] = self.base_pedestrian + 12
            self.warning_light = "on"
            self.emergency_light = "on"

        # 38. 高事件+低紧急车辆+极低能见度
        if special_event > 70 and emergency_vehicles < 100 and visibility < 200:
            actions.append("高级别事件，少量紧急车辆，极低能见度，三重危机")
            adjustments["main"] = 22
            adjustments["secondary"] = 18
            adjustments["pedestrian"] = self.base_pedestrian + 22
            self.emergency_light = "on"
            self.fog_light = "on"
            self.warning_light = "on"
            self.traffic_police_alert = "on"
            self.speed_limit_display = "very_reduced"

        # 39. 极高事件+高紧急车辆+高能见度
        if special_event > 85 and emergency_vehicles >= 250 and visibility > 700:
            actions.append("极高级别事件，大量紧急车辆，能见度良好，最高响应")
            adjustments["main"] = self.max_main + 28
            adjustments["secondary"] = self.min_secondary - 8
            self.emergency_light = "on"
            self.traffic_police_alert = "on"
            self.warning_light = "on"

        # 40. 极高事件+极高紧急车辆+低能见度
        if special_event > 85 and emergency_vehicles > 300 and visibility < 350:
            actions.append("极高级别事件，极多紧急车辆，低能见度，终极危机管理")
            adjustments["main"] = 25
            adjustments["secondary"] = 20
            self.emergency_light = "on"
            self.fog_light = "on"
            self.warning_light = "on"
            self.traffic_police_alert = "on"
            self.speed_limit_display = "very_reduced"

        # 41-50: 跨范围复合条件
        # 41. 事件与紧急车辆成正比+高能见度
        if special_event // 10 == emergency_vehicles // 40 and visibility > 700:
            actions.append("事件级别与紧急车辆匹配，能见度良好，协调响应")
            adjustments["main"] = self.max_main + int(special_event / 20)
            adjustments["secondary"] = self.max_secondary - int(emergency_vehicles / 100)

        # 42. 事件值高+紧急车辆少+能见度差(矛盾情况)
        if special_event > 60 and emergency_vehicles < 120 and visibility < 400:
            actions.append("高级别事件但紧急车辆不足且能见度差，资源调配警告")
            adjustments["main"] = 35
            adjustments["secondary"] = 30
            self.warning_light = "on"
            self.traffic_police_alert = "on"

        # 43. 三个变量都在中等范围
        if 40 < special_event < 60 and 150 < emergency_vehicles < 250 and 400 < visibility < 600:
            actions.append("三参数均处中等水平，标准均衡配时")
            adjustments["main"] = 42
            adjustments["secondary"] = 38
            adjustments["pedestrian"] = self.base_pedestrian + 12

        # 44. 事件值与能见度反比关系
        if special_event > 70 and visibility < (1000 - special_event * 8):
            actions.append("高级别事件伴随低能见度，恶劣条件组合")
            adjustments["main"] = max(self.max_main - 18, 25)
            adjustments["secondary"] = max(self.max_secondary - 15, 22)
            self.fog_light = "on"
            self.emergency_light = "on"

        # 45. 紧急车辆数量超过事件级别要求
        if emergency_vehicles > special_event * 3 and visibility > 500:
            actions.append("紧急车辆超配，可能是重大事故，全力支援")
            adjustments["main"] = self.max_main + 20
            adjustments["secondary"] = self.min_secondary
            self.emergency_light = "on"
            self.traffic_police_alert = "on"

        # 46. 三个变量都在低值范围
        if special_event < 30 and emergency_vehicles < 100 and visibility < 350:
            actions.append("三参数均偏低，低级别低能见度模式")
            adjustments["main"] = 30
            adjustments["secondary"] = 26
            self.fog_light = "on"

        # 47. 三个变量都在高值范围
        if special_event > 75 and emergency_vehicles > 280 and visibility > 750:
            actions.append("三参数均偏高，大型事件高响应高能见度")
            adjustments["main"] = self.max_main + 25
            adjustments["secondary"] = self.min_secondary - 5
            self.emergency_light = "on"
            self.traffic_police_alert = "on"

        # 48. 紧急车辆与能见度负相关
        if emergency_vehicles > 250 and visibility < (1200 - emergency_vehicles):
            actions.append("大量紧急车辆伴随低能见度，复杂应急环境")
            adjustments["main"] = 32
            adjustments["secondary"] = 28
            self.emergency_light = "on"
            self.fog_light = "on"
            self.traffic_police_alert = "on"

        # 49. 事件级别模运算特殊值
        if special_event % 17 == 0 and emergency_vehicles > 200:
            actions.append("特殊事件编码匹配，大量紧急车辆，预案模式")
            adjustments["main"] = self.max_main + 12
            adjustments["secondary"] = self.max_secondary - 10
            self.emergency_light = "on"

        # 50. 综合评分阈值判断
        if (special_event * 0.4 + emergency_vehicles * 0.1 + (1000 - visibility) * 0.05) > 80:
            actions.append("综合危机指数超标，启动最高级别管控")
            adjustments["main"] = max(self.max_main - 15, 28)
            adjustments["secondary"] = max(self.max_secondary - 12, 25)
            self.emergency_light = "on"
            self.fog_light = "on"
            self.warning_light = "on"
            self.traffic_police_alert = "on"

        # === Additional conditions extracted from first code ===

        # Single variable conditions
        if special_event < 50:
            actions.append("特殊事件级别较低，基础模式")
            adjustments["main"] = self.max_main - 3

        if special_event > 50:
            actions.append("特殊事件级别较高，提升响应")
            adjustments["main"] = self.max_main + 5

        if special_event <= 25:
            actions.append("极低级别事件，最小影响")
            adjustments["secondary"] = self.max_secondary - 2

        if special_event >= 25:
            actions.append("事件级别25以上，标准处理")
            adjustments["pedestrian"] = self.base_pedestrian + 5

        if emergency_vehicles < 200:
            actions.append("紧急车辆数量较少，轻度调整")
            adjustments["main"] = self.max_main - 2

        if emergency_vehicles > 200:
            actions.append("紧急车辆数量较多，增强响应")
            adjustments["main"] = self.max_main + 8

        if emergency_vehicles <= 150:
            actions.append("紧急车辆很少，常规配时")
            adjustments["secondary"] = self.max_secondary

        if emergency_vehicles >= 150:
            actions.append("紧急车辆达到阈值，开始应急")
            self.emergency_light = "on"

        if emergency_vehicles % 2 == 0:
            actions.append("紧急车辆偶数，特殊配时")
            adjustments["main"] = adjustments.get("main", 40) + 2

        if visibility < 500:
            actions.append("能见度较低，安全优先")
            adjustments["main"] = max(adjustments.get("main", 40) - 8, 25)

        if visibility > 500:
            actions.append("能见度良好，优化通行")
            adjustments["main"] = adjustments.get("main", 40) + 3

        if visibility <= 400:
            actions.append("能见度差，启动雾灯")
            self.fog_light = "on"

        if visibility >= 400:
            actions.append("能见度可接受，标准操作")
            adjustments["secondary"] = self.max_secondary - 1

        if visibility % 10 < 5:
            actions.append("能见度特殊值，精细调节")
            adjustments["main"] = adjustments.get("main", 40) + 1

        # Two variable combinations not already covered
        if special_event <= 40 or emergency_vehicles <= 180:
            actions.append("低事件或低紧急车辆，宽松配时")
            adjustments["main"] = self.max_main - 4

        if emergency_vehicles <= 350 or visibility <= 700:
            actions.append("中等紧急车辆或中等能见度")
            adjustments["secondary"] = self.max_secondary - 3

        if emergency_vehicles <= 250 or visibility <= 700:
            actions.append("适中条件组合")
            adjustments["main"] = 42

        if special_event < 40 or emergency_vehicles > 300:
            actions.append("低事件或高紧急车辆")
            adjustments["main"] = 45

        if special_event > 60 or visibility < 400:
            actions.append("高事件或低能见度")
            adjustments["main"] = 38
            self.warning_light = "on"

        if emergency_vehicles < 150 or visibility > 700:
            actions.append("少紧急车辆或高能见度")
            adjustments["secondary"] = self.max_secondary + 2

        # Mathematical operations
        if special_event + emergency_vehicles < 350:
            actions.append("事件和紧急车辆总和较低")
            adjustments["main"] = self.max_main - 6

        if special_event * 2 < emergency_vehicles:
            actions.append("紧急车辆数量超过事件的两倍")
            adjustments["main"] = self.max_main + 12
            self.emergency_light = "on"

        if visibility - special_event > 400:
            actions.append("能见度远超事件级别")
            adjustments["main"] = self.max_main + 3

        if emergency_vehicles % 50 == 0:
            actions.append("紧急车辆为50的倍数")
            adjustments["main"] = adjustments.get("main", 40) + 3

        if special_event + visibility > 600:
            actions.append("事件加能见度超过600")
            adjustments["secondary"] = self.max_secondary - 5

        if abs(special_event - 50) < 20:
            actions.append("事件级别接近50")
            adjustments["main"] = 43

        # Range conditions
        if 20 <= special_event <= 80 and emergency_vehicles < 250:
            actions.append("中等事件范围，低紧急车辆")
            adjustments["main"] = 41

        if 100 <= emergency_vehicles <= 300 and visibility > 400:
            actions.append("标准紧急车辆范围，好能见度")
            adjustments["secondary"] = self.max_secondary - 4

        if 300 <= visibility <= 700 and special_event > 25:
            actions.append("中等能见度范围，有事件")
            adjustments["main"] = 40

        # Complex range conditions
        if (10 <= special_event <= 40) and (150 <= emergency_vehicles <= 350):
            actions.append("双参数中等范围")
            adjustments["main"] = 42
            adjustments["secondary"] = 37

        if (special_event > 60) and (200 < visibility < 800):
            actions.append("高事件，中高能见度")
            adjustments["main"] = self.max_main + 6

        # Complex mathematical conditions
        if special_event + emergency_vehicles // 2 > 150:
            actions.append("事件加半数紧急车辆超150")
            adjustments["main"] = self.max_main + 4

        if visibility // 10 + special_event > 80:
            actions.append("能见度十分位加事件超80")
            adjustments["secondary"] = self.max_secondary - 6

        if emergency_vehicles % 10 + special_event % 10 < 15:
            actions.append("两参数个位数和小于15")
            adjustments["main"] = adjustments.get("main", 40) + 2

        # Min/max functions
        if max(special_event, emergency_vehicles // 4) > 45:
            actions.append("事件或紧急车辆四分位超45")
            adjustments["main"] = self.max_main + 3

        if min(visibility // 10, special_event) < 35:
            actions.append("能见度十分位或事件小于35")
            adjustments["main"] = max(adjustments.get("main", 40) - 5, 25)

        if special_event > visibility // 20:
            actions.append("事件超过能见度的二十分位")
            adjustments["secondary"] = self.max_secondary - 2

        # Modulo and bitwise operations
        if special_event % 2 == emergency_vehicles % 2:
            actions.append("事件和紧急车辆奇偶性相同")
            adjustments["main"] = adjustments.get("main", 40) + 1

        if visibility % 7 == 0:
            actions.append("能见度为7的倍数")
            adjustments["secondary"] = adjustments.get("secondary", 35) + 2

        if (special_event + emergency_vehicles) % 5 == 0:
            actions.append("事件和紧急车辆和为5的倍数")
            adjustments["main"] = adjustments.get("main", 40) + 2

        # Bitwise operations
        if special_event & 1 == visibility & 1:
            actions.append("事件和能见度最低位相同")
            adjustments["main"] = adjustments.get("main", 40) + 1

        if emergency_vehicles >> 2 > special_event:
            actions.append("紧急车辆右移2位大于事件")
            adjustments["main"] = self.max_main + 5

        if (special_event ^ emergency_vehicles) % 50 < 25:
            actions.append("事件异或紧急车辆模50小于25")
            adjustments["secondary"] = adjustments.get("secondary", 35) + 1

        # String and complex operations
        if len(str(special_event)) + len(str(emergency_vehicles)) > 4:
            actions.append("事件和紧急车辆位数和超过4")
            adjustments["main"] = adjustments.get("main", 40) + 2

        if sum(int(d) for d in str(visibility)) > 10:
            actions.append("能见度各位数字和超过10")
            adjustments["secondary"] = adjustments.get("secondary", 35) + 1

        if special_event ** 2 % 100 < 50:
            actions.append("事件平方模100小于50")
            adjustments["main"] = adjustments.get("main", 40) + 1

        # Final complex conditions
        if abs(special_event - emergency_vehicles // 4) < 20:
            actions.append("事件与紧急车辆四分位差小于20")
            adjustments["secondary"] = adjustments.get("secondary", 35) + 2

        if (visibility + special_event + emergency_vehicles) % 100 > 50:
            actions.append("三参数和模100大于50")
            adjustments["main"] = adjustments.get("main", 40) + 3

        return actions, adjustments

    def get_weather_name(self, weather_code):
        """将天气代码转换为名称"""
        weather_map = {1: "sunny", 2: "rainy", 3: "foggy", 4: "snowy", 5: "windy", 6: "stormy"}
        return weather_map.get(weather_code, "unknown")

    def get_time_name(self, time_code):
        """将时间代码转换为名称"""
        time_map = {1: "rush_morning", 2: "rush_evening", 3: "lunch", 4: "night", 5: "weekend", 6: "holiday"}
        return time_map.get(time_code, "unknown")

    def control_traffic(self):
        """主控制函数（四大块结构大规模扩展版）"""
        # 获取传感器数据
        (x, y, z, weather, time_period, special_event, vehicle_types,
         emergency_vehicles, air_quality, noise_level, visibility,
         temperature, wind_speed, school_zones_active, business_district_active,
         hospital_nearby, construction_zone, accident_probability, road_condition) = self.get_sensor_data()

        print(f"当前交通数据 - 主路: {x:.1f}%, 辅路: {y:.1f}%, 行人: {z:.0f}")
        print(
            f"天气: {self.get_weather_name(weather)}({weather}), 时段: {self.get_time_name(time_period)}({time_period}), 特殊事件: {special_event}")
        print(f"车辆类型 - 轿车: {vehicle_types['cars']:.1f}%, 货车: {vehicle_types['trucks']:.1f}%")
        print(
            f"        公交: {vehicle_types['buses']:.1f}%, 摩托: {vehicle_types['motorcycles']:.1f}%, 自行车: {vehicle_types['bicycles']:.1f}%")
        print(f"紧急车辆: {emergency_vehicles}, 空气质量: {air_quality:.0f}, 能见度: {visibility:.0f}m")
        print(f"温度: {temperature:.1f}°C, 风速: {wind_speed:.1f}m/s, 噪音: {noise_level:.1f}dB")
        print(f"学校活跃: {school_zones_active}, 商业区活跃: {business_district_active}")
        print(f"医院附近: {hospital_nearby}, 施工区域: {construction_zone}")
        print(f"事故概率: {accident_probability:.1f}%, 路况: {road_condition}")

        # 初始化参数
        actions = []
        adjustments = {}
        blocks_executed = []

        # 重置所有警示状态
        self.warning_light = "off"
        self.traffic_police_alert = "off"
        self.emergency_light = "off"
        self.fog_light = "off"
        self.speed_limit_display = "normal"
        self.noise_alert = "off"
        self.air_quality_alert = "off"

        # 参数传递元组
        params = (x, y, z, weather, time_period, special_event, vehicle_types,
                  emergency_vehicles, air_quality, noise_level, visibility,
                  temperature, wind_speed, school_zones_active, business_district_active,
                  hospital_nearby, construction_zone, accident_probability, road_condition)

        # 按四大块执行策略
        # 大块1: X, Y, Z组合（基础流量关系）- 45个IF语句
        blocks_executed.append(1)
        acts, adj = self.block1_x_y_z_combinations(*params)
        actions.extend(acts)
        adjustments.update(adj)

        # 大块2: weather(数字), X, Y组合（天气与车流）- 40个IF语句
        blocks_executed.append(2)
        acts, adj = self.block2_weather_x_y_combinations(*params)
        actions.extend(acts)
        adjustments.update(adj)

        # 大块3: weather(数字), time_period(数字), Z组合（天气时间与行人）- 35个IF语句
        blocks_executed.append(3)
        acts, adj = self.block3_weather_time_z_combinations(*params)
        actions.extend(acts)
        adjustments.update(adj)

        # 大块4: time_period(数字), X, Y组合（时间与车流）- 35个IF语句
        blocks_executed.append(4)
        acts, adj = self.block4_time_x_y_combinations(*params)
        actions.extend(acts)
        adjustments.update(adj)

        # 设置默认值
        if "main" not in adjustments:
            adjustments["main"] = 40
        if "secondary" not in adjustments:
            adjustments["secondary"] = 35
        if "pedestrian" not in adjustments:
            adjustments["pedestrian"] = self.base_pedestrian
        if "turning" not in adjustments:
            adjustments["turning"] = self.base_turning
        if "bicycle" not in adjustments:
            adjustments["bicycle"] = self.base_bicycle
        if "bus" not in adjustments:
            adjustments["bus"] = self.base_bus

        # 更新信号灯状态
        self.main_road_light = "green"
        self.secondary_road_light = "red"

        print(f"\n执行大块 {blocks_executed} 的控制策略")
        print(f"总IF语句数量: 155个 (大块1: 45个, 大块2: 40个, 大块3: 35个, 大块4: 35个)")
        for action in actions:
            print(f"- {action}")

        print("\n信号灯配置:")
        print(f"主路绿灯时长: {adjustments['main']}秒")
        print(f"辅路绿灯时长: {adjustments['secondary']}秒")
        print(f"行人绿灯时长: {adjustments['pedestrian']}秒")
        print(f"转弯绿灯时长: {adjustments['turning']}秒")
        print(f"自行车绿灯时长: {adjustments['bicycle']}秒")
        print(f"公交专用道时长: {adjustments['bus']}秒")

        print("\n系统状态:")
        print(f"警示灯状态: {self.warning_light}")
        print(f"交警支援请求: {self.traffic_police_alert}")
        print(f"应急灯状态: {self.emergency_light}")
        print(f"雾灯状态: {self.fog_light}")
        print(f"限速显示: {self.speed_limit_display}")
        print(f"噪音警示: {self.noise_alert}")
        print(f"空气质量警示: {self.air_quality_alert}")
        print(f"{'-' * 100}")


# 运行大规模扩展版四大块结构的交通控制系统
if __name__ == "__main__":
    controller = AdvancedTrafficControllerMassiveExtended()
    # 模拟3个周期的控制
    for i in range(3):
        print(f"=== 控制周期 {i + 1} ===")
        controller.control_traffic()
        print()