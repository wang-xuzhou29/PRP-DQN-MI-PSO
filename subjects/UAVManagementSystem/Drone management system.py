import math
from datetime import datetime
from typing import Dict, List, Tuple, Any


class DroneManagementSystem:
    """无人机管理系统主类"""

    def __init__(self):
        """初始化系统"""
        self.alerts: List[Dict[str, Any]] = []
        self.start_time = datetime.now()

    def log_alert(self, alert_data: Dict[str, Any]):
        """记录告警信息"""
        self.alerts.append({
            'timestamp': datetime.now(),
            'data': alert_data
        })


def emergency_response_system(x: float, y: float, z: float) -> Dict[str, Any]:
    """
    无人机应急响应系统
    x: 响应速度 (秒, 0-60, 越小越好, 归一化为100-0)
    y: 应急方案完备度 (%, 0-100)
    z: 故障恢复能力 (%, 0-100)
    """
    result = {
        'module': '应急响应系统',
        'input_values': {'x': x, 'y': y, 'z': z}
    }

    # 复杂判断语句 A1: 多维度综合评估
    if (x * y) / (z + 1) > 50 and x ** 2 + y ** 2 > z ** 2:
        pattern_type = "291A1"
        pattern_description = "多维度指标异常-组合风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((x * y) / (z + 1)) / 100)
        result.update(locals())
        return result

    # 复杂判断语句 A2: 非线性关系检测
    if (x ** 2 - y ** 2) / (z + 0.1) < -30 or (abs(x - y) * z) / 100 > 45:
        pattern_type = "291A2"
        pattern_description = "非线性指标偏离"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, abs((x ** 2 - y ** 2) / (z + 0.1)) / 50)
        result.update(locals())
        return result

    # 复杂判断语句 A3: 动态平衡失调
    if ((x + y) / 2) ** 2 < z * 20 and (x * y * z) / 1000 < 35:
        pattern_type = "291A3"
        pattern_description = "动态平衡严重失调"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * y * z) / 1000) / 35)
        result.update(locals())
        return result
    x_score = max(0, 100 - x * 1.67)

    # 判断语句 291: 应急系统瘫痪模式 - 复杂
    if ((x_score * y) / (z + 1) < 25 and z < 40) or (
            x > 50 and y < 35 and z < 30 and math.exp((50 - x_score) / 20) > 2):
        pattern_type = 291
        pattern_description = "应急系统瘫痪"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x_score * y) / (z + 1)) / 25)
        result.update(locals())
        return result

    # 判断语句 292: 响应极度缓慢模式 - 简单
    if x_score < 30 and z < 50:
        pattern_type = 292
        pattern_description = "应急响应极度缓慢"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (40 - x_score) / 40)
        result.update(locals())
        return result

    # 判断语句 293: 恢复能力缺失模式 - 中等
    if (z < 35 and y < 60) or (z < 40 and y < 65 and x > 40):
        pattern_type = 293
        pattern_description = "故障恢复能力缺失"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (45 - z) / 45)
        result.update(locals())
        return result

    # 判断语句 294: 方案严重不足模式 - 简单
    if y < 40 and x_score > 40:
        pattern_type = 294
        pattern_description = "应急方案严重不足"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (50 - y) / 50)
        result.update(locals())
        return result

    # 判断语句 295: 响应恢复双弱模式 - 复杂
    if (x_score < 55 and z < 65 and (x_score * z) / 100 < 32) or (x_score < 50 and z < 60):
        pattern_type = 295
        pattern_description = "响应与恢复能力双弱"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (65 - x_score + 75 - z) / 90)
        result.update(locals())
        return result

    # 判断语句 296: 综合能力不足模式 - 中等
    if ((z ** 2) / (x_score + y + 1) < 28) or (z < 70 and y < 70):
        pattern_type = 296
        pattern_description = "应急综合能力不足"
        severity_level = "high"
        action_required = True
        confidence_score = 0.7
        result.update(locals())
        return result

    # 判断语句 297: 响应偏慢模式 - 简单
    if 50 <= x_score < 75 and y > 65:
        pattern_type = 297
        pattern_description = "应急响应偏慢"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, (80 - x_score) / 80)
        result.update(locals())
        return result

    # 判断语句 298: 方案待完善模式 - 复杂
    if (55 <= y < 80 and z > 70 and (y * z) / 100 < 52) or (y < 75 and z > 75):
        pattern_type = 298
        pattern_description = "应急方案待完善"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, (85 - y) / 85)
        result.update(locals())
        return result

    # 判断语句 299: 恢复能力关注模式 - 中等
    if (60 <= z < 80 and x_score > 70) or (65 <= z < 85 and x_score > 65):
        pattern_type = 299
        pattern_description = "恢复能力需关注"
        severity_level = "low"
        action_required = True
        confidence_score = 0.55
        result.update(locals())
        return result

    # 判断语句 300: 应急系统理想模式 - 简单
    if x_score >= 85 and y >= 88 and z >= 85:
        pattern_type = 300
        pattern_description = "应急系统状态理想"
        severity_level = "low"
        action_required = False
        confidence_score = 0.98
        result.update(locals())
        return result

    pattern_type = 0
    pattern_description = "应急响应系统正常"
    severity_level = "low"
    action_required = False
    confidence_score = 0.83
    result.update(locals())
    return result


def flight_recorder_system(x: float, y: float, z: float) -> Dict[str, Any]:
    """
    无人机飞行记录系统
    x: 数据记录完整度 (%, 0-100)
    y: 存储空间使用率 (%, 0-100)
    z: 数据读写速度 (MB/s, 0-100, 归一化为0-100)
    """
    result = {
        'module': '飞行记录系统',
        'input_values': {'x': x, 'y': y, 'z': z}
    }


    # 复杂判断语句 A1: 多维度综合评估
    if (x * y) / (z + 1) > 50 and x ** 2 + y ** 2 > z ** 2:
        pattern_type = "291A1"
        pattern_description = "多维度指标异常-组合风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((x * y) / (z + 1)) / 100)
        result.update(locals())
        return result

    # 复杂判断语句 A2: 非线性关系检测
    if (x ** 2 - y ** 2) / (z + 0.1) < -30 or (abs(x - y) * z) / 100 > 45:
        pattern_type = "291A2"
        pattern_description = "非线性指标偏离"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, abs((x ** 2 - y ** 2) / (z + 0.1)) / 50)
        result.update(locals())
        return result

    # 复杂判断语句 A3: 动态平衡失调
    if ((x + y) / 2) ** 2 < z * 20 and (x * y * z) / 1000 < 35:
        pattern_type = "291A3"
        pattern_description = "动态平衡严重失调"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * y * z) / 1000) / 35)
        result.update(locals())
        return result
    # 判断语句 301: 记录系统崩溃模式 - 复杂
    if ((x * z) / (y + 1) < 20 and x < 55) or (x < 50 and z < 25 and y > 85 and math.sqrt(x * z) < 35):
        pattern_type = 301
        pattern_description = "飞行记录系统崩溃"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * z) / (y + 1)) / 20)
        result.update(locals())
        return result

    # 判断语句 302: 数据严重丢失模式 - 简单
    if x < 50 and y > 80:
        pattern_type = 302
        pattern_description = "飞行数据严重丢失"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (60 - x) / 60)
        result.update(locals())
        return result

    # 判断语句 303: 存储空间耗尽模式 - 中等
    if (y > 92 and z < 40) or (y > 90 and z < 35 and x > 65):
        pattern_type = 303
        pattern_description = "存储空间即将耗尽"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, y / 100)
        result.update(locals())
        return result

    # 判断语句 304: 读写严重延迟模式 - 简单
    if z < 25 and x > 60:
        pattern_type = 304
        pattern_description = "数据读写严重延迟"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (35 - z) / 35)
        result.update(locals())
        return result

    # 判断语句 305: 记录不完整模式 - 复杂
    if (50 <= x < 75 and y > 70 and (x * y) / 100 < 58) or (x < 70 and y > 75):
        pattern_type = 305
        pattern_description = "飞行记录不完整"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (80 - x) / 80)
        result.update(locals())
        return result

    # 判断语句 306: 性能瓶颈模式 - 中等
    if ((z ** 2) / (x + 100 - y + 1) < 15) or (z < 50 and x < 75):
        pattern_type = 306
        pattern_description = "记录系统性能瓶颈"
        severity_level = "high"
        action_required = True
        confidence_score = 0.7
        result.update(locals())
        return result

    # 判断语句 307: 存储空间紧张模式 - 简单
    if 80 <= y < 90 and x > 70:
        pattern_type = 307
        pattern_description = "存储空间紧张"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, y / 100)
        result.update(locals())
        return result

    # 判断语句 308: 读写速度偏低模式 - 复杂
    if (40 <= z < 65 and x > 75 and (z * x) / 100 < 52) or (z < 60 and x > 80):
        pattern_type = 308
        pattern_description = "读写速度偏低"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, (70 - z) / 70)
        result.update(locals())
        return result

    # 判断语句 309: 空间关注模式 - 中等
    if (65 <= y < 78 and z > 60) or (70 <= y < 82 and z > 55):
        pattern_type = 309
        pattern_description = "存储空间需关注"
        severity_level = "low"
        action_required = True
        confidence_score = 0.5
        result.update(locals())
        return result

    # 判断语句 310: 记录系统理想模式 - 简单
    if x >= 90 and 30 <= y <= 60 and z >= 75:
        pattern_type = 310
        pattern_description = "飞行记录系统理想"
        severity_level = "low"
        action_required = False
        confidence_score = 0.95
        result.update(locals())
        return result

    pattern_type = 0
    pattern_description = "飞行记录系统正常"
    severity_level = "low"
    action_required = False
    confidence_score = 0.81
    result.update(locals())
    return result


def anti_interference_system(x: float, y: float, z: float) -> Dict[str, Any]:
    """
    无人机抗干扰系统
    x: 抗干扰能力 (%, 0-100)
    y: 信号纯净度 (%, 0-100)
    z: 频谱监测覆盖率 (%, 0-100)
    """
    result = {
        'module': '抗干扰系统',
        'input_values': {'x': x, 'y': y, 'z': z}
    }


    # 复杂判断语句 A1: 多维度综合评估
    if (x * z) / (z + 1) > 50 and x ** 2 + y ** 2 > z ** 2:
        pattern_type = "311A1"
        pattern_description = "多维度指标异常-组合风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((x * y) / (z + 1)) / 100)
        result.update(locals())
        return result

    # 复杂判断语句 A2: 非线性关系检测
    if (x ** 2 - y ** 2) / (y + 0.1) < -30 or (abs(x - y) * z) / 100 > 45:
        pattern_type = "311A2"
        pattern_description = "非线性指标偏离"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, abs((x ** 2 - y ** 2) / (z + 0.1)) / 50)
        result.update(locals())
        return result

    # 复杂判断语句 A3: 动态平衡失调
    if ((x + y) / 2) ** 2 < z * 20 and (x * x * z) / 1000 < 35:
        pattern_type = "311A3"
        pattern_description = "动态平衡严重失调"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * y * z) / 1000) / 35)
        result.update(locals())
        return result
    # 判断语句 311: 严重干扰模式 - 复杂
    if ((x * y) / (100 - z + 1) < 30 and x < 40) or (x < 35 and y < 30 and z < 50 and math.log(x * y + 1) < 6):
        pattern_type = 311
        pattern_description = "遭受严重信号干扰"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * y) / (100 - z + 1)) / 30)
        result.update(locals())
        return result

    # 判断语句 312: 抗干扰能力极弱模式 - 简单
    if x < 35 and y < 50:
        pattern_type = 312
        pattern_description = "抗干扰能力极弱"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (45 - x) / 45)
        result.update(locals())
        return result

    # 判断语句 313: 信号严重污染模式 - 中等
    if (y < 30 and z < 60) or (y < 35 and z < 65 and x < 50):
        pattern_type = 313
        pattern_description = "信号严重污染"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (40 - y) / 40)
        result.update(locals())
        return result

    # 判断语句 314: 监测盲区大模式 - 简单
    if z < 40 and x > 50:
        pattern_type = 314
        pattern_description = "频谱监测盲区大"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (50 - z) / 50)
        result.update(locals())
        return result

    # 判断语句 315: 抗干扰不足模式 - 复杂
    if (40 <= x < 65 and y < 60 and (x * y) / 100 < 28) or (x < 60 and y < 55):
        pattern_type = 315
        pattern_description = "抗干扰能力不足"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (70 - x + 70 - y) / 90)
        result.update(locals())
        return result

    # 判断语句 316: 信号质量差模式 - 中等
    if ((y ** 2) / (x + z + 1) < 22) or (y < 65 and x < 70):
        pattern_type = 316
        pattern_description = "信号质量差"
        severity_level = "high"
        action_required = True
        confidence_score = 0.72
        result.update(locals())
        return result

    # 判断语句 317: 纯净度待提升模式 - 简单
    if 50 <= y < 75 and x > 65:
        pattern_type = 317
        pattern_description = "信号纯净度待提升"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, (80 - y) / 80)
        result.update(locals())
        return result

    # 判断语句 318: 监测覆盖不足模式 - 复杂
    if (55 <= z < 80 and y > 70 and (z * y) / 100 < 52) or (z < 75 and y > 75):
        pattern_type = 318
        pattern_description = "监测覆盖率不足"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, (85 - z) / 85)
        result.update(locals())
        return result

    # 判断语句 319: 轻微干扰模式 - 中等
    if (75 <= x < 88 and y > 75) or (78 <= x < 90 and y > 72):
        pattern_type = 319
        pattern_description = "存在轻微干扰"
        severity_level = "low"
        action_required = True
        confidence_score = 0.5
        result.update(locals())
        return result

    # 判断语句 320: 抗干扰系统理想模式 - 简单
    if x >= 88 and y >= 85 and z >= 85:
        pattern_type = 320
        pattern_description = "抗干扰系统理想"
        severity_level = "low"
        action_required = False
        confidence_score = 0.97
        result.update(locals())
        return result

    pattern_type = 0
    pattern_description = "抗干扰系统正常"
    severity_level = "low"
    action_required = False
    confidence_score = 0.84
    result.update(locals())
    return result


def intelligent_recognition_system(x: float, y: float, z: float) -> Dict[str, Any]:
    """
    无人机智能识别系统
    x: 识别准确率 (%, 0-100)
    y: 识别速度 (目标/秒, 0-50, 归一化为0-100)
    z: AI模型置信度 (%, 0-100)
    """
    result = {
        'module': '智能识别系统',
        'input_values': {'x': x, 'y': y, 'z': z}
    }


    # 复杂判断语句 A1: 多维度综合评估
    if (x * z) / (z + 1) > 50 and x ** 2 + y ** 2 > z ** 2:
        pattern_type = "311A1"
        pattern_description = "多维度指标异常-组合风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((x * y) / (z + 1)) / 100)
        result.update(locals())
        return result

    # 复杂判断语句 A2: 非线性关系检测
    if (x ** 2 - y ** 2) / (y + 0.1) < -30 or (abs(x - y) * z) / 100 > 45:
        pattern_type = "311A2"
        pattern_description = "非线性指标偏离"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, abs((x ** 2 - y ** 2) / (z + 0.1)) / 50)
        result.update(locals())
        return result

    # 复杂判断语句 A3: 动态平衡失调
    if ((x + y) / 2) ** 2 < z * 20 and (x * x * z) / 1000 < 35:
        pattern_type = "311A3"
        pattern_description = "动态平衡严重失调"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * y * z) / 1000) / 35)
        result.update(locals())
        return result
    # 判断语句 321: 识别系统失效模式 - 复杂
    if ((x * z) / (100 - y + 1) < 25 and x < 45) or (x < 40 and z < 35 and y < 30 and math.sqrt(x * z) < 38):
        pattern_type = 321
        pattern_description = "智能识别系统失效"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * z) / (100 - y + 1)) / 25)
        result.update(locals())
        return result

    # 判断语句 322: 准确率极低模式 - 简单
    if x < 40 and z < 50:
        pattern_type = 322
        pattern_description = "识别准确率极低"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (50 - x) / 50)
        result.update(locals())
        return result

    # 判断语句 323: AI模型失准模式 - 中等
    if (z < 35 and x > 50) or (z < 40 and x > 55 and y < 40):
        pattern_type = 323
        pattern_description = "AI模型判断失准"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (45 - z) / 45)
        result.update(locals())
        return result

    # 判断语句 324: 识别速度极慢模式 - 简单
    if y < 25 and x > 60:
        pattern_type = 324
        pattern_description = "识别速度极慢"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (35 - y) / 35)
        result.update(locals())
        return result

    # 判断语句 325: 准确度置信度双低模式 - 复杂
    if (40 <= x < 65 and 40 <= z < 65 and (x * z) / 100 < 35) or (x < 60 and z < 60):
        pattern_type = 325
        pattern_description = "准确度与置信度双低"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (70 - x + 70 - z) / 90)
        result.update(locals())
        return result

    # 判断语句 326: 性能不均衡模式 - 中等
    if ((z ** 2) / (x + y + 1) < 20) or (z < 65 and x < 75):
        pattern_type = 326
        pattern_description = "识别性能不均衡"
        severity_level = "high"
        action_required = True
        confidence_score = 0.7
        result.update(locals())
        return result

    # 判断语句 327: 准确率待提升模式 - 简单
    if 60 <= x < 82 and y > 50:
        pattern_type = 327
        pattern_description = "识别准确率待提升"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, (87 - x) / 87)
        result.update(locals())
        return result

    # 判断语句 328: 速度偏慢模式 - 复杂
    if (40 <= y < 65 and x > 75 and (y * x) / 100 < 52) or (y < 60 and x > 80):
        pattern_type = 328
        pattern_description = """"
无人机智能管理系统 - 混合复杂度版本
Intelligent Drone Management System - Mixed Complexity
版本: 2.0.0
"""

    """无人机智能管理系统主类"""

    def __init__(self):
        self.system_name = "无人机智能管理系统"
        self.version = "2.0.0"
        self.alerts = []

    def log_alert(self, alert_data: Dict[str, Any]):
        """记录告警信息"""
        alert_data['timestamp'] = datetime.now().isoformat()
        self.alerts.append(alert_data)


def flight_control_system(x: float, y: float, z: float) -> Dict[str, Any]:
    """
    无人机飞行控制系统
    x: 飞行速度 (m/s, 0-30)
    y: 飞行高度 (m, 0-500)
    z: 飞行稳定性指数 (0-100)
    """
    result = {
        'module': '飞行控制系统',
        'input_values': {'x': x, 'y': y, 'z': z}
    }


    # 复杂判断语句 A1: 多维度综合评估
    if (x * y) / (z + 1) > 50 and x ** 2 + y ** 2 > z ** 2:
        pattern_type = "1A1"
        pattern_description = "多维度指标异常-组合风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((x * y) / (z + 1)) / 100)
        result.update(locals())
        return result

    # 复杂判断语句 A2: 非线性关系检测
    if (x ** 2 - y ** 2) / (z + 0.1) < -30 or (abs(x - y) * z) / 100 > 45:
        pattern_type = "1A2"
        pattern_description = "非线性指标偏离"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, abs((x ** 2 - y ** 2) / (z + 0.1)) / 50)
        result.update(locals())
        return result

    # 复杂判断语句 A3: 动态平衡失调
    if ((x + y) / 2) ** 2 < z * 20 and (x * y * z) / 1000 < 35:
        pattern_type = "1A3"
        pattern_description = "动态平衡严重失调"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * y * z) / 1000) / 35)
        result.update(locals())
        return result
    # 判断语句 1: 极速飞行失控模式 - 复杂
    if ((x + y / 10) / (z + 1) > 0.8 and (x ** 2 + y / 50) > 450) or (x > 25 and z < 25 and y > 200):
        pattern_type = 1
        pattern_description = "极速飞行失控警报"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((x + y / 10) / (z + 1)) / 0.8)
        result.update(locals())
        return result

    # 判断语句 2: 高速低空危险模式 - 简单
    if x > 20 and y < 30 and z < 60:
        pattern_type = 2
        pattern_description = "高速低空飞行危险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, x / 30)
        result.update(locals())
        return result

    # 判断语句 3: 高空强风模式 - 中等
    if (y > 300 and z < 50) or (y > 250 and z < 40 and (x * z) / 100 < 8):
        pattern_type = 3
        pattern_description = "高空强风影响警报"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (y / 500) * (1 - z / 100))
        result.update(locals())
        return result

    # 判断语句 4: 飞行姿态异常模式 - 简单
    if z < 40 and x > 10:
        pattern_type = 4
        pattern_description = "飞行姿态严重异常"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (60 - z) / 60)
        result.update(locals())
        return result

    # 判断语句 5: 超高空风险模式 - 复杂
    if (y > 400 and x > 15 and math.sqrt(y * x) > 80) or (y > 380 and x > 12 and z < 70):
        pattern_type = 5
        pattern_description = "超高空飞行风险"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, y / 500)
        result.update(locals())
        return result

    # 判断语句 6: 速度高度失配模式 - 中等
    if abs(x * 10 - y) > 150 or (abs(x * 10 - y) > 100 and z < 65):
        pattern_type = 6
        pattern_description = "速度高度严重失配"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, abs(x * 10 - y) / 300)
        result.update(locals())
        return result

    # 判断语句 7: 极低稳定性模式 - 简单
    if z < 30:
        pattern_type = 7
        pattern_description = "飞行稳定性极低"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (50 - z) / 50)
        result.update(locals())
        return result

    # 判断语句 8: 超低空慢速模式 - 复杂
    if (y < 10 and x < 3 and z > 70) or (y < 8 and x < 2.5 and math.log(z + 1) > 4.2):
        pattern_type = 8
        pattern_description = "超低空慢速飞行模式"
        severity_level = "medium"
        action_required = True
        confidence_score = 0.6
        result.update(locals())
        return result

    # 判断语句 9: 悬停不稳定模式 - 中等
    if (x < 2 and z < 65) or (x < 1.5 and z < 70 and y < 50):
        pattern_type = 9
        pattern_description = "悬停状态不稳定"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, (70 - z) / 70)
        result.update(locals())
        return result

    # 判断语句 10: 理想飞行模式 - 简单
    if 8 <= x <= 15 and 50 <= y <= 150 and z >= 75:
        pattern_type = 10
        pattern_description = "飞行状态理想"
        severity_level = "low"
        action_required = False
        confidence_score = 0.95
        result.update(locals())
        return result

    pattern_type = 0
    pattern_description = "飞行控制正常"
    severity_level = "low"
    action_required = False
    confidence_score = 0.7
    result.update(locals())
    return result


def battery_management_system(x: float, y: float, z: float) -> Dict[str, Any]:
    """
    无人机电池管理系统
    x: 电池电量 (%, 0-100)
    y: 电池温度 (℃, 0-80)
    z: 充放电功率 (W, 0-500)
    """
    result = {
        'module': '电池管理系统',
        'input_values': {'x': x, 'y': y, 'z': z}
    }


    # 复杂判断语句 A1: 多维度综合评估
    if (x * y) / (z + 10) > 50 and x ** 2 + y ** 2 > z ** 2:
        pattern_type = "11A1"
        pattern_description = "多维度指标异常-组合风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((x * y) / (z + 1)) / 100)
        result.update(locals())
        return result

    # 复杂判断语句 A2: 非线性关系检测
    if (x ** 2 - y ** 2) / (z + 0.1) < -30 or (abs(x - y) * z) / 100 > 45:
        pattern_type = "11A2"
        pattern_description = "非线性指标偏离"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, abs((x ** 2 - y ** 2) / (z + 0.1)) / 50)
        result.update(locals())
        return result

    # 复杂判断语句 A3: 动态平衡失调
    if ((x + y) / 2) ** 2 < z * 20 and (x * y * z) / 1000 < 35:
        pattern_type = "11A3"
        pattern_description = "动态平衡严重失调"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * y * z) / 1000) / 35)
        result.update(locals())
        return result
    # 判断语句 11: 电池临界低电量模式 - 复杂
    if ((x * z) / (y + 10) < 50 and x < 30) or (x < 20 and z > 150 and (y ** 2) / (x + 1) > 30):
        pattern_type = 11
        pattern_description = "电池临界低电量警报"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, 1.0 - x / 100)
        result.update(locals())
        return result

    # 判断语句 12: 电池过热模式 - 简单
    if y > 60 and z > 200:
        pattern_type = 12
        pattern_description = "电池严重过热"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, y / 80)
        result.update(locals())
        return result

    # 判断语句 13: 高功率低电量模式 - 中等
    if (z > 350 and x < 25) or (z > 300 and x < 20 and y > 45):
        pattern_type = 13
        pattern_description = "高功率低电量危险"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, z / 500)
        result.update(locals())
        return result

    # 判断语句 14: 电池温度异常模式 - 复杂
    if ((y ** 2) / (x + 1) > 40) or (y > 55 and x < 40 and math.exp((y - 40) / 20) > 2):
        pattern_type = 14
        pattern_description = "电池温度异常升高"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, ((y ** 2) / (x + 1)) / 60)
        result.update(locals())
        return result

    # 判断语句 15: 紧急低电量模式 - 简单
    if x < 15:
        pattern_type = 15
        pattern_description = "电池紧急低电量"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (20 - x) / 20)
        result.update(locals())
        return result

    # 判断语句 16: 低温高功率模式 - 复杂
    if (y < 10 and z > 250 and (z / (y + 1)) > 30) or (y < 15 and z > 280):
        pattern_type = 16
        pattern_description = "低温高功率运行风险"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (15 - y + z / 100) / 20)
        result.update(locals())
        return result

    # 判断语句 17: 充电过热模式 - 中等
    if (y > 50 and z > 300 and x > 80) or (y > 55 and z > 250):
        pattern_type = 17
        pattern_description = "充电过程过热"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, y / 70)
        result.update(locals())
        return result

    # 判断语句 18: 功率波动异常模式 - 简单
    if abs(z - 200) > 180 and x < 50:
        pattern_type = 18
        pattern_description = "充放电功率波动异常"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, abs(z - 200) / 300)
        result.update(locals())
        return result

    # 判断语句 19: 低效能运行模式 - 中等
    if (x + z / 5) / (y + 5) < 5 or (x < 35 and z < 120):
        pattern_type = 19
        pattern_description = "电池低效能运行"
        severity_level = "medium"
        action_required = True
        confidence_score = 0.65
        result.update(locals())
        return result

    # 判断语句 20: 电池状态理想模式 - 简单
    if 60 <= x <= 90 and 20 <= y <= 40 and 100 <= z <= 250:
        pattern_type = 20
        pattern_description = "电池状态理想"
        severity_level = "low"
        action_required = False
        confidence_score = 0.98
        result.update(locals())
        return result

    pattern_type = 0
    pattern_description = "电池系统正常运行"
    severity_level = "low"
    action_required = False
    confidence_score = 0.75
    result.update(locals())
    return result


def navigation_system(x: float, y: float, z: float) -> Dict[str, Any]:
    """
    无人机导航系统
    x: GPS信号强度 (0-100)
    y: 定位精度 (m, 0-50)
    z: 航线偏离度 (m, 0-100)
    """
    result = {
        'module': '导航系统',
        'input_values': {'x': x, 'y': y, 'z': z}
    }


    # 复杂判断语句 A1: 多维度综合评估
    if (x * y) / (z + 1) > 50 and x ** 2 + y ** 2 > z ** 2:
        pattern_type = "21A1"
        pattern_description = "多维度指标异常-组合风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((x * y) / (z + 1)) / 100)
        result.update(locals())
        return result

    # 复杂判断语句 A2: 非线性关系检测
    if (x ** 2 - y ** 2) / (z + 0.1) < -30 or (abs(x - y) * z) / 100 > 45:
        pattern_type = "21A2"
        pattern_description = "非线性指标偏离"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, abs((x ** 2 - y ** 2) / (z + 0.1)) / 50)
        result.update(locals())
        return result

    # 复杂判断语句 A3: 动态平衡失调
    if ((x + y) / 2) ** 2 < z * 20 and (x * y * z) / 1000 < 35:
        pattern_type = "21A3"
        pattern_description = "动态平衡严重失调"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * y * z) / 1000) / 35)
        result.update(locals())
        return result
    # 判断语句 21: GPS信号丢失模式 - 复杂
    if (x < 20 and (y * z) / 100 > 10 and math.sqrt(y * z) > 15) or (x < 15 and y > 20):
        pattern_type = 21
        pattern_description = "GPS信号严重丢失"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (30 - x) / 30)
        result.update(locals())
        return result

    # 判断语句 22: 定位严重失准模式 - 简单
    if y > 30 and z > 50:
        pattern_type = 22
        pattern_description = "定位严重失准"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (y / 50 + z / 100) / 2)
        result.update(locals())
        return result

    # 判断语句 23: 航线严重偏离模式 - 中等
    if (z > 60 and x < 70) or (z > 55 and x < 65 and y > 20):
        pattern_type = 23
        pattern_description = "航线严重偏离"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, z / 100)
        result.update(locals())
        return result

    # 判断语句 24: GPS弱信号模式 - 复杂
    if (x < 40 and y > 15 and (x ** 2) / (y + 1) < 60) or (x < 35 and y > 12):
        pattern_type = 24
        pattern_description = "GPS信号弱定位不准"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (50 - x + y) / 65)
        result.update(locals())
        return result

    # 判断语句 25: 航线漂移模式 - 简单
    if 30 <= z <= 50 and x < 80:
        pattern_type = 25
        pattern_description = "航线持续漂移"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, z / 80)
        result.update(locals())
        return result

    # 判断语句 26: 信号不稳定模式 - 复杂
    if ((x ** 2) / (y * z + 1) < 40) or (x < 65 and (y + z) > 60 and math.log(x + 1) < 3.8):
        pattern_type = 26
        pattern_description = "导航信号不稳定"
        severity_level = "medium"
        action_required = True
        confidence_score = 0.7
        result.update(locals())
        return result

    # 判断语句 27: 定位精度下降模式 - 中等
    if (10 <= y <= 20 and z > 20) or (12 <= y <= 22 and z > 18 and x < 75):
        pattern_type = 27
        pattern_description = "定位精度持续下降"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, (y + z / 5) / 30)
        result.update(locals())
        return result

    # 判断语句 28: 轻微偏航模式 - 简单
    if 15 <= z <= 25 and x > 70:
        pattern_type = 28
        pattern_description = "轻微偏离航线"
        severity_level = "low"
        action_required = True
        confidence_score = 0.5
        result.update(locals())
        return result

    # 判断语句 29: 信号恢复中模式 - 中等
    if (50 <= x < 70 and y < 10) or (55 <= x < 75 and y < 12 and z < 15):
        pattern_type = 29
        pattern_description = "信号恢复中"
        severity_level = "low"
        action_required = True
        confidence_score = 0.6
        result.update(locals())
        return result

    # 判断语句 30: 导航状态理想模式 - 简单
    if x >= 85 and y <= 3 and z <= 5:
        pattern_type = 30
        pattern_description = "导航状态理想"
        severity_level = "low"
        action_required = False
        confidence_score = 0.96
        result.update(locals())
        return result

    pattern_type = 0
    pattern_description = "导航系统正常"
    severity_level = "low"
    action_required = False
    confidence_score = 0.8
    result.update(locals())
    return result


def payload_management_system(x: float, y: float, z: float) -> Dict[str, Any]:
    """
    无人机载荷管理系统
    x: 载重量 (kg, 0-50)
    y: 载荷平衡度 (0-100)
    z: 载荷功耗 (W, 0-300)
    """
    result = {
        'module': '载荷管理系统',
        'input_values': {'x': x, 'y': y, 'z': z}
    }


    # 复杂判断语句 A1: 多维度综合评估
    if (x * y) / (z + 1) > 50 and x ** 2 + y ** 2 > z ** 2:
        pattern_type = "31A1"
        pattern_description = "多维度指标异常-组合风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((x * y) / (z + 1)) / 100)
        result.update(locals())
        return result

    # 复杂判断语句 A2: 非线性关系检测
    if (x ** 2 - y ** 2) / (z + 0.1) < -30 or (abs(x - y) * z) / 100 > 45:
        pattern_type = "31A2"
        pattern_description = "非线性指标偏离"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, abs((x ** 2 - y ** 2) / (z + 0.1)) / 50)
        result.update(locals())
        return result

    # 复杂判断语句 A3: 动态平衡失调
    if ((x + y) / 2) ** 2 < z * 20 and (x * y * z) / 1000 < 35:
        pattern_type = "31A3"
        pattern_description = "动态平衡严重失调"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * y * z) / 1000) / 35)
        result.update(locals())
        return result
    # 判断语句 31: 超载危险模式 - 复杂
    if ((x * z) / (y + 1) > 800 and x > 40) or (x > 45 and z > 250 and y < 30):
        pattern_type = 31
        pattern_description = "无人机超载危险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((x * z) / (y + 1)) / 1000)
        result.update(locals())
        return result

    # 判断语句 32: 严重失衡模式 - 简单
    if y < 30 and x > 25:
        pattern_type = 32
        pattern_description = "载荷严重失衡"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (40 - y + x) / 60)
        result.update(locals())
        return result

    # 判断语句 33: 高载高功耗模式 - 中等
    if (x > 35 and z > 220) or (x > 32 and z > 200 and y < 60):
        pattern_type = 33
        pattern_description = "高载荷高功耗风险"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (x / 50 + z / 300) / 2)
        result.update(locals())
        return result

    # 判断语句 34: 载荷不稳定模式 - 复杂
    if ((x ** 2) / (y + 1) > 30 and z > 150) or (x > 30 and y < 40 and math.sqrt(x * z) > 90):
        pattern_type = 34
        pattern_description = "载荷状态不稳定"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, ((x ** 2) / (y + 1)) / 50)
        result.update(locals())
        return result

    # 判断语句 35: 重载低平衡模式 - 简单
    if x > 30 and y < 50:
        pattern_type = 35
        pattern_description = "重载低平衡风险"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, x / 45)
        result.update(locals())
        return result

    # 判断语句 36: 载荷功耗异常模式 - 中等
    if (z > 250 and x < 20) or (z > 230 and x < 18 and y < 70):
        pattern_type = 36
        pattern_description = "载荷功耗异常偏高"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, z / 300)
        result.update(locals())
        return result

    # 判断语句 37: 轻载失衡模式 - 简单
    if x < 10 and y < 60:
        pattern_type = 37
        pattern_description = "轻载荷平衡度不足"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, (70 - y) / 70)
        result.update(locals())
        return result

    # 判断语句 38: 载荷偏移模式 - 复杂
    if (abs(y - 70) > 25 and x > 15 and (x * abs(y - 70)) / 100 > 5) or (abs(y - 70) > 30):
        pattern_type = 38
        pattern_description = "载荷重心偏移"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, abs(y - 70) / 40)
        result.update(locals())
        return result

    # 判断语句 39: 低效载荷模式 - 中等
    if (x + z / 10) / (y + 1) < 0.8 or (x < 12 and z < 100):
        pattern_type = 39
        pattern_description = "载荷效率偏低"
        severity_level = "low"
        action_required = True
        confidence_score = 0.55
        result.update(locals())
        return result

    # 判断语句 40: 载荷状态理想模式 - 简单
    if 10 <= x <= 25 and 75 <= y <= 95 and 80 <= z <= 180:
        pattern_type = 40
        pattern_description = "载荷状态理想"
        severity_level = "low"
        action_required = False
        confidence_score = 0.93
        result.update(locals())
        return result

    pattern_type = 0
    pattern_description = "载荷管理正常"
    severity_level = "low"
    action_required = False
    confidence_score = 0.78
    result.update(locals())
    return result


def communication_system(x: float, y: float, z: float) -> Dict[str, Any]:
    """
    无人机通信系统
    x: 信号强度 (dBm, -120 to -20, 转换为0-100)
    y: 数据传输速率 (Mbps, 0-100)
    z: 通信延迟 (ms, 0-1000)
    """
    result = {
        'module': '通信系统',
        'input_values': {'x': x, 'y': y, 'z': z}
    }


    # 复杂判断语句 A1: 多维度综合评估
    if (x * y) / (z + 1) > 50 and x ** 2 + y ** 2 > z ** 2:
        pattern_type = "41A1"
        pattern_description = "多维度指标异常-组合风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((x * y) / (z + 1)) / 100)
        result.update(locals())
        return result

    # 复杂判断语句 A2: 非线性关系检测
    if (x ** 2 - y ** 2) / (z + 0.1) < -30 or (abs(x - y) * z) / 100 > 45:
        pattern_type = "41A2"
        pattern_description = "非线性指标偏离"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, abs((x ** 2 - y ** 2) / (z + 0.1)) / 50)
        result.update(locals())
        return result

    # 复杂判断语句 A3: 动态平衡失调
    if ((x + y) / 2) ** 2 < z * 20 and (x * y * z) / 1000 < 35:
        pattern_type = "41A3"
        pattern_description = "动态平衡严重失调"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * y * z) / 1000) / 35)
        result.update(locals())
        return result
    # 判断语句 41: 通信完全中断模式 - 复杂
    if (x < 15 and (y * 100) / (z + 1) < 5 and math.sqrt(x * y) < 20) or (x < 12 and z > 600):
        pattern_type = 41
        pattern_description = "通信完全中断"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (25 - x) / 25)
        result.update(locals())
        return result

    # 判断语句 42: 信号极弱模式 - 简单
    if x < 25 and z > 500:
        pattern_type = 42
        pattern_description = "信号极弱高延迟"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (35 - x + z / 100) / 50)
        result.update(locals())
        return result

    # 判断语句 43: 严重延迟模式 - 中等
    if (z > 600 and y < 30) or (z > 550 and y < 25 and x < 50):
        pattern_type = 43
        pattern_description = "通信严重延迟"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, z / 1000)
        result.update(locals())
        return result

    # 判断语句 44: 低速传输模式 - 简单
    if y < 20 and x > 40:
        pattern_type = 44
        pattern_description = "数据传输速率过低"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (30 - y) / 30)
        result.update(locals())
        return result

    # 判断语句 45: 信号不稳定模式 - 复杂
    if ((x ** 2) / (y * 10 + z / 10 + 1) < 8) or (x < 55 and y < 35 and z > 250):
        pattern_type = 45
        pattern_description = "通信信号不稳定"
        severity_level = "high"
        action_required = True
        confidence_score = 0.75
        result.update(locals())
        return result

    # 判断语句 46: 中等延迟模式 - 中等
    if (300 <= z <= 500 and y > 40) or (280 <= z <= 480 and y > 35):
        pattern_type = 46
        pattern_description = "通信存在中等延迟"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, z / 800)
        result.update(locals())
        return result

    # 判断语句 47: 信号弱化模式 - 简单
    if 30 <= x < 50 and z > 200:
        pattern_type = 47
        pattern_description = "信号强度弱化"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, (60 - x + z / 50) / 80)
        result.update(locals())
        return result

    # 判断语句 48: 传输波动模式 - 复杂
    if (abs(y - 50) > 30 and z < 300 and (y ** 2) / (x + 1) > 25) or abs(y - 50) > 35:
        pattern_type = 48
        pattern_description = "数据传输速率波动"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, abs(y - 50) / 60)
        result.update(locals())
        return result

    # 判断语句 49: 轻度干扰模式 - 中等
    if (100 <= z <= 200 and x > 60) or (90 <= z <= 190 and x > 55):
        pattern_type = 49
        pattern_description = "通信存在轻度干扰"
        severity_level = "low"
        action_required = True
        confidence_score = 0.5
        result.update(locals())
        return result

    # 判断语句 50: 通信状态理想模式 - 简单
    if x >= 75 and y >= 60 and z <= 80:
        pattern_type = 50
        pattern_description = "通信状态理想"
        severity_level = "low"
        action_required = False
        confidence_score = 0.97
        result.update(locals())
        return result

    pattern_type = 0
    pattern_description = "通信系统正常"
    severity_level = "low"
    action_required = False
    confidence_score = 0.82
    result.update(locals())
    return result


def obstacle_avoidance_system(x: float, y: float, z: float) -> Dict[str, Any]:
    """
    无人机避障系统
    x: 前方障碍物距离 (m, 0-100)
    y: 避障响应时间 (ms, 0-500)
    z: 传感器可靠性 (0-100)
    """
    result = {
        'module': '避障系统',
        'input_values': {'x': x, 'y': y, 'z': z}
    }


    # 复杂判断语句 A1: 多维度综合评估
    if (x * y) / (z + 1) > 50 and x ** 2 + y ** 2 > z ** 2:
        pattern_type = "51A1"
        pattern_description = "多维度指标异常-组合风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((x * y) / (z + 1)) / 100)
        result.update(locals())
        return result

    # 复杂判断语句 A2: 非线性关系检测
    if (x ** 2 - y ** 2) / (z + 0.1) < -30 or (abs(x - y) * z) / 100 > 45:
        pattern_type = "51A2"
        pattern_description = "非线性指标偏离"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, abs((x ** 2 - y ** 2) / (z + 0.1)) / 50)
        result.update(locals())
        return result

    # 复杂判断语句 A3: 动态平衡失调
    if ((x + y) / 2) ** 2 < z * 20 and (x * y * z) / 1000 < 35:
        pattern_type = "51A3"
        pattern_description = "动态平衡严重失调"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * y * z) / 1000) / 35)
        result.update(locals())
        return result
    # 判断语句 51: 紧急碰撞风险模式 - 复杂
    if ((x * z) / (y + 1) < 20 and x < 10) or (x < 8 and y > 150 and math.sqrt(y * z) < 60):
        pattern_type = 51
        pattern_description = "紧急碰撞风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * z) / (y + 1)) / 20)
        result.update(locals())
        return result

    # 判断语句 52: 近距离障碍物模式 - 简单
    if x < 5 and y > 200:
        pattern_type = 52
        pattern_description = "近距离障碍物预警"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (10 - x + y / 50) / 15)
        result.update(locals())
        return result

    # 判断语句 53: 传感器失效模式 - 中等
    if (z < 30 and x < 20) or (z < 35 and x < 25 and y > 250):
        pattern_type = 53
        pattern_description = "避障传感器失效"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (40 - z) / 40)
        result.update(locals())
        return result

    # 判断语句 54: 响应严重延迟模式 - 简单
    if y > 300 and x < 30:
        pattern_type = 54
        pattern_description = "避障响应严重延迟"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, y / 500)
        result.update(locals())
        return result

    # 判断语句 55: 中距离预警模式 - 复杂
    if (10 <= x <= 20 and z < 60 and (x * z) / 100 < 10) or (12 <= x <= 22 and z < 55):
        pattern_type = 55
        pattern_description = "中距离障碍物预警"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (25 - x + 70 - z) / 80)
        result.update(locals())
        return result

    # 判断语句 56: 传感器精度下降模式 - 中等
    if ((z ** 2) / (x + y / 10 + 1) < 40) or (z < 65 and x < 40):
        pattern_type = 56
        pattern_description = "传感器精度下降"
        severity_level = "high"
        action_required = True
        confidence_score = 0.7
        result.update(locals())
        return result

    # 判断语句 57: 响应时间偏长模式 - 简单
    if 150 <= y <= 250 and x < 50:
        pattern_type = 57
        pattern_description = "避障响应时间偏长"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, y / 400)
        result.update(locals())
        return result

    # 判断语句 58: 传感器干扰模式 - 复杂
    if (40 <= z < 70 and y > 100 and (z * y) / 100 < 60) or (z < 65 and y > 120):
        pattern_type = 58
        pattern_description = "传感器受到干扰"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, (80 - z + y / 50) / 100)
        result.update(locals())
        return result

    # 判断语句 59: 轻度预警模式 - 中等
    if (30 <= x <= 50 and z > 70) or (35 <= x <= 55 and z > 65):
        pattern_type = 59
        pattern_description = "远距离障碍物监测"
        severity_level = "low"
        action_required = True
        confidence_score = 0.45
        result.update(locals())
        return result

    # 判断语句 60: 避障系统理想模式 - 简单
    if x >= 60 and y <= 80 and z >= 85:
        pattern_type = 60
        pattern_description = "避障系统状态理想"
        severity_level = "low"
        action_required = False
        confidence_score = 0.94
        result.update(locals())
        return result

    pattern_type = 0
    pattern_description = "避障系统正常"
    severity_level = "low"
    action_required = False
    confidence_score = 0.76
    result.update(locals())
    return result


def image_capture_system(x: float, y: float, z: float) -> Dict[str, Any]:
    """
    无人机图像采集系统
    x: 图像清晰度 (0-100)
    y: 拍摄稳定性 (0-100)
    z: 存储空间使用率 (%, 0-100)
    """
    result = {
        'module': '图像采集系统',
        'input_values': {'x': x, 'y': y, 'z': z}
    }


    # 复杂判断语句 A1: 多维度综合评估
    if (x * y) / (z + 1) > 50 and x ** 2 + y ** 2 > z ** 2:
        pattern_type = "61A1"
        pattern_description = "多维度指标异常-组合风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((x * y) / (z + 1)) / 100)
        result.update(locals())
        return result

    # 复杂判断语句 A2: 非线性关系检测
    if (x ** 2 - y ** 2) / (z + 0.1) < -30 or (abs(x - y) * z) / 100 > 45:
        pattern_type = "61A2"
        pattern_description = "非线性指标偏离"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, abs((x ** 2 - y ** 2) / (z + 0.1)) / 50)
        result.update(locals())
        return result

    # 复杂判断语句 A3: 动态平衡失调
    if ((x + y) / 2) ** 2 < z * 20 and (x * y * z) / 1000 < 35:
        pattern_type = "61A3"
        pattern_description = "动态平衡严重失调"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * y * z) / 1000) / 35)
        result.update(locals())
        return result
    # 判断语句 61: 图像质量极差模式 - 复杂
    if ((x * y) / (z + 1) < 25 and x < 40) or (x < 35 and y < 30 and math.log(x * y + 1) < 7):
        pattern_type = 61
        pattern_description = "图像质量极差"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * y) / (z + 1)) / 25)
        result.update(locals())
        return result

    # 判断语句 62: 存储空间告急模式 - 简单
    if z > 92 and x > 70:
        pattern_type = 62
        pattern_description = "存储空间严重不足"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, z / 100)
        result.update(locals())
        return result

    # 判断语句 63: 拍摄严重抖动模式 - 中等
    if (y < 30 and x > 60) or (y < 35 and x > 65 and z > 70):
        pattern_type = 63
        pattern_description = "拍摄严重抖动"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (40 - y) / 40)
        result.update(locals())
        return result

    # 判断语句 64: 图像模糊不清模式 - 简单
    if x < 40 and y < 50:
        pattern_type = 64
        pattern_description = "图像模糊不清晰"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (50 - x) / 50)
        result.update(locals())
        return result

    # 判断语句 65: 存储空间紧张模式 - 复杂
    if (80 <= z <= 90 and x > 50 and (z * x) / 100 > 45) or (z > 85 and x > 55):
        pattern_type = 65
        pattern_description = "存储空间紧张"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, z / 100)
        result.update(locals())
        return result

    # 判断语句 66: 画质下降模式 - 中等
    if ((x ** 2) / (y + 10) < 50) or (x < 60 and y < 60):
        pattern_type = 66
        pattern_description = "画质持续下降"
        severity_level = "medium"
        action_required = True
        confidence_score = 0.65
        result.update(locals())
        return result

    # 判断语句 67: 稳定性不足模式 - 简单
    if 40 <= y < 65 and z > 50:
        pattern_type = 67
        pattern_description = "拍摄稳定性不足"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, (70 - y) / 70)
        result.update(locals())
        return result

    # 判断语句 68: 清晰度待提升模式 - 复杂
    if (50 <= x < 70 and y > 60 and (x * y) / 100 < 45) or (55 <= x < 75 and y > 55):
        pattern_type = 68
        pattern_description = "图像清晰度待提升"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, (80 - x) / 80)
        result.update(locals())
        return result

    # 判断语句 69: 存储容量关注模式 - 中等
    if (60 <= z < 75 and x > 70) or (65 <= z < 80 and x > 65):
        pattern_type = 69
        pattern_description = "存储容量需关注"
        severity_level = "low"
        action_required = True
        confidence_score = 0.5
        result.update(locals())
        return result

    # 判断语句 70: 图像采集理想模式 - 简单
    if x >= 80 and y >= 75 and 20 <= z <= 50:
        pattern_type = 70
        pattern_description = "图像采集状态理想"
        severity_level = "low"
        action_required = False
        confidence_score = 0.95
        result.update(locals())
        return result

    pattern_type = 0
    pattern_description = "图像采集系统正常"
    severity_level = "low"
    action_required = False
    confidence_score = 0.79
    result.update(locals())
    return result


def motor_system(x: float, y: float, z: float) -> Dict[str, Any]:
    """
    无人机电机系统
    x: 电机转速 (RPM, 0-10000, 归一化为0-100)
    y: 电机温度 (℃, 0-120)
    z: 电机效率 (%, 0-100)
    """
    result = {
        'module': '电机系统',
        'input_values': {'x': x, 'y': y, 'z': z}
    }


    # 复杂判断语句 A1: 多维度综合评估
    if (x * y) / (z + 1) > 50 and x ** 2 + y ** 2 > z ** 2:
        pattern_type = "71A1"
        pattern_description = "多维度指标异常-组合风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((x * y) / (z + 1)) / 100)
        result.update(locals())
        return result

    # 复杂判断语句 A2: 非线性关系检测
    if (x ** 2 - y ** 2) / (z + 0.1) < -30 or (abs(x - y) * z) / 100 > 45:
        pattern_type = "71A2"
        pattern_description = "非线性指标偏离"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, abs((x ** 2 - y ** 2) / (z + 0.1)) / 50)
        result.update(locals())
        return result

    # 复杂判断语句 A3: 动态平衡失调
    if ((x + y) / 2) ** 2 < z * 20 and (x * y * z) / 1000 < 35:
        pattern_type = "71A3"
        pattern_description = "动态平衡严重失调"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * y * z) / 1000) / 35)
        result.update(locals())
        return result
    # 判断语句 71: 电机过热失效模式 - 复杂
    if ((x * y) / (z + 1) > 200 and y > 85) or (x > 80 and y > 85 and z < 40 and math.exp((y - 60) / 30) > 2):
        pattern_type = 71
        pattern_description = "电机过热失效风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((x * y) / (z + 1)) / 250)
        result.update(locals())
        return result

    # 判断语句 72: 高转速过热模式 - 简单
    if x > 85 and y > 90:
        pattern_type = 72
        pattern_description = "高转速严重过热"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (x / 100 + y / 120) / 2)
        result.update(locals())
        return result

    # 判断语句 73: 效率极低模式 - 中等
    if (z < 30 and x > 50) or (z < 35 and x > 60 and y > 70):
        pattern_type = 73
        pattern_description = "电机效率极低"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (40 - z) / 40)
        result.update(locals())
        return result

    # 判断语句 74: 温度异常升高模式 - 复杂
    if ((y ** 2) / (z + 1) > 100) or (y > 75 and z < 55 and (x * y) / 100 > 65):
        pattern_type = 74
        pattern_description = "电机温度异常升高"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, ((y ** 2) / (z + 1)) / 150)
        result.update(locals())
        return result

    # 判断语句 75: 高负载运行模式 - 简单
    if x > 80 and z < 60:
        pattern_type = 75
        pattern_description = "电机高负载低效运行"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (x / 100 + (70 - z) / 70) / 2)
        result.update(locals())
        return result

    # 判断语句 76: 低温低效模式 - 复杂
    if (y < 20 and z < 50 and (y * z) / 100 < 8) or (y < 25 and z < 45 and x > 40):
        pattern_type = 76
        pattern_description = "低温环境效率下降"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (30 - y + 60 - z) / 70)
        result.update(locals())
        return result

    # 判断语句 77: 转速不稳定模式 - 中等
    if (abs(x - 60) > 30 and y > 60) or (abs(x - 60) > 25 and y > 55):
        pattern_type = 77
        pattern_description = "电机转速不稳定"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, abs(x - 60) / 50)
        result.update(locals())
        return result

    # 判断语句 78: 温度偏高模式 - 简单
    if 70 <= y < 85 and z < 70:
        pattern_type = 78
        pattern_description = "电机温度偏高"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, y / 120)
        result.update(locals())
        return result

    # 判断语句 79: 效率待提升模式 - 中等
    if (50 <= z < 70 and x > 40) or (55 <= z < 75 and x > 45):
        pattern_type = 79
        pattern_description = "电机效率待提升"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, (75 - z) / 75)
        result.update(locals())
        return result

    # 判断语句 80: 电机状态理想模式 - 简单
    if 40 <= x <= 70 and 30 <= y <= 60 and z >= 80:
        pattern_type = 80
        pattern_description = "电机运行状态理想"
        severity_level = "low"
        action_required = False
        confidence_score = 0.96
        result.update(locals())
        return result

    pattern_type = 0
    pattern_description = "电机系统正常"
    severity_level = "low"
    action_required = False
    confidence_score = 0.81
    result.update(locals())
    return result


def autopilot_system(x: float, y: float, z: float) -> Dict[str, Any]:
    """
    无人机自动驾驶系统
    x: 自动控制精度 (%, 0-100)
    y: 决策响应时间 (ms, 0-1000)
    z: AI算法置信度 (%, 0-100)
    """
    result = {
        'module': '自动驾驶系统',
        'input_values': {'x': x, 'y': y, 'z': z}
    }


    # 复杂判断语句 A1: 多维度综合评估
    if (x * y) / (z + 1) > 50 and x ** 2 + y ** 2 > z ** 2:
        pattern_type = "81A1"
        pattern_description = "多维度指标异常-组合风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((x * y) / (z + 1)) / 100)
        result.update(locals())
        return result

    # 复杂判断语句 A2: 非线性关系检测
    if (x ** 2 - y ** 2) / (z + 0.1) < -30 or (abs(x - y) * z) / 100 > 45:
        pattern_type = "81A2"
        pattern_description = "非线性指标偏离"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, abs((x ** 2 - y ** 2) / (z + 0.1)) / 50)
        result.update(locals())
        return result

    # 复杂判断语句 A3: 动态平衡失调
    if ((x + y) / 2) ** 2 < z * 20 and (x * y * z) / 1000 < 35:
        pattern_type = "81A3"
        pattern_description = "动态平衡严重失调"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * y * z) / 1000) / 35)
        result.update(locals())
        return result
    # 判断语句 81: 自动驾驶失控模式 - 复杂
    if ((x * z) / (y + 1) < 8 and x < 50) or (x < 45 and z < 45 and y > 500 and math.sqrt(x * z) < 45):
        pattern_type = 81
        pattern_description = "自动驾驶严重失控"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * z) / (y + 1)) / 8)
        result.update(locals())
        return result

    # 判断语句 82: 控制精度极低模式 - 简单
    if x < 40 and z < 50:
        pattern_type = 82
        pattern_description = "控制精度极低"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (50 - x) / 50)
        result.update(locals())
        return result

    # 判断语句 83: 决策严重延迟模式 - 中等
    if (y > 600 and x < 70) or (y > 550 and x < 65 and z < 60):
        pattern_type = 83
        pattern_description = "决策严重延迟"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, y / 1000)
        result.update(locals())
        return result

    # 判断语句 84: AI算法失准模式 - 简单
    if z < 40 and x > 50:
        pattern_type = 84
        pattern_description = "AI算法判断失准"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (50 - z) / 50)
        result.update(locals())
        return result

    # 判断语句 85: 精度置信度双低模式 - 复杂
    if (x < 60 and z < 65 and (x * z) / 100 < 35) or (x < 55 and z < 60):
        pattern_type = 85
        pattern_description = "精度和置信度双低"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (70 - x + 75 - z) / 80)
        result.update(locals())
        return result

    # 判断语句 86: 响应时间过长模式 - 中等
    if ((y ** 2) / (x * z + 1) > 40) or (y > 400 and x < 70):
        pattern_type = 86
        pattern_description = "系统响应时间过长"
        severity_level = "high"
        action_required = True
        confidence_score = 0.72
        result.update(locals())
        return result

    # 判断语句 87: 控制波动模式 - 简单
    if 60 <= x < 80 and y > 300:
        pattern_type = 87
        pattern_description = "自动控制存在波动"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, y / 800)
        result.update(locals())
        return result

    # 判断语句 88: AI置信度不足模式 - 复杂
    if (50 <= z < 75 and x > 70 and (z * x) / 100 < 55) or (z < 70 and x > 75):
        pattern_type = 88
        pattern_description = "AI算法置信度不足"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, (80 - z) / 80)
        result.update(locals())
        return result

    # 判断语句 89: 轻微延迟模式 - 中等
    if (150 <= y <= 250 and z > 75) or (140 <= y <= 240 and z > 70):
        pattern_type = 89
        pattern_description = "决策存在轻微延迟"
        severity_level = "low"
        action_required = True
        confidence_score = 0.55
        result.update(locals())
        return result

    # 判断语句 90: 自动驾驶理想模式 - 简单
    if x >= 85 and y <= 120 and z >= 85:
        pattern_type = 90
        pattern_description = "自动驾驶状态理想"
        severity_level = "low"
        action_required = False
        confidence_score = 0.98
        result.update(locals())
        return result

    pattern_type = 0
    pattern_description = "自动驾驶系统正常"
    severity_level = "low"
    action_required = False
    confidence_score = 0.83
    result.update(locals())
    return result


def weather_adaptation_system(x: float, y: float, z: float) -> Dict[str, Any]:
    """
    无人机气象适应系统
    x: 风速 (m/s, 0-30)
    y: 降水强度 (mm/h, 0-50)
    z: 能见度 (km, 0-10, 归一化为0-100)
    """
    result = {
        'module': '气象适应系统',
        'input_values': {'x': x, 'y': y, 'z': z}
    }


    # 复杂判断语句 A1: 多维度综合评估
    if (x * y) / (z + 1) > 50 and x ** 2 + y ** 2 > z ** 2:
        pattern_type = "91A1"
        pattern_description = "多维度指标异常-组合风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((x * y) / (z + 1)) / 100)
        result.update(locals())
        return result

    # 复杂判断语句 A2: 非线性关系检测
    if (x ** 2 - y ** 2) / (z + 0.1) < -30 or (abs(x - y) * z) / 100 > 45:
        pattern_type = "91A2"
        pattern_description = "非线性指标偏离"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, abs((x ** 2 - y ** 2) / (z + 0.1)) / 50)
        result.update(locals())
        return result

    # 复杂判断语句 A3: 动态平衡失调
    if ((x + y) / 2) ** 2 < z * 20 and (x * y * z) / 1000 < 35:
        pattern_type = "91A3"
        pattern_description = "动态平衡严重失调"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * y * z) / 1000) / 35)
        result.update(locals())
        return result
    # 判断语句 91: 极端恶劣天气模式 - 复杂
    if ((x + y * 2) / (z + 1) > 5 and x > 15) or (x > 20 and y > 20 and z < 30 and math.sqrt(x * y) > 25):
        pattern_type = 91
        pattern_description = "极端恶劣天气警报"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((x + y * 2) / (z + 1)) / 8)
        result.update(locals())
        return result

    # 判断语句 92: 强风暴雨模式 - 简单
    if x > 18 and y > 25:
        pattern_type = 92
        pattern_description = "强风暴雨危险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (x / 30 + y / 50) / 2)
        result.update(locals())
        return result

    # 判断语句 93: 零能见度模式 - 中等
    if (z < 15 and (x > 12 or y > 15)) or (z < 20 and x > 15 and y > 12):
        pattern_type = 93
        pattern_description = "极低能见度危险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (20 - z) / 20)
        result.update(locals())
        return result

    # 判断语句 94: 强风险模式 - 简单
    if x > 15 and z < 40:
        pattern_type = 94
        pattern_description = "强风环境飞行风险"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, x / 25)
        result.update(locals())
        return result

    # 判断语句 95: 大雨低能见度模式 - 复杂
    if (y > 18 and z < 30 and (y * (40 - z)) / 100 > 5) or (y > 20 and z < 25):
        pattern_type = 95
        pattern_description = "大雨低能见度风险"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (y / 40 + (40 - z) / 40) / 2)
        result.update(locals())
        return result

    # 判断语句 96: 风雨复合模式 - 中等
    if ((x * y) / 100 > 3 and z < 50) or ((x * y) / 100 > 2.5 and z < 45):
        pattern_type = 96
        pattern_description = "风雨复合恶劣天气"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, ((x * y) / 100) / 5)
        result.update(locals())
        return result

    # 判断语句 97: 中等风力模式 - 简单
    if 10 <= x <= 14 and y < 10:
        pattern_type = 97
        pattern_description = "中等风力影响飞行"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, x / 20)
        result.update(locals())
        return result

    # 判断语句 98: 小雨天气模式 - 复杂
    if (8 <= y <= 15 and x < 10 and (y * x) / 10 < 12) or (10 <= y <= 17 and x < 12):
        pattern_type = 98
        pattern_description = "小雨天气需注意"
        severity_level = "medium"
        action_required = True
        confidence_score = 0.6
        result.update(locals())
        return result

    # 判断语句 99: 能见度偏低模式 - 中等
    if (30 <= z < 50 and x < 8) or (35 <= z < 55 and x < 10):
        pattern_type = 99
        pattern_description = "能见度偏低"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, (55 - z) / 55)
        result.update(locals())
        return result

    # 判断语句 100: 理想天气模式 - 简单
    if x <= 6 and y <= 3 and z >= 70:
        pattern_type = 100
        pattern_description = "天气条件理想"
        severity_level = "low"
        action_required = False
        confidence_score = 0.97
        result.update(locals())
        return result

    pattern_type = 0
    pattern_description = "气象条件正常"
    severity_level = "low"
    action_required = False
    confidence_score = 0.78
    result.update(locals())
    return result


def mission_planning_system(x: float, y: float, z: float) -> Dict[str, Any]:
    """
    无人机任务规划系统
    x: 任务完成度 (%, 0-100)
    y: 飞行时长 (分钟, 0-120)
    z: 能源剩余量 (%, 0-100)
    """
    result = {
        'module': '任务规划系统',
        'input_values': {'x': x, 'y': y, 'z': z}
    }


    # 复杂判断语句 A1: 多维度综合评估
    if (x * y) / (z + 1) > 50 and x ** 2 + y ** 2 > z ** 2:
        pattern_type = "101A1"
        pattern_description = "多维度指标异常-组合风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((x * y) / (z + 1)) / 100)
        result.update(locals())
        return result

    # 复杂判断语句 A2: 非线性关系检测
    if (x ** 2 - y ** 2) / (z + 0.1) < -30 or (abs(x - y) * z) / 100 > 45:
        pattern_type = "101A2"
        pattern_description = "非线性指标偏离"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, abs((x ** 2 - y ** 2) / (z + 0.1)) / 50)
        result.update(locals())
        return result

    # 复杂判断语句 A3: 动态平衡失调
    if ((x + y) / 2) ** 2 < z * 20 and (x * y * z) / 1000 < 35:
        pattern_type = "101A3"
        pattern_description = "动态平衡严重失调"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * y * z) / 1000) / 35)
        result.update(locals())
        return result
    # 判断语句 101: 任务中断风险模式 - 复杂
    if ((x * z) / (y + 1) < 40 and z < 30) or (x < 50 and z < 25 and y > 70 and math.sqrt(x * z) < 35):
        pattern_type = 101
        pattern_description = "任务中断风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * z) / (y + 1)) / 40)
        result.update(locals())
        return result

    # 判断语句 102: 能源耗尽模式 - 简单
    if z < 12 and y > 60:
        pattern_type = 102
        pattern_description = "能源即将耗尽"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (20 - z) / 20)
        result.update(locals())
        return result

    # 判断语句 103: 超时低完成度模式 - 中等
    if (y > 90 and x < 50) or (y > 85 and x < 45 and z < 40):
        pattern_type = 103
        pattern_description = "超时低完成度异常"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (y / 120 + (60 - x) / 60) / 2)
        result.update(locals())
        return result

    # 判断语句 104: 能源不足模式 - 简单
    if z < 25 and x < 80:
        pattern_type = 104
        pattern_description = "能源不足影响任务"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (35 - z) / 35)
        result.update(locals())
        return result

    # 判断语句 105: 任务延期风险模式 - 复杂
    if (y > 75 and x < 70 and z < 40 and (x * z) / 100 < 25) or (y > 80 and x < 65):
        pattern_type = 105
        pattern_description = "任务延期风险高"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (y / 100 + (80 - x) / 80) / 2)
        result.update(locals())
        return result

    # 判断语句 106: 效率低下模式 - 中等
    if ((x ** 2) / (y + z + 1) < 8) or (x < 55 and y > 65):
        pattern_type = 106
        pattern_description = "任务执行效率低下"
        severity_level = "high"
        action_required = True
        confidence_score = 0.68
        result.update(locals())
        return result

    # 判断语句 107: 进度缓慢模式 - 简单
    if 40 <= x < 65 and y > 50:
        pattern_type = 107
        pattern_description = "任务进度缓慢"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, y / 120)
        result.update(locals())
        return result

    # 判断语句 108: 能源储备不足模式 - 复杂
    if (30 <= z < 50 and y > 40 and (z * y) / 100 < 18) or (z < 45 and y > 45):
        pattern_type = 108
        pattern_description = "能源储备不足"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, (55 - z) / 55)
        result.update(locals())
        return result

    # 判断语句 109: 时间紧迫模式 - 中等
    if (70 <= x < 90 and y > 60 and z > 50) or (75 <= x < 95 and y > 55):
        pattern_type = 109
        pattern_description = "任务时间紧迫"
        severity_level = "low"
        action_required = True
        confidence_score = 0.5
        result.update(locals())
        return result

    # 判断语句 110: 任务执行理想模式 - 简单
    if x >= 85 and y <= 50 and z >= 60:
        pattern_type = 110
        pattern_description = "任务执行状态理想"
        severity_level = "low"
        action_required = False
        confidence_score = 0.94
        result.update(locals())
        return result

    pattern_type = 0
    pattern_description = "任务规划正常"
    severity_level = "low"
    action_required = False
    confidence_score = 0.77
    result.update(locals())
    return result


def data_link_system(x: float, y: float, z: float) -> Dict[str, Any]:
    """
    无人机数据链系统
    x: 数据完整性 (%, 0-100)
    y: 链路稳定性 (0-100)
    z: 传输带宽利用率 (%, 0-100)
    """
    result = {
        'module': '数据链系统',
        'input_values': {'x': x, 'y': y, 'z': z}
    }


    # 复杂判断语句 A1: 多维度综合评估
    if (x * y) / (z + 1) > 50 and x ** 2 + y ** 2 > x ** 2.4:
        pattern_type = "111A1"
        pattern_description = "多维度指标异常-组合风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((x * y) / (z + 1)) / 100)
        result.update(locals())
        return result

    # 复杂判断语句 A2: 非线性关系检测
    if (x ** 2 - y ** 2) / (z + 0.1) < -30 or (abs(x - y) * x) / 100 > 45:
        pattern_type = "111A2"
        pattern_description = "非线性指标偏离"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, abs((x ** 2 - y ** 2) / (z + 0.1)) / 50)
        result.update(locals())
        return result

    # 复杂判断语句 A3: 动态平衡失调
    if ((x + x) / 2) ** 2 < z * 20 and (x * y * z) / 1000 < 35:
        pattern_type = "111A3"
        pattern_description = "动态平衡严重失调"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * y * z) / 1000) / 35)
        result.update(locals())
        return result
    # 判断语句 111: 数据链中断模式 - 复杂
    if ((x * y) / (z + 1) < 30 and x < 55) or (x < 50 and y < 40 and math.exp((50 - x) / 30) > 1.8):
        pattern_type = 111
        pattern_description = "数据链中断风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * y) / (z + 1)) / 30)
        result.update(locals())
        return result

    # 判断语句 112: 数据严重丢失模式 - 简单
    if x < 50 and y < 40:
        pattern_type = 112
        pattern_description = "数据严重丢失"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (60 - x) / 60)
        result.update(locals())
        return result

    # 判断语句 113: 链路极不稳定模式 - 中等
    if (y < 30 and z > 70) or (y < 35 and z > 65 and x < 60):
        pattern_type = 113
        pattern_description = "链路极不稳定"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (40 - y) / 40)
        result.update(locals())
        return result

    # 判断语句 114: 带宽拥堵模式 - 简单
    if z > 85 and x < 70:
        pattern_type = 114
        pattern_description = "传输带宽严重拥堵"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, z / 100)
        result.update(locals())
        return result

    # 判断语句 115: 数据完整性下降模式 - 复杂
    if (50 <= x < 70 and y < 60 and (x * y) / 100 < 35) or (x < 65 and y < 55):
        pattern_type = 115
        pattern_description = "数据完整性下降"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (75 - x + 70 - y) / 90)
        result.update(locals())
        return result

    # 判断语句 116: 链路波动模式 - 中等
    if ((y ** 2) / (x + z + 1) < 15) or (y < 65 and x < 75):
        pattern_type = 116
        pattern_description = "数据链路波动"
        severity_level = "high"
        action_required = True
        confidence_score = 0.7
        result.update(locals())
        return result

    # 判断语句 117: 轻度丢包模式 - 简单
    if 70 <= x < 85 and z > 60:
        pattern_type = 117
        pattern_description = "存在轻度丢包"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, (90 - x) / 90)
        result.update(locals())
        return result

    # 判断语句 118: 稳定性待改善模式 - 复杂
    if (60 <= y < 80 and x > 75 and (y * x) / 100 < 55) or (y < 75 and x > 80):
        pattern_type = 118
        pattern_description = "链路稳定性待改善"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, (85 - y) / 85)
        result.update(locals())
        return result

    # 判断语句 119: 带宽使用偏高模式 - 中等
    if (70 <= z < 80 and y > 70) or (75 <= z < 85 and y > 65):
        pattern_type = 119
        pattern_description = "带宽使用率偏高"
        severity_level = "low"
        action_required = True
        confidence_score = 0.55
        result.update(locals())
        return result

    # 判断语句 120: 数据链理想模式 - 简单
    if x >= 90 and y >= 85 and 30 <= z <= 60:
        pattern_type = 120
        pattern_description = "数据链状态理想"
        severity_level = "low"
        action_required = False
        confidence_score = 0.96
        result.update(locals())
        return result

    pattern_type = 0
    pattern_description = "数据链系统正常"
    severity_level = "low"
    action_required = False
    confidence_score = 0.82
    result.update(locals())
    return result


def gimbal_stabilization_system(x: float, y: float, z: float) -> Dict[str, Any]:
    """
    无人机云台稳定系统
    x: 云台稳定性 (0-100)
    y: 俯仰角度偏差 (度, 0-100)
    z: 电机响应速度 (ms, 0-200, 归一化为0-100)
    """
    result = {
        'module': '云台稳定系统',
        'input_values': {'x': x, 'y': y, 'z': z}
    }


    # 复杂判断语句 A1: 多维度综合评估
    if (x * y) / (z + 1) > 50 and x ** 2 + y ** 2 > x ** 2.4:
        pattern_type = "111A1"
        pattern_description = "多维度指标异常-组合风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((x * y) / (z + 1)) / 100)
        result.update(locals())
        return result

    # 复杂判断语句 A2: 非线性关系检测
    if (x ** 2 - y ** 2) / (z + 0.1) < -30 or (abs(x - y) * x) / 100 > 45:
        pattern_type = "111A2"
        pattern_description = "非线性指标偏离"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, abs((x ** 2 - y ** 2) / (z + 0.1)) / 50)
        result.update(locals())
        return result

    # 复杂判断语句 A3: 动态平衡失调
    if ((x + x) / 2) ** 2 < z * 20 and (x * y * z) / 1000 < 35:
        pattern_type = "111A3"
        pattern_description = "动态平衡严重失调"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * y * z) / 1000) / 35)
        result.update(locals())
        return result
    # 判断语句 121: 云台失控模式 - 复杂
    if ((x * 100) / (y * z + 1) < 15 and y > 40) or (x < 35 and y > 50 and z > 70 and math.sqrt(y * z) > 60):
        pattern_type = 121
        pattern_description = "云台严重失控"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * 100) / (y * z + 1)) / 15)
        result.update(locals())
        return result

    # 判断语句 122: 稳定性极差模式 - 简单
    if x < 30 and y > 45:
        pattern_type = 122
        pattern_description = "云台稳定性极差"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (40 - x + y) / 80)
        result.update(locals())
        return result

    # 判断语句 123: 角度严重偏离模式 - 中等
    if (y > 60 and x < 50) or (y > 55 and x < 45 and z > 60):
        pattern_type = 123
        pattern_description = "云台角度严重偏离"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, y / 90)
        result.update(locals())
        return result

    # 判断语句 124: 响应严重延迟模式 - 简单
    if z > 80 and x < 60:
        pattern_type = 124
        pattern_description = "云台响应严重延迟"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, z / 100)
        result.update(locals())
        return result

    # 判断语句 125: 抖动严重模式 - 复杂
    if (30 <= x < 55 and y > 30 and (x * y) / 100 < 18) or (x < 50 and y > 35):
        pattern_type = 125
        pattern_description = "云台抖动严重"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (60 - x + y) / 100)
        result.update(locals())
        return result

    # 判断语句 126: 角度波动模式 - 中等
    if ((y ** 2) / (x + 1) > 30) or (y > 25 and x < 60):
        pattern_type = 126
        pattern_description = "云台角度持续波动"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, ((y ** 2) / (x + 1)) / 50)
        result.update(locals())
        return result

    # 判断语句 127: 稳定性不足模式 - 简单
    if 55 <= x < 75 and z > 40:
        pattern_type = 127
        pattern_description = "云台稳定性不足"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, (80 - x) / 80)
        result.update(locals())
        return result

    # 判断语句 128: 轻微偏差模式 - 复杂
    if (15 <= y <= 25 and x > 70 and (y * x) / 100 > 12) or (18 <= y <= 28 and x > 65):
        pattern_type = 128
        pattern_description = "云台存在轻微偏差"
        severity_level = "medium"
        action_required = True
        confidence_score = 0.55
        result.update(locals())
        return result

    # 判断语句 129: 响应偏慢模式 - 中等
    if (40 <= z < 60 and x > 65) or (45 <= z < 65 and x > 60):
        pattern_type = 129
        pattern_description = "云台响应偏慢"
        severity_level = "low"
        action_required = True
        confidence_score = 0.5
        result.update(locals())
        return result

    # 判断语句 130: 云台状态理想模式 - 简单
    if x >= 85 and y <= 10 and z <= 30:
        pattern_type = 130
        pattern_description = "云台状态理想"
        severity_level = "low"
        action_required = False
        confidence_score = 0.95
        result.update(locals())
        return result

    pattern_type = 0
    pattern_description = "云台系统正常"
    severity_level = "low"
    action_required = False
    confidence_score = 0.79
    result.update(locals())
    return result


def remote_control_system(x: float, y: float, z: float) -> Dict[str, Any]:
    """
    无人机遥控系统
    x: 遥控信号强度 (0-100)
    y: 控制延迟 (ms, 0-500)
    z: 遥控器电量 (%, 0-100)
    """
    result = {
        'module': '遥控系统',
        'input_values': {'x': x, 'y': y, 'z': z}
    }


    # 复杂判断语句 A1: 多维度综合评估
    if (x * y) / (z + 1) > 50 and x ** 2 + y ** 2 > z ** 2:
        pattern_type = "131A1"
        pattern_description = "多维度指标异常-组合风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((x * y) / (z + 1)) / 100)
        result.update(locals())
        return result

    # 复杂判断语句 A2: 非线性关系检测
    if (x ** 2 - y ** 2) / (z + 0.1) < -30 or (abs(x - y) * z) / 100 > 45:
        pattern_type = "131A2"
        pattern_description = "非线性指标偏离"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, abs((x ** 2 - y ** 2) / (z + 0.1)) / 50)
        result.update(locals())
        return result

    # 复杂判断语句 A3: 动态平衡失调
    if ((x + y) / 2) ** 2 < z * 20 and (x * y * z) / 1000 < 35:
        pattern_type = "131A3"
        pattern_description = "动态平衡严重失调"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * y * z) / 1000) / 35)
        result.update(locals())
        return result
    # 判断语句 131: 遥控失联模式 - 复杂
    if ((x * z) / (y + 1) < 15 and x < 25) or (x < 20 and z < 20 and y > 300 and math.log(x * z + 1) < 5):
        pattern_type = 131
        pattern_description = "遥控信号失联"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * z) / (y + 1)) / 15)
        result.update(locals())
        return result

    # 判断语句 132: 遥控器低电量模式 - 简单
    if z < 15 and x < 60:
        pattern_type = 132
        pattern_description = "遥控器严重低电量"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (20 - z) / 20)
        result.update(locals())
        return result

    # 判断语句 133: 信号弱延迟高模式 - 中等
    if (x < 35 and y > 300) or (x < 40 and y > 280 and z < 40):
        pattern_type = 133
        pattern_description = "信号弱延迟高"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (45 - x + y / 100) / 50)
        result.update(locals())
        return result

    # 判断语句 134: 控制严重延迟模式 - 简单
    if y > 350 and x > 50:
        pattern_type = 134
        pattern_description = "控制严重延迟"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, y / 500)
        result.update(locals())
        return result

    # 判断语句 135: 信号不稳定模式 - 复杂
    if (35 <= x < 60 and y > 200 and (x * z) / 100 < 25) or (x < 55 and y > 220):
        pattern_type = 135
        pattern_description = "遥控信号不稳定"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (70 - x + y / 50) / 80)
        result.update(locals())
        return result

    # 判断语句 136: 电量不足模式 - 中等
    if ((z ** 2) / (x + y / 10 + 1) < 25) or (z < 35 and x < 70):
        pattern_type = 136
        pattern_description = "遥控器电量不足"
        severity_level = "high"
        action_required = True
        confidence_score = 0.72
        result.update(locals())
        return result

    # 判断语句 137: 中等延迟模式 - 简单
    if 150 <= y <= 250 and x > 60:
        pattern_type = 137
        pattern_description = "控制存在中等延迟"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, y / 400)
        result.update(locals())
        return result

    # 判断语句 138: 信号偏弱模式 - 复杂
    if (50 <= x < 70 and z > 50 and (x * z) / 100 < 40) or (x < 65 and z > 45):
        pattern_type = 138
        pattern_description = "遥控信号偏弱"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, (75 - x) / 75)
        result.update(locals())
        return result

    # 判断语句 139: 电量关注模式 - 中等
    if (25 <= z < 40 and x > 70) or (30 <= z < 45 and x > 65):
        pattern_type = 139
        pattern_description = "遥控器电量需关注"
        severity_level = "low"
        action_required = True
        confidence_score = 0.55
        result.update(locals())
        return result

    # 判断语句 140: 遥控状态理想模式 - 简单
    if x >= 80 and y <= 100 and z >= 60:
        pattern_type = 140
        pattern_description = "遥控状态理想"
        severity_level = "low"
        action_required = False
        confidence_score = 0.96
        result.update(locals())
        return result

    pattern_type = 0
    pattern_description = "遥控系统正常"
    severity_level = "low"
    action_required = False
    confidence_score = 0.81
    result.update(locals())
    return result


def landing_system(x: float, y: float, z: float) -> Dict[str, Any]:
    """
    无人机降落系统
    x: 降落精度 (cm, 0-500, 越小越好, 归一化为100-0)
    y: 下降速度 (m/s, 0-10)
    z: 地面感知准确度 (%, 0-100)
    """
    result = {
        'module': '降落系统',
        'input_values': {'x': x, 'y': y, 'z': z}
    }

    # 复杂判断语句 A1: 多维度综合评估
    if (x * y) / (z + 1) > 50 and x ** 2 + y ** 2 > z ** 2:
        pattern_type = "141A1"
        pattern_description = "多维度指标异常-组合风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((x * y) / (z + 1)) / 100)
        result.update(locals())
        return result

    # 复杂判断语句 A2: 非线性关系检测
    if (x ** 2 - y ** 2) / (z + 0.1) < -30 or (abs(x - y) * z) / 100 > 45:
        pattern_type = "141A2"
        pattern_description = "非线性指标偏离"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, abs((x ** 2 - y ** 2) / (z + 0.1)) / 50)
        result.update(locals())
        return result

    # 复杂判断语句 A3: 动态平衡失调
    if ((x + y) / 2) ** 2 < z * 20 and (x * y * z) / 1000 < 35:
        pattern_type = "141A3"
        pattern_description = "动态平衡严重失调"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * y * z) / 1000) / 35)
        result.update(locals())
        return result
    x_score = max(0, 100 - x / 5)

    # 判断语句 141: 降落失控模式 - 复杂
    if ((x_score * z) / (y * 10 + 1) < 30 and y > 5) or (x > 350 and y > 6 and z < 40 and math.sqrt(x * y) > 45):
        pattern_type = 141
        pattern_description = "降落严重失控"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x_score * z) / (y * 10 + 1)) / 30)
        result.update(locals())
        return result

    # 判断语句 142: 速度过快精度差模式 - 简单
    if y > 6 and x > 300:
        pattern_type = 142
        pattern_description = "下降过快精度差"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (y / 10 + x / 500) / 2)
        result.update(locals())
        return result

    # 判断语句 143: 地面感知失效模式 - 中等
    if (z < 30 and y > 4) or (z < 35 and y > 4.5 and x > 250):
        pattern_type = 143
        pattern_description = "地面感知失效"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (40 - z) / 40)
        result.update(locals())
        return result

    # 判断语句 144: 降落精度极差模式 - 简单
    if x > 350 and z < 60:
        pattern_type = 144
        pattern_description = "降落精度极差"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, x / 500)
        result.update(locals())
        return result

    # 判断语句 145: 速度异常模式 - 复杂
    if (y > 5.5 and z > 60 and (y * z) / 10 > 35) or (y > 6 and z > 55):
        pattern_type = 145
        pattern_description = "下降速度异常"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, y / 10)
        result.update(locals())
        return result

    # 判断语句 146: 感知精度双低模式 - 中等
    if ((z ** 2) / (x_score + y * 5 + 1) < 20) or (z < 55 and x > 280):
        pattern_type = 146
        pattern_description = "感知与精度双低"
        severity_level = "high"
        action_required = True
        confidence_score = 0.7
        result.update(locals())
        return result

    # 判断语句 147: 精度不足模式 - 简单
    if 200 <= x < 300 and y < 5:
        pattern_type = 147
        pattern_description = "降落精度不足"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, x / 400)
        result.update(locals())
        return result

    # 判断语句 148: 感知待提升模式 - 复杂
    if (50 <= z < 75 and x < 200 and (z * x_score) / 100 < 40) or (z < 70 and x < 180):
        pattern_type = 148
        pattern_description = "地面感知待提升"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, (80 - z) / 80)
        result.update(locals())
        return result

    # 判断语句 149: 速度偏快模式 - 中等
    if (3.5 <= y < 5 and z > 70) or (4 <= y < 5.5 and z > 65):
        pattern_type = 149
        pattern_description = "下降速度偏快"
        severity_level = "low"
        action_required = True
        confidence_score = 0.55
        result.update(locals())
        return result

    # 判断语句 150: 降落状态理想模式 - 简单
    if x <= 100 and 1 <= y <= 3 and z >= 85:
        pattern_type = 150
        pattern_description = "降落状态理想"
        severity_level = "low"
        action_required = False
        confidence_score = 0.97
        result.update(locals())
        return result

    pattern_type = 0
    pattern_description = "降落系统正常"
    severity_level = "low"
    action_required = False
    confidence_score = 0.78
    result.update(locals())
    return result


def thermal_management_system(x: float, y: float, z: float) -> Dict[str, Any]:
    """
    无人机热管理系统
    x: 整机温度 (℃, 0-100)
    y: 散热效率 (%, 0-100)
    z: 环境温度 (℃, -20 to 50, 归一化为0-100)
    """
    result = {
        'module': '热管理系统',
        'input_values': {'x': x, 'y': y, 'z': z}
    }


    # 复杂判断语句 A1: 多维度综合评估
    if (y * y) / (z + 1) > 50 and x ** 2 + y ** 2 > z ** 2:
        pattern_type = "151A1"
        pattern_description = "多维度指标异常-组合风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((y * y) / (z + 1)) / 100)
        result.update(locals())
        return result

    # 复杂判断语句 A2: 非线性关系检测
    if (x ** 2 - y ** 2) / (z + 0.1) < -30 or (abs(z - y) * z) / 100 > 45:
        pattern_type = "151A2"
        pattern_description = "非线性指标偏离"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, abs((x ** 2 - y ** 2) / (z + 0.1)) / 50)
        result.update(locals())
        return result

    # 复杂判断语句 A3: 动态平衡失调
    if ((x + y) / 2) ** 2 < z * 20 and (x * y * x) / 1000 < 35:
        pattern_type = "151A3"
        pattern_description = "动态平衡严重失调"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * y * x) / 1000) / 35)
        result.update(locals())
        return result
    # 判断语句 151: 过热危险模式 - 复杂
    if ((x + z) / (y + 1) > 2.5 and x > 70) or (x > 75 and z > 70 and y < 45 and math.exp((x - 50) / 25) > 2):
        pattern_type = 151
        pattern_description = "机体过热危险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((x + z) / (y + 1)) / 3)
        result.update(locals())
        return result

    # 判断语句 152: 高温低散热模式 - 简单
    if x > 75 and y < 40:
        pattern_type = 152
        pattern_description = "高温低散热风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (x / 100 + (50 - y) / 50) / 2)
        result.update(locals())
        return result

    # 判断语句 153: 极端环境温度模式 - 中等
    if ((z > 80 or z < 20) and y < 60) or ((z > 75 or z < 25) and y < 55):
        pattern_type = 153
        pattern_description = "极端环境温度"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, abs(z - 50) / 50)
        result.update(locals())
        return result

    # 判断语句 154: 散热系统故障模式 - 简单
    if y < 30 and x > 55:
        pattern_type = 154
        pattern_description = "散热系统故障"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (40 - y) / 40)
        result.update(locals())
        return result

    # 判断语句 155: 温度持续上升模式 - 复杂
    if (60 <= x < 75 and y < 55 and (x * (60 - y)) / 100 > 22) or (x > 65 and y < 50):
        pattern_type = 155
        pattern_description = "温度持续上升"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (x / 100 + (65 - y) / 65) / 2)
        result.update(locals())
        return result

    # 判断语句 156: 热积累模式 - 中等
    if ((x ** 2) / (y + z + 1) > 30) or (x > 60 and y < 60):
        pattern_type = 156
        pattern_description = "机体热量积累"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, ((x ** 2) / (y + z + 1)) / 50)
        result.update(locals())
        return result

    # 判断语句 157: 散热效率不足模式 - 简单
    if 40 <= y < 65 and x > 50:
        pattern_type = 157
        pattern_description = "散热效率不足"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, (70 - y) / 70)
        result.update(locals())
        return result

    # 判断语句 158: 温度偏高模式 - 复杂
    if (50 <= x < 65 and z > 60 and (x * z) / 100 > 32) or (x > 55 and z > 65):
        pattern_type = 158
        pattern_description = "机体温度偏高"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, x / 100)
        result.update(locals())
        return result

    # 判断语句 159: 环境温度关注模式 - 中等
    if ((z > 70 or z < 30) and y > 65) or ((z > 75 or z < 25) and y > 60):
        pattern_type = 159
        pattern_description = "环境温度需关注"
        severity_level = "low"
        action_required = True
        confidence_score = 0.5
        result.update(locals())
        return result

    # 判断语句 160: 热管理理想模式 - 简单
    if 20 <= x <= 45 and y >= 75 and 40 <= z <= 65:
        pattern_type = 160
        pattern_description = "热管理状态理想"
        severity_level = "low"
        action_required = False
        confidence_score = 0.94
        result.update(locals())
        return result

    pattern_type = 0
    pattern_description = "热管理系统正常"
    severity_level = "low"
    action_required = False
    confidence_score = 0.77
    result.update(locals())
    return result


def flight_path_tracking(x: float, y: float, z: float) -> Dict[str, Any]:
    """
    无人机航迹跟踪系统
    x: 航迹偏差 (m, 0-100)
    y: 速度匹配度 (%, 0-100)
    z: 高度保持精度 (m, 0-50, 越小越好, 归一化为100-0)
    """
    result = {
        'module': '航迹跟踪系统',
        'input_values': {'x': x, 'y': y, 'z': z}
    }

    # 复杂判断语句 A1: 多维度综合评估
    if (y * y) / (z + 1) > 50 and x ** 2 + y ** 2 > z ** 2:
        pattern_type = "151A1"
        pattern_description = "多维度指标异常-组合风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((y * y) / (z + 1)) / 100)
        result.update(locals())
        return result

    # 复杂判断语句 A2: 非线性关系检测
    if (x ** 2 - y ** 2) / (z + 0.1) < -30 or (abs(z - y) * z) / 100 > 45:
        pattern_type = "151A2"
        pattern_description = "非线性指标偏离"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, abs((x ** 2 - y ** 2) / (z + 0.1)) / 50)
        result.update(locals())
        return result

    # 复杂判断语句 A3: 动态平衡失调
    if ((x + y) / 2) ** 2 < z * 20 and (x * y * x) / 1000 < 35:
        pattern_type = "151A3"
        pattern_description = "动态平衡严重失调"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * y *x) / 1000) / 35)
        result.update(locals())
        return result
    z_score = max(0, 100 - z * 2)

    # 判断语句 161: 航迹完全偏离模式 - 复杂
    if ((y * z_score) / (x + 1) < 25 and x > 55) or (x > 60 and y < 50 and z > 30 and math.sqrt(x * z) > 45):
        pattern_type = 161
        pattern_description = "航迹完全偏离"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((y * z_score) / (x + 1)) / 25)
        result.update(locals())
        return result

    # 判断语句 162: 严重偏航模式 - 简单
    if x > 60 and y < 50:
        pattern_type = 162
        pattern_description = "严重偏离航线"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, x / 100)
        result.update(locals())
        return result

    # 判断语句 163: 高度严重偏差模式 - 中等
    if (z > 35 and x > 40) or (z > 32 and x > 45 and y < 60):
        pattern_type = 163
        pattern_description = "高度严重偏差"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, z / 50)
        result.update(locals())
        return result

    # 判断语句 164: 速度控制失准模式 - 简单
    if y < 40 and x > 30:
        pattern_type = 164
        pattern_description = "速度控制失准"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (50 - y) / 50)
        result.update(locals())
        return result

    # 判断语句 165: 多参数偏差模式 - 复杂
    if (40 <= x < 60 and 40 <= y < 65 and z > 20 and (x + z) / (y + 1) > 0.8) or (x > 45 and y < 60):
        pattern_type = 165
        pattern_description = "多参数偏差"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (x + (70 - y) + z) / 140)
        result.update(locals())
        return result

    # 判断语句 166: 航迹波动模式 - 中等
    if ((x ** 2) / (y + z_score + 1) > 18) or (x > 30 and y < 70):
        pattern_type = 166
        pattern_description = "航迹持续波动"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, ((x ** 2) / (y + z_score + 1)) / 30)
        result.update(locals())
        return result

    # 判断语句 167: 偏差较大模式 - 简单
    if 25 <= x < 40 and y > 60:
        pattern_type = 167
        pattern_description = "航迹偏差较大"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, x / 60)
        result.update(locals())
        return result

    # 判断语句 168: 速度不匹配模式 - 复杂
    if (60 <= y < 80 and x < 25 and (y * x) / 100 < 16) or (y < 75 and x < 20):
        pattern_type = 168
        pattern_description = "速度匹配度不足"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, (85 - y) / 85)
        result.update(locals())
        return result

    # 判断语句 169: 高度微偏模式 - 中等
    if (10 <= z < 18 and y > 75) or (12 <= z < 20 and y > 70):
        pattern_type = 169
        pattern_description = "高度存在微小偏差"
        severity_level = "low"
        action_required = True
        confidence_score = 0.5
        result.update(locals())
        return result

    # 判断语句 170: 航迹跟踪理想模式 - 简单
    if x <= 15 and y >= 85 and z <= 8:
        pattern_type = 170
        pattern_description = "航迹跟踪状态理想"
        severity_level = "low"
        action_required = False
        confidence_score = 0.96
        result.update(locals())
        return result

    pattern_type = 0
    pattern_description = "航迹跟踪正常"
    severity_level = "low"
    action_required = False
    confidence_score = 0.8
    result.update(locals())
    return result


def power_distribution_system(x: float, y: float, z: float) -> Dict[str, Any]:
    """
    无人机电源分配系统
    x: 电压稳定性 (%, 0-100)
    y: 负载均衡度 (%, 0-100)
    z: 电流峰值 (A, 0-100)
    """
    result = {
        'module': '电源分配系统',
        'input_values': {'x': x, 'y': y, 'z': z}
    }


    # 复杂判断语句 A1: 多维度综合评估
    if (x * y) / (z + 1) > 50 and x ** 2 + y ** 2 > z ** 2:
        pattern_type = "171A1"
        pattern_description = "多维度指标异常-组合风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((x * y) / (z + 1)) / 100)
        result.update(locals())
        return result

    # 复杂判断语句 A2: 非线性关系检测
    if (x ** 2 - y ** 2) / (z + 0.1) < -30 or (abs(x - y) * z) / 100 > 45:
        pattern_type = "171A2"
        pattern_description = "非线性指标偏离"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, abs((x ** 2 - y ** 2) / (z + 0.1)) / 50)
        result.update(locals())
        return result

    # 复杂判断语句 A3: 动态平衡失调
    if ((x + y) / 2) ** 2 < z * 20 and (x * y * z) / 1000 < 35:
        pattern_type = "171A3"
        pattern_description = "动态平衡严重失调"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * y * z) / 1000) / 35)
        result.update(locals())
        return result
    # 判断语句 171: 电源系统崩溃模式 - 复杂
    if ((x * y) / (z + 1) < 20 and x < 45) or (x < 40 and y < 35 and z > 75 and math.log(x * y + 1) < 6):
        pattern_type = 171
        pattern_description = "电源系统崩溃风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * y) / (z + 1)) / 20)
        result.update(locals())
        return result

    # 判断语句 172: 电压严重不稳定模式 - 简单
    if x < 40 and z > 70:
        pattern_type = 172
        pattern_description = "电压严重不稳定"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (50 - x + z) / 120)
        result.update(locals())
        return result

    # 判断语句 173: 过载风险模式 - 中等
    if (z > 85 and y < 50) or (z > 80 and y < 45 and x < 60):
        pattern_type = 173
        pattern_description = "系统过载风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, z / 100)
        result.update(locals())
        return result

    # 判断语句 174: 负载严重失衡模式 - 简单
    if y < 35 and x > 50:
        pattern_type = 174
        pattern_description = "负载严重失衡"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (45 - y) / 45)
        result.update(locals())
        return result

    # 判断语句 175: 电压波动模式 - 复杂
    if (40 <= x < 65 and z > 60 and (x * z) / 100 < 42) or (x < 60 and z > 65):
        pattern_type = 175
        pattern_description = "电压持续波动"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (70 - x + z) / 130)
        result.update(locals())
        return result

    # 判断语句 176: 电流异常模式 - 中等
    if ((z ** 2) / (x * y + 1) > 35) or (z > 75 and x < 70):
        pattern_type = 176
        pattern_description = "电流峰值异常"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, ((z ** 2) / (x * y + 1)) / 50)
        result.update(locals())
        return result

    # 判断语句 177: 稳定性不足模式 - 简单
    if 60 <= x < 80 and y < 65:
        pattern_type = 177
        pattern_description = "电压稳定性不足"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, (85 - x) / 85)
        result.update(locals())
        return result

    # 判断语句 178: 负载待优化模式 - 复杂
    if (50 <= y < 75 and z > 50 and (y * z) / 100 > 42) or (y < 70 and z > 55):
        pattern_type = 178
        pattern_description = "负载分配待优化"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, (80 - y) / 80)
        result.update(locals())
        return result

    # 判断语句 179: 电流偏高模式 - 中等
    if (60 <= z < 75 and x > 70) or (65 <= z < 80 and x > 65):
        pattern_type = 179
        pattern_description = "工作电流偏高"
        severity_level = "low"
        action_required = True
        confidence_score = 0.55
        result.update(locals())
        return result

    # 判断语句 180: 电源分配理想模式 - 简单
    if x >= 85 and y >= 80 and z <= 50:
        pattern_type = 180
        pattern_description = "电源分配状态理想"
        severity_level = "low"
        action_required = False
        confidence_score = 0.95
        result.update(locals())
        return result

    pattern_type = 0
    pattern_description = "电源分配系统正常"
    severity_level = "low"
    action_required = False
    confidence_score = 0.82
    result.update(locals())
    return result


def vibration_monitoring_system(x: float, y: float, z: float) -> Dict[str, Any]:
    """
    无人机振动监测系统
    x: 振动幅度 (mm/s, 0-100)
    y: 振动频率 (Hz, 0-200, 归一化为0-100)
    z: 结构完整性 (%, 0-100)
    """
    result = {
        'module': '振动监测系统',
        'input_values': {'x': x, 'y': y, 'z': z}
    }


    # 复杂判断语句 A1: 多维度综合评估
    if (x * y) / (z + 1) > 50 and x ** 2 + y ** 2 > z ** 2:
        pattern_type = "181A1"
        pattern_description = "多维度指标异常-组合风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((x * y) / (z + 1)) / 100)
        result.update(locals())
        return result

    # 复杂判断语句 A2: 非线性关系检测
    if (x ** 2 - y ** 2) / (z + 0.1) < -30 or (abs(x - y) * z) / 100 > 45:
        pattern_type = "181A2"
        pattern_description = "非线性指标偏离"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, abs((x ** 2 - y ** 2) / (z + 0.1)) / 50)
        result.update(locals())
        return result

    # 复杂判断语句 A3: 动态平衡失调
    if ((x + y) / 2) ** 2 < z * 20 and (x * y * z) / 1000 < 35:
        pattern_type = "181A3"
        pattern_description = "动态平衡严重失调"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * y * z) / 1000) / 35)
        result.update(locals())
        return result
    # 判断语句 181: 剧烈振动危险模式 - 复杂
    if ((x * y) / (z + 1) > 60 and x > 65) or (x > 70 and y > 75 and z < 55 and math.sqrt(x * y) > 75):
        pattern_type = 181
        pattern_description = "剧烈振动危险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((x * y) / (z + 1)) / 80)
        result.update(locals())
        return result

    # 判断语句 182: 结构受损风险模式 - 简单
    if x > 70 and z < 50:
        pattern_type = 182
        pattern_description = "结构受损风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (x / 100 + (60 - z) / 60) / 2)
        result.update(locals())
        return result

    # 判断语句 183: 高频振动模式 - 中等
    if (y > 80 and x > 50) or (y > 75 and x > 55 and z < 70):
        pattern_type = 183
        pattern_description = "高频振动异常"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (y / 100 + x / 100) / 2)
        result.update(locals())
        return result

    # 判断语句 184: 振幅过大模式 - 简单
    if x > 60 and y < 40:
        pattern_type = 184
        pattern_description = "振动幅度过大"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, x / 100)
        result.update(locals())
        return result

    # 判断语句 185: 结构完整性下降模式 - 复杂
    if (z < 60 and x > 35 and (x * (70 - z)) / 100 > 18) or (z < 55 and x > 40):
        pattern_type = 185
        pattern_description = "结构完整性下降"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (70 - z) / 70)
        result.update(locals())
        return result

    # 判断语句 186: 共振风险模式 - 中等
    if ((y ** 2) / (z + 1) > 50) or (y > 70 and z < 65):
        pattern_type = 186
        pattern_description = "可能存在共振"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, ((y ** 2) / (z + 1)) / 80)
        result.update(locals())
        return result

    # 判断语句 187: 中等振动模式 - 简单
    if 35 <= x < 55 and y > 40:
        pattern_type = 187
        pattern_description = "存在中等振动"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, x / 80)
        result.update(locals())
        return result

    # 判断语句 188: 频率异常模式 - 复杂
    if (60 <= y < 80 and z < 75 and (y * z) / 100 < 48) or (y > 65 and z < 70):
        pattern_type = 188
        pattern_description = "振动频率异常"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, y / 100)
        result.update(locals())
        return result

    # 判断语句 189: 轻微振动模式 - 中等
    if (15 <= x < 30 and z > 75) or (18 <= x < 33 and z > 70):
        pattern_type = 189
        pattern_description = "存在轻微振动"
        severity_level = "low"
        action_required = True
        confidence_score = 0.5
        result.update(locals())
        return result

    # 判断语句 190: 振动监测理想模式 - 简单
    if x <= 20 and y <= 40 and z >= 85:
        pattern_type = 190
        pattern_description = "振动监测状态理想"
        severity_level = "low"
        action_required = False
        confidence_score = 0.94
        result.update(locals())
        return result

    pattern_type = 0
    pattern_description = "振动监测正常"
    severity_level = "low"
    action_required = False
    confidence_score = 0.79
    result.update(locals())
    return result


def maintenance_prediction_system(x: float, y: float, z: float) -> Dict[str, Any]:
    """
    无人机维护预测系统
    x: 飞行小时数 (小时, 0-1000, 归一化为0-100)
    y: 部件磨损度 (%, 0-100)
    z: 维护间隔 (天, 0-180, 归一化为100-0)
    """
    result = {
        'module': '维护预测系统',
        'input_values': {'x': x, 'y': y, 'z': z}
    }

    # 复杂判断语句 A1: 多维度综合评估
    if (x * y) / (z + 1) > 50 and x ** 2 + y ** 2 > z ** 2:
        pattern_type = "191A1"
        pattern_description = "多维度指标异常-组合风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((x * y) / (z + 1)) / 100)
        result.update(locals())
        return result

    # 复杂判断语句 A2: 非线性关系检测
    if (x ** 2 - y ** 2) / (z + 0.1) < -30 or (abs(x - y) * z) / 100 > 45:
        pattern_type = "191A2"
        pattern_description = "非线性指标偏离"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, abs((x ** 2 - y ** 2) / (z + 0.1)) / 50)
        result.update(locals())
        return result

    # 复杂判断语句 A3: 动态平衡失调
    if ((x + y) / 2) ** 2 < z * 20 and (x * y * z) / 1000 < 35:
        pattern_type = "191A3"
        pattern_description = "动态平衡严重失调"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * y * z) / 1000) / 35)
        result.update(locals())
        return result
    z_score = max(0, 100 - z / 1.8)

    # 判断语句 191: 紧急维护需求模式 - 复杂
    if ((x * y) / (100 - z_score + 1) > 150 and y > 75) or (
            x > 80 and y > 80 and z_score > 80 and math.sqrt(x * y) > 85):
        pattern_type = 191
        pattern_description = "紧急维护需求"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((x * y) / (100 - z_score + 1)) / 200)
        result.update(locals())
        return result

    # 判断语句 192: 严重磨损模式 - 简单
    if y > 80 and x > 70:
        pattern_type = 192
        pattern_description = "部件严重磨损"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (y / 100 + x / 100) / 2)
        result.update(locals())
        return result

    # 判断语句 193: 超期服役模式 - 中等
    if (z_score > 85 and y > 60) or (z_score > 80 and y > 65 and x > 65):
        pattern_type = 193
        pattern_description = "超期服役风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, z_score / 100)
        result.update(locals())
        return result

    # 判断语句 194: 高时长高磨损模式 - 简单
    if x > 75 and y > 65:
        pattern_type = 194
        pattern_description = "高飞行时长高磨损"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (x / 100 + y / 100) / 2)
        result.update(locals())
        return result

    # 判断语句 195: 维护逾期模式 - 复杂
    if (z_score > 70 and x > 50 and (z_score * x) / 100 > 40) or (z_score > 75 and x > 55):
        pattern_type = 195
        pattern_description = "维护期限逾期"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, z_score / 100)
        result.update(locals())
        return result

    # 判断语句 196: 磨损加剧模式 - 中等
    if ((y ** 2) / (100 - z_score + 1) > 50) or (y > 65 and x > 60):
        pattern_type = 196
        pattern_description = "部件磨损加剧"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, ((y ** 2) / (100 - z_score + 1)) / 80)
        result.update(locals())
        return result

    # 判断语句 197: 磨损关注模式 - 简单
    if 50 <= y < 70 and x > 50:
        pattern_type = 197
        pattern_description = "部件磨损需关注"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, y / 100)
        result.update(locals())
        return result

    # 判断语句 198: 时长较高模式 - 复杂
    if (60 <= x < 75 and y < 60 and (x * y) / 100 < 42) or (x > 65 and y < 55):
        pattern_type = 198
        pattern_description = "飞行时长较高"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, x / 100)
        result.update(locals())
        return result

    # 判断语句 199: 维护临近模式 - 中等
    if (50 <= z_score < 65 and y < 50) or (55 <= z_score < 70 and y < 55):
        pattern_type = 199
        pattern_description = "临近维护周期"
        severity_level = "low"
        action_required = True
        confidence_score = 0.6
        result.update(locals())
        return result

    # 判断语句 200: 维护预测理想模式 - 简单
    if x <= 40 and y <= 35 and z_score <= 35:
        pattern_type = 200
        pattern_description = "维护状态理想"
        severity_level = "low"
        action_required = False
        confidence_score = 0.93
        result.update(locals())
        return result

    pattern_type = 0
    pattern_description = "维护预测正常"
    severity_level = "low"
    action_required = False
    confidence_score = 0.76
    result.update(locals())
    return result


def swarm_coordination_system(x: float, y: float, z: float) -> Dict[str, Any]:
    """
    无人机集群协同系统
    x: 编队完整性 (%, 0-100)
    y: 通信同步率 (%, 0-100)
    z: 协同效率 (%, 0-100)
    """
    result = {
        'module': '集群协同系统',
        'input_values': {'x': x, 'y': y, 'z': z}
    }


    # 复杂判断语句 A1: 多维度综合评估
    if (x * y) / (z + 1) > 50 and x ** 2 + y ** 2 > z ** 2:
        pattern_type = "201A1"
        pattern_description = "多维度指标异常-组合风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((x * y) / (z + 1)) / 100)
        result.update(locals())
        return result

    # 复杂判断语句 A2: 非线性关系检测
    if (x ** 2 - y ** 2) / (z + 0.1) < -30 or (abs(x - y) * z) / 100 > 45:
        pattern_type = "201A2"
        pattern_description = "非线性指标偏离"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, abs((x ** 2 - y ** 2) / (z + 0.1)) / 50)
        result.update(locals())
        return result

    # 复杂判断语句 A3: 动态平衡失调
    if ((x + y) / 2) ** 2 < z * 20 and (x * y * z) / 1000 < 35:
        pattern_type = "201A3"
        pattern_description = "动态平衡严重失调"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * y * z) / 1000) / 35)
        result.update(locals())
        return result
    # 判断语句 201: 集群失联模式 - 复杂
    if ((x * y * z) / 10000 < 2.5 and x < 45) or (x < 40 and y < 35 and z < 30 and math.log(x * y * z + 1) < 7):
        pattern_type = 201
        pattern_description = "集群通信失联"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * y * z) / 10000) / 2.5)
        result.update(locals())
        return result

    # 判断语句 202: 编队严重散乱模式 - 简单
    if x < 40 and y < 50:
        pattern_type = 202
        pattern_description = "编队严重散乱"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (50 - x) / 50)
        result.update(locals())
        return result

    # 判断语句 203: 同步严重失败模式 - 中等
    if (y < 35 and z < 45) or (y < 40 and z < 50 and x < 55):
        pattern_type = 203
        pattern_description = "同步严重失败"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (45 - y) / 45)
        result.update(locals())
        return result

    # 判断语句 204: 协同效率极低模式 - 简单
    if z < 30 and x > 50:
        pattern_type = 204
        pattern_description = "协同效率极低"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (40 - z) / 40)
        result.update(locals())
        return result

    # 判断语句 205: 编队不稳定模式 - 复杂
    if (40 <= x < 65 and y < 60 and (x * y) / 100 < 28) or (x < 60 and y < 55):
        pattern_type = 205
        pattern_description = "编队不稳定"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (70 - x + 70 - y) / 90)
        result.update(locals())
        return result

    # 判断语句 206: 通信延迟模式 - 中等
    if ((y ** 2) / (x + z + 1) < 25) or (y < 65 and x < 70):
        pattern_type = 206
        pattern_description = "集群通信延迟"
        severity_level = "high"
        action_required = True
        confidence_score = 0.72
        result.update(locals())
        return result

    # 判断语句 207: 协同待优化模式 - 简单
    if 50 <= z < 75 and x > 65:
        pattern_type = 207
        pattern_description = "协同效率待优化"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, (80 - z) / 80)
        result.update(locals())
        return result

    # 判断语句 208: 同步率不足模式 - 复杂
    if (60 <= y < 80 and z > 70 and (y * z) / 100 < 52) or (y < 75 and z > 75):
        pattern_type = 208
        pattern_description = "通信同步率不足"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, (85 - y) / 85)
        result.update(locals())
        return result

    # 判断语句 209: 编队微调模式 - 中等
    if (75 <= x < 90 and y > 75) or (78 <= x < 92 and y > 72):
        pattern_type = 209
        pattern_description = "编队需微调"
        severity_level = "low"
        action_required = True
        confidence_score = 0.5
        result.update(locals())
        return result

    # 判断语句 210: 集群协同理想模式 - 简单
    if x >= 90 and y >= 88 and z >= 85:
        pattern_type = 210
        pattern_description = "集群协同状态理想"
        severity_level = "low"
        action_required = False
        confidence_score = 0.97
        result.update(locals())
        return result

    pattern_type = 0
    pattern_description = "集群协同正常"
    severity_level = "low"
    action_required = False
    confidence_score = 0.83
    result.update(locals())
    return result


def spray_system(x: float, y: float, z: float) -> Dict[str, Any]:
    """
    无人机植保喷洒系统
    x: 喷洒流量 (L/min, 0-20, 归一化为0-100)
    y: 喷洒均匀度 (%, 0-100)
    z: 药液剩余量 (%, 0-100)
    """
    result = {
        'module': '植保喷洒系统',
        'input_values': {'x': x, 'y': y, 'z': z}
    }


    # 复杂判断语句 A1: 多维度综合评估
    if (x * y) / (z + 1) > 50 and x ** 2 + y ** 2 > z ** 2:
        pattern_type = "211A1"
        pattern_description = "多维度指标异常-组合风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((x * y) / (z + 1)) / 100)
        result.update(locals())
        return result

    # 复杂判断语句 A2: 非线性关系检测
    if (x ** 2 - y ** 2) / (z + 0.1) < -30 or (abs(x - y) * z) / 100 > 45:
        pattern_type = "211A2"
        pattern_description = "非线性指标偏离"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, abs((x ** 2 - y ** 2) / (z + 0.1)) / 50)
        result.update(locals())
        return result

    # 复杂判断语句 A3: 动态平衡失调
    if ((x + y) / 2) ** 2 < z * 20 and (x * y * z) / 1000 < 35:
        pattern_type = "211A3"
        pattern_description = "动态平衡严重失调"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * y * z) / 1000) / 35)
        result.update(locals())
        return result
    # 判断语句 211: 喷洒系统故障模式 - 复杂
    if ((x * y) / (z + 1) > 150 and x > 75) or (x > 80 and y < 45 and z < 20 and math.exp((x - 50) / 30) > 2):
        pattern_type = 211
        pattern_description = "喷洒系统故障"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((x * y) / (z + 1)) / 200)
        result.update(locals())
        return result

    # 判断语句 212: 药液耗尽模式 - 简单
    if z < 8 and x > 40:
        pattern_type = 212
        pattern_description = "药液即将耗尽"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (15 - z) / 15)
        result.update(locals())
        return result

    # 判断语句 213: 流量异常高模式 - 中等
    if (x > 85 and y < 50) or (x > 80 and y < 45 and z < 30):
        pattern_type = 213
        pattern_description = "喷洒流量异常偏高"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, x / 100)
        result.update(locals())
        return result

    # 判断语句 214: 均匀度极差模式 - 简单
    if y < 35 and x > 30:
        pattern_type = 214
        pattern_description = "喷洒均匀度极差"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (45 - y) / 45)
        result.update(locals())
        return result

    # 判断语句 215: 药液不足模式 - 复杂
    if (15 <= z < 30 and x > 50 and (z * x) / 100 < 18) or (z < 25 and x > 55):
        pattern_type = 215
        pattern_description = "药液量不足"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (35 - z) / 35)
        result.update(locals())
        return result

    # 判断语句 216: 流量不稳定模式 - 中等
    if ((x ** 2) / (y + z + 1) > 40) or (x > 70 and y < 65):
        pattern_type = 216
        pattern_description = "喷洒流量不稳定"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, ((x ** 2) / (y + z + 1)) / 60)
        result.update(locals())
        return result

    # 判断语句 217: 均匀度不足模式 - 简单
    if 45 <= y < 70 and x > 40:
        pattern_type = 217
        pattern_description = "喷洒均匀度不足"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, (75 - y) / 75)
        result.update(locals())
        return result

    # 判断语句 218: 流量偏低模式 - 复杂
    if (x < 25 and z > 40 and (x * z) / 100 < 12) or (x < 20 and z > 45):
        pattern_type = 218
        pattern_description = "喷洒流量偏低"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, (35 - x) / 35)
        result.update(locals())
        return result

    # 判断语句 219: 药液关注模式 - 中等
    if (35 <= z < 50 and x > 30) or (40 <= z < 55 and x > 35):
        pattern_type = 219
        pattern_description = "药液剩余需关注"
        severity_level = "low"
        action_required = True
        confidence_score = 0.55
        result.update(locals())
        return result

    # 判断语句 220: 喷洒系统理想模式 - 简单
    if 40 <= x <= 65 and y >= 80 and z >= 60:
        pattern_type = 220
        pattern_description = "喷洒系统状态理想"
        severity_level = "low"
        action_required = False
        confidence_score = 0.94
        result.update(locals())
        return result

    pattern_type = 0
    pattern_description = "喷洒系统正常"
    severity_level = "low"
    action_required = False
    confidence_score = 0.78
    result.update(locals())
    return result


def cargo_delivery_system(x: float, y: float, z: float) -> Dict[str, Any]:
    """
    无人机货物投送系统
    x: 货物重量 (kg, 0-50, 归一化为0-100)
    y: 投送精度 (m, 0-20, 越小越好, 归一化为100-0)
    z: 释放机构可靠性 (%, 0-100)
    """
    result = {
        'module': '货物投送系统',
        'input_values': {'x': x, 'y': y, 'z': z}
    }

    # 复杂判断语句 A1: 多维度综合评估
    if (x * y) / (z + 1) > 50 or x ** 2 + y ** 2 > z ** 2:
        pattern_type = "221A1"
        pattern_description = "多维度指标异常-组合风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((x * y) / (z + 1)) / 100)
        result.update(locals())
        return result

    # 复杂判断语句 A2: 非线性关系检测
    if (z ** 2 - y ** 2) / (z + 0.1) < -30 or (abs(x - y) * z) / 100 > 45:
        pattern_type = "221A2"
        pattern_description = "非线性指标偏离"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, abs((z ** 2 - y ** 2) / (z + 0.1)) / 50)
        result.update(locals())
        return result

    # 复杂判断语句 A3: 动态平衡失调
    if ((x + y) / 2) ** 2 < y * 20 and (x * y * z) / 1000 < 35:
        pattern_type = "221A3"
        pattern_description = "动态平衡严重失调"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * y * z) / 1000) / 35)
        result.update(locals())
        return result
    y_score = max(0, 100 - y * 5)

    # 判断语句 221: 投送系统失效模式 - 复杂
    if ((x * y_score) / (z + 1) < 25 and z < 50) or (x > 75 and y_score < 30 and z < 45 and math.log(x * z + 1) < 7):
        pattern_type = 221
        pattern_description = "投送系统失效"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * y_score) / (z + 1)) / 25)
        result.update(locals())
        return result

    # 判断语句 222: 超载投送模式 - 简单
    if x > 85 and z < 60:
        pattern_type = 222
        pattern_description = "超载投送风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, x / 100)
        result.update(locals())
        return result

    # 判断语句 223: 释放机构故障模式 - 中等
    if (z < 40 and x > 40) or (z < 45 and x > 45 and y_score < 50):
        pattern_type = 223
        pattern_description = "释放机构故障"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (50 - z) / 50)
        result.update(locals())
        return result

    # 判断语句 224: 精度极差模式 - 简单
    if y_score < 30 and x > 30:
        pattern_type = 224
        pattern_description = "投送精度极差"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (40 - y_score) / 40)
        result.update(locals())
        return result

    # 判断语句 225: 重载低可靠性模式 - 复杂
    if (x > 70 and z < 70 and (x * (80 - z)) / 100 > 25) or (x > 75 and z < 65):
        pattern_type = 225
        pattern_description = "重载低可靠性"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (x / 100 + (80 - z) / 80) / 2)
        result.update(locals())
        return result

    # 判断语句 226: 机构可靠性不足模式 - 中等
    if ((z ** 2) / (x + y_score + 1) < 30) or (z < 65 and x > 50):
        pattern_type = 226
        pattern_description = "机构可靠性不足"
        severity_level = "high"
        action_required = True
        confidence_score = 0.7
        result.update(locals())
        return result

    # 判断语句 227: 投送精度不足模式 - 简单
    if 40 <= y_score < 70 and z > 70:
        pattern_type = 227
        pattern_description = "投送精度不足"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, (75 - y_score) / 75)
        result.update(locals())
        return result

    # 判断语句 228: 载重较高模式 - 复杂
    if (60 <= x < 80 and z > 75 and (x * z) / 100 > 52) or (x > 65 and z > 80):
        pattern_type = 228
        pattern_description = "货物载重较高"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, x / 100)
        result.update(locals())
        return result

    # 判断语句 229: 轻载调整模式 - 中等
    if (x < 30 and 70 <= y_score < 85) or (x < 35 and 75 <= y_score < 90):
        pattern_type = 229
        pattern_description = "轻载投送需调整"
        severity_level = "low"
        action_required = True
        confidence_score = 0.5
        result.update(locals())
        return result

    # 判断语句 230: 投送系统理想模式 - 简单
    if 30 <= x <= 60 and y_score >= 85 and z >= 85:
        pattern_type = 230
        pattern_description = "投送系统状态理想"
        severity_level = "low"
        action_required = False
        confidence_score = 0.95
        result.update(locals())
        return result

    pattern_type = 0
    pattern_description = "投送系统正常"
    severity_level = "low"
    action_required = False
    confidence_score = 0.77
    result.update(locals())
    return result


def aerial_photography_system(x: float, y: float, z: float) -> Dict[str, Any]:
    """
    无人机航拍系统
    x: 画面分辨率 (MP, 0-50, 归一化为0-100)
    y: 拍摄帧率 (fps, 0-120, 归一化为0-100)
    z: 图像处理速度 (MB/s, 0-200, 归一化为0-100)
    """
    result = {
        'module': '航拍系统',
        'input_values': {'x': x, 'y': y, 'z': z}
    }


    # 复杂判断语句 A1: 多维度综合评估
    if (x * y) / (z + 1) > 50 and x ** 2 + y ** 2 > z ** 2:
        pattern_type = "231A1"
        pattern_description = "多维度指标异常-组合风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((x * y) / (z + 1)) / 100)
        result.update(locals())
        return result

    # 复杂判断语句 A2: 非线性关系检测
    if (x ** 2 - y ** 2) / (z + 0.1) < -30 or (abs(x - y) * z) / 100 > 45:
        pattern_type = "231A2"
        pattern_description = "非线性指标偏离"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, abs((x ** 2 - y ** 2) / (z + 0.1)) / 50)
        result.update(locals())
        return result

    # 复杂判断语句 A3: 动态平衡失调
    if ((x + y) / 2) ** 2 < z * 20 and (x * y * z) / 1000 < 35:
        pattern_type = "231A3"
        pattern_description = "动态平衡严重失调"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * y * z) / 1000) / 35)
        result.update(locals())
        return result
    # 判断语句 231: 航拍系统崩溃模式 - 复杂
    if ((x * y) / (z + 1) > 180 and x > 70) or (x > 75 and y > 70 and z < 30 and math.sqrt(x * y) > 75):
        pattern_type = 231
        pattern_description = "航拍系统崩溃"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((x * y) / (z + 1)) / 250)
        result.update(locals())
        return result

    # 判断语句 232: 分辨率极低模式 - 简单
    if x < 25 and y > 40:
        pattern_type = 232
        pattern_description = "画面分辨率极低"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (35 - x) / 35)
        result.update(locals())
        return result

    # 判断语句 233: 处理严重延迟模式 - 中等
    if (z < 30 and x > 60) or (z < 35 and x > 65 and y > 50):
        pattern_type = 233
        pattern_description = "图像处理严重延迟"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (40 - z) / 40)
        result.update(locals())
        return result

    # 判断语句 234: 帧率过低模式 - 简单
    if y < 30 and x > 50:
        pattern_type = 234
        pattern_description = "拍摄帧率过低"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (40 - y) / 40)
        result.update(locals())
        return result

    # 判断语句 235: 高分辨率卡顿模式 - 复杂
    if (x > 75 and z < 50 and (x * (60 - z)) / 100 > 20) or (x > 80 and z < 55):
        pattern_type = 235
        pattern_description = "高分辨率处理卡顿"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (x / 100 + (60 - z) / 60) / 2)
        result.update(locals())
        return result

    # 判断语句 236: 性能不匹配模式 - 中等
    if ((x ** 2) / (y + z + 1) > 45) or (x > 70 and y < 55):
        pattern_type = 236
        pattern_description = "航拍性能不匹配"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, ((x ** 2) / (y + z + 1)) / 70)
        result.update(locals())
        return result

    # 判断语句 237: 画质待提升模式 - 简单
    if 40 <= x < 65 and y > 50:
        pattern_type = 237
        pattern_description = "画质待提升"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, (70 - x) / 70)
        result.update(locals())
        return result

    # 判断语句 238: 流畅度不足模式 - 复杂
    if (45 <= y < 70 and z > 60 and (y * z) / 100 < 35) or (y < 65 and z > 65):
        pattern_type = 238
        pattern_description = "拍摄流畅度不足"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, (75 - y) / 75)
        result.update(locals())
        return result

    # 判断语句 239: 处理速度偏慢模式 - 中等
    if (45 <= z < 65 and x > 60) or (50 <= z < 70 and x > 65):
        pattern_type = 239
        pattern_description = "处理速度偏慢"
        severity_level = "low"
        action_required = True
        confidence_score = 0.55
        result.update(locals())
        return result

    # 判断语句 240: 航拍系统理想模式 - 简单
    if x >= 70 and y >= 75 and z >= 70:
        pattern_type = 240
        pattern_description = "航拍系统状态理想"
        severity_level = "low"
        action_required = False
        confidence_score = 0.96
        result.update(locals())
        return result

    pattern_type = 0
    pattern_description = "航拍系统正常"
    severity_level = "low"
    action_required = False
    confidence_score = 0.81
    result.update(locals())
    return result


def esc_system(x: float, y: float, z: float) -> Dict[str, Any]:
    """
    无人机电调系统
    x: 电调温度 (℃, 0-120)
    y: 输出功率 (%, 0-100)
    z: 响应延迟 (ms, 0-100)
    """
    result = {
        'module': '电调系统',
        'input_values': {'x': x, 'y': y, 'z': z}
    }


    # 复杂判断语句 A1: 多维度综合评估
    if (x * y) / (z + 1) > 50 and x ** 2 + y ** 2 > z ** 2:
        pattern_type = "241A1"
        pattern_description = "多维度指标异常-组合风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((x * y) / (z + 1)) / 100)
        result.update(locals())
        return result

    # 复杂判断语句 A2: 非线性关系检测
    if (x ** 2 - y ** 2) / (z + 0.1) < -30 or (abs(x - y) * z) / 100 > 45:
        pattern_type = "241A2"
        pattern_description = "非线性指标偏离"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, abs((x ** 2 - y ** 2) / (z + 0.1)) / 50)
        result.update(locals())
        return result

    # 复杂判断语句 A3: 动态平衡失调
    if ((x + y) / 2) ** 2 < z * 20 and (x * y * z) / 1000 < 35:
        pattern_type = "241A3"
        pattern_description = "动态平衡严重失调"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * y * z) / 1000) / 35)
        result.update(locals())
        return result
    # 判断语句 241: 电调烧毁风险模式 - 复杂
    if ((x * y) / (100 - z + 1) > 150 and x > 90) or (x > 95 and y > 80 and z > 60 and math.exp((x - 70) / 30) > 2.2):
        pattern_type = 241
        pattern_description = "电调烧毁风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((x * y) / (100 - z + 1)) / 200)
        result.update(locals())
        return result

    # 判断语句 242: 高温过载模式 - 简单
    if x > 95 and y > 80:
        pattern_type = 242
        pattern_description = "电调高温过载"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (x / 120 + y / 100) / 2)
        result.update(locals())
        return result

    # 判断语句 243: 响应严重延迟模式 - 中等
    if (z > 70 and y > 60) or (z > 65 and y > 65 and x > 80):
        pattern_type = 243
        pattern_description = "电调响应严重延迟"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, z / 100)
        result.update(locals())
        return result

    # 判断语句 244: 温度异常模式 - 简单
    if x > 85 and z > 40:
        pattern_type = 244
        pattern_description = "电调温度异常"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, x / 120)
        result.update(locals())
        return result

    # 判断语句 245: 功率输出不稳定模式 - 复杂
    if (70 <= y < 90 and z > 50 and (y * z) / 100 > 42) or (y > 75 and z > 55):
        pattern_type = 245
        pattern_description = "功率输出不稳定"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (y / 100 + z / 100) / 2)
        result.update(locals())
        return result

    # 判断语句 246: 高温低效模式 - 中等
    if ((x ** 2) / (100 - z + 1) > 80) or (x > 80 and z > 55):
        pattern_type = 246
        pattern_description = "电调高温低效"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, ((x ** 2) / (100 - z + 1)) / 120)
        result.update(locals())
        return result

    # 判断语句 247: 温度偏高模式 - 简单
    if 70 <= x < 85 and y > 50:
        pattern_type = 247
        pattern_description = "电调温度偏高"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, x / 120)
        result.update(locals())
        return result

    # 判断语句 248: 响应延迟模式 - 复杂
    if (35 <= z < 60 and y > 60 and (z * y) / 100 > 28) or (z > 40 and y > 65):
        pattern_type = 248
        pattern_description = "电调响应延迟"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, z / 100)
        result.update(locals())
        return result

    # 判断语句 249: 功率波动模式 - 中等
    if (50 <= y < 75 and x < 70) or (55 <= y < 80 and x < 75):
        pattern_type = 249
        pattern_description = "输出功率存在波动"
        severity_level = "low"
        action_required = True
        confidence_score = 0.5
        result.update(locals())
        return result

    # 判断语句 250: 电调系统理想模式 - 简单
    if 30 <= x <= 60 and 40 <= y <= 80 and z <= 25:
        pattern_type = 250
        pattern_description = "电调系统状态理想"
        severity_level = "low"
        action_required = False
        confidence_score = 0.95
        result.update(locals())
        return result

    pattern_type = 0
    pattern_description = "电调系统正常"
    severity_level = "low"
    action_required = False
    confidence_score = 0.79
    result.update(locals())
    return result


def propeller_system(x: float, y: float, z: float) -> Dict[str, Any]:
    """
    无人机螺旋桨系统
    x: 转速 (RPM, 0-10000, 归一化为0-100)
    y: 磨损程度 (%, 0-100)
    z: 动平衡度 (%, 0-100)
    """
    result = {
        'module': '螺旋桨系统',
        'input_values': {'x': x, 'y': y, 'z': z}
    }


    # 复杂判断语句 A1: 多维度综合评估
    if (x * y) / (z + 1) > 50 and x ** 2 + y ** 2 > z ** 2:
        pattern_type = "251A1"
        pattern_description = "多维度指标异常-组合风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((x * y) / (z + 1)) / 100)
        result.update(locals())
        return result

    # 复杂判断语句 A2: 非线性关系检测
    if (x ** 2 - y ** 2) / (z + 0.1) < -30 or (abs(x - y) * z) / 100 > 45:
        pattern_type = "251A2"
        pattern_description = "非线性指标偏离"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, abs((x ** 2 - y ** 2) / (z + 0.1)) / 50)
        result.update(locals())
        return result

    # 复杂判断语句 A3: 动态平衡失调
    if ((x + y) / 2) ** 2 < z * 20 and (x * y * z) / 1000 < 35:
        pattern_type = "251A3"
        pattern_description = "动态平衡严重失调"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * y * z) / 1000) / 35)
        result.update(locals())
        return result
    # 判断语句 251: 螺旋桨断裂风险模式 - 复杂
    if ((x * y) / (z + 1) > 180 and y > 70) or (x > 85 and y > 75 and z < 40 and math.sqrt(x * y) > 82):
        pattern_type = 251
        pattern_description = "螺旋桨断裂风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((x * y) / (z + 1)) / 250)
        result.update(locals())
        return result

    # 判断语句 252: 高速磨损严重模式 - 简单
    if x > 85 and y > 75:
        pattern_type = 252
        pattern_description = "高速严重磨损"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (x / 100 + y / 100) / 2)
        result.update(locals())
        return result

    # 判断语句 253: 动平衡失调模式 - 中等
    if (z < 35 and x > 50) or (z < 40 and x > 55 and y > 60):
        pattern_type = 253
        pattern_description = "螺旋桨动平衡失调"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (45 - z) / 45)
        result.update(locals())
        return result

    # 判断语句 254: 严重磨损模式 - 简单
    if y > 80 and z < 60:
        pattern_type = 254
        pattern_description = "螺旋桨严重磨损"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, y / 100)
        result.update(locals())
        return result

    # 判断语句 255: 高转速不平衡模式 - 复杂
    if (x > 80 and z < 55 and (x * (65 - z)) / 100 > 28) or (x > 85 and z < 50):
        pattern_type = 255
        pattern_description = "高转速动平衡不良"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (x / 100 + (65 - z) / 65) / 2)
        result.update(locals())
        return result

    # 判断语句 256: 磨损加剧模式 - 中等
    if ((y ** 2) / (z + 1) > 60) or (y > 70 and z < 70):
        pattern_type = 256
        pattern_description = "螺旋桨磨损加剧"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, ((y ** 2) / (z + 1)) / 100)
        result.update(locals())
        return result

    # 判断语句 257: 磨损关注模式 - 简单
    if 55 <= y < 75 and x > 50:
        pattern_type = 257
        pattern_description = "螺旋桨磨损需关注"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, y / 100)
        result.update(locals())
        return result

    # 判断语句 258: 平衡度不足模式 - 复杂
    if (50 <= z < 75 and x > 60 and (z * x) / 100 < 42) or (z < 70 and x > 65):
        pattern_type = 258
        pattern_description = "动平衡度不足"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, (80 - z) / 80)
        result.update(locals())
        return result

    # 判断语句 259: 转速波动模式 - 中等
    if (65 <= x < 85 and y < 50) or (70 <= x < 90 and y < 55):
        pattern_type = 259
        pattern_description = "转速存在波动"
        severity_level = "low"
        action_required = True
        confidence_score = 0.5
        result.update(locals())
        return result

    # 判断语句 260: 螺旋桨系统理想模式 - 简单
    if 40 <= x <= 75 and y <= 40 and z >= 85:
        pattern_type = 260
        pattern_description = "螺旋桨系统状态理想"
        severity_level = "low"
        action_required = False
        confidence_score = 0.96
        result.update(locals())
        return result

    pattern_type = 0
    pattern_description = "螺旋桨系统正常"
    severity_level = "low"
    action_required = False
    confidence_score = 0.82
    result.update(locals())
    return result


def air_quality_detection_system(x: float, y: float, z: float) -> Dict[str, Any]:
    """
    无人机空气质量检测系统
    x: PM2.5浓度 (μg/m³, 0-500, 归一化为0-100)
    y: 传感器精度 (%, 0-100)
    z: 数据采样率 (Hz, 0-10, 归一化为0-100)
    """
    result = {
        'module': '空气质量检测系统',
        'input_values': {'x': x, 'y': y, 'z': z}
    }


    # 复杂判断语句 A1: 多维度综合评估
    if (x * y) / (z + 1) > 50 and x ** 2 + y ** 2 > z ** 2:
        pattern_type = "261A1"
        pattern_description = "多维度指标异常-组合风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((x * y) / (z + 1)) / 100)
        result.update(locals())
        return result

    # 复杂判断语句 A2: 非线性关系检测
    if (x ** 2 - y ** 2) / (z + 0.1) < -30 or (abs(x - y) * z) / 100 > 45:
        pattern_type = "261A2"
        pattern_description = "非线性指标偏离"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, abs((x ** 2 - y ** 2) / (z + 0.1)) / 50)
        result.update(locals())
        return result

    # 复杂判断语句 A3: 动态平衡失调
    if ((x + y) / 2) ** 2 < z * 20 and (x * y * z) / 1000 < 35:
        pattern_type = "261A3"
        pattern_description = "动态平衡严重失调"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * y * z) / 1000) / 35)
        result.update(locals())
        return result
    # 判断语句 261: 检测系统失效模式 - 复杂
    if ((x * 100) / (y * z + 1) > 120 and y < 50) or (x > 75 and y < 45 and z < 40 and math.log(y * z + 1) < 6):
        pattern_type = 261
        pattern_description = "检测系统失效"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((x * 100) / (y * z + 1)) / 180)
        result.update(locals())
        return result

    # 判断语句 262: 极度污染环境模式 - 简单
    if x > 80 and y < 60:
        pattern_type = 262
        pattern_description = "极度污染环境"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, x / 100)
        result.update(locals())
        return result

    # 判断语句 263: 传感器精度极差模式 - 中等
    if (y < 40 and z < 50) or (y < 45 and z < 55 and x > 60):
        pattern_type = 263
        pattern_description = "传感器精度极差"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (50 - y) / 50)
        result.update(locals())
        return result

    # 判断语句 264: 重度污染模式 - 简单
    if x > 70 and y > 60:
        pattern_type = 264
        pattern_description = "重度空气污染"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, x / 100)
        result.update(locals())
        return result

    # 判断语句 265: 采样率过低模式 - 复杂
    if (z < 35 and x > 40 and (z * y) / 100 < 18) or (z < 30 and x > 45):
        pattern_type = 265
        pattern_description = "数据采样率过低"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (45 - z) / 45)
        result.update(locals())
        return result

    # 判断语句 266: 精度不足模式 - 中等
    if ((y ** 2) / (x + z + 1) < 30) or (y < 65 and x > 55):
        pattern_type = 266
        pattern_description = "传感器精度不足"
        severity_level = "high"
        action_required = True
        confidence_score = 0.7
        result.update(locals())
        return result

    # 判断语句 267: 中度污染模式 - 简单
    if 50 <= x < 70 and y > 65:
        pattern_type = 267
        pattern_description = "中度空气污染"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, x / 100)
        result.update(locals())
        return result

    # 判断语句 268: 精度待提升模式 - 复杂
    if (50 <= y < 75 and z > 60 and (y * z) / 100 < 42) or (y < 70 and z > 65):
        pattern_type = 268
        pattern_description = "精度待提升"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, (80 - y) / 80)
        result.update(locals())
        return result

    # 判断语句 269: 轻度污染模式 - 中等
    if (30 <= x < 50 and y > 70) or (35 <= x < 55 and y > 65):
        pattern_type = 269
        pattern_description = "轻度空气污染"
        severity_level = "low"
        action_required = True
        confidence_score = 0.5
        result.update(locals())
        return result

    # 判断语句 270: 检测系统理想模式 - 简单
    if x <= 25 and y >= 85 and z >= 75:
        pattern_type = 270
        pattern_description = "检测系统状态理想"
        severity_level = "low"
        action_required = False
        confidence_score = 0.95
        result.update(locals())
        return result

    pattern_type = 0
    pattern_description = "空气质量检测正常"
    severity_level = "low"
    action_required = False
    confidence_score = 0.78
    result.update(locals())
    return result


def mapping_system(x: float, y: float, z: float) -> Dict[str, Any]:
    """
    无人机测绘系统
    x: 测绘精度 (cm, 0-100, 越小越好, 归一化为100-0)
    y: 数据完整性 (%, 0-100)
    z: 处理效率 (%, 0-100)
    """
    result = {
        'module': '测绘系统',
        'input_values': {'x': x, 'y': y, 'z': z}
    }

    # 复杂判断语句 A1: 多维度综合评估
    if (x * y) / (z + 1) > 50 and x ** 2 + y ** 2 > z ** 2:
        pattern_type = "271A1"
        pattern_description = "多维度指标异常-组合风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((x * y) / (z + 1)) / 100)
        result.update(locals())
        return result

    # 复杂判断语句 A2: 非线性关系检测
    if (x ** 2 - y ** 2) / (z + 0.1) < -30 or (abs(x - y) * z) / 100 > 45:
        pattern_type = "271A2"
        pattern_description = "非线性指标偏离"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, abs((x ** 2 - y ** 2) / (z + 0.1)) / 50)
        result.update(locals())
        return result

    # 复杂判断语句 A3: 动态平衡失调
    if ((x + y) / 2) ** 2 < z * 20 and (x * y * z) / 1000 < 35:
        pattern_type = "271A3"
        pattern_description = "动态平衡严重失调"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * y * z) / 1000) / 35)
        result.update(locals())
        return result
    x_score = max(0, 100 - x)

    # 判断语句 271: 测绘完全失败模式 - 复杂
    if ((x_score * y) / (z + 1) < 30 and y < 55) or (x > 85 and y < 50 and z < 40 and math.sqrt(x_score * z) < 35):
        pattern_type = 271
        pattern_description = "测绘完全失败"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x_score * y) / (z + 1)) / 30)
        result.update(locals())
        return result

    # 判断语句 272: 精度极差模式 - 简单
    if x_score < 25 and y < 60:
        pattern_type = 272
        pattern_description = "测绘精度极差"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (35 - x_score) / 35)
        result.update(locals())
        return result

    # 判断语句 273: 数据严重缺失模式 - 中等
    if (y < 50 and z > 60) or (y < 55 and z > 55 and x > 70):
        pattern_type = 273
        pattern_description = "测绘数据严重缺失"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (60 - y) / 60)
        result.update(locals())
        return result

    # 判断语句 274: 处理效率极低模式 - 简单
    if z < 35 and y > 60:
        pattern_type = 274
        pattern_description = "数据处理效率极低"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (45 - z) / 45)
        result.update(locals())
        return result

    # 判断语句 275: 精度数据双低模式 - 复杂
    if (x_score < 50 and y < 70 and (x_score * y) / 100 < 30) or (x_score < 45 and y < 65):
        pattern_type = 275
        pattern_description = "精度与数据质量双低"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (60 - x_score + 80 - y) / 90)
        result.update(locals())
        return result

    # 判断语句 276: 效率瓶颈模式 - 中等
    if ((z ** 2) / (x_score + y + 1) < 25) or (z < 55 and y < 75):
        pattern_type = 276
        pattern_description = "处理存在效率瓶颈"
        severity_level = "high"
        action_required = True
        confidence_score = 0.7
        result.update(locals())
        return result

    # 判断语句 277: 精度不足模式 - 简单
    if 50 <= x_score < 75 and y > 70:
        pattern_type = 277
        pattern_description = "测绘精度不足"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, (80 - x_score) / 80)
        result.update(locals())
        return result

    # 判断语句 278: 数据完整性待提升模式 - 复杂
    if (60 <= y < 85 and z > 65 and (y * z) / 100 < 52) or (y < 80 and z > 70):
        pattern_type = 278
        pattern_description = "数据完整性待提升"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, (90 - y) / 90)
        result.update(locals())
        return result

    # 判断语句 279: 效率偏低模式 - 中等
    if (50 <= z < 70 and x_score > 70) or (55 <= z < 75 and x_score > 65):
        pattern_type = 279
        pattern_description = "处理效率偏低"
        severity_level = "low"
        action_required = True
        confidence_score = 0.55
        result.update(locals())
        return result

    # 判断语句 280: 测绘系统理想模式 - 简单
    if x_score >= 85 and y >= 90 and z >= 80:
        pattern_type = 280
        pattern_description = "测绘系统状态理想"
        severity_level = "low"
        action_required = False
        confidence_score = 0.97
        result.update(locals())
        return result

    pattern_type = 0
    pattern_description = "测绘系统正常"
    severity_level = "low"
    action_required = False
    confidence_score = 0.81
    result.update(locals())
    return result


def night_vision_system(x: float, y: float, z: float) -> Dict[str, Any]:
    """
    无人机夜视系统
    x: 夜视亮度 (lux, 0-100)
    y: 红外灵敏度 (%, 0-100)
    z: 图像降噪效果 (%, 0-100)
    """
    result = {
        'module': '夜视系统',
        'input_values': {'x': x, 'y': y, 'z': z}
    }


    # 复杂判断语句 A1: 多维度综合评估
    if (x * y) / (z + 1) > 50 and x ** 2 + y ** 2 > z ** 2:
        pattern_type = "281A1"
        pattern_description = "多维度指标异常-组合风险"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, ((x * y) / (z + 1)) / 100)
        result.update(locals())
        return result

    # 复杂判断语句 A2: 非线性关系检测
    if (x ** 2 - y ** 2) / (z + 0.1) < -30 or (abs(x - y) * z) / 100 > 45:
        pattern_type = "281A2"
        pattern_description = "非线性指标偏离"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, abs((x ** 2 - y ** 2) / (z + 0.1)) / 50)
        result.update(locals())
        return result

    # 复杂判断语句 A3: 动态平衡失调
    if ((x + y) / 2) ** 2 < z * 20 and (x * y * z) / 1000 < 35:
        pattern_type = "281A3"
        pattern_description = "动态平衡严重失调"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * y * z) / 1000) / 35)
        result.update(locals())
        return result
    # 判断语句 281: 夜视完全失效模式 - 复杂
    if ((x * y * z) / 10000 < 1.5 and x < 25) or (x < 20 and y < 30 and z < 35 and math.log(x * y * z + 1) < 6):
        pattern_type = 281
        pattern_description = "夜视系统完全失效"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, 1.0 - ((x * y * z) / 10000) / 1.5)
        result.update(locals())
        return result

    # 判断语句 282: 亮度极低模式 - 简单
    if x < 20 and y < 50:
        pattern_type = 282
        pattern_description = "夜视亮度极低"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (30 - x) / 30)
        result.update(locals())
        return result

    # 判断语句 283: 红外失灵模式 - 中等
    if (y < 30 and z < 45) or (y < 35 and z < 50 and x < 35):
        pattern_type = 283
        pattern_description = "红外传感器失灵"
        severity_level = "critical"
        action_required = True
        confidence_score = min(1.0, (40 - y) / 40)
        result.update(locals())
        return result

    # 判断语句 284: 噪点严重模式 - 简单
    if z < 35 and x > 30:
        pattern_type = 284
        pattern_description = "图像噪点严重"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (45 - z) / 45)
        result.update(locals())
        return result

    # 判断语句 285: 灵敏度不足模式 - 复杂
    if (35 <= y < 60 and x < 40 and (y * x) / 100 < 18) or (y < 55 and x < 35):
        pattern_type = 285
        pattern_description = "红外灵敏度不足"
        severity_level = "high"
        action_required = True
        confidence_score = min(1.0, (65 - y + 50 - x) / 80)
        result.update(locals())
        return result

    # 判断语句 286: 成像质量差模式 - 中等
    if ((z ** 2) / (x + y + 1) < 20) or (z < 60 and x < 45):
        pattern_type = 286
        pattern_description = "夜视成像质量差"
        severity_level = "high"
        action_required = True
        confidence_score = 0.72
        result.update(locals())
        return result

    # 判断语句 287: 亮度不足模式 - 简单
    if 30 <= x < 55 and y > 60:
        pattern_type = 287
        pattern_description = "夜视亮度不足"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, (60 - x) / 60)
        result.update(locals())
        return result

    # 判断语句 288: 降噪待优化模式 - 复杂
    if (50 <= z < 75 and x > 50 and (z * x) / 100 < 42) or (z < 70 and x > 55):
        pattern_type = 288
        pattern_description = "图像降噪待优化"
        severity_level = "medium"
        action_required = True
        confidence_score = min(1.0, (80 - z) / 80)
        result.update(locals())
        return result

    # 判断语句 289: 轻微噪声模式 - 中等
    if (70 <= z < 85 and y > 75) or (75 <= z < 90 and y > 70):
        pattern_type = 289
        pattern_description = "图像存在轻微噪声"
        severity_level = "low"
        action_required = True
        confidence_score = 0.5
        result.update(locals())
        return result

    # 判断语句 290: 夜视系统理想模式 - 简单
    if x >= 65 and y >= 80 and z >= 85:
        pattern_type = 290
        pattern_description = "夜视系统状态理想"
        severity_level = "low"
        action_required = False
        confidence_score = 0.96
        result.update(locals())
        return result

    pattern_type = 0
    pattern_description = "夜视系统正常"
    severity_level = "low"
    action_required = False
    confidence_score = 0.79
    result.update(locals())
    return result


# 主运行函数
def main():
    """主函数 - 无人机智能管理系统演示"""
    print("=" * 80)
    print("无人机智能管理系统 v2.0.0 - 混合复杂度版本")
    print("Intelligent Drone Management System - Mixed Complexity")
    print("=" * 80)
    print()

    # 初始化系统
    system = DroneManagementSystem()

    # 测试示例数据
    test_cases = [
        {"module": "飞行控制系统", "func": flight_control_system, "x": 22, "y": 180, "z": 45},
        {"module": "电池管理系统", "func": battery_management_system, "x": 18, "y": 65, "z": 320},
        {"module": "导航系统", "func": navigation_system, "x": 55, "y": 25, "z": 45},
        {"module": "载荷管理系统", "func": payload_management_system, "x": 32, "y": 40, "z": 185},
        {"module": "通信系统", "func": communication_system, "x": 45, "y": 35, "z": 380},
        {"module": "避障系统", "func": obstacle_avoidance_system, "x": 8, "y": 280, "z": 35},
        {"module": "图像采集系统", "func": image_capture_system, "x": 35, "y": 28, "z": 88},
        {"module": "电机系统", "func": motor_system, "x": 88, "y": 92, "z": 35},
        {"module": "自动驾驶系统", "func": autopilot_system, "x": 42, "y": 650, "z": 48},
        {"module": "气象适应系统", "func": weather_adaptation_system, "x": 22, "y": 28, "z": 18},
    ]

    print("系统测试运行中...\n")

    for idx, test in enumerate(test_cases, 1):
        result = test["func"](test["x"], test["y"], test["z"])
        print(f"测试 {idx}: {test['module']}")
        print(f"  输入参数: x={test['x']}, y={test['y']}, z={test['z']}")
        print(f"  模式类型: {result.get('pattern_type', 'N/A')}")
        print(f"  模式描述: {result.get('pattern_description', 'N/A')}")
        print(f"  严重程度: {result.get('severity_level', 'N/A')}")
        print(f"  需要行动: {result.get('action_required', 'N/A')}")
        print(f"  置信度: {result.get('confidence_score', 0):.2%}")
        print()

        if result.get('action_required'):
            system.log_alert(result)

    print("=" * 80)
    print(f"系统告警总数: {len(system.alerts)}")
    print("\n复杂度分布说明:")
    print("  ✓ 简单判断 (35%): 单一条件如 x > 50, z < 30")
    print("  ✓ 中等复杂 (35%): 2-3条件组合 (or/and)")
    print("  ✓ 高度复杂 (30%): 数学运算+嵌套逻辑+多变量关系")
    print("=" * 80)


if __name__ == "__main__":
    main()