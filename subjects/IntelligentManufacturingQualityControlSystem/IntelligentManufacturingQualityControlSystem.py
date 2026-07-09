import random
import json
import csv
import sqlite3
import datetime
import statistics
import numpy as np
import matplotlib.pyplot as plt
from math import sqrt, sin, cos, pi
from collections import defaultdict, deque
from dataclasses import dataclass, asdict
from typing import List, Dict, Tuple, Optional, Any
from enum import Enum
import threading
import time
import logging
import pickle
import os
import scipy.fft
import scipy.signal
import scipy.stats
from sklearn.ensemble import IsolationForest, RandomForestRegressor
from sklearn.preprocessing import StandardScaler
from sklearn.model_selection import train_test_split
from sklearn.metrics import mean_squared_error
import joblib
import asyncio
from flask import Flask, jsonify, request
import smtplib
import requests
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

class AlertLevel(Enum):
    """警报级别枚举"""
    NORMAL = "normal"
    LOW = "low"
    MEDIUM = "medium"
    HIGH = "high"
    CRITICAL = "critical"
    EMERGENCY = "emergency"


class SystemStatus(Enum):
    """系统状态枚举"""
    RUNNING = "running"
    STOPPED = "stopped"
    MAINTENANCE = "maintenance"
    ERROR = "error"
    STARTUP = "startup"
    SHUTDOWN = "shutdown"


@dataclass
class ProductionRecord:
    """生产记录数据类"""
    timestamp: datetime.datetime
    cycle_id: int
    temperature: float
    pressure: float
    density: float
    humidity: float
    speed: float
    vibration: float
    flow_rate: float
    voltage: float
    current: float
    rpm: float
    torque: float
    power: float
    section_used: int
    alert_level: str
    actions_count: int
    adjustments_count: int
    quality_score: float
    efficiency_score: float


@dataclass
class CalibrationRecord:
    """校准记录数据类"""
    parameter: str
    last_calibration: datetime.datetime
    next_calibration: datetime.datetime
    calibrator: str
    status: str
    deviation: float


class DataLogger:
    """数据记录器"""

    def __init__(self, db_path: str = "production_data.db"):
        self.db_path = db_path
        self.init_database()

    def init_database(self):
        """初始化数据库"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        # 创建生产数据表
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS production_records (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                timestamp TEXT,
                cycle_id INTEGER,
                temperature REAL,
                pressure REAL,
                density REAL,
                humidity REAL,
                speed REAL,
                vibration REAL,
                flow_rate REAL,
                voltage REAL,
                current REAL,
                rpm REAL,
                torque REAL,
                power REAL,
                section_used INTEGER,
                alert_level TEXT,
                actions_count INTEGER,
                adjustments_count INTEGER,
                quality_score REAL,
                efficiency_score REAL
            )
        ''')

        # 创建报警日志表
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS alert_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                timestamp TEXT,
                level TEXT,
                parameter TEXT,
                value REAL,
                threshold REAL,
                message TEXT,
                resolved INTEGER DEFAULT 0
            )
        ''')

        # 创建校准记录表
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS calibration_records (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                parameter TEXT,
                last_calibration TEXT,
                next_calibration TEXT,
                calibrator TEXT,
                status TEXT,
                deviation REAL
            )
        ''')

        # 创建维护记录表
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS maintenance_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                timestamp TEXT,
                type TEXT,
                description TEXT,
                technician TEXT,
                duration_hours REAL,
                cost REAL,
                parts_replaced TEXT
            )
        ''')

        conn.commit()
        conn.close()

    def log_production_record(self, record: ProductionRecord):
        """记录生产数据"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()


        cursor.execute('''
            INSERT INTO production_records VALUES (
                NULL, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?
            )
        ''', (
            record.timestamp.isoformat(),
            record.cycle_id,
            record.temperature,
            record.pressure,
            record.density,
            record.humidity,
            record.speed,
            record.vibration,
            record.flow_rate,
            record.voltage,
            record.current,
            record.rpm,
            record.torque,
            record.power,
            record.section_used,
            record.alert_level,
            record.actions_count,
            record.adjustments_count,
            record.quality_score,
            record.efficiency_score
        ))

        conn.commit()
        conn.close()

    def log_alert(self, level: str, parameter: str, value: float, threshold: float, message: str):
        """记录报警"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute('''
            INSERT INTO alert_logs VALUES (NULL, ?, ?, ?, ?, ?, ?, 0)
        ''', (datetime.datetime.now().isoformat(), level, parameter, value, threshold, message))

        conn.commit()
        conn.close()



    def get_recent_records(self, hours: int = 24) -> List[ProductionRecord]:
        """获取最近的生产记录"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        since = datetime.datetime.now() - datetime.timedelta(hours=hours)
        cursor.execute('''
            SELECT * FROM production_records 
            WHERE timestamp > ?
            ORDER BY timestamp DESC
        ''', (since.isoformat(),))



        records = []
        for row in cursor.fetchall():
            records.append(ProductionRecord(
                timestamp=datetime.datetime.fromisoformat(row[1]),
                cycle_id=row[2],
                temperature=row[3],
                pressure=row[4],
                density=row[5],
                humidity=row[6],
                speed=row[7],
                vibration=row[8],
                flow_rate=row[9],
                voltage=row[10],
                current=row[11],
                rpm=row[12],
                torque=row[13],
                power=row[14],
                section_used=row[15],
                alert_level=row[16],
                actions_count=row[17],
                adjustments_count=row[18],
                quality_score=row[19],
                efficiency_score=row[20]
            ))

        conn.close()
        return records



class StatisticalAnalyzer:
    """统计分析器"""

    def __init__(self):
        self.parameter_stats = defaultdict(list)
        self.trend_window = 50  # 趋势分析窗口

    def update_stats(self, data: Dict[str, float]):
        """更新统计数据"""
        for param, value in data.items():
            if len(self.parameter_stats[param]) >= 1000:  # 保持最近1000个数据点
                self.parameter_stats[param].pop(0)
            self.parameter_stats[param].append(value)

    def get_parameter_statistics(self, parameter: str) -> Dict[str, float]:
        """获取参数统计信息"""
        if parameter not in self.parameter_stats or not self.parameter_stats[parameter]:
            return {}




        values = self.parameter_stats[parameter]
        return {
            'mean': statistics.mean(values),
            'median': statistics.median(values),
            'std_dev': statistics.stdev(values) if len(values) > 1 else 0,
            'min': min(values),
            'max': max(values),
            'range': max(values) - min(values),
            'count': len(values),
            'variance': statistics.variance(values) if len(values) > 1 else 0
        }




    def detect_trend(self, parameter: str) -> Dict[str, Any]:
        """检测参数趋势"""
        if parameter not in self.parameter_stats:
            return {'trend': 'unknown', 'confidence': 0}

        values = self.parameter_stats[parameter][-self.trend_window:]
        if len(values) < 10:
            return {'trend': 'insufficient_data', 'confidence': 0}




        # 简单线性回归检测趋势
        n = len(values)
        x = list(range(n))
        y = values

        x_mean = sum(x) / n
        y_mean = sum(y) / n

        numerator = sum((x[i] - x_mean) * (y[i] - y_mean) for i in range(n))
        denominator = sum((x[i] - x_mean) ** 2 for i in range(n))

        if denominator == 0:
            return {'trend': 'stable', 'confidence': 0}

        slope = numerator / denominator

        # 计算置信度
        y_pred = [slope * (i - x_mean) + y_mean for i in x]
        mse = sum((y[i] - y_pred[i]) ** 2 for i in range(n)) / n
        confidence = max(0, min(1, 1 - mse / statistics.variance(y) if statistics.variance(y) > 0 else 0))

        if abs(slope) < 0.001:
            trend = 'stable'
        elif slope > 0:
            trend = 'increasing'
        else:
            trend = 'decreasing'

        return {
            'trend': trend,
            'slope': slope,
            'confidence': confidence,
            'mse': mse
        }

    def detect_anomalies(self, parameter: str, threshold: float = 2.0) -> List[Tuple[int, float]]:
        """检测异常值（基于Z-score）"""
        if parameter not in self.parameter_stats:
            return []

        values = self.parameter_stats[parameter]
        if len(values) < 10:
            return []

        mean_val = statistics.mean(values)
        std_val = statistics.stdev(values)

        anomalies = []
        for i, value in enumerate(values):
            if std_val > 0:
                z_score = abs((value - mean_val) / std_val)
                if z_score > threshold:
                    anomalies.append((i, value))

        return anomalies


class PredictiveMaintenance:
    """预测性维护模块"""

    def __init__(self):
        self.component_health = {
            'temperature_sensor': 1.0,
            'pressure_sensor': 1.0,
            'flow_sensor': 1.0,
            'motor': 1.0,
            'pump': 1.0,
            'valve': 1.0,
            'heater': 1.0,
            'cooler': 1.0
        }
        self.maintenance_history = defaultdict(list)
        self.failure_predictors = {}

    def update_component_health(self, data: Dict[str, float], alert_level: str):
        """更新组件健康度"""
        degradation_factors = {
            'normal': 0.0001,
            'low': 0.0005,
            'medium': 0.001,
            'high': 0.005,
            'critical': 0.01,
            'emergency': 0.02
        }

        factor = degradation_factors.get(alert_level, 0.001)

        # 根据参数值和警报级别更新健康度
        for component in self.component_health:
            self.component_health[component] = max(0, self.component_health[component] - factor)

            # 根据具体参数调整
            if 'temperature' in component and 'temperature' in data:
                temp = data['temperature']
                if temp > 230 or temp < 170:
                    self.component_health[component] -= 0.002

            if 'pressure' in component and 'pressure' in data:
                pressure = data['pressure']
                if pressure > 130 or pressure < 70:
                    self.component_health[component] -= 0.002

    def predict_failure_risk(self, component: str) -> Dict[str, Any]:
        """预测组件故障风险"""
        health = self.component_health.get(component, 1.0)

        if health > 0.8:
            risk_level = "low"
            days_to_maintenance = 365
        elif health > 0.6:
            risk_level = "medium"
            days_to_maintenance = 180
        elif health > 0.4:
            risk_level = "high"
            days_to_maintenance = 90
        elif health > 0.2:
            risk_level = "critical"
            days_to_maintenance = 30
        else:
            risk_level = "imminent"
            days_to_maintenance = 7

        return {
            'component': component,
            'health_score': health,
            'risk_level': risk_level,
            'recommended_maintenance_days': days_to_maintenance,
            'confidence': min(1.0, (1.0 - health) * 2)  # 健康度越低，预测置信度越高
        }

    def get_maintenance_schedule(self) -> List[Dict[str, Any]]:
        """获取维护计划"""
        schedule = []
        for component in self.component_health:
            prediction = self.predict_failure_risk(component)
            if prediction['risk_level'] in ['high', 'critical', 'imminent']:
                schedule.append(prediction)

        return sorted(schedule, key=lambda x: x['recommended_maintenance_days'])


class AlarmManager:
    """警报管理器"""

    def __init__(self):
        self.active_alarms = {}
        self.alarm_history = []
        self.alarm_thresholds = {
            'temperature': {'low': 170, 'high': 230, 'critical_low': 160, 'critical_high': 240},
            'pressure': {'low': 70, 'high': 130, 'critical_low': 60, 'critical_high': 140},
            'humidity': {'low': 35, 'high': 65, 'critical_low': 30, 'critical_high': 70},
            'vibration': {'low': 0.4, 'high': 2.2, 'critical_low': 0.3, 'critical_high': 2.5}
        }

    def check_alarms(self, data: Dict[str, float]) -> List[Dict[str, Any]]:
        """检查并生成警报"""
        new_alarms = []

        for param, value in data.items():
            if param in self.alarm_thresholds:
                thresholds = self.alarm_thresholds[param]
                alarm_key = f"{param}_{int(time.time())}"

                if value < thresholds['critical_low'] or value > thresholds['critical_high']:
                    alarm = {
                        'id': alarm_key,
                        'parameter': param,
                        'value': value,
                        'level': 'critical',
                        'message': f"{param} 严重超出正常范围: {value}",
                        'timestamp': datetime.datetime.now(),
                        'acknowledged': False
                    }
                    new_alarms.append(alarm)
                    self.active_alarms[alarm_key] = alarm

                elif value < thresholds['low'] or value > thresholds['high']:
                    alarm = {
                        'id': alarm_key,
                        'parameter': param,
                        'value': value,
                        'level': 'warning',
                        'message': f"{param} 超出推荐范围: {value}",
                        'timestamp': datetime.datetime.now(),
                        'acknowledged': False
                    }
                    new_alarms.append(alarm)
                    self.active_alarms[alarm_key] = alarm

        self.alarm_history.extend(new_alarms)
        return new_alarms

    def acknowledge_alarm(self, alarm_id: str):
        """确认警报"""
        if alarm_id in self.active_alarms:
            self.active_alarms[alarm_id]['acknowledged'] = True

    def clear_alarm(self, alarm_id: str):
        """清除警报"""
        if alarm_id in self.active_alarms:
            del self.active_alarms[alarm_id]

    def get_active_alarms(self, level: str = None) -> List[Dict[str, Any]]:
        """获取活跃警报"""
        alarms = list(self.active_alarms.values())
        if level:
            alarms = [a for a in alarms if a['level'] == level]
        return sorted(alarms, key=lambda x: x['timestamp'], reverse=True)


class ConfigurationManager:
    """配置管理器"""

    def __init__(self, config_file: str = "system_config.json"):
        self.config_file = config_file
        self.config = self.load_config()

    def load_config(self) -> Dict[str, Any]:
        """加载配置"""
        default_config = {
            'system': {
                'name': 'MultiVariable Quality Control System',
                'version': '2.0',
                'max_cycles': 1000,
                'sampling_rate': 1.0,
                'auto_adjustment': True
            },
            'standards': {
                'temperature': [160, 240],
                'pressure': [60, 140],
                'density': [1.2, 1.5],
                'humidity': [40, 60],
                'speed': [1200, 1800],
                'vibration': [0.5, 2.0],
                'flow_rate': [50, 80],
                'voltage': [220, 240],
                'current': [10, 15],
                'rpm': [2800, 3200],
                'torque': [150, 200],
                'power': [45, 65]
            },
            'alerts': {
                'email_notifications': True,
                'sms_notifications': False,
                'sound_alerts': True,
                'escalation_time': 300
            },
            'maintenance': {
                'preventive_interval_days': 90,
                'calibration_interval_days': 180,
                'backup_interval_hours': 24
            }
        }

        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                loaded_config = json.load(f)
                # 合并默认配置和加载的配置
                self._merge_configs(default_config, loaded_config)
                return default_config
        except FileNotFoundError:
            self.save_config(default_config)
            return default_config

    def _merge_configs(self, default: dict, loaded: dict):
        """合并配置"""
        for key, value in loaded.items():
            if key in default:
                if isinstance(default[key], dict) and isinstance(value, dict):
                    self._merge_configs(default[key], value)
                else:
                    default[key] = value

    def save_config(self, config: Dict[str, Any] = None):
        """保存配置"""
        config_to_save = config or self.config
        with open(self.config_file, 'w', encoding='utf-8') as f:
            json.dump(config_to_save, f, indent=2, ensure_ascii=False)

    def get(self, key_path: str, default=None):
        """获取配置值"""
        keys = key_path.split('.')
        value = self.config
        for key in keys:
            if isinstance(value, dict) and key in value:
                value = value[key]
            else:
                return default
        return value

    def set(self, key_path: str, value):
        """设置配置值"""
        keys = key_path.split('.')
        config = self.config
        for key in keys[:-1]:
            if key not in config:
                config[key] = {}
            config = config[key]
        config[keys[-1]] = value
        self.save_config()


class ReportGenerator:
    """报告生成器"""

    def __init__(self, data_logger: DataLogger):
        self.data_logger = data_logger

    def generate_daily_report(self, date: datetime.date = None) -> Dict[str, Any]:
        """生成日报"""
        if date is None:
            date = datetime.date.today()

        start_time = datetime.datetime.combine(date, datetime.time.min)
        end_time = datetime.datetime.combine(date, datetime.time.max)

        records = self.data_logger.get_recent_records(24)
        daily_records = [r for r in records if start_time <= r.timestamp <= end_time]

        if not daily_records:
            return {'error': '无数据可用于生成报告'}

        # 计算统计信息
        quality_scores = [r.quality_score for r in daily_records]
        efficiency_scores = [r.efficiency_score for r in daily_records]
        alert_levels = [r.alert_level for r in daily_records]

        alert_counts = {level.value: alert_levels.count(level.value) for level in AlertLevel}

        report = {
            'date': date.isoformat(),
            'total_cycles': len(daily_records),
            'quality_metrics': {
                'average_quality_score': statistics.mean(quality_scores),
                'min_quality_score': min(quality_scores),
                'max_quality_score': max(quality_scores),
                'quality_std_dev': statistics.stdev(quality_scores) if len(quality_scores) > 1 else 0
            },
            'efficiency_metrics': {
                'average_efficiency_score': statistics.mean(efficiency_scores),
                'min_efficiency_score': min(efficiency_scores),
                'max_efficiency_score': max(efficiency_scores),
                'efficiency_std_dev': statistics.stdev(efficiency_scores) if len(efficiency_scores) > 1 else 0
            },
            'alert_summary': alert_counts,
            'parameter_averages': {
                'temperature': statistics.mean([r.temperature for r in daily_records]),
                'pressure': statistics.mean([r.pressure for r in daily_records]),
                'density': statistics.mean([r.density for r in daily_records]),
                'humidity': statistics.mean([r.humidity for r in daily_records]),
                'speed': statistics.mean([r.speed for r in daily_records]),
                'vibration': statistics.mean([r.vibration for r in daily_records]),
                'flow_rate': statistics.mean([r.flow_rate for r in daily_records]),
                'voltage': statistics.mean([r.voltage for r in daily_records]),
                'current': statistics.mean([r.current for r in daily_records]),
                'rpm': statistics.mean([r.rpm for r in daily_records]),
                'torque': statistics.mean([r.torque for r in daily_records]),
                'power': statistics.mean([r.power for r in daily_records])
            },
            'section_usage': {
                f'section_{i}': len([r for r in daily_records if r.section_used == i])
                for i in range(1, 11)
            }
        }

        return report

    def export_report_to_csv(self, report: Dict[str, Any], filename: str):
        """导出报告到CSV"""
        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)

            # 写入标题
            writer.writerow(['报告类型', '日期', '参数', '值'])

            # 写入基本信息
            writer.writerow(['基本信息', report['date'], '总周期数', report['total_cycles']])

            # 写入质量指标
            for key, value in report['quality_metrics'].items():
                writer.writerow(['质量指标', report['date'], key, value])

            # 写入效率指标
            for key, value in report['efficiency_metrics'].items():
                writer.writerow(['效率指标', report['date'], key, value])

            # 写入参数平均值
            for key, value in report['parameter_averages'].items():
                writer.writerow(['参数平均值', report['date'], key, value])


class PerformanceMonitor:
    """性能监控器"""

    def __init__(self):
        self.metrics = {
            'system_uptime': time.time(),
            'total_cycles_processed': 0,
            'average_cycle_time': 0,
            'cpu_usage_history': deque(maxlen=100),
            'memory_usage_history': deque(maxlen=100),
            'processing_times': deque(maxlen=1000)
        }

    def record_cycle_time(self, cycle_time: float):
        """记录周期处理时间"""
        self.metrics['processing_times'].append(cycle_time)
        self.metrics['total_cycles_processed'] += 1

        if len(self.metrics['processing_times']) > 0:
            self.metrics['average_cycle_time'] = statistics.mean(self.metrics['processing_times'])

    def get_performance_stats(self) -> Dict[str, Any]:
        """获取性能统计"""
        uptime = time.time() - self.metrics['system_uptime']

        return {
            'uptime_seconds': uptime,
            'uptime_hours': uptime / 3600,
            'total_cycles': self.metrics['total_cycles_processed'],
            'average_cycle_time': self.metrics['average_cycle_time'],
            'cycles_per_hour': self.metrics['total_cycles_processed'] / (uptime / 3600) if uptime > 0 else 0,
            'current_throughput': 3600 / self.metrics['average_cycle_time'] if self.metrics[
                                                                                   'average_cycle_time'] > 0 else 0
        }


class DataVisualization:
    """数据可视化"""

    def __init__(self, data_logger: DataLogger):
        self.data_logger = data_logger

    def plot_parameter_trend(self, parameter: str, hours: int = 24, save_path: str = None):
        """绘制参数趋势图"""
        records = self.data_logger.get_recent_records(hours)

        if not records:
            print("没有数据用于绘图")
            return

        times = [r.timestamp for r in records]
        values = [getattr(r, parameter) for r in records]

        plt.figure(figsize=(12, 6))
        plt.plot(times, values, 'b-', linewidth=2, label=parameter)
        plt.title(f'{parameter} 趋势 (最近 {hours} 小时)')
        plt.xlabel('时间')
        plt.ylabel(parameter)
        plt.grid(True, alpha=0.3)
        plt.legend()
        plt.xticks(rotation=45)
        plt.tight_layout()

        if save_path:
            plt.savefig(save_path, dpi=300, bbox_inches='tight')
        else:
            plt.show()

        plt.close()

    def plot_quality_distribution(self, hours: int = 24, save_path: str = None):
        """绘制质量分布图"""
        records = self.data_logger.get_recent_records(hours)

        if not records:
            print("没有数据用于绘图")
            return

        quality_scores = [r.quality_score for r in records]

        plt.figure(figsize=(10, 6))
        plt.hist(quality_scores, bins=20, alpha=0.7, color='green', edgecolor='black')
        plt.title(f'质量分数分布 (最近 {hours} 小时)')
        plt.xlabel('质量分数')
        plt.ylabel('频次')
        plt.grid(True, alpha=0.3)

        # 添加统计信息
        mean_score = statistics.mean(quality_scores)
        plt.axvline(mean_score, color='red', linestyle='--', label=f'平均值: {mean_score:.2f}')
        plt.legend()

        if save_path:
            plt.savefig(save_path, dpi=300, bbox_inches='tight')
        else:
            plt.show()

        plt.close()

    def plot_correlation_matrix(self, hours: int = 24, save_path: str = None):
        """绘制参数相关性矩阵"""
        records = self.data_logger.get_recent_records(hours)

        if not records:
            print("没有数据用于绘图")
            return

        # 提取数值参数
        parameters = ['temperature', 'pressure', 'density', 'humidity', 'speed',
                      'vibration', 'flow_rate', 'voltage', 'current', 'rpm', 'torque', 'power']

        data_matrix = []
        for param in parameters:
            data_matrix.append([getattr(r, param) for r in records])

        # 计算相关性矩阵
        correlation_matrix = np.corrcoef(data_matrix)

        plt.figure(figsize=(12, 10))
        plt.imshow(correlation_matrix, cmap='coolwarm', vmin=-1, vmax=1)
        plt.colorbar(label='相关系数')
        plt.title('参数相关性矩阵')
        plt.xticks(range(len(parameters)), parameters, rotation=45)
        plt.yticks(range(len(parameters)), parameters)

        # 添加数值标签
        for i in range(len(parameters)):
            for j in range(len(parameters)):
                plt.text(j, i, f'{correlation_matrix[i, j]:.2f}',
                         ha='center', va='center',
                         color='white' if abs(correlation_matrix[i, j]) > 0.5 else 'black')

        plt.tight_layout()

        if save_path:
            plt.savefig(save_path, dpi=300, bbox_inches='tight')
        else:
            plt.show()

        plt.close()


class SystemHealthMonitor:
    """系统健康监控"""

    def __init__(self):
        self.health_metrics = {
            'sensor_health': defaultdict(float),
            'communication_status': True,
            'database_status': True,
            'processing_errors': 0,
            'last_error_time': None,
            'system_temperature': 25.0,
            'disk_usage': 0.0,
            'memory_usage': 0.0
        }

    def check_sensor_health(self, data: Dict[str, float]) -> Dict[str, str]:
        """检查传感器健康状态"""
        sensor_status = {}

        for param, value in data.items():
            # 简单的传感器健康检查
            if value is None or value < 0:
                sensor_status[param] = 'error'
                self.health_metrics['sensor_health'][param] = 0.0
            elif param == 'temperature' and (value < -50 or value > 500):
                sensor_status[param] = 'warning'
                self.health_metrics['sensor_health'][param] = 0.5
            elif param == 'pressure' and (value < 0 or value > 1000):
                sensor_status[param] = 'warning'
                self.health_metrics['sensor_health'][param] = 0.5
            else:
                sensor_status[param] = 'healthy'
                self.health_metrics['sensor_health'][param] = 1.0

        return sensor_status

    def get_system_health_score(self) -> float:
        """计算系统健康分数"""
        if not self.health_metrics['sensor_health']:
            return 0.0

        sensor_scores = list(self.health_metrics['sensor_health'].values())
        avg_sensor_health = statistics.mean(sensor_scores)

        # 系统健康分数基于多个因素
        health_score = avg_sensor_health * 0.6

        if self.health_metrics['communication_status']:
            health_score += 0.2

        if self.health_metrics['database_status']:
            health_score += 0.1

        if self.health_metrics['processing_errors'] < 10:
            health_score += 0.1

        return min(1.0, health_score)


class BackupManager:
    """备份管理器"""

    def __init__(self, data_logger: DataLogger):
        self.data_logger = data_logger
        self.backup_dir = "backups"
        os.makedirs(self.backup_dir, exist_ok=True)

    def create_backup(self) -> str:
        """创建系统备份"""
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_filename = f"system_backup_{timestamp}.sqlite"
        backup_path = os.path.join(self.backup_dir, backup_filename)

        # 复制数据库文件
        import shutil
        shutil.copy2(self.data_logger.db_path, backup_path)

        return backup_path

    def restore_backup(self, backup_path: str) -> bool:
        """恢复备份"""
        try:
            import shutil
            shutil.copy2(backup_path, self.data_logger.db_path)
            return True
        except Exception as e:
            print(f"备份恢复失败: {e}")
            return False

    def cleanup_old_backups(self, keep_days: int = 30):
        """清理旧备份"""
        cutoff_date = datetime.datetime.now() - datetime.timedelta(days=keep_days)

        for filename in os.listdir(self.backup_dir):
            if filename.startswith("system_backup_") and filename.endswith(".sqlite"):
                file_path = os.path.join(self.backup_dir, filename)
                file_time = datetime.datetime.fromtimestamp(os.path.getctime(file_path))

                if file_time < cutoff_date:
                    os.remove(file_path)
                    print(f"已删除旧备份: {filename}")


class MultiVariableQualityController:
    """增强版多变量质量控制器"""

    def __init__(self):
        # 初始化系统状态
        self.production_line = SystemStatus.RUNNING.value
        self.alert_status = AlertLevel.NORMAL.value
        self.adjustments = []
        self.inspection_required = False
        self.cycle_counter = 0

        # 初始化子系统
        self.config_manager = ConfigurationManager()
        self.data_logger = DataLogger()
        self.statistics_analyzer = StatisticalAnalyzer()
        self.predictive_maintenance = PredictiveMaintenance()
        self.alarm_manager = AlarmManager()
        self.report_generator = ReportGenerator(self.data_logger)
        self.performance_monitor = PerformanceMonitor()
        self.data_visualization = DataVisualization(self.data_logger)
        self.health_monitor = SystemHealthMonitor()
        self.backup_manager = BackupManager(self.data_logger)

        # 获取标准范围从配置
        self.standards = {}
        for param, range_vals in self.config_manager.get('standards').items():
            self.standards[param] = tuple(range_vals)

        # 设置日志
        self.setup_logging()

        # 启动后台任务
        self.start_background_tasks()

    def setup_logging(self):
        """设置日志系统"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('quality_control.log', encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)

    def start_background_tasks(self):
        """启动后台任务"""

        # 定期备份任务
        def backup_task():
            while True:
                time.sleep(self.config_manager.get('maintenance.backup_interval_hours', 24) * 3600)
                try:
                    backup_path = self.backup_manager.create_backup()
                    self.logger.info(f"自动备份完成: {backup_path}")
                    self.backup_manager.cleanup_old_backups()
                except Exception as e:
                    self.logger.error(f"自动备份失败: {e}")

        backup_thread = threading.Thread(target=backup_task, daemon=True)
        backup_thread.start()

    def get_production_data(self):
        """获取所有生产过程数据（模拟）"""
        data = {}
        # 添加一些真实的变化模式
        base_time = time.time()

        # 温度有日周期变化
        temp_variation = 5 * sin(2 * pi * base_time / 86400)  # 日周期
        data['temperature'] = random.uniform(160, 240) + temp_variation

        # 压力随温度变化
        temp_factor = (data['temperature'] - 200) / 40
        data['pressure'] = random.uniform(60, 140) + temp_factor * 10

        # 其他参数
        data['density'] = random.uniform(1.0, 1.7)
        data['humidity'] = random.uniform(30, 70) + 10 * sin(2 * pi * base_time / 43200)  # 半日周期
        data['speed'] = random.uniform(1000, 2000)
        data['vibration'] = random.uniform(0.3, 2.5) + random.random() * 0.5  # 添加噪声
        data['flow_rate'] = random.uniform(40, 90)
        data['voltage'] = random.uniform(210, 250) + random.choice([-5, 0, 5])  # 电压波动
        data['current'] = random.uniform(8, 18)
        data['rpm'] = random.uniform(2600, 3400)
        data['torque'] = random.uniform(130, 220)
        data['power'] = random.uniform(35, 75)

        # 添加偶发异常
        if random.random() < 0.05:  # 5%概率出现异常
            anomaly_param = random.choice(list(data.keys()))
            if random.random() < 0.5:
                data[anomaly_param] *= 1.3  # 偏高
            else:
                data[anomaly_param] *= 0.7  # 偏低

        return data
    
    def section1_thermal_control(self, x, y, z):
        """第1类: 热力控制系统 (温度x, 压力y, 密度z) - 80个if语句"""
        actions = []

        # 基础范围检查 (1-20)
        if self.standards['temperature'][0] < x < self.standards['temperature'][1]:
            actions.append("热力控制：温度在标准范围内")
        if self.standards['pressure'][0] < y < self.standards['pressure'][1]:
            actions.append("热力控制：压力在标准范围内")
        if self.standards['density'][0] < z < self.standards['density'][1]:
            actions.append("热力控制：密度在标准范围内")
        if x > 190 and x < 210 and y > 90 and y < 110:
            actions.append("热力控制：温度压力在核心区间")
        if x > 185 and x < 215 and z > 1.25 and z < 1.45:
            actions.append("热力控制：温度密度匹配良好")
        if y > 85 and y < 115 and z > 1.25 and z < 1.45:
            actions.append("热力控制：压力密度协调稳定")
        if x > 195 and x < 205:
            actions.append("热力控制：温度精确控制")
        if y > 95 and y < 105:
            actions.append("热力控制：压力精确控制")
        if z > 1.3 and z < 1.4:
            actions.append("热力控制：密度精确控制")
        if abs(x - 200) < 8:
            actions.append("热力控制：温度稳定性良好")
        if abs(y - 100) < 8:
            actions.append("热力控制：压力稳定性良好")
        if abs(z - 1.35) < 0.08:
            actions.append("热力控制：密度稳定性良好")
        if x > self.standards['temperature'][0] + 10:
            actions.append("热力控制：温度安全缓冲区内")
        if y > self.standards['pressure'][0] + 10:
            actions.append("热力控制：压力安全缓冲区内")
        if z > self.standards['density'][0] + 0.1:
            actions.append("热力控制：密度安全缓冲区内")
        if x < self.standards['temperature'][1] - 10:
            actions.append("热力控制：温度上限安全")
        if y < self.standards['pressure'][1] - 10:
            actions.append("热力控制：压力上限安全")
        if z < self.standards['density'][1] - 0.1:
            actions.append("热力控制：密度上限安全")
        if x > 188 and x < 212 and y > 88 and y < 112 and z > 1.28 and z < 1.42:
            actions.append("热力控制：全系统协调运行")
        if (x + y + z * 100) > 420 and (x + y + z * 100) < 460:
            actions.append("热力控制：综合热力指标正常")

        # 比例关系检查 (21-40)
        if x / y > 1.8 and x / y < 2.2:
            actions.append("热力控制：温压比例理想")
        if x / z > 130 and x / z < 170:
            actions.append("热力控制：温密比例正常")
        if y / z > 60 and y / z < 90:
            actions.append("热力控制：压密比例适当")
        if x * y > 16000 and x * y < 24000:
            actions.append("热力控制：热力功率正常")
        if x * z > 240 and x * z < 320:
            actions.append("热力控制：温度密度积正常")
        if y * z > 100 and y * z < 170:
            actions.append("热力控制：压力密度积正常")
        if (x + y) / 2 > 145 and (x + y) / 2 < 155:
            actions.append("热力控制：温压平均值理想")
        if x - y > 60 and x - y < 120:
            actions.append("热力控制：温压差值正常")
        if abs(x / y - 2.0) < 0.3:
            actions.append("热力控制：温压比例稳定")
        if x / (y + 50) > 1.2 and x / (y + 50) < 1.8:
            actions.append("热力控制：调整温压比正常")
        if y / (z * 50) > 1.2 and y / (z * 50) < 1.8:
            actions.append("热力控制：调整压密比正常")
        if z / (x / 200) > 1.2 and z / (x / 200) < 1.8:
            actions.append("热力控制：调整密温比正常")
        if (x * y * z) > 25000 and (x * y * z) < 35000:
            actions.append("热力控制：三元热力积正常")
        if (x + y + z) / 3 > 108 and (x + y + z) / 3 < 118:
            actions.append("热力控制：参数平均值正常")
        if x ** 2 + y ** 2 + z ** 2 * 10000 > 50000:
            actions.append("热力控制：参数能量总和正常")
        if sqrt((x - 200) ** 2 + (y - 100) ** 2) < 15:
            actions.append("热力控制：温压距离理想点近")
        if (x - 180) * (y - 80) > 200 and (x - 180) * (y - 80) < 1000:
            actions.append("热力控制：温压偏差积正常")
        if abs((x + y) - 300) < 20:
            actions.append("热力控制：温压总和适当")
        if x / 200 > 0.9 and x / 200 < 1.1:
            actions.append("热力控制：温度相对值正常")
        if y / 100 > 0.9 and y / 100 < 1.1:
            actions.append("热力控制：压力相对值正常")

        # 临界监控 (41-60)
        if x < 175:
            actions.append("热力控制：温度偏低需要调整")
            self.adjustments.append("提高加热功率")
        if x > 225:
            actions.append("热力控制：温度偏高需要调整")
            self.adjustments.append("降低加热功率")
        if y < 75:
            actions.append("热力控制：压力偏低需要调整")
            self.adjustments.append("增加压力设定")
        if y > 125:
            actions.append("热力控制：压力偏高需要调整")
            self.adjustments.append("降低压力设定")
        if z < 1.15:
            actions.append("热力控制：密度偏低需要调整")
            self.adjustments.append("调整材料配比")
        if z > 1.55:
            actions.append("热力控制：密度偏高需要调整")
            self.adjustments.append("优化材料配比")
        if x < 170 or x > 230:
            actions.append("热力控制：温度在危险区间")
            self.alert_status = "high"
        if y < 70 or y > 130:
            actions.append("热力控制：压力在危险区间")
            self.alert_status = "high"
        if z < 1.1 or z > 1.6:
            actions.append("热力控制：密度在危险区间")
            self.alert_status = "high"
        if x < 165 and y < 75:
            actions.append("热力控制：温压同时偏低")
            self.alert_status = "medium"
        if x > 235 and y > 125:
            actions.append("热力控制：温压同时偏高")
            self.alert_status = "medium"
        if x < 170 and z < 1.15:
            actions.append("热力控制：温密同时偏低")
            self.alert_status = "medium"
        if x > 230 and z > 1.55:
            actions.append("热力控制：温密同时偏高")
            self.alert_status = "medium"
        if y < 75 and z < 1.15:
            actions.append("热力控制：压密同时偏低")
            self.alert_status = "medium"
        if y > 125 and z > 1.55:
            actions.append("热力控制：压密同时偏高")
            self.alert_status = "medium"
        if abs(x / y - 2.0) > 0.5:
            actions.append("热力控制：温压比例失衡")
            self.alert_status = "medium"
        if x * y < 14000 or x * y > 26000:
            actions.append("热力控制：热力功率异常")
            self.alert_status = "medium"
        if (x + y + z * 100) < 400 or (x + y + z * 100) > 480:
            actions.append("热力控制：综合指标异常")
            self.alert_status = "medium"
        if x < 160 or x > 240:
            actions.append("热力控制：温度极限范围")
            self.alert_status = "emergency"
        if y < 60 or y > 140:
            actions.append("热力控制：压力极限范围")
            self.alert_status = "emergency"

        # 优化控制 (61-80)
        if x > 198 and x < 202 and y > 98 and y < 102:
            actions.append("热力控制：温压精密控制优秀")
        if z > 1.33 and z < 1.37:
            actions.append("热力控制：密度精密控制优秀")
        if abs(x - 200) < 3 and abs(y - 100) < 3:
            actions.append("热力控制：核心参数超精密")
        if abs(z - 1.35) < 0.03:
            actions.append("热力控制：密度超精密控制")
        if x / y > 1.95 and x / y < 2.05:
            actions.append("热力控制：温压比最优化")
        if x * y > 19000 and x * y < 21000:
            actions.append("热力控制：热力功率最优")
        if (x + y + z * 100) > 430 and (x + y + z * 100) < 440:
            actions.append("热力控制：综合指标最优")
        if (x * y * z) > 29000 and (x * y * z) < 31000:
            actions.append("热力控制：三元积最优")
        if ((x - 200) ** 2 + (y - 100) ** 2) < 25:
            actions.append("热力控制：温压距离最优")
        if abs((x + y) / 2 - 150) < 2:
            actions.append("热力控制：温压平均最优")
        if x / (z * 100) > 1.45 and x / (z * 100) < 1.55:
            actions.append("热力控制：温密比最优")
        if y / (z * 50) > 1.45 and y / (z * 50) < 1.55:
            actions.append("热力控制：压密比最优")
        if (x - 180) / 40 > 0.45 and (x - 180) / 40 < 0.55:
            actions.append("热力控制：温度相对位置最优")
        if (y - 80) / 40 > 0.45 and (y - 80) / 40 < 0.55:
            actions.append("热力控制：压力相对位置最优")
        if (z - 1.2) / 0.3 > 0.45 and (z - 1.2) / 0.3 < 0.55:
            actions.append("热力控制：密度相对位置最优")
        if min(x - 180, 220 - x) > 15:
            actions.append("热力控制：温度安全边距充足")
        if min(y - 80, 120 - y) > 15:
            actions.append("热力控制：压力安全边距充足")
        if min(z - 1.2, 1.5 - z) > 0.12:
            actions.append("热力控制：密度安全边距充足")
        if abs(x * y - 20000) < 500:
            actions.append("热力控制：热力功率精确")
        if abs((x * y * z) / 30000 - 1) < 0.05:
            actions.append("热力控制：归一化积完美")

        return actions

    def section2_environmental_control(self, x, y, z):
        """第2类: 环境控制系统 (湿度x, 速度y, 振动z) - 80个if语句"""
        actions = []

        # 基础范围检查 (1-20) - 调整振动范围为1-100
        if 30 < x < 70:
            actions.append("环境控制：湿度在标准范围内")
        if 1000 < y < 2000:
            actions.append("环境控制：速度在标准范围内")
        if 1 < z < 100:  # 振动范围调整为1-100
            actions.append("环境控制：振动在标准范围内")
        if x > 45 and x < 55 and y > 1400 and y < 1600:
            actions.append("环境控制：湿度速度在核心区间")
        if x > 42 and x < 58 and z > 40 and z < 60:  # 振动核心区间40-60
            actions.append("环境控制：湿度振动匹配良好")
        if y > 1300 and y < 1700 and z > 40 and z < 60:  # 振动核心区间40-60
            actions.append("环境控制：速度振动协调稳定")
        if x > 48 and x < 52:
            actions.append("环境控制：湿度精确控制")
        if y > 1480 and y < 1520:
            actions.append("环境控制：速度精确控制")
        if z > 45 and z < 55:  # 振动精确控制范围45-55
            actions.append("环境控制：振动精确控制")
        if abs(x - 50) < 4:
            actions.append("环境控制：湿度稳定性良好")
        if abs(y - 1500) < 80:
            actions.append("环境控制：速度稳定性良好")
        if abs(z - 50) < 10:  # 振动中心值50
            actions.append("环境控制：振动稳定性良好")
        if x > 35:
            actions.append("环境控制：湿度安全缓冲区内")
        if y > 1100:
            actions.append("环境控制：速度安全缓冲区内")
        if z > 20:  # 振动安全缓冲区
            actions.append("环境控制：振动安全缓冲区内")
        if x < 65:
            actions.append("环境控制：湿度上限安全")
        if y < 1900:
            actions.append("环境控制：速度上限安全")
        if z < 80:  # 振动上限安全
            actions.append("环境控制：振动上限安全")
        if x > 43 and x < 57 and y > 1350 and y < 1650 and z > 35 and z < 65:  # 振动范围35-65
            actions.append("环境控制：全系统协调运行")
        if (x + y / 30 + z / 2) > 95 and (x + y / 30 + z / 2) < 125:  # 调整z系数
            actions.append("环境控制：综合环境指标正常")

        # 比例关系检查 (21-40) - 调整振动相关比例
        if x / (y / 30) > 0.8 and x / (y / 30) < 1.2:
            actions.append("环境控制：湿速比例理想")
        if x / (z / 2) > 1.8 and x / (z / 2) < 2.8:  # 调整z系数
            actions.append("环境控制：湿振比例正常")
        if (y / 30) / (z / 50) > 35 and (y / 30) / (z / 50) < 65:  # 调整z系数
            actions.append("环境控制：速振比例适当")
        if x * (y / 30) > 2000 and x * (y / 30) < 3200:
            actions.append("环境控制：湿速功率正常")
        if x * (z / 50) > 30 and x * (z / 50) < 90:  # 调整z系数
            actions.append("环境控制：湿度振动积正常")
        if (y / 30) * (z / 50) > 40 and (y / 30) * (z / 50) < 80:  # 调整z系数
            actions.append("环境控制：速度振动积正常")
        if (x + y / 30) / 2 > 45 and (x + y / 30) / 2 < 55:
            actions.append("环境控制：湿速平均值理想")
        if x - y / 30 > -10 and x - y / 30 < 10:
            actions.append("环境控制：湿速差值正常")
        if abs(x / (y / 30) - 1.0) < 0.3:
            actions.append("环境控制：湿速比例稳定")
        if x / ((y / 30) + 20) > 0.8 and x / ((y / 30) + 20) < 1.4:
            actions.append("环境控制：调整湿速比正常")
        if (y / 30) / (z / 50 + 0.5) > 20 and (y / 30) / (z / 50 + 0.5) < 40:  # 调整z系数
            actions.append("环境控制：调整速振比正常")
        if (z / 50) / (x / 50) > 0.8 and (z / 50) / (x / 50) < 2.2:  # 调整z系数
            actions.append("环境控制：调整振湿比正常")
        if (x * (y / 30) * (z / 50)) > 1000 and (x * (y / 30) * (z / 50)) < 4000:  # 调整z系数
            actions.append("环境控制：三元环境积正常")
        if (x + y / 30 + z / 2) / 3 > 35 and (x + y / 30 + z / 2) / 3 < 45:  # 调整z系数
            actions.append("环境控制：参数平均值正常")
        if x ** 2 + (y / 30) ** 2 + (z / 2) ** 2 > 2500:  # 调整z系数
            actions.append("环境控制：参数能量总和正常")
        if sqrt((x - 50) ** 2 + (y / 30 - 50) ** 2) < 8:
            actions.append("环境控制：湿速距离理想点近")
        if (x - 40) * (y / 30 - 40) > 50 and (x - 40) * (y / 30 - 40) < 200:
            actions.append("环境控制：湿速偏差积正常")
        if abs((x + y / 30) - 100) < 15:
            actions.append("环境控制：湿速总和适当")
        if x / 50 > 0.85 and x / 50 < 1.15:
            actions.append("环境控制：湿度相对值正常")
        if (y / 30) / 50 > 0.85 and (y / 30) / 50 < 1.15:
            actions.append("环境控制：速度相对值正常")

        # 临界监控 (41-60) - 调整振动阈值
        if x < 35:
            actions.append("环境控制：湿度偏低需要调整")
            self.adjustments.append("增加加湿功率")
        if x > 65:
            actions.append("环境控制：湿度偏高需要调整")
            self.adjustments.append("增加除湿功率")
        if y < 1100:
            actions.append("环境控制：速度偏低需要调整")
            self.adjustments.append("提高运行速度")
        if y > 1900:
            actions.append("环境控制：速度偏高需要调整")
            self.adjustments.append("降低运行速度")
        if z < 20:  # 振动偏低阈值
            actions.append("环境控制：振动偏低需要调整")
            self.adjustments.append("检查振动源")
        if z > 80:  # 振动偏高阈值
            actions.append("环境控制：振动偏高需要调整")
            self.adjustments.append("增加减振措施")
        if x < 32 or x > 68:
            actions.append("环境控制：湿度在危险区间")
            self.alert_status = "high"
        if y < 1000 or y > 2000:
            actions.append("环境控制：速度在危险区间")
            self.alert_status = "high"
        if z < 15 or z > 85:  # 振动危险区间
            actions.append("环境控制：振动在危险区间")
            self.alert_status = "high"
        if x < 35 and y < 1200:
            actions.append("环境控制：湿速同时偏低")
            self.alert_status = "medium"
        if x > 65 and y > 1800:
            actions.append("环境控制：湿速同时偏高")
            self.alert_status = "medium"
        if x < 35 and z < 30:  # 振动偏低阈值
            actions.append("环境控制：湿振同时偏低")
            self.alert_status = "medium"
        if x > 65 and z > 70:  # 振动偏高阈值
            actions.append("环境控制：湿振同时偏高")
            self.alert_status = "medium"
        if y < 1200 and z < 30:  # 振动偏低阈值
            actions.append("环境控制：速振同时偏低")
            self.alert_status = "medium"
        if y > 1800 and z > 70:  # 振动偏高阈值
            actions.append("环境控制：速振同时偏高")
            self.alert_status = "medium"
        if abs(x / (y / 30) - 1.0) > 0.6:
            actions.append("环境控制：湿速比例失衡")
            self.alert_status = "medium"
        if x * (y / 30) < 1800 or x * (y / 30) > 3500:
            actions.append("环境控制：湿速功率异常")
            self.alert_status = "medium"
        if (x + y / 30 + z / 2) < 85 or (x + y / 30 + z / 2) > 135:  # 调整z系数
            actions.append("环境控制：综合指标异常")
            self.alert_status = "medium"
        if x < 30 or x > 70:
            actions.append("环境控制：湿度极限范围")
            self.alert_status = "emergency"
        if y < 950 or y > 2050:
            actions.append("环境控制：速度极限范围")
            self.alert_status = "emergency"

        # 优化控制 (61-80) - 调整振动优化参数
        if x > 49 and x < 51 and y > 1490 and y < 1510:
            actions.append("环境控制：湿速精密控制优秀")
        if z > 48 and z < 52:  # 振动精密控制范围48-52
            actions.append("环境控制：振动精密控制优秀")
        if abs(x - 50) < 1.5 and abs(y / 30 - 50) < 1.5:
            actions.append("环境控制：核心参数超精密")
        if abs(z - 50) < 2:  # 振动超精密控制
            actions.append("环境控制：振动超精密控制")
        if x / (y / 30) > 0.95 and x / (y / 30) < 1.05:
            actions.append("环境控制：湿速比最优化")
        if x * (y / 30) > 2400 and x * (y / 30) < 2600:
            actions.append("环境控制：湿速功率最优")
        if (x + y / 30 + z / 2) > 108 and (x + y / 30 + z / 2) < 112:  # 调整z系数
            actions.append("环境控制：综合指标最优")
        if (x * (y / 30) * (z / 50)) > 2800 and (x * (y / 30) * (z / 50)) < 3200:  # 调整z系数
            actions.append("环境控制：三元积最优")
        if ((x - 50) ** 2 + (y / 30 - 50) ** 2) < 9:
            actions.append("环境控制：湿速距离最优")
        if abs((x + y / 30) / 2 - 50) < 1:
            actions.append("环境控制：湿速平均最优")
        if x / (z / 2) > 2.2 and x / (z / 2) < 2.4:  # 调整z系数
            actions.append("环境控制：湿振比最优")
        if (y / 30) / (z / 50) > 38 and (y / 30) / (z / 50) < 42:  # 调整z系数
            actions.append("环境控制：速振比最优")
        if (x - 40) / 20 > 0.45 and (x - 40) / 20 < 0.55:
            actions.append("环境控制：湿度相对位置最优")
        if (y / 30 - 40) / 20 > 0.45 and (y / 30 - 40) / 20 < 0.55:
            actions.append("环境控制：速度相对位置最优")
        if (z - 20) / 60 > 0.45 and (z - 20) / 60 < 0.55:  # 振动相对位置
            actions.append("环境控制：振动相对位置最优")
        if min(x - 40, 60 - x) > 8:
            actions.append("环境控制：湿度安全边距充足")
        if min(y - 1200, 1800 - y) > 150:
            actions.append("环境控制：速度安全边距充足")
        if min(z - 20, 80 - z) > 15:  # 振动安全边距
            actions.append("环境控制：振动安全边距充足")
        if abs(x * (y / 30) - 2500) < 50:
            actions.append("环境控制：湿速功率精确")
        if abs((x * (y / 30) * (z / 50)) / 3000 - 1) < 0.05:  # 调整z系数
            actions.append("环境控制：归一化积完美")

        return actions

    def section3_electrical_control(self, x, y, z):
        """第3类: 电气控制系统 (流量x, 电压y, 电流z) - 80个if语句"""
        actions = []

        # 基础范围检查 (1-20)
        if self.standards['flow_rate'][0] < x < self.standards['flow_rate'][1]:
            actions.append("电气控制：流量在标准范围内")
        if self.standards['voltage'][0] < y < self.standards['voltage'][1]:
            actions.append("电气控制：电压在标准范围内")
        if self.standards['current'][0] < z < self.standards['current'][1]:
            actions.append("电气控制：电流在标准范围内")
        if x > 60 and x < 70 and y > 225 and y < 235:
            actions.append("电气控制：流量电压在核心区间")
        if x > 55 and x < 75 and z > 11 and z < 14:
            actions.append("电气控制：流量电流匹配良好")
        if y > 225 and y < 235 and z > 11 and z < 14:
            actions.append("电气控制：电压电流协调稳定")
        if x > 63 and x < 67:
            actions.append("电气控制：流量精确控制")
        if y > 228 and y < 232:
            actions.append("电气控制：电压精确控制")
        if z > 12 and z < 13:
            actions.append("电气控制：电流精确控制")
        if abs(x - 65) < 3:
            actions.append("电气控制：流量稳定性良好")
        if abs(y - 230) < 4:
            actions.append("电气控制：电压稳定性良好")
        if abs(z - 12.5) < 0.8:
            actions.append("电气控制：电流稳定性良好")
        if x > self.standards['flow_rate'][0] + 5:
            actions.append("电气控制：流量安全缓冲区内")
        if y > self.standards['voltage'][0] + 8:
            actions.append("电气控制：电压安全缓冲区内")
        if z > self.standards['current'][0] + 1:
            actions.append("电气控制：电流安全缓冲区内")
        if x < self.standards['flow_rate'][1] - 5:
            actions.append("电气控制：流量上限安全")
        if y < self.standards['voltage'][1] - 8:
            actions.append("电气控制：电压上限安全")
        if z < self.standards['current'][1] - 1:
            actions.append("电气控制：电流上限安全")
        if x > 58 and x < 72 and y > 223 and y < 237 and z > 10.5 and z < 14.5:
            actions.append("电气控制：全系统协调运行")
        if (x + y / 10 + z * 5) > 105 and (x + y / 10 + z * 5) < 125:
            actions.append("电气控制：综合电气指标正常")

        # 比例关系检查 (21-40)
        if x / (y / 10) > 2.5 and x / (y / 10) < 3.5:
            actions.append("电气控制：流压比例理想")
        if x / z > 4.5 and x / z < 6.5:
            actions.append("电气控制：流流比例正常")
        if (y / 10) / z > 1.6 and (y / 10) / z < 2.4:
            actions.append("电气控制：压流比例适当")
        if x * (y / 10) > 1400 and x * (y / 10) < 1800:
            actions.append("电气控制：流压功率正常")
        if x * z > 600 and x * z < 1000:
            actions.append("电气控制：流量电流积正常")
        if (y / 10) * z > 250 and (y / 10) * z < 350:
            actions.append("电气控制：电压电流积正常")
        if (x + y / 10) / 2 > 40 and (x + y / 10) / 2 < 50:
            actions.append("电气控制：流压平均值理想")
        if x - y / 10 > 30 and x - y / 10 < 50:
            actions.append("电气控制：流压差值正常")
        if abs(x / (y / 10) - 3.0) < 0.4:
            actions.append("电气控制：流压比例稳定")
        if x / ((y / 10) + 10) > 1.5 and x / ((y / 10) + 10) < 2.5:
            actions.append("电气控制：调整流压比正常")
        if (y / 10) / (z + 2) > 1.4 and (y / 10) / (z + 2) < 2.2:
            actions.append("电气控制：调整压流比正常")
        if z / (x / 50) > 8 and z / (x / 50) < 12:
            actions.append("电气控制：调整流流比正常")
        if (x * (y / 10) * z) > 18000 and (x * (y / 10) * z) < 26000:
            actions.append("电气控制：三元电气积正常")
        if (x + y / 10 + z * 5) / 3 > 35 and (x + y / 10 + z * 5) / 3 < 45:
            actions.append("电气控制：参数平均值正常")
        if x ** 2 + (y / 10) ** 2 + (z * 5) ** 2 > 4500:
            actions.append("电气控制：参数能量总和正常")
        if sqrt((x - 65) ** 2 + (y / 10 - 23) ** 2) < 5:
            actions.append("电气控制：流压距离理想点近")
        if (x - 50) * (y / 10 - 22) > 20 and (x - 50) * (y / 10 - 22) < 80:
            actions.append("电气控制：流压偏差积正常")
        if abs((x + y / 10) - 88) < 8:
            actions.append("电气控制：流压总和适当")
        if x / 65 > 0.85 and x / 65 < 1.15:
            actions.append("电气控制：流量相对值正常")
        if (y / 10) / 23 > 0.85 and (y / 10) / 23 < 1.15:
            actions.append("电气控制：电压相对值正常")

        # 临界监控 (41-60)
        if x < 45:
            actions.append("电气控制：流量偏低需要调整")
            self.adjustments.append("增加泵流量")
        if x > 85:
            actions.append("电气控制：流量偏高需要调整")
            self.adjustments.append("降低泵流量")
        if y < 215:
            actions.append("电气控制：电压偏低需要调整")
            self.adjustments.append("提高电压设定")
        if y > 245:
            actions.append("电气控制：电压偏高需要调整")
            self.adjustments.append("降低电压设定")
        if z < 8.5:
            actions.append("电气控制：电流偏低需要调整")
            self.adjustments.append("检查负载连接")
        if z > 16.5:
            actions.append("电气控制：电流偏高需要调整")
            self.adjustments.append("检查短路风险")
        if x < 42 or x > 88:
            actions.append("电气控制：流量在危险区间")
            self.alert_status = "high"
        if y < 212 or y > 248:
            actions.append("电气控制：电压在危险区间")
            self.alert_status = "high"
        if z < 8 or z > 17:
            actions.append("电气控制：电流在危险区间")
            self.alert_status = "high"
        if x < 45 and y < 220:
            actions.append("电气控制：流压同时偏低")
            self.alert_status = "medium"
        if x > 85 and y > 240:
            actions.append("电气控制：流压同时偏高")
            self.alert_status = "medium"
        if x < 45 and z < 9:
            actions.append("电气控制：流流同时偏低")
            self.alert_status = "medium"
        if x > 85 and z > 16:
            actions.append("电气控制：流流同时偏高")
            self.alert_status = "medium"
        if y < 220 and z < 9:
            actions.append("电气控制：压流同时偏低")
            self.alert_status = "medium"
        if y > 240 and z > 16:
            actions.append("电气控制：压流同时偏高")
            self.alert_status = "medium"
        if abs(x / (y / 10) - 3.0) > 0.8:
            actions.append("电气控制：流压比例失衡")
            self.alert_status = "medium"
        if x * (y / 10) < 1200 or x * (y / 10) > 2000:
            actions.append("电气控制：流压功率异常")
            self.alert_status = "medium"
        if (x + y / 10 + z * 5) < 95 or (x + y / 10 + z * 5) > 135:
            actions.append("电气控制：综合指标异常")
            self.alert_status = "medium"
        if x < 40 or x > 90:
            actions.append("电气控制：流量极限范围")
            self.alert_status = "emergency"
        if y < 210 or y > 250:
            actions.append("电气控制：电压极限范围")
            self.alert_status = "emergency"

        # 优化控制 (61-80)
        if x > 64 and x < 66 and y > 229 and y < 231:
            actions.append("电气控制：流压精密控制优秀")
        if z > 12.3 and z < 12.7:
            actions.append("电气控制：电流精密控制优秀")
        if abs(x - 65) < 1 and abs(y / 10 - 23) < 0.5:
            actions.append("电气控制：核心参数超精密")
        if abs(z - 12.5) < 0.2:
            actions.append("电气控制：电流超精密控制")
        if x / (y / 10) > 2.9 and x / (y / 10) < 3.1:
            actions.append("电气控制：流压比最优化")
        if x * (y / 10) > 1590 and x * (y / 10) < 1610:
            actions.append("电气控制：流压功率最优")
        if (x + y / 10 + z * 5) > 113 and (x + y / 10 + z * 5) < 117:
            actions.append("电气控制：综合指标最优")
        if (x * (y / 10) * z) > 21500 and (x * (y / 10) * z) < 22500:
            actions.append("电气控制：三元积最优")
        if ((x - 65) ** 2 + (y / 10 - 23) ** 2) < 4:
            actions.append("电气控制：流压距离最优")
        if abs((x + y / 10) / 2 - 44) < 1:
            actions.append("电气控制：流压平均最优")
        if x / z > 5.1 and x / z < 5.3:
            actions.append("电气控制：流流比最优")
        if (y / 10) / z > 1.8 and (y / 10) / z < 1.9:
            actions.append("电气控制：压流比最优")
        if (x - 50) / 30 > 0.45 and (x - 50) / 30 < 0.55:
            actions.append("电气控制：流量相对位置最优")
        if (y / 10 - 22) / 2 > 0.45 and (y / 10 - 22) / 2 < 0.55:
            actions.append("电气控制：电压相对位置最优")
        if (z - 10) / 5 > 0.45 and (z - 10) / 5 < 0.55:
            actions.append("电气控制：电流相对位置最优")
        if min(x - 50, 80 - x) > 12:
            actions.append("电气控制：流量安全边距充足")
        if min(y - 220, 240 - y) > 8:
            actions.append("电气控制：电压安全边距充足")
        if min(z - 10, 15 - z) > 2:
            actions.append("电气控制：电流安全边距充足")
        if abs(x * (y / 10) - 1600) < 20:
            actions.append("电气控制：流压功率精确")
        if abs((x * (y / 10) * z) / 22000 - 1) < 0.05:
            actions.append("电气控制：归一化积完美")

        return actions

    def section4_mechanical_control(self, x, y, z):
        """第4类: 机械控制系统 (转速x, 扭矩y, 功率z) - 80个if语句"""
        actions = []



        # 基础范围检查 (1-20)
        if self.standards['rpm'][0] < x < self.standards['rpm'][1]:
            actions.append("机械控制：转速在标准范围内")
        if self.standards['torque'][0] < y < self.standards['torque'][1]:
            actions.append("机械控制：扭矩在标准范围内")
        if self.standards['power'][0] < z < self.standards['power'][1]:
            actions.append("机械控制：功率在标准范围内")
        if x > 2900 and x < 3100 and y > 160 and y < 190:
            actions.append("机械控制：转速扭矩在核心区间")
        if x > 2850 and x < 3150 and z > 50 and z < 60:
            actions.append("机械控制：转速功率匹配良好")
        if y > 160 and y < 190 and z > 50 and z < 60:
            actions.append("机械控制：扭矩功率协调稳定")
        if x > 2980 and x < 3020:
            actions.append("机械控制：转速精确控制")
        if y > 173 and y < 177:
            actions.append("机械控制：扭矩精确控制")
        if z > 53 and z < 57:
            actions.append("机械控制：功率精确控制")
        if abs(x - 3000) < 50:
            actions.append("机械控制：转速稳定性良好")
        if abs(y - 175) < 8:
            actions.append("机械控制：扭矩稳定性良好")
        if abs(z - 55) < 3:
            actions.append("机械控制：功率稳定性良好")
        if x > self.standards['rpm'][0] + 80:
            actions.append("机械控制：转速安全缓冲区内")
        if y > self.standards['torque'][0] + 10:
            actions.append("机械控制：扭矩安全缓冲区内")
        if z > self.standards['power'][0] + 5:
            actions.append("机械控制：功率安全缓冲区内")
        if x < self.standards['rpm'][1] - 80:
            actions.append("机械控制：转速上限安全")
        if y < self.standards['torque'][1] - 10:
            actions.append("机械控制：扭矩上限安全")
        if z < self.standards['power'][1] - 5:
            actions.append("机械控制：功率上限安全")
        if x > 2850 and x < 3150 and y > 155 and y < 195 and z > 48 and z < 62:
            actions.append("机械控制：全系统协调运行")
        if (x / 50 + y + z) > 250 and (x / 50 + y + z) < 290:
            actions.append("机械控制：综合机械指标正常")

        # 比例关系检查 (21-40)
        if x / (y * 10) > 1.5 and x / (y * 10) < 2.0:
            actions.append("机械控制：转扭比例理想")
        if x / z > 50 and x / z < 60:
            actions.append("机械控制：转功比例正常")
        if y / z > 2.8 and y / z < 3.5:
            actions.append("机械控制：扭功比例适当")
        if x * y / 1000 > 500 and x * y / 1000 < 600:
            actions.append("机械控制：转扭功率正常")
        if x * z / 100 > 1500 and x * z / 100 < 2000:
            actions.append("机械控制：转速功率积正常")
        if y * z > 8000 and y * z < 12000:
            actions.append("机械控制：扭矩功率积正常")
        if (x / 50 + y) / 2 > 85 and (x / 50 + y) / 2 < 95:
            actions.append("机械控制：转扭平均值理想")
        if x / 50 - y > -110 and x / 50 - y > -130:
            actions.append("机械控制：转扭差值正常")
        if abs(x / (y * 10) - 1.7) < 0.2:
            actions.append("机械控制：转扭比例稳定")
        if x / (y * 10 + 50) > 1.2 and x / (y * 10 + 50) < 1.6:
            actions.append("机械控制：调整转扭比正常")
        if y / (z + 20) > 2.2 and y / (z + 20) < 2.8:
            actions.append("机械控制：调整扭功比正常")
        if z / (x / 100) > 1.6 and z / (x / 100) < 2.2:
            actions.append("机械控制：调整功转比正常")
        if (x * y * z / 10000) > 25 and (x * y * z / 10000) < 35:
            actions.append("机械控制：三元机械积正常")
        if (x / 50 + y + z) / 3 > 85 and (x / 50 + y + z) / 3 < 95:
            actions.append("机械控制：参数平均值正常")
        if (x / 50) ** 2 + y ** 2 + z ** 2 > 35000:
            actions.append("机械控制：参数能量总和正常")
        if sqrt((x / 50 - 60) ** 2 + (y - 175) ** 2) < 15:
            actions.append("机械控制：转扭距离理想点近")
        if (x / 50 - 56) * (y - 150) > 200 and (x / 50 - 56) * (y - 150) < 800:
            actions.append("机械控制：转扭偏差积正常")
        if abs((x / 50 + y) - 235) < 15:
            actions.append("机械控制：转扭总和适当")
        if (x / 50) / 60 > 0.85 and (x / 50) / 60 < 1.15:
            actions.append("机械控制：转速相对值正常")
        if y / 175 > 0.85 and y / 175 < 1.15:
            actions.append("机械控制：扭矩相对值正常")

        # 临界监控 (41-60)
        if x < 2700:
            actions.append("机械控制：转速偏低需要调整")
            self.adjustments.append("提高电机转速")
        if x > 3300:
            actions.append("机械控制：转速偏高需要调整")
            self.adjustments.append("降低电机转速")
        if y < 140:
            actions.append("机械控制：扭矩偏低需要调整")
            self.adjustments.append("增加负载扭矩")
        if y > 210:
            actions.append("机械控制：扭矩偏高需要调整")
            self.adjustments.append("减少负载扭矩")
        if z < 40:
            actions.append("机械控制：功率偏低需要调整")
            self.adjustments.append("检查功率输出")
        if z > 70:
            actions.append("机械控制：功率偏高需要调整")
            self.adjustments.append("限制功率输出")
        if x < 2650 or x > 3350:
            actions.append("机械控制：转速在危险区间")
            self.alert_status = "high"
        if y < 135 or y > 215:
            actions.append("机械控制：扭矩在危险区间")
            self.alert_status = "high"
        if z < 38 or z > 72:
            actions.append("机械控制：功率在危险区间")
            self.alert_status = "high"
        if x < 2750 and y < 145:
            actions.append("机械控制：转扭同时偏低")
            self.alert_status = "medium"
        if x > 3250 and y > 205:
            actions.append("机械控制：转扭同时偏高")
            self.alert_status = "medium"
        if x < 2750 and z < 42:
            actions.append("机械控制：转功同时偏低")
            self.alert_status = "medium"
        if x > 3250 and z > 68:
            actions.append("机械控制：转功同时偏高")
            self.alert_status = "medium"
        if y < 145 and z < 42:
            actions.append("机械控制：扭功同时偏低")
            self.alert_status = "medium"
        if y > 205 and z > 68:
            actions.append("机械控制：扭功同时偏高")
            self.alert_status = "medium"
        if abs(x / (y * 10) - 1.7) > 0.4:
            actions.append("机械控制：转扭比例失衡")
            self.alert_status = "medium"
        if x * y / 1000 < 450 or x * y / 1000 > 650:
            actions.append("机械控制：转扭功率异常")
            self.alert_status = "medium"
        if (x / 50 + y + z) < 240 or (x / 50 + y + z) > 300:
            actions.append("机械控制：综合指标异常")
            self.alert_status = "medium"
        if x < 2600 or x > 3400:
            actions.append("机械控制：转速极限范围")
            self.alert_status = "emergency"
        if y < 130 or y > 220:
            actions.append("机械控制：扭矩极限范围")
            self.alert_status = "emergency"

        # 优化控制 (61-80)
        if x > 2990 and x < 3010 and y > 173 and y < 177:
            actions.append("机械控制：转扭精密控制优秀")
        if z > 54 and z < 56:
            actions.append("机械控制：功率精密控制优秀")
        if abs(x / 50 - 60) < 1 and abs(y - 175) < 2:
            actions.append("机械控制：核心参数超精密")
        if abs(z - 55) < 1:
            actions.append("机械控制：功率超精密控制")
        if x / (y * 10) > 1.68 and x / (y * 10) < 1.72:
            actions.append("机械控制：转扭比最优化")
        if x * y / 1000 > 520 and x * y / 1000 < 530:
            actions.append("机械控制：转扭功率最优")
        if (x / 50 + y + z) > 268 and (x / 50 + y + z) < 272:
            actions.append("机械控制：综合指标最优")
        if (x * y * z / 10000) > 29 and (x * y * z / 10000) < 31:
            actions.append("机械控制：三元积最优")
        if ((x / 50 - 60) ** 2 + (y - 175) ** 2) < 25:
            actions.append("机械控制：转扭距离最优")
        if abs((x / 50 + y) / 2 - 117.5) < 2:
            actions.append("机械控制：转扭平均最优")
        if x / z > 53 and x / z < 55:
            actions.append("机械控制：转功比最优")
        if y / z > 3.1 and y / z < 3.3:
            actions.append("机械控制：扭功比最优")
        if (x / 50 - 56) / 8 > 0.45 and (x / 50 - 56) / 8 < 0.55:
            actions.append("机械控制：转速相对位置最优")
        if (y - 150) / 50 > 0.45 and (y - 150) / 50 < 0.55:
            actions.append("机械控制：扭矩相对位置最优")
        if (z - 45) / 20 > 0.45 and (z - 45) / 20 < 0.55:
            actions.append("机械控制：功率相对位置最优")
        if min(x - 2800, 3200 - x) > 120:
            actions.append("机械控制：转速安全边距充足")
        if min(y - 150, 200 - y) > 20:
            actions.append("机械控制：扭矩安全边距充足")
        if min(z - 45, 65 - z) > 8:
            actions.append("机械控制：功率安全边距充足")
        if abs(x * y / 1000 - 525) < 5:
            actions.append("机械控制：转扭功率精确")
        if abs((x * y * z / 10000) / 30 - 1) < 0.05:
            actions.append("机械控制：归一化积完美")

        return actions




    def section5_thermal_environment_hybrid(self, x, y, z):
        """第5类: 热环境混合控制 (温度x, 湿度y, 振动z) - 80个if语句"""
        actions = []

        # 混合系统基础检查 (1-20)
        if self.standards['temperature'][0] < x < self.standards['temperature'][1]:
            actions.append("混合控制：温度基础参数正常")
        if self.standards['humidity'][0] < y < self.standards['humidity'][1]:
            actions.append("混合控制：湿度基础参数正常")
        if self.standards['vibration'][0] < z < self.standards['vibration'][1]:
            actions.append("混合控制：振动基础参数正常")
        if x > 190 and x < 210 and y > 45 and y < 55:
            actions.append("混合控制：温湿度核心区间协调")
        if x > 185 and x < 215 and z > 0.8 and z < 1.6:
            actions.append("混合控制：温度振动匹配良好")
        if y > 42 and y < 58 and z > 0.8 and z < 1.6:
            actions.append("混合控制：湿度振动协调稳定")
        if x > 195 and x < 205:
            actions.append("混合控制：温度精确控制")
        if y > 48 and y < 52:
            actions.append("混合控制：湿度精确控制")
        if z > 1.1 and z < 1.4:
            actions.append("混合控制：振动精确控制")
        if abs(x - 200) < 6:
            actions.append("混合控制：温度稳定性良好")
        if abs(y - 50) < 3:
            actions.append("混合控制：湿度稳定性良好")
        if abs(z - 1.25) < 0.15:
            actions.append("混合控制：振动稳定性良好")
        if x > 185 and y > 43 and z > 0.7:
            actions.append("混合控制：所有参数在安全范围")
        if x < 215 and y < 57 and z < 1.8:
            actions.append("混合控制：所有参数未超上限")
        if x + y + z * 100 > 360 and x + y + z * 100 < 420:
            actions.append("混合控制：综合环境指标正常")
        if x * y / 100 > 85 and x * y / 100 < 115:
            actions.append("混合控制：温湿度乘积正常")
        if x * z > 180 and x * z < 280:
            actions.append("混合控制：温度振动乘积正常")
        if y * z > 40 and y * z < 80:
            actions.append("混合控制：湿度振动乘积正常")
        if (x + y + z * 100) / 3 > 115 and (x + y + z * 100) / 3 < 135:
            actions.append("混合控制：平均环境参数正常")
        if sqrt(x ** 2 + y ** 2 + (z * 100) ** 2) > 220:
            actions.append("混合控制：环境向量模长正常")

        # 交互影响分析 (21-40)
        if x / y > 3.5 and x / y < 4.5:
            actions.append("混合控制：温湿比例理想")
        if x / (z * 100) > 1.4 and x / (z * 100) < 1.8:
            actions.append("混合控制：温振比例正常")
        if y / (z * 50) > 0.8 and y / (z * 50) < 1.2:
            actions.append("混合控制：湿振比例适当")
        if (x - 180) + (y - 40) > 15 and (x - 180) + (y - 40) < 35:
            actions.append("混合控制：温湿偏差和正常")
        if (x - 180) + (z - 0.5) * 100 > 25 and (x - 180) + (z - 0.5) * 100 < 45:
            actions.append("混合控制：温振偏差和正常")
        if (y - 40) + (z - 0.5) * 50 > 25 and (y - 40) + (z - 0.5) * 50 < 45:
            actions.append("混合控制：湿振偏差和正常")
        if abs((x - 200) - (y - 50) * 4) < 15:
            actions.append("混合控制：温湿偏差关系平衡")
        if abs((x - 200) - (z - 1.25) * 160) < 20:
            actions.append("混合控制：温振偏差关系平衡")
        if abs((y - 50) - (z - 1.25) * 40) < 8:
            actions.append("混合控制：湿振偏差关系平衡")
        if x / (y + 20) > 2.5 and x / (y + 20) < 3.5:
            actions.append("混合控制：调整温湿比正常")
        if y / (z + 0.5) > 20 and y / (z + 0.5) < 35:
            actions.append("混合控制：调整湿振比正常")
        if z / (x / 200) > 1.0 and z / (x / 200) < 1.6:
            actions.append("混合控制：调整振温比正常")
        if (x * y * z) > 8000 and (x * y * z) < 15000:
            actions.append("混合控制：三元环境积正常")
        if x + y * 4 + z * 160 > 540 and x + y * 4 + z * 160 < 620:
            actions.append("混合控制：加权环境和正常")
        if x ** 0.6 * y ** 0.3 * (z * 100) ** 0.1 > 60:
            actions.append("混合控制：加权几何平均正常")
        if (x - 200) * (y - 50) > -100 and (x - 200) * (y - 50) < 100:
            actions.append("混合控制：温湿偏差积平衡")
        if (x - 200) * (z - 1.25) > -15 and (x - 200) * (z - 1.25) < 15:
            actions.append("混合控制：温振偏差积平衡")
        if (y - 50) * (z - 1.25) > -6 and (y - 50) * (z - 1.25) < 6:
            actions.append("混合控制：湿振偏差积平衡")
        if abs(x / 200 + y / 50 + z / 1.25 - 3) < 0.3:
            actions.append("混合控制：归一化和接近理想")
        if abs((x / 200) * (y / 50) * (z / 1.25) - 1) < 0.2:
            actions.append("混合控制：归一化积接近理想")

        # 协调性评估 (41-60)
        if x > 192 and x < 208 and y > 47 and y < 53 and z > 1.1 and z < 1.4:
            actions.append("混合控制：全参数协调运行")
        if abs(x / y - 4.0) < 0.4 and abs(z - 1.25) < 0.1:
            actions.append("混合控制：温湿比例与振动协调")
        if x * y > 9500 and x * y < 10500 and z > 1.2 and z < 1.3:
            actions.append("混合控制：温湿积与振动协调")
        if (x + y) / 2 > 122 and (x + y) / 2 < 128 and z > 1.15 and z < 1.35:
            actions.append("混合控制：温湿平均与振动协调")
        if x - y > 145 and x - y < 155 and z > 1.1 and z < 1.4:
            actions.append("混合控制：温湿差与振动协调")
        if sqrt((x - 200) ** 2 + (y - 50) ** 2) < 8 and abs(z - 1.25) < 0.1:
            actions.append("混合控制：温湿距离与振动优秀")
        if x / 200 > 0.96 and x / 200 < 1.04 and y / 50 > 0.94 and y / 50 < 1.06:
            actions.append("混合控制：温湿相对值协调")
        if z / 1.25 > 0.92 and z / 1.25 < 1.08:
            actions.append("混合控制：振动相对值协调")
        if (x / 200 + y / 50 + z / 1.25) / 3 > 0.97 and (x / 200 + y / 50 + z / 1.25) / 3 < 1.03:
            actions.append("混合控制：归一化平均协调")
        if max(abs(x / 200 - 1), abs(y / 50 - 1), abs(z / 1.25 - 1)) < 0.06:
            actions.append("混合控制：最大偏差协调良好")
        if min(x / 200, y / 50, z / 1.25) > 0.95:
            actions.append("混合控制：最小相对值协调良好")
        if max(x / 200, y / 50, z / 1.25) < 1.05:
            actions.append("混合控制：最大相对值协调良好")
        if abs(max(x, y * 4, z * 160) - min(x, y * 4, z * 160)) < 30:
            actions.append("混合控制：标准化后范围协调")
        if (x + y + z * 100) > 375 and (x + y + z * 100) < 385:
            actions.append("混合控制：综合参数协调优秀")
        if (x * y * z) > 11000 and (x * y * z) < 13000:
            actions.append("混合控制：三元积协调优秀")
        if abs((x + y + z * 100) / 3 - 126.7) < 3:
            actions.append("混合控制：平均值协调优秀")
        if x > 196 and x < 204 and y > 49 and y < 51:
            actions.append("混合控制：温湿度超精密协调")
        if z > 1.22 and z < 1.28:
            actions.append("混合控制：振动超精密协调")
        if ((x - 200) ** 2 + (y - 50) ** 2 + (z - 1.25) ** 2 * 10000) < 100:
            actions.append("混合控制：三维距离协调完美")
        if abs(x / y / z - 160) < 10:
            actions.append("混合控制：连续比例协调完美")

        # 优化策略建议 (61-80)
        if x < 188 or y < 45 or z < 0.9:
            actions.append("混合控制：检测到参数偏低趋势")
            self.adjustments.append("预防性提升相关参数")
        if x > 212 or y > 55 or z > 1.6:
            actions.append("混合控制：检测到参数偏高趋势")
            self.adjustments.append("预防性降低相关参数")
        if abs(x / y - 4.0) > 0.3:
            actions.append("混合控制：温湿比例需要调整")
            self.adjustments.append("优化温湿度协调控制")
        if abs(x / (z * 100) - 1.6) > 0.15:
            actions.append("混合控制：温振比例需要调整")
            self.adjustments.append("优化温度振动协调控制")
        if abs(y / (z * 50) - 1.0) > 0.1:
            actions.append("混合控制：湿振比例需要调整")
            self.adjustments.append("优化湿度振动协调控制")
        if (x + y + z * 100) < 370 or (x + y + z * 100) > 410:
            actions.append("混合控制：综合指标需要调整")
            self.adjustments.append("全面优化环境参数")
        if (x * y * z) < 10000 or (x * y * z) > 14000:
            actions.append("混合控制：三元积需要调整")
            self.adjustments.append
        if x > 215 and y > 55:
            actions.append("混合控制：温湿度同时偏高")
            self.alert_status = "medium"("平衡三参数关系")
        if x < 185 and y < 45:
            actions.append("混合控制：温湿度同时偏高")
            self.alert_status = "medium"("平衡三参数关系")
        if x < 185 and z < 0.8:
            actions.append("混合控制：温度振动同时偏低")
            self.alert_status = "medium"
        if x > 215 and z > 1.7:
            actions.append("混合控制：温度振动同时偏高")
            self.alert_status = "medium"
        if y < 45 and z < 0.8:
            actions.append("混合控制：湿度振动同时偏低")
            self.alert_status = "medium"
        if y > 55 and z > 1.7:
            actions.append("混合控制：湿度振动同时偏高")
            self.alert_status = "medium"
        if x < 180 or x > 220:
            actions.append("混合控制：温度在临界范围")
            self.alert_status = "high"
        if y < 38 or y > 62:
            actions.append("混合控制：湿度在临界范围")
            self.alert_status = "high"
        if z < 0.4 or z > 2.2:
            actions.append("混合控制：振动在临界范围")
            self.alert_status = "high"
        if x < 175 and y < 40 and z < 0.6:
            actions.append("混合控制：所有参数严重偏低")
            self.alert_status = "emergency"
        if x > 225 and y > 60 and z > 2.0:
            actions.append("混合控制：所有参数严重偏高")
            self.alert_status = "emergency"
        if ((x - 200) ** 2 + (y - 50) ** 2 + (z - 1.25) ** 2 * 10000) > 500:
            actions.append("混合控制：三维偏离过大")
            self.alert_status = "high"
        if abs((x * y * z) / 12000 - 1) > 0.25:
            actions.append("混合控制：三元积偏离过大")
            self.alert_status = "medium"

        return actions

    def section6_electro_mechanical_hybrid(self, x, y, z):
        """第6类: 电机混合控制 (电压x, 转速y, 功率z) - 80个if语句"""
        actions = []



        # 电机系统基础检查 (1-20)
        if self.standards['voltage'][0] < x < self.standards['voltage'][1]:
            actions.append("电机控制：电压基础参数正常")
        if self.standards['rpm'][0] < y < self.standards['rpm'][1]:
            actions.append("电机控制：转速基础参数正常")
        if self.standards['power'][0] < z < self.standards['power'][1]:
            actions.append("电机控制：功率基础参数正常")
        if x > 225 and x < 235 and y > 2900 and y < 3100:
            actions.append("电机控制：电压转速核心区间协调")
        if x > 223 and x < 237 and z > 50 and z < 60:
            actions.append("电机控制：电压功率匹配良好")
        if y > 2850 and y < 3150 and z > 50 and z < 60:
            actions.append("电机控制：转速功率协调稳定")
        if x > 228 and x < 232:
            actions.append("电机控制：电压精确控制")
        if y > 2980 and y < 3020:
            actions.append("电机控制：转速精确控制")
        if z > 53 and z < 57:
            actions.append("电机控制：功率精确控制")
        if abs(x - 230) < 3:
            actions.append("电机控制：电压稳定性良好")
        if abs(y - 3000) < 40:
            actions.append("电机控制：转速稳定性良好")
        if abs(z - 55) < 2.5:
            actions.append("电机控制：功率稳定性良好")
        if x > 223 and y > 2850 and z > 48:
            actions.append("电机控制：所有参数在安全范围")
        if x < 237 and y < 3150 and z < 62:
            actions.append("电机控制：所有参数未超上限")
        if x / 10 + y / 100 + z > 110 and x / 10 + y / 100 + z < 130:
            actions.append("电机控制：综合电机指标正常")
        if x * y / 1000 > 650 and x * y / 1000 < 750:
            actions.append("电机控制：电压转速乘积正常")
        if x * z > 12000 and x * z > 14000:
            actions.append("电机控制：电压功率乘积正常")
        if y * z / 1000 > 150 and y * z / 1000 < 200:
            actions.append("电机控制：转速功率乘积正常")
        if (x / 10 + y / 100 + z) / 3 > 36 and (x / 10 + y / 100 + z) / 3 < 42:
            actions.append("电机控制：平均电机参数正常")
        if sqrt((x / 10) ** 2 + (y / 100) ** 2 + z ** 2) > 50:
            actions.append("电机控制：电机向量模长正常")

        # 电机特性分析 (21-40)
        if x / (y / 100) > 7.5 and x / (y / 100) < 8.5:
            actions.append("电机控制：电压转速比理想")
        if x / z > 4.0 and x / z < 4.5:
            actions.append("电机控制：电压功率比正常")
        if (y / 100) / z > 0.5 and (y / 100) / z < 0.6:
            actions.append("电机控制：转速功率比适当")
        if (x - 220) + (y - 2800) / 10 > 15 and (x - 220) + (y - 2800) / 10 < 35:
            actions.append("电机控制：电压转速偏差和正常")
        if (x - 220) + (z - 45) * 2 > 20 and (x - 220) + (z - 45) * 2 < 40:
            actions.append("电机控制：电压功率偏差和正常")
        if (y - 2800) / 10 + (z - 45) > 25 and (y - 2800) / 10 + (z - 45) < 45:
            actions.append("电机控制：转速功率偏差和正常")
        if abs((x - 230) - (y - 3000) / 50) < 8:
            actions.append("电机控制：电压转速偏差关系平衡")
        if abs((x - 230) - (z - 55) * 4) < 12:
            actions.append("电机控制：电压功率偏差关系平衡")
        if abs((y - 3000) / 50 - (z - 55)) < 3:
            actions.append("电机控制：转速功率偏差关系平衡")
        if x / (y / 100 + 10) > 6.5 and x / (y / 100 + 10) < 8.5:
            actions.append("电机控制：调整电压转速比正常")
        if (y / 100) / (z + 20) > 0.35 and (y / 100) / (z + 20) < 0.55:
            actions.append("电机控制：调整转速功率比正常")
        if z / (x / 50) > 10 and z / (x / 50) < 14:
            actions.append("电机控制：调整功率电压比正常")
        if (x * y * z / 100000) > 35 and (x * y * z / 100000) < 45:
            actions.append("电机控制：三元电机积正常")
        if x / 10 + y / 100 * 0.8 + z * 1.2 > 110 and x / 10 + y / 100 * 0.8 + z * 1.2 < 130:
            actions.append("电机控制：加权电机和正常")
        if (x / 230) ** 0.4 * (y / 3000) ** 0.4 * (z / 55) ** 0.2 > 0.95:
            actions.append("电机控制：加权几何平均正常")
        if (x - 230) * (y - 3000) / 100 > -300 and (x - 230) * (y - 3000) / 100 < 300:
            actions.append("电机控制：电压转速偏差积平衡")
        if (x - 230) * (z - 55) > -60 and (x - 230) * (z - 55) < 60:
            actions.append("电机控制：电压功率偏差积平衡")
        if (y - 3000) / 50 * (z - 55) > -30 and (y - 3000) / 50 * (z - 55) < 30:
            actions.append("电机控制：转速功率偏差积平衡")
        if abs(x / 230 + (y / 100) / 30 + z / 55 - 3) < 0.2:
            actions.append("电机控制：归一化和接近理想")
        if abs((x / 230) * (y / 3000) * (z / 55) - 1) < 0.15:
            actions.append("电机控制：归一化积接近理想")

        # 效率优化分析 (41-60)
        if x > 228 and x < 232 and y > 2980 and y < 3020 and z > 53 and z < 57:
            actions.append("电机控制：全参数高效运行")
        if abs(x / (y / 100) - 8.0) < 0.3 and abs(z - 55) < 2:
            actions.append("电机控制：电压转速比与功率协调")
        if x * y / 1000 > 690 and x * y / 1000 < 710 and z > 54 and z < 56:
            actions.append("电机控制：电压转速积与功率协调")
        if (x + y / 100) / 2 > 52 and (x + y / 100) / 2 < 56 and z > 53 and z < 57:
            actions.append("电机控制：电压转速平均与功率协调")
        if x - y / 100 > 195 and x - y / 100 < 205 and z > 53 and z < 57:
            actions.append("电机控制：电压转速差与功率协调")
        if sqrt((x - 230) ** 2 + (y / 100 - 30) ** 2) < 3 and abs(z - 55) < 2:
            actions.append("电机控制：电压转速距离与功率优秀")
        if x / 230 > 0.985 and x / 230 < 1.015 and (y / 100) / 30 > 0.98 and (y / 100) / 30 < 1.02:
            actions.append("电机控制：电压转速相对值协调")
        if z / 55 > 0.982 and z / 55 < 1.018:
            actions.append("电机控制：功率相对值协调")
        if (x / 230 + (y / 100) / 30 + z / 55) / 3 > 0.985 and (x / 230 + (y / 100) / 30 + z / 55) / 3 < 1.015:
            actions.append("电机控制：归一化平均协调")
        if max(abs(x / 230 - 1), abs((y / 100) / 30 - 1), abs(z / 55 - 1)) < 0.02:
            actions.append("电机控制：最大偏差协调良好")
        if min(x / 230, (y / 100) / 30, z / 55) > 0.985:
            actions.append("电机控制：最小相对值协调良好")
        if max(x / 230, (y / 100) / 30, z / 55) < 1.015:
            actions.append("电机控制：最大相对值协调良好")
        if abs(max(x, y / 100 * 7.7, z * 4.2) - min(x, y / 100 * 7.7, z * 4.2)) < 8:
            actions.append("电机控制：标准化后范围协调")
        if (x / 10 + y / 100 + z) > 118 and (x / 10 + y / 100 + z) < 122:
            actions.append("电机控制：综合参数协调优秀")
        if (x * y * z / 100000) > 39 and (x * y * z / 100000) < 41:
            actions.append("电机控制：三元积协调优秀")
        if abs((x / 10 + y / 100 + z) / 3 - 40) < 1:
            actions.append("电机控制：平均值协调优秀")
        if x > 229 and x < 231 and y > 2995 and y < 3005:
            actions.append("电机控制：电压转速超精密协调")
        if z > 54.5 and z < 55.5:
            actions.append("电机控制：功率超精密协调")
        if ((x - 230) ** 2 + (y / 100 - 30) ** 2 + (z - 55) ** 2) < 9:
            actions.append("电机控制：三维距离协调完美")
        if abs(x / (y / 100) / z - 1.45) < 0.05:
            actions.append("电机控制：连续比例协调完美")

        # 预测维护建议 (61-80)
        if x < 225 or y < 2900 or z < 50:
            actions.append("电机控制：检测到参数偏低趋势")
            self.adjustments.append("预防性提升电机性能")
        if x > 235 or y > 3100 or z > 60:
            actions.append("电机控制：检测到参数偏高趋势")
            self.adjustments.append("预防性限制电机负荷")
        if abs(x / (y / 100) - 8.0) > 0.4:
            actions.append("电机控制：电压转速比需要调整")
            self.adjustments.append("优化电压转速协调控制")
        if abs(x / z - 4.2) > 0.2:
            actions.append("电机控制：电压功率比需要调整")
            self.adjustments.append("优化电压功率协调控制")
        if abs((y / 100) / z - 0.55) > 0.05:
            actions.append("电机控制：转速功率比需要调整")
            self.adjustments.append("优化转速功率协调控制")
        if (x / 10 + y / 100 + z) < 115 or (x / 10 + y / 100 + z) > 125:
            actions.append("电机控制：综合指标需要调整")
            self.adjustments.append("全面优化电机参数")
        if (x * y * z / 100000) < 37 or (x * y * z / 100000) > 43:
            actions.append("电机控制：三元积需要调整")
            self.adjustments.append("平衡三参数关系")
        if x < 223 and y < 2900:
            actions.append("电机控制：电压转速同时偏低")
            self.alert_status = "medium"
        if x > 237 and y > 3100:
            actions.append("电机控制：电压转速同时偏高")
            self.alert_status = "medium"
        if x < 223 and z < 50:
            actions.append("电机控制：电压功率同时偏低")
            self.alert_status = "medium"
        if x > 237 and z > 60:
            actions.append("电机控制：电压功率同时偏高")
            self.alert_status = "medium"
        if y < 2900 and z < 50:
            actions.append("电机控制：转速功率同时偏低")
            self.alert_status = "medium"
        if y > 3100 and z > 60:
            actions.append("电机控制：转速功率同时偏高")
            self.alert_status = "medium"
        if x < 218 or x > 242:
            actions.append("电机控制：电压在临界范围")
            self.alert_status = "high"
        if y < 2750 or y > 3250:
            actions.append("电机控制：转速在临界范围")
            self.alert_status = "high"
        if z < 42 or z > 68:
            actions.append("电机控制：功率在临界范围")
            self.alert_status = "high"
        if x < 215 and y < 2800 and z < 45:
            actions.append("电机控制：所有参数严重偏低")
            self.alert_status = "emergency"
        if x > 245 and y > 3200 and z > 65:
            actions.append("电机控制：所有参数严重偏高")
            self.alert_status = "emergency"
        if ((x - 230) ** 2 + (y / 100 - 30) ** 2 + (z - 55) ** 2) > 50:
            actions.append("电机控制：三维偏离过大")
            self.alert_status = "high"
        if abs((x * y * z / 100000) / 40 - 1) > 0.2:
            actions.append("电机控制：三元积偏离过大")
            self.alert_status = "medium"

        return actions

    def section7_flow_pressure_density_hybrid(self, x, y, z):
        """第7类: 流体控制系统 (流量x, 压力y, 密度z) - 80个if语句"""
        actions = []

        # 流体系统基础检查 (1-20)
        if self.standards['flow_rate'][0] < x < self.standards['flow_rate'][1]:
            actions.append("流体控制：流量基础参数正常")
        if self.standards['pressure'][0] < y < self.standards['pressure'][1]:
            actions.append("流体控制：压力基础参数正常")
        if self.standards['density'][0] < z < self.standards['density'][1]:
            actions.append("流体控制：密度基础参数正常")
        if x > 60 and x < 70 and y > 90 and y < 110:
            actions.append("流体控制：流量压力核心区间协调")
        if x > 55 and x < 75 and z > 1.25 and z < 1.45:
            actions.append("流体控制：流量密度匹配良好")
        if y > 85 and y < 115 and z > 1.25 and z < 1.45:
            actions.append("流体控制：压力密度协调稳定")
        if x > 63 and x < 67:
            actions.append("流体控制：流量精确控制")
        if y > 98 and y < 102:
            actions.append("流体控制：压力精确控制")
        if z > 1.33 and z < 1.37:
            actions.append("流体控制：密度精确控制")
        if abs(x - 65) < 2.5:
            actions.append("流体控制：流量稳定性良好")
        if abs(y - 100) < 5:
            actions.append("流体控制：压力稳定性良好")
        if abs(z - 1.35) < 0.05:
            actions.append("流体控制：密度稳定性良好")
        if x > 55 and y > 85 and z > 1.25:
            actions.append("流体控制：所有参数在安全范围")
        if x < 75 and y < 115 and z < 1.45:
            actions.append("流体控制：所有参数未超上限")
        if x + y + z * 100 > 280 and x + y + z * 100 < 320:
            actions.append("流体控制：综合流体指标正常")
        if x * y > 5500 and x * y < 7500:
            actions.append("流体控制：流量压力乘积正常")
        if x * z > 70 and x * z < 110:
            actions.append("流体控制：流量密度乘积正常")
        if y * z > 110 and y * z < 150:
            actions.append("流体控制：压力密度乘积正常")
        if (x + y + z * 100) / 3 > 93 and (x + y + z * 100) / 3 < 107:
            actions.append("流体控制：平均流体参数正常")
        if sqrt(x ** 2 + y ** 2 + (z * 100) ** 2) > 150:
            actions.append("流体控制：流体向量模长正常")

        # 流体动力学分析 (21-40)
        if x / y > 0.55 and x / y < 0.85:
            actions.append("流体控制：流量压力比理想")
        if x / z > 40 and x / z < 60:
            actions.append("流体控制：流量密度比正常")
        if y / z > 60 and y / z < 90:
            actions.append("流体控制：压力密度比适当")
        if (x - 50) + (y - 80) > 20 and (x - 50) + (y - 80) < 40:
            actions.append("流体控制：流量压力偏差和正常")
        if (x - 50) + (z - 1.2) * 100 > 25 and (x - 50) + (z - 1.2) * 100 < 45:
            actions.append("流体控制：流量密度偏差和正常")
        if (y - 80) + (z - 1.2) * 100 > 35 and (y - 80) + (z - 1.2) * 100 < 55:
            actions.append("流体控制：压力密度偏差和正常")
        if abs((x - 65) - (y - 100) * 0.65) < 8:
            actions.append("流体控制：流量压力偏差关系平衡")
        if abs((x - 65) - (z - 1.35) * 50) < 10:
            actions.append("流体控制：流量密度偏差关系平衡")
        if abs((y - 100) - (z - 1.35) * 100) < 12:
            actions.append("流体控制：压力密度偏差关系平衡")
        if x / (y + 20) > 0.5 and x / (y + 20) < 0.8:
            actions.append("流体控制：调整流量压力比正常")
        if y / (z + 0.5) > 45 and y / (z + 0.5) < 75:
            actions.append("流体控制：调整压力密度比正常")
        if z / (x / 50) > 0.9 and z / (x / 50) < 1.4:
            actions.append("流体控制：调整密度流量比正常")
        if (x * y * z) > 7000 and (x * y * z) < 11000:
            actions.append("流体控制：三元流体积正常")
        if x * 0.6 + y * 0.3 + z * 10 > 65 and x * 0.6 + y * 0.3 + z * 10 < 75:
            actions.append("流体控制：加权流体和正常")
        if (x / 65) ** 0.5 * (y / 100) ** 0.3 * (z / 1.35) ** 0.2 > 0.92:
            actions.append("流体控制：加权几何平均正常")
        if (x - 65) * (y - 100) > -150 and (x - 65) * (y - 100) < 150:
            actions.append("流体控制：流量压力偏差积平衡")
        if (x - 65) * (z - 1.35) > -8 and (x - 65) * (z - 1.35) < 8:
            actions.append("流体控制：流量密度偏差积平衡")
        if (y - 100) * (z - 1.35) > -10 and (y - 100) * (z - 1.35) < 10:
            actions.append("流体控制：压力密度偏差积平衡")
        if abs(x / 65 + y / 100 + z / 1.35 - 3) < 0.25:
            actions.append("流体控制：归一化和接近理想")
        if abs((x / 65) * (y / 100) * (z / 1.35) - 1) < 0.18:
            actions.append("流体控制：归一化积接近理想")

        # 流体平衡控制 (41-60)
        if x > 62 and x < 68 and y > 97 and y < 103 and z > 1.32 and z < 1.38:
            actions.append("流体控制：全参数平衡运行")
        if abs(x / y - 0.65) < 0.08 and abs(z - 1.35) < 0.03:
            actions.append("流体控制：流量压力比与密度协调")
        if x * y > 6300 and x * y < 6700 and z > 1.33 and z < 1.37:
            actions.append("流体控制：流量压力积与密度协调")
        if (x + y) / 2 > 80 and (x + y) / 2 < 86 and z > 1.32 and z < 1.38:
            actions.append("流体控制：流量压力平均与密度协调")
        if abs(x - y) < 40 and z > 1.32 and z < 1.38:
            actions.append("流体控制：流量压力差与密度协调")
        if sqrt((x - 65) ** 2 + (y - 100) ** 2) < 5 and abs(z - 1.35) < 0.03:
            actions.append("流体控制：流量压力距离与密度优秀")
        if x / 65 > 0.975 and x / 65 < 1.025 and y / 100 > 0.97 and y / 100 < 1.03:
            actions.append("流体控制：流量压力相对值协调")
        if z / 1.35 > 0.978 and z / 1.35 < 1.022:
            actions.append("流体控制：密度相对值协调")
        if (x / 65 + y / 100 + z / 1.35) / 3 > 0.975 and (x / 65 + y / 100 + z / 1.35) / 3 < 1.025:
            actions.append("流体控制：归一化平均协调")
        if max(abs(x / 65 - 1), abs(y / 100 - 1), abs(z / 1.35 - 1)) < 0.025:
            actions.append("流体控制：最大偏差协调良好")
        if min(x / 65, y / 100, z / 1.35) > 0.975:
            actions.append("流体控制：最小相对值协调良好")
        if max(x / 65, y / 100, z / 1.35) < 1.025:
            actions.append("流体控制：最大相对值协调良好")
        if abs(max(x, y, z * 100) - min(x, y, z * 100)) < 40:
            actions.append("流体控制：标准化后范围协调")
        if (x + y + z * 100) > 298 and (x + y + z * 100) < 302:
            actions.append("流体控制：综合参数协调优秀")
        if (x * y * z) > 8700 and (x * y * z) < 9300:
            actions.append("流体控制：三元积协调优秀")
        if abs((x + y + z * 100) / 3 - 100) < 2:
            actions.append("流体控制：平均值协调优秀")
        if x > 64 and x < 66 and y > 99 and y < 101:
            actions.append("流体控制：流量压力超精密协调")
        if z > 1.345 and z < 1.355:
            actions.append("流体控制：密度超精密协调")
        if ((x - 65) ** 2 + (y - 100) ** 2 + (z - 1.35) ** 2 * 10000) < 25:
            actions.append("流体控制：三维距离协调完美")
        if abs(x / y / z - 48) < 2:
            actions.append("流体控制：连续比例协调完美")

        # 流体安全控制 (61-80)
        if x < 52 or y < 85 or z < 1.22:
            actions.append("流体控制：检测到参数偏低趋势")
            self.adjustments.append("预防性提升流体参数")
        if x > 78 or y > 115 or z > 1.48:
            actions.append("流体控制：检测到参数偏高趋势")
            self.adjustments.append("预防性限制流体负荷")
        if abs(x / y - 0.65) > 0.12:
            actions.append("流体控制：流量压力比需要调整")
            self.adjustments.append("优化流量压力协调控制")
        if abs(x / z - 48) > 6:
            actions.append("流体控制：流量密度比需要调整")
            self.adjustments.append("优化流量密度协调控制")
        if abs(y / z - 74) > 8:
            actions.append("流体控制：压力密度比需要调整")
            self.adjustments.append("优化压力密度协调控制")
        if (x + y + z * 100) < 285 or (x + y + z * 100) > 315:
            actions.append("流体控制：综合指标需要调整")
            self.adjustments.append("全面优化流体参数")
        if (x * y * z) < 8000 or (x * y * z) > 10000:
            actions.append("流体控制：三元积需要调整")
            self.adjustments.append("平衡三参数关系")
        if x < 55 and y < 90:
            actions.append("流体控制：流量压力同时偏低")
            self.alert_status = "medium"
        if x > 75 and y > 110:
            actions.append("流体控制：流量压力同时偏高")
            self.alert_status = "medium"
        if x < 55 and z < 1.25:
            actions.append("流体控制：流量密度同时偏低")
            self.alert_status = "medium"
        if x > 75 and z > 1.45:
            actions.append("流体控制：流量密度同时偏高")
            self.alert_status = "medium"
        if y < 90 and z < 1.25:
            actions.append("流体控制：压力密度同时偏低")
            self.alert_status = "medium"
        if y > 110 and z > 1.45:
            actions.append("流体控制：压力密度同时偏高")
            self.alert_status = "medium"
        if x < 48 or x > 82:
            actions.append("流体控制：流量在临界范围")
            self.alert_status = "high"
        if y < 75 or y > 125:
            actions.append("流体控制：压力在临界范围")
            self.alert_status = "high"
        if z < 1.15 or z > 1.55:
            actions.append("流体控制：密度在临界范围")
            self.alert_status = "high"
        if x < 45 and y < 80 and z < 1.2:
            actions.append("流体控制：所有参数严重偏低")
            self.alert_status = "emergency"
        if x > 85 and y > 120 and z > 1.5:
            actions.append("流体控制：所有参数严重偏高")
            self.alert_status = "emergency"
        if ((x - 65) ** 2 + (y - 100) ** 2 + (z - 1.35) ** 2 * 10000) > 200:
            actions.append("流体控制：三维偏离过大")
            self.alert_status = "high"
        if abs((x * y * z) / 9000 - 1) > 0.22:
            actions.append("流体控制：三元积偏离过大")
            self.alert_status = "medium"

        return actions

    def section8_hybrid_speed_torque_current(self, x, y, z):
        """第8类: 速度扭矩电流混合控制 (速度x, 扭矩y, 电流z) - 80个if语句"""
        actions = []

        # 速度扭矩电流基础检查 (1-20)
        if self.standards['speed'][0] < x < self.standards['speed'][1]:
            actions.append("混合控制：速度基础参数正常")
        if self.standards['torque'][0] < y < self.standards['torque'][1]:
            actions.append("混合控制：扭矩基础参数正常")
        if self.standards['current'][0] < z < self.standards['current'][1]:
            actions.append("混合控制：电流基础参数正常")
        if x > 1400 and x < 1600 and y > 160 and y < 190:
            actions.append("混合控制：速度扭矩核心区间协调")
        if x > 1350 and x < 1650 and z > 11 and z < 14:
            actions.append("混合控制：速度电流匹配良好")
        if y > 155 and y < 195 and z > 11 and z < 14:
            actions.append("混合控制：扭矩电流协调稳定")
        if x > 1480 and x < 1520:
            actions.append("混合控制：速度精确控制")
        if y > 173 and y < 177:
            actions.append("混合控制：扭矩精确控制")
        if z > 12.2 and z < 12.8:
            actions.append("混合控制：电流精确控制")
        if abs(x - 1500) < 40:
            actions.append("混合控制：速度稳定性良好")
        if abs(y - 175) < 6:
            actions.append("混合控制：扭矩稳定性良好")
        if abs(z - 12.5) < 0.4:
            actions.append("混合控制：电流稳定性良好")
        if x > 1350 and y > 155 and z > 11:
            actions.append("混合控制：所有参数在安全范围")
        if x < 1650 and y < 195 and z < 14:
            actions.append("混合控制：所有参数未超上限")
        if x / 10 + y + z * 10 > 350 and x / 10 + y + z * 10 < 410:
            actions.append("混合控制：综合驱动指标正常")
        if x * y / 1000 > 230 and x * y / 1000 < 310:
            actions.append("混合控制：速度扭矩乘积正常")
        if x * z / 100 > 160 and x * z / 100 < 220:
            actions.append("混合控制：速度电流乘积正常")
        if y * z > 2000 and y * z < 2600:
            actions.append("混合控制：扭矩电流乘积正常")
        if (x / 10 + y + z * 10) / 3 > 116 and (x / 10 + y + z * 10) / 3 < 136:
            actions.append("混合控制：平均驱动参数正常")
        if sqrt((x / 10) ** 2 + y ** 2 + (z * 10) ** 2) > 220:
            actions.append("混合控制：驱动向量模长正常")

        # 驱动系统分析 (21-40)
        if x / (y * 10) > 0.8 and x / (y * 10) < 1.0:
            actions.append("混合控制：速度扭矩比理想")
        if x / z > 110 and x / z < 140:
            actions.append("混合控制：速度电流比正常")
        if y / z > 12 and y / z < 16:
            actions.append("混合控制：扭矩电流比适当")
        if (x - 1200) / 10 + (y - 150) > 35 and (x - 1200) / 10 + (y - 150) < 55:
            actions.append("混合控制：速度扭矩偏差和正常")
        if (x - 1200) / 10 + (z - 10) * 5 > 40 and (x - 1200) / 10 + (z - 10) * 5 < 60:
            actions.append("混合控制：速度电流偏差和正常")
        if (y - 150) + (z - 10) * 2 > 30 and (y - 150) + (z - 10) * 2 < 50:
            actions.append("混合控制：扭矩电流偏差和正常")
        if abs((x - 1500) / 10 - (y - 175)) < 12:
            actions.append("混合控制：速度扭矩偏差关系平衡")
        if abs((x - 1500) / 100 - (z - 12.5)) < 3:
            actions.append("混合控制：速度电流偏差关系平衡")
        if abs((y - 175) - (z - 12.5) * 12) < 8:
            actions.append("混合控制：扭矩电流偏差关系平衡")
        if x / (y * 10 + 500) > 0.7 and x / (y * 10 + 500) < 0.9:
            actions.append("混合控制：调整速度扭矩比正常")
        if y / (z + 5) > 9 and y / (z + 5) < 13:
            actions.append("混合控制：调整扭矩电流比正常")
        if z / (x / 150) > 1.1 and z / (x / 150) < 1.5:
            actions.append("混合控制：调整电流速度比正常")
        if (x * y * z / 10000) > 28 and (x * y * z / 10000) < 38:
            actions.append("混合控制：三元驱动积正常")
        if x / 20 + y * 0.6 + z * 8 > 260 and x / 20 + y * 0.6 + z * 8 < 300:
            actions.append("混合控制：加权驱动和正常")
        if (x / 1500) ** 0.4 * (y / 175) ** 0.4 * (z / 12.5) ** 0.2 > 0.93:
            actions.append("混合控制：加权几何平均正常")
        if (x - 1500) / 10 * (y - 175) > -200 and (x - 1500) / 10 * (y - 175) < 200:
            actions.append("混合控制：速度扭矩偏差积平衡")
        if (x - 1500) / 100 * (z - 12.5) > -20 and (x - 1500) / 100 * (z - 12.5) < 20:
            actions.append("混合控制：速度电流偏差积平衡")
        if (y - 175) * (z - 12.5) > -20 and (y - 175) * (z - 12.5) < 20:
            actions.append("混合控制：扭矩电流偏差积平衡")
        if abs(x / 1500 + y / 175 + z / 12.5 - 3) < 0.2:
            actions.append("混合控制：归一化和接近理想")
        if abs((x / 1500) * (y / 175) * (z / 12.5) - 1) < 0.15:
            actions.append("混合控制：归一化积接近理想")

        # 动力传动协调 (41-60)
        if x > 1480 and x < 1520 and y > 172 and y < 178 and z > 12.2 and z < 12.8:
            actions.append("混合控制：全参数动力协调")
        if abs(x / (y * 10) - 0.86) < 0.06 and abs(z - 12.5) < 0.3:
            actions.append("混合控制：速度扭矩比与电流协调")
        if x * y / 1000 > 258 and x * y / 1000 < 268 and z > 12.3 and z < 12.7:
            actions.append("混合控制：速度扭矩积与电流协调")
        if (x / 10 + y) / 2 > 137 and (x / 10 + y) / 2 < 143 and z > 12.2 and z < 12.8:
            actions.append("混合控制：速度扭矩平均与电流协调")
        if abs(x / 10 - y) < 25 and z > 12.2 and z < 12.8:
            actions.append("混合控制：速度扭矩差与电流协调")
        if sqrt((x / 10 - 150) ** 2 + (y - 175) ** 2) < 6 and abs(z - 12.5) < 0.3:
            actions.append("混合控制：速度扭矩距离与电流优秀")
        if x / 1500 > 0.98 and x / 1500 < 1.02 and y / 175 > 0.97 and y / 175 < 1.03:
            actions.append("混合控制：速度扭矩相对值协调")
        if z / 12.5 > 0.976 and z / 12.5 < 1.024:
            actions.append("混合控制：电流相对值协调")
        if (x / 1500 + y / 175 + z / 12.5) / 3 > 0.98 and (x / 1500 + y / 175 + z / 12.5) / 3 < 1.02:
            actions.append("混合控制：归一化平均协调")
        if max(abs(x / 1500 - 1), abs(y / 175 - 1), abs(z / 12.5 - 1)) < 0.025:
            actions.append("混合控制：最大偏差协调良好")
        if min(x / 1500, y / 175, z / 12.5) > 0.98:
            actions.append("混合控制：最小相对值协调良好")
        if max(x / 1500, y / 175, z / 12.5) < 1.02:
            actions.append("混合控制：最大相对值协调良好")
        if abs(max(x / 10, y, z * 10) - min(x / 10, y, z * 10)) < 30:
            actions.append("混合控制：标准化后范围协调")
        if (x / 10 + y + z * 10) > 378 and (x / 10 + y + z * 10) < 382:
            actions.append("混合控制：综合参数协调优秀")
        if (x * y * z / 10000) > 32.5 and (x * y * z / 10000) < 33.5:
            actions.append("混合控制：三元积协调优秀")
        if abs((x / 10 + y + z * 10) / 3 - 126.7) < 1.5:
            actions.append("混合控制：平均值协调优秀")
        if x > 1495 and x < 1505 and y > 174 and y < 176:
            actions.append("混合控制：速度扭矩超精密协调")
        if z > 12.45 and z < 12.55:
            actions.append("混合控制：电流超精密协调")
        if ((x / 10 - 150) ** 2 + (y - 175) ** 2 + (z - 12.5) ** 2 * 100) < 25:
            actions.append("混合控制：三维距离协调完美")
        if abs(x / y / z - 9.6) < 0.5:
            actions.append("混合控制：连续比例协调完美")

        # 动力控制优化 (61-80)
        if x < 1350 or y < 160 or z < 11.5:
            actions.append("混合控制：检测到参数偏低趋势")
            self.adjustments.append("预防性提升动力参数")
        if x > 1650 or y > 190 or z > 13.5:
            actions.append("混合控制：检测到参数偏高趋势")
            self.adjustments.append("预防性限制动力负荷")
        if abs(x / (y * 10) - 0.86) > 0.1:
            actions.append("混合控制：速度扭矩比需要调整")
            self.adjustments.append("优化速度扭矩协调控制")
        if abs(x / z - 120) > 15:
            actions.append("混合控制：速度电流比需要调整")
            self.adjustments.append("优化速度电流协调控制")
        if abs(y / z - 14) > 2:
            actions.append("混合控制：扭矩电流比需要调整")
            self.adjustments.append("优化扭矩电流协调控制")
        if (x / 10 + y + z * 10) < 360 or (x / 10 + y + z * 10) > 400:
            actions.append("混合控制：综合指标需要调整")
            self.adjustments.append("全面优化动力参数")
        if (x * y * z / 10000) < 30 or (x * y * z / 10000) > 36:
            actions.append("混合控制：三元积需要调整")
            self.adjustments.append("平衡三参数关系")
        if x < 1300 and y < 160:
            actions.append("混合控制：速度扭矩同时偏低")
            self.alert_status = "medium"
        if x > 1700 and y > 190:
            actions.append("混合控制：速度扭矩同时偏高")
            self.alert_status = "medium"
        if x < 1300 and z < 11:
            actions.append("混合控制：速度电流同时偏低")
            self.alert_status = "medium"
        if x > 1700 and z > 14:
            actions.append("混合控制：速度电流同时偏高")
            self.alert_status = "medium"
        if y < 160 and z < 11:
            actions.append("混合控制：扭矩电流同时偏低")
            self.alert_status = "medium"
        if y > 190 and z > 14:
            actions.append("混合控制：扭矩电流同时偏高")
            self.alert_status = "medium"
        if x < 1250 or x > 1750:
            actions.append("混合控制：速度在临界范围")
            self.alert_status = "high"
        if y < 145 or y > 205:
            actions.append("混合控制：扭矩在临界范围")
            self.alert_status = "high"
        if z < 9.5 or z > 15.5:
            actions.append("混合控制：电流在临界范围")
            self.alert_status = "high"
        if x < 1200 and y < 155 and z < 10.5:
            actions.append("混合控制：所有参数严重偏低")
            self.alert_status = "emergency"
        if x > 1800 and y > 195 and z > 14.5:
            actions.append("混合控制：所有参数严重偏高")
            self.alert_status = "emergency"
        if ((x / 10 - 150) ** 2 + (y - 175) ** 2 + (z - 12.5) ** 2 * 100) > 150:
            actions.append("混合控制：三维偏离过大")
            self.alert_status = "high"
        if abs((x * y * z / 10000) / 33 - 1) > 0.18:
            actions.append("混合控制：三元积偏离过大")
            self.alert_status = "medium"

        return actions

    def section9_thermal_electrical_flow_hybrid(self, x, y, z):
        """第9类: 热电流体混合控制 (温度x, 电压y, 流量z) - 80个if语句"""
        actions = []

        # 热电流体基础检查 (1-20)
        if self.standards['temperature'][0] < x < self.standards['temperature'][1]:
            actions.append("热电流体：温度基础参数正常")
        if self.standards['voltage'][0] < y < self.standards['voltage'][1]:
            actions.append("热电流体：电压基础参数正常")
        if self.standards['flow_rate'][0] < z < self.standards['flow_rate'][1]:
            actions.append("热电流体：流量基础参数正常")
        if x > 190 and x < 210 and y > 225 and y < 235:
            actions.append("热电流体：温度电压核心区间协调")
        if x > 185 and x < 215 and z > 60 and z < 70:
            actions.append("热电流体：温度流量匹配良好")
        if y > 223 and y < 237 and z > 60 and z < 70:
            actions.append("热电流体：电压流量协调稳定")
        if x > 198 and x < 202:
            actions.append("热电流体：温度精确控制")
        if y > 228 and y < 232:
            actions.append("热电流体：电压精确控制")
        if z > 63 and z < 67:
            actions.append("热电流体：流量精确控制")
        if abs(x - 200) < 5:
            actions.append("热电流体：温度稳定性良好")
        if abs(y - 230) < 3:
            actions.append("热电流体：电压稳定性良好")
        if abs(z - 65) < 2:
            actions.append("热电流体：流量稳定性良好")
        if x > 185 and y > 223 and z > 55:
            actions.append("热电流体：所有参数在安全范围")
        if x < 215 and y < 237 and z < 75:
            actions.append("热电流体：所有参数未超上限")
        if x + y / 10 + z > 288 and x + y / 10 + z < 308:
            actions.append("热电流体：综合系统指标正常")
        if x * y / 100 > 440 and x * y / 100 < 500:
            actions.append("热电流体：温度电压乘积正常")
        if x * z > 12500 and x * z < 14500:
            actions.append("热电流体：温度流量乘积正常")
        if y / 10 * z > 1400 and y / 10 * z < 1600:
            actions.append("热电流体：电压流量乘积正常")
        if (x + y / 10 + z) / 3 > 96 and (x + y / 10 + z) / 3 < 102:
            actions.append("热电流体：平均系统参数正常")
        if sqrt(x ** 2 + (y / 10) ** 2 + z ** 2) > 220:
            actions.append("热电流体：系统向量模长正常")

        # 热电耦合分析 (21-40)
        if x / (y / 10) > 8.2 and x / (y / 10) < 9.2:
            actions.append("热电流体：温度电压比理想")
        if x / z > 2.8 and x / z < 3.4:
            actions.append("热电流体：温度流量比正常")
        if (y / 10) / z > 0.32 and (y / 10) / z < 0.38:
            actions.append("热电流体：电压流量比适当")
        if (x - 180) + (y - 220) / 10 > 22 and (x - 180) + (y - 220) / 10 < 32:
            actions.append("热电流体：温度电压偏差和正常")
        if (x - 180) + (z - 50) > 35 and (x - 180) + (z - 50) < 45:
            actions.append("热电流体：温度流量偏差和正常")
        if (y - 220) / 10 + (z - 50) > 25 and (y - 220) / 10 + (z - 50) < 35:
            actions.append("热电流体：电压流量偏差和正常")
        if abs((x - 200) - (y - 230) / 2) < 8:
            actions.append("热电流体：温度电压偏差关系平衡")
        if abs((x - 200) - (z - 65) * 3) < 12:
            actions.append("热电流体：温度流量偏差关系平衡")
        if abs((y - 230) / 10 - (z - 65) / 5) < 3:
            actions.append("热电流体：电压流量偏差关系平衡")
        if x / (y / 10 + 5) > 6.5 and x / (y / 10 + 5) < 8.5:
            actions.append("热电流体：调整温度电压比正常")
        if (y / 10) / (z + 15) > 0.26 and (y / 10) / (z + 15) < 0.34:
            actions.append("热电流体：调整电压流量比正常")
        if z / (x / 5) > 1.4 and z / (x / 5) < 1.8:
            actions.append("热电流体：调整流量温度比正常")
        if (x * y * z / 10000) > 28 and (x * y * z / 10000) < 36:
            actions.append("热电流体：三元系统积正常")
        if x * 0.5 + y / 10 * 0.3 + z * 0.2 > 125 and x * 0.5 + y / 10 * 0.3 + z * 0.2 < 135:
            actions.append("热电流体：加权系统和正常")
        if (x / 200) ** 0.4 * (y / 230) ** 0.3 * (z / 65) ** 0.3 > 0.92:
            actions.append("热电流体：加权几何平均正常")
        if (x - 200) * (y - 230) / 10 > -80 and (x - 200) * (y - 230) / 10 < 80:
            actions.append("热电流体：温度电压偏差积平衡")
        if (x - 200) * (z - 65) > -120 and (x - 200) * (z - 65) < 120:
            actions.append("热电流体：温度流量偏差积平衡")
        if (y - 230) / 10 * (z - 65) > -30 and (y - 230) / 10 * (z - 65) < 30:
            actions.append("热电流体：电压流量偏差积平衡")
        if abs(x / 200 + y / 230 + z / 65 - 3) < 0.18:
            actions.append("热电流体：归一化和接近理想")
        if abs((x / 200) * (y / 230) * (z / 65) - 1) < 0.12:
            actions.append("热电流体：归一化积接近理想")

        # 系统综合协调 (41-60)
        if x > 197 and x < 203 and y > 228 and y < 232 and z > 63.5 and z < 66.5:
            actions.append("热电流体：全参数系统协调")
        if abs(x / (y / 10) - 8.7) < 0.3 and abs(z - 65) < 1.5:
            actions.append("热电流体：温度电压比与流量协调")
        if x * y / 100 > 455 and x * y / 100 < 465 and z > 64 and z < 66:
            actions.append("热电流体：温度电压积与流量协调")
        if (x + y / 10) / 2 > 112 and (x + y / 10) / 2 < 116 and z > 63.5 and z < 66.5:
            actions.append("热电流体：温度电压平均与流量协调")
        if abs(x - y / 10) < 175 and z > 63.5 and z < 66.5:
            actions.append("热电流体：温度电压差与流量协调")
        if sqrt((x - 200) ** 2 + (y / 10 - 23) ** 2) < 3 and abs(z - 65) < 1.5:
            actions.append("热电流体：温度电压距离与流量优秀")
        if x / 200 > 0.985 and x / 200 < 1.015 and y / 230 > 0.987 and y / 230 < 1.013:
            actions.append("热电流体：温度电压相对值协调")
        if z / 65 > 0.985 and z / 65 < 1.015:
            actions.append("热电流体：流量相对值协调")
        if (x / 200 + y / 230 + z / 65) / 3 > 0.985 and (x / 200 + y / 230 + z / 65) / 3 < 1.015:
            actions.append("热电流体：归一化平均协调")
        if max(abs(x / 200 - 1), abs(y / 230 - 1), abs(z / 65 - 1)) < 0.02:
            actions.append("热电流体：最大偏差协调良好")
        if min(x / 200, y / 230, z / 65) > 0.985:
            actions.append("热电流体：最小相对值协调良好")
        if max(x / 200, y / 230, z / 65) < 1.015:
            actions.append("热电流体：最大相对值协调良好")
        if abs(max(x, y / 10 * 8.7, z) - min(x, y / 10 * 8.7, z)) < 15:
            actions.append("热电流体：标准化后范围协调")
        if (x + y / 10 + z) > 297 and (x + y / 10 + z) < 299:
            actions.append("热电流体：综合参数协调优秀")
        if (x * y * z / 10000) > 31.5 and (x * y * z / 10000) < 32.5:
            actions.append("热电流体：三元积协调优秀")
        if abs((x + y / 10 + z) / 3 - 99.3) < 0.8:
            actions.append("热电流体：平均值协调优秀")
        if x > 199 and x < 201 and y > 229 and y < 231:
            actions.append("热电流体：温度电压超精密协调")
        if z > 64.5 and z < 65.5:
            actions.append("热电流体：流量超精密协调")
        if ((x - 200) ** 2 + (y / 10 - 23) ** 2 + (z - 65) ** 2) < 9:
            actions.append("热电流体：三维距离协调完美")
        if abs(x / (y / 10) / z - 0.134) < 0.005:
            actions.append("热电流体：连续比例协调完美")

        # 综合控制策略 (61-80)
        if x < 188 or y < 225 or z < 58:
            actions.append("热电流体：检测到参数偏低趋势")
            self.adjustments.append("预防性提升系统参数")
        if x > 212 or y > 235 or z > 72:
            actions.append("热电流体：检测到参数偏高趋势")
            self.adjustments.append("预防性限制系统负荷")
        if abs(x / (y / 10) - 8.7) > 0.5:
            actions.append("热电流体：温度电压比需要调整")
            self.adjustments.append("优化温度电压协调控制")
        if abs(x / z - 3.08) > 0.3:
            actions.append("热电流体：温度流量比需要调整")
            self.adjustments.append("优化温度流量协调控制")
        if abs((y / 10) / z - 0.354) > 0.03:
            actions.append("热电流体：电压流量比需要调整")
            self.adjustments.append("优化电压流量协调控制")
        if (x + y / 10 + z) < 292 or (x + y / 10 + z) > 304:
            actions.append("热电流体：综合指标需要调整")
            self.adjustments.append("全面优化系统参数")
        if (x * y * z / 10000) < 29 or (x * y * z / 10000) > 35:
            actions.append("热电流体：三元积需要调整")
            self.adjustments.append("平衡三参数关系")
        if x < 185 and y < 225:
            actions.append("热电流体：温度电压同时偏低")
            self.alert_status = "medium"
        if x > 215 and y > 235:
            actions.append("热电流体：温度电压同时偏高")
            self.alert_status = "medium"
        if x < 185 and z < 60:
            actions.append("热电流体：温度流量同时偏低")
            self.alert_status = "medium"
        if x > 215 and z > 70:
            actions.append("热电流体：温度流量同时偏高")
            self.alert_status = "medium"
        if y < 225 and z < 60:
            actions.append("热电流体：电压流量同时偏低")
            self.alert_status = "medium"
        if y > 235 and z > 70:
            actions.append("热电流体：电压流量同时偏高")
            self.alert_status = "medium"
        if x < 180 or x > 220:
            actions.append("热电流体：温度在临界范围")
            self.alert_status = "high"
        if y < 218 or y > 242:
            actions.append("热电流体：电压在临界范围")
            self.alert_status = "high"
        if z < 48 or z > 82:
            actions.append("热电流体：流量在临界范围")
            self.alert_status = "high"
        if x < 175 and y < 220 and z < 55:
            actions.append("热电流体：所有参数严重偏低")
            self.alert_status = "emergency"
        if x > 225 and y > 240 and z > 75:
            actions.append("热电流体：所有参数严重偏高")
            self.alert_status = "emergency"
        if ((x - 200) ** 2 + (y / 10 - 23) ** 2 + (z - 65) ** 2) > 100:
            actions.append("热电流体：三维偏离过大")
            self.alert_status = "high"
        if abs((x * y * z / 10000) / 32 - 1) > 0.18:
            actions.append("热电流体：三元积偏离过大")
            self.alert_status = "medium"

        return actions

    def section10_comprehensive_hybrid_control(self, x, y, z):
        """第10类: 综合混合控制 (湿度x, 扭矩y, 电流z) - 80个if语句"""
        actions = []

        # 综合系统基础检查 (1-20)
        if self.standards['humidity'][0] < x < self.standards['humidity'][1]:
            actions.append("综合控制：湿度基础参数正常")
        if self.standards['torque'][0] < y < self.standards['torque'][1]:
            actions.append("综合控制：扭矩基础参数正常")
        if self.standards['current'][0] < z < self.standards['current'][1]:
            actions.append("综合控制：电流基础参数正常")
        if x > 45 and x < 55 and y > 160 and y < 190:
            actions.append("综合控制：湿度扭矩核心区间协调")
        if x > 42 and x < 58 and z > 11 and z < 14:
            actions.append("综合控制：湿度电流匹配良好")
        if y > 155 and y < 195 and z > 11 and z < 14:
            actions.append("综合控制：扭矩电流协调稳定")
        if x > 48 and x < 52:
            actions.append("综合控制：湿度精确控制")
        if y > 173 and y < 177:
            actions.append("综合控制：扭矩精确控制")
        if z > 12.2 and z < 12.8:
            actions.append("综合控制：电流精确控制")
        if abs(x - 50) < 2:
            actions.append("综合控制：湿度稳定性良好")
        if abs(y - 175) < 5:
            actions.append("综合控制：扭矩稳定性良好")
        if abs(z - 12.5) < 0.3:
            actions.append("综合控制：电流稳定性良好")
        if x > 42 and y > 155 and z > 11:
            actions.append("综合控制：所有参数在安全范围")
        if x < 58 and y < 195 and z < 14:
            actions.append("综合控制：所有参数未超上限")
        if x + y + z * 10 > 260 and x + y + z * 10 < 300:
            actions.append("综合控制：综合系统指标正常")
        if x * y > 8000 and x * y > 10000:
            actions.append("综合控制：湿度扭矩乘积正常")
        if x * z > 600 and x * z < 750:
            actions.append("综合控制：湿度电流乘积正常")
        if y * z > 2100 and y * z < 2500:
            actions.append("综合控制：扭矩电流乘积正常")
        if (x + y + z * 10) / 3 > 86 and (x + y + z * 10) / 3 < 100:
            actions.append("综合控制：平均系统参数正常")
        if sqrt(x ** 2 + y ** 2 + (z * 10) ** 2) > 200:
            actions.append("综合控制：系统向量模长正常")

        # 跨域参数分析 (21-40)
        if x / (y / 4) > 1.0 and x / (y / 4) < 1.3:
            actions.append("综合控制：湿度扭矩比理想")
        if x / z > 3.5 and x / z < 4.5:
            actions.append("综合控制：湿度电流比正常")
        if y / z > 12 and y / z < 16:
            actions.append("综合控制：扭矩电流比适当")
        if (x - 40) + (y - 150) / 5 > 15 and (x - 40) + (y - 150) / 5 < 25:
            actions.append("综合控制：湿度扭矩偏差和正常")
        if (x - 40) + (z - 10) * 3 > 15 and (x - 40) + (z - 10) * 3 < 25:
            actions.append("综合控制：湿度电流偏差和正常")
        if (y - 150) / 5 + (z - 10) > 8 and (y - 150) / 5 + (z - 10) < 12:
            actions.append("综合控制：扭矩电流偏差和正常")
        if abs((x - 50) - (y - 175) / 4) < 5:
            actions.append("综合控制：湿度扭矩偏差关系平衡")
        if abs((x - 50) - (z - 12.5) * 3) < 6:
            actions.append("综合控制：湿度电流偏差关系平衡")
        if abs((y - 175) / 14 - (z - 12.5)) < 2:
            actions.append("综合控制：扭矩电流偏差关系平衡")
        if x / (y / 4 + 10) > 0.9 and x / (y / 4 + 10) < 1.2:
            actions.append("综合控制：调整湿度扭矩比正常")
        if y / (z + 5) > 9 and y / (z + 5) < 13:
            actions.append("综合控制：调整扭矩电流比正常")
        if z / (x / 10) > 2.2 and z / (x / 10) < 2.8:
            actions.append("综合控制：调整电流湿度比正常")
        if (x * y * z) > 20000 and (x * y * z) < 30000:
            actions.append("综合控制：三元系统积正常")
        if x * 0.8 + y * 0.1 + z * 8 > 160 and x * 0.8 + y * 0.1 + z * 8 < 180:
            actions.append("综合控制：加权系统和正常")
        if (x / 50) ** 0.3 * (y / 175) ** 0.4 * (z / 12.5) ** 0.3 > 0.9:
            actions.append("综合控制：加权几何平均正常")
        if (x - 50) * (y - 175) / 4 > -40 and (x - 50) * (y - 175) / 4 < 40:
            actions.append("综合控制：湿度扭矩偏差积平衡")
        if (x - 50) * (z - 12.5) > -15 and (x - 50) * (z - 12.5) < 15:
            actions.append("综合控制：湿度电流偏差积平衡")
        if (y - 175) / 14 * (z - 12.5) > -8 and (y - 175) / 14 * (z - 12.5) < 8:
            actions.append("综合控制：扭矩电流偏差积平衡")
        if abs(x / 50 + y / 175 + z / 12.5 - 3) < 0.15:
            actions.append("综合控制：归一化和接近理想")
        if abs((x / 50) * (y / 175) * (z / 12.5) - 1) < 0.1:
            actions.append("综合控制：归一化积接近理想")

        # 全系统优化协调 (41-60)
        if x > 49 and x < 51 and y > 173 and y < 177 and z > 12.3 and z < 12.7:
            actions.append("综合控制：全参数系统优化")
        if abs(x / (y / 4) - 1.14) < 0.1 and abs(z - 12.5) < 0.2:
            actions.append("综合控制：湿度扭矩比与电流协调")
        if x * y > 8700 and x * y < 8800 and z > 12.4 and z < 12.6:
            actions.append("综合控制：湿度扭矩积与电流协调")
        if (x + y / 4) / 2 > 46 and (x + y / 4) / 2 < 48 and z > 12.3 and z < 12.7:
            actions.append("综合控制：湿度扭矩平均与电流协调")
        if abs(x - y / 4) < 6 and z > 12.3 and z < 12.7:
            actions.append("综合控制：湿度扭矩差与电流协调")
        if sqrt((x - 50) ** 2 + (y / 4 - 43.75) ** 2) < 2 and abs(z - 12.5) < 0.2:
            actions.append("综合控制：湿度扭矩距离与电流优秀")
        if x / 50 > 0.98 and x / 50 < 1.02 and y / 175 > 0.985 and y / 175 < 1.015:
            actions.append("综合控制：湿度扭矩相对值协调")
        if z / 12.5 > 0.984 and z / 12.5 < 1.016:
            actions.append("综合控制：电流相对值协调")
        if (x / 50 + y / 175 + z / 12.5) / 3 > 0.985 and (x / 50 + y / 175 + z / 12.5) / 3 < 1.015:
            actions.append("综合控制：归一化平均协调")
        if max(abs(x / 50 - 1), abs(y / 175 - 1), abs(z / 12.5 - 1)) < 0.02:
            actions.append("综合控制：最大偏差协调良好")
        if min(x / 50, y / 175, z / 12.5) > 0.985:
            actions.append("综合控制：最小相对值协调良好")
        if max(x / 50, y / 175, z / 12.5) < 1.015:
            actions.append("综合控制：最大相对值协调良好")
        if abs(max(x, y / 4, z * 4) - min(x, y / 4, z * 4)) < 8:
            actions.append("综合控制：标准化后范围协调")
        if (x + y + z * 10) > 278 and (x + y + z * 10) < 282:
            actions.append("综合控制：综合参数协调优秀")
        if (x * y * z) > 24500 and (x * y * z) < 25500:
            actions.append("综合控制：三元积协调优秀")
        if abs((x + y + z * 10) / 3 - 93.3) < 1:
            actions.append("综合控制：平均值协调优秀")
        if x > 49.5 and x < 50.5 and y > 174 and y < 176:
            actions.append("综合控制：湿度扭矩超精密协调")
        if z > 12.45 and z < 12.55:
            actions.append("综合控制：电流超精密协调")
        if ((x - 50) ** 2 + (y / 4 - 43.75) ** 2 + (z - 12.5) ** 2 * 16) < 16:
            actions.append("综合控制：三维距离协调完美")
        if abs(x / (y / 4) / z - 0.091) < 0.003:
            actions.append("综合控制：连续比例协调完美")

        # 智能预测控制 (61-80)
        if x < 45 or y < 160 or z < 11.5:
            actions.append("综合控制：检测到参数偏低趋势")
            self.adjustments.append("预防性提升综合参数")
        if x > 55 or y > 190 or z > 13.5:
            actions.append("综合控制：检测到参数偏高趋势")
            self.adjustments.append("预防性限制综合负荷")
        if abs(x / (y / 4) - 1.14) > 0.15:
            actions.append("综合控制：湿度扭矩比需要调整")
            self.adjustments.append("优化湿度扭矩协调控制")
        if abs(x / z - 4.0) > 0.5:
            actions.append("综合控制：湿度电流比需要调整")
            self.adjustments.append("优化湿度电流协调控制")
        if abs(y / z - 14) > 2:
            actions.append("综合控制：扭矩电流比需要调整")
            self.adjustments.append("优化扭矩电流协调控制")
        if (x + y + z * 10) < 270 or (x + y + z * 10) > 290:
            actions.append("综合控制：综合指标需要调整")
            self.adjustments.append("全面优化综合参数")
        if (x * y * z) < 22000 or (x * y * z) > 28000:
            actions.append("综合控制：三元积需要调整")
            self.adjustments.append("平衡三参数关系")
        if x < 43 and y < 160:
            actions.append("综合控制：湿度扭矩同时偏低")
            self.alert_status = "medium"
        if x > 57 and y > 190:
            actions.append("综合控制：湿度扭矩同时偏高")
            self.alert_status = "medium"
        if x < 43 and z < 11:
            actions.append("综合控制：湿度电流同时偏低")
            self.alert_status = "medium"
        if x > 57 and z > 14:
            actions.append("综合控制：湿度电流同时偏高")
            self.alert_status = "medium"
        if y < 160 and z < 11:
            actions.append("综合控制：扭矩电流同时偏低")
            self.alert_status = "medium"
        if y > 190 and z > 14:
            actions.append("综合控制：扭矩电流同时偏高")
            self.alert_status = "medium"
        if x < 38 or x > 62:
            actions.append("综合控制：湿度在临界范围")
            self.alert_status = "high"
        if y < 145 or y > 205:
            actions.append("综合控制：扭矩在临界范围")
            self.alert_status = "high"
        if z < 9.5 or z > 15.5:
            actions.append("综合控制：电流在临界范围")
            self.alert_status = "high"
        if x < 35 and y < 155 and z < 10.5:
            actions.append("综合控制：所有参数严重偏低")
            self.alert_status = "emergency"
        if x > 65 and y > 195 and z > 14.5:
            actions.append("综合控制：所有参数严重偏高")
            self.alert_status = "emergency"
        if ((x - 50) ** 2 + (y / 4 - 43.75) ** 2 + (z - 12.5) ** 2 * 16) > 100:
            actions.append("综合控制：三维偏离过大")
            self.alert_status = "high"
        if abs((x * y * z) / 25000 - 1) > 0.2:
            actions.append("综合控制：三元积偏离过大")
            self.alert_status = "medium"

        return actions

        def calculate_quality_score(self, data: Dict[str, float], actions: List[str]) -> float:
            """计算质量分数"""
            base_score = 100.0

            # 根据参数偏离程度扣分
            for param, value in data.items():
                if param in self.standards:
                    min_val, max_val = self.standards[param]
                    if value < min_val or value > max_val:
                        # 偏离程度越大，扣分越多
                        deviation = max(0, min_val - value) + max(0, value - max_val)
                        normalized_deviation = deviation / ((max_val - min_val) / 2)
                        base_score -= min(30, normalized_deviation * 10)

            # 根据警报级别扣分
            alert_penalties = {
                'normal': 0,
                'low': 2,
                'medium': 5,
                'high': 15,
                'critical': 25,
                'emergency': 40
            }
            base_score -= alert_penalties.get(self.alert_status, 0)

            # 根据调整建议数量扣分
            base_score -= len(self.adjustments) * 2

            return max(0, min(100, base_score))

        def calculate_efficiency_score(self, data: Dict[str, float], processing_time: float) -> float:
            """计算效率分数"""
            base_efficiency = 100.0

            # 处理时间影响
            if processing_time > 1.0:
                base_efficiency -= (processing_time - 1.0) * 10

            # 参数稳定性影响
            stability_bonus = 0
            for param, value in data.items():
                if param in self.standards:
                    min_val, max_val = self.standards[param]
                    mid_val = (min_val + max_val) / 2
                    range_val = max_val - min_val

                    # 越接近中心值，稳定性越好
                    deviation = abs(value - mid_val) / (range_val / 2)
                    stability_bonus += max(0, 5 * (1 - deviation))

            efficiency_score = base_efficiency + stability_bonus / len(data)
            return max(0, min(100, efficiency_score))

        def control_quality(self):
            """主控制函数"""
            start_time = time.time()
            self.cycle_counter += 1

            try:
                # 获取所有生产数据
                data = self.get_production_data()

                # 检查传感器健康状态
                sensor_status = self.health_monitor.check_sensor_health(data)
                unhealthy_sensors = [k for k, v in sensor_status.items() if v != 'healthy']

                if unhealthy_sensors:
                    self.logger.warning(f"传感器异常: {unhealthy_sensors}")

                # 更新统计分析器
                self.statistics_analyzer.update_stats(data)

                # 检查警报
                new_alarms = self.alarm_manager.check_alarms(data)
                if new_alarms:
                    for alarm in new_alarms:
                        self.logger.warning(f"新警报: {alarm['message']}")

                print(f"当前生产数据 (周期 {self.cycle_counter}):")
                print(
                    f"  温度: {data['temperature']:.1f}℃  压力: {data['pressure']:.1f}bar  密度: {data['density']:.2f}g/cm³")
                print(
                    f"  湿度: {data['humidity']:.1f}%   速度: {data['speed']:.0f}rpm   振动: {data['vibration']:.2f}mm/s")
                print(
                    f"  流量: {data['flow_rate']:.1f}L/min  电压: {data['voltage']:.1f}V  电流: {data['current']:.1f}A")
                print(f"  转速: {data['rpm']:.0f}rpm  扭矩: {data['torque']:.1f}N·m  功率: {data['power']:.1f}kW")

                # 初始化参数
                self.alert_status = AlertLevel.NORMAL.value
                self.adjustments = []

                # 智能选择控制策略（保持原有逻辑）
                section = 0
                actions = []

                # 根据参数状态选择最合适的控制类别
                temp_issues = not (
                            self.standards['temperature'][0] < data['temperature'] < self.standards['temperature'][1])
                pressure_issues = not (self.standards['pressure'][0] < data['pressure'] < self.standards['pressure'][1])
                density_issues = not (self.standards['density'][0] < data['density'] < self.standards['density'][1])

                humidity_issues = not (self.standards['humidity'][0] < data['humidity'] < self.standards['humidity'][1])
                speed_issues = not (self.standards['speed'][0] < data['speed'] < self.standards['speed'][1])
                vibration_issues = not (
                            self.standards['vibration'][0] < data['vibration'] < self.standards['vibration'][1])

                flow_issues = not (self.standards['flow_rate'][0] < data['flow_rate'] < self.standards['flow_rate'][1])
                voltage_issues = not (self.standards['voltage'][0] < data['voltage'] < self.standards['voltage'][1])
                current_issues = not (self.standards['current'][0] < data['current'] < self.standards['current'][1])

                rpm_issues = not (self.standards['rpm'][0] < data['rpm'] < self.standards['rpm'][1])
                torque_issues = not (self.standards['torque'][0] < data['torque'] < self.standards['torque'][1])
                power_issues = not (self.standards['power'][0] < data['power'] < self.standards['power'][1])

                # 优先级判断逻辑
                if temp_issues or pressure_issues or density_issues:
                    section = 1
                    actions = self.section1_thermal_control(data['temperature'], data['pressure'], data['density'])
                elif humidity_issues or speed_issues or vibration_issues:
                    section = 2
                    actions = self.section2_environmental_control(data['humidity'], data['speed'], data['vibration'])
                elif flow_issues or voltage_issues or current_issues:
                    section = 3
                    actions = self.section3_electrical_control(data['flow_rate'], data['voltage'], data['current'])
                elif rpm_issues or torque_issues or power_issues:
                    section = 4
                    actions = self.section4_mechanical_control(data['rpm'], data['torque'], data['power'])
                elif temp_issues or humidity_issues or vibration_issues:
                    section = 5
                    actions = self.section5_thermal_environment_hybrid(data['temperature'], data['humidity'],
                                                                       data['vibration'])
                elif voltage_issues or rpm_issues or power_issues:
                    section = 6
                    actions = self.section6_electro_mechanical_hybrid(data['voltage'], data['rpm'], data['power'])
                elif flow_issues or pressure_issues or density_issues:
                    section = 7
                    actions = self.section7_flow_pressure_density_hybrid(data['flow_rate'], data['pressure'],
                                                                         data['density'])
                elif speed_issues or torque_issues or current_issues:
                    section = 8
                    actions = self.section8_hybrid_speed_torque_current(data['speed'], data['torque'], data['current'])
                elif temp_issues or voltage_issues or flow_issues:
                    section = 9
                    actions = self.section9_thermal_electrical_flow_hybrid(data['temperature'], data['voltage'],
                                                                           data['flow_rate'])
                else:
                    section = 10
                    actions = self.section10_comprehensive_hybrid_control(data['humidity'], data['torque'],
                                                                          data['current'])

                # 计算处理时间
                processing_time = time.time() - start_time

                # 计算质量和效率分数
                quality_score = self.calculate_quality_score(data, actions)
                efficiency_score = self.calculate_efficiency_score(data, processing_time)

                # 更新预测维护
                self.predictive_maintenance.update_component_health(data, self.alert_status)

                # 创建生产记录
                record = ProductionRecord(
                    timestamp=datetime.datetime.now(),
                    cycle_id=self.cycle_counter,
                    temperature=data['temperature'],
                    pressure=data['pressure'],
                    density=data['density'],
                    humidity=data['humidity'],
                    speed=data['speed'],
                    vibration=data['vibration'],
                    flow_rate=data['flow_rate'],
                    voltage=data['voltage'],
                    current=data['current'],
                    rpm=data['rpm'],
                    torque=data['torque'],
                    power=data['power'],
                    section_used=section,
                    alert_level=self.alert_status,
                    actions_count=len(actions),
                    adjustments_count=len(self.adjustments),
                    quality_score=quality_score,
                    efficiency_score=efficiency_score
                )

                # 记录数据
                self.data_logger.log_production_record(record)
                self.performance_monitor.record_cycle_time(processing_time)

                print(f"\n执行第 {section} 类质量控制策略 (共80个if语句)")
                for action in actions[:5]:  # 显示前5个动作
                    print(f"- {action}")

                if len(actions) > 5:
                    print(f"  ... 以及其他 {len(actions) - 5} 项检查")

                if self.adjustments:
                    print("\n工艺调整建议:")
                    for adj in self.adjustments[:3]:
                        print(f"- {adj}")
                    if len(self.adjustments) > 3:
                        print(f"  ... 以及其他 {len(self.adjustments) - 3} 项调整建议")

                # 显示增强的系统信息
                print(f"\n系统状态:")
                print(f"生产线状态: {self.production_line}")
                print(f"警报级别: {self.alert_status}")
                print(f"需要人工检查: {'是' if self.inspection_required else '否'}")
                print(f"执行的if语句数量: {len(actions)}")
                print(f"质量分数: {quality_score:.1f}")
                print(f"效率分数: {efficiency_score:.1f}")
                print(f"处理时间: {processing_time:.3f}秒")

                # 显示传感器状态
                if unhealthy_sensors:
                    print(f"传感器异常: {', '.join(unhealthy_sensors)}")

                # 显示活跃警报
                active_alarms = self.alarm_manager.get_active_alarms()
                if active_alarms:
                    print(f"活跃警报数量: {len(active_alarms)}")
                    for alarm in active_alarms[:2]:  # 显示前2个
                        print(f"  - {alarm['message']} ({alarm['level']})")

                # 趋势分析
                if self.cycle_counter % 10 == 0:  # 每10个周期分析一次趋势
                    temp_trend = self.statistics_analyzer.detect_trend('temperature')
                    print(f"\n趋势分析:")
                    print(f"温度趋势: {temp_trend['trend']} (置信度: {temp_trend['confidence']:.2f})")

                    # 检查预测维护
                    maintenance_schedule = self.predictive_maintenance.get_maintenance_schedule()
                    if maintenance_schedule:
                        print("维护提醒:")
                        for item in maintenance_schedule[:2]:
                            print(
                                f"  - {item['component']}: {item['risk_level']} 风险，建议 {item['recommended_maintenance_days']} 天内维护")

                # 每100个周期生成报告
                if self.cycle_counter % 100 == 0:
                    try:
                        daily_report = self.report_generator.generate_daily_report()
                        if 'error' not in daily_report:
                            print(f"\n=== 第{self.cycle_counter}周期报告 ===")
                            print(f"平均质量分数: {daily_report['quality_metrics']['average_quality_score']:.1f}")
                            print(f"平均效率分数: {daily_report['efficiency_metrics']['average_efficiency_score']:.1f}")
                            print(
                                f"最常用控制策略: section_{max(daily_report['section_usage'], key=daily_report['section_usage'].get)}")

                            # 导出报告
                            report_filename = f"report_cycle_{self.cycle_counter}.csv"
                            self.report_generator.export_report_to_csv(daily_report, report_filename)
                            print(f"报告已导出到: {report_filename}")
                    except Exception as e:
                        self.logger.error(f"报告生成失败: {e}")

                print(f"{'-' * 90}")

            except Exception as e:
                self.logger.error(f"质量控制处理错误: {e}")
                self.health_monitor.health_metrics['processing_errors'] += 1
                self.health_monitor.health_metrics['last_error_time'] = datetime.datetime.now()
                raise

        def get_system_dashboard(self) -> Dict[str, Any]:
            """获取系统仪表板数据"""
            performance_stats = self.performance_monitor.get_performance_stats()
            system_health = self.health_monitor.get_system_health_score()
            active_alarms = self.alarm_manager.get_active_alarms()
            maintenance_schedule = self.predictive_maintenance.get_maintenance_schedule()

            return {
                'system_info': {
                    'name': self.config_manager.get('system.name'),
                    'version': self.config_manager.get('system.version'),
                    'uptime_hours': performance_stats['uptime_hours'],
                    'current_cycle': self.cycle_counter,
                    'health_score': system_health
                },
                'performance': performance_stats,
                'alarms': {
                    'active_count': len(active_alarms),
                    'critical_count': len([a for a in active_alarms if a['level'] == 'critical']),
                    'recent_alarms': active_alarms[:5]
                },
                'maintenance': {
                    'pending_count': len(maintenance_schedule),
                    'urgent_items': [item for item in maintenance_schedule if
                                     item['risk_level'] in ['critical', 'imminent']]
                }
            }

        def shutdown_system(self):
            """系统关闭"""
            self.logger.info("系统正在关闭...")

            # 创建最终备份
            try:
                backup_path = self.backup_manager.create_backup()
                self.logger.info(f"关闭前备份完成: {backup_path}")
            except Exception as e:
                self.logger.error(f"关闭备份失败: {e}")

            # 生成最终报告
            try:
                final_report = self.report_generator.generate_daily_report()
                if 'error' not in final_report:
                    self.report_generator.export_report_to_csv(final_report, "final_report.csv")
                    self.logger.info("最终报告已生成")
            except Exception as e:
                self.logger.error(f"最终报告生成失败: {e}")

            self.production_line = SystemStatus.SHUTDOWN.value
            self.logger.info("系统已关闭")

    # 系统使用示例
    def run_enhanced_system(self):
        """运行增强版系统"""
        controller = MultiVariableQualityController()

        try:
            print("=== 增强版多变量智能制造质量控制系统启动 ===")
            print(f"系统版本: {controller.config_manager.get('system.version')}")
            print(f"配置加载完成")
            print(f"数据库初始化完成")
            print(f"开始生产监控...")
            print("=" * 70)

            # 模拟生产运行
            for i in range(50):  # 运行50个周期进行演示
                controller.control_quality()

                # 模拟一些交互
                if i == 20:
                    print("\n>>> 系统仪表板数据 <<<")
                    dashboard = controller.get_system_dashboard()
                    print(f"系统健康分数: {dashboard['system_info']['health_score']:.2f}")
                    print(f"平均周期时间: {dashboard['performance']['average_cycle_time']:.3f}秒")
                    print(f"当前吞吐量: {dashboard['performance']['current_throughput']:.1f} 周期/小时")
                    print(f"活跃警报数: {dashboard['alarms']['active_count']}")
                    print(f"待维护项目: {dashboard['maintenance']['pending_count']}")

                if i == 40:
                    print("\n>>> 参数统计分析 <<<")
                    temp_stats = controller.statistics_analyzer.get_parameter_statistics('temperature')
                    if temp_stats:
                        print(f"温度统计: 平均值={temp_stats['mean']:.1f}, 标准差={temp_stats['std_dev']:.2f}")

                    pressure_trend = controller.statistics_analyzer.detect_trend('pressure')
                    print(f"压力趋势: {pressure_trend['trend']} (置信度: {pressure_trend.get('confidence', 0):.2f})")

                # 短暂延时以便观察
                time.sleep(0.1)

            print("\n>>> 生成数据可视化 <<<")
            controller.data_visualization.plot_parameter_trend('temperature', hours=1,
                                                               save_path='temperature_trend.png')
            controller.data_visualization.plot_quality_distribution(hours=1, save_path='quality_distribution.png')
            print("图表已保存")

        except KeyboardInterrupt:
            print("\n用户中断，正在安全关闭系统...")
        except Exception as e:
            print(f"系统运行错误: {e}")
            controller.logger.error(f"系统运行错误: {e}")
        finally:
            controller.shutdown_system()
            print("系统已安全关闭")

    class MLPredictor:
        """机器学习预测器"""

        def __init__(self):
            self.anomaly_detector = IsolationForest(contamination=0.1, random_state=42)
            self.quality_predictor = RandomForestRegressor(n_estimators=100, random_state=42)
            self.scaler = StandardScaler()
            self.is_trained = False

        def prepare_features(self, records: List[ProductionRecord]) -> np.ndarray:
            """准备特征数据"""
            features = []
            for record in records:
                feature_row = [
                    record.temperature, record.pressure, record.density,
                    record.humidity, record.speed, record.vibration,
                    record.flow_rate, record.voltage, record.current,
                    record.rpm, record.torque, record.power,
                    record.actions_count, record.adjustments_count
                ]
                features.append(feature_row)
            return np.array(features)

        def train_models(self, records: List[ProductionRecord]):
            """训练机器学习模型"""
            if len(records) < 100:
                return False

            features = self.prepare_features(records)
            quality_scores = np.array([r.quality_score for r in records])

            # 标准化特征
            features_scaled = self.scaler.fit_transform(features)

            # 训练异常检测模型
            self.anomaly_detector.fit(features_scaled)

            # 训练质量预测模型
            X_train, X_test, y_train, y_test = train_test_split(
                features_scaled, quality_scores, test_size=0.2, random_state=42
            )
            self.quality_predictor.fit(X_train, y_train)

            # 评估模型
            y_pred = self.quality_predictor.predict(X_test)
            mse = mean_squared_error(y_test, y_pred)

            self.is_trained = True
            return {'mse': mse, 'training_samples': len(records)}

        def detect_anomaly(self, data: Dict[str, float]) -> Dict[str, Any]:
            """检测异常"""
            if not self.is_trained:
                return {'is_anomaly': False, 'confidence': 0}

            feature_row = np.array([[
                data['temperature'], data['pressure'], data['density'],
                data['humidity'], data['speed'], data['vibration'],
                data['flow_rate'], data['voltage'], data['current'],
                data['rpm'], data['torque'], data['power'], 0, 0
            ]])

            feature_scaled = self.scaler.transform(feature_row)
            anomaly_score = self.anomaly_detector.decision_function(feature_scaled)[0]
            is_anomaly = self.anomaly_detector.predict(feature_scaled)[0] == -1

            return {
                'is_anomaly': is_anomaly,
                'anomaly_score': anomaly_score,
                'confidence': abs(anomaly_score)
            }

        def predict_quality(self, data: Dict[str, float]) -> float:
            """预测质量分数"""
            if not self.is_trained:
                return 75.0  # 默认值

            feature_row = np.array([[
                data['temperature'], data['pressure'], data['density'],
                data['humidity'], data['speed'], data['vibration'],
                data['flow_rate'], data['voltage'], data['current'],
                data['rpm'], data['torque'], data['power'], 0, 0
            ]])

            feature_scaled = self.scaler.transform(feature_row)
            predicted_quality = self.quality_predictor.predict(feature_scaled)[0]
            return max(0, min(100, predicted_quality))

        def save_models(self, path: str = "models/"):
            """保存模型"""
            os.makedirs(path, exist_ok=True)
            joblib.dump(self.anomaly_detector, f"{path}anomaly_detector.pkl")
            joblib.dump(self.quality_predictor, f"{path}quality_predictor.pkl")
            joblib.dump(self.scaler, f"{path}scaler.pkl")

        def load_models(self, path: str = "models/"):
            """加载模型"""
            try:
                self.anomaly_detector = joblib.load(f"{path}anomaly_detector.pkl")
                self.quality_predictor = joblib.load(f"{path}quality_predictor.pkl")
                self.scaler = joblib.load(f"{path}scaler.pkl")
                self.is_trained = True
                return True
            except:
                return False


def __init__(self, quality_controller):
    self.controller = quality_controller
    self.websocket_clients = set()
    self.flask_app = Flask(__name__)
    self.setup_routes()

def setup_routes(self):
    """设置REST API路由"""

    @self.flask_app.route('/api/status', methods=['GET'])
    def get_status():
        dashboard = self.controller.get_system_dashboard()
        return jsonify(dashboard)

    @self.flask_app.route('/api/alarms', methods=['GET'])
    def get_alarms():
        alarms = self.controller.alarm_manager.get_active_alarms()
        return jsonify(alarms)

    @self.flask_app.route('/api/parameters', methods=['GET'])
    def get_parameters():
        data = self.controller.get_production_data()
        return jsonify(data)

    @self.flask_app.route('/api/reports/daily', methods=['GET'])
    def get_daily_report():
        report = self.controller.report_generator.generate_daily_report()
        return jsonify(report)

    @self.flask_app.route('/api/maintenance', methods=['GET'])
    def get_maintenance():
        schedule = self.controller.predictive_maintenance.get_maintenance_schedule()
        return jsonify(schedule)

    @self.flask_app.route('/api/config', methods=['GET', 'POST'])
    def handle_config():
        if request.method == 'GET':
            return jsonify(self.controller.config_manager.config)
        else:
            new_config = request.json
            for key, value in new_config.items():
                self.controller.config_manager.set(key, value)
            return jsonify({'status': 'updated'})

async def websocket_handler(self, websocket, path, websockets=None):
    """WebSocket连接处理"""
    self.websocket_clients.add(websocket)
    try:
        async for message in websocket:
            data = json.loads(message)
            response = await self.handle_websocket_message(data)
            if response:
                await websocket.send(json.dumps(response))
    except websockets.exceptions.ConnectionClosed:
        pass
    finally:
        self.websocket_clients.remove(websocket)

async def handle_websocket_message(self, data):
    """处理WebSocket消息"""
    msg_type = data.get('type')

    if msg_type == 'subscribe':
        return {'type': 'subscribed', 'status': 'success'}
    elif msg_type == 'get_status':
        dashboard = self.controller.get_system_dashboard()
        return {'type': 'status_update', 'data': dashboard}
    elif msg_type == 'acknowledge_alarm':
        alarm_id = data.get('alarm_id')
        self.controller.alarm_manager.acknowledge_alarm(alarm_id)
        return {'type': 'alarm_acknowledged', 'alarm_id': alarm_id}

    return None

async def broadcast_update(self, update_type, data):
    """向所有WebSocket客户端广播更新"""
    if self.websocket_clients:
        message = json.dumps({
            'type': update_type,
            'data': data,
            'timestamp': datetime.datetime.now().isoformat()
        })
        await asyncio.gather(
            *[client.send(message) for client in self.websocket_clients],
            return_exceptions=True
        )

def start_websocket_server(self, host='localhost', port=8765, websockets=None):
    """启动WebSocket服务器"""
    start_server = websockets.serve(self.websocket_handler, host, port)
    asyncio.get_event_loop().run_until_complete(start_server)

def start_rest_api(self, host='localhost', port=5000):
    """启动REST API服务器"""
    self.flask_app.run(host=host, port=port, threaded=True)
    class AdvancedAnalytics:
        """高级分析模块"""

        def __init__(self, window_size=1000):
            self.window_size = window_size
            self.frequency_data = defaultdict(lambda: deque(maxlen=window_size))

        def fourier_analysis(self, parameter: str, values: List[float]) -> Dict[str, Any]:
            """傅里叶频域分析"""
            if len(values) < 64:
                return {'error': 'insufficient_data'}

            # 计算FFT
            fft_values = scipy.fft.fft(values)
            frequencies = scipy.fft.fftfreq(len(values))

            # 找到主要频率分量
            magnitudes = np.abs(fft_values)
            dominant_freq_idx = np.argmax(magnitudes[1:len(magnitudes) // 2]) + 1
            dominant_frequency = frequencies[dominant_freq_idx]

            return {
                'dominant_frequency': float(dominant_frequency),
                'magnitude': float(magnitudes[dominant_freq_idx]),
                'total_energy': float(np.sum(magnitudes ** 2)),
                'spectral_centroid': float(np.sum(frequencies * magnitudes) / np.sum(magnitudes))
            }

        def process_capability_analysis(self, values: List[float], lower_spec: float,
                                        upper_spec: float, target: float = None) -> Dict[str, float]:
            """过程能力分析"""
            if len(values) < 30:
                return {'error': 'insufficient_data'}

            mean_val = np.mean(values)
            std_val = np.std(values, ddof=1)

            if target is None:
                target = (lower_spec + upper_spec) / 2

            # 计算过程能力指数
            cp = (upper_spec - lower_spec) / (6 * std_val)  # 过程潜在能力
            cpk = min((upper_spec - mean_val) / (3 * std_val),
                      (mean_val - lower_spec) / (3 * std_val))  # 过程实际能力
            cpm = cp / np.sqrt(1 + ((mean_val - target) / std_val) ** 2)  # 修正过程能力

            return {
                'cp': cp,
                'cpk': cpk,
                'cpm': cpm,
                'mean': mean_val,
                'std': std_val,
                'interpretation': self._interpret_capability(cp, cpk)
            }



        def _interpret_capability(self, cp: float, cpk: float) -> str:
            """解释过程能力"""
            if cpk >= 1.33:
                return "excellent"
            elif cpk >= 1.0:
                return "adequate"
            elif cpk >= 0.67:
                return "marginal"
            else:
                return "inadequate"

        def correlation_analysis(self, data_dict: Dict[str, List[float]]) -> Dict[str, Dict[str, float]]:
            """参数相关性分析"""
            parameters = list(data_dict.keys())
            correlation_matrix = {}

            for i, param1 in enumerate(parameters):
                correlation_matrix[param1] = {}
                for j, param2 in enumerate(parameters):
                    if len(data_dict[param1]) == len(data_dict[param2]) and len(data_dict[param1]) > 10:
                        corr_coef, p_value = scipy.stats.pearsonr(data_dict[param1], data_dict[param2])
                        correlation_matrix[param1][param2] = {
                            'correlation': float(corr_coef),
                            'p_value': float(p_value),
                            'significant': p_value < 0.05
                        }



            return correlation_matrix

        def detect_cycles(self, values: List[float], min_cycle_length: int = 10) -> List[Dict[str, Any]]:
            """检测周期性模式"""
            if len(values) < min_cycle_length * 3:
                return []

            # 使用自相关检测周期
            autocorr = np.correlate(values, values, mode='full')
            autocorr = autocorr[autocorr.size // 2:]

            # 寻找峰值
            peaks, _ = scipy.signal.find_peaks(autocorr[min_cycle_length:],
                                               height=0.3 * np.max(autocorr))

            cycles = []
            for peak in peaks:
                cycle_length = peak + min_cycle_length
                confidence = autocorr[cycle_length] / np.max(autocorr)

                cycles.append({
                    'cycle_length': int(cycle_length),
                    'confidence': float(confidence),
                    'strength': float(autocorr[cycle_length])
                })

            return sorted(cycles, key=lambda x: x['confidence'], reverse=True)[:5]




        class NotificationSystem:
            """通知系统"""

            def __init__(self, config_manager):
                self.config = config_manager
                self.notification_history = deque(maxlen=1000)

            def send_email(self, subject: str, body: str, recipients: List[str]) -> bool:
                """发送邮件通知"""
                try:
                    smtp_server = self.config.get('notifications.email.smtp_server', 'smtp.gmail.com')
                    smtp_port = self.config.get('notifications.email.smtp_port', 587)
                    username = self.config.get('notifications.email.username')
                    password = self.config.get('notifications.email.password')

                    if not username or not password:
                        return False


                    msg = MIMEMultipart()
                    msg['From'] = username
                    msg['Subject'] = subject
                    msg.attach(MIMEText(body, 'plain'))

                    server = smtplib.SMTP(smtp_server, smtp_port)
                    server.starttls()
                    server.login(username, password)

                    for recipient in recipients:
                        msg['To'] = recipient
                        server.send_message(msg)
                        del msg['To']

                    server.quit()

                    self.notification_history.append({
                        'type': 'email',
                        'timestamp': datetime.datetime.now(),
                        'subject': subject,
                        'recipients': recipients,
                        'status': 'sent'
                    })

                    return True

                except Exception as e:
                    self.notification_history.append({
                        'type': 'email',
                        'timestamp': datetime.datetime.now(),
                        'subject': subject,
                        'recipients': recipients,
                        'status': 'failed',
                        'error': str(e)
                    })
                    return False

            def send_webhook(self, url: str, data: Dict[str, Any]) -> bool:
                """发送Webhook通知"""
                try:
                    response = requests.post(url, json=data, timeout=10)
                    response.raise_for_status()

                    self.notification_history.append({
                        'type': 'webhook',
                        'timestamp': datetime.datetime.now(),
                        'url': url,
                        'status': 'sent',
                        'response_code': response.status_code
                    })

                    return True

                except Exception as e:
                    self.notification_history.append({
                        'type': 'webhook',
                        'timestamp': datetime.datetime.now(),
                        'url': url,
                        'status': 'failed',
                        'error': str(e)
                    })
                    return False


            def send_slack_message(self, message: str, channel: str = None) -> bool:
                """发送Slack消息"""
                webhook_url = self.config.get('notifications.slack.webhook_url')
                if not webhook_url:
                    return False

                payload = {
                    'text': message,
                    'username': 'QualityBot',
                    'icon_emoji': ':gear:'
                }

                if channel:
                    payload['channel'] = channel

                return self.send_webhook(webhook_url, payload)



            def notify_alarm(self, alarm: Dict[str, Any]) -> List[str]:
                """发送警报通知"""
                sent_methods = []

                # 根据警报级别确定通知方式
                if alarm['level'] in ['critical', 'emergency']:
                    # 紧急警报 - 所有渠道
                    email_recipients = self.config.get('notifications.email.emergency_recipients', [])
                    if email_recipients:
                        subject = f"紧急警报: {alarm['message']}"
                        body = f"""
                        警报级别: {alarm['level']}
                        参数: {alarm['parameter']}
                        当前值: {alarm['value']}
                        阈值: {alarm.get('threshold', 'N/A')}
                        时间: {alarm['timestamp']}

                        请立即检查系统状态。
                        """
                        if self.send_email(subject, body, email_recipients):
                            sent_methods.append('email')

                    # Slack通知
                    slack_message = f"🚨 紧急警报: {alarm['message']} (级别: {alarm['level']})"
                    if self.send_slack_message(slack_message):
                        sent_methods.append('slack')

                elif alarm['level'] == 'high':
                    # 高级警报 - Slack + 部分邮件
                    slack_message = f"⚠️ 高级警报: {alarm['message']}"
                    if self.send_slack_message(slack_message):
                        sent_methods.append('slack')

                return sent_methods



            def notify_maintenance_due(self, maintenance_item: Dict[str, Any]):
                """维护到期通知"""
                message = f"维护提醒: {maintenance_item['component']} 需要维护 (风险级别: {maintenance_item['risk_level']})"

                if maintenance_item['risk_level'] in ['critical', 'imminent']:
                    # 紧急维护
                    email_recipients = self.config.get('notifications.email.maintenance_recipients', [])
                    if email_recipients:
                        subject = "紧急维护提醒"
                        body = f"""
                        设备: {maintenance_item['component']}
                        健康分数: {maintenance_item['health_score']}
                        风险级别: {maintenance_item['risk_level']}
                        建议维护时间: {maintenance_item['recommended_maintenance_days']} 天内

                        请安排维护人员检查设备。
                        """
                        self.send_email(subject, body, email_recipients)

                self.send_slack_message(f"🔧 {message}")

            def get_notification_history(self, limit: int = 50) -> List[Dict[str, Any]]:
                """获取通知历史"""
                return list(self.notification_history)[-limit:]

        class EnergyManagement:
            """能耗管理模块"""



            def __init__(self):
                self.energy_data = deque(maxlen=10000)
                self.baseline_power = 0
                self.efficiency_targets = {
                    'power_factor': 0.95,
                    'energy_per_cycle': 2.5,  # kWh per production cycle
                    'idle_power_ratio': 0.3
                }



            def record_energy_data(self, voltage: float, current: float, power: float,
                                   cycle_active: bool = True):
                """记录能耗数据"""
                timestamp = datetime.datetime.now()
                apparent_power = voltage * current
                power_factor = power / apparent_power if apparent_power > 0 else 0

                energy_record = {
                    'timestamp': timestamp,
                    'voltage': voltage,
                    'current': current,
                    'active_power': power,
                    'apparent_power': apparent_power,
                    'power_factor': power_factor,
                    'cycle_active': cycle_active
                }

                self.energy_data.append(energy_record)

                # 更新基准功率
                if cycle_active:
                    recent_active_power = [r['active_power'] for r in self.energy_data
                                           if r['cycle_active'] and
                                           (timestamp - r['timestamp']).seconds < 300]
                    if recent_active_power:
                        self.baseline_power = statistics.mean(recent_active_power)



            def calculate_energy_efficiency(self, time_window_hours: float = 1.0) -> Dict[str, float]:
                """计算能效指标"""
                cutoff_time = datetime.datetime.now() - datetime.timedelta(hours=time_window_hours)
                recent_data = [r for r in self.energy_data if r['timestamp'] > cutoff_time]

                if not recent_data:
                    return {'error': 'no_data'}

                active_data = [r for r in recent_data if r['cycle_active']]
                idle_data = [r for r in recent_data if not r['cycle_active']]

                # 计算各项指标
                avg_power_factor = statistics.mean([r['power_factor'] for r in recent_data])
                total_energy = sum([r['active_power'] for r in recent_data]) / 60  # kWh

                active_energy = sum([r['active_power'] for r in active_data]) / 60 if active_data else 0
                idle_energy = sum([r['active_power'] for r in idle_data]) / 60 if idle_data else 0

                idle_ratio = idle_energy / total_energy if total_energy > 0 else 0

                # 计算效率评分
                pf_score = min(100, (avg_power_factor / self.efficiency_targets['power_factor']) * 100)
                idle_score = max(0, 100 - (idle_ratio / self.efficiency_targets['idle_power_ratio']) * 100)

                overall_efficiency = (pf_score + idle_score) / 2

                return {
                    'overall_efficiency': overall_efficiency,
                    'power_factor': avg_power_factor,
                    'power_factor_score': pf_score,
                    'total_energy_kwh': total_energy,
                    'active_energy_kwh': active_energy,
                    'idle_energy_kwh': idle_energy,
                    'idle_ratio': idle_ratio,
                    'idle_score': idle_score,
                    'baseline_power': self.baseline_power
                }



            def detect_energy_anomalies(self) -> List[Dict[str, Any]]:
                """检测能耗异常"""
                anomalies = []

                if len(self.energy_data) < 100:
                    return anomalies

                recent_power = [r['active_power'] for r in list(self.energy_data)[-100:]]
                mean_power = statistics.mean(recent_power)
                std_power = statistics.stdev(recent_power)

                # 检测功率尖峰
                for record in list(self.energy_data)[-10:]:
                    z_score = abs(record['active_power'] - mean_power) / std_power
                    if z_score > 3:
                        anomalies.append({
                            'type': 'power_spike',
                            'timestamp': record['timestamp'],
                            'value': record['active_power'],
                            'expected_range': f"{mean_power - 2 * std_power:.1f}-{mean_power + 2 * std_power:.1f}",
                            'severity': 'high' if z_score > 4 else 'medium'
                        })



                # 检测功率因数异常
                recent_pf = [r['power_factor'] for r in list(self.energy_data)[-50:]]
                avg_pf = statistics.mean(recent_pf)
                if avg_pf < 0.85:
                    anomalies.append({
                        'type': 'low_power_factor',
                        'timestamp': datetime.datetime.now(),
                        'value': avg_pf,
                        'threshold': 0.85,
                        'severity': 'high' if avg_pf < 0.75 else 'medium'
                    })

                return anomalies



            def generate_energy_report(self) -> Dict[str, Any]:
                """生成能耗报告"""
                daily_efficiency = self.calculate_energy_efficiency(24)
                hourly_efficiency = self.calculate_energy_efficiency(1)
                anomalies = self.detect_energy_anomalies()

                # 计算成本（假设电价）
                electricity_rate = 0.12  # $/kWh
                daily_cost = daily_efficiency.get('total_energy_kwh', 0) * electricity_rate

                return {
                    'daily_efficiency': daily_efficiency,
                    'hourly_efficiency': hourly_efficiency,
                    'anomalies': anomalies,
                    'estimated_daily_cost': daily_cost,
                    'efficiency_targets': self.efficiency_targets,
                    'recommendations': self._generate_energy_recommendations(daily_efficiency)
                }





            def _generate_energy_recommendations(self, efficiency_data: Dict[str, float]) -> List[str]:
                """生成节能建议"""
                recommendations = []

                pf = efficiency_data.get('power_factor', 1.0)
                if pf < 0.9:
                    recommendations.append("建议安装功率因数补偿设备提高功率因数")

                idle_ratio = efficiency_data.get('idle_ratio', 0)
                if idle_ratio > 0.4:
                    recommendations.append("建议优化生产计划减少空闲时间")

                if efficiency_data.get('overall_efficiency', 100) < 75:
                    recommendations.append("建议进行能效审计识别改进机会")

                return recommendations