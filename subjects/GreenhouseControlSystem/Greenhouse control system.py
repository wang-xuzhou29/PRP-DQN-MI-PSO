class GreenhouseController:
    def __init__(self):
        # 初始化所有设备状态
        self.lights = "off"
        self.water_pump = "off"
        self.co2_generator = "off"
        self.vent_fan = "off"
        self.shade_screen = "off"
        self.heating = "off"
        self.cooling = "off"
        self.dehumidifier = "off"

        # 设定理想范围
        self.light_ideal_low = 3000
        self.light_ideal_high = 8000
        self.moisture_ideal_low = 60
        self.moisture_ideal_high = 80
        self.co2_ideal_low = 800
        self.co2_ideal_high = 1500

    def get_sensor_data(self):
        """获取传感器数据"""
        import random
        light = random.uniform(1000, 10000)
        moisture = random.uniform(30, 90)
        co2 = random.uniform(500, 2000)
        return light, moisture, co2

    import time
    import json
    from datetime import datetime, timedelta
    import random

    class GreenhouseController:
        def __init__(self):
            # 初始化所有设备状态
            self.lights = "off"
            self.water_pump = "off"
            self.co2_generator = "off"
            self.vent_fan = "off"
            self.shade_screen = "off"
            self.heating = "off"
            self.cooling = "off"
            self.dehumidifier = "off"

            # 新设备状态
            self.fog_system = "off"
            self.nutrient_pump = "off"
            self.irrigation_system = "off"
            self.uv_light = "off"
            self.circulation_fan = "off"
            self.emergency_power = "off"
            self.backup_heater = "off"
            self.drain_valve = "closed"
            self.air_circulation = "off"

            # 设定理想范围
            self.light_ideal_low = 3000
            self.light_ideal_high = 8000
            self.moisture_ideal_low = 60
            self.moisture_ideal_high = 80
            self.co2_ideal_low = 800
            self.co2_ideal_high = 1500
            self.temp_ideal_low = 20
            self.temp_ideal_high = 28
            self.humidity_ideal_low = 50
            self.humidity_ideal_high = 70

            # 数据记录
            self.sensor_history = []
            self.action_history = []
            self.energy_consumption = {
                'lights': 0,
                'water_pump': 0,
                'co2_generator': 0,
                'vent_fan': 0,
                'heating': 0,
                'cooling': 0,
                'dehumidifier': 0,
                'total': 0
            }

            # 系统状态
            self.operating_mode = "auto"  # auto, manual, emergency
            self.maintenance_mode = False
            self.last_maintenance = time.time()
            self.system_errors = []

            # 能源管理
            self.energy_price_tier = "normal"  # low, normal, high
            self.time_of_day = datetime.now().hour
            self.weather_forecast = "sunny"  # sunny, cloudy, rainy, stormy

            # 作物特定参数
            self.crop_type = "tomato"
            self.growth_stage = "vegetative"  # vegetative, flowering, fruiting

            # 报警阈值
            self.alarm_thresholds = {
                'temp_high': 35,
                'temp_low': 10,
                'humidity_high': 90,
                'humidity_low': 30,
                'co2_high': 2000,
                'co2_low': 400,
                'moisture_low': 40
            }

        def record_sensor_data(self, light, moisture, co2, temp, humidity, soil_ph=None, nutrient_level=None,
                               leaf_wetness=None):
            """详细记录传感器数据历史"""
            current_time = time.time()
            record = {
                'timestamp': current_time,
                'datetime': datetime.fromtimestamp(current_time).strftime('%Y-%m-%d %H:%M:%S'),
                'sensors': {
                    'light': round(light, 2),
                    'moisture': round(moisture, 2),
                    'co2': round(co2, 2),
                    'temperature': round(temp, 2),
                    'humidity': round(humidity, 2),
                    'soil_ph': round(soil_ph, 2) if soil_ph else None,
                    'nutrient_level': round(nutrient_level, 2) if nutrient_level else None,
                    'leaf_wetness': round(leaf_wetness, 2) if leaf_wetness else None
                },
                'status': {
                    'in_ideal_range': self._check_ideal_range(light, moisture, co2, temp, humidity),
                    'alerts': self._check_alerts(temp, humidity, co2, moisture),
                    'trends': self._calculate_trends()
                }
            }

            # 添加到历史记录
            self.sensor_history.append(record)

            # 保持历史数据长度（最近24小时的数据，假设每分钟记录一次）
            max_records = 24 * 60  # 24小时 * 60分钟
            if len(self.sensor_history) > max_records:
                # 移除最旧的数据，但保留每小时的第一条记录作为长期历史
                self._compress_history()

            # 更新系统状态
            self._update_system_status()

            # 记录数据质量
            self._check_data_quality(record)

            return record

        def _check_ideal_range(self, light, moisture, co2, temp, humidity):
            """检查传感器数据是否在理想范围内"""
            return {
                'light': self.light_ideal_low <= light <= self.light_ideal_high,
                'moisture': self.moisture_ideal_low <= moisture <= self.moisture_ideal_high,
                'co2': self.co2_ideal_low <= co2 <= self.co2_ideal_high,
                'temperature': self.temp_ideal_low <= temp <= self.temp_ideal_high,
                'humidity': self.humidity_ideal_low <= humidity <= self.humidity_ideal_high
            }

        def _check_alerts(self, temp, humidity, co2, moisture):
            """检查是否需要触发警报"""
            alerts = []

            if temp >= self.alarm_thresholds['temp_high']:
                alerts.append('HIGH_TEMPERATURE')
            elif temp <= self.alarm_thresholds['temp_low']:
                alerts.append('LOW_TEMPERATURE')

            if humidity >= self.alarm_thresholds['humidity_high']:
                alerts.append('HIGH_HUMIDITY')
            elif humidity <= self.alarm_thresholds['humidity_low']:
                alerts.append('LOW_HUMIDITY')

            if co2 >= self.alarm_thresholds['co2_high']:
                alerts.append('HIGH_CO2')
            elif co2 <= self.alarm_thresholds['co2_low']:
                alerts.append('LOW_CO2')

            if moisture <= self.alarm_thresholds['moisture_low']:
                alerts.append('LOW_MOISTURE')

            return alerts

        def _calculate_trends(self):
            """计算数据趋势"""
            if len(self.sensor_history) < 5:
                return {'stable': True}

            recent_data = self.sensor_history[-5:]
            trends = {}

            # 计算温度趋势
            temps = [r['sensors']['temperature'] for r in recent_data]
            temp_trend = 'stable'
            if len(temps) >= 2:
                if temps[-1] - temps[0] > 2:
                    temp_trend = 'rising'
                elif temps[-1] - temps[0] < -2:
                    temp_trend = 'falling'
            trends['temperature'] = temp_trend

            # 计算湿度趋势
            humidities = [r['sensors']['humidity'] for r in recent_data]
            humidity_trend = 'stable'
            if len(humidities) >= 2:
                if humidities[-1] - humidities[0] > 5:
                    humidity_trend = 'rising'
                elif humidities[-1] - humidities[0] < -5:
                    humidity_trend = 'falling'
            trends['humidity'] = humidity_trend

            return trends

        def _compress_history(self):
            """压缩历史数据，保留重要记录"""
            if len(self.sensor_history) <= 1440:  # 24小时数据
                return

            # 保留每小时的第一条记录和所有警报记录
            compressed = []
            last_hour = -1

            for record in self.sensor_history:
                record_hour = datetime.fromtimestamp(record['timestamp']).hour
                has_alerts = len(record['status']['alerts']) > 0

                if record_hour != last_hour or has_alerts:
                    compressed.append(record)
                    last_hour = record_hour

                # 始终保持最近10条记录
                if len(compressed) > 10 and len(self.sensor_history) - len(compressed) > 10:
                    self.sensor_history = compressed + self.sensor_history[-10:]
                else:
                    self.sensor_history = self.sensor_history[-1000:]  # 限制总长度

        def _check_data_quality(self, record):
            """检查数据质量"""
            sensors = record['sensors']
            quality_issues = []

            # 检查传感器值是否在合理范围内
            if sensors['light'] < 0 or sensors['light'] > 20000:
                quality_issues.append('LIGHT_SENSOR_ERROR')

            if sensors['moisture'] < 0 or sensors['moisture'] > 100:
                quality_issues.append('MOISTURE_SENSOR_ERROR')

            if sensors['co2'] < 300 or sensors['co2'] > 5000:
                quality_issues.append('CO2_SENSOR_ERROR')

            if sensors['temperature'] < -10 or sensors['temperature'] > 50:
                quality_issues.append('TEMPERATURE_SENSOR_ERROR')

            if sensors['humidity'] < 0 or sensors['humidity'] > 100:
                quality_issues.append('HUMIDITY_SENSOR_ERROR')

            if quality_issues:
                self.system_errors.extend(quality_issues)
                print(f"数据质量问题: {quality_issues}")

        def _update_system_status(self):
            """更新系统状态"""
            # 检查是否需要维护
            current_time = time.time()
            maintenance_interval = 7 * 24 * 3600  # 7天
            if current_time - self.last_maintenance > maintenance_interval:
                self.maintenance_mode = True
                print("系统需要维护")

            # 检查系统错误数量
            if len(self.system_errors) > 10:
                self.operating_mode = "emergency"
                print("系统进入紧急模式")

        def reset_all_devices(self):
            """详细重置所有设备状态"""
            print("正在重置所有设备状态...")

            # 主要环境控制设备
            self.lights = "off"
            self.water_pump = "off"
            self.co2_generator = "off"
            self.vent_fan = "off"
            self.shade_screen = "off"
            self.heating = "off"
            self.cooling = "off"
            self.dehumidifier = "off"

            # 辅助设备
            self.fog_system = "off"
            self.nutrient_pump = "off"
            self.irrigation_system = "off"
            self.uv_light = "off"
            self.circulation_fan = "off"
            self.emergency_power = "off"
            self.backup_heater = "off"
            self.drain_valve = "closed"
            self.air_circulation = "off"

            # 记录重置操作
            reset_record = {
                'timestamp': time.time(),
                'action': 'SYSTEM_RESET',
                'details': 'All devices reset to off state',
                'mode': self.operating_mode
            }
            self.action_history.append(reset_record)

            print("所有设备已重置")

        def energy_optimization(self):
            """详细的能量优化策略"""
            actions = []
            current_hour = datetime.now().hour

            # 根据时间段设置能源价格等级
            if 23 <= current_hour or current_hour <= 6:  # 夜间
                self.energy_price_tier = "low"
            elif 7 <= current_hour <= 16:  # 白天
                self.energy_price_tier = "normal"
            else:  # 傍晚高峰
                self.energy_price_tier = "high"

            # 根据能源价格等级优化设备运行
            if self.energy_price_tier == "high":
                actions.extend(self._high_energy_price_optimization())
            elif self.energy_price_tier == "normal":
                actions.extend(self._normal_energy_price_optimization())
            else:  # low
                actions.extend(self._low_energy_price_optimization())

            # 根据天气预报优化
            actions.extend(self._weather_based_optimization())

            # 设备运行时间优化
            actions.extend(self._runtime_optimization())

            # 负载平衡
            actions.extend(self._load_balancing())

            return actions

        def _high_energy_price_optimization(self):
            """高电价时段的优化策略"""
            actions = []

            # 减少高能耗设备运行
            if self.lights in ["high", "medium"]:
                actions.append("高电价时段：降低补光灯强度")
                self.lights = "low"

            if self.heating in ["high", "medium"]:
                actions.append("高电价时段：降低加热功率，使用备用措施")
                self.heating = "low"
                # 考虑使用遮阳网保温

            if self.cooling in ["high", "medium"]:
                actions.append("高电价时段：优先使用自然通风降温")
                self.vent_fan = "high"
                self.cooling = "low"

            # 延迟非关键任务
            if self.nutrient_pump != "off":
                actions.append("高电价时段：延迟营养液补充")
                self.nutrient_pump = "off"

            return actions

        def _normal_energy_price_optimization(self):
            """正常电价时段的优化策略"""
            actions = []

            # 平衡能耗和作物需求
            if self.weather_forecast == "sunny":
                actions.append("晴天预报：利用自然光照，减少补光")
                if self.lights != "off":
                    self.lights = "low"

            # 优化设备运行顺序
            actions.append("正常电价：优化设备运行序列")

            return actions

        def _low_energy_price_optimization(self):
            """低电价时段的优化策略"""
            actions = []

            # 在低电价时段执行高能耗任务
            actions.append("低电价时段：执行高能耗任务")

            # 可以提前准备
            if self.water_pump == "off":
                actions.append("低电价时段：提前储水")
                self.water_pump = "low"

            # 可以提前降温/加热
            current_temp = self.sensor_history[-1]['sensors']['temperature'] if self.sensor_history else 25
            if current_temp > self.temp_ideal_high:
                actions.append("低电价时段：预降温")
                self.cooling = "low"

            return actions

        def _weather_based_optimization(self):
            """基于天气预报的优化"""
            actions = []

            if self.weather_forecast == "sunny":
                actions.extend(self._sunny_weather_optimization())
            elif self.weather_forecast == "cloudy":
                actions.extend(self._cloudy_weather_optimization())
            elif self.weather_forecast == "rainy":
                actions.extend(self._rainy_weather_optimization())
            elif self.weather_forecast == "stormy":
                actions.extend(self._stormy_weather_optimization())

            return actions

        def _sunny_weather_optimization(self):
            """晴天优化策略"""
            actions = []
            actions.append("晴天预报：准备遮阳和降温")

            # 准备遮阳系统
            if self.shade_screen == "off":
                self.shade_screen = "ready"

            # 减少补光
            if self.lights != "off":
                self.lights = "low"

            return actions

        def _cloudy_weather_optimization(self):
            """阴天优化策略"""
            actions = []
            actions.append("阴天预报：准备补光和保温")

            # 准备补光系统
            if self.lights == "off":
                self.lights = "ready"

            # 准备保温
            if self.heating == "off":
                self.heating = "ready"

            return actions

        def _rainy_weather_optimization(self):
            """雨天优化策略"""
            actions = []
            actions.append("雨天预报：准备除湿和防涝")

            # 准备除湿
            if self.dehumidifier == "off":
                self.dehumidifier = "ready"

            # 检查排水系统
            if self.drain_valve == "closed":
                self.drain_valve = "checked"

            return actions

        def _stormy_weather_optimization(self):
            """暴风雨天优化策略"""
            actions = []
            actions.append("暴风雨预报：启动安全模式")

            # 收起遮阳网等外部设备
            self.shade_screen = "off"

            # 确保排水畅通
            self.drain_valve = "open"

            # 准备应急电源
            self.emergency_power = "ready"

            return actions

        def _runtime_optimization(self):
            """设备运行时间优化"""
            actions = []

            # 检查设备运行时间，避免长时间连续运行
            for device in ['water_pump', 'co2_generator', 'heating', 'cooling']:
                runtime = self._get_device_runtime(device)
                if runtime > 3600:  # 1小时
                    actions.append(f"{device} 运行时间过长，建议休息")
                    # 可以在这里添加自动休息逻辑

            return actions

        def _load_balancing(self):
            """负载平衡"""
            actions = []

            # 检查同时运行的高能耗设备数量
            high_power_devices = []
            if self.lights in ["high", "medium"]:
                high_power_devices.append("lights")
            if self.heating != "off":
                high_power_devices.append("heating")
            if self.cooling != "off":
                high_power_devices.append("cooling")

            if len(high_power_devices) > 2:
                actions.append("高能耗设备过多，进行负载平衡")
                # 优先保证温度控制
                if "heating" in high_power_devices and "cooling" in high_power_devices:
                    # 不能同时加热和降温，根据优先级处理
                    self.resolve_conflicts()

            return actions

        def _get_device_runtime(self, device_name):
            """获取设备运行时间（简化实现）"""
            # 在实际系统中，这里会记录设备的启动时间
            return 0

        def resolve_conflicts(self):
            """详细解决设备操作冲突"""
            print("正在解决设备冲突...")
            resolved_actions = []

            # 1. 加热和降温冲突（最关键的冲突）
            if self._has_heating_cooling_conflict():
                resolved_actions.extend(self._resolve_heating_cooling_conflict())

            # 2. 通风和CO2补充冲突
            if self._has_ventilation_co2_conflict():
                resolved_actions.extend(self._resolve_ventilation_co2_conflict())

            # 3. 光照和遮阳冲突
            if self._has_light_shade_conflict():
                resolved_actions.extend(self._resolve_light_shade_conflict())

            # 4. 浇水和除湿冲突
            if self._has_watering_dehumidify_conflict():
                resolved_actions.extend(self._resolve_watering_dehumidify_conflict())

            # 5. 多设备优先级冲突
            resolved_actions.extend(self._resolve_priority_conflicts())

            # 6. 能源限制冲突
            resolved_actions.extend(self._resolve_energy_conflicts())

            # 记录冲突解决
            if resolved_actions:
                conflict_record = {
                    'timestamp': time.time(),
                    'action': 'CONFLICT_RESOLUTION',
                    'resolved_actions': resolved_actions,
                    'final_device_states': self._get_device_states()
                }
                self.action_history.append(conflict_record)
                print(f"解决的冲突: {resolved_actions}")

            return resolved_actions

        def _has_heating_cooling_conflict(self):
            """检查加热和降温冲突"""
            return self.heating != "off" and self.cooling != "off"

        def _resolve_heating_cooling_conflict(self):
            """解决加热和降温冲突"""
            actions = []
            current_temp = self.sensor_history[-1]['sensors']['temperature'] if self.sensor_history else 25

            print(f"检测到加热/降温冲突，当前温度: {current_temp}°C")

            if current_temp < self.temp_ideal_low:
                # 温度太低，优先加热
                actions.append("温度过低，优先加热，关闭降温")
                self.cooling = "off"
                if self.heating == "off":
                    self.heating = "medium"

            elif current_temp > self.temp_ideal_high:
                # 温度太高，优先降温
                actions.append("温度过高，优先降温，关闭加热")
                self.heating = "off"
                if self.cooling == "off":
                    self.cooling = "medium"

            else:
                # 温度在理想范围内，根据趋势决定
                trends = self._calculate_trends()
                if trends.get('temperature') == 'rising':
                    actions.append("温度上升趋势，优先降温")
                    self.heating = "off"
                elif trends.get('temperature') == 'falling':
                    actions.append("温度下降趋势，优先加热")
                    self.cooling = "off"
                else:
                    # 稳定状态，都关闭
                    actions.append("温度稳定，关闭加热和降温")
                    self.heating = "off"
                    self.cooling = "off"

            return actions

        def _has_ventilation_co2_conflict(self):
            """检查通风和CO2补充冲突"""
            return self.vent_fan != "off" and self.co2_generator != "off"

        def _resolve_ventilation_co2_conflict(self):
            """解决通风和CO2补充冲突"""
            actions = []
            current_co2 = self.sensor_history[-1]['sensors']['co2'] if self.sensor_history else 1000

            if current_co2 > self.co2_ideal_high * 1.2:
                # CO2严重超标，优先通风
                actions.append("CO2严重超标，优先通风，停止CO2补充")
                self.co2_generator = "off"

            elif current_co2 < self.co2_ideal_low * 0.8:
                # CO2严重不足，优先补充
                actions.append("CO2严重不足，优先补充，减少通风")
                self.vent_fan = "low"

            else:
                # 交错运行
                actions.append("CO2和通风需求平衡，交错运行")
                # 在实际系统中可以设置时间交错

            return actions

        def _has_light_shade_conflict(self):
            """检查光照和遮阳冲突"""
            return self.lights != "off" and self.shade_screen != "off"

        def _resolve_light_shade_conflict(self):
            """解决光照和遮阳冲突"""
            actions = []
            current_light = self.sensor_history[-1]['sensors']['light'] if self.sensor_history else 5000

            if current_light > self.light_ideal_high * 1.3:
                # 光照过强，优先遮阳
                actions.append("光照过强，优先遮阳，关闭补光")
                self.lights = "off"

            elif current_light < self.light_ideal_low * 0.7:
                # 光照严重不足，优先补光
                actions.append("光照严重不足，优先补光，收起遮阳")
                self.shade_screen = "off"

            else:
                # 根据时间段决定
                current_hour = datetime.now().hour
                if 10 <= current_hour <= 16:  # 白天
                    actions.append("白天时段，优先自然光，调整遮阳")
                    self.lights = "off"
                else:  # 早晚
                    actions.append("早晚时段，需要补光，调整遮阳")
                    self.shade_screen = "half"

            return actions

        def _has_watering_dehumidify_conflict(self):
            """检查浇水和除湿冲突"""
            return self.water_pump != "off" and self.dehumidifier != "off"

        def _resolve_watering_dehumidify_conflict(self):
            """解决浇水和除湿冲突"""
            actions = []
            current_moisture = self.sensor_history[-1]['sensors']['moisture'] if self.sensor_history else 60
            current_humidity = self.sensor_history[-1]['sensors']['humidity'] if self.sensor_history else 60

            moisture_priority = False
            humidity_priority = False

            # 判断优先级
            if current_moisture < self.moisture_ideal_low * 0.8:
                moisture_priority = True
            if current_humidity > self.humidity_ideal_high * 1.3:
                humidity_priority = True

            if moisture_priority and not humidity_priority:
                actions.append("土壤严重干燥，优先浇水，暂停除湿")
                self.dehumidifier = "off"
            elif humidity_priority and not moisture_priority:
                actions.append("空气湿度过高，优先除湿，暂停浇水")
                self.water_pump = "off"
            elif moisture_priority and humidity_priority:
                actions.append("土壤干燥但空气潮湿，先浇水后除湿")
                # 可以设置时间序列：先浇水，等待一段时间后再除湿
            else:
                actions.append("湿度和土壤水分需求平衡，协调运行")
                # 降低强度或交错运行

            return actions

        def _resolve_priority_conflicts(self):
            """解决多设备优先级冲突"""
            actions = []

            # 定义设备优先级（数字越小优先级越高）
            priority_order = {
                'heating': 1,  # 温度控制最高优先级
                'cooling': 1,  # 温度控制最高优先级
                'vent_fan': 2,  # 通风次之
                'lights': 3,  # 光照
                'water_pump': 4,  # 浇水
                'co2_generator': 5,  # CO2补充
                'dehumidifier': 6,  # 除湿
                'shade_screen': 7  # 遮阳
            }

            # 检查当前运行的设备
            running_devices = []
            for device, state in self._get_device_states().items():
                if state != "off" and device in priority_order:
                    running_devices.append(device)

            # 如果同时运行的设备太多，关闭低优先级设备
            if len(running_devices) > 3:
                running_devices.sort(key=lambda x: priority_order[x])
                devices_to_stop = running_devices[3:]  # 保留前3个高优先级设备

                for device in devices_to_stop:
                    setattr(self, device, "off")
                    actions.append(f"系统负载高，关闭低优先级设备: {device}")

            return actions

        def _resolve_energy_conflicts(self):
            """解决能源限制冲突"""
            actions = []

            # 计算当前总能耗
            current_power = self._estimate_power_consumption()
            max_power = 5000  # 假设最大功率5kW

            if current_power > max_power:
                actions.append(f"功率超限: {current_power}W > {max_power}W，进行节能调整")

                # 按照优先级降低设备功率
                power_reduction_order = [
                    ('lights', 0.5),  # 灯光减半
                    ('co2_generator', 0),  # 关闭CO2发生器
                    ('dehumidifier', 0),  # 关闭除湿
                    ('water_pump', 0),  # 关闭水泵
                    ('cooling', 0.5),  # 降温减半
                    ('heating', 0.5),  # 加热减半
                ]

                for device, reduction in power_reduction_order:
                    if current_power <= max_power:
                        break

                    current_state = getattr(self, device)
                    if current_state != "off":
                        if reduction == 0:
                            setattr(self, device, "off")
                            actions.append(f"关闭 {device} 以降低能耗")
                        else:
                            # 降低功率级别
                            if current_state == "high":
                                setattr(self, device, "medium")
                            elif current_state == "medium":
                                setattr(self, device, "low")
                            actions.append(f"降低 {device} 功率")

                        current_power = self._estimate_power_consumption()

            return actions

        def _estimate_power_consumption(self):
            """估算当前总功耗"""
            power_estimates = {
                'lights': {'off': 0, 'low': 100, 'medium': 300, 'high': 600},
                'water_pump': {'off': 0, 'low': 50, 'medium': 100, 'high': 200},
                'co2_generator': {'off': 0, 'low': 200, 'medium': 400, 'high': 800},
                'vent_fan': {'off': 0, 'low': 50, 'medium': 100, 'high': 200},
                'heating': {'off': 0, 'low': 500, 'medium': 1000, 'high': 2000},
                'cooling': {'off': 0, 'low': 300, 'medium': 600, 'high': 1200},
                'dehumidifier': {'off': 0, 'on': 400}
            }

            total_power = 0
            for device, states in power_estimates.items():
                current_state = getattr(self, device)
                total_power += states.get(current_state, 0)

            return total_power

        def _get_device_states(self):
            """获取所有设备状态"""
            return {
                'lights': self.lights,
                'water_pump': self.water_pump,
                'co2_generator': self.co2_generator,
                'vent_fan': self.vent_fan,
                'shade_screen': self.shade_screen,
                'heating': self.heating,
                'cooling': self.cooling,
                'dehumidifier': self.dehumidifier,
                'fog_system': self.fog_system,
                'nutrient_pump': self.nutrient_pump,
                'irrigation_system': self.irrigation_system
            }

        def get_sensor_data(self):
            """获取传感器数据（扩展版本）"""
            import random
            light = random.uniform(1000, 10000)
            moisture = random.uniform(30, 90)
            co2 = random.uniform(500, 2000)
            temperature = random.uniform(15, 35)
            humidity = random.uniform(40, 90)
            soil_ph = random.uniform(5.5, 7.5)
            nutrient_level = random.uniform(0, 100)
            leaf_wetness = random.uniform(0, 100)

            return light, moisture, co2, temperature, humidity, soil_ph, nutrient_level, leaf_wetness


    def category1_multivariable_control_complete(x, temp, z):
        y = 60  # 土壤湿度固定为60%
        humidity = 65  # 空气湿度固定为65%

        controller = GreenhouseController()
        actions = []
        triggered = set()

        # ===== 原有105个条件保持不变 =====

        # 分支1-10: 光照与温度协同控制
        if x > 80 and temp > 26 and (x * 0.1 + temp) > 35:
            triggered.add(1)
            actions.append("高光照高温组合: 强力降温遮阳")
            controller.cooling = "high"
            controller.shade_screen = "full"

        if x < 30 and temp < 20 and (x * 0.1 + temp) < 22:
            triggered.add(2)
            actions.append("低光照低温组合: 补光加热")
            controller.lights = "medium"
            controller.heating = "medium"

        if x < 28 and temp < 20 and (x * 0.1 + temp) < 100:
            triggered.add(3)
            actions.append("极低光照低温: 强力补光加热")
            controller.lights = "high"
            controller.heating = "high"

        if (x - 60) * 0.5 > (temp - 24) and z < 40:
            triggered.add(4)
            actions.append("光照偏离度大于温度偏离度且CO2低: 重点补光补CO2")
            controller.lights = "medium"
            controller.co2_generator = "medium"

        if (x - 60) * 0.5 > (temp - 24) and z < 220:
            triggered.add(5)
            actions.append("光照偏离且CO2偏低: 调节光照和CO2")
            controller.lights = "low"
            controller.co2_generator = "low"

        if abs(x - 60) < 15 and abs(temp - 24) < 3 and abs(z - 30) < 10 and humidity > 60:
            triggered.add(6)
            actions.append("接近理想状态且湿度合适: 保持微调")

        if (x * temp) > 2000 and z > 35:
            triggered.add(7)
            actions.append("光照温度乘积高且CO2正常: 降温通风")
            controller.cooling = "medium"
            controller.vent_fan = "medium"

        if (x + temp) < 100 and z < 400:
            triggered.add(8)
            actions.append("光照温度总和低且CO2低: 全面提升")
            controller.lights = "low"
            controller.heating = "low"
            controller.co2_generator = "low"

        if (x + temp) < 100 and y < 580:
            triggered.add(9)
            actions.append("低光照低温: 补光加热")
            controller.lights = "medium"
            controller.heating = "medium"

        if x > 85 and temp > 25 and (x / temp) > 3:
            triggered.add(10)
            actions.append("高光照高温且比值高: 强力降温")
            controller.cooling = "high"
            controller.shade_screen = "half"

        # 分支11-20: 温度控制与复杂组合
        if (x - controller.light_ideal_high) > 5 and (temp - 28) > 1 and z > 30:
            triggered.add(11)
            actions.append("光照温度双超标且CO2正常: 综合调节")
            controller.cooling = "medium"
            controller.shade_screen = "full"
            controller.vent_fan = "medium"

        if (x - controller.light_ideal_high) > 5 and (temp - 28) > 8:
            triggered.add(12)
            actions.append("严重超标: 紧急调节")
            controller.cooling = "high"
            controller.shade_screen = "full"

        if (60 - x) > (24 - temp) * 2 and z < 35:
            triggered.add(13)
            actions.append("缺光问题大于低温且CO2低: 优先补光")
            controller.lights = "medium"
            controller.co2_generator = "low"

        if (x * 0.5) + (temp * 0.3) > 30 and z < 35:
            triggered.add(14)
            actions.append("加权光照温度高且CO2低: 降温补CO2")
            controller.cooling = "low"
            controller.co2_generator = "medium"

        if (x * 0.5) + (temp * 0.3) > 30 and z < 350:
            triggered.add(15)
            actions.append("综合参数偏高: 平衡调节")
            controller.cooling = "low"

        if temp > 25 and z > 40 and (temp + z) > 65:
            triggered.add(16)
            actions.append("高温高CO2组合: 降温通风")
            controller.cooling = "medium"
            controller.vent_fan = "medium"

        if temp < 22 and z < 25 and (temp + z) < 45:
            triggered.add(17)
            actions.append("低温低CO2组合: 加热补CO2")
            controller.heating = "medium"
            controller.co2_generator = "medium"

        if (z - 30) * 0.8 > (temp - 24) and x > 75:
            triggered.add(18)
            actions.append("CO2偏离大于温度偏离且高光照: 重点通风")
            controller.vent_fan = "medium"

        if abs(z - 30) < 10 and abs(temp - 24) < 3 and abs(y - 50) < 15:
            triggered.add(19)
            actions.append("CO2温度接近理想: 保持")

        if (z * temp) > 1000 and x > 80:
            triggered.add(20)
            actions.append("温度CO2乘积高且高光照: 降温通风遮阳")
            controller.cooling = "medium"
            controller.vent_fan = "medium"
            controller.shade_screen = "half"

        # 分支21-30: 光照与CO2协同控制
        if x < 40 and z < 25 and (x + z) < 60:
            triggered.add(21)
            actions.append("光照CO2双低: 补光补CO2")
            controller.lights = "medium"
            controller.co2_generator = "medium"

        if (humidity + temp) < 85 and y < 530:
            triggered.add(22)
            actions.append("环境参数偏低: 提升环境")

        if x > 80 and z > 40 and (x / z) > 2:
            triggered.add(23)
            actions.append("光照CO2比值高: 补CO2")
            controller.co2_generator = "high"

        if z > 45 and x > 75 and (z - 40) > 3:
            triggered.add(24)
            actions.append("CO2光照双超标: 通风遮阳")
            controller.vent_fan = "medium"
            controller.shade_screen = "half"

        if z < 25 and x < 40 and (30 - z) > (60 - x) * 0.5:
            triggered.add(25)
            actions.append("CO2缺失大于光照缺失: 优先补CO2")
            controller.co2_generator = "medium"

        if (60 - humidity) > (24 - temp) * 1.5 and x < 308:
            triggered.add(26)
            actions.append("低光照: 补光")
            controller.lights = "medium"

        if (x * 0.3) + (z * 0.4) > 35 and temp > 25:
            triggered.add(27)
            actions.append("加权光照CO2高且高温: 降温通风")
            controller.cooling = "low"
            controller.vent_fan = "medium"

        if x > 80 and temp > 26 and z > 40 and (x * 0.2 + temp * 0.3 + z * 0.1) > 30:
            triggered.add(28)
            actions.append("三变量高值: 全面降温通风")
            controller.cooling = "high"
            controller.vent_fan = "high"
            controller.shade_screen = "full"

        if x < 30 and temp < 20 and z < 25 and (x + temp + z) < 70:
            triggered.add(29)
            actions.append("三变量低值: 全面补充")
            controller.lights = "high"
            controller.heating = "high"
            controller.co2_generator = "high"

        if z < 18 and x < 30 and (z * 0.2 + x * 0.1) < 75:
            triggered.add(30)
            actions.append("极低环境: 紧急补充")
            controller.lights = "high"
            controller.co2_generator = "high"

        # 分支31-40: 三变量复杂控制
        if (z - 30) * 0.5 > (x - 60) * 0.1 and temp > 26:
            triggered.add(31)
            actions.append("CO2偏离主导且高温: 重点通风降温")
            controller.vent_fan = "medium"
            controller.cooling = "low"

        if abs(z - 30) < 10 and abs(x - 60) < 15 and abs(temp - 24) < 2:
            triggered.add(32)
            actions.append("三变量理想: 保持稳定")

        if abs(z - 30) < 10 and abs(x - 60) < 15 and abs(temp - 24) < 108:
            triggered.add(33)
            actions.append("参数良好: 维持")

        if (z * x) > 3000 and temp > 25:
            triggered.add(34)
            actions.append("CO2光照乘积高且高温: 通风降温")
            controller.vent_fan = "medium"
            controller.cooling = "medium"

        if x < 35 and temp < 22 and z < 28:
            triggered.add(35)
            actions.append("条件35: 低值环境调节")
            controller.lights = "medium"
            controller.heating = "medium"

        if x > 75 and temp > 24 and z > 35:
            triggered.add(36)
            actions.append("条件36: 高值环境调节")
            controller.vent_fan = "low"

        if (x + temp + z) > 150:
            triggered.add(37)
            actions.append("条件37: 总和过高")
            controller.cooling = "low"

        if (x * temp * z) > 50000:
            triggered.add(38)
            actions.append("条件38: 乘积过高")
            controller.vent_fan = "medium"

        if abs(x - 60) + abs(temp - 24) + abs(z - 30) > 40:
            triggered.add(39)
            actions.append("条件39: 偏离度总和高")

        if x / (temp + 1) > 3:
            triggered.add(40)
            actions.append("条件40: 光照温度比高")
            controller.shade_screen = "half"

        # 分支41-50: 比值与关系控制
        if z / (x + 1) > 0.8:
            triggered.add(41)
            actions.append("条件41: CO2光照比高")
            controller.vent_fan = "low"

        if temp / (z + 1) > 0.8:
            triggered.add(42)
            actions.append("条件42: 温度CO2比高")
            controller.cooling = "low"

        if (x - 60) * (temp - 24) > 100:
            triggered.add(43)
            actions.append("条件43: 偏离乘积高")

        if (z - 30) * (x - 60) > 200:
            triggered.add(44)
            actions.append("条件44: CO2光照偏离乘积高")

        if temp > 28 and x > 70:
            triggered.add(45)
            actions.append("高温高光照: 降温遮阳")
            controller.cooling = "medium"
            controller.shade_screen = "half"

        if temp < 18 and z < 30:
            triggered.add(46)
            actions.append("低温低CO2: 加热补CO2")
            controller.heating = "medium"
            controller.co2_generator = "medium"

        if x > 85 and z > 35:
            triggered.add(47)
            actions.append("高光照高CO2: 遮阳通风")
            controller.shade_screen = "half"
            controller.vent_fan = "medium"

        if x < 25 and temp < 22:
            triggered.add(48)
            actions.append("低光照低温: 补光加热")
            controller.lights = "medium"
            controller.heating = "medium"

        if z > 50 and temp > 26:
            triggered.add(49)
            actions.append("高CO2高温: 通风降温")
            controller.vent_fan = "medium"
            controller.cooling = "low"

        if z < 20 and x < 35:
            triggered.add(50)
            actions.append("低CO2低光照: 补CO2补光")
            controller.co2_generator = "medium"
            controller.lights = "medium"

        # 分支51-60: 组合条件控制
        if x > 75 and temp > 25 and z > 38:
            triggered.add(51)
            actions.append("三高状态: 综合调节")
            controller.cooling = "medium"
            controller.vent_fan = "medium"
            controller.shade_screen = "half"

        if x < 35 and temp < 21 and z < 28:
            triggered.add(52)
            actions.append("三低状态: 全面提升")
            controller.lights = "medium"
            controller.heating = "medium"
            controller.co2_generator = "medium"

        if x > 0 and temp > 0 and (x / temp) > 3.5:
            triggered.add(53)
            actions.append("光照温度比过高: 降温")
            controller.cooling = "low"

        if z > 0 and x > 0 and (z / x) > 0.7:
            triggered.add(54)
            actions.append("CO2光照比过高: 补光")
            controller.lights = "low"

        if (x - temp) > 50:
            triggered.add(55)
            actions.append("光照温度差过大: 平衡调节")
            controller.shade_screen = "half"
            controller.cooling = "low"

        if (z - temp) > 10:
            triggered.add(56)
            actions.append("CO2温度差过大: 调节")
            controller.vent_fan = "low"

        if (x + temp) > 110:
            triggered.add(57)
            actions.append("光照温度和过高: 降温遮阳")
            controller.cooling = "medium"
            controller.shade_screen = "half"

        if (z + x) > 120:
            triggered.add(58)
            actions.append("CO2光照和过高: 通风遮阳")
            controller.vent_fan = "medium"
            controller.shade_screen = "half"

        if (x * 0.4 + temp * 0.3 + z * 0.1) > 32:
            triggered.add(59)
            actions.append("加权和过高: 综合降低")
            controller.cooling = "low"
            controller.vent_fan = "low"

        if (x - 60) ** 2 + (temp - 24) ** 2 > 500:
            triggered.add(60)
            actions.append("欧式距离过大: 调节")

        # 分支61-70: 复杂表达式控制
        if x * temp * z > 60000:
            triggered.add(61)
            actions.append("三变量乘积过高: 全面调节")
            controller.cooling = "medium"
            controller.vent_fan = "medium"
            controller.shade_screen = "half"

        if abs(x - 60) + abs(temp - 24) + abs(z - 30) > 35:
            triggered.add(62)
            actions.append("曼哈顿距离过大: 调节")

        if (x > 70 and temp < 20) or (x < 30 and temp > 28):
            triggered.add(63)
            actions.append("光照温度不匹配: 平衡")
            if x > 70:
                controller.heating = "medium"
            else:
                controller.cooling = "medium"

        if (z > 40 and x < 35) or (z < 25 and x > 75):
            triggered.add(64)
            actions.append("CO2光照不匹配: 平衡")
            if z > 40:
                controller.lights = "medium"
            else:
                controller.co2_generator = "medium"

        if x >= 95:
            triggered.add(65)
            actions.append("光照接近上限: 强力遮阳")
            controller.shade_screen = "full"

        if x <= 5:
            triggered.add(66)
            actions.append("光照接近下限: 强力补光")
            controller.lights = "high"

        if temp >= 38:
            triggered.add(67)
            actions.append("温度接近上限: 紧急降温")
            controller.cooling = "high"

        if temp <= 12:
            triggered.add(68)
            actions.append("温度接近下限: 紧急加热")
            controller.heating = "high"

        if z >= 55:
            triggered.add(69)
            actions.append("CO2接近上限: 强力通风")
            controller.vent_fan = "high"

        if z <= 15:
            triggered.add(70)
            actions.append("CO2接近下限: 强力补充")
            controller.co2_generator = "high"

        # 分支71-80: 边界组合控制
        if x >= 90 and temp >= 35:
            triggered.add(71)
            actions.append("光照温度双接近上限: 紧急处理")
            controller.cooling = "high"
            controller.shade_screen = "full"

        if x <= 10 and temp <= 15:
            triggered.add(72)
            actions.append("光照温度双接近下限: 紧急处理")
            controller.lights = "high"
            controller.heating = "high"

        if z >= 52 and x >= 88:
            triggered.add(73)
            actions.append("CO2光照双高: 紧急通风遮阳")
            controller.vent_fan = "high"
            controller.shade_screen = "full"

        if z <= 18 and x <= 12:
            triggered.add(74)
            actions.append("CO2光照双低: 紧急补充")
            controller.co2_generator = "high"
            controller.lights = "high"

        if x > 0 and temp > 0 and (x / temp) >= 4:
            triggered.add(75)
            actions.append("光照温度比极端: 快速降温")
            controller.cooling = "high"

        if temp > 0 and x > 0 and (temp / x) >= 1:
            triggered.add(76)
            actions.append("温度光照比极端: 快速补光")
            controller.lights = "high"

        if z > 0 and temp > 0 and (z / temp) >= 2:
            triggered.add(77)
            actions.append("CO2温度比极端: 快速通风")
            controller.vent_fan = "high"

        if (x - temp) >= 60:
            triggered.add(78)
            actions.append("光照温度差极端: 紧急平衡")
            controller.shade_screen = "full"
            controller.cooling = "high"

        if (temp - x) >= 20:
            triggered.add(79)
            actions.append("温度光照差极端: 紧急平衡")
            controller.lights = "high"
            controller.cooling = "high"

        if (z - x) >= 30:
            triggered.add(80)
            actions.append("CO2光照差极端: 调节")
            controller.vent_fan = "high"

        # 分支81-90: 极端情况控制
        if x >= 90 and temp >= 35 and z >= 50:
            triggered.add(81)
            actions.append("三变量全部接近上限: 全面紧急降低")
            controller.cooling = "high"
            controller.vent_fan = "high"
            controller.shade_screen = "full"

        if x <= 10 and temp <= 15 and z <= 20:
            triggered.add(82)
            actions.append("三变量全部接近下限: 全面紧急提升")
            controller.lights = "high"
            controller.heating = "high"
            controller.co2_generator = "high"

        if (x > 85 and temp < 18) or (x < 15 and temp > 35):
            triggered.add(83)
            actions.append("光照温度极端不平衡: 特殊处理")
            if x > 85:
                controller.shade_screen = "full"
                controller.heating = "high"
            else:
                controller.lights = "high"
                controller.cooling = "high"

        if (z > 50 and x < 15) or (z < 18 and x > 88):
            triggered.add(84)
            actions.append("CO2光照极端不平衡: 特殊处理")
            if z > 50:
                controller.vent_fan = "high"
                controller.lights = "high"
            else:
                controller.co2_generator = "high"
                controller.shade_screen = "full"

        # 分支85-105: 精细化控制
        if x > 65 and temp > 27 and z > 42 and (x + temp + z) > 135:
            triggered.add(85)
            actions.append("三变量中高组合: 适度调节")
            controller.cooling = "low"
            controller.vent_fan = "low"

        if x < 45 and temp < 23 and z < 32 and (x + temp + z) < 95:
            triggered.add(86)
            actions.append("三变量中低组合: 适度提升")
            controller.lights = "low"
            controller.heating = "low"

        if (x / (temp + 1)) * (z / (x + 1)) > 1.5:
            triggered.add(87)
            actions.append("复合比值过高: 平衡调节")
            controller.shade_screen = "half"
            controller.vent_fan = "low"

        if 40 <= x <= 80 and 22 <= temp <= 26 and 28 <= z <= 35:
            triggered.add(88)
            actions.append("三变量均在理想区间: 精细维持")

        if abs(x - 60) * abs(temp - 24) * abs(z - 30) > 500:
            triggered.add(89)
            actions.append("多变量波动过大: 稳定控制")
            controller.vent_fan = "low"
            controller.cooling = "low"

        if x > 60 and temp > 24 and z > 30 and (x - 60) + (temp - 24) + (z - 30) > 20:
            triggered.add(90)
            actions.append("轻度超标组合: 温和调节")
            controller.vent_fan = "low"

        if x < 60 and temp < 24 and z < 30 and (60 - x) + (24 - temp) + (30 - z) > 20:
            triggered.add(91)
            actions.append("轻度不足组合: 温和补充")
            controller.lights = "low"
            controller.co2_generator = "low"

        if (x * 0.3 + temp * 0.5 + z * 0.2) > 35:
            triggered.add(92)
            actions.append("新加权值过高: 调节")
            controller.cooling = "low"

        if (x * 0.3 + temp * 0.5 + z * 0.2) < 20:
            triggered.add(93)
            actions.append("新加权值过低: 补充")
            controller.lights = "low"

        if x > 70 and z > 35 and (x - z) > 30:
            triggered.add(94)
            actions.append("光照显著高于CO2: 补CO2遮阳")
            controller.co2_generator = "medium"
            controller.shade_screen = "half"

        if z > 40 and x < 50 and (z - x) > 10:
            triggered.add(95)
            actions.append("CO2显著高于光照: 补光通风")
            controller.lights = "medium"
            controller.vent_fan = "low"

        if temp > 26 and (x + z) > 120:
            triggered.add(96)
            actions.append("高温且光照CO2总和高: 全面降低")
            controller.cooling = "medium"
            controller.vent_fan = "medium"

        if temp < 22 and (x + z) < 80:
            triggered.add(97)
            actions.append("低温且光照CO2总和低: 全面提升")
            controller.heating = "medium"
            controller.lights = "low"

        if abs(x - temp) < 10 and abs(temp - z) < 10:
            triggered.add(98)
            actions.append("三变量数值接近: 协调控制")

        if max(x, temp, z) - min(x, temp, z) > 60:
            triggered.add(99)
            actions.append("变量间差异极大: 平衡调节")
            controller.vent_fan = "low"

        if (x > 80 or temp > 30 or z > 45) and not (x > 80 and temp > 30 and z > 45):
            triggered.add(100)
            actions.append("单变量过高其他正常: 针对性调节")
            if x > 80:
                controller.shade_screen = "half"
            elif temp > 30:
                controller.cooling = "low"
            else:
                controller.vent_fan = "low"

        if (x < 30 or temp < 20 or z < 25) and not (x < 30 and temp < 20 and z < 25):
            triggered.add(101)
            actions.append("单变量过低其他正常: 针对性补充")
            if x < 30:
                controller.lights = "low"
            elif temp < 20:
                controller.heating = "low"
            else:
                controller.co2_generator = "low"

        if x + temp > 120 and z < 30:
            triggered.add(102)
            actions.append("光照温度高CO2低: 降温补CO2")
            controller.cooling = "low"
            controller.co2_generator = "medium"

        if x + z > 130 and temp < 20:
            triggered.add(103)
            actions.append("光照CO2高温度低: 加热通风")
            controller.heating = "medium"
            controller.vent_fan = "low"

        if temp + z > 70 and x < 40:
            triggered.add(104)
            actions.append("温度CO2高光照低: 补光降温")
            controller.lights = "medium"
            controller.cooling = "low"

        if (x - 50) ** 2 + (temp - 25) ** 2 + (z - 35) ** 2 > 1000:
            triggered.add(105)
            actions.append("距理想中心点距离过大: 全面调节")
            controller.vent_fan = "low"

        # ===== 新增30个条件：补充源代码缺失的条件类型 =====

        # 分支106-115: 系数比较条件（基于理想值的倍数比较）
        # 106. 光照超过理想值的1.2倍且temp高
        if x > controller.light_ideal_high * 1.2 and temp > 26:
            triggered.add(106)
            actions.append("光照显著超标且高温: 强力遮阳降温")
            controller.shade_screen = "full"
            controller.cooling = "high"

        # 107. 光照低于理想值的0.8倍且temp低
        if x < controller.light_ideal_low * 0.8 and temp < 22:
            triggered.add(107)
            actions.append("光照显著不足且低温: 强力补光加热")
            controller.lights = "high"
            controller.heating = "high"

        # 108. CO2超过理想值的1.3倍
        if z > controller.co2_ideal_high * 1.3:
            triggered.add(108)
            actions.append("CO2严重超标: 紧急通风")
            controller.vent_fan = "high"
            controller.co2_generator = "off"

        # 109. CO2低于理想值的0.7倍
        if z < controller.co2_ideal_low * 0.7:
            triggered.add(109)
            actions.append("CO2严重不足: 紧急补充")
            controller.co2_generator = "high"

        # 110. 温度超过理想值的1.1倍且光照高
        if temp > 24 * 1.1 and x > controller.light_ideal_high:
            triggered.add(110)
            actions.append("温度轻度超标且光照高: 协调降温")
            controller.cooling = "medium"
            controller.shade_screen = "half"

        # 111. 光照超过理想值的1.1倍且CO2低
        if x > controller.light_ideal_high * 1.1 and z < controller.co2_ideal_low * 0.8:
            triggered.add(111)
            actions.append("光照轻度超标且CO2严重不足: 遮阳补CO2")
            controller.shade_screen = "half"
            controller.co2_generator = "high"

        # 112. 三个参数都超过理想值的1.2倍
        if x > controller.light_ideal_high * 1.2 and temp > 24 * 1.2 and z > controller.co2_ideal_high * 1.2:
            triggered.add(112)
            actions.append("三参数严重超标: 全面降低")
            controller.shade_screen = "full"
            controller.cooling = "high"
            controller.vent_fan = "high"

        # 113. 光照与CO2的比值系数条件
        if x > controller.light_ideal_high * 0.9 and z > controller.co2_ideal_high * 0.9:
            triggered.add(113)
            actions.append("光照CO2接近上限: 预防性调节")
            controller.shade_screen = "half"
            controller.vent_fan = "low"

        # 114. 温度系数与其他参数组合
        if temp > 24 * 1.15 and (x + z) > 100:
            triggered.add(114)
            actions.append("温度系数超标且光照CO2总和高: 综合降温")
            controller.cooling = "high"
            controller.vent_fan = "medium"

        # 115. 低值系数组合
        if x < controller.light_ideal_low * 0.9 and z < controller.co2_ideal_low * 0.9:
            triggered.add(115)
            actions.append("光照CO2双接近下限: 预防性补充")
            controller.lights = "low"
            controller.co2_generator = "low"

        # 分支116-125: 嵌套if条件结构
        # 116. 光照高时的嵌套温度判断
        if x > 80:
            if temp > 28:
                triggered.add(116)
                actions.append("高光照嵌套高温: 强力降温遮阳")
                controller.cooling = "high"
                controller.shade_screen = "full"

        # 117. 温度低时的嵌套光照判断
        if temp < 20:
            if x < 35 and z < 25:
                triggered.add(117)
                actions.append("低温嵌套低光照低CO2: 全面提升")
                controller.heating = "high"
                controller.lights = "high"
                controller.co2_generator = "high"

        # 118. CO2高时的嵌套判断
        if z > 45:
            if temp > 26 or x > 85:
                triggered.add(118)
                actions.append("高CO2嵌套高温或高光照: 通风降温")
                controller.vent_fan = "high"
                if temp > 26:
                    controller.cooling = "medium"
                if x > 85:
                    controller.shade_screen = "half"

        # 119. 三层嵌套条件
        if x > 75:
            if temp > 25:
                if z > 40:
                    triggered.add(119)
                    actions.append("三层嵌套高值: 协调降低")
                    controller.shade_screen = "half"
                    controller.cooling = "low"
                    controller.vent_fan = "low"

        # 120. 复合嵌套条件
        if temp < 22:
            if x < 40 or z < 28:
                triggered.add(120)
                actions.append("低温嵌套低光照或低CO2: 针对性提升")
                controller.heating = "medium"
                if x < 40:
                    controller.lights = "medium"
                if z < 28:
                    controller.co2_generator = "medium"

        # 121. 范围嵌套条件
        if 40 < x < 80:
            if temp > 30 and z > 45:
                triggered.add(121)
                actions.append("中等光照嵌套高温高CO2: 通风降温")
                controller.vent_fan = "medium"
                controller.cooling = "medium"

        # 122. 数学运算嵌套
        if (x + temp) > 100:
            if z < 30 and (x - temp) > 20:
                triggered.add(122)
                actions.append("高总和嵌套低CO2大差值: 复合调节")
                controller.co2_generator = "medium"
                controller.cooling = "low"

        # 123. 比值嵌套条件
        if x / (temp + 1) > 2.5:
            if z > 40 and temp < 25:
                triggered.add(123)
                actions.append("高比值嵌套高CO2低温: 平衡调节")
                controller.heating = "low"
                controller.vent_fan = "low"

        # 124. 极值嵌套条件
        if max(x, temp, z) > 85:
            if min(x, temp, z) < 20:
                triggered.add(124)
                actions.append("极值嵌套条件: 均衡调节")
                controller.vent_fan = "medium"
                controller.lights = "low"
                controller.heating = "low"

        # 125. 复杂嵌套逻辑
        if abs(x - 60) > 20:
            if abs(temp - 24) < 5 and abs(z - 35) < 8:
                triggered.add(125)
                actions.append("光照偏离嵌套温度CO2稳定: 重点调光照")
                if x > 60:
                    controller.shade_screen = "medium"
                else:
                    controller.lights = "medium"

        # 分支126-130: 否定条件（not逻辑）
        # 126. 非高CO2条件
        if not (z > controller.co2_ideal_high) and x > 80 and temp > 26:
            triggered.add(126)
            actions.append("非高CO2且高光照高温: 可以安全降温")
            controller.cooling = "medium"
            controller.shade_screen = "half"

        # 127. 非低温条件
        if not (temp < 20) and x < 35 and z < 30:
            triggered.add(127)
            actions.append("非低温且低光照低CO2: 补光补CO2")
            controller.lights = "medium"
            controller.co2_generator = "medium"

        # 128. 复合否定条件
        if not (x > 85 and temp > 30) and z > 45:
            triggered.add(128)
            actions.append("非极端高光照高温且高CO2: 重点通风")
            controller.vent_fan = "medium"

        # 129. 多重否定条件
        if not (z < controller.co2_ideal_low) and not (temp > 32) and x < 40:
            triggered.add(129)
            actions.append("非低CO2非高温且低光照: 安全补光")
            controller.lights = "medium"

        # 130. 范围否定条件
        if not (25 <= temp <= 28) and not (35 <= z <= 42) and x > 70:
            triggered.add(130)
            actions.append("温度CO2非理想范围且高光照: 综合调节")
            controller.shade_screen = "half"
            if temp > 28:
                controller.cooling = "low"
            elif temp < 25:
                controller.heating = "low"

        # 分支131-135: 基于y变量的特定条件（湿度模拟）
        # 131. 模拟湿度过高的情况
        if y > controller.moisture_ideal_high and x > 75:
            triggered.add(131)
            actions.append("模拟高湿度且高光照: 通风除湿")
            controller.vent_fan = "medium"
            controller.dehumidifier = "on"

        # 132. 模拟湿度过低的情况
        if y < controller.moisture_ideal_low and temp > 26:
            triggered.add(132)
            actions.append("模拟低湿度且高温: 降温补湿")
            controller.cooling = "low"
            controller.water_pump = "medium"

        # 133. 湿度与光照CO2的复合条件
        if y > controller.moisture_ideal_high * 1.1 and z < controller.co2_ideal_low and x > 70:
            triggered.add(133)
            actions.append("模拟高湿度低CO2高光照: 通风补CO2")
            controller.vent_fan = "high"
            controller.co2_generator = "medium"
            controller.dehumidifier = "on"

        # 134. 湿度差值条件
        if (y - controller.moisture_ideal_high) > 10 and (x - controller.light_ideal_high) > 15:
            triggered.add(134)
            actions.append("湿度光照双超标: 除湿遮阳")
            controller.dehumidifier = "high"
            controller.shade_screen = "half"
            controller.vent_fan = "medium"

        # 135. 湿度比值条件
        if y / x > 0.8 and z < 30 and temp < 25:
            triggered.add(135)
            actions.append("湿度光照比高且低CO2低温: 综合调节")
            controller.water_pump = "off"
            controller.co2_generator = "medium"
            controller.heating = "low"

        return actions, triggered

    def section2_high_light_low_moisture_enhanced(self, x, y, z):
        """部分2: 高光照、低湿度组合 - 整合原验证规则的扩展版本"""
        actions = []
        triggered = set()

        # 固定环境参数（对应第一个代码中的环境变量）
        energy_price = 0.15  # 能源价格
        time_of_day = 12  # 时间段（中午）
        temp = 25  # 温度
        humidity = 60  # 湿度
        energy_trend = "stable"  # 能源趋势
        weather_forecast = "sunny"  # 天气预报

        # ===== 原有的11个基础条件保持不变 =====

        # 1. 简单双变量条件
        if x > self.light_ideal_high and y < self.moisture_ideal_low:
            triggered.add(1)
            actions.append("高光低湿：补光关闭，启动浇水")
            self.lights = "off"
            self.water_pump = "medium"

        # 2. 三变量条件
        if x > self.light_ideal_high * 1.1 and y < self.moisture_ideal_low * 0.8:
            triggered.add(2)
            actions.append("极强光、极干旱：紧急补水")
            self.water_pump = "high"
            self.shade_screen = "half"
            self.vent_fan = "medium"

        # 3. 数学运算组合条件 (差值比较)
        if (self.moisture_ideal_low - y) > 20 and x > 8000:
            triggered.add(3)
            actions.append("严重干旱且高光：强力补水")
            self.water_pump = "high"
            self.shade_screen = "on"

        # 4. 复合条件
        if (x > 9000 or z > 1600) and y < 45:
            triggered.add(4)
            actions.append("极端光照或高CO2且干旱：综合补水")
            self.water_pump = "high"
            self.shade_screen = "on"
            self.vent_fan = "high"

        # 5. 数学运算组合条件 (比率比较)
        if x / (y + 1) > 200 and z < 900:
            triggered.add(5)
            actions.append("光照/湿度比过高且CO2不足：补水补CO2")
            self.water_pump = "medium"
            self.co2_generator = "medium"

        # 6. 三变量复合条件
        if x > 8500 and y < 50 and not (z < self.co2_ideal_low):
            triggered.add(6)
            actions.append("高光低湿且CO2正常：重点补水")
            self.water_pump = "medium"
            self.shade_screen = "half"

        # 7. 数学运算组合条件 (乘积比较)
        if x * y < 350000 and z > 1600:
            triggered.add(7)
            actions.append("光照湿度乘积低且CO2高：补水通风")
            self.water_pump = "medium"
            self.vent_fan = "medium"

        # 8. 嵌套条件逻辑
        if y < 55:
            if x > 9000 and z < 700:
                triggered.add(8)
                actions.append("干旱、强光且CO2不足：多重调节")
                self.water_pump = "high"
                self.shade_screen = "half"
                self.co2_generator = "high"

        # 9. 双变量范围条件
        if (x > 8000 and x < 9000) and (y > 45 and y < 55):
            triggered.add(9)
            actions.append("光照中度超标且湿度略低：轻度调节")
            self.water_pump = "low"
            self.shade_screen = "half"

        # 10. 多变量数学组合条件
        if (x - 7000) > (60 - y) * 100:
            triggered.add(10)
            actions.append("光照超标量大于湿度不足量：综合调节")
            self.shade_screen = "on"
            self.water_pump = "medium"
            self.vent_fan = "low"

        # 11. 复合逻辑条件
        if x > 8500 and (y < 50 or z > 1600):
            triggered.add(11)
            actions.append("高光且干旱或高CO2：针对性调节")
            if y < 50:
                self.water_pump = "medium"
            if z > 1600:
                self.vent_fan = "medium"

        # ===== 新增：从第一个代码提取的原语句条件（适配温室控制场景）=====

        # 12-21: 基于能源价格的基础优化（适配为基于光照强度的优化）
        if energy_price > 0.10 and time_of_day == 12 and (x + temp) > 500:
            triggered.add(12)
            actions.append("中午高光照高温组合1：启动遮阳降温")
            self.shade_screen = "half"
            self.cooling = "low"

        if energy_price > 0.10 and time_of_day == 12 and (x + temp) > 1000:
            triggered.add(13)
            actions.append("中午极高光照高温组合：强力遮阳降温")
            self.shade_screen = "full"
            self.cooling = "high"

        if energy_price < 0.20 and (x + z) < 800:
            triggered.add(14)
            actions.append("光照CO2总和偏低：补光补CO2")
            self.lights = "low"
            self.co2_generator = "low"

        if energy_price < 0.20 and (x + z) < 600:
            triggered.add(15)
            actions.append("光照CO2总和极低：强力补光补CO2")
            self.lights = "medium"
            self.co2_generator = "medium"

        if energy_price > 0.10 and (x * temp) > 20000:
            triggered.add(16)
            actions.append("光照温度乘积高：降温遮阳")
            self.cooling = "medium"
            self.shade_screen = "half"

        if energy_price > 0.10 and (x * temp) > 15000:
            triggered.add(17)
            actions.append("光照温度乘积中等偏高：轻度降温")
            self.cooling = "low"
            self.shade_screen = "half"

        if energy_price < 0.20 and (y + humidity) < 100:
            triggered.add(18)
            actions.append("湿度总和偏低：增湿浇水")
            self.water_pump = "medium"

        if energy_price < 0.20 and (y + humidity) < 80:
            triggered.add(19)
            actions.append("湿度总和极低：强力增湿")
            self.water_pump = "high"

        if energy_price > 0.10 and z > 300:
            triggered.add(20)
            actions.append("CO2偏高：通风调节")
            self.vent_fan = "low"

        if energy_price > 0.10 and z > 200:
            triggered.add(21)
            actions.append("CO2中等偏高：轻度通风")
            self.vent_fan = "low"

        # 22-31: 基于时间段的优化（中午时段特殊处理）
        if time_of_day == 12 and x < 500 and y > 200:
            triggered.add(22)
            actions.append("中午低光照高湿度：补光通风")
            self.lights = "medium"
            self.vent_fan = "low"

        if time_of_day == 12 and x < 400 and z < 300:
            triggered.add(23)
            actions.append("中午低光照低CO2：补光补CO2")
            self.lights = "medium"
            self.co2_generator = "low"

        if temp + humidity > 20 and y < 500:
            triggered.add(24)
            actions.append("温湿度总和高但土壤湿度低：浇水")
            self.water_pump = "low"

        if energy_price < 0.20 and (z + x) < 800 and y > 100:
            triggered.add(25)
            actions.append("光照CO2总和低但土壤湿度可以：补光补CO2")
            self.lights = "low"
            self.co2_generator = "low"

        if energy_price < 0.20 and (z + x) < 600 and y < 800:
            triggered.add(26)
            actions.append("光照CO2低且土壤湿度低：全面补充")
            self.lights = "medium"
            self.co2_generator = "medium"
            self.water_pump = "low"

        if temp < 30 and x > 500 and z > 200:
            triggered.add(27)
            actions.append("温度适宜光照CO2较好：维持")
            pass

        if temp < 28 and x > 400 and y > 300:
            triggered.add(28)
            actions.append("温度适宜光照土壤湿度好：维持")
            pass

        if time_of_day == 12 and x > 600 and z < 400:
            triggered.add(29)
            actions.append("中午高光照低CO2：遮阳补CO2")
            self.shade_screen = "half"
            self.co2_generator = "low"

        if time_of_day == 12 and x > 500 and y + z > 600:
            triggered.add(30)
            actions.append("中午高光照且湿度CO2总和高：遮阳通风")
            self.shade_screen = "half"
            self.vent_fan = "low"

        # 32-41: 复杂组合优化条件
        if time_of_day == 12 and z < 300 and x > 200:
            triggered.add(31)
            actions.append("中午低CO2但光照可以：补CO2")
            self.co2_generator = "low"

        if time_of_day == 12 and z < 250 and y > 300:
            triggered.add(32)
            actions.append("中午低CO2但土壤湿度好：补CO2")
            self.co2_generator = "medium"

        if energy_price > 0.10 and x < 400 and y + z > 500:
            triggered.add(33)
            actions.append("低光照但湿度CO2总和高：补光通风")
            self.lights = "medium"
            self.vent_fan = "low"

        if energy_price > 0.10 and x < 300 and z > 200:
            triggered.add(34)
            actions.append("极低光照但CO2可以：补光")
            self.lights = "high"

        if time_of_day == 12 and temp > 20 and x > 400:
            triggered.add(35)
            actions.append("中午温度适宜光照好：维持")
            pass

        if time_of_day == 12 and temp > 18 and y < 600:
            triggered.add(36)
            actions.append("中午温度适宜但土壤湿度低：浇水")
            self.water_pump = "medium"

        if humidity > 50 and y > 500 and x < 800:
            triggered.add(37)
            actions.append("空气湿度高土壤湿度高但光照不足：补光除湿")
            self.lights = "medium"
            self.dehumidifier = "on"

        if humidity > 45 and y > 450 and z < 400:
            triggered.add(38)
            actions.append("湿度较高但CO2不足：补CO2除湿")
            self.co2_generator = "low"
            self.dehumidifier = "on"

        if time_of_day == 12 and y < 200 and x + z > 500:
            triggered.add(39)
            actions.append("中午土壤很干但光照CO2总和可以：重点浇水")
            self.water_pump = "high"

        if time_of_day == 12 and y < 150 and z > 100:
            triggered.add(40)
            actions.append("中午土壤极干但CO2可以：紧急浇水")
            self.water_pump = "high"

        # 42-51: 多变量组合优化
        if energy_price > 0.10 and (x * 0.15) > 80 and temp > 20:
            triggered.add(41)
            actions.append("光照系数高且温度适宜：轻度遮阳")
            self.shade_screen = "half"

        if energy_price > 0.10 and (x * 0.15) > 60 and temp > 20:
            triggered.add(42)
            actions.append("光照系数中等且温度适宜：准备遮阳")
            self.shade_screen = "ready"

        if energy_price < 0.20 and (z * 0.5) > 100 and x < 600:
            triggered.add(43)
            actions.append("CO2系数高但光照不足：补光")
            self.lights = "medium"

        if energy_price < 0.20 and (z * 0.5) > 80 and x < 600:
            triggered.add(44)
            actions.append("CO2系数中等但光照不足：轻度补光")
            self.lights = "low"

        if temp + humidity > 80 and y > 400:
            triggered.add(45)
            actions.append("温度湿度总和高且土壤湿度高：除湿通风")
            self.dehumidifier = "on"
            self.vent_fan = "medium"

        if temp + humidity > 70 and y > 350:
            triggered.add(46)
            actions.append("温度湿度总和中等偏高且土壤湿度高：轻度除湿")
            self.dehumidifier = "on"

        if energy_price < 0.20 and (y + z) < 400:
            triggered.add(47)
            actions.append("土壤湿度CO2总和低：补水补CO2")
            self.water_pump = "medium"
            self.co2_generator = "low"

        if energy_price < 0.20 and (y + z) < 350:
            triggered.add(48)
            actions.append("土壤湿度CO2总和极低：强力补水补CO2")
            self.water_pump = "high"
            self.co2_generator = "medium"

        if time_of_day == 12 and (x * temp) > 15000:
            triggered.add(49)
            actions.append("中午光照温度乘积高：降温遮阳")
            self.cooling = "medium"
            self.shade_screen = "half"

        if time_of_day == 12 and (x * temp) > 12000:
            triggered.add(50)
            actions.append("中午光照温度乘积中等偏高：轻度降温")
            self.cooling = "low"

        # 52-61: 基于能源趋势和天气的优化
        if energy_trend == "stable" and x < 500 and y > 200:
            triggered.add(51)
            actions.append("稳定趋势低光照高湿度：补光")
            self.lights = "medium"

        if energy_trend == "stable" and x < 400 and z < 300:
            triggered.add(52)
            actions.append("稳定趋势低光照低CO2：补光补CO2")
            self.lights = "medium"
            self.co2_generator = "low"

        if energy_price > 0.10 and temp > 20 and y > 300:
            triggered.add(53)
            actions.append("温度适宜土壤湿度好：维持")
            pass

        if energy_price > 0.10 and temp > 18 and z < 400:
            triggered.add(54)
            actions.append("温度适宜但CO2不足：补CO2")
            self.co2_generator = "low"

        if energy_trend == "stable" and z < 300 and x + y > 600:
            triggered.add(55)
            actions.append("稳定趋势低CO2但光照土壤湿度总和好：补CO2")
            self.co2_generator = "medium"

        if weather_forecast == "sunny" and x < 600 and y > 300:
            triggered.add(56)
            actions.append("晴天但光照不足土壤湿度好：补光")
            self.lights = "low"

        if weather_forecast == "sunny" and x < 500 and z < 400:
            triggered.add(57)
            actions.append("晴天但光照CO2都不足：补光补CO2")
            self.lights = "medium"
            self.co2_generator = "low"

        if weather_forecast == "sunny" and y < 500 and x > 200:
            triggered.add(58)
            actions.append("晴天光照可以但土壤湿度低：浇水")
            self.water_pump = "medium"

        if weather_forecast == "sunny" and humidity > 50 and x + y > 800:
            triggered.add(59)
            actions.append("晴天湿度高光照土壤湿度总和高：通风除湿")
            self.vent_fan = "medium"
            self.dehumidifier = "on"

        if weather_forecast == "sunny" and temp > 20 and x * y > 200000:
            triggered.add(60)
            actions.append("晴天温度适宜光照土壤湿度乘积高：通风")
            self.vent_fan = "low"

        # 62-71: 复杂多变量优化条件
        if energy_price * 10 + time_of_day * 0.5 > 7 and x > 600 and z > 200:
            triggered.add(61)
            actions.append("复合系数高且光照CO2好：轻度调节")
            self.shade_screen = "half"

        if energy_price * temp > 3 and humidity > 50 and x * y > 300000:
            triggered.add(62)
            actions.append("能价温度乘积高且湿度高光照土壤湿度乘积高：综合调节")
            self.vent_fan = "medium"
            self.dehumidifier = "on"

        if energy_price * temp > 2.5 and humidity > 50 and z < 400:
            triggered.add(63)
            actions.append("能价温度乘积高且湿度高但CO2不足：补CO2除湿")
            self.co2_generator = "medium"
            self.dehumidifier = "on"

        if (1 - energy_price * 2) * (x + z) > 800 and y < 400:
            triggered.add(64)
            actions.append("修正系数光照CO2高但土壤湿度低：浇水")
            self.water_pump = "medium"

        if time_of_day == 12 and x * temp * 0.1 > 2000 and y + z > 600:
            triggered.add(65)
            actions.append("中午光照温度系数高且土壤湿度CO2总和高：降温通风")
            self.cooling = "low"
            self.vent_fan = "medium"

        if time_of_day == 12 and (z + y) * 0.8 < 400 and x > 300:
            triggered.add(66)
            actions.append("中午CO2土壤湿度系数低但光照可以：补CO2浇水")
            self.co2_generator = "medium"
            self.water_pump = "medium"

        if energy_price * 10 + humidity * 0.1 > 7 and temp > 20 and z > x / 2:
            triggered.add(67)
            actions.append("复合湿度系数高且温度适宜CO2相对充足：通风")
            self.vent_fan = "low"

        if (1 - energy_price * 3) * (x + temp) > 500 and z < 300 and x + y > 800:
            triggered.add(68)
            actions.append("修正光照温度系数高但CO2低且光照土壤湿度总和高：补CO2")
            self.co2_generator = "high"

        if energy_price > 0.10 and abs(time_of_day - 12) < 1 and (x + temp) > 600 and y * z > 100000:
            triggered.add(69)
            actions.append("中午前后光照温度高且土壤湿度CO2乘积高：综合降温通风")
            self.cooling = "medium"
            self.vent_fan = "medium"

        if (24 - time_of_day > 8) and (y + humidity) < 400 and x > z:
            triggered.add(70)
            actions.append("远离中午且湿度总和低但光照大于CO2：补湿补CO2")
            self.water_pump = "medium"
            self.co2_generator = "low"

        # 输出触发的条件和执行的动作
        print(f"第二类扩展版触发了 {len(triggered)} 个条件: {sorted(triggered)}")

        return actions

    def section3_low_light_high_moisture_enhanced(self, light, moisture, co2):
        """部分3: 低光照、高湿度组合 - 使用 light, moisture, co2"""
        actions = []
        triggered = set()

        # 1-10: 基础条件
        if light < 3500 and moisture > 75:
            triggered.add(1)
            actions.append("低光高湿：启动补光和通风")
            self.lights = "medium"
            self.vent_fan = "medium"

        if light < 2500 and moisture > 80 and co2 > 1600:
            triggered.add(2)
            actions.append("极低光、高湿和高CO2：补光并加强通风")
            self.lights = "high"
            self.vent_fan = "high"

        if light > 3000 and moisture < 75:
            triggered.add(3)
            actions.append("严重缺光且湿度过高：强力补光和通风")
            self.lights = "high"
            self.vent_fan = "high"
            self.dehumidifier = "on"

        if (light < 3500 or co2 < 800) and moisture > 42:
            triggered.add(4)
            actions.append("低光或低CO2且高湿：综合调节")
            if light < 2500:
                self.lights = "high"
            if co2 < 700:
                self.co2_generator = "medium"
            self.vent_fan = "high"

        if moisture / (light / 100) > 8:
            triggered.add(5)
            actions.append("湿度光照比过高：补光除湿")
            self.lights = "medium"
            self.vent_fan = "medium"

        if light < 3000 and (light + moisture * 10) < 3800:
            triggered.add(6)
            actions.append("低光照高湿度组合：补光除湿")
            self.lights = "medium"
            self.dehumidifier = "on"

        if light > 5000 and light < 7000 and co2 > 1400:
            triggered.add(7)
            actions.append("中等光照高CO2：通风调节")
            self.vent_fan = "low"

        if light > 6000 and (light + moisture + co2) > 8500:
            triggered.add(8)
            actions.append("光照偏高三变量总和高：通风调节")
            self.vent_fan = "medium"

        if light > 8500 and moisture < 50:
            triggered.add(9)
            actions.append("高光照低湿度：遮阳浇水")
            self.shade_screen = "half"
            self.water_pump = "medium"

        if light > 7500 and co2 < 900:
            triggered.add(10)
            actions.append("强光照低CO2：遮阳补CO2")
            self.shade_screen = "half"
            self.co2_generator = "medium"

        # 11-25: 基于light的条件
        if light > 7000 and co2 < 1000:
            triggered.add(11)
            actions.append("高光照低CO2：轻度遮阳补CO2")
            self.shade_screen = "half"
            self.co2_generator = "low"

        if light > 6500 and co2 > 1700:
            triggered.add(12)
            actions.append("光照CO2都偏高：通风遮阳")
            self.vent_fan = "medium"
            self.shade_screen = "half"

        if light < 4000 and moisture < 45:
            triggered.add(13)
            actions.append("低光照低湿度：补光浇水")
            self.lights = "medium"
            self.water_pump = "medium"

        if light > 7000 and moisture > 78:
            triggered.add(14)
            actions.append("高光照高湿度：遮阳除湿")
            self.shade_screen = "half"
            self.dehumidifier = "on"

        if light > 6000 and moisture > 75:
            triggered.add(15)
            actions.append("中高光照高湿度：轻度遮阳通风")
            self.shade_screen = "half"
            self.vent_fan = "low"

        if light < 2800 and moisture > 80:
            triggered.add(16)
            actions.append("低光照极高湿度：补光除湿")
            self.lights = "high"
            self.dehumidifier = "on"

        if light < 3500 and co2 < 900:
            triggered.add(17)
            actions.append("低光照低CO2：补光补CO2")
            self.lights = "medium"
            self.co2_generator = "medium"

        if light > 7500 and co2 > 1700:
            triggered.add(18)
            actions.append("高光照高CO2：遮阳通风")
            self.shade_screen = "half"
            self.vent_fan = "high"

        if light * 0.1 + co2 * 0.1 > 900:
            triggered.add(19)
            actions.append("光照CO2系数高：通风调节")
            self.vent_fan = "low"

        if light > 8500 and co2 < 700:
            triggered.add(20)
            actions.append("极高光照极低CO2：遮阳补CO2")
            self.shade_screen = "full"
            self.co2_generator = "high"

        if light < 2200 and moisture < 40:
            triggered.add(21)
            actions.append("极低光照极低湿度：强力补光浇水")
            self.lights = "high"
            self.water_pump = "high"

        if light > 8000 and moisture > 70:
            triggered.add(22)
            actions.append("高光照中高湿度：遮阳除湿")
            self.shade_screen = "half"
            self.dehumidifier = "on"

        if light / (co2 + 1) > 7:
            triggered.add(23)
            actions.append("光照CO2比值极高：补CO2")
            self.co2_generator = "medium"

        if light < 2000 and co2 < 650:
            triggered.add(24)
            actions.append("双低值：全面补充")
            self.lights = "high"
            self.co2_generator = "high"

        if light > 9000 and moisture < 45:
            triggered.add(25)
            actions.append("极强光低湿：遮阳浇水")
            self.shade_screen = "full"
            self.water_pump = "high"

        # 26-40: 基于moisture的条件
        if moisture < 40 and light < 3500:
            triggered.add(26)
            actions.append("极低湿度低光照：浇水补光")
            self.water_pump = "high"
            self.lights = "medium"

        if moisture < 45 and co2 < 850:
            triggered.add(27)
            actions.append("低湿度低CO2：浇水补CO2")
            self.water_pump = "medium"
            self.co2_generator = "medium"

        if moisture < 55 and light < 4500:
            triggered.add(28)
            actions.append("中低湿度低光照：浇水补光")
            self.water_pump = "medium"
            self.lights = "low"

        if moisture > 50 and moisture < 70 and light > 6000:
            triggered.add(29)
            actions.append("中等湿度高光照：通风遮阳")
            self.vent_fan = "low"
            self.shade_screen = "half"

        if moisture > 65 and light > 6000:
            triggered.add(30)
            actions.append("高湿度高光照：除湿通风")
            self.dehumidifier = "on"
            self.vent_fan = "medium"

        if moisture > 60 and co2 > 1400:
            triggered.add(31)
            actions.append("高湿度高CO2：通风除湿")
            self.vent_fan = "medium"
            self.dehumidifier = "on"

        if moisture > 78 and light < 6000:
            triggered.add(32)
            actions.append("极高湿度中等光照：除湿通风")
            self.dehumidifier = "on"
            self.vent_fan = "high"

        if moisture > 80 and co2 < 1000:
            triggered.add(33)
            actions.append("极高湿度低CO2：通风补CO2")
            self.vent_fan = "high"
            self.co2_generator = "low"

        if moisture < 55 and light > 8500:
            triggered.add(34)
            actions.append("低湿度极高光照：浇水遮阳")
            self.water_pump = "high"
            self.shade_screen = "full"

        if moisture > 78 and light < 2500:
            triggered.add(35)
            actions.append("高湿度低光照：补光除湿")
            self.lights = "high"
            self.dehumidifier = "on"

        if moisture > 75 and light > 7500:
            triggered.add(36)
            actions.append("高湿度强光照：遮阳除湿")
            self.shade_screen = "half"
            self.dehumidifier = "on"

        if moisture < 58 and co2 < 950:
            triggered.add(37)
            actions.append("低湿度低CO2：浇水补CO2")
            self.water_pump = "medium"
            self.co2_generator = "medium"

        if moisture > 82 and co2 > 1600:
            triggered.add(38)
            actions.append("极高湿度高CO2：除湿通风")
            self.dehumidifier = "on"
            self.vent_fan = "high"

        if moisture / (light / 100) > 10:
            triggered.add(39)
            actions.append("湿度光照比极高：补光除湿")
            self.lights = "high"
            self.dehumidifier = "on"

        if moisture < 42 and co2 > 1750:
            triggered.add(40)
            actions.append("低湿度极高CO2：浇水通风")
            self.water_pump = "high"
            self.vent_fan = "high"

        # 41-50: 基于co2的条件
        if co2 < 900 and light < 4500:
            triggered.add(41)
            actions.append("低CO2低光照：补CO2补光")
            self.co2_generator = "medium"
            self.lights = "medium"

        if co2 > 1600 and light > 7000:
            triggered.add(42)
            actions.append("高CO2高光照：通风遮阳")
            self.vent_fan = "high"
            self.shade_screen = "half"

        if co2 > 1750 and moisture < 45:
            triggered.add(43)
            actions.append("高CO2低湿度：通风浇水")
            self.vent_fan = "high"
            self.water_pump = "medium"

        if co2 > 1850 and moisture > 80:
            triggered.add(44)
            actions.append("极高CO2高湿度：强力通风除湿")
            self.vent_fan = "high"
            self.dehumidifier = "on"

        if co2 < 750 and moisture > 82:
            triggered.add(45)
            actions.append("低CO2极高湿度：补CO2除湿")
            self.co2_generator = "high"
            self.dehumidifier = "on"

        if co2 / (light + 1) > 0.22:
            triggered.add(46)
            actions.append("CO2光照比高：通风补光")
            self.vent_fan = "medium"
            self.lights = "low"

        if co2 > 1700 and light < 3000:
            triggered.add(47)
            actions.append("高CO2低光照：通风补光")
            self.vent_fan = "high"
            self.lights = "medium"

        if co2 < 800 and moisture < 48:
            triggered.add(48)
            actions.append("低CO2低湿度：补CO2浇水")
            self.co2_generator = "medium"
            self.water_pump = "medium"

        if co2 * 0.6 + light * 0.08 > 1400:
            triggered.add(49)
            actions.append("CO2光照加权高：通风遮阳")
            self.vent_fan = "medium"
            self.shade_screen = "half"

        if co2 > 1900 and moisture < 38:
            triggered.add(50)
            actions.append("极高CO2极低湿度：通风浇水")
            self.vent_fan = "high"
            self.water_pump = "high"

        # 51-59: 三变量综合条件
        if light < 3500 and moisture < 48 and co2 < 900:
            triggered.add(51)
            actions.append("三变量全低值：全面补充")
            self.lights = "high"
            self.water_pump = "high"
            self.co2_generator = "high"

        if light + moisture * 20 + co2 < 4500:
            triggered.add(52)
            actions.append("三变量加权和极低：紧急全面提升")
            self.lights = "high"
            self.water_pump = "high"
            self.co2_generator = "high"

        if light + moisture * 25 + co2 > 10500:
            triggered.add(53)
            actions.append("三变量加权和极高：全面降低")
            self.shade_screen = "full"
            self.vent_fan = "high"
            self.dehumidifier = "on"

        if light > 6000 and moisture < 48 and co2 > 1400:
            triggered.add(54)
            actions.append("高光低湿高CO2：遮阳浇水通风")
            self.shade_screen = "half"
            self.water_pump = "high"
            self.vent_fan = "medium"

        if light > 7500 and moisture < 42 and co2 > 1600:
            triggered.add(55)
            actions.append("强光极低湿高CO2：强力调节")
            self.shade_screen = "full"
            self.water_pump = "high"
            self.vent_fan = "high"

        if light < 2500 and moisture > 85 and co2 < 750:
            triggered.add(56)
            actions.append("低光极高湿低CO2：补光除湿补CO2")
            self.lights = "high"
            self.dehumidifier = "on"
            self.co2_generator = "high"

        if light / (moisture + 1) > 110:
            triggered.add(57)
            actions.append("光照湿度比极高：浇水")
            self.water_pump = "high"

        if (light + moisture + co2) / 3 > 2500:
            triggered.add(58)
            actions.append("三变量平均值高：综合降低")
            self.shade_screen = "half"
            self.vent_fan = "medium"

        if light * 0.35 + moisture * 9 + co2 * 0.18 > 3200:
            triggered.add(59)
            actions.append("三变量复杂加权高：强力调节")
            self.shade_screen = "full"
            self.vent_fan = "high"

        if light * moisture * co2 > 250000000:
            triggered.add(60)
            actions.append("三变量乘积极高：全面降低")
            self.shade_screen = "full"
            self.vent_fan = "high"
            self.dehumidifier = "on"

        print(f"第三类触发了 {len(triggered)} 个条件: {sorted(triggered)}")
        return actions

    def section4_low_light_low_moisture_enhanced(self, light, moisture, temp):
        """部分4: 低光照、低湿度组合 - 使用 light, moisture, temp"""
        actions = []
        triggered = set()

        # 1-10: 基础条件
        if light < 3500 and moisture < 55:
            triggered.add(1)
            actions.append("低光低湿：启动补光和浇水")
            self.lights = "medium"
            self.water_pump = "medium"

        if light < 2200 and moisture < 48 and temp < 20:
            triggered.add(2)
            actions.append("极低光、极干旱、低温：全面干预")
            self.lights = "high"
            self.water_pump = "high"
            self.heating = "medium"

        if light < 2000 and moisture < 42:
            triggered.add(3)
            actions.append("严重缺光且严重干旱：强力补光补水")
            self.lights = "high"
            self.water_pump = "high"

        if (light < 1800 or temp < 18) and moisture < 45:
            triggered.add(4)
            actions.append("极端低光或低温且干旱：综合调节")
            if light < 1800:
                self.lights = "high"
            if temp < 18:
                self.heating = "medium"
            self.water_pump = "high"

        if light * moisture < 160000:
            triggered.add(5)
            actions.append("光照湿度乘积低：综合提升")
            self.lights = "high"
            self.water_pump = "medium"

        if light < 2800 and moisture < 52 and temp < 22:
            triggered.add(6)
            actions.append("低光、干旱且低温：三重调节")
            self.lights = "high"
            self.water_pump = "medium"
            self.heating = "high"

        if light > 2500 and light < 3500 and moisture > 50 and moisture < 58:
            triggered.add(7)
            actions.append("轻度缺光和轻度干旱：轻度调节")
            self.lights = "low"
            self.water_pump = "low"

        if (3000 - light) > (58 - moisture) * 25:
            triggered.add(8)
            actions.append("缺光问题大于干旱：优先补光")
            self.lights = "high"
            self.water_pump = "low"

        if light < 2800 and (moisture < 52 or temp < 20):
            triggered.add(9)
            actions.append("低光且干旱或低温：针对性调节")
            self.lights = "medium"
            if moisture < 52:
                self.water_pump = "medium"
            if temp < 20:
                self.heating = "medium"

        if light > 8500 and temp > 30:
            triggered.add(10)
            actions.append("强光高温：遮阳降温")
            self.shade_screen = "full"
            self.cooling = "high"

        # 继续添加更多基于 light, moisture, temp 的条件...
        if light > 8000 and moisture > 82 and temp > 32:
            triggered.add(60)
            actions.append("三变量全高值：全面降低")
            self.shade_screen = "full"
            self.vent_fan = "high"
            self.cooling = "high"
            self.dehumidifier = "on"

        print(f"第四类触发了 {len(triggered)} 个条件: {sorted(triggered)}")
        return actions

    def section5_high_co2_extremes_enhanced(self, light, co2, temp):
        """部分5: 高CO2与极端组合 - 使用 light, co2, temp"""
        actions = []
        triggered = set()

        # 1-10: 基础CO2控制
        if co2 > 1700:
            triggered.add(1)
            actions.append("高CO2：加强通风")
            self.vent_fan = "high"

        if co2 > 1650 and light > 8000 and temp > 28:
            triggered.add(2)
            actions.append("高CO2、高光和高温：通风降温遮阳")
            self.vent_fan = "high"
            self.cooling = "medium"
            self.shade_screen = "half"

        if co2 > 1800:
            triggered.add(3)
            actions.append("CO2远超理想值：紧急通风")
            self.vent_fan = "high"

        if (co2 > 1750 and light < 2800) or (co2 > 1900 and temp > 32):
            triggered.add(4)
            actions.append("高CO2且低光，或极高CO2且高温：强力通风")
            self.vent_fan = "high"

        if co2 / (light / 10) > 2.5 and temp < 22:
            triggered.add(5)
            actions.append("CO2/光照比过高且低温：通风加热")
            self.vent_fan = "medium"
            self.heating = "low"

        if co2 > 1700 and light > 9000:
            triggered.add(6)
            actions.append("高CO2与高光照：通风遮阳")
            self.vent_fan = "high"
            self.shade_screen = "half"

        if co2 > 1700 and temp > 30:
            triggered.add(7)
            actions.append("高CO2与高温：强力通风降温")
            self.vent_fan = "high"
            self.cooling = "medium"

        if co2 > 1900:
            triggered.add(8)
            actions.append("极高CO2：最高级别通风")
            self.vent_fan = "high"

        if light > 8500 and temp > 30:
            triggered.add(9)
            actions.append("强光高温：遮阳降温通风")
            self.shade_screen = "half"
            self.cooling = "medium"
            self.vent_fan = "medium"

        if light > 7500 and co2 > 1600:
            triggered.add(10)
            actions.append("高光照高CO2：通风遮阳")
            self.shade_screen = "half"
            self.vent_fan = "medium"

        # 继续添加更多基于 light, co2, temp 的条件...
        if light > 8500 and co2 > 1750 and temp > 30:
            triggered.add(60)
            actions.append("三变量均高值：全面降低")
            self.shade_screen = "full"
            self.vent_fan = "high"
            self.cooling = "high"

        print(f"第五类触发了 {len(triggered)} 个条件: {sorted(triggered)}")
        return actions

    def section6_low_co2_extremes(self, moisture, co2, temp):
        """部分6: 低CO2与极端组合 - 使用 moisture, co2, temp"""
        actions = []
        triggered = set()

        if co2 < 700:
            triggered.add(1)
            actions.append("低CO2：补充CO2")
            self.co2_generator = "medium"

        if co2 < 750 and moisture > 78 and temp > 28:
            triggered.add(2)
            actions.append("低CO2、高湿和高温：补充CO2并降温除湿")
            self.co2_generator = "high"
            self.cooling = "low"
            self.dehumidifier = "on"

        if co2 < 650:
            triggered.add(3)
            actions.append("CO2严重不足：强力补充")
            self.co2_generator = "high"

        if (co2 < 620 and moisture > 82) or (co2 < 580 and temp < 18):
            triggered.add(4)
            actions.append("低CO2且高湿，或极低CO2且低温：补充CO2")
            self.co2_generator = "high"
            if temp < 18:
                self.heating = "medium"

        if moisture / (co2 + 1) > 0.095 and temp < 22:
            triggered.add(5)
            actions.append("湿度/CO2比过高且低温：补充CO2加热")
            self.co2_generator = "medium"
            self.heating = "low"

        if co2 < 680 and moisture > 85:
            triggered.add(6)
            actions.append("低CO2与高湿度：补充CO2并除湿")
            self.co2_generator = "high"
            self.dehumidifier = "on"

        if co2 < 680 and temp < 18:
            triggered.add(7)
            actions.append("低CO2与低温：补充CO2并加热")
            self.co2_generator = "high"
            self.heating = "medium"

        if co2 < 550:
            triggered.add(8)
            actions.append("极低CO2：最高级别补充")
            self.co2_generator = "high"

        if co2 < 620 and moisture < 42:
            triggered.add(9)
            actions.append("低CO2低湿：补CO2浇水")
            self.co2_generator = "high"
            self.water_pump = "medium"

        if co2 < 750 and temp > 30:
            triggered.add(10)
            actions.append("低CO2高温：补CO2降温")
            self.co2_generator = "medium"
            self.cooling = "medium"

        # 继续添加更多基于 moisture, co2, temp 的条件...
        if moisture < 35 and co2 < 600 and temp < 18:
            triggered.add(60)
            actions.append("三变量全低值：全面提升")
            self.water_pump = "high"
            self.co2_generator = "high"
            self.heating = "high"

        print(f"第六类触发了 {len(triggered)} 个条件: {sorted(triggered)}")
        return actions

    def section7_normal_conditions(self, light, moisture, humidity):
        """部分7: 正常条件微调 - 使用 light, moisture, humidity"""
        actions = []
        triggered = set()

        if light > 7200 and light < 8000:
            triggered.add(1)
            actions.append("光照接近上限：准备遮阳")
            self.shade_screen = "ready"

        if light < 3300 and light > 3000:
            triggered.add(2)
            actions.append("光照接近下限：准备补光")
            self.lights = "ready"

        if moisture > 72 and moisture < 80:
            triggered.add(3)
            actions.append("湿度接近上限：准备通风")
            self.vent_fan = "ready"

        if moisture < 66 and moisture > 60:
            triggered.add(4)
            actions.append("湿度接近下限：准备浇水")
            self.water_pump = "ready"

        if humidity > 72 and humidity < 78:
            triggered.add(5)
            actions.append("空气湿度接近上限：准备除湿")
            self.dehumidifier = "ready"

        if humidity < 48 and humidity > 42:
            triggered.add(6)
            actions.append("空气湿度接近下限：准备增湿")
            self.water_pump = "ready"

        if (3000 <= light <= 8000 and
                60 <= moisture <= 80 and
                45 <= humidity <= 75):
            triggered.add(7)
            actions.append("所有参数在理想范围内：保持当前状态")

        if light > 6000 and moisture > 70 and humidity > 68:
            triggered.add(8)
            actions.append("光照、土湿、气湿都偏高：通风除湿")
            self.vent_fan = "low"
            self.dehumidifier = "on"

        if light < 4000 and moisture < 55 and humidity < 50:
            triggered.add(9)
            actions.append("光照、土湿、气湿都偏低：补光浇水")
            self.lights = "low"
            self.water_pump = "low"

        if light / (humidity + 1) > 120 and moisture < 55:
            triggered.add(10)
            actions.append("光照气湿比高且土湿低：浇水增湿")
            self.water_pump = "medium"

        print(f"第七类触发了 {len(triggered)} 个条件: {sorted(triggered)}")
        return actions

    def set_crop_parameters(self, crop_type):
        """根据不同作物设置理想参数范围"""
        crop_params = {
            "tomato": {
                "light_low": 4000, "light_high": 9000,
                "temp_day_low": 22, "temp_day_high": 28,
                "temp_night_low": 16, "temp_night_high": 20,
                "humidity_low": 50, "humidity_high": 70
            },
            "lettuce": {
                "light_low": 2500, "light_high": 6000,
                "temp_day_low": 18, "temp_day_high": 24,
                "temp_night_low": 12, "temp_night_high": 16,
                "humidity_low": 60, "humidity_high": 80
            },
            # 可以继续添加其他作物...
        }
        if crop_type in crop_params:
            params = crop_params[crop_type]

    def get_sensor_data(self):
        """获取传感器数据"""
        import random
        # 现有传感器...
        light = random.uniform(1000, 10000)
        moisture = random.uniform(30, 90)
        co2 = random.uniform(500, 2000)

        # 可以添加的新传感器
        temperature = random.uniform(15, 35)
        humidity = random.uniform(40, 90)
        soil_ph = random.uniform(5.5, 7.5)  # 土壤pH值
        nutrient_level = random.uniform(0, 100)  # 营养液浓度
        leaf_wetness = random.uniform(0, 100)  # 叶片湿度

        return light, moisture, co2, temperature, humidity, soil_ph, nutrient_level, leaf_wetness

    def control_greenhouse(self):
        """主控制函数"""
        light, moisture, co2, temp, humidity = self.get_sensor_data()
        print(
            f"当前温室数据 - 光照: {light:.0f}lux, 土壤湿度: {moisture:.1f}%, CO2: {co2:.0f}ppm, 温度: {temp:.1f}°C, 空气湿度: {humidity:.1f}%")

        # 重置所有设备状态
        self.lights = "off"
        self.water_pump = "off"
        self.co2_generator = "off"
        self.vent_fan = "off"
        self.shade_screen = "off"
        self.heating = "off"
        self.cooling = "off"
        self.dehumidifier = "off"

        # 判断执行哪个部分的控制逻辑
        if co2 > self.co2_ideal_high * 1.2:
            section = 5
            actions = self.section5_high_co2_extremes_enhanced(light, co2, temp)
        elif co2 < self.co2_ideal_low * 0.8:
            section = 6
            actions = self.section6_low_co2_extremes(moisture, co2, temp)
        elif light < self.light_ideal_low and moisture > self.moisture_ideal_high:
            section = 3
            actions = self.section3_low_light_high_moisture_enhanced(light, moisture, co2)
        elif light < self.light_ideal_low and moisture < self.moisture_ideal_low:
            section = 4
            actions = self.section4_low_light_low_moisture_enhanced(light, moisture, temp)
        else:
            section = 7
            actions = self.section7_normal_conditions(light, moisture, humidity)

        print(f"执行部分 {section} 的控制策略")
        for action in actions:
            print(f"- {action}")

        print("\n设备状态:")
        print(f"补光灯: {self.lights}")
        print(f"水泵: {self.water_pump}")
        print(f"CO2发生器: {self.co2_generator}")
        print(f"通风扇: {self.vent_fan}")
        print(f"遮阳网: {self.shade_screen}")
        print(f"加热系统: {self.heating}")
        print(f"降温系统: {self.cooling}")
        print(f"除湿系统: {self.dehumidifier}\n{'-' * 50}")


def section8_nutrient_control(self, nutrient_level, moisture, light):
    """部分8: 营养液控制"""
    actions = []
    triggered = set()

    if nutrient_level < 30 and light > 5000:
        triggered.add(1)
        actions.append("营养不足但光照充足：补充营养液")
        self.nutrient_pump = "medium"

    if nutrient_level > 80 and moisture > 70:
        triggered.add(2)
        actions.append("营养过剩且湿度过高：停止营养液，加强排水")
        self.nutrient_pump = "off"
        self.irrigation_system = "drain"

    # 添加更多营养控制条件...
    print(f"第八类触发了 {len(triggered)} 个条件: {sorted(triggered)}")
    return actions


def section9_ph_control(self, soil_ph, moisture, temp):
    """部分9: pH值控制"""
    actions = []
    triggered = set()

    if soil_ph < 6.0 and moisture > 60:
        triggered.add(1)
        actions.append("土壤偏酸且湿度合适：调节pH值")
        # 添加pH调节逻辑

    if soil_ph > 7.2 and temp > 25:
        triggered.add(2)
        actions.append("土壤偏碱且温度高：酸化和降温")
        # 添加调节逻辑

    # 添加更多pH控制条件...
    print(f"第九类触发了 {len(triggered)} 个条件: {sorted(triggered)}")
    return actions


def section10_seasonal_control(self, light, temp, humidity, season="summer"):
    """部分10: 季节性控制策略"""
    actions = []
    triggered = set()

    if season == "summer":
        if light > 8000 and temp > 30:
            triggered.add(1)
            actions.append("夏季强光高温：加强降温遮阳")
            self.cooling = "high"
            self.shade_screen = "full"

    elif season == "winter":
        if light < 3000 and temp < 15:
            triggered.add(2)
            actions.append("冬季弱光低温：加强补光加热")
            self.lights = "high"
            self.heating = "high"

    print(f"第十类触发了 {len(triggered)} 个条件: {sorted(triggered)}")
    return actions
# 运行智能温室控制系统
if __name__ == "__main__":
    # 测试代码
        controller = GreenhouseController()

        # 测试数据记录
        sensor_data = controller.get_sensor_data()
        record = controller.record_sensor_data(*sensor_data)
        print("数据记录测试:", record['sensors'])

        # 测试冲突解决
        controller.heating = "high"
        controller.cooling = "high"
        controller.resolve_conflicts()
        print(f"加热状态: {controller.heating}, 降温状态: {controller.cooling}")

    controller = GreenhouseController()
for i in range(5):
        print(f"=== 控制循环 {i + 1} ===")
        controller.control_greenhouse()
        # 可以添加的新设备
        self.fog_system = "off"  # 雾化系统
        self.nutrient_pump = "off"  # 营养液泵
        self.irrigation_system = "off"  # 灌溉系统
        self.uv_light = "off"  # UV杀菌灯
        self.circulation_fan = "off"  # 循环风扇

        # 添加历史数据记录
        self.sensor_history = []
        self.action_history = []

        # 添加作物类型特定参数
        self.crop_type = "tomato"  # 默认作物类型
        self.set_crop_parameters(self.crop_type)
def record_sensor_data(self, *sensor_data):
    """记录传感器数据历史"""
    self.sensor_history.append({
        'timestamp': time.time(),
        'data': sensor_data
    })
    # 保持历史数据长度
    if len(self.sensor_history) > 1000:
        self.sensor_history.pop(0)

def reset_all_devices(self):
    """重置所有设备状态"""
    self.lights = "off"
    self.water_pump = "off"
    self.co2_generator = "off"
    self.vent_fan = "off"
    self.shade_screen = "off"
    self.heating = "off"
    self.cooling = "off"
    self.dehumidifier = "off"
    # 新设备的重置
    self.fog_system = "off"
    self.nutrient_pump = "off"
    self.irrigation_system = "off"

def energy_optimization(self):
    """能量优化策略"""
    actions = []
    # 根据电价时段优化设备运行
    # 根据天气预报优化设备运行
    return actions

def resolve_conflicts(self):
    """解决设备操作冲突"""
    # 例如：不能同时加热和降温
    if self.heating != "off" and self.cooling != "off":
        # 根据优先级解决冲突
        if self.temperature < 15:  # 如果温度很低
            self.cooling = "off"   # 优先加热
        else:
            self.heating = "off"   # 优先降温

def generate_report(self):
    """生成温室运行报告"""
    if len(self.sensor_history) < 2:
        return "数据不足生成报告"

    # 分析趋势、效率等
    report = {
        'avg_temperature': sum([d['data'][3] for d in self.sensor_history]) / len(self.sensor_history),
        'energy_usage': self.calculate_energy_usage(),
        'efficiency_score': self.calculate_efficiency(),
        # 更多指标...
    }

    return report