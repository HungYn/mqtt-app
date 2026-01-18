import datetime
import os
import time
import ctypes
import configparser
import paho.mqtt.client as mqtt
import logging
import uuid
import random
import string

INI_FILE = "limit_time.ini"
LOG_FILE = "limit_time.log"
mqtt_client = None

#打包成單一 .exe 檔案 →pyinstaller --noconsole --onefile limit_time.py

# 設定 logging
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)

def log_event(message):
    print(message)
    logging.info(message)

def load_config():
    config = configparser.ConfigParser()
    config.read(INI_FILE, encoding="utf-8")
    return config

# 星期對照表
WEEKDAY_MAP = {
    "monday": "monday", "mon": "monday", "星期一": "monday", "週一": "monday",
    "tuesday": "tuesday", "tue": "tuesday", "星期二": "tuesday", "週二": "tuesday",
    "wednesday": "wednesday", "wed": "wednesday", "星期三": "wednesday", "週三": "wednesday",
    "thursday": "thursday", "thu": "thursday", "星期四": "thursday", "週四": "thursday",
    "friday": "friday", "fri": "friday", "星期五": "friday", "週五": "friday",
    "saturday": "saturday", "sat": "saturday", "星期六": "saturday", "週六": "saturday",
    "sunday": "sunday", "sun": "sunday", "星期日": "sunday", "週日": "sunday", "星期天": "sunday"
}

def normalize_weekday(name: str) -> str:
    """將多語言星期轉換成標準英文名稱"""
    return WEEKDAY_MAP.get(name.lower(), name.lower())

def load_allowed_periods(config):
    periods_by_day = {}
    if "AllowedTimes" in config:
        for day in config["AllowedTimes"]:
            normalized_day = normalize_weekday(day)
            raw_periods = config["AllowedTimes"][day].split(",")
            day_periods = []
            for p in raw_periods:
                try:
                    start_str, end_str = p.strip().split("-")
                    start = datetime.time(int(start_str.split(":")[0]), int(start_str.split(":")[1]))
                    end   = datetime.time(int(end_str.split(":")[0]), int(end_str.split(":")[1]))
                    day_periods.append((start, end))
                except Exception as e:
                    log_event(f"⚠️ 格式錯誤: {p} ({e})")
            periods_by_day[normalized_day] = day_periods
    return periods_by_day

def check_time(periods_by_day):
    now = datetime.datetime.now()
    weekday = now.strftime("%A").lower()  # e.g. monday, tuesday...
    current_time = now.time()

    if weekday in periods_by_day:
        for start, end in periods_by_day[weekday]:
            if start <= current_time <= end:
                return True
    return False


def execute_action(config, reason=""):
    action = config.get("Action", "action", fallback="lock").lower()
    if action == "shutdown":
        log_event(f"⚠️ {reason} → 電腦關機")
        os.system("shutdown /s /f /t 0")
    elif action == "lock":
        log_event(f"⚠️ {reason} → 電腦鎖定螢幕")
        os.system("rundll32.exe user32.dll,LockWorkStation")

# MQTT 回呼函數
def on_message(client, userdata, msg):
    payload = msg.payload.decode("utf-8").strip()
    log_event(f"📩 收到 MQTT 訊息: {payload}")

    config = load_config()
    publish_topic = config["MQTT"]["publish_topic"]

    # 收到 shutdown關機 或 lock鎖定 → 立即執行動作
    if payload in ["shutdown","關機", "lock","鎖定"]:        
        client.publish(publish_topic, f"⚠️ 電腦即將執行 {payload}")
        if payload in ["shutdown", "關機"]:
            log_event("⚠️ 收到 shutdown 指令 → 電腦關機")
            os.system("shutdown /s /f /t 0")
        elif payload in ["lock", "鎖定"]:
            log_event("⚠️ 收到 lock 指令 → 電腦鎖定螢幕")
            os.system("rundll32.exe user32.dll,LockWorkStation")

    # 收到 action = lock 或 action = shutdown → 改寫 ini
    elif payload.startswith("action"):
        new_action = payload.replace("action =", "").strip()
        if new_action in ["lock", "shutdown"]:
            if "Action" not in config:
                config["Action"] = {}
            config["Action"]["action"] = new_action
            with open(INI_FILE, "w", encoding="utf-8") as f:
                config.write(f)
                
            client.publish(publish_topic, f"✅ 已更新動作模式: {new_action}")
            log_event(f"✅ 已更新動作模式: {new_action}")
        else:
            log_event(f"⚠️ 收到未知 action: {new_action}")

    # 收到 periods 設定 → 改寫 ini (支援多段星期更新)
    elif payload.startswith("periods"):
        try:
            # 允許多行設定，用分號或換行分隔
            lines = payload.replace("periods", "").strip().split(";")
            if "AllowedTimes" not in config:
                config["AllowedTimes"] = {}

            updated_days = []
            for line in lines:
                if "=" in line:
                    day, times = line.split("=", 1)
                    day = day.strip().lower()
                    times = times.strip()
                    if day:
                        config["AllowedTimes"][day] = times
                        updated_days.append(f"{day}={times}")

            # 寫回 ini 檔
            with open(INI_FILE, "w", encoding="utf-8") as f:
                config.write(f)

            # 推播回應
            if updated_days:
                msg = "✅ 已更新允許時段: " + "; ".join(updated_days)
                client.publish(publish_topic, msg)
                log_event(msg)
            else:
                client.publish(publish_topic, "⚠️ 未找到有效的 periods 設定")
                log_event("⚠️ 未找到有效的 periods 設定")

        except Exception as e:
            log_event(f"⚠️ 設定 periods 格式錯誤: {payload} ({e})")
            client.publish(publish_topic, f"⚠️ 設定 periods 格式錯誤: {payload}")

        
    # 收到 reset 或 重設 → 將 action 和 AllowedTimes 重設回 Defaults
    elif payload in ["reset", "重設"]:
        # 如果 [Defaults] 不存在 → 自動建立
        if "Defaults" not in config:
            config["Defaults"] = {}
            config["Defaults"]["action"] = "lock"
            # 預設一週的時段
            config["Defaults"]["星期一"] = "08:00-19:20,20:00-22:00"
            config["Defaults"]["星期二"] = "08:00-19:20,20:00-22:00"
            config["Defaults"]["星期三"] = "08:00-19:20,20:00-22:00"
            config["Defaults"]["星期四"] = "08:00-19:20,20:00-22:00"
            config["Defaults"]["星期五"] = "08:00-19:20,20:00-22:00"
            config["Defaults"]["星期六"] = "10:00-17:00"
            config["Defaults"]["星期日"] = "14:00-18:00"
            log_event("⚠️ [Defaults] 不存在，已自動建立")

        # 讀取 Defaults 的 action
        default_action = config["Defaults"].get("action", "lock")

        # 讀取 Defaults 的 AllowedTimes（星期格式）
        default_times = {day: times for day, times in config["Defaults"].items() if day != "action"}

        # 確保 Action 與 AllowedTimes 區塊存在
        if "Action" not in config:
            config["Action"] = {}
        if "AllowedTimes" not in config:
            config["AllowedTimes"] = {}

        # 更新 Action
        config["Action"]["action"] = default_action

        # 更新 AllowedTimes（逐一星期）
        for day, times in default_times.items():
            config["AllowedTimes"][day] = times

        # 寫回 ini 檔
        with open(INI_FILE, "w", encoding="utf-8") as f:
            config.write(f)

        client.publish(publish_topic, f"✅ 已重設動作模式與允許時段為 Defaults: action={default_action}, AllowedTimes={default_times}")
        log_event(f"✅ 已重設動作模式與允許時段為 Defaults: action={default_action}, AllowedTimes={default_times}")

            
    # 收到 status 或 狀態 → 查詢目前設定
    elif payload in ["status", "狀態"]:
        current_action = config.get("Action", "action", fallback="(未設定)")

        # 讀取所有星期的允許時段
        allowed_times = []
        if "AllowedTimes" in config:
            for day, times in config["AllowedTimes"].items():
                allowed_times.append(f"periods {day} = {times}")
        else:
            allowed_times.append("(未設定)")

        # 組合回傳訊息
        allowed_times_str = "; ".join(allowed_times)
        client.publish(publish_topic, f"ℹ️ 目前設定 → action = {current_action}, AllowedTimes = {allowed_times_str}")
        log_event(f"ℹ️ 查詢目前設定 → action = {current_action}, AllowedTimes = {allowed_times_str}")

            
# 自動重連回呼
def on_disconnect(client, userdata, rc):
    if rc != 0:
        log_event("⚠️ MQTT 斷線1，嘗試重新連線...")
        try:
            reconnect(client)
            log_event("✅ MQTT 已重新連線")
        except Exception as e:
            log_event(f"❌ 重連失敗: {e}")

def reconnect(client):
    config = load_config()
    broker = config["MQTT"]["broker"]
    port = int(config["MQTT"]["port"])
    subscribe_topic = config["MQTT"]["subscribe_topic"]

    while True:
        try:
            #client.connect(broker, port, 60)
            client.reconnect() 
            client.subscribe(subscribe_topic)
            log_event("✅ MQTT 已重新連線")
            client.loop_start() # 確保背景 loop 重新啟動
            break
        except Exception as e:
            log_event(f"❌ 重連失敗: {e}，60 秒後再試...")
            time.sleep(60)
            

def setup_mqtt(config):   
    broker = config["MQTT"]["broker"]
    port = int(config["MQTT"]["port"])
    client_id = "client-" + str(uuid.uuid4())    
    log_event(client_id)    
    subscribe_topic = config["MQTT"]["subscribe_topic"]

    client = mqtt.Client(client_id=client_id, protocol=mqtt.MQTTv311, transport="tcp")
    client.on_message = on_message
    client.on_disconnect = on_disconnect

    try:
        client.connect(broker, port, 60)        
        client.subscribe(subscribe_topic)
        client.loop_start()  # 背景執行
        log_event("✅ MQTT 已連線")
    except Exception as e:
        log_event(f"⚠️ 無法連線到 MQTT broker: {e}，將持續嘗試重連")
        client.loop_start()        
        # 啟動背景重連機制
        # 這裡不 raise，避免程式崩潰
    return client

if __name__ == "__main__":
    # 隱藏 console 視窗 (Windows only)
    whnd = ctypes.windll.kernel32.GetConsoleWindow()
    if whnd != 0:
        ctypes.windll.user32.ShowWindow(whnd, 0)
        ctypes.windll.kernel32.CloseHandle(whnd)

    config = load_config()
    mqtt_client = setup_mqtt(config)

    # 無限迴圈，每 60 秒重新載入 ini 並檢查
    while True:

        # 檢查 MQTT 是否斷線
        if not mqtt_client.is_connected():
            log_event("⚠️ MQTT 斷線2，嘗試重新連線...")
            try:
                #reconnect(mqtt_client) # 使用 paho-mqtt 內建 reconnect
                mqtt_client = setup_mqtt(config)
                client.loop_start() # 確保背景 loop 重新啟動
                log_event("✅ MQTT 已重新連線")
            except Exception as e:
                log_event(f"❌ MQTT 重連失敗: {e}")
            
        # 重新載入 ini
        config = load_config()
        periods_by_day = load_allowed_periods(config)


        # 判斷是否在允許時段
        if not check_time(periods_by_day):
            try:
                mqtt_client.publish(config["MQTT"]["publish_topic"], "⚠️ 不在允許時段，電腦即將關機")
            except Exception as e:
                log_event(f"⚠️ MQTT publish 失敗: {e}")
            execute_action(config, "不在允許時段")
        else:
            log_event("✅ 在允許時段 → 電腦正常使用")
            try:
                mqtt_client.publish(config["MQTT"]["publish_topic"], "✅ 在允許時段 → 電腦正常使用")
            except Exception as e:
                log_event(f"⚠️ MQTT publish 失敗: {e}")
        time.sleep(60) # 每 60 秒檢查一次
