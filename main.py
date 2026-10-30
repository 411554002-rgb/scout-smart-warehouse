"""
=============================================================================
童軍倉庫智慧櫃位引導與借還終端機 (Scout Smart Warehouse & Asset Guidance System)
=============================================================================
【系統概述】
專為童軍團裝備庫房設計之智慧終端管理系統。
透過 4x4 矩陣薄膜按鍵（未教過零件）輸入櫃位代碼（如 A1-2、B3-4），
系統即時查詢內建器材資料庫，並提供：
 1. I2C OLED 螢幕：顯示器材品名、現存數量與操作提示
 2. NeoPixel 燈條：智慧尋架引導燈 (Pick-to-Light)，指引櫃位方向與庫存狀態
 3. Buzzer 蜂鳴器：按鍵聲、借還成功提示音、缺貨警報
 4. Servo 伺服馬達：智慧櫃門電子鎖，核准時自動解鎖，延遲後自動鎖定

【硬體配置與接線】
 1. 4x4 矩陣鍵盤 (Membrane Keypad) -> 未教過零件
    Rows: GPIO 19, 18, 5, 17
    Cols: GPIO 16, 4, 0, 2 (或自選通用GPIO)
 2. I2C OLED 螢幕 (SSD1306, 128x64):
    SDA -> GPIO 21, SCL -> GPIO 22
 3. NeoPixel RGB LED 燈條 (8 LEDs):
    DIN -> GPIO 15
 4. 伺服馬達 (Servo Motor - 電子櫃鎖):
    PWM -> GPIO 13
 5. 蜂鳴器 (Passive Buzzer):
    PWM -> GPIO 12
=============================================================================
"""

import time
import machine
from machine import Pin, I2C, PWM
import neopixel

# ---------------------------------------------------------------------------
# 1. 零件初始化：OLED 螢幕 (SSD1306 128x64)
# ---------------------------------------------------------------------------
try:
    import ssd1306
    i2c = I2C(0, scl=Pin(22), sda=Pin(21))
    oled = ssd1306.SSD1306_I2C(128, 64, i2c)
    HAS_OLED = True
except Exception as e:
    print("OLED 初始化警告:", e)
    HAS_OLED = False

# ---------------------------------------------------------------------------
# 2. 零件初始化：NeoPixel RGB 燈條 (8 顆 LED) - 智慧尋架引導 (Pick-to-Light)
# ---------------------------------------------------------------------------
NUM_LEDS = 8
np_pin = Pin(15, Pin.OUT)
np = neopixel.NeoPixel(np_pin, NUM_LEDS)

def set_all_leds(r, g, b):
    for i in range(NUM_LEDS):
        np[i] = (r, g, b)
    np.write()

def clear_leds():
    set_all_leds(0, 0, 0)

# ---------------------------------------------------------------------------
# 3. 零件初始化：蜂鳴器 (Buzzer)
# ---------------------------------------------------------------------------
buzzer = PWM(Pin(12))
buzzer.duty(0)

def beep(freq, duration_ms):
    try:
        buzzer.freq(freq)
        buzzer.duty(512)
        time.sleep_ms(duration_ms)
        buzzer.duty(0)
    except:
        buzzer.duty(0)

def sound_key_click():
    beep(1200, 30)

def sound_success():
    beep(880, 80)
    time.sleep_ms(40)
    beep(1320, 120)

def sound_alert():
    for _ in range(3):
        beep(400, 100)
        time.sleep_ms(60)

# ---------------------------------------------------------------------------
# 4. 零件初始化：伺服馬達 (Servo Motor - 電子櫃鎖)
# ---------------------------------------------------------------------------
servo = PWM(Pin(13), freq=50)

def set_servo_angle(angle):
    # 0 度 ~ 180 度對應 pulse 寬度 (約 0.5ms ~ 2.5ms)
    # 週期 20ms (50Hz), 0度: duty約 26~40, 90度: duty約 75~77, 180度: duty約 120~125
    duty = int(26 + (angle / 180.0) * (125 - 26))
    servo.duty(duty)

def lock_cabinet():
    set_servo_angle(0)  # 0度鎖上

def unlock_cabinet():
    set_servo_angle(90) # 90度開啟

# ---------------------------------------------------------------------------
# 5. 【沒教過的零件】4x4 矩陣薄膜按鍵驅動類別
# ---------------------------------------------------------------------------
class MatrixKeypad:
    """
    4x4 矩陣鍵盤掃描驅動：
    原理：將 Rows 設為輸出並輪流拉低 (LOW)，
          將 Cols 設為輸入並開啟上拉電阻 (Pull-up)。
          當某鍵被按下時，對應的 Col 會被拉為 LOW，藉此辨識行列交叉點。
    """
    def __init__(self, row_pins, col_pins):
        self.row_pins = [Pin(p, Pin.OUT) for p in row_pins]
        self.col_pins = [Pin(p, Pin.IN, Pin.PULL_UP) for p in col_pins]
        self.keys = [
            ['1', '2', '3', 'A'],
            ['4', '5', '6', 'B'],
            ['7', '8', '9', 'C'],
            ['*', '0', '#', 'D']
        ]
        # 初始將所有 Row 拉高
        for r in self.row_pins:
            r.value(1)

    def scan(self):
        pressed_key = None
        for r_idx, r_pin in enumerate(self.row_pins):
            r_pin.value(0) # 啟用目前 Row
            for c_idx, c_pin in enumerate(self.col_pins):
                if c_pin.value() == 0: # 偵測到按下
                    pressed_key = self.keys[r_idx][c_idx]
                    # 等待放開防彈跳 (Debounce)
                    while c_pin.value() == 0:
                        time.sleep_ms(10)
                    r_pin.value(1)
                    return pressed_key
            r_pin.value(1) # 復位
        return None

# 接腳配置 (Row: 19, 18, 5, 17 / Col: 16, 4, 0, 2)
# 註：ESP32 GPIO 0, 2, 4, 16 常用於輸入
keypad = MatrixKeypad(row_pins=[19, 18, 5, 17], col_pins=[16, 4, 0, 2])

# ---------------------------------------------------------------------------
# 6. 童軍器材資料庫 (對應真實童軍倉庫清單)
# ---------------------------------------------------------------------------
# 格式: 櫃位代碼: {"name": 品名, "total": 總量, "stock": 現存量, "unit": 單位}
INVENTORY_DB = {
    "A1-1": {"name": "Glow Rope",    "cname": "光繩",       "stock": 15, "total": 15, "unit": "條"},
    "A1-2": {"name": "Pliers/Tools", "cname": "虎頭鉗工具", "stock": 4,  "total": 4,  "unit": "支"},
    "A1-3": {"name": "Fire Stove",   "cname": "高山爐架",   "stock": 2,  "total": 2,  "unit": "個"},
    "A1-7": {"name": "Scout Ropes",  "cname": "粗童軍繩",   "stock": 42, "total": 42, "unit": "條"},
    "A2-1": {"name": "Stainless Pot","cname": "不銹鋼炒鍋", "stock": 4,  "total": 4,  "unit": "個"},
    "B3-1": {"name": "Gas Stove",    "cname": "卡式瓦斯爐", "stock": 1,  "total": 1,  "unit": "個"},
    "B3-4": {"name": "Camp Lantern", "cname": "露營營燈",   "stock": 2,  "total": 2,  "unit": "個"},
    "B3-6": {"name": "Paring Knife", "cname": "削皮刀/小刀","stock": 14, "total": 14, "unit": "支"},
    "C1-1": {"name": "Kitchen Tools","cname": "料理廚具砧板","stock": 5, "total": 5,  "unit": "個"},
    "C2-3": {"name": "Kettle/Pots",  "cname": "煮水壺鍋具", "stock": 2,  "total": 2,  "unit": "個"},
    "D2-1": {"name": "Ribbon/Wires", "cname": "活動彩線卷", "stock": 10, "total": 10, "unit": "捲"},
    "D2-7": {"name": "Walkie Talkie","cname": "無線對講機", "stock": 4,  "total": 4,  "unit": "台"}
}

# ---------------------------------------------------------------------------
# 7. UI 與視覺引導函式
# ---------------------------------------------------------------------------
MODE_BORROW = 0  # 借出模式
MODE_RETURN = 1  # 歸還模式
MODE_QUERY  = 2  # 查詢模式
MODE_NAMES  = ["BORROW (借出)", "RETURN (歸還)", "QUERY (查詢)"]

current_mode = MODE_BORROW
input_buffer = ""

def display_screen(header, line1, line2, line3, line4=""):
    if not HAS_OLED:
        print(f"[{header}] {line1} | {line2} | {line3} | {line4}")
        return
    oled.fill(0)
    oled.text(header[:16], 0, 0)
    oled.hline(0, 10, 128, 1)
    oled.text(line1[:16], 0, 14)
    oled.text(line2[:16], 0, 26)
    oled.text(line3[:16], 0, 38)
    if line4:
        oled.text(line4[:16], 0, 52)
    oled.show()

def update_idle_screen():
    m_name = MODE_NAMES[current_mode]
    code_show = input_buffer if input_buffer else "[Input Code]"
    display_screen(
        "SCOUT WAREHOUSE",
        f"Mode:{m_name.split()[0]}",
        f"Loc : {code_show}",
        "Press A-D/0-9",
        "*:Mode  #:Enter"
    )

def light_guide_for_cabinet(cab_char):
    """
    根據櫃位字母點亮不同區域的貨架指示燈 (Pick-to-Light)
    A 櫃: LED 0, 1 亮藍光
    B 櫃: LED 2, 3 亮綠光
    C 櫃: LED 4, 5 亮黃光
    D 櫃: LED 6, 7 亮紫光
    """
    clear_leds()
    cab = cab_char.upper()
    if cab == 'A':
        np[0] = (0, 40, 255)
        np[1] = (0, 40, 255)
    elif cab == 'B':
        np[2] = (0, 255, 40)
        np[3] = (0, 255, 40)
    elif cab == 'C':
        np[4] = (255, 200, 0)
        np[5] = (255, 200, 0)
    elif cab == 'D':
        np[6] = (200, 0, 255)
        np[7] = (200, 0, 255)
    else:
        set_all_leds(20, 20, 20)
    np.write()

def flash_leds_result(success):
    if success:
        # 成功流水綠光
        for _ in range(2):
            for i in range(NUM_LEDS):
                clear_leds()
                np[i] = (0, 255, 0)
                np.write()
                time.sleep_ms(30)
        set_all_leds(0, 120, 0)
        np.write()
    else:
        # 失敗全體紅光閃爍
        for _ in range(3):
            set_all_leds(255, 0, 0)
            np.write()
            time.sleep_ms(100)
            clear_leds()
            time.sleep_ms(100)

# ---------------------------------------------------------------------------
# 8. 雲端試算表同步模組 (Cloud Sync via Webhook)
# ---------------------------------------------------------------------------
# 可選功能：若啟用 WiFi，可將每一次借還即時同步至 Google 試算表 / LINE Bot 後端
ENABLE_WIFI_SYNC = False  # 在 Wokwi 離線純展示時設為 False；若需展示連網設為 True
GAS_WEBHOOK_URL = ""      # 佈署 Google Apps Script 後取得的 Web App 網址 (作業繳交時勿放真實通行碼)

def sync_to_cloud(code, action, user="ESP32-Terminal"):
    if not ENABLE_WIFI_SYNC or not GAS_WEBHOOK_URL:
        print(f"[Cloud Log] (離線模式) 已記錄本地異動: {action} {code} by {user}")
        return
    try:
        import urequests
        payload = {
            "source": "ESP32_TERMINAL",
            "code": code,
            "action": action,
            "user": user
        }
        res = urequests.post(GAS_WEBHOOK_URL, json=payload)
        print("[Cloud Sync] 試算表同步結果:", res.text)
        res.close()
    except Exception as e:
        print("[Cloud Sync] 連線雲端失敗:", e)

# ---------------------------------------------------------------------------
# 9. 核心業務邏輯處理 (輸入 -> 判斷 -> 輸出)
# ---------------------------------------------------------------------------
def execute_action(code):
    global current_mode
    # 格式校正：如果使用者輸入 A12，自動轉為標準編號 A1-2
    norm_code = code
    if len(code) == 3 and code[1].isdigit() and code[2].isdigit():
        norm_code = f"{code[0]}{code[1]}-{code[2]}"
    elif len(code) == 2 and code[1].isdigit():
        norm_code = f"{code[0]}{code[1]}-1"

    # 1. 判斷資料庫是否存在此器材
    if norm_code not in INVENTORY_DB:
        sound_alert()
        flash_leds_result(False)
        display_screen("ERROR!", "Unknown Loc:", norm_code, "Check Cabinet ID", "Press any key")
        time.sleep(2)
        return

    item = INVENTORY_DB[norm_code]
    cab_char = norm_code[0]
    light_guide_for_cabinet(cab_char) # 立即亮起該櫃指示燈！

    # 2. 依照當前模式判斷庫存
    if current_mode == MODE_QUERY:
        sound_success()
        display_screen(
            f"LOC: {norm_code}",
            item["name"][:16],
            f"Stock: {item['stock']}/{item['total']}",
            f"Unit : {item['unit']}",
            "Status: Ready"
        )
        time.sleep(3)

    elif current_mode == MODE_BORROW:
        if item["stock"] > 0:
            # 庫存充裕，執行借出
            item["stock"] -= 1
            sound_success()
            flash_leds_result(True)
            unlock_cabinet() # 轉動伺服馬達解鎖櫃門！
            
            display_screen(
                "BORROW OK!",
                f"Item:{item['name'][:10]}",
                f"Left: {item['stock']} {item['unit']}",
                "Lock Opened!",
                "Auto-lock in 3s"
            )
            # 同步雲端試算表
            sync_to_cloud(norm_code, "borrow", "ESP32-Locker")
            time.sleep(3)
            lock_cabinet() # 自動上鎖
        else:
            # 庫存為0，拒絕借出
            sound_alert()
            flash_leds_result(False)
            display_screen(
                "OUT OF STOCK!",
                f"Item:{item['name'][:10]}",
                "Stock is 0!",
                "Cannot Borrow",
                "All Lent Out!"
            )
            time.sleep(2.5)

    elif current_mode == MODE_RETURN:
        if item["stock"] < item["total"]:
            item["stock"] += 1
            sound_success()
            flash_leds_result(True)
            unlock_cabinet() # 開鎖放回器材
            display_screen(
                "RETURN OK!",
                f"Item:{item['name'][:10]}",
                f"Stock Now: {item['stock']}",
                "Thanks Scout!",
                "Cabinet Locked"
            )
            # 同步雲端試算表
            sync_to_cloud(norm_code, "return", "ESP32-Locker")
            time.sleep(3)
            lock_cabinet()
        else:
            sound_alert()
            display_screen(
                "WARN: FULL!",
                f"Stock Already Max",
                f"Total: {item['total']}",
                "No need return"
            )
            time.sleep(2)

    clear_leds()

# ---------------------------------------------------------------------------
# 9. 系統啟動迎賓與主循環 (重複操作狀態機)
# ---------------------------------------------------------------------------
def system_startup():
    print("=== Scout Smart Warehouse System Initializing ===")
    lock_cabinet()
    clear_leds()
    # 開機燈光效果
    for i in range(NUM_LEDS):
        np[i] = (50, 50, 50)
        np.write()
        time.sleep_ms(30)
    beep(988, 100)
    beep(1318, 150)
    clear_leds()

    display_screen(
        "* SCOUT LOCKER *",
        "Smart Terminal",
        "System Ready",
        "ESP32 + MicroPy"
    )
    time.sleep(1.5)
    update_idle_screen()

def main():
    global current_mode, input_buffer
    system_startup()

    while True:
        key = keypad.scan()
        if key:
            sound_key_click()
            print("Key Pressed:", key)

            if key == '*':
                # 切換操作模式 (借出 -> 歸還 -> 查詢)
                current_mode = (current_mode + 1) % 3
                input_buffer = ""
                update_idle_screen()

            elif key == '#':
                # 確認送出
                if len(input_buffer) > 0:
                    display_screen("PROCESSING...", f"Checking {input_buffer}", "Please wait...")
                    execute_action(input_buffer)
                    input_buffer = ""
                    update_idle_screen()
                else:
                    sound_alert()

            elif key == 'D':
                # D 鍵當作 Backspace / Clear
                input_buffer = ""
                clear_leds()
                update_idle_screen()

            else:
                # 輸入英文字母 (A/B/C) 或數字 (0-9)
                if len(input_buffer) < 5:
                    input_buffer += key
                    # 如果使用者剛輸入第一個字母 (A/B/C/D)，立即點亮對應櫃位導引燈！
                    if len(input_buffer) == 1 and input_buffer in ['A', 'B', 'C', 'D']:
                        light_guide_for_cabinet(input_buffer)
                    update_idle_screen()
                else:
                    sound_alert()

        time.sleep_ms(20)

if __name__ == '__main__':
    main()
