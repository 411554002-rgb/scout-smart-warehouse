import os
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import pptx
from pptx.util import Inches as PInches, Pt as PPt
from pptx.dml.color import RGBColor as PRGBColor

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tcPr.append(shd)

# ==============================================================================
# 課程一專屬：【創意應用 / 創意實作】(純 LINE Bot + Google Sheets + Gemini，不用 Wokwi)
# ==============================================================================

def create_course1_report_docx(output_path):
    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = Inches(0.8)
        s.bottom_margin = Inches(0.8)
        s.left_margin = Inches(0.8)
        s.right_margin = Inches(0.8)
        
    title = doc.add_paragraph()
    r = title.add_run("期中專題報告：童軍庫房 LINE Bot 雲端智慧管家系統")
    r.font.size = Pt(20)
    r.font.bold = True
    r.font.color.rgb = RGBColor(16, 124, 65) # LINE綠
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    sub = doc.add_paragraph()
    sub_r = sub.add_run("整合 Google Apps Script、雲端試算表與 Gemini AI 之無人化器材管理方案")
    sub_r.font.size = Pt(12)
    sub_r.font.color.rgb = RGBColor(100, 100, 100)
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph("─" * 45).alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 摘要
    h_abs = doc.add_heading(level=2)
    h_abs.add_run("【摘要 Abstract】").font.color.rgb = RGBColor(16, 124, 65)
    p_abs = doc.add_paragraph()
    p_abs.paragraph_format.line_spacing = 1.2
    p_abs.add_run(
        "針對童軍團器材庫房物品繁多、紙本登記容易疏漏、外勤露營無法即時查閱庫存等痛點，"
        "本專案開發一套基於 LINE Bot 的雲端智慧管家系統。透過 Google Apps Script (GAS) 架設雲端後端，"
        "小隊員可在手機 LINE 直接輸入自然語言（如「查 虎頭鉗」、「借 虎頭鉗」、「A區清單」），"
        "系統自動於 Google 試算表即時扣減可用數、累計借出數、防呆阻擋超額借出，並在借還紀錄自動登記時間與人員。"
        "同時導入 Google Gemini AI，即時解答露營裝備配置與童軍工程技能，達成高效無紙化管理。"
    )

    # 壹、做什麼 (What)
    h1 = doc.add_heading(level=2)
    h1.add_run("壹、 做什麼 (What)").font.color.rgb = RGBColor(16, 124, 65)
    p1 = doc.add_paragraph()
    p1.paragraph_format.line_spacing = 1.2
    p1.add_run(
        "本專案研發「童軍庫房 LINE Bot 雲端智慧管家」，主要功能涵蓋：\n"
        "1. 即時庫存查詢：隊員輸入「查 虎頭鉗」或「查 A1-2」，即時回報品名、櫃位位置、現存可用數量與狀態。\n"
        "2. 線上借還即時扣帳：輸入「借 虎頭鉗」或「還 虎頭鉗」，後端自動至 Google 試算表增減庫存，並於「借還紀錄」自動寫入時間戳記與借用人員。\n"
        "3. 全庫房與分區清單總覽：輸入「所有器材」查看品項統計；輸入「A區」、「B區」即時列出該分區所有裝備。\n"
        "4. 庫存防呆保護：當現存可用數為 0 時，機器人主動拒絕借出並發出缺貨警告；在庫滿額時提醒無需重複歸還。\n"
        "5. Gemini AI 童軍諮詢：當使用者詢問露營技巧、生火工具或繩結架設時，由 AI 提供專業指引與裝備建議。"
    )

    # 貳、給誰用 (Who)
    h2 = doc.add_heading(level=2)
    h2.add_run("貳、 給誰用 (Who)").font.color.rgb = RGBColor(16, 124, 65)
    p2 = doc.add_paragraph()
    p2.paragraph_format.line_spacing = 1.2
    p2.add_run(
        "1. 童軍團小隊長與小隊員：在露營營地或外出活動時，直接用手機 LINE 查詢器材放哪、剩多少，並能隨時手機借還。\n"
        "2. 童軍器材長與團內幹部：告別繁瑣混亂的紙本登記簿，由試算表自動彙整歷程，盤點效率提升 90% 以上。\n"
        "3. 學校社團道具庫與實驗室：模組化架構能快速移植至各類公共物資管理情境。"
    )

    # 參、怎麼做 (How)
    h3 = doc.add_heading(level=2)
    h3.add_run("參、 怎麼做 (How)").font.color.rgb = RGBColor(16, 124, 65)
    p3 = doc.add_paragraph()
    p3.paragraph_format.line_spacing = 1.2
    p3.add_run(
        "1. 前端介面 (LINE Messaging API)：\n"
        "   - 使用者透過 LINE 官方帳號輸入訊息或點擊底部 Quick Reply 快捷按鈕。\n"
        "2. 雲端後端核心 (Google Apps Script)：\n"
        "   - 部署為 Web App 接收 LINE Webhook POST 請求。\n"
        "   - 智慧正則表達式解析借還指令（如匹配「借/還 + 器材名 + 數量」）。\n"
        "   - 動態表頭對齊演算法：自動定位試算表中的「編號、名稱、建檔總數量、現存可用數」欄位。\n"
        "   - 呼叫 SpreadsheetApp.flush() 強制將數值即刻寫入試算表實體儲存。\n"
        "3. 資料庫與 AI (Google Sheets + Gemini API)：\n"
        "   - 以雲端試算表為持久化資料庫，記錄最新庫存與流水帳。\n"
        "   - 串接 Google Gemini-1.5-Flash，賦予系統自然語言問答與童軍知識庫能力。"
    )

    # 肆、怎麼測 (Test)
    h4 = doc.add_heading(level=2)
    h4.add_run("肆、 怎麼測 (Test)").font.color.rgb = RGBColor(16, 124, 65)
    
    table = doc.add_table(rows=1, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = table.rows[0].cells
    for i, title_text in enumerate(["測試情境", "操作輸入", "預期反應", "實測結果"]):
        hdr[i].text = title_text
        set_cell_background(hdr[i], "107C41")
        hdr[i].paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
        hdr[i].paragraphs[0].runs[0].font.bold = True
        
    test_cases = [
        ("1. 查詢特定品項", "傳送「查 虎頭鉗」", "回覆櫃位 A1-2、可用數 1 支、狀態正常", "✅ 通過，秒級回傳"),
        ("2. 登記借出扣帳", "傳送「借 虎頭鉗」", "回覆借出成功，試算表可用數減 1，借還紀錄新增一筆", "✅ 通過，試算表即時連動扣帳"),
        ("3. 缺貨防呆阻擋", "再次傳送「借 虎頭鉗」", "回覆無法借出，提示庫存僅剩 0 支", "✅ 通過，成功阻擋超額借出"),
        ("4. 登記歸還入庫", "傳送「還 虎頭鉗」", "回覆歸還成功，試算表可用數加 1 回復為 1", "✅ 通過，帳物正確恢復"),
        ("5. 分區清單總覽", "傳送「A區」或「所有器材」", "列出 A 櫃所有器材，或全庫房統計看板", "✅ 通過，自動排版分頁"),
        ("6. AI 技能諮詢", "詢問「露營生火需要什麼器材？」", "Gemini AI 條列建議裝備與安全注意事項", "✅ 通過，專業解答")
    ]
    for row in test_cases:
        r_cells = table.add_row().cells
        for idx, text in enumerate(row):
            r_cells[idx].text = text
            r_cells[idx].paragraphs[0].paragraph_format.line_spacing = 1.1

    # 伍、下一步 (Next Steps)
    h5 = doc.add_heading(level=2)
    h5.add_run("伍、 下一步 (Next Steps)").font.color.rgb = RGBColor(16, 124, 65)
    p5 = doc.add_paragraph()
    p5.paragraph_format.line_spacing = 1.2
    p5.add_run(
        "1. Flex Message 視覺化卡片升級：將純文字清單升級為圖文並茂的互動式卡片，附帶一鍵借還按鈕。\n"
        "2. 逾期未還定時推播：利用 Apps Script 觸發器 (Trigger)，每週一定時檢查借用名單並發送提醒。\n"
        "3. QR Code 掃描即借：為庫房每個器材箱張貼 QR Code，隊員掃描直接開啟 LINE 填入代碼。"
    )

    doc.save(output_path)
    print(f"Course 1 Report DOCX saved to {output_path}")

def create_course1_summary_docx(output_path):
    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = Inches(0.8)
        s.bottom_margin = Inches(0.8)
        s.left_margin = Inches(0.8)
        s.right_margin = Inches(0.8)
        
    title = doc.add_paragraph()
    r = title.add_run("AI 協同開發過程摘要報告 (LINE Bot 雲端智慧管家)")
    r.font.size = Pt(18)
    r.font.bold = True
    r.font.color.rgb = RGBColor(16, 124, 65)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph("─" * 45).alignment = WD_ALIGN_PARAGRAPH.CENTER

    stages = [
        ("一、 需求釐清與系統功能定位",
         "• 開發者需求：\n"
         "  希望利用 LINE 官方帳號做為操作入口，讓童軍隊員能直接在手機上查詢庫存，並能即時更動 Google 試算表（扣減可用數、記錄借用人與時間），同時具備 AI 顧問功能。\n"
         "• AI 建議架構：\n"
         "  選用 Google Apps Script (GAS) 部署 Webhook，無須自租伺服器即可 24 小時免費用戶連線，並直接操作原生 Google 試算表 API。"),
         
        ("二、 指令邏輯設計與試算表欄位對齊",
         "• 問題：童軍試算表資料龐雜（幾百行），且包含「編號、名稱、數量、單位、借出數、可用數」等不同表頭。\n"
         "• AI 解決方案：\n"
         "  撰寫全表動態表頭定位器，自動判斷表頭列位置。設計正則表達式，支援「查 虎頭鉗」、「借 虎頭鉗」、「還 虎頭鉗 1」、「A區」等多元自然語言指令。"),
         
        ("三、 LINE Webhook 302 錯誤與權限突破 (重大除錯)",
         "• 問題：在 LINE Developers 按下 Verify 時出現「HTTP 302 Found」驗證失敗。\n"
         "• 深入診斷：GAS 在未進行手動執行授權前，所有外部訪問都會被 Google 強制轉址 (302) 到登入頁面；此外 Verify 送出空事件時若無妥善處理亦會被拒。\n"
         "• AI 解決方案：\n"
         "  1. 撰寫專門的 testAuthorizeAll() 函式，手動執行強制觸發試算表與網路權限核准。\n"
         "  2. 改寫 doPost，在接收到空事件或 Verify ping 時直接秒回純文字「OK」(HTTP 200)。"),
         
        ("四、 試算表數值持久化延遲與可用數修正 (重大除錯)",
         "• 問題：初期借出成功後，立即歸還或查詢時數值未同步，甚至可用數顯示為布林值 false。\n"
         "• 深入診斷：試算表部分欄位含有舊標記，且 GAS 寫入試算表存在快取延遲，尚未寫入硬碟就被下一次查詢讀到舊值。\n"
         "• AI 解決方案：\n"
         "  在每次借還更新儲存格後，強制執行 SpreadsheetApp.flush() 即時寫入實體儲存，並修正可用數抓取邏輯，確保帳物 100% 同步。"),
         
        ("五、 使用者體驗升級：Quick Reply 底部快捷氣泡",
         "• AI 優化方案：\n"
         "  關閉 LINE 官方後台預設罐頭回應，並在每次回覆時動態附加 6 顆 Quick Reply 氣泡按鈕（📊 庫存總覽、🔨 A區、⛺ B區、🍳 C區、⚠️ 缺貨清單、📖 說明），隊員點擊按鈕即可一鍵查詢，大幅提升操作流暢度。")
    ]
    
    for title_text, desc_text in stages:
        h = doc.add_heading(level=2)
        h.add_run(title_text).font.color.rgb = RGBColor(16, 124, 65)
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.2
        p.add_run(desc_text)
        
    doc.save(output_path)
    print(f"Course 1 Summary DOCX saved to {output_path}")

def create_course1_presentation_pptx(output_path):
    prs = pptx.Presentation()
    prs.slide_width = PInches(13.333)
    prs.slide_height = PInches(7.5)
    blank_layout = prs.slide_layouts[6]
    
    slides = [
        {
            "title": "童軍庫房 LINE Bot 雲端智慧管家系統",
            "sub": "【創意應用課】期中專題 90 秒展示簡報",
            "bullets": [
                "• 核心定位：手機端 LINE 官方帳號 ＋ Google 試算表雲端資料庫 ＋ Gemini AI",
                "• 解決痛點：童軍庫房裝備數百項，外勤露營無從查閱、紙本登記遺漏、帳物嚴重不符",
                "• 亮點特色：自然語言查借還、試算表自動扣減、防呆阻擋缺貨、AI 露營諮詢"
            ]
        },
        {
            "title": "90 秒展示流程：打開什麼、做什麼、看到什麼",
            "sub": "精準時間分配，流暢展示真實軟體連動全貌",
            "bullets": [
                "• 【打開什麼】(00~20秒)：投影打開「Google 雲端試算表」並向大家展示「手機 LINE 官方帳號」",
                "• 【做什麼】(20~60秒)：",
                "   1. 手機輸入「查 虎頭鉗」確認目前庫存 1 支",
                "   2. 手機輸入「借 虎頭鉗」登記借出，接著再次輸入「借 虎頭鉗」測試防呆",
                "   3. 手機輸入「還 虎頭鉗」完成歸還，並展示底部快捷氣泡一鍵查「A區」",
                "• 【看到什麼】(60~90秒)：",
                "   1. LINE 秒回漂亮卡片；試算表可用數即時從 1 變 0，借還紀錄自動新增時間與姓名！",
                "   2. 缺貨時機器人立刻警報防呆；歸還後試算表自動回復正常，全自動無紙化！"
            ]
        },
        {
            "title": "系統架構與業務邏輯流 (Architecture)",
            "sub": "Google Apps Script 雲端微服務 × 雙向資料即時同步",
            "bullets": [
                "• 使用者端 (LINE App)：透過自然語言或 Quick Reply 快捷按鈕發送指令",
                "• 雲端中介層 (Google Apps Script)：",
                "   - 處理 LINE Webhook 請求，智慧正則表達式解析指令 (查/借/還/分區)",
                "   - 動態表頭定位器適應複雜試算表欄位，SpreadsheetApp.flush() 杜絕快取延遲",
                "• 資料庫與智慧大腦 (Google Sheets + Gemini API)：",
                "   - 試算表即時記錄庫存數量與借還日誌，Gemini 賦予專業童軍露營問答能力"
            ]
        },
        {
            "title": "重大卡關與除錯歷程 (Debug Process)",
            "sub": "從連線錯誤到即時資料持久化的關鍵突破",
            "bullets": [
                "• 1. LINE Webhook 302 Found 轉址報錯：",
                "   - 撰寫 testAuthorizeAll() 強制觸發 Google 試算表與網路權限授權，並針對 Verify 請求秒回 200 OK",
                "• 2. 試算表數值未即時變更與 false 布林值異常：",
                "   - 重寫表頭定位器，全表動態鎖定「現存可用數」；呼叫 flush() 強制實體儲存，徹底消除快取延遲",
                "• 3. LINE 官方預設罐頭回應干擾：",
                "   - 調整 LINE Official Account Manager 設定，停用自動回應，使介面極致純淨俐落"
            ]
        },
        {
            "title": "專題效益與下一步展望",
            "sub": "告別紙本混亂，實現童軍團部智慧化管理",
            "bullets": [
                "• 專題實務成效：",
                "   - 借還登記零漏記：手機點擊即時連動試算表，權責分明",
                "   - 盤點時間大幅縮短 90%：一鍵產出「缺貨清單」與「各區器材清單」",
                "• 未來拓展：",
                "   - 整合 Flex Message 圖文卡片與 QR Code 一掃即借",
                "   - 加入每週定時借用未還自動推播催繳提醒"
            ]
        }
    ]
    
    for s_data in slides:
        slide = prs.slides.add_slide(blank_layout)
        top_bar = slide.shapes.add_shape(
            pptx.enum.shapes.MSO_SHAPE.RECTANGLE,
            PInches(0), PInches(0), PInches(13.333), PInches(1.5)
        )
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = PRGBColor(16, 124, 65) # LINE綠
        top_bar.line.color.rgb = PRGBColor(16, 124, 65)
        
        tx_box = slide.shapes.add_textbox(PInches(0.8), PInches(0.2), PInches(11.7), PInches(0.6))
        tf = tx_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = s_data["title"]
        p.font.size = PPt(26)
        p.font.bold = True
        p.font.color.rgb = PRGBColor(255, 255, 255)
        
        sub_box = slide.shapes.add_textbox(PInches(0.8), PInches(0.85), PInches(11.7), PInches(0.5))
        sub_tf = sub_box.text_frame
        sub_p = sub_tf.paragraphs[0]
        sub_p.text = s_data["sub"]
        sub_p.font.size = PPt(15)
        sub_p.font.color.rgb = PRGBColor(230, 245, 235)
        
        c_box = slide.shapes.add_textbox(PInches(0.8), PInches(1.8), PInches(11.7), PInches(5.0))
        c_tf = c_box.text_frame
        c_tf.word_wrap = True
        for idx, bullet in enumerate(s_data["bullets"]):
            cp = c_tf.paragraphs[0] if idx == 0 else c_tf.add_paragraph()
            cp.text = bullet
            cp.font.size = PPt(18)
            cp.font.color.rgb = PRGBColor(40, 40, 40)
            cp.space_before = PPt(10)
            cp.space_after = PPt(6)
            
    prs.save(output_path)
    print(f"Course 1 PPTX saved to {output_path}")

# ==============================================================================
# 課程二專屬：【智慧生活創意實作】(100% Wokwi ESP32 + 4x4 Keypad + MicroPython)
# ==============================================================================

def create_course2_presentation_pptx(output_path):
    prs = pptx.Presentation()
    prs.slide_width = PInches(13.333)
    prs.slide_height = PInches(7.5)
    blank_layout = prs.slide_layouts[6]
    
    slides = [
        {
            "title": "童軍倉庫智慧櫃位引導與借還終端機",
            "sub": "【智慧生活創意實作】期中專題報告 (ESP32 + MicroPython on Wokwi)",
            "bullets": [
                "• 模擬平台：Wokwi 模擬器 (ESP32 DevKit v1 ＋ MicroPython)",
                "• 系統架構：完整輸入 ➔ 判斷 ➔ 輸出閉環，支援無限次重複操作",
                "• 課堂未教過零件：4×4 薄膜矩陣鍵盤 (Membrane Keypad)，直覺輸入櫃位代碼"
            ]
        },
        {
            "title": "① 這個系統做什麼、給誰用？",
            "sub": "解決童軍庫房器材成百上千、昏暗混亂、借還無防呆之核心痛點",
            "bullets": [
                "• 做什麼 (What)：",
                "   1. 櫃位快速檢索：輸入童軍編碼 (如 A1-2 虎頭鉗、B3-4 營燈)，即時於 OLED 顯示器材與庫存",
                "   2. 智慧尋架導引 (Pick-to-Light)：按 A 亮藍燈、按 B 亮綠燈、按 C 亮黃燈，黑暗中循光取物",
                "   3. 借還門禁防呆：借出成功馬達開鎖 90 度並扣庫存，3 秒自動鎖閉；缺貨時蜂鳴警報鎖死",
                "• 給誰用 (Who)：",
                "   - 童軍團小隊長/隊員：團集會前快速領取裝備",
                "   - 童軍器材長/社團保管人：無人化借還計數、免除盤點缺漏"
            ]
        },
        {
            "title": "② 用了哪個沒教過的零件、它怎麼接？",
            "sub": "4×4 矩陣薄膜按鍵 (Membrane Keypad) 原理與接線圖說明",
            "bullets": [
                "• 沒教過的零件：4×4 矩陣薄膜按鍵 (16 鍵：A-D, 0-9, *, #)",
                "• 工作原理：動態行列掃描。4 條 Row 輸出 LOW，4 條 Col 開啟內部上拉 (Pull-up) 讀取",
                "• 完整接線規範 (Pin Mapping)：",
                "   - 4x4 Keypad Rows ➔ ESP32 GPIO 19, 18, 5, 17",
                "   - 4x4 Keypad Cols ➔ ESP32 GPIO 16, 4, 0, 2 (內部上拉 PULL_UP)",
                "   - I2C OLED (SSD1306) ➔ SDA: GPIO 21, SCL: GPIO 22",
                "   - NeoPixel 8 顆 RGB ➔ DIN: GPIO 15",
                "   - Passive Buzzer ➔ PWM: GPIO 12",
                "   - Servo Motor (電子鎖) ➔ PWM: GPIO 13"
            ]
        },
        {
            "title": "③ 程式怎麼判斷？(流程說明與關鍵程式碼)",
            "sub": "狀態機架構：輸入解析 ➔ 庫存比對 ➔ 模式分支 ➔ 聲光與馬達輸出",
            "bullets": [
                "• 核心判斷流程：",
                "   Keypad 按鍵輸入 ➔ 格式自動校正 (A12 ➔ A1-2) ➔ 查表 INVENTORY_DB ➔ 模式判斷：",
                "   - 借出模式：若 stock > 0 ➔ stock - 1、亮流水綠燈、馬達轉 90 度開鎖；若 stock == 0 ➔ 警報防呆",
                "   - 歸還模式：若 stock < total ➔ stock + 1、開鎖放回；星號 [*] 切換借/還/查模式",
                "• 關鍵程式碼 (Keypad 行列掃描)：",
                "   for r_idx, r_pin in enumerate(self.row_pins):",
                "       r_pin.value(0) # 輪流拉低 Row",
                "       for c_idx, c_pin in enumerate(self.col_pins):",
                "           if c_pin.value() == 0: # 偵測按下",
                "               while c_pin.value() == 0: time.sleep_ms(10) # 軟體防彈跳 (Debounce)",
                "               r_pin.value(1); return self.keys[r_idx][c_idx]"
            ]
        },
        {
            "title": "④ 跑起來的畫面截圖與功能展示",
            "sub": "Wokwi 現場模擬運作之 4 大情境",
            "bullets": [
                "• 1. 待機循環首頁：OLED 顯示「SCOUT WAREHOUSE / Mode: BORROW / Loc: [Input Code]」",
                "• 2. 智慧尋架導引 (Pick-to-Light)：按下「A」鍵瞬間，NeoPixel 前兩顆亮出高亮藍光指引 A 櫃！",
                "• 3. 借出成功與開鎖：輸入「A12#」，OLED 顯示「BORROW OK! / Left: 3」，馬達轉至 90 度開鎖",
                "• 4. 缺貨防呆阻擋：庫存為 0 時，OLED 顯示「OUT OF STOCK!」，燈條狂閃紅光、蜂鳴器長嗶"
            ]
        },
        {
            "title": "⑤ 卡過什麼、怎麼解的？(除錯歷程)",
            "sub": "真實開發中遇到的 3 大硬體與程式痛點與解決之道",
            "bullets": [
                "• 卡關點 1：Keypad 掃描連點與幽靈字元 (Bounce Issue)",
                "   - 解法：在偵測到低電位時加入 while c_pin.value() == 0: sleep_ms(10)，待放開才回傳鍵值完成去彈跳",
                "• 卡關點 2：ESP32 開機限制腳位衝突",
                "   - 解法：避開 Strapping Pins，精選 GPIO 16, 4, 0, 2 並啟用內部上拉，確保電位穩定不飄移",
                "• 卡關點 3：伺服馬達 PWM 抖動與角度不準",
                "   - 解法：查閱 SG90 規格書 (50Hz 週期 20ms)，換算 10-bit PWM Duty 為 26(0度) 至 125(180度)，精準鎖定 90 度開鎖"
            ]
        }
    ]
    
    for s_data in slides:
        slide = prs.slides.add_slide(blank_layout)
        top_bar = slide.shapes.add_shape(
            pptx.enum.shapes.MSO_SHAPE.RECTANGLE,
            PInches(0), PInches(0), PInches(13.333), PInches(1.5)
        )
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = PRGBColor(30, 60, 114) # 科技藍
        top_bar.line.color.rgb = PRGBColor(30, 60, 114)
        
        tx_box = slide.shapes.add_textbox(PInches(0.8), PInches(0.2), PInches(11.7), PInches(0.6))
        tf = tx_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = s_data["title"]
        p.font.size = PPt(26)
        p.font.bold = True
        p.font.color.rgb = PRGBColor(255, 255, 255)
        
        sub_box = slide.shapes.add_textbox(PInches(0.8), PInches(0.85), PInches(11.7), PInches(0.5))
        sub_tf = sub_box.text_frame
        sub_p = sub_tf.paragraphs[0]
        sub_p.text = s_data["sub"]
        sub_p.font.size = PPt(15)
        sub_p.font.color.rgb = PRGBColor(220, 230, 245)
        
        c_box = slide.shapes.add_textbox(PInches(0.8), PInches(1.8), PInches(11.7), PInches(5.0))
        c_tf = c_box.text_frame
        c_tf.word_wrap = True
        for idx, bullet in enumerate(s_data["bullets"]):
            cp = c_tf.paragraphs[0] if idx == 0 else c_tf.add_paragraph()
            cp.text = bullet
            cp.font.size = PPt(18)
            cp.font.color.rgb = PRGBColor(40, 40, 40)
            cp.space_before = PPt(10)
            cp.space_after = PPt(6)
            
    prs.save(output_path)
    print(f"Course 2 PPTX saved to {output_path}")

if __name__ == "__main__":
    docs_dir = os.path.join(os.path.dirname(__file__), "docs")
    
    # 課程一檔案
    c1_report = os.path.join(docs_dir, "【創意應用課】專題報告_童軍庫房LINE_Bot雲端智慧管家.docx")
    c1_summary = os.path.join(docs_dir, "【創意應用課】AI協同開發過程摘要_LINE_Bot雲端管家.docx")
    c1_slides = os.path.join(docs_dir, "【創意應用課】90秒展示投影片_LINE_Bot雲端管家.pptx")
    
    # 課程二檔案
    c2_slides = os.path.join(docs_dir, "【智慧生活課】期中專題投影片_Wokwi_ESP32智慧櫃位引導終端.pptx")
    
    create_course1_report_docx(c1_report)
    create_course1_summary_docx(c1_summary)
    create_course1_presentation_pptx(c1_slides)
    create_course2_presentation_pptx(c2_slides)
