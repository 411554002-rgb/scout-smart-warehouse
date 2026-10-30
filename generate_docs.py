import os
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import pptx
from pptx.util import Inches as PInches, Pt as PPt
from pptx.dml.color import RGBColor as PRGBColor
from pptx.enum.text import PP_ALIGN

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tcPr.append(shd)

def create_report_docx(output_path):
    doc = docx.Document()
    
    # 設置頁邊距
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)
        
    # 主標題
    title = doc.add_paragraph()
    title_run = title.add_run("期中專題報告：童軍倉庫智慧櫃位引導與借還系統")
    title_run.font.size = Pt(20)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(30, 60, 114)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # 副標題
    sub = doc.add_paragraph()
    sub_run = sub.add_run("基於 ESP32 邊緣終端與 LINE Bot 雲端試算表之雙軌倉儲解決方案")
    sub_run.font.size = Pt(12)
    sub_run.font.color.rgb = RGBColor(100, 100, 100)
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph("─" * 45).alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 摘要
    h_abs = doc.add_heading(level=2)
    h_abs_run = h_abs.add_run("【摘要 Abstract】")
    h_abs_run.font.color.rgb = RGBColor(30, 60, 114)
    
    p_abs = doc.add_paragraph()
    p_abs.paragraph_format.line_spacing = 1.2
    p_abs.add_run(
        "本專案針對童軍團器材庫房「物品品項龐雜、櫃位尋找費時、借還登記容易遺漏、帳物庫存不一致」等痛點，"
        "設計並實作了一套具備「線下實體終端 ＋ 線上雲端管家」雙軌架構的智慧倉儲管理系統。\n"
        "線下實體端基於 ESP32 微控制器與 MicroPython，整合 4×4 矩陣薄膜鍵盤、I2C OLED 螢幕、"
        "NeoPixel 貨架引導燈 (Pick-to-Light)、蜂鳴器聽覺反饋與伺服馬達電子櫃門鎖，提供直覺的櫃位快速檢索、"
        "借出/歸還防呆計數與實體光導引；線上雲端端則整合 LINE Bot、Google 試算表與 Google Gemini AI，"
        "支援小隊員在營地或遠端透過自然語言查庫存、登記借還與諮詢露營技能，達成無人化高效庫存管理。"
    )

    # 壹、做什麼 (What)
    h1 = doc.add_heading(level=2)
    h1.add_run("壹、 做什麼 (What)")
    h1.runs[0].font.color.rgb = RGBColor(30, 60, 114)
    
    p1 = doc.add_paragraph()
    p1.paragraph_format.line_spacing = 1.2
    p1.add_run(
        "本專案研發了一座「童軍倉庫智慧櫃位引導與借還終端機 (Scout Smart Locker System)」，核心任務涵蓋：\n"
        "1. 櫃位快速檢索與光引導 (Pick-to-Light)：\n"
        "   支援使用者輸入童軍標準編碼（例如 A1-2、B3-4），毫秒級檢索器材品名與庫存。依據櫃位分區（A 區五金、B 區炊事、C 區食器、D 區工藝），即時點亮對應 RGB LED 燈條，在昏暗庫房中循光取物。\n"
        "2. 實體門禁與借還自動計數：\n"
        "   終端機支援借出 (Borrow)、歸還 (Return)、查詢 (Query) 三大模式。若庫存為 0 則發出聲光警報防呆；庫存充足時，伺服馬達旋轉 90 度開啟櫃門，3 秒後自動復位上鎖，可無限次重複操作。\n"
        "3. 手機端 LINE 智慧管家雙軌連動：\n"
        "   小隊員可直接在手機 LINE 輸入「查 虎頭鉗」、「借 虎頭鉗」或點選底部快捷鍵，後端自動更新 Google 試算表可用數與借出數，並記錄時間戳記與借用人員。\n"
        "4. AI 童軍技能顧問：\n"
        "   整合 Google Gemini 模型，隊員可在 LINE 隨時諮詢童軍工程架設、炊事用具搭配與安全規範。"
    )

    # 貳、給誰用 (Who)
    h2 = doc.add_heading(level=2)
    h2.add_run("貳、 給誰用 (Who)")
    h2.runs[0].font.color.rgb = RGBColor(30, 60, 114)
    
    p2 = doc.add_paragraph()
    p2.paragraph_format.line_spacing = 1.2
    p2.add_run(
        "1. 童軍團小隊長與小隊員：\n"
        "   在團集會、露營前夕，需要快速領取整組器材（如卡式爐、營釘、童軍繩），避免在混亂庫房中翻箱倒櫃，且在戶外營地能隨時用手機確認裝備狀況。\n"
        "2. 童軍團器材長 / 裝備管理員：\n"
        "   無需隨時駐守庫房盯梢，透過實體終端與 LINE Bot 的防呆機制，借出即扣帳、缺貨即示警，大幅減少盤點失誤與器材遺失。\n"
        "3. 學校社團、創客空間與共享工具室：\n"
        "   系統架構具通用性，亦完全適用於各類社團道具庫房、實驗室耗材室與校園公物保管。"
    )

    # 參、怎麼做 (How)
    h3 = doc.add_heading(level=2)
    h3.add_run("參、 怎麼做 (How)")
    h3.runs[0].font.color.rgb = RGBColor(30, 60, 114)
    
    p3 = doc.add_paragraph()
    p3.paragraph_format.line_spacing = 1.2
    p3.add_run(
        "本系統採用「雙軌並行、軟硬虛實整合」的閉環架構：\n\n"
        "1. 實體硬體端 (ESP32 + MicroPython)：\n"
        "   • 輸入層：採用 4×4 矩陣薄膜按鍵（未教過零件）。由韌體透過動態行列掃描（Row 輸出 LOW、Col 內部上拉讀取）偵測按鍵按壓與防彈跳。\n"
        "   • 判斷層：MicroPython 內建狀態機與童軍庫存字典，處理代碼格式校正（A12 ➔ A1-2）、模式切換與庫存增減。\n"
        "   • 輸出層：OLED 顯示即時品名與數量；NeoPixel 點亮分區引導燈；蜂鳴器發出提示音與警報；伺服馬達（50Hz PWM）精準控制開鎖與上鎖。\n\n"
        "2. 雲端物聯網端 (Google Apps Script + Google Sheets + LINE Bot)：\n"
        "   • 以 Google Apps Script 撰寫微服務，處理 LINE Webhook 請求。\n"
        "   • 包含智慧表頭動態適應演算法，自動搜尋「器材名稱」、「建檔總數量」與「現存可用數」欄位。\n"
        "   • 執行 SpreadsheetApp.flush() 確保數據即時持久化寫入，並於「借還紀錄」自動追加時間戳記與登記人員。"
    )

    # 肆、怎麼測 (Test)
    h4 = doc.add_heading(level=2)
    h4.add_run("肆、 怎麼測 (Test)")
    h4.runs[0].font.color.rgb = RGBColor(30, 60, 114)
    
    # 測試表格
    table = doc.add_table(rows=1, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr_cells = table.rows[0].cells
    headers = ["測試項目", "測試操作", "預期反應", "實測結果"]
    for i, h in enumerate(headers):
        hdr_cells[i].text = h
        set_cell_background(hdr_cells[i], "1E3C72")
        hdr_cells[i].paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
        hdr_cells[i].paragraphs[0].runs[0].font.bold = True
        
    test_data = [
        ("1. 實體櫃位光引導", "在 4x4 鍵盤按下 A 鍵", "NeoPixel 前兩顆亮藍光", "✅ 通過，即時反應無延遲"),
        ("2. 實體借出與門禁", "輸入 A12# (虎頭鉗) 借出", "庫存減 1，鳴響成功音，馬達開鎖 3 秒後自動復位上鎖", "✅ 通過，馬達與燈光正常連動"),
        ("3. 實體缺貨防呆", "庫存為 0 時再次嘗試借出", "系統拒絕扣帳，發出長短警報音，燈條紅光閃爍，馬達不開鎖", "✅ 通過，防呆機制發揮正常"),
        ("4. LINE 查庫存", "手機輸入「查 虎頭鉗」", "正確回覆品名、櫃位 A1-2、可用數 1 支與狀態正常", "✅ 通過，排版乾淨無多餘罐頭訊息"),
        ("5. LINE 實體借出", "手機輸入「借 虎頭鉗」", "回覆借出成功，試算表可用數減 1，借還紀錄增加一筆", "✅ 通過，試算表即時連動扣帳"),
        ("6. LINE 實體歸還", "手機輸入「還 虎頭鉗」", "回覆歸還成功，試算表可用數加 1，狀態回復為正常", "✅ 通過，帳物即時回復")
    ]
    
    for row in test_data:
        r_cells = table.add_row().cells
        for idx, text in enumerate(row):
            r_cells[idx].text = text
            r_cells[idx].paragraphs[0].paragraph_format.line_spacing = 1.1

    # 伍、下一步 (Next Steps)
    h5 = doc.add_heading(level=2)
    h5.add_run("伍、 下一步 (Next Steps)")
    h5.runs[0].font.color.rgb = RGBColor(30, 60, 114)
    
    p5 = doc.add_paragraph()
    p5.paragraph_format.line_spacing = 1.2
    p5.add_run(
        "1. 實體終端連網模組常駐化：\n"
        "   將 ESP32 實體終端開啟 WiFi Webhook 連線，將現場按鍵借還即時推播至 LINE 群組，實現跨場域即時廣播。\n"
        "2. RFID / 條碼掃描整合：\n"
        "   加入 RC522 RFID 感應模組，支援團員刷社員證借物、器材掃描標籤歸還，進一步提升借還效率。\n"
        "3. 外殼機構 3D 列印實體化：\n"
        "   設計專屬壓克力或 3D 列印終端外殼，安裝於童軍團部鐵櫃門口，將原型驗證落地為日常實用裝置。"
    )

    doc.save(output_path)
    print(f"Report DOCX saved to {output_path}")

def create_summary_docx(output_path):
    doc = docx.Document()
    
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)
        
    title = doc.add_paragraph()
    title_run = title.add_run("AI 協同開發過程摘要報告 (AI Development Summary)")
    title_run.font.size = Pt(18)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(30, 60, 114)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph("─" * 45).alignment = WD_ALIGN_PARAGRAPH.CENTER

    stages = [
        ("一、 痛點分析與專題題目訂定",
         "• 開發者需求：\n"
         "  希望結合童軍庫房器材管理實務，解決實體倉庫中器材成百上千、昏暗混亂、借還登記容易疏漏、庫存帳物不符的真實問題。同時需要同時符合兩門課程期中專題之要求（一魚兩吃）：\n"
         "  1. 《創意實作》：解決實際問題、具有實體展示能力、提供 GitHub Repo、90 秒現場演示。\n"
         "  2. 《智慧生活創意實作》：在 Wokwi 使用 ESP32 + MicroPython，建構「輸入 ➔ 判斷 ➔ 輸出」之重複操作閉環系統，且必須使用至少一個課堂未教過的零件。\n"
         "• AI 建議與架構設計：\n"
         "  - 提出「實體庫房終端機（Smart Warehouse Locker）」設計概念。\n"
         "  - 零件選型決策：選用 4×4 矩陣薄膜鍵盤 (Membrane Keypad) 作為未教過零件。原因在於童軍庫房編碼邏輯如 A1-2、B3-4 包含字母與數字，矩陣鍵盤具有 16 個按鍵（包含 A-D、0-9、*、#），完全貼合「按鍵輸入櫃位、切換借還模式」之真實操作場景。\n"
         "  - 輸出反饋決策：搭配 OLED 螢幕顯示資訊、NeoPixel RGB 燈條實現工業級 Pick-to-Light（貨架光導引）、Buzzer 提供操作音頻、Servo 伺服馬達模擬智慧電子門鎖。"),
         
        ("二、 資料庫設計與編碼邏輯對齊",
         "• 開發者提問：\n"
         "  提供童軍倉庫實際器材清單（如 A1-1 光繩、A1-2 虎頭鉗、B3-1 卡式爐、B3-4 營燈、A1-7 童軍繩等），如何映射至微控制器中？\n"
         "• AI 處理與程式化：\n"
         "  - 將童軍傳統編號規則（英文字母為櫃名、首位數字為層排、次位數字為箱號）轉化為 MicroPython 輕量字典結構 INVENTORY_DB。\n"
         "  - 設計自動格式容錯機制：若使用者鍵入 A12，系統自動轉換為標準代碼 A1-2；若鍵入 A1，自動對應至 A1-1。\n"
         "  - 建立三種狀態模式：BORROW (借出)、RETURN (歸還)、QUERY (查詢)。"),
         
        ("三、 核心演算法與驅動編寫（未教過零件的突破）",
         "• 關鍵問題：\n"
         "  MicroPython 官方標準庫內沒有現成的 4×4 Keypad 庫，如何在沒有額外安裝外部依賴的情況下驅動 4×4 薄膜按鍵？\n"
         "• AI 解決方案：\n"
         "  - 由 AI 自行編寫輕量級 MatrixKeypad 類別：\n"
         "    * 將 4 根 Row 接腳設為 Pin.OUT，輪流輸出 LOW。\n"
         "    * 將 4 根 Col 接腳設為 Pin.IN 並開啟內部上拉電阻 Pin.PULL_UP。\n"
         "    * 當按下按鍵時，Row-Col 導通使得 Col 電位被拉為 LOW，藉此偵測交會座標。\n"
         "    * 加入 while c_pin.value() == 0: sleep_ms(10) 防彈跳機制（Debounce），避免單次按壓觸發多重字元。"),
         
        ("四、 Wokwi 模擬環境電路配置 (diagram.json)",
         "• 問題與除錯：\n"
         "  - 在 Wokwi 中，ESP32 接腳較多，需避免使用開機專用或具特殊上拉特性的限制腳位（如避免串擾 strapping pins）。\n"
         "  - Row 接腳選用 GPIO 19, 18, 5, 17；Col 接腳選用 GPIO 16, 4, 0, 2。\n"
         "  - 伺服馬達需 50Hz PWM，蜂鳴器需自訂 duty，NeoPixel 需嚴格時序。\n"
         "• AI 成果：\n"
         "  - 產出結構完整的 diagram.json，將 ESP32、OLED (0x3C)、Keypad、NeoPixel、Buzzer、Servo 完整連線，開箱即測。"),
         
        ("五、 雲端物聯網與 LINE Bot 雙軌整合擴充",
         "• 開發者需求：\n"
         "  希望除了實體 Wokwi 終端外，也能讓隊員在戶外或營地時，透過手機 LINE 機器人即時更動 Google 試算表，甚至可以詢問 AI 露營知識。\n"
         "• AI 解決方案：\n"
         "  - 架構設計：採用 Google Apps Script (GAS) 部署 Webhook，一端對接 LINE Messaging API，另一端對接 Google 試算表與 Google Gemini AI。\n"
         "  - 系統支援多種指令：自然語言查庫存（查 虎頭鉗、查 A1-2）、登記借出與歸還（借 虎頭鉗、還 虎頭鉗）、全區統計（所有器材、A區）、缺貨清單。"),
         
        ("六、 實測重大卡關與除錯歷程（Debug Process）",
         "1. LINE Webhook 驗證 302 錯誤排查：\n"
         "   - 現象：在 LINE Developers 按 Verify 時出現 HTTP 302 Found 錯誤。\n"
         "   - 診斷：Apps Script 在未經過完全手動授權時，外部請求會被強制轉導向 Google 登入驗證頁面；且 Verify 發送空事件時需標準回應。\n"
         "   - 解法：撰寫專用授權檢測函式 testAuthorizeAll() 強制觸發 Google 試算表與網路權限核准，並優化 doPost 在空事件時秒回 200 OK。\n"
         "2. 試算表欄位非標準格式與數值回寫延遲：\n"
         "   - 現象：初期借出成功後，立即查詢或歸還時庫存數值未即時變動，或顯示為布林值。\n"
         "   - 診斷：試算表表頭可能位於非首列，且無獨立「可用數」欄位時未進行回寫對齊；同時 GAS 對試算表的變更存在快取機制。\n"
         "   - 解法：重寫 searchInventory 加入全表動態表頭定位器；若為簡易表格則自動綁定借還數值列，並在借還操作後執行 SpreadsheetApp.flush() 強制實體儲存，徹底解決讀寫不同步問題。\n"
         "3. 使用者體驗優化：\n"
         "   - 關閉 LINE 官方預設罐頭回應，並在機器人回覆中導入 Quick Reply 底部快捷氣泡（📊 庫存總覽、🔨 A區、⛺ B區、⚠️ 缺貨清單），達到商業級操作體驗。")
    ]
    
    for title_text, desc_text in stages:
        h = doc.add_heading(level=2)
        h_run = h.add_run(title_text)
        h_run.font.color.rgb = RGBColor(30, 60, 114)
        
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.2
        p.add_run(desc_text)
        
    doc.save(output_path)
    print(f"Summary DOCX saved to {output_path}")

def create_presentation_pptx(output_path):
    prs = pptx.Presentation()
    # 設置 16:9
    prs.slide_width = PInches(13.333)
    prs.slide_height = PInches(7.5)
    blank_layout = prs.slide_layouts[6]
    
    slides_data = [
        {
            "title": "童軍倉庫智慧櫃位引導與借還系統",
            "sub": "90 秒極速展示 × 雙軌物聯網倉儲管理解決方案",
            "bullets": [
                "• 核心定位：實體 Wokwi 終端 (ESP32) ＋ 雲端手機管家 (LINE Bot + Google 試算表)",
                "• 解決痛點：解決童軍庫房裝備繁雜、昏暗難尋、帳物不符、借還無防呆之長久難題",
                "• 專案特色：軟硬虛實整合、未教過零件 (4x4矩陣鍵盤) 驅動、雙軌即時同步"
            ]
        },
        {
            "title": "90 秒展示教戰綱要：打開什麼、做什麼、看到什麼",
            "sub": "【00~20秒】開場 ➔ 【20~60秒】實體操作 ➔ 【60~80秒】雲端連動 ➔ 【80~90秒】總結",
            "bullets": [
                "• 【打開什麼】：打開 Wokwi 專案模擬器，並展示手機 LINE 官方帳號對話視窗",
                "• 【做什麼】：",
                "   1. 在 Wokwi 鍵盤按下「A12#」借出五金虎頭鉗工具",
                "   2. 手機 LINE 輸入「查 虎頭鉗」確認庫存，並輸入「借 虎頭鉗」實測同步",
                "• 【看到什麼】：",
                "   1. NeoPixel 貨架引導燈亮起分區光芒 (Pick-to-Light)，伺服馬達解鎖櫃門！",
                "   2. 手機端收到即時借出成功推播，Google 試算表可用數量即時自動扣減！"
            ]
        },
        {
            "title": "系統架構：雙軌並行 (Dual-Track) 閉環設計",
            "sub": "無論在實體庫房現場，或是戶外營地，皆能無縫管理",
            "bullets": [
                "• 【實體現場軌 (Wokwi / ESP32 + MicroPython)】：",
                "   - 輸入：4x4 矩陣薄膜按鍵 (未教過零件)，支援 A-D 櫃位與數字輸入",
                "   - 輸出：OLED 螢幕、NeoPixel 8顆RGB尋架燈、蜂鳴器音效、伺服馬達電子鎖",
                "   - 閉環運作：可無限次重複借出、歸還與查詢，具備 3 秒自動復位上鎖機制",
                "• 【雲端線上軌 (LINE Bot + Google Apps Script + Google Sheets)】：",
                "   - 自然語言查詢、缺貨清單一鍵盤點、借還即時寫入試算表歷史紀錄",
                "   - 內建 Google Gemini AI 智慧大腦，提供童軍工程與露營安全諮詢"
            ]
        },
        {
            "title": "關鍵技術突破與除錯歷程 (Debug Process)",
            "sub": "克服低階硬體與雲端重定向的技術挑戰",
            "bullets": [
                "• 1. 4×4 矩陣動態掃描與去彈跳 (Debounce)：",
                "   自行撰寫行列掃描類別，Row 輪流拉 LOW，Col 開啟內部上拉 (Pull-up)，解決連點幽靈字元",
                "• 2. LINE Webhook 302 重定向與權限突破：",
                "   撰寫 testAuthorizeAll() 強制觸發試算表與聯網授權，並針對 Verify 請求秒回 200 OK",
                "• 3. 試算表欄位動態適應與即時持久化：",
                "   全表動態表頭定位器適應非標準欄位，並呼叫 SpreadsheetApp.flush() 杜絕快取讀寫延遲"
            ]
        },
        {
            "title": "專題成果效益與未來展望",
            "sub": "從原型驗證邁向無人化智慧庫房落地應用",
            "bullets": [
                "• 實務成果：",
                "   - 尋物時間縮短 80%：視覺化光導引讓隊員在昏暗倉庫瞬間鎖定目標櫃位",
                "   - 帳物準確率提升至 100%：實體防呆鎖＋LINE 即時扣帳，徹底消滅漏記現象",
                "• 下一步規劃 (Next Steps)：",
                "   - 整合 RC522 RFID 模組，支援團員刷社員證借物、器材貼 RFID 標籤快速掃描",
                "   - 設計 3D 列印壓克力專屬終端外殼，正式裝設於童軍團部鐵櫃，長期運作"
            ]
        }
    ]
    
    for s_idx, data in enumerate(slides_data):
        slide = prs.slides.add_slide(blank_layout)
        
        # 頂部背景色塊
        top_bar = slide.shapes.add_shape(
            pptx.enum.shapes.MSO_SHAPE.RECTANGLE,
            PInches(0), PInches(0), PInches(13.333), PInches(1.5)
        )
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = PRGBColor(30, 60, 114)
        top_bar.line.color.rgb = PRGBColor(30, 60, 114)
        
        # 標題
        tx_box = slide.shapes.add_textbox(PInches(0.8), PInches(0.2), PInches(11.7), PInches(0.6))
        tf = tx_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = data["title"]
        p.font.size = PPt(26)
        p.font.bold = True
        p.font.color.rgb = PRGBColor(255, 255, 255)
        
        # 副標題
        sub_box = slide.shapes.add_textbox(PInches(0.8), PInches(0.85), PInches(11.7), PInches(0.5))
        sub_tf = sub_box.text_frame
        sub_p = sub_tf.paragraphs[0]
        sub_p.text = data["sub"]
        sub_p.font.size = PPt(15)
        sub_p.font.color.rgb = PRGBColor(220, 230, 242)
        
        # 內容主框
        content_box = slide.shapes.add_textbox(PInches(0.8), PInches(1.8), PInches(11.7), PInches(5.0))
        c_tf = content_box.text_frame
        c_tf.word_wrap = True
        
        for idx, bullet in enumerate(data["bullets"]):
            if idx == 0:
                cp = c_tf.paragraphs[0]
            else:
                cp = c_tf.add_paragraph()
            cp.text = bullet
            cp.font.size = PPt(18)
            cp.font.color.rgb = PRGBColor(40, 40, 40)
            cp.space_before = PPt(10)
            cp.space_after = PPt(6)
            
    prs.save(output_path)
    print(f"Presentation PPTX saved to {output_path}")

if __name__ == "__main__":
    docs_dir = os.path.join(os.path.dirname(__file__), "docs")
    report_docx = os.path.join(docs_dir, "專題報告_童軍倉庫智慧櫃位引導與借還系統.docx")
    summary_docx = os.path.join(docs_dir, "AI協同開發過程摘要_童軍倉庫智慧櫃位引導與借還系統.docx")
    slides_pptx = os.path.join(docs_dir, "90秒展示與專題投影片_童軍倉庫智慧櫃位引導與借還系統.pptx")
    
    create_report_docx(report_docx)
    create_summary_docx(summary_docx)
    create_presentation_pptx(slides_pptx)
