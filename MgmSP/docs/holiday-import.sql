-- =========================================================
-- 台灣政府行事曆匯入腳本
--
-- 資料來源：政府資料開放平台 (data.gov.tw)
-- 搜尋：政府行政機關辦公日曆表
--
-- 使用方式：
-- 1. 從 https://data.gov.tw/dataset/14718 下載當年度 CSV
-- 2. 手動將假日資料轉換為以下 INSERT 語句格式
-- 3. 執行此腳本
-- =========================================================

-- 範例：2026 年假日資料（需根據實際行事曆更新）
-- 欄位說明：HolidayDate=日期, HolidayName=假日名稱, Year=年份, Source=來源, IsWorkday=1補班/0放假

-- 清除舊資料（可選）
-- DELETE FROM jHolidays WHERE Year = 2026

-- === 2026 年國定假日（資料來源：行政院人事行政總處）===
-- 資料下載：https://data.gov.tw/dataset/14718

-- 元旦
INSERT INTO jHolidays (HolidayDate, HolidayName, Year, Source, IsWorkday) VALUES ('2026-01-01', '開國紀念日', 2026, 'Government', 0);

-- 農曆春節
INSERT INTO jHolidays (HolidayDate, HolidayName, Year, Source, IsWorkday) VALUES ('2026-02-14', '春節調整放假', 2026, 'Government', 0);
INSERT INTO jHolidays (HolidayDate, HolidayName, Year, Source, IsWorkday) VALUES ('2026-02-15', '小年夜', 2026, 'Government', 0);
INSERT INTO jHolidays (HolidayDate, HolidayName, Year, Source, IsWorkday) VALUES ('2026-02-16', '農曆除夕', 2026, 'Government', 0);
INSERT INTO jHolidays (HolidayDate, HolidayName, Year, Source, IsWorkday) VALUES ('2026-02-17', '春節', 2026, 'Government', 0);
INSERT INTO jHolidays (HolidayDate, HolidayName, Year, Source, IsWorkday) VALUES ('2026-02-18', '春節', 2026, 'Government', 0);
INSERT INTO jHolidays (HolidayDate, HolidayName, Year, Source, IsWorkday) VALUES ('2026-02-19', '春節', 2026, 'Government', 0);

-- 補班日（週六/週五需上班）
INSERT INTO jHolidays (HolidayDate, HolidayName, Year, Source, IsWorkday) VALUES ('2026-02-20', '春節補班', 2026, 'Government', 1);
INSERT INTO jHolidays (HolidayDate, HolidayName, Year, Source, IsWorkday) VALUES ('2026-02-27', '和平紀念日補班', 2026, 'Government', 1);

-- 和平紀念日
INSERT INTO jHolidays (HolidayDate, HolidayName, Year, Source, IsWorkday) VALUES ('2026-02-28', '和平紀念日', 2026, 'Government', 0);

-- 兒童節/清明節
INSERT INTO jHolidays (HolidayDate, HolidayName, Year, Source, IsWorkday) VALUES ('2026-04-03', '兒童節調整放假', 2026, 'Government', 0);
INSERT INTO jHolidays (HolidayDate, HolidayName, Year, Source, IsWorkday) VALUES ('2026-04-04', '兒童節', 2026, 'Government', 0);
INSERT INTO jHolidays (HolidayDate, HolidayName, Year, Source, IsWorkday) VALUES ('2026-04-05', '清明節', 2026, 'Government', 0);
INSERT INTO jHolidays (HolidayDate, HolidayName, Year, Source, IsWorkday) VALUES ('2026-04-06', '清明節調整放假', 2026, 'Government', 0);

-- 補班日
INSERT INTO jHolidays (HolidayDate, HolidayName, Year, Source, IsWorkday) VALUES ('2026-03-28', '清明節補班', 2026, 'Government', 1);

-- 勞動節
INSERT INTO jHolidays (HolidayDate, HolidayName, Year, Source, IsWorkday) VALUES ('2026-05-01', '勞動節', 2026, 'Government', 0);

-- 端午節
INSERT INTO jHolidays (HolidayDate, HolidayName, Year, Source, IsWorkday) VALUES ('2026-06-19', '端午節', 2026, 'Government', 0);

-- 中秋節
INSERT INTO jHolidays (HolidayDate, HolidayName, Year, Source, IsWorkday) VALUES ('2026-09-25', '中秋節', 2026, 'Government', 0);

-- 國慶日
INSERT INTO jHolidays (HolidayDate, HolidayName, Year, Source, IsWorkday) VALUES ('2026-10-09', '國慶日調整放假', 2026, 'Government', 0);
INSERT INTO jHolidays (HolidayDate, HolidayName, Year, Source, IsWorkday) VALUES ('2026-10-10', '國慶日', 2026, 'Government', 0);

-- 補班日
INSERT INTO jHolidays (HolidayDate, HolidayName, Year, Source, IsWorkday) VALUES ('2026-10-17', '國慶日補班', 2026, 'Government', 1);

-- === 公司特殊假日（可自行新增）===
-- INSERT INTO jHolidays (HolidayDate, HolidayName, Year, Source, IsWorkday) VALUES ('2026-12-31', '公司年假', 2026, 'Company', 0);

-- =========================================================
-- 查詢確認
-- =========================================================
SELECT * FROM jHolidays WHERE Year = 2026 ORDER BY HolidayDate;

-- =========================================================
-- 快速查詢函數測試
-- =========================================================
-- 檢查某天是否在假日表中
-- SELECT * FROM jHolidays WHERE HolidayDate = '2026-01-01';
