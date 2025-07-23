-- UPDATE script for table M_CORP_INFO (columns containing NAME, TEL, FAX, POST, ADDRESS, TANTOU, CREATE_USER, UPDATE_USER, FURIGANA)
-- Số dòng dữ liệu: 1

;WITH T AS (SELECT *, ROW_NUMBER() OVER (ORDER BY (SELECT 1)) AS rn FROM M_CORP_INFO)
UPDATE T SET CORP_NAME = N'エジソン商事株式会社', CORP_RYAKU_NAME = N'エジソン商事', CORP_FURIGANA = N'エジソンショウジテストヘンシュウ', KOUZA_NAME = N'エジソンショウジ（カ', CREATE_USER = N'edison', UPDATE_USER = N'鈴木　次郎', KOUZA_NAME_2 = N'エジソンショウジ（カ', KOUZA_NAME_3 = N'エジソンショウジ（カ', FURIKOMI_MOTO_KOUZA_NAME = N'' WHERE rn = 1;

