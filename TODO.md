# TODO: בניית אתר Streamlit לדוח גיול

## שלב 1: config/constants.py
- [x] פישוט קבועים (ZERO_THRESHOLD, HEADERS, SHEET_NAME, SECTION_NAMES)

## שלב 2: parsers/bkmv_parser.py
- [x] שכתוב מלא - fixed-width CP862 parser
- [x] parse_b11_records() - כרטיסיות חשבון
- [x] parse_b1_transactions() - תנועות יומן

## שלב 3: parsers/pdf_parser.py
- [x] קובץ חדש - פרסור דוח 331
- [x] parse_pdf_section() - חילוץ חשבונות לפי קוד סעיף

## שלב 4: processors/aging_calculator.py
- [x] שכתוב - חישוב גיול (סכימה אחורה)
- [x] process_accounts() - עיבוד כל החשבונות

## שלב 5: excel/excel_generator.py
- [x] שכתוב - גיליון בודד מעוצב RTL
- [x] generate_excel() → bytes

## שלב 6: utils/file_utils.py
- [x] פישוט - extract_zip, find_bkmvdata, cleanup

## שלב 7: app.py - ממשק Streamlit
- [x] RTL CSS + page config
- [x] 3 שדות קלט (ZIP, PDF, קוד סעיף)
- [x] פס התקדמות + הודעות בעברית
- [x] תצוגה מקדימה + סיכום
- [x] כפתור הורדת Excel

## שלב 8: ניקוי + requirements
- [x] requirements.txt מינימלי
- [x] מחיקת קבצים ישנים
- [ ] בדיקה סופית
