## 👶🏻 Dolbom Reporter
- 반복되는 한글 파일 수정 업무를 자동화 합니다.
- .hwp -> .html 파일 변환 후 내용을 수정합니다.
- 특정 문서 형식에 의존성을 가집니다.

## 환경 설정
```
# python version 3.9
python -m venv .venv
source .venv/bin/activate

pip install -r requirements.txt

# Google Sheet ID 설정
echo "GOOGLE_SHEET_ID={}" > .env
```

## 실행
```
python main.py --child_name "김철수" --city "서울시" --teacher_name "박영희"
```

## 결과 확인
```
./output/doc.html 파일 클릭 후 브라우저에서 인쇄
```
