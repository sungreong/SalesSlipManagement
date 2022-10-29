
결제 영수증 관리하는 코드

# 수정은 언제나 환영합니다.


# API KET 필요

https://www.data.go.kr/data/15012690/openapi.do

위의 링크를 통해 API 키가 있어야지, 휴일을 확인할 수 있음.

# 이미지 형식

`yyyymmdd+[tag_name]+[금액]+[특이사항].PNG`

## 예시
- 20221004+점심+4_000
- 20221004+점심+4_000+[인원2명]

# 실행 코드

```
python main_v2.py
```

# 실행 화면

![화면](./output.PNG)

# 최신 코드
main_v2.py

# EXE 파일 만들기
pyinstaller --noconsole --onefile --icon=agilesoda.ico main.py
pyinstaller --noconsole --onefile --icon=agilesoda.ico main_v2.py

# NEXT

- [ ] 기존에 다른 분께서 만든 코드 결과물 포맷 반영 
- [ ] 맥에서 실행할 수 있는 파일 만들기
