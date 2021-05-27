# Estate

## 1. Parsing.py

### 1-1. 기능설명

네이버 부동산에서 3개의 동(논현동, 고잔동, 남촌동)에서 최신 매물들 정보를 google spread sheet로 가져오게 된다.

이때 고유 매물번호도 함께 저장하여 이후에 중복된 매물인지 확인한다.

가져오는 정보는 총 10개로
- 매물번호
- 업로드 날짜
- 매매금액
- 대지/연면적
- 지상층/지하층
- 입주가능일
- 현재용도
- 사용승인일
- 매물설명
- 링크
이다.

### 1-2. 기술설명
 
selenium과 chrome webdriver를 사용하여 웹크롤링을 사용하였고, gspread를 통해 google spread sheet에 접근하여 데이터를 저장하였다.


## 2. matching_data.py

### 2-1. 기능설명

1에서 가져온 데이터들을 바탕으로 탐색을 원하는 데이터 NO의 시작점과 끝점을 입력하면 기존 시트에 정보와 대조해서 사용승인일이 일치하거나, 대지/연면적 크기의 오차가 0.01% 이내라면 예상매물 리스트에 추가하여 gogole spread sheet에 추가한다. 이때 해당 상호명과 시트내의 링크를 동시에 추가한다.

### 2-2. 기술설명

파싱으로 가져온 데이터들의 사용승인일과 대지연면적 데이터를 가져오고, 기존 매물데이터들의 사용승인일과 대지연면적 데이터를 가져와서 일치여부를 확인한 후 2차원 배열에 해당 정보들을 담아 마지막에 google spread sheet에 업데이트한다. 



## Chromedriver

Chromedriver version과 사용중인 chrome의 version이 일치해야한다.
