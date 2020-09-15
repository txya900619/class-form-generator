# Class Form Generator

a tool which deploy on google apps script can auto generate group class signup form and feedback form.

## How to use

1. `npm i -g @google/clasp`
2. `clasp login`
3. `clasp create`
4. `clasp push`

### classInformation should like that (for email)

```
【課程資訊】
課程名稱：初階 Web 應用 - Python 爬蟲-2
課程日期：03/23 (一) 19:00~21:00 ( 18:30 入場 )
課程地點：臺北科技大學 共同科館 412 教室

【課程資訊】
1. 解析網站原始碼
2. 介紹爬蟲所需要的套件
3. 利用套件抓取網頁資訊

【備 註】
1. 建議需要具備 Python 的基礎語言知識
2. 於教室內時，請勿飲食
3. 若需提前離開，請務必在當天告知工作人員
4. 非社員需繳交 50 元課程費用
5. 進入教室前建議戴上口罩
6. 體溫異常不得進場
```

### Success email will be like that

```
Hi,

感謝您報名 NPC 北科程式設計研究社 - 初階 Web 應用 - Python 爬蟲-2
在課程開始前一天，我們會再次寄信提醒您！

另外，由於資源寶貴，若臨時未能前來請您務必及早回信告知，讓備取學員得以遞補，謝謝您。

【課程資訊】
課程名稱：初階 Web 應用 - Python 爬蟲-2
課程日期：03/23 (一) 19:00~21:00 ( 18:30 入場 )
課程地點：臺北科技大學 共同科館 412 教室

【課程資訊】
1. 解析網站原始碼
2. 介紹爬蟲所需要的套件
3. 利用套件抓取網頁資訊

【備 註】
1. 建議需要具備 Python 的基礎語言知識
2. 於教室內時，請勿飲食
3. 若需提前離開，請務必在當天告知工作人員
4. 非社員需繳交 50 元課程費用
5. 進入教室前建議戴上口罩
6. 體溫異常不得進場

若有任何疑問，歡迎隨時連絡我們。\
期待在課程與您相見:)

Best regards,
NPC 北科程式設計研究社
```
