import requests

# LINE Notify 權杖
token = 'dLF5rbxPQOTzjO2QaaleksNHRL5qzaywjFQZl6LbHVR'

# 要發送的訊息
message = 'http://www.horngenplastic.com/'

# HTTP 標頭參數與資料
headers = { "Authorization": "Bearer " + token }
data = { 'message': message }

# 以 requests 發送 POST 請求
requests.post("https://notify-api.line.me/api/notify",
    headers = headers, data = data)