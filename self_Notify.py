import test_secret
import requests

token = test_secret.line_Notify
message = "123"
# HTTP 標頭參數與資料
headers = {"Authorization": "Bearer " + token}
data = {'message': message}

response = requests.post("https://notify-api.line.me/api/notify",
                         headers=headers, data=data)