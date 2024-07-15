import requests
import xml.etree.ElementTree as ET
import sapAPI
import token

# 設定API端點URL
url = "https://my303711-api.s4hana.ondemand.com:443/sap/opu/odata/sap/API_OPLACCTGDOCITEMCUBE_SRV/A_OperationalAcctgDocItemCube?$filter=AccountingDocument eq '2000000049' and  CompanyCode eq '6320'"

# 設定驗證資訊
username = sapAPI.username
password = sapAPI.password

# 設定HTTP請求標頭
headers = {
    "Accept": "application/xml",  # 設定接受的回應格式為XML
}

response = requests.get(url, auth=(username, password), headers=headers)

if response.status_code == 200:
    # 解析XML回應
    root = ET.fromstring(response.content)

    # 提取並列印回應中的所有元素名稱
    for element in root.iter():
        if element.tag == "{http://schemas.microsoft.com/ado/2007/08/dataservices}SupplierName":  # 替換為要提取的標籤
            print(f"({element.tag}: {element.text})")

            #處理與LINE API連動
            text = element.text
            message = text
            token = token.sapAPI
            headers1 = {"Authorization": "Bearer " + token}
            data1 = {'message': message}
            requests.post("https://notify-api.line.me/api/notify",
                          headers=headers1, data=data1)

else:
    print("發生錯誤，HTTP回應狀態碼:", response.status_code)


