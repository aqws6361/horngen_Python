def DecimalToAZ52(dec):
    if dec == 0:
        return 'a'

    chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
    base = 52
    result = []

    while dec > 0:
        remainder = dec % base
        result.append(chars[remainder])
        dec = dec // base

    return ''.join(result[::-1])


def AZ52ToDecimal(az52):
    chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
    base = 52
    result = 0

    for char in az52:
        result = result * base + chars.index(char)

    return result


# 預期執行結果
DecimalNumber = 4073126405
AZ52Number = "kLdXLX"

print(DecimalToAZ52(DecimalNumber))  # kLdXLX
print(DecimalToAZ52(0))  # a
print(DecimalToAZ52(26))  # A

print(AZ52ToDecimal(AZ52Number))  # 4073126405
print("測試github")