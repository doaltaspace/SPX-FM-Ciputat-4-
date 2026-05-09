import zipfile
import re
with zipfile.ZipFile('Retur Pergudanan - Ciputat 4 First Mile.xlsx', 'r') as zh:
    # First search sharedStrings.xml
    try:
        ss_content = zh.read('xl/sharedStrings.xml').decode('utf-8')
        strings = re.findall(r'<t>(.*?)</t>', ss_content)
        print("Shared Strings:")
        for idx, s in enumerate(strings[:30]):
            print(f"{idx}: {s}")
    except Exception as e:
        print("Error reading shared strings", e)
