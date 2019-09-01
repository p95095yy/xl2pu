import zipfile
import xml.etree.ElementTree as ET

v = []

with zipfile.ZipFile('excel.xlsx', 'r') as zf:
    with zf.open('xl/sharedStrings.xml') as shareStrings:
        root = ET.fromstring(shareStrings.read())
        strings=[]
        for i in root:
            strings.append(i[0].text)

    with zf.open('xl/worksheets/sheet1.xml') as sheet1:
        root = ET.fromstring(sheet1.read())
        sheetData = root.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheetData')
        for row in sheetData:
            a = ['']*26
            num = 0;
            for c in row:
                num = ord(c.attrib['r'][0])-ord("A")
                a[num] = strings[int(c[0].text)];
            del a[num+1:]
            v.append([row.attrib['r'],a])

names = v[0][1]

print("@startuml")
for idx,line in enumerate(v[1:]):
    f = 0
    for i,elem in enumerate(line[1]):
        if(elem != ''):
            if(f == 0):
                print(names[i], elem, " ", end="")
                f=1
            else:
                print(names[i], ":", line[0]+")",elem)
print("@enduml")
