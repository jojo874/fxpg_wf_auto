


str1="支持SSL 64位块大小的密码套件（SWEET32）"

str2="gSOAP 安全漏洞(CVE-2017-9765)"

if "CVE" in str2:
    print("有")
else:
    print(str2)


for i in range(0,3):
    print(i)

unit="奎文区大数"
attrList_shi = ["高密市", "青州市", "诸城市", "寿光市", "安丘市", "昌邑市"]
attrList_qu = ["奎文区", "潍城区", "房子区", "寒亭区", "高新区", "滨海区", "保税区", "峡山区", "经济开发区"]
attrList_xian = ["临朐县", "昌乐县"]
attr="单位"
for i in attrList_shi:
    if i in unit:
        attr="市"
for i in attrList_qu:
    if i in unit:
        attr="区"
for i in attrList_xian:
    if i in unit:
        attr="县"

print(attr)
