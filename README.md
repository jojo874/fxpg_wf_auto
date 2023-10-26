# fxpg_wf_auto
### 软件架构
?

### 安装教程
pip3 install -r requirements.txt

### 使用说明:
搭建mysql数据库,导入数据库结构3k.sql,导入相关数据,run.py修改数据库连接信息,python3 run.py, 脚本会根据数据库中的导入的数据自动生成附件1和附件2,然后生成文档到根目录

### 模版文件:
+ 2023_2.docx,根据情况自行修改.
+ 相关变量:
1. ph_unit 对应数据库中的unit字段
2. ph_x 对应数据库中同一unit的个数
3. ph_att 对应run.py脚本中的变量attr

### 输出文档演示
![image](https://github.com/jojo874/fxpg_wf_auto/assets/57160501/f7144013-5674-4a84-9d5a-6bf659674908)
![image](https://github.com/jojo874/fxpg_wf_auto/assets/57160501/4293004e-3541-4efa-973f-05f2a96a8218)
![image](https://github.com/jojo874/fxpg_wf_auto/assets/57160501/00f2bbd1-121d-4b80-aff6-93b4e0887502)

