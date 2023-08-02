import requests
from lxml import etree
import pandas as pd
import time
import streamlit as st
import os
import base64
def hospital(hospital_name):
    url = 'https://zgylbx.com/index.php?m=content&c=index&a=lists&catid=106&steps=&search=1&pc_hash=&k1=0&k2=0&k3=0&title={hospital_name}'.format(hospital_name=hospital_name)
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36 Edg/115.0.1901.183',
    }
    response = requests.get(url, headers=headers)
    html = etree.HTML(response.text)
    grade = ''
    if len(html.xpath('//table/tr[@class=" tr-dt"]/td/text()')) >= 3:
        grade = html.xpath('//table/tr[@class=" tr-dt"]/td/text()')[2]
    return grade

def main():
    st.title('医院等级查询')
    # 使用Streamlit的文件上传组件，允许用户上传Excel文件
    uploaded_file = st.file_uploader("上传Excel文件", type=["xlsx"])

    if uploaded_file is not None:
        # 使用pandas读取上传的Excel文件，并指定格式引擎为'openpyxl'
        df = pd.read_excel(uploaded_file, engine='openpyxl')

        # 假设要读取的列名为'医院名称'，替换成你实际需要读取的列名
        column_name = '医院名称'
        # 读取指定列的数据，进行查询并更新等级
        for i in range(len(df)):
            title = df.loc[i, column_name]
            grade = hospital(title)
            st.write("医院：", title)
            st.write("等级：", grade)
            if grade:
                df.loc[i, '医院等级'] = grade
            time.sleep(2)
        # 显示DataFrame内容
        st.write(df)
        # 保存更新后的DataFrame到临时文件
        tmp_file_path = os.path.join(os.getcwd(), 'tmp_data.xlsx')
        df.to_excel(tmp_file_path, index=False)
        # 将临时文件转换为Base64编码
        with open(tmp_file_path, 'rb') as file:
            base64_data = base64.b64encode(file.read()).decode()
        # 以Markdown形式提供下载链接
        st.markdown(
            f"### [点击下载更新后的Excel文件](data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{base64_data})")
        # 删除临时文件
        os.remove(tmp_file_path)
        st.success("Excel文件已更新！")

if __name__ == '__main__':
    main()
