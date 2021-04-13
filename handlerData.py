import pandas as pd
from natsort import natsorted
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use('Agg')  # 解决 RuntimeError: main thread is not in main loop
import datetime
import json
from docx.shared import Cm
from docxtpl import DocxTemplate, InlineImage
import os
import comtypes.client

class RiskPartsAnalysis():

    def __init__(self,BOM_filename,part_status_filename):
        file_path = BOM_filename
        self.risk_data_path = part_status_filename
        (path, filename) = os.path.split(file_path)
        self.project = filename.split('.')[0]
        self.df = pd.read_excel(file_path)
        self.df.fillna('No Value in Excel',inplace = True)
        self.obj = {'date': datetime.date.today()}  # 当天时间
        self.obj['project_name'] = self.project

    def handler_BOM(self):

        all_df = self.df.groupby(['Component number', 'MFG Name', 'Mfg Part Number'])['Sub item name'].apply(lambda x: ' '.join(x)).to_frame().reset_index()

        all_df['Sub item name'] = all_df['Sub item name'].apply(lambda i: ' '.join(natsorted(i.split(' '))))  # 排序，如C33 C35 C23 重新C23 C33 C35

        duplicate_con = all_df.duplicated(subset=['Component number'], keep=False)  # 根据component判断，是否重复，全部标识为False

        unique_df = all_df[~duplicate_con].reset_index(drop=True)  # 单一Source的Component

        multiple_df = all_df[duplicate_con].reset_index(drop=True)  # 多Source的Component

        self.obj['total_count'] = len(all_df['Component number'].unique())
        self.obj['unique_count'] = len(unique_df)
        self.obj['multiple_count'] = len(multiple_df['Component number'].unique())
        self.obj['unique_source'] = json.loads(unique_df.to_json(orient='records'))  ##转成json格式
        self.obj['unique_source_lables'] = unique_df.columns.to_list()

        return all_df,unique_df, multiple_df

    def get_EOL_start(self,df):
        start_eol_con = df['Mfg Part Number'].str.startswith('EOL')
        eol_start_df = df[start_eol_con]
        return eol_start_df.assign(Status=lambda i: 'EOL').reset_index(drop=True)


    def merge_data(self,left_df):
        usecols = ['Component category', 'EOL MPN', 'Status']
        full_df = pd.read_excel(self.risk_data_path,sheet_name='Full EOL List', usecols=usecols)
        merge_df = pd.merge(left = left_df, right=full_df, left_on='Mfg Part Number', right_on='EOL MPN', how='left')
        merge_df = merge_df[~merge_df['Status'].isna()].reset_index(drop=True)  # 过滤掉状态为空的
        merge_df['Status'] = merge_df['Status'].apply(lambda x: 'EOL' if (('EOL' in x) or ('Decline' in x ) or (' Phase out part' in x ))  else x)  # 包含EOL，就直接用EOL替代
        merge_df.drop(['EOL MPN'], inplace=True, axis=1)  # 去除EOL MPN列
        merge_df = merge_df.drop_duplicates(['Component number', 'MFG Name', 'Mfg Part Number'],
                                            keep='first').reset_index(drop=True)  # 去重，保留第一个
        merge_df.sort_values('Status', inplace=True, ignore_index=True)  # 排序
        return merge_df

    def main(self):

        all_df,unique_df,multiple_df = self.handler_BOM()
        merge_df = self.merge_data(all_df)
        multiple_eol_df = self.get_EOL_start(multiple_df)
        multiple_eol_df['Mfg Part Number'] = multiple_eol_df['Mfg Part Number'].apply(lambda x: x.strip('EOL '))
        concat_df = pd.concat([merge_df, multiple_eol_df], axis=0, ignore_index=True).fillna('')
        concat_df.sort_values('Status', inplace=True, ignore_index=True)  # 排序

        self.get_risk_parts(unique_df)

        df = self.handler_dict(concat_df)

        self.save_file(df)

    def handler_dict(self,df):

        self.obj['p_status'] = json.loads(df.to_json(orient='records'))
        con = df.duplicated(subset=['Component number', 'Status'], keep='first')
        df = df[~con].reset_index(drop=True)
        self.obj['eol_count'] = len(df[df['Status'] == 'EOL'])
        self.obj['shortage_count'] = len(df[df['Status'] == 'Shortage'])
        self.obj['nrnd_count'] = len(df[df['Status'] == 'NRND']['Status'])
        self.obj['risk_count'] = self.obj['eol_count'] + self.obj['shortage_count'] + self.obj['nrnd_count']
        self.obj['active_count'] = self.obj['total_count'] - self.obj['risk_count']
        return df

    def get_risk_parts(self,df):
        risk_df = self.merge_data(df)
        self.obj['col_labels'] = risk_df.columns.to_list()
        self.obj['risk_parts'] = json.loads(risk_df.to_json(orient='records'))
        self.obj['high_risk_count'] = len(risk_df)
        ## 单一source，带状态
        unique_status_df = df.merge(risk_df,
                                           on=['Component number', 'MFG Name', 'Mfg Part Number', 'Sub item name'],
                                           how='outer').fillna('')
        self.obj['unique_parts'] = json.loads(unique_status_df.to_json(orient='records'))

    def my_fmt(self,x):
        return '{:0.0f}%\n[{:.0f}]'.format(x, self.obj['total_count'] * x / 100)

    def save_file(self,df):

        status_df = df['Status'].value_counts().rename_axis('Status').reset_index(name='Counts')
        status_df = status_df.append(pd.DataFrame({'Status': 'Active', 'Counts': [self.obj['active_count']]}), ignore_index=True)
        status_df.set_index('Status', inplace=True)  # 设置index

        self.plot_pie_graph(status_df)

        self.insert_img()

        self.word2pdf()

    def insert_img(self):

        doc = DocxTemplate(f'template.docx')
        img_size = Cm(9)  # sets the size of the image
        pic = InlineImage(doc, f'{self.project}.png', img_size)
        self.obj['chart_pic'] = pic
        doc.render(self.obj)
        print('进行保存word文件...')
        doc.save(f'{self.project}.docx')

    def plot_pie_graph(self, df):
        print('进行饼图文件生成...')
        colours = {'EOL': '#FF3333',
                   'NRND': '#33A8FF',
                   'Shortage': '#FFE033',
                   'Active': '#73D637'}
        explode = [0.1] + [0] * (len(df) -1)
        plt.rcParams["figure.dpi"] = 140  # 设置dpi
        p = df.plot.pie(y='Counts',
                       legend=False,  # 不显示图例
                       figsize=(6, 6),
                       radius=1.2,  # 设置半径，默认1
                       autopct=self.my_fmt,  # 设置显示格式
                       explode=explode,
                       counterclock=True,  # 百分比顺时针增大
                       textprops={'fontsize': 12, 'color': 'black', 'font': 'Arial', 'fontweight': 'bold'},
                       # 字体设置，字号，颜色，字体，粗体'fontweight':'bold'
                       startangle=90,
                       colors=[colours[key] for key in df.index.to_list()],  # 根据方块设置颜色
                       subplots=True,  # subplots=True 为了后续保存图片，不知道为啥！
                       pctdistance=0.8,  # 设置比例位置
                       labeldistance=1.05  # 设置label位置，图中EOL等
                       )
        p[0].yaxis.set_visible(False) #隐藏 y坐标
        # p[0].legend(loc="lower right",bbox_to_anchor=(1.17,0),fontsize=10)
        p[0].set_title('Proportion of Parts Status Counts', pad=28,
                       fontdict={'fontsize': 18, 'fontweight': 'bold', 'color': '#004d7c'})  # pad 标题间距
        p[0].get_figure().savefig(f'{self.project}.png', dpi=1200)
        plt.close()

    def word2pdf(self):
        print('进行word转PDF....')

        wdFormatPDF = 17

        in_file = os.path.abspath(f'{self.project}.docx')
        out_file = os.path.abspath(f'{self.project}.pdf')

        word = comtypes.client.CreateObject('Word.Application')
        doc = word.Documents.Open(in_file)
        doc.SaveAs(out_file, FileFormat=wdFormatPDF)
        doc.Close()
        word.Quit()


if __name__ == '__main__':
    BOM_filename = r'C:\Users\h290602\Desktop\50122789-0131602.xlsx'
    part_status_filename = r'C:\Users\h290602\Desktop\EE parts risk analyze Bom Scrub.xlsm'
    RiskPartsAnalysis(BOM_filename,part_status_filename).main()
