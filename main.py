import pandas
import re


class ComparaFun():

    # 初始化一些设置
    def __init__(self) -> None:
        pandas.set_option('display.precision', 2)
        pandas.set_option('display.unicode.east_asian_width', True)
        pandas.set_option('display.unicode.ambiguous_as_wide', True)

        self.splitList = ['=', '+', '-', '*', '/', '(', ')']
        self.dfFun = None
        self.outputPath = None
        self.inputPath = None

    # 自述

    def __repr__(self) -> str:
        print('首先实例化一个对象如F = ComparaFun()')
        print('第一步，通过setInputFile属性，指定指标和公式所在的Excel文件')
        print('第二步，通过setOutputFile属性，指定输出文件夹')
        print('第三步，通过run方法执行主程序')
        return ''

    @property
    def setOutputFile(self):
        return self.outputPath

    @setOutputFile.setter
    def setOutputFile(self, path: str):
        self.outputPath = path

    @property
    def setInputFile(self):
        return self.inputPath

    @setInputFile.setter
    def setInputFile(self, path):
        self.inputPath = path

        self.dfPara = pandas.read_excel(self.inputPath, sheet_name='指标')
        self.dfPara.set_index('序列', inplace=True)
        self.dfPara = self.dfPara.to_dict()['指标']

        self.dfFun = pandas.read_excel(self.inputPath, sheet_name='公式')
        self.dfFun['验证'] = [0]*len(self.dfFun)

    def SplitFun(self, userInput, splitStr: str) -> list:
        '''SplitFun用于将输入的字符串基于分隔符拆解，并输出一个保留了分隔符的列表'''
        outList = []  # 存储结果
        try:
            if isinstance(userInput, str):
                if userInput.find(splitStr) != -1:  # 如果字符串存在分隔符，则进行如下处理
                    userInput.replace(' ', '')
                    tmpList = userInput.split(splitStr)  # 按分隔符分割得到的列表
                    for i in tmpList:
                        outList.append(i)  # 将分割后的元素一个一个写入结果表
                        outList.append(splitStr)  # 将分隔符写入结果表
                    outList.pop(-1)
                    return outList
                else:  # 如果不存在则直接返回
                    return userInput
            elif isinstance(userInput, list):
                outList = []
                for i in userInput:
                    if isinstance(self.SplitFun(i, splitStr), list):
                        for j in self.SplitFun(i, splitStr):
                            outList.append(j)
                    else:
                        outList.append(i)
                return outList
            else:
                raise ValueError
        except ValueError:
            print('未知的输入类型')

    def ReplaceFun(self, userInput):
        '''ReplaceFun是利用正则表达式实现指标替换的方法'''
        ...

    def main(self, userInput):
        for i in self.splitList:
            userInput = self.SplitFun(userInput, i)
        userInput = [i for i in userInput if i != '']
        for i in range(len(userInput)):
            if userInput[i].startswith('$'):
                userInput[i] = self.dfPara[userInput[i]]
        return ''.join(userInput)

    def run(self):
        for i in range(len(self.dfFun)):
            self.dfFun.iloc[i, 2] = self.main(self.dfFun.iloc[i, 0])
        self.dfFun.to_excel(self.outputPath + r'\输出结果.xlsx')

# from pycomparafun import *
# F = ComparaFun()
# F.setInputFile = r"C:\Users\zhhon\Desktop\新建 Microsoft Excel 工作表.xlsx"
# F.setOutputFile = r"C:\Users\zhhon\Desktop"
# F.run()
