import json
import xlrd
import base64


def Base64Encode(content):
    """Base64 加密"""
    return base64.b64encode(content.encode()).decode()


def ReadExcel(excelFile):
    """读取 Excel 文件，返回节点字典列表"""
    workbook = xlrd.open_workbook(excelFile)
    sheet = workbook.sheets()[0]  # 仅读取第一个表

    nodesDictList = []
    for row in range(1, sheet.nrows):
        nodeDict = {}
        for Col in range(0, sheet.ncols):
            nodeDict[sheet.row(0)[Col].value] = sheet.row(row)[Col].value
        nodesDictList.append(nodeDict)

    return nodesDictList


def SaveToSubFile(nodesDictList, subFile):
    """将节点字典列表保存到订阅文件"""
    VMessList = list(map(lambda NodeDict: 'vmess://' + Base64Encode(json.dumps(NodeDict)), nodesDictList))
    VMessContent = '\n'.join(VMessList)
    subContent = Base64Encode(VMessContent)

    with open(subFile, 'w') as f:
        f.write(subContent)

    return True


if __name__ == '__main__':
    # nodesDictList = ReadExcel('./sub.xls')
    # SaveToSubFile(nodesDictList, './sub')
    pass
