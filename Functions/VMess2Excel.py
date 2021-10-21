import os
import json
import xlwt
import base64
import requests


def Base64Decode(content):
    """Base64 解密"""
    return base64.b64decode(str.encode(content)).decode()


def LoadSubFile(localFile):
    """加载本地经 Base64 加密的订阅文件，返回节点列表"""
    with open(localFile) as f:
        subContent = f.read()
        VMessContent = Base64Decode(subContent)
        VMessList = VMessContent.replace('\r', '').split('\n')
        nodesDictList = list(map(lambda VMess: Base64Decode(VMess.replace('vmess://', '')), VMessList))
        return nodesDictList


def LoadSubUrl(subUrl):
    """加密网络经 Base64 加密的订阅网址，返回节点列表"""
    getHeaders = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.75 Safari/537.36'}
    getResponse = requests.get(url=subUrl, headers=getHeaders)
    subContent = getResponse.content.decode()
    VMessContent = Base64Decode(subContent)
    VMessList = VMessContent.replace('\r', '').split('\n')
    nodesDictList = list(map(lambda x: Base64Decode(x.replace('vmess://', '')), VMessList))
    return nodesDictList


def SaveToExcel(nodesDictList, excelFile):
    """将节点字典列表保存到 Excel 文件"""
    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet = workbook.add_sheet('VMess')

    keyDict = {}
    for row, node in enumerate(nodesDictList, start=1):

        try:
            NodeDict = json.loads(node)
        except:
            continue

        for Key in NodeDict:
            if Key not in keyDict:
                keyDict[Key] = {'Col': len(keyDict), 'Width': 10}
            worksheet.write(row, keyDict[Key]['Col'], label=NodeDict[Key])
            if keyDict[Key]['Width'] < len(NodeDict[Key].encode('gbk')) + 2:
                keyDict[Key]['Width'] = len(NodeDict[Key].encode('gbk')) + 2

    for col, key in enumerate(keyDict):
        worksheet.write(0, col, label=key)
        worksheet.col(keyDict[key]['Col']).width = 256 * (keyDict[key]['Width'])

    if os.path.isfile(excelFile):
        os.remove(excelFile)
    workbook.save(excelFile)

    return True


if __name__ == '__main__':
    # NodesDictList = LoadSubFile('./sub2')
    # SaveToExcel(NodesDictList, './sub.xls')
    pass
