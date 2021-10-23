import os
import json
import xlwt
import base64
import requests


def Base64Decode(content):
    """Base64 解密"""
    return base64.b64decode(str.encode(content)).decode()


def VMessFilter(NodeList):
    """过滤非 VMess 节点"""
    VMessList = []
    for node in NodeList:
        if 'vmess://' in node:
            VMessList.append(node)
        else:
            continue
    return VMessList


def LoadSubFile(localFile):
    """加载本地经 Base64 加密的订阅文件，返回节点列表"""
    with open(localFile) as f:
        subContent = f.read()
        VMessContent = Base64Decode(subContent)
        NodeList = VMessContent.replace('\r', '').split('\n')
        VMessList = VMessFilter(NodeList)
        nodesDictList = list(map(lambda x: Base64Decode(x.replace('vmess://', '')), VMessList))
        return nodesDictList


def LoadSubUrl(subUrl):
    """加密网络经 Base64 加密的订阅网址，返回节点列表"""
    getHeaders = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.75 Safari/537.36'}
    getResponse = requests.get(url=subUrl, headers=getHeaders)
    subContent = getResponse.content.decode()
    NodeContent = Base64Decode(subContent)
    NodeList = NodeContent.replace('\r', '').split('\n')
    VMessList = VMessFilter(NodeList)
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
            try:
                if keyDict[Key]['Width'] < len(str(NodeDict[Key]).encode('gbk')) + 2:
                    keyDict[Key]['Width'] = len(str(NodeDict[Key]).encode('gbk')) + 2
            except:
                pass

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
