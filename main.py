from Functions.VMess2Excel import *
from Functions.Excel2VMess import *

if __name__ == '__main__':
    # 示例：
    # 将多个机场的订阅合并并转换保存至 nodes.xls
    # 此时可以打开 nodes.xls 用各种方法批量编辑节点
    # 编辑完成后按回车将 nodes.xls 转换为订阅文件 sub
    # 然后就可以把 sub 订阅文件上传至网络中，并用网址订阅
    # 注：暂未开发异常处理，以及更多应用场景和方法自行开发

    subUrls = [
        'http://xxx.xxx/sub1',
        'http://xxx.xxx/sub2',
        'http://xxx.xxx/sub3'
    ]

    nodesDictList = []
    for subUrl in subUrls:
        print(f'获取：{subUrl}')
        nodesDictList += LoadSubUrl(subUrl)
    SaveToExcel(nodesDictList, './nodes.xls')

    input('回车后将 nodes.xls 转换为 sub 订阅文件')

    newNodesDictList = ReadExcel('./nodes.xls')
    SaveToSubFile(newNodesDictList, './sub')

    pass
