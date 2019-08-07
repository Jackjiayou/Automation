import time
from concurrent.futures import ThreadPoolExecutor  # 线程池模块
import os

def get(url):
    print('GET %s' % url)
    #response = requests.get(url)  # 下载页面
    time.sleep(3)  # 模拟网络延时
    return {'url': url, 'content':"hahahahah"}  # 页面地址和页面内容

def parse(res):
    res = res.result()  # !取到res结果 【回调函数】带参数需要这样
    print('%s res is %s' % (res['url'], len(res['content'])))

if __name__ == '__main__':
    #print( os.environ['HOME'])
    #print (os.path.expandvars('$HOME'))
    #print (os.path.expanduser('~'))
    urls = {
        'http://www.baidu.com',
        'http://www.360.com',
        'http://www.iqiyi.com'
    }

    pool = ThreadPoolExecutor(2)
    for i in urls:
        pool.submit(get, i).add_done_callback(parse)  # 【回调函数】执行完线程后，跟一个函数 

