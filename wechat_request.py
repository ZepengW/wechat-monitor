import json
import requests
import time
import random
from error import *
import sys
from csv_io import write_csv
from requests.packages.urllib3.exceptions import InsecureRequestWarning

# 关闭报警
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)
url = "https://mp.weixin.qq.com/cgi-bin/appmsg"

headers = {
    "Cookie": 'ptui_loginuin=1371946804; RK=htes1sHRPN; ptcz=a3d21d75959b623656aaf3d38ca3bcb81becb27a3fe5b88388508e5d255e9b5e; tvfe_boss_uuid=d6c5598666965073; pgv_pvid=5496426140; o_cookie=1371946804; pac_uid=1_1371946804; ua_id=7W6H0n1CFOdFthOwAAAAAKOdm07Oq_oHrMgVp6s3EyU=; wxuin=77647480655089; pgv_info=ssid=s9211065855; rewardsn=; wxtokenkey=777; uuid=cc9f83622c77ccd40542a1420db10cef; cert=h6PV27b8sOnD9zh1Cb2KdxbNA0iBFNsI; sig=h01287ce8c797385c72256ba300b6a846b1be64cb454c04eaf2d30d057e8458960d5cf84381652e2958; data_bizuin=3890910163; bizuin=3890910163; master_user=gh_a67b7e12998e; master_sid=ZTZmbkVzeWZzN3podEpITDdaczBaUnloSGFFVUNaM2RWWFRzeHZGb3Z1YnJBMUhDZng4MjY5QVZqbkVxTG1Idkg5cHB3ZEREZ1lIYldtZGtYYVlaVFRaeG5DeXk0WXJkVlJIa0xVaWpSUGNicGxwdzlzS3VFTkFkcG5nY0FkbGk5S3Q1UnVWdlZia0VjRDA5; master_ticket=e357c81aa4f4f3d52ca1ae1bcc8e1e22; data_ticket=7vXNsZ2ViFPIEq4EZ0vxGTHE63KBn4W0XhUw5Rt6Dsf+JthuvOaz81yqg+Cn+1zz; rand_info=CAESID0hnHKqC2APtlSjpvdwyscJEv7ir6LJ+tL+XRA+BhBw; slave_bizuin=3890910163; slave_user=gh_a67b7e12998e; slave_sid=OXU3dUNFTFpEM2YyektqY0VSV291YWY4dHJRYmJ1UXZWZFpibTkyUU5zSnlRcjdlek9TdXNYdVdZc3BmYmxvNzN4YUQxM05WUUU0OXdxVDJmeXRtZHMzMnYycEV2cDIwUXc1Zjl4Rk9kRHUwZ2d1M1hwMU9xZXlkQnJjUXk2WDNKeTZ4NU5kQ0tqOWZIcGVS',
    "User-Agent": 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.0.0 Safari/537.36'
}



def get_all_articles(fakeid, token, name=''):
    count = 10
    article_l = []
    times_try = 5  # 尝试次数
    while(times_try > 0):
        try:
            time.sleep(random.randint(3,15))
            return_articles = request_once(fakeid, token, len(article_l), count)
            #print(return_articles)
        except WechatNetworkError:
            time.sleep(3600)
            times_try -= 1
            continue
        except:
            print("Unexpected error:", sys.exc_info()[0])
            sys.exit()
        if len(return_articles) == 0:
            #搜索完成
            print(f'[{name if name == "" else fakeid}] search finish')
            return article_l
        article_l += return_articles
        times_try = 10
        print(f'already search articles: {len(article_l)}')
    print('[WARNING] 访问失败次数达到上限')
    return article_l
    
    


#
def request_once(fakeid, token, begin_id = 0, count = 5):
    '''request articles list once
    params:
        fakeid: related to account
        token: related to account
        begin_id: begin to search from [begin_id]th article
        count: number of articles returned
    '''
    params = {
        "action": "list_ex",
        "begin": f"{begin_id}",
        "count": f"{count}",
        "fakeid":  fakeid,
        "type": "9",
        "token": token,
        "lang": "zh_CN",
        "f": "json",
        "ajax": "1"
        }  
    
    resp = requests.get(url, headers=headers, params = params, verify=False)
    # 微信流量控制, 退出
    if resp.json()['base_resp']['ret'] == 200013:
        #print("frequencey control, stop at {}".format(str(begin)))
        raise WechatNetworkError()
    msg = resp.json()
    if not "app_msg_list" in msg:
        raise WechatUndifineError()
    if len(msg["app_msg_list"]) == 0:
        return []
    article_l = []
    for item in msg["app_msg_list"]:
        time_str = time.strftime("%Y-%m-%d %H:%M:%S",time.localtime(item['create_time']))
        title_str = item['title']
        link_str = item['link']
        aid = str(item["aid"])
        article_l.append({
            'time': time_str,
            'title': title_str,
            'link': link_str,
            'aid': aid
        })
    return article_l


if __name__ == '__main__':
    article_l = get_all_articles('MjM5MDM5MzEyMQ==', '1804298720')
    write_csv('./rizhao.csv', article_l)