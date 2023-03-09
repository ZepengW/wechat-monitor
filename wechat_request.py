import json
import requests
import time
import random
from error import *
import sys
from csv_io import write_csv
from requests.packages.urllib3.exceptions import InsecureRequestWarning
from urllib.parse import urlparse, parse_qs
import logging

# 关闭报警
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)
url = "https://mp.weixin.qq.com/cgi-bin/appmsg"


def cookies_str2dict(cookies_str, cookies_dict = None):
    '''Convert string format cookies to dict format'''
    if cookies_dict == None:
        res_dict = dict()
    else:
        res_dict = cookies_dict
    for cookies in cookies_str.split(';'):
        name, value = cookie.strip().split('=', 1)
        res_dict[name] = value
    return res_dict

def cookies_dict2str(cookies_dict):
    cookies_str = ''
    for name, value in cookies_dict.items():
        cookies_str += name + '=' + value + '; '
    return cookies_str.strip('; ')



def get_articles(fakeid, token, name='', exist_aid = None):
    """
    
    """
    count = 5
    article_l = []
    times_try = 5  # 尝试次数
    while(times_try > 0):
        try:
            time.sleep(random.randint(3,15))
            return_articles = request_once(fakeid, token, len(article_l), count)
            #print(return_articles)
        except WechatNetworkError:
            t_wait = 120
            times_try -= 1
            logging.warning(f'Network Error [Left try times: {times_try}], waiting {t_wait}s')
            time.sleep(t_wait)
            continue
        except:
            print("Unexpected error:", sys.exc_info()[0])
            sys.exit()
        if len(return_articles) == 0:
            #搜索完成
            print(f'[{name if name == "" else fakeid}] search finish')
            return article_l
        for item in return_articles:
            if end_aid in exist_aid:
                # search newest finish
                return article_l
            article_l.append(item)
        times_try = 5
        logging.info(f'already search articles: {len(article_l)}')
        break
    logging.warning('访问失败次数达到上限')
    return article_l
    
    


#
def request_once(fakeid, token, cookies_dict, begin_id = 0, count = 5):
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
    headers = {
        "User-Agent": 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.0.0 Safari/537.36',
        "Cookie": cookies_dict2str(cookies_dict=cookies_dict)
    }
    resp = requests.get(url, headers=headers, params = params, verify=False)
    
    # 微信流量控制, 退出
    if resp.json()['base_resp']['ret'] == 200013:
        logging.debug(f"frequencey control for fakeid[{fakeid}]")
        raise WechatNetworkError()
    msg = resp.json()
    if not "app_msg_list" in msg:
        logging.debug(msg)
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


def create_session(cookies_str):
    """ http resquest with session
    Args:
        cookies_str (_type_): cookies for logining without password
    """
    s = requests.session()
    s.headers.update(
        {
            "User-Agent": 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.0.0 Safari/537.36',
            "Cookie": cookies_str
        }
    )
    # login
    resp = s.get("https://mp.weixin.qq.com/")
    url_redict_home = resp.url
    # check if login success
    if not 'token=' in url_redict_home:
        logging.error('Login Fail, please update login cookies')
        sys.exit(0)
    # get token
    # 解析url
    parsed_url = urlparse(url_redict_home)
    # 获取查询参数
    query_params = parse_qs(parsed_url.query)
    # 获取token值
    token = query_params['token'][0]
    # login home
    resp = s.get(url_redict_home)
    if resp.status_code == 200:
        logging.info('Login Success')
        return s, token
    # login fail
    logging.error(f'Login Fail with Status Code [{resp.status_code}]')
    sys.exit(0)


def get_articles_session(session:requests.Session, fakeid, name='', exist_aid = []):
    """
    
    """
    count = 5
    article_l = []
    times_try = 5  # 尝试次数
    while(times_try > 0):
        try:
            time.sleep(random.randint(3,15))
            return_articles = request_once(session, fakeid, len(article_l), count)
            #print(return_articles)
        except WechatNetworkError:
            t_wait = 120
            times_try -= 1
            logging.warning(f'Network Error [Left try times: {times_try}], waiting {t_wait}s')
            time.sleep(t_wait)
            continue
        except:
            logging.error("Unexpected error:", sys.exc_info()[0])
            sys.exit()
        if len(return_articles) == 0:
            #搜索完成
            logging.debug(f'[{name if name == "" else fakeid}] search finish')
            return article_l
        for item in return_articles:
            if end_aid in exist_aid:
                # search newest finish
                return article_l
            article_l.append(item)
        times_try = 5
        logging.debug(f'already search articles: {len(article_l)}')
        break
    logging.warning('访问失败次数达到上限')
    return article_l

def request_once_session(session:requests.Session, fakeid, token, begin_id = 0, count = 5):
    """get article list from one fakeid with session mode

    Args:
        session (requests.Session): _description_
        fakeid (_type_): _description_
        token (_type_): _description_
        begin_id (int, optional): _description_. Defaults to 0.
        count (int, optional): _description_. Defaults to 5.

    Raises:
        WechatNetworkError: _description_
        WechatUndifineError: _description_

    Returns:
        _type_: _description_
    """
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
    resp = session.get(url, headers=headers, params = params, verify=False)
    
    # 微信流量控制, 退出
    if resp.json()['base_resp']['ret'] == 200013:
        logging.debug(f"frequencey control for fakeid[{fakeid}]")
        raise WechatNetworkError()
    msg = resp.json()
    if not "app_msg_list" in msg:
        logging.debug(msg)
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
    cookies_str = 'ptui_loginuin=1371946804; RK=htes1sHRPN; ptcz=a3d21d75959b623656aaf3d38ca3bcb81becb27a3fe5b88388508e5d255e9b5e; tvfe_boss_uuid=d6c5598666965073; pgv_pvid=5496426140; o_cookie=1371946804; pac_uid=1_1371946804; ua_id=7W6H0n1CFOdFthOwAAAAAKOdm07Oq_oHrMgVp6s3EyU=; wxuin=77647480655089; uuid=c33c01ddbcba6273ecf26494b18a7b8d; bizuin=3890910163; ticket=4940d1f627494592c5268e613265f48d9f1de3ff; ticket_id=gh_a67b7e12998e; slave_bizuin=3890910163; cert=rHs0QqmnR6AYHkH6XkogdgeLGgif3EGB; noticeLoginFlag=1; remember_acct=oliver_ck%40outlook.com; rand_info=CAESIP9SfzlJqj3BQkL0FuGDbY9pRTgRlFSe317Mn7GuyH3K; data_bizuin=3890910163; data_ticket=wQku4QfMRJuFV5JS9B+skw6bOa4KWsgdM2b2yD91c0K6lE7JEZZaAzVCWOIJ7iNh; slave_sid=UUhrSlJxc05LVW9paUJabVc0YTY3bGJiS0RiMjNqcjFjaHBVRTNsajRScWc1TWpQdXFoQW5BRjJZMEdlUko2VmY4c1BzQUIwWGRrZ0VQRVFYaVlGaF9pMkpVdFo4NFlWZ3dfSHpWX1M0cFJESzRtTWxOOFpXaE9VeGZQdGRFbTlqOUFHTWhuZTdHNFNmY0ts; slave_user=gh_a67b7e12998e; xid=7ed49574623f7156efb33390fbfcbbfd; openid2ticket_o05Pq5-QxhkTYn3M78Z0_yduOsCI=AFaipTt96ERDcluq9tC1uizfkVyJsPSOUFL9eu5Ardg=; mm_lang=zh_CN; uin=o1371946804; skey=@8pyRF3FET'
    #cookies_str = ''
    create_session(cookies_str)