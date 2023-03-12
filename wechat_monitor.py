from excel_io import get_history_aid, write_to_excel
from wechat_request import create_session, get_articles_session, get_cookies_dict, cookies_str2dict, request_once_session
import time
import random
import logging
import yaml
import os


class WechatMonitor:
    def __init__(self, cfg_path, cookies_str=None, history_path = './wechat_report.xlsx', output_path = './wechat_report.xlsx'):
        logging.info('初始化 [WechatMonitor]')
        # cfg_path
        self.cfg_path = cfg_path
        self.load_cfg()
        # list history
        self.aid_history = get_history_aid(history_path)
        # config new cookies
        if cookies_str is not None:
            self.cfg_dict['cookies'] = cookies_str2dict(cookies_str)
        # check ready
        self.check_ready()

        self.output_path = output_path
        logging.info('初始化 [WechatMonitor] 完成')

    def listen_wechats(self, itervals = (10, 600)):
        """_summary_

        Args:
            itervals (tuple, optional): (itervals for fakeid, itervals for times). Defaults to (10, 600).
        """
        logging.info(f'开始监听， 监听公众号数目{len(self.fake_id_dict)}!!!')
        logging.info(f'监听列表: {self.fake_id_dict.keys()}')
        # save cfg every 1 hour,
        iterval_save_cfg = 3600 // itervals[1]
        times_save_cfg = 0
        while len(self.fake_id_dict):
            if times_save_cfg % iterval_save_cfg == 0:
                # update cookies
                times_save_cfg = 0
                self.cfg_dict['cookies'] = get_cookies_dict(self.session)
                with open(self.cfg_path, 'w') as f:
                    yaml.dump(self.cfg_dict, f)
                    
            for account_name in self.fake_id_dict.keys():
                print(f'\r[{time.asctime( time.localtime(time.time()))}] 检索公众号[{account_name}]中', end='')
                fake_id = self.fake_id_dict[account_name]
                article_l = get_articles_session(self.session, fake_id, self.token, exist_aid=self.aid_history.get(account_name, []))
                if len(article_l) != 0:
                    self.report(account_name, article_l)
                # 随机时延
                time.sleep(random.randint(itervals[0] // 2, itervals[0] * 2))
            times_save_cfg += 1
            # 随机时延
            time.sleep(random.randint(itervals[1] // 2, itervals[1] * 2))
    
    def get_newest(self):
        # 检查没有历史记录的fakeid
        logging.info('[初始化] 检索历史文章')
        for name in self.fake_id_dict.keys():
            if name in self.aid_history.keys():
                continue
            print(f'\r[{time.asctime( time.localtime(time.time()))}] 检索公众号[{name}]历史文章中', end='')
            article_l = request_once_session(self.session, self.fake_id_dict[name], self.token)
            if len(article_l) > 0:
                self.report(self.fake_id_dict[name], article_l)
        print('')
        logging.info('[初始化] 检索历史文章 [完成]')

    def report(self, account_name, article_l):
        print('')
        logging.warning(f'[监听到新文章!] 公众号[{account_name}] 数量[{len(article_l)}]')
        for article in article_l[::-1]:
            self.aid_history[account_name].append(article['aid'])
            logging.warning(article["link"])
            article['公众号'] = account_name
        write_to_excel(self.output_path, article_l)
        logging.info(f'[已写入日志]')

    def load_cfg(self):
        if os.path.isfile(self.cfg_path):
            with open(self.cfg_path, 'r') as f:
                self.cfg_dict = yaml.safe_load(f)
                if 'cookies' in self.cfg_dict.keys() and isinstance(self.cfg_dict['cookies'],str):
                    self.cfg_dict['cookies'] = cookies_str2dict(self.cfg_dict['cookies'])
        else:
            self.cfg_dict = dict()
        self.fake_id_dict =  self.cfg_dict.get('fake_id', dict())
        
    
    def check_ready(self):
        if self.cfg_dict.get('cookies', None) == None:
            logging.warning('请输入登录cookies')
            self.cfg_dict = cookies_str2dict(input())
        while(1):
            # login
            self.session, self.token = create_session(self.cfg_dict.get('cookies',''))
            if self.session != None:
                logging.info('登录成功')
                break
            logging.warning('[登录失败] 请更新cookies(输入字符串):')
            self.cfg_dict.update(cookies_str2dict(input()))
        print('是否要监听新的公众号？([n]/y)')
        if (input() == 'y') and (len(self.fake_id_dict) == 0):
            logging.warning('请按如下格式输入新的微信公众号: [公众号]:[fakeid], 输入任意按键退出')
            try:
                while(1):
                    account_name, fake_id = input().split(':')
                    self.fake_id_dict[account_name] = fake_id
            except:
                pass
