from excel_io import search_aid, write_to_excel
from wechat_request import create_session, get_articles_session, get_cookies_dict, cookies_str2dict, request_once_session
import time
import random
import logging
import yaml
import os


class WechatMonitor:
    def __init__(self, cfg_path, cookies_str=None, history_path = './wechat_report.xlsx', output_path = './wechat_report.xlsx'):
        logging.info('beign to initial [WechatMonitor]')
        # cfg_path
        self.cfg_path = cfg_path
        self.load_cfg()
        # list history
        self.aid_history = search_aid(history_path)
        # config new cookies
        if cookies_str is not None:
            self.cfg_dict['cookies'] = cookies_str2dict(cookies_str)
        # check ready
        self.check_ready()
        # login
        self.session, self.token = create_session(self.cfg_dict.get('cookies',''))
        self.output_path = output_path
        logging.info('initial [WechatMonitor] finish')

    def listen_wechats(self, itervals = (10, 600)):
        """_summary_

        Args:
            itervals (tuple, optional): (itervals for fakeid, itervals for times). Defaults to (10, 600).
        """
        fake_id_dict = self.cfg_dict.get('fake_id', dict())
        logging.info(f'begin to listen: {len(fake_id_dict)}!!!')
        # save cfg every 1 hour,
        iterval_save_cfg = 3600 // itervals[1]
        times_save_cfg = 0
        while len(fake_id_dict):
            if times_save_cfg % iterval_save_cfg == 0:
                # update cookies
                times_save_cfg = 0
                self.cfg_dict['cookies'] = get_cookies_dict(self.session)
                with open(self.cfg_path, 'w') as f:
                    yaml.dump(self.cfg_dict, f)
                    
            for fake_id in fake_id_dict.keys():
                print(f'\r[{time.asctime( time.localtime(time.time()))}] Listen: {fake_id_dict[fake_id]}', end='')
                article_l = get_articles_session(self.session, fake_id, self.token, exist_aid=self.aid_history.get(fake_id, []))
                if len(article_l) != 0:
                    self.report(fake_id, article_l)
                # 随机时延
                time.sleep(random.randint(itervals[0] // 2, itervals[0] * 2))
            times_save_cfg += 1
            # 随机时延
            time.sleep(random.randint(itervals[1] // 2, itervals[1] * 2))
    
    def get_newest(self):
        # 检查没有历史记录的fakeid
        logging.info('Request previous articles')
        fake_id_dict = self.cfg_dict.get('fake_id', dict())
        for fake_id in fake_id_dict.keys():
            if fake_id in self.aid_history.keys():
                continue
            print(f'\r[{time.asctime( time.localtime(time.time()))}] request: {fake_id_dict[fake_id]}', end='')
            article_l = request_once_session(self.session, fake_id, self.token)
            if len(article_l) > 0:
                self.report(fake_id, article_l)
        print('')
        logging.info('Finish')

    def report(self, fake_id, article_l):
        print('')
        for article in article_l[::-1]:
            logging.warning(f'[New Report!] {fake_id}-{article["time"]}-{article["link"]}')
            self.aid_history[fake_id].append(article['aid'])
        write_to_excel(self.output_path, fake_id, article_l)

    def load_cfg(self):
        if os.path.isfile(self.cfg_path):
            with open(self.cfg_path, 'r') as f:
                self.cfg_dict = yaml.safe_load(f)
                if 'cookies' in self.cfg_dict.keys() and isinstance(self.cfg_dict['cookies'],str):
                    self.cfg_dict['cookies'] = cookies_str2dict(self.cfg_dict['cookies'])
        else:
            self.cfg_dict = dict()
    
    def check_ready(self):
        if self.cfg_dict.get('cookies', None) == None:
            logging.warning('Please input cookies for login')
            self.cfg_dict = cookies_str2dict(input())
        fake_id_dict = self.cfg_dict.get('fake_id', dict())
        if len(fake_id_dict) == 0:
            logging.warning('Please input wechat account with format: [fakeid]:[name], press p for stop')
            try:
                while(1):
                    fake_id, account_name = input().split(':')
                    fake_id_dict[fake_id] = account_name
            except:
                self.cfg_dict['fake_id'] = fake_id_dict
        logging.info('Check Ready')
